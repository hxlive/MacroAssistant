# -*- coding: utf-8 -*-
# core_engine.py
# 描述:自动化宏的核心功能引擎
# Version: 1.8.3
CORE_VERSION = "1.8.3"
# ======================================================================
# Module index
# ======================================================================
# 1. Exceptions, imports, optional engines, and global configuration
# 2. Macro schema, persistence, runtime context, and loop cache
# 3. Screenshot, image matching, variables, expressions, and file safety
# 4. Linear action handlers and dispatch table
# 5. Control-flow handlers, dispatch table, and execution loop
# 6. Performance monitoring, thread stopping, and macro validation

# ======================================================================
# 即时中断异常
# ======================================================================
class MacroStopException(BaseException):
    """快捷键触发时注入到执行线程的异常，强制立刻中断宏。
    继承 BaseException 而非 Exception，确保不被 except Exception 误吞。
    """
    pass

class LoopConditionCheckError(RuntimeError):
    """Raised when a loop exit condition cannot be checked safely."""
    pass

class ScreenshotUnavailableError(RuntimeError):
    """Raised when the desktop cannot be captured."""
    pass

import pyautogui
import time

import re
import pyperclip
import json
import os
import sys
import subprocess
import shlex
import ctypes
from collections import OrderedDict
import threading
import tempfile
import ast
import math
import operator as operator_module
from decimal import Decimal, InvalidOperation
from sys_utils import HotkeyUtils, capture_physical_bbox

import logging
logger = logging.getLogger(__name__)

if sys.platform == 'win32':
    ctypes.windll.shell32.CommandLineToArgvW.argtypes = (ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int))
    ctypes.windll.shell32.CommandLineToArgvW.restype = ctypes.POINTER(ctypes.c_wchar_p)
    ctypes.windll.kernel32.LocalFree.argtypes = (ctypes.c_void_p,)
    ctypes.windll.kernel32.LocalFree.restype = ctypes.c_void_p

# ======================================================================
# 宏文件持久化工具 (属模块索引第 2 节)
# ======================================================================
class MacroPersistence:
    @staticmethod
    def convert_to_native(obj):
        """递归转换所有值为 Python 原生类型 (处理 numpy 等类型)"""
        try:
            import numpy as np
            numpy_types = (np.integer, np.floating)
        except ImportError:
            numpy_types = ()
            
        if isinstance(obj, dict):
            return {k: MacroPersistence.convert_to_native(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [MacroPersistence.convert_to_native(item) for item in obj]
        elif numpy_types and isinstance(obj, numpy_types):
            return obj.item()
        else:
            return obj

    @staticmethod
    def save(file_path, steps):
        native_steps = MacroPersistence.convert_to_native(steps)
        tmp_path = file_path + '.tmp'
        with open(tmp_path, 'w', encoding='utf-8') as f:
            f.write('[\n')
            for i, step in enumerate(native_steps):
                step_str = json.dumps(step, ensure_ascii=False)
                if i < len(native_steps) - 1:
                    f.write(f'    {step_str},\n')
                else:
                    f.write(f'    {step_str}\n')
            f.write(']\n')
        os.replace(tmp_path, file_path)

    @staticmethod
    def load(file_path):
        """从 JSON 文件加载宏"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data



try:
    import pygetwindow as gw
    PYGETWINDOW_AVAILABLE = True
except ImportError:
    PYGETWINDOW_AVAILABLE = False
    logger.error("FAIL 未找到 pygetwindow 库 (pip install pygetwindow)。'激活窗口' 功能将不可用。")

# ======================================================================
# 全局配置
# ======================================================================
FORCE_OCR_ENGINE = None 
ENABLE_GLOBAL_FALLBACK = True
# 条件循环检测间隔 (秒) - 平衡流畅度与准确率
LOOP_CHECK_INTERVAL = 0.2  # 从 0.5s 降低到 0.2s，平衡响应速度与检测开销
_MOUSE_MOVE_FRAME_INTERVAL = 1.0 / 60.0  # 约 60 FPS，兼顾平滑度与停止响应
# 性能与缓存相关常量
LOOP_PHYSICAL_COOLDOWN = 0.05  # 循环物理冷却时间（秒），防止队列瞬间爆炸
CACHE_BOX_PADDING = 50  # 缓存区域扩展边距（像素）
TEMPLATE_CACHE_SIZE = 200  # 模板缓存最大条目数
TEMPLATE_CACHE_MAX_BYTES = 128 * 1024 * 1024  # 模板缓存最大总内存（128 MB）
QUICK_CHECK_SCALES = (1.0, 0.9, 1.1)  # 快速检查尝试的缩放比例
GOTO_LABEL_DEFAULT_MAX_JUMPS = 100
APP_DIR = os.path.dirname(os.path.abspath(__file__))
_VAR_PATTERN = re.compile(r'\{([^{}]+)\}')
_GOTO_LABEL_PATTERN = re.compile(r'^\s*(?:LABEL|标签)\s*[:：]\s*(.+?)\s*$', re.IGNORECASE)
_FIELD_SPLIT_PATTERN = re.compile(r'[,，\t|]+')

TEXT_OUTPUT_EXTENSIONS = {
    '.txt', '.log', '.ini', '.cfg', '.csv', '.tsv',
    '.json', '.jsonl', '.md', '.yaml', '.yml', '.xml'
}


try:
    import ocr_engine
except ImportError:
    logger.error("未找到 'ocr_engine.py'。")
    class ocr_engine:
        def find_text_location(*args, **kwargs): return None
        WINOCR_AVAILABLE = False
        TESSERACT_AVAILABLE = False
        RAPIDOCR_AVAILABLE = False

try:
    import vlm_engine
except ImportError:
    logger.error("FAIL 未找到 'vlm_engine.py'。AI 自然语言指令功能将不可用。")
    class vlm_engine:
        def find_location_by_vlm(*args, **kwargs): return None
        VLM_AVAILABLE = False

try:
    import cv2
    import numpy as np 
    OPENCV_AVAILABLE = True
    logger.info("OpenCV engine ready")
except ImportError:
    OPENCV_AVAILABLE = False
    logger.warning("OpenCV not found, fallback to slower image matching")

# ======================================================================
# 宏定义元数据
# ======================================================================
class MacroSchema:
    ACTION_TRANSLATIONS = {
        'CLICK':          '01. 点击鼠标',
        'MOVE_TO':        '02. 移动到 (绝对坐标)',
        'MOVE_OFFSET':    '03. 相对移动',
        'SCROLL':         '04. 滚动滚轮',
        'WAIT':           '05. 等待',
        'TYPE_TEXT':      '06. 输入文本',
        'PRESS_KEY':      '07. 按下按键',
        'ACTIVATE_WINDOW': '08. 激活窗口 (按标题)',
        'NOTE':           "09. 备注",
        'FIND_IMAGE':     '10. 查找图像',
        'FIND_TEXT':      '11. 查找文本 (OCR)',
        'IF_IMAGE_FOUND':  '12. IF 找到图像',
        'IF_TEXT_FOUND':   '13. IF 找到文本',
        'IF_VAR':         '14. IF 变量比较',
        'ELSE':           '15. ELSE（否则）',
        'END_IF':         '16. END_IF（结束 IF）',
        'LOOP_START':     '17. 循环开始',          # Loop
        'END_LOOP':       '18. 结束循环',       # EndLoop
        'SET_VAR':        '19. 设置变量',        # Set Var
        'CALCULATE':      '20. 变量计算',      # Calculate
        'EXTRACT_VAR':    '21. 正则提取变量',    # Extract
        'PROMPT_INPUT':    '22. 人工输入',   # Prompt Input
        'GOTO_LABEL':     '23. 跳转到标签',
        'GOTO_IF':        '24. 条件跳转',        # Goto If
        'READ_FILE':      '25. 读取文本文件到变量',
        'WRITE_FILE':     '26. 写入文本文件',     # Write File
        'FOREACH_LINE':    '27. 批量处理文本行',   # Batch Lines
        'END_FOREACH':     '28. 结束批量处理',
        'RUN':            '29. 执行命令/脚本',
        'AI_COMMAND':      '30. AI 自然语言指令',
    }
    ACTION_KEYS_TO_NAME = {v: k for k, v in ACTION_TRANSLATIONS.items()}
    RUN_TYPE_OPTIONS = {
        'command (命令)': 'command',
        'script (脚本)': 'script',
    }
    RUN_TYPE_DISPLAY_BY_VALUE = {v: k for k, v in RUN_TYPE_OPTIONS.items()}
    CONTROL_FLOW_ACTIONS = {'IF_IMAGE_FOUND', 'IF_TEXT_FOUND', 'IF_VAR', 'ELSE', 'END_IF', 'LOOP_START', 'END_LOOP', 'FOREACH_LINE', 'END_FOREACH'}
    
    LANG_OPTIONS = {'chi_sim (简体中文)': 'chi_sim', 'eng (英文)': 'eng'}
    LANG_VALUES_TO_NAME = {v: k for k, v in LANG_OPTIONS.items()}
    
    CLICK_OPTIONS = {'left (左键)': 'left', 'right (右键)': 'right', 'middle (中键)': 'middle'}
    CLICK_VALUES_TO_NAME = {v: k for k, v in CLICK_OPTIONS.items()}

# ======================================================================
# 性能监控
# ======================================================================
class PerformanceMonitor:
    def __init__(self): self.reset()
    def reset(self):
        self.image_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}
        self.ocr_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}
    def _get_stats_for(self, stats_dict):
        total = stats_dict['hits'] + stats_dict['misses']
        if total == 0: return "(无记录)"
        unique_hits = stats_dict['hits'] - stats_dict['loop_hits']
        total_valid = unique_hits + stats_dict['misses']
        hit_rate = (unique_hits / total_valid * 100) if total_valid > 0 else 0
        avg_ms = (sum(stats_dict['times']) / len(stats_dict['times']) * 1000) if stats_dict['times'] else 0
        return f"(命中{hit_rate:.0f}% | 循环{stats_dict['loop_hits']} | 均耗{avg_ms:.0f}ms)"
    def record_hit(self, is_loop, is_ocr):
        s = self.ocr_stats if is_ocr else self.image_stats
        s['hits'] += 1
        if is_loop: s['loop_hits'] += 1
    def record_miss(self, is_ocr): (self.ocr_stats if is_ocr else self.image_stats)['misses'] += 1
    def record_time(self, dt, is_ocr): (self.ocr_stats if is_ocr else self.image_stats)['times'].append(dt)
    def get_stats(self): return f"图像{self._get_stats_for(self.image_stats)} | OCR{self._get_stats_for(self.ocr_stats)}"

perf = PerformanceMonitor()

# ======================================================================
# 循环缓存管理器
# ======================================================================
class LoopCacheManager:
    def __init__(self):
        self._lock = threading.RLock()
        self.reset()
    
    def reset(self):
        with self._lock:
            self.caches = {}
            self.stack = []
        
    def get_current_loop_id(self):
        with self._lock:
            return self.stack[-1] if self.stack else None

    def enter(self, loop_id):
        with self._lock:
            if loop_id not in self.caches:
                self.caches[loop_id] = {}
            self.stack.append(loop_id)

    def exit(self):
        with self._lock:
            if self.stack:
                loop_id = self.stack.pop()
                # Clear this loop cache here; execute_steps finally also resets as a fallback.
                if loop_id in self.caches:
                    del self.caches[loop_id]


    def get(self, sig): 
        with self._lock:
            loop_id = self.stack[-1] if self.stack else None
            return self.caches.get(loop_id, {}).get(sig) if loop_id else None

    def set(self, sig, loc): 
        with self._lock:
            loop_id = self.stack[-1] if self.stack else None
            if loop_id:
                if loop_id not in self.caches:
                     self.caches[loop_id] = {}
                self.caches[loop_id][sig] = loc

loop_cache = LoopCacheManager()

class RunContext:
    """
    宏执行上下文，封装单次宏运行的所有状态。
    通过魔术方法实现对旧代码字典式访问的 100% 兼容。
    不定义 __bool__，依赖 Python 默认对象真值（永远为 True），
    以正确兼容代码中 `if ctx:` / `if not ctx:` 的 None 判断语义。
    """
    def __new__(cls, raw_dict=None):
        if isinstance(raw_dict, cls):
            return raw_dict
        return super().__new__(cls)

    def __init__(self, raw_dict=None):
        if isinstance(raw_dict, RunContext):
            return

        self._data = raw_dict if isinstance(raw_dict, dict) else {}
        self._data.setdefault('vars', {})
        self._data.setdefault('stop_requested', False)
        stop_event = self._data.get('stop_event')
        if not callable(getattr(stop_event, 'is_set', None)) or not callable(getattr(stop_event, 'set', None)):
            stop_event = threading.Event()
            self._data['stop_event'] = stop_event
        if self._data['stop_requested']:
            stop_event.set()
        self._data.setdefault('last_pos', (None, None))
        self._data.setdefault('clipboard_var', '')
        self._data.setdefault('_active_processes', set())
        self._data.setdefault('_active_process_lock', threading.RLock())
        self._data['_goto_counts'] = {}

        self.perf = PerformanceMonitor()
        self.loop_cache = LoopCacheManager()

    def check_stop(self):
        stop_event = self._data.get('stop_event')
        return bool(
            self._data.get('stop_requested', False)
            or (callable(getattr(stop_event, 'is_set', None)) and stop_event.is_set())
        )

    def request_stop(self):
        self['stop_requested'] = True

    @property
    def vars(self):
        return self._data.setdefault('vars', {})

    def get_var(self, name, default=None):
        return self.vars.get(name, default)

    def set_var(self, name, value):
        self.vars[name] = value

    def get_last_pos(self):
        return self._data.setdefault('last_pos', (None, None))

    def set_last_pos(self, x, y=None):
        self._data['last_pos'] = x if y is None else (x, y)

    def get_clipboard_var(self, default=''):
        return self._data.get('clipboard_var', default)

    def set_clipboard_var(self, value):
        self._data['clipboard_var'] = value

    def get_macro_base_dir(self):
        return self._data.get('macro_base_dir')

    def get_allowed_file_roots(self):
        return self._data.get('allowed_file_roots', []) or []

    def allow_external_paths(self):
        return bool(self._data.get('allow_external_paths', False))

    def __getitem__(self, key): return self._data[key]
    def __setitem__(self, key, val):
        self._data[key] = val
        if key == 'stop_requested':
            stop_event = self._data.get('stop_event')
            if callable(getattr(stop_event, 'set', None)) and val:
                stop_event.set()
            elif callable(getattr(stop_event, 'clear', None)):
                stop_event.clear()
    def get(self, key, default=None): return self._data.get(key, default)
    def setdefault(self, key, default=None): return self._data.setdefault(key, default)
    def __contains__(self, key): return key in self._data

def _get_perf(ctx): return getattr(ctx, 'perf', perf)
def _get_loop_cache(ctx): return getattr(ctx, 'loop_cache', loop_cache)


def _ctx_check_stop(ctx):
    if ctx is None:
        return False
    if hasattr(ctx, 'check_stop'):
        return ctx.check_stop()
    stop_event = ctx.get('stop_event')
    return bool(
        ctx.get('stop_requested', False)
        or (callable(getattr(stop_event, 'is_set', None)) and stop_event.is_set())
    )


def is_stop_requested(ctx):
    """Return whether either the legacy flag or cooperative stop event is set."""
    return _ctx_check_stop(ctx)


def request_stop(ctx):
    """Set both cooperative stop representations while preserving dict contexts."""
    if ctx is None:
        return
    if hasattr(ctx, 'request_stop'):
        ctx.request_stop()
        return
    stop_event = ctx.get('stop_event')
    if not callable(getattr(stop_event, 'set', None)):
        stop_event = threading.Event()
        ctx['stop_event'] = stop_event
    ctx['stop_requested'] = True
    stop_event.set()


def _ctx_vars(ctx):
    if ctx is None:
        return {}
    if hasattr(ctx, 'vars'):
        return ctx.vars
    return ctx.setdefault('vars', {})


def _ctx_get_var(ctx, name, default=None):
    if hasattr(ctx, 'get_var'):
        return ctx.get_var(name, default)
    return _ctx_vars(ctx).get(name, default)


def _ctx_set_var(ctx, name, value):
    if ctx is None:
        return
    if hasattr(ctx, 'set_var'):
        ctx.set_var(name, value)
    else:
        _ctx_vars(ctx)[name] = value


def _ctx_get_last_pos(ctx):
    if ctx is None:
        return (None, None)
    if hasattr(ctx, 'get_last_pos'):
        return ctx.get_last_pos()
    return ctx.setdefault('last_pos', (None, None))


def _ctx_set_last_pos(ctx, x, y=None):
    if ctx is None:
        return
    if hasattr(ctx, 'set_last_pos'):
        ctx.set_last_pos(x, y)
    else:
        ctx['last_pos'] = x if y is None else (x, y)


def _ctx_get_clipboard_var(ctx, default=''):
    if ctx is None:
        return default
    if hasattr(ctx, 'get_clipboard_var'):
        return ctx.get_clipboard_var(default)
    return ctx.get('clipboard_var', default)


def _ctx_set_clipboard_var(ctx, value):
    if ctx is None:
        return
    if hasattr(ctx, 'set_clipboard_var'):
        ctx.set_clipboard_var(value)
    else:
        ctx['clipboard_var'] = value


def _ctx_get_macro_base_dir(ctx):
    if ctx is None:
        return None
    if hasattr(ctx, 'get_macro_base_dir'):
        return ctx.get_macro_base_dir()
    return ctx.get('macro_base_dir')


def _ctx_get_allowed_file_roots(ctx):
    if ctx is None:
        return []
    if hasattr(ctx, 'get_allowed_file_roots'):
        return ctx.get_allowed_file_roots()
    return ctx.get('allowed_file_roots', []) or []


def _ctx_allow_external_paths(ctx):
    if ctx is None:
        return False
    if hasattr(ctx, 'allow_external_paths'):
        return ctx.allow_external_paths()
    return bool(ctx.get('allow_external_paths', False))


# ======================================================================
# 核心工具函数
# ======================================================================
def _safe_int(value, default=None, min_value=None, max_value=None):
    try:
        result = int(value)
    except (TypeError, ValueError):
        return default, False
    if min_value is not None and result < min_value:
        return default, False
    if max_value is not None and result > max_value:
        return default, False
    return result, True

def _safe_float(value, default=None, min_value=None, max_value=None):
    try:
        result = float(value)
    except (TypeError, ValueError):
        return default, False
    if not math.isfinite(result):
        return default, False
    if min_value is not None and result < min_value:
        return default, False
    if max_value is not None and result > max_value:
        return default, False
    return result, True

def _safe_param_int(p, name, default=None, min_value=None, max_value=None):
    value = p.get(name, default)
    if value is None or (isinstance(value, str) and not value.strip()):
        value = default
    return _safe_int(value, default, min_value=min_value, max_value=max_value)


def _safe_param_float(p, name, default=None, min_value=None, max_value=None):
    value = p.get(name, default)
    if value is None or (isinstance(value, str) and not value.strip()):
        value = default
    return _safe_float(value, default, min_value=min_value, max_value=max_value)


def _parse_optional_coords_or_break(p, action, x_key='x', y_key='y'):
    """解析可选的整数坐标，失败时返回 _ACTION_BREAK。"""
    try:
        x = int(p.get(x_key, '')) if str(p.get(x_key, '')).strip() else None
        y = int(p.get(y_key, '')) if str(p.get(y_key, '')).strip() else None
    except (ValueError, TypeError):
        logger.error(f"  {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_required_coords_or_break(p, action, x_key='x', y_key='y'):
    """解析必填的整数坐标，失败时返回 _ACTION_BREAK。"""
    try:
        x, y = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError):
        logger.error(f"  {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_offset_or_break(p, action, x_key='x_offset', y_key='y_offset'):
    """解析整数偏移量，失败时返回 _ACTION_BREAK。"""
    try:
        ox, oy = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError):
        logger.error(f"  {action} invalid offset")
        return _ACTION_BREAK
    return (ox, oy)


def _parse_positive_int_or_break(p, action, key, default=None):
    """解析正整数参数，失败时返回 _ACTION_BREAK。"""
    value, ok = _safe_param_int(p, key, default, min_value=1)
    if not ok:
        logger.error(f"  {action} {key} must be a positive integer")
        return _ACTION_BREAK
    return value


def _parse_int_or_break(p, action, key, default=0):
    """解析整数参数，失败时返回 _ACTION_BREAK。"""
    value, ok = _safe_param_int(p, key, default)
    if not ok:
        logger.error(f"  {action} {key} must be an integer")
        return _ACTION_BREAK
    return value


def _warn_param_default(action, name, default):
    logger.warning(f"  {action} invalid parameter '{name}', using default: {default}")

def _error_param_skip(action, name, expected):
    logger.error(f"  {action} invalid parameter '{name}' expected {expected}; step skipped")


def _run_with_fail_stop(action, p, fn, ctx=None, var_name=None, default_value=''):
    """执行可能抛异常的 IO/提取操作，统一处理 fail_stop 行为。"""
    try:
        return fn()
    except Exception as e:
        if ctx is not None and var_name is not None:
            _ctx_set_var(ctx, var_name, default_value)
        logger.error(f"  {action} 失败: {e}")
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_DONE


def _close_image_quietly(image):
    if image:
        try:
            image.close()
        except Exception:
            pass


def _log_value_summary(value):
    """Describe runtime values without persisting their content in logs."""
    if value is None:
        return 'None'
    if isinstance(value, str):
        return f'str({len(value)} chars)'
    if isinstance(value, (bytes, bytearray)):
        return f'{type(value).__name__}({len(value)} bytes)'
    if isinstance(value, (list, tuple, set, dict)):
        return f'{type(value).__name__}({len(value)} items)'
    return type(value).__name__

def _log_display_text(value):
    """Render diagnostic text on one log line while preserving its content."""
    return json.dumps(str(value), ensure_ascii=False)

def _log_display_path(path, ctx=None, basename_only=False):
    """Return a useful path label without exposing macro-local absolute roots."""
    raw_path = os.fspath(path) if path else ''
    if not raw_path:
        return ''
    if basename_only:
        return os.path.basename(raw_path)

    absolute_path = os.path.abspath(raw_path)
    macro_base = ctx.get('macro_base_dir') if ctx else None
    if macro_base and _is_path_inside(absolute_path, macro_base):
        return os.path.relpath(absolute_path, os.path.abspath(macro_base))
    return absolute_path


def _coerce_bbox(raw_bbox):
    if not isinstance(raw_bbox, (list, tuple)):
        return None
    try:
        bbox = [int(value) for value in raw_bbox]
    except (TypeError, ValueError):
        return None
    if len(bbox) == 2:
        bbox = [bbox[0], bbox[1], bbox[0] + 1, bbox[1] + 1]
    if len(bbox) < 4 or bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
        return None
    return bbox[:4]

def smart_screenshot(region=None, pad=0):
    try:
        bbox = None
        if region:
            if sys.platform == 'win32':
                x1 = region[0] - pad
                y1 = region[1] - pad
            else:
                x1 = max(0, region[0] - pad)
                y1 = max(0, region[1] - pad)
            x2 = region[0] + region[2] + pad
            y2 = region[1] + region[3] + pad
            bbox = (x1, y1, x2, y2)

        return capture_physical_bbox(bbox)
    except OSError as e:
        raise ScreenshotUnavailableError("Screen capture is unavailable") from e
def bbox_to_region(raw_bbox):
    """Convert an (x1, y1, x2, y2) box into a screenshot region."""
    bbox = _coerce_bbox(raw_bbox)
    if not bbox:
        return None
    return (bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1])

def _padded_bbox_to_region(raw_bbox, pad):
    bbox = _coerce_bbox(raw_bbox)
    if not bbox:
        return None
    left = bbox[0] - pad
    top = bbox[1] - pad
    if sys.platform != 'win32':
        left = max(0, left)
        top = max(0, top)
    right = bbox[2] + pad
    bottom = bbox[3] + pad
    return (left, top, right - left, bottom - top)

def _copy_to_clipboard_with_retry(text, ctx=None, retries=3, delay=0.2):
    for _ in range(retries):
        if _ctx_check_stop(ctx):
            raise MacroStopException("Stop requested while writing to clipboard")
        try:
            pyperclip.copy(text)
            return True
        except Exception:
            time.sleep(delay)
    return False

SCALES = (1.0, 0.9, 1.1, 0.8, 1.2)


_TEMPLATE_CACHE = OrderedDict()
_TEMPLATE_CACHE_BYTES = 0
_TEMPLATE_CACHE_LOCK = threading.RLock()


def _clear_template_cache():
    global _TEMPLATE_CACHE_BYTES
    with _TEMPLATE_CACHE_LOCK:
        _TEMPLATE_CACHE.clear()
        _TEMPLATE_CACHE_BYTES = 0


def _get_template_cached(path, scale, modified_ns, file_size):
    global _TEMPLATE_CACHE_BYTES
    key = (path, scale, modified_ns, file_size)
    with _TEMPLATE_CACHE_LOCK:
        cached = _TEMPLATE_CACHE.pop(key, None)
        if cached is not None:
            _TEMPLATE_CACHE[key] = cached
            return cached[:3]
        stale_keys = [cached_key for cached_key in _TEMPLATE_CACHE
                      if cached_key[:2] == key[:2]]
        for stale_key in stale_keys:
            _TEMPLATE_CACHE_BYTES -= _TEMPLATE_CACHE.pop(stale_key)[3]

    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        return None, 0, 0
    if scale != 1.0:
        h, w = img.shape[:2]
        img = cv2.resize(
            img,
            (int(w * scale), int(h * scale)),
            interpolation=cv2.INTER_AREA,
        )
    result = (img, img.shape[1], img.shape[0])
    item_bytes = int(getattr(img, 'nbytes', 0))
    if item_bytes <= 0 or item_bytes > TEMPLATE_CACHE_MAX_BYTES:
        return result

    with _TEMPLATE_CACHE_LOCK:
        existing = _TEMPLATE_CACHE.pop(key, None)
        if existing is not None:
            _TEMPLATE_CACHE_BYTES -= existing[3]
        _TEMPLATE_CACHE[key] = (*result, item_bytes)
        _TEMPLATE_CACHE_BYTES += item_bytes
        while (len(_TEMPLATE_CACHE) > TEMPLATE_CACHE_SIZE or
               _TEMPLATE_CACHE_BYTES > TEMPLATE_CACHE_MAX_BYTES):
            _, removed = _TEMPLATE_CACHE.popitem(last=False)
            _TEMPLATE_CACHE_BYTES -= removed[3]
    return result


_get_template_cached.cache_clear = _clear_template_cache


def _get_template_cache_info():
    with _TEMPLATE_CACHE_LOCK:
        return len(_TEMPLATE_CACHE), _TEMPLATE_CACHE_BYTES


def _get_template(path, scale):
    """Return a cached template, invalidating it when the source file changes."""
    try:
        stat = os.stat(path)
        signature = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        signature = (None, None)
    return _get_template_cached(path, scale, *signature)

def find_image_cv2(path, conf, screenshot_pil, offset=(0,0), enhanced_mode=False, ctx=None):
    if not OPENCV_AVAILABLE: return None
    try:
        t0 = time.monotonic()
        screen_gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
        best = (-1, None, 0, 0)
        scales_to_try = SCALES if enhanced_mode else [1.0]
        for scale in scales_to_try:
            _raise_if_stop_requested(ctx, "Stop requested during image matching")
            tmpl, tw, th = _get_template(path, scale)
            if tmpl is None or th > screen_gray.shape[0] or tw > screen_gray.shape[1]: continue
            res = cv2.matchTemplate(screen_gray, tmpl, cv2.TM_CCOEFF_NORMED)
            min_v, max_v, min_l, max_l = cv2.minMaxLoc(res)
            if max_v > best[0]: best = (max_v, max_l, tw, th)
            if best[0] >= 0.95 and best[0] >= conf: break 
        val, loc, w, h = best
        if val >= conf and loc:
            cx, cy = offset[0] + loc[0] + w//2, offset[1] + loc[1] + h//2
            _get_perf(ctx).record_time(time.monotonic()-t0, False)
            return (cx, cy, w, h), val
    except (cv2.error, ValueError, TypeError, AttributeError) as e:
        logger.error(f"CV2找图错误: {e}")
    return None

def quick_check_cv2(path, conf, screenshot_pil, offset, target_loc, enhanced_mode=False):
    """
    [补丁优化] 快速检查图片是否仍在缓存位置
    
    优化: 支持多缩放比例检查，避免缓存失效
    
    Args:
        path: 图片文件路径
        conf: 置信度阈值
        screenshot_pil: PIL截图对象
        offset: 截图偏移量 (x, y)
        target_loc: 目标位置 (x, y)
        enhanced_mode: 是否启用增强模式
        
    Returns:
        bool: 是否在目标位置找到图片
    """
    if not OPENCV_AVAILABLE: return False
    try:
        # [补丁优化] 尝试多个缩放比例，避免因缩放不匹配导致误判
        scales_to_try = QUICK_CHECK_SCALES if enhanced_mode else [1.0]
        for scale in scales_to_try:
            tmpl, tw, th = _get_template(path, scale)
            if tmpl is None: continue
            
            pad_w, pad_h = tw//2 + 15, th//2 + 15
            rel_x, rel_y = target_loc[0] - offset[0], target_loc[1] - offset[1]
            l, t = max(0, rel_x - pad_w), max(0, rel_y - pad_h)
            r, b = min(screenshot_pil.width, rel_x + pad_w), min(screenshot_pil.height, rel_y + pad_h)
            if r <= l or b <= t: continue
            
            crop = cv2.cvtColor(np.array(screenshot_pil.crop((l, t, r, b))), cv2.COLOR_RGB2GRAY)
            if crop.shape[0] < th or crop.shape[1] < tw:
                continue
            _, max_v, _, _ = cv2.minMaxLoc(cv2.matchTemplate(crop, tmpl, cv2.TM_CCOEFF_NORMED))
            
            if max_v >= conf:
                return True  # 找到匹配，立即返回
        
        return False  # 所有缩放比例都不匹配
    except (ValueError, TypeError, AttributeError, IndexError) + ((cv2.error,) if OPENCV_AVAILABLE else ()) as e:
        # Loop cache misses can happen frequently; keep this hot path quiet.
        logger.error(f"异常: {e}")
        return False

# ======================================================================
# 主执行引擎
# ======================================================================
def _normalize_goto_label(label):
    return str(label or '').strip().casefold()

def _extract_goto_label_from_note(text):
    match = _GOTO_LABEL_PATTERN.match(str(text or ''))
    if not match:
        return None
    label = match.group(1).strip()
    return label or None

def _build_goto_label_table(steps):
    labels = {}
    loop_depth = 0
    for idx, step in enumerate(steps):
        action = step.get('action', '')
        if action in ('END_LOOP', 'END_FOREACH') and loop_depth > 0:
            loop_depth -= 1

        if action == 'NOTE' and step.get('enabled', True):
            label = _extract_goto_label_from_note(step.get('params', {}).get('text', ''))
            if label:
                if loop_depth > 0:
                    raise ValueError(f"标签 '{label}' 位于循环块内部，当前版本暂不支持跳转到循环内部标签")
                key = _normalize_goto_label(label)
                if key in labels:
                    prev_idx = labels[key]['index']
                    raise ValueError(f"标签重复: '{label}' 同时出现在第 {prev_idx + 1} 步和第 {idx + 1} 步")
                labels[key] = {'name': label, 'index': idx}

        if action in ('LOOP_START', 'FOREACH_LINE'):
            loop_depth += 1
    return labels

def _normalize_path_for_compare(path):
    return os.path.normcase(os.path.realpath(os.path.abspath(path)))

def _is_path_inside(path, root):
    try:
        return os.path.commonpath([
            _normalize_path_for_compare(path),
            _normalize_path_for_compare(root),
        ]) == _normalize_path_for_compare(root)
    except (ValueError, OSError):
        return False

def _get_allowed_file_roots(ctx):
    roots = [APP_DIR, os.getcwd(), tempfile.gettempdir()]
    base_dir = _ctx_get_macro_base_dir(ctx)
    if base_dir:
        roots.append(base_dir)
    for root in _ctx_get_allowed_file_roots(ctx):
        if root:
            roots.append(root)

    normalized = []
    for root in roots:
        try:
            normalized_root = _normalize_path_for_compare(root)
            if normalized_root not in normalized:
                normalized.append(normalized_root)
        except (TypeError, ValueError, OSError):
            continue
    return normalized

def _resolve_safe_file_path(file_path, ctx, purpose='file access'):
    file_path = str(file_path or '').strip()
    if not file_path:
        raise ValueError('路径为空')

    base_dir = _ctx_get_macro_base_dir(ctx) or os.getcwd()
    expanded = os.path.expandvars(os.path.expanduser(file_path))
    if not os.path.isabs(expanded):
        expanded = os.path.join(base_dir, expanded)
    resolved = _normalize_path_for_compare(expanded)

    if _ctx_allow_external_paths(ctx):
        return resolved

    for root in _get_allowed_file_roots(ctx):
        if _is_path_inside(resolved, root):
            return resolved

    raise PermissionError(f"{purpose} 路径不在允许目录内: {file_path}")

def _ensure_text_output_path(safe_path, purpose='WRITE_FILE'):
    ext = os.path.splitext(str(safe_path))[1].lower()
    if ext not in TEXT_OUTPUT_EXTENSIONS:
        allowed = ', '.join(sorted(TEXT_OUTPUT_EXTENSIONS))
        raise ValueError(f"{purpose} 仅支持文本输出文件，允许扩展名: {allowed}")
    return safe_path


def _render_vars(param_str, ctx):
    if not isinstance(param_str, str) or not param_str: return param_str
    if '{' not in param_str: return param_str
    vars_dict = _ctx_vars(ctx)
    def repl(match):
        var_name = match.group(1)
        return str(vars_dict.get(var_name, match.group(0)))
    return _VAR_PATTERN.sub(repl, param_str)


def _render_param(p, name, ctx, default=''):
    return _render_vars(str(p.get(name, default)), ctx)


def _render_param_stripped(p, name, ctx, default=''):
    return _render_param(p, name, ctx, default).strip()


def _render_string_params(params, ctx):
    return {k: _render_vars(v, ctx) if isinstance(v, str) else v for k, v in params.items()}


def _call_vlm_with_stop(instruction, region, ctx):
    done = threading.Event()
    result = {'coords': None, 'error': None}

    def worker():
        try:
            result['coords'] = vlm_engine.find_location_by_vlm(instruction, region=region)
        except Exception as e:
            result['error'] = e
        finally:
            done.set()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    while not done.wait(0.1):
        if _ctx_check_stop(ctx):
            raise MacroStopException("Stop requested while VLM request is active")
    if result['error'] is not None:
        raise result['error']
    return result['coords']

def _parse_bool_param(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in ('1', 'true', 'yes', 'y', 'on'):
        return True
    if text in ('0', 'false', 'no', 'n', 'off'):
        return False
    return default

def _split_field_names(value):
    if isinstance(value, (list, tuple)):
        raw_names = value
    else:
        raw_names = _FIELD_SPLIT_PATTERN.split(str(value or ''))
    return [str(name).strip() for name in raw_names if str(name).strip()]

def _normalize_split_delimiter(value):
    text = str(value or '')
    lower = text.strip().lower()
    if lower in ('\\t', 'tab'):
        return '\t'
    if lower in ('\\s', 'space'):
        return ' '
    return text

def _set_foreach_line_vars(loop_data, ctx):
    vars_dict = _ctx_vars(ctx)
    items = loop_data.get('items', [])
    index = loop_data.get('index', 0)
    total = len(items)
    line = items[index] if 0 <= index < total else ''

    current_line_var = loop_data.get('current_line_var') or 'current_line'
    index_var = loop_data.get('index_var') or 'loop_index'
    total_var = loop_data.get('total_var') or 'loop_total'
    vars_dict[current_line_var] = line
    vars_dict[index_var] = str(index + 1)
    vars_dict[total_var] = str(total)

    field_names = loop_data.get('field_names') or []
    delimiter = loop_data.get('delimiter', '')
    if field_names:
        if delimiter:
            values = line.split(delimiter)
        else:
            values = [line]
        if loop_data.get('strip_fields', True):
            values = [value.strip() for value in values]
        for field_index, name in enumerate(field_names):
            vars_dict[name] = values[field_index] if field_index < len(values) else ''

def _compare_values(left, operator, right):
    left = str(left)
    right = str(right)
    operator = str(operator or '==')

    if operator == '包含':
        return right in left
    if operator == '不包含':
        return right not in left

    numeric_ops = {
        '==': lambda a, b: a == b,
        '!=': lambda a, b: a != b,
        '>': lambda a, b: a > b,
        '<': lambda a, b: a < b,
        '>=': lambda a, b: a >= b,
        '<=': lambda a, b: a <= b,
    }
    if operator not in numeric_ops:
        return False

    if operator in ('==', '!='):
        try:
            return numeric_ops[operator](_parse_compare_number(left), _parse_compare_number(right))
        except (TypeError, ValueError, InvalidOperation):
            return numeric_ops[operator](left, right)

    try:
        return numeric_ops[operator](_parse_compare_number(left), _parse_compare_number(right))
    except (TypeError, ValueError, InvalidOperation):
        return False

def _parse_compare_number(value):
    text = str(value).strip()
    if not text:
        raise ValueError("empty numeric value")
    if re.fullmatch(r'[+-]?\d+', text):
        return int(text)
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', text):
        raise ValueError(f"not a numeric value: {value}")
    number = Decimal(text)
    if not number.is_finite():
        raise ValueError(f"not a finite numeric value: {value}")
    return number

_CALC_OPERATORS = {
    ast.Add: operator_module.add,
    ast.Sub: operator_module.sub,
    ast.Mult: operator_module.mul,
    ast.Div: operator_module.truediv,
    ast.FloorDiv: operator_module.floordiv,
    ast.Mod: operator_module.mod,
}
_CALC_UNARY_OPERATORS = {
    ast.UAdd: operator_module.pos,
    ast.USub: operator_module.neg,
}
_CALC_MAX_ABS_VALUE = 1e18

def _validate_calc_number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError("Only numeric constants are supported")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("Non-finite numeric values are not supported")
    if abs(value) > _CALC_MAX_ABS_VALUE:
        raise OverflowError("Calculation result is too large")
    return value

def _raise_if_stop_requested(ctx, message="Stop requested"):
    if _ctx_check_stop(ctx):
        raise MacroStopException(message)

def _sleep_with_stop_check(seconds, ctx=None, interval=0.1):
    end_time = time.monotonic() + max(0.0, float(seconds or 0))
    while True:
        _raise_if_stop_requested(ctx, "Stop requested during wait")
        remaining = end_time - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(interval, remaining))


def _pause_after_pyautogui_call(ctx=None):
    pause = max(0.0, float(getattr(pyautogui, 'PAUSE', 0.0) or 0.0))
    if pause > 0:
        _sleep_with_stop_check(pause, ctx, interval=0.02)


def _move_to_with_stop(x, y, duration=0.0, ctx=None, pause_after=True):
    _raise_if_stop_requested(ctx, "Stop requested before mouse move")
    duration = max(0.0, float(duration or 0.0))
    if duration <= 0:
        pyautogui.moveTo(x, y, duration=0)
        _raise_if_stop_requested(ctx, "Stop requested after mouse move")
        return

    start_x, start_y = pyautogui.position()
    steps = max(1, int(math.ceil(duration / _MOUSE_MOVE_FRAME_INTERVAL)))
    step_duration = duration / steps
    for i in range(1, steps + 1):
        _raise_if_stop_requested(ctx, "Stop requested during mouse move")
        time.sleep(step_duration)
        _raise_if_stop_requested(ctx, "Stop requested during mouse move")
        ratio = i / steps
        nx = start_x + (x - start_x) * ratio
        ny = start_y + (y - start_y) * ratio
        pyautogui.moveTo(nx, ny, duration=0, _pause=False)
    if pause_after:
        _pause_after_pyautogui_call(ctx)

def _click_with_stop(x, y, button, clicks=1, interval=0.0, duration=0.0, ctx=None):
    _raise_if_stop_requested(ctx, "Stop requested before click")
    duration = max(0.0, float(duration or 0.0))
    interval = max(0.0, float(interval or 0.0))
    clicks = max(1, int(clicks or 1))

    if clicks == 1 and interval == 0.0 and duration == 0.0:
        pyautogui.click(x=x, y=y, button=button, clicks=1, interval=0.0, duration=0.0)
        _raise_if_stop_requested(ctx, "Stop requested after click")
        return

    if x is not None and y is not None:
        _move_to_with_stop(x, y, duration, ctx, pause_after=False)
    for index in range(clicks):
        _raise_if_stop_requested(ctx, "Stop requested during click")
        pyautogui.click(button=button, clicks=1, interval=0.0, duration=0.0, _pause=False)
        if index < clicks - 1 and interval > 0:
            _sleep_with_stop_check(interval, ctx, interval=0.02)
    _pause_after_pyautogui_call(ctx)
    _raise_if_stop_requested(ctx, "Stop requested after click")

# ======================================================================
# Expressions and value calculation
# ======================================================================
def _safe_calculate_expression(expression):
    def eval_node(node):
        if isinstance(node, ast.Expression):
            return eval_node(node.body)
        if isinstance(node, ast.Constant):
            return _validate_calc_number(node.value)
        if isinstance(node, ast.BinOp):
            op_func = _CALC_OPERATORS.get(type(node.op))
            if op_func is None:
                raise TypeError("Unsupported operator")
            return _validate_calc_number(op_func(eval_node(node.left), eval_node(node.right)))
        if isinstance(node, ast.UnaryOp):
            op_func = _CALC_UNARY_OPERATORS.get(type(node.op))
            if op_func is None:
                raise TypeError("Unsupported unary operator")
            return _validate_calc_number(op_func(eval_node(node.operand)))
        raise TypeError("Unsupported node")

    return eval_node(ast.parse(str(expression), mode='eval'))

_ACTION_DONE = "done"
_ACTION_SKIP = "skip"
_ACTION_BREAK = "break"


# ======================================================================
# Linear action handlers
# ======================================================================
def _handle_click_action(p, ctx):
    btn = p.get('button', 'left').lower()
    clicks, ok = _safe_param_int(p, 'clicks', 1, min_value=1)
    if not ok:
        _error_param_skip('CLICK', 'clicks', 'positive integer')
        return _ACTION_SKIP
    interval_ms, ok = _safe_param_float(p, 'interval', 0.0, min_value=0)
    if not ok:
        _warn_param_default('CLICK', 'interval', interval_ms)
    interval = interval_ms / 1000.0
    duration, ok = _safe_param_float(p, 'duration', 0.0, min_value=0)
    if not ok:
        _warn_param_default('CLICK', 'duration', duration)
    coords = _parse_optional_coords_or_break(p, 'CLICK')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    _click_with_stop(x, y, btn, clicks=clicks, interval=interval, duration=duration, ctx=ctx)
    if x is not None and y is not None:
        _ctx_set_last_pos(ctx, x, y)
    return _ACTION_DONE


def _handle_move_to_action(p, ctx):
    coords = _parse_required_coords_or_break(p, 'MOVE_TO')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
    if not ok:
        _warn_param_default('MOVE_TO', 'duration', duration)
    _move_to_with_stop(x, y, duration, ctx)
    _ctx_set_last_pos(ctx, x, y)
    return _ACTION_DONE


def _handle_move_offset_action(p, ctx):
    last_x, last_y = _ctx_get_last_pos(ctx)
    if last_x is None or last_y is None:
        logger.error("  MOVE_OFFSET has no last position")
        return _ACTION_BREAK
    offset = _parse_offset_or_break(p, 'MOVE_OFFSET')
    if offset == _ACTION_BREAK:
        return _ACTION_BREAK
    ox, oy = offset
    duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
    if not ok:
        _warn_param_default('MOVE_OFFSET', 'duration', duration)
    current_x, current_y = pyautogui.position()
    _move_to_with_stop(current_x + ox, current_y + oy, duration, ctx)
    _ctx_set_last_pos(ctx, last_x + ox, last_y + oy)
    return _ACTION_DONE


def _handle_scroll_action(p, ctx):
    clicks = _parse_int_or_break(p, 'SCROLL', 'amount', 0)
    if clicks == _ACTION_BREAK:
        return _ACTION_BREAK
    coords = _parse_optional_coords_or_break(p, 'SCROLL')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    if x is not None and y is not None:
        pyautogui.moveTo(x, y)
    pyautogui.scroll(clicks)
    return _ACTION_DONE


def _handle_wait_action(p, ctx):
    total_ms = _parse_positive_int_or_break(p, 'WAIT', 'ms', 0)
    if total_ms == _ACTION_BREAK:
        return _ACTION_SKIP
    _sleep_with_stop_check(total_ms / 1000.0, ctx)
    return _ACTION_DONE


def _handle_type_text_action(p, ctx):
    interval, ok = _safe_param_float(p, 'interval', 0.0, min_value=0)
    if not ok:
        _warn_param_default('TYPE_TEXT', 'interval', interval)
    text = _render_param(p, 'text', ctx)
    if not text:
        logger.warning("  TYPE_TEXT has no text; step skipped")
        return _ACTION_SKIP

    if '{CLIPBOARD}' in text:
        clipboard_content = _ctx_get_clipboard_var(ctx)
        if not clipboard_content:
            try:
                clipboard_content = pyperclip.paste()
            except Exception:
                clipboard_content = ''
        text = text.replace('{CLIPBOARD}', clipboard_content)
        logger.info("  replaced clipboard placeholder (%s chars)", len(text))

    if interval > 0:
        pyautogui.write(text, interval=interval)
    else:
        copy_success = _copy_to_clipboard_with_retry(text, ctx)
        if copy_success:
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
        else:
            logger.error("  clipboard copy failed; paste skipped")
    return _ACTION_DONE


def _handle_press_key_action(p, ctx):
    keys = [k for k in p.get('key', '').lower().replace(' ', '').split('+') if k]
    if keys:
        pyautogui.hotkey(*keys)
    return _ACTION_DONE



def _handle_activate_window_action(p, ctx):
    if not PYGETWINDOW_AVAILABLE:
        msg = "pygetwindow is unavailable; cannot activate target window"
        if p.get('ignore_fail', False):
            logger.warning(f"  {msg}")
            return _ACTION_SKIP
        raise RuntimeError(msg)
    title = p.get('title')
    if not title:
        msg = "ACTIVATE_WINDOW is missing a window title"
        if p.get('ignore_fail', False):
            logger.warning(f"  {msg}")
            return _ACTION_SKIP
        raise RuntimeError(msg)

    try:
        wins = gw.getWindowsWithTitle(title)
        if not wins:
            raise RuntimeError(f"No window title contains '{title}'")
        
        target_win = wins[0]
        if target_win.isMinimized:
            target_win.restore()
        target_win.activate()
        logger.info(f"  已激活窗口: {target_win.title}")
        for _ in range(5):
            if _ctx_check_stop(ctx):
                raise MacroStopException("用户在窗口激活期间请求停止")
            time.sleep(0.1) 
    except Exception as e:
        msg = str(e)
        if p.get('ignore_fail', False) and not msg.startswith("No window title contains"):
            logger.warning(f"  ACTIVATE_WINDOW failed and was ignored: {e}")
            return _ACTION_SKIP
        raise RuntimeError(f"ACTIVATE_WINDOW failed: {e}") from e

    return _ACTION_DONE

def _handle_note_action(p, ctx):
    note_text = p.get('text', '')
    if note_text:
        logger.info("  备注步骤已执行 (%s chars)", len(note_text))
    return _ACTION_SKIP


def _handle_set_var_action(p, ctx):
    var_name = p.get('var_name', '').strip()
    var_value = _render_param(p, 'var_value', ctx)
    if var_name:
        _ctx_set_var(ctx, var_name, var_value)
        logger.info("  %s 已设置: %s", var_name, _log_value_summary(var_value))
    return _ACTION_DONE



def _handle_prompt_input_action(p, ctx):
    var_name = p.get('var_name', '').strip()
    title = _render_param(p, 'title', ctx, '智点助手输入')
    prompt = _render_param(p, 'prompt', ctx, '请输入内容:')
    default_value = _render_param(p, 'default_value', ctx)
    if not var_name:
        logger.error("  PROMPT_INPUT 缺少变量名")
        return _ACTION_BREAK
    if _ctx_check_stop(ctx):
        raise MacroStopException("用户在输入前请求停止")
    callback = ctx.get('prompt_input_callback')
    try:
        if callable(callback):
            value = callback(title, prompt, default_value, ctx)
        else:
            suffix = f" [{default_value}]" if default_value else ""
            raw = input(f"{title} - {prompt}{suffix}: ")
            value = default_value if raw == '' and default_value else raw
    except KeyboardInterrupt as e:
        raise MacroStopException("用户取消输入") from e
    if value is None:
        raise MacroStopException("用户取消输入")
    _ctx_set_var(ctx, var_name, str(value))
    logger.info("  %s 已接收用户输入: %s", var_name, _log_value_summary(_ctx_get_var(ctx, var_name)))
         
    return _ACTION_DONE

def _handle_read_file_action(p, ctx):
    file_path = _render_param(p, 'file_path', ctx)
    var_name = p.get('var_name', '').strip()
    encoding = p.get('encoding', 'utf-8')
    if not file_path or not var_name:
        logger.error("  READ_FILE 参数不完整")
        return _ACTION_BREAK

    def _do_read():
        safe_path = _resolve_safe_file_path(file_path, ctx, 'READ_FILE')
        with open(safe_path, 'r', encoding=encoding) as f:
            _ctx_set_var(ctx, var_name, f.read())
        logger.info(f"  {var_name} 已读取文本文件 ({len(_ctx_get_var(ctx, var_name))} 字符)")
        return _ACTION_DONE

    return _run_with_fail_stop('READ_FILE', p, _do_read, ctx, var_name)

def _handle_extract_var_action(p, ctx):
    source = _render_param(p, 'source_text', ctx)
    pattern = str(p.get('regex', ''))
    var_name = p.get('var_name', '').strip()
    if not pattern or not var_name:
        logger.error("  EXTRACT_VAR 参数不完整")
        return _ACTION_BREAK

    def _do_extract():
        match = re.search(pattern, source)
        val = match.group(1) if match and match.lastindex else (match.group(0) if match else '')
        _ctx_set_var(ctx, var_name, val)
        logger.info("  %s 已提取: %s", var_name, _log_value_summary(val))
        return _ACTION_DONE

    return _run_with_fail_stop('EXTRACT_VAR', p, _do_extract, ctx, var_name)

def _handle_write_file_action(p, ctx):
    file_path = _render_param(p, 'file_path', ctx)
    content = _render_param(p, 'content', ctx)
    encoding = p.get('encoding', 'utf-8')
    append = _parse_bool_param(p.get('append', False), False)
    if not file_path:
        logger.error("  WRITE_FILE 未指定路径")
        return _ACTION_BREAK

    def _do_write():
        mode = 'a' if append else 'w'
        safe_path = _resolve_safe_file_path(file_path, ctx, 'WRITE_FILE')
        _ensure_text_output_path(safe_path, 'WRITE_FILE')
        dir_path = os.path.dirname(safe_path)
        if dir_path and not os.path.exists(dir_path):
            os.makedirs(dir_path, exist_ok=True)
        with open(safe_path, mode, encoding=encoding) as f:
            f.write(content)
        logger.info("  写入路径=%s (追加: %s)", _log_display_text(_log_display_path(safe_path, ctx)), append)
        return _ACTION_DONE

    return _run_with_fail_stop('WRITE_FILE', p, _do_write)


def _resolve_goto_step(ctx, pc, goto_labels, action_name, label, max_jumps_value, warn_on_default=False):
    if not label:
        raise RuntimeError(f"{action_name} 缺少标签名")
    target = goto_labels.get(_normalize_goto_label(label))
    if not target:
        raise RuntimeError(f"{action_name} 找不到标签: {label}")
    max_jumps, ok = _safe_int(
        max_jumps_value,
        GOTO_LABEL_DEFAULT_MAX_JUMPS,
        min_value=1
    )
    if warn_on_default and not ok:
        _warn_param_default(action_name, 'max_jumps', max_jumps)
    goto_counts = ctx.setdefault('_goto_counts', {})
    count = goto_counts.get(pc, 0)
    if count >= max_jumps:
        raise RuntimeError(f"{action_name} 超过最大跳转次数 {max_jumps}: {label}")
    goto_counts[pc] = count + 1
    return target


def _parse_run_timeout(p):
    timeout, ok = _safe_param_int(p, 'timeout', 30, min_value=1)
    if not ok:
        logger.warning("  RUN 参数 'timeout' 必须是正整数，已使用默认值 30")
    return timeout

def _handle_run_action(p, ctx):
    run_result = _handle_run(p, ctx)
    if run_result == 'SKIPPED':
        logger.info("  已跳过，继续执行后续步骤")
        return _ACTION_SKIP
    if run_result is False:
        logger.error("  执行失败")
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_DONE
    logger.info("  执行成功")
    return _ACTION_DONE


def _handle_calculate_action(p, ctx):
    expr = _render_param(p, 'expression', ctx)
    var_name = p.get('var_name', '').strip()
    if not expr or not var_name:
        logger.error("  CALCULATE missing expression or var_name")
        return _ACTION_BREAK
    try:
        val = _safe_calculate_expression(expr)
        _ctx_set_var(ctx, var_name, str(val))
        logger.info("  %s 已计算: %s", var_name, _log_value_summary(val))
    except (ValueError, TypeError, OverflowError, SyntaxError, ZeroDivisionError) as e:
        logger.error("  CALCULATE failed (expression length: %s): %s", len(expr), e)
        return _ACTION_SKIP
    return _ACTION_DONE

# ======================================================================
# Linear action dispatch
# ======================================================================
_LINEAR_ACTION_HANDLERS = {
    'CLICK': _handle_click_action,
    'MOVE_TO': _handle_move_to_action,
    'MOVE_OFFSET': _handle_move_offset_action,
    'SCROLL': _handle_scroll_action,
    'WAIT': _handle_wait_action,
    'TYPE_TEXT': _handle_type_text_action,
    'PRESS_KEY': _handle_press_key_action,
    'ACTIVATE_WINDOW': _handle_activate_window_action,
    'NOTE': _handle_note_action,
    'SET_VAR': _handle_set_var_action,
    'READ_FILE': _handle_read_file_action,
    'EXTRACT_VAR': _handle_extract_var_action,
    'CALCULATE': _handle_calculate_action,
    'WRITE_FILE': _handle_write_file_action,
    'PROMPT_INPUT': _handle_prompt_input_action,
    'RUN': _handle_run_action,
}


def _dispatch_linear_action(act, p, ctx):
    handler = _LINEAR_ACTION_HANDLERS.get(act)
    if handler is None:
        return None
    return handler(p, ctx)


def _handle_if_var_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    var_val = _render_param(p, 'var_value', ctx)
    op = p.get('operator', '==')
    expected = _render_param(p, 'expected_val', ctx)
    res_bool = _compare_values(var_val, op, expected)

    if not res_bool:
        logger.info("  -> IF条件不满足 (%s)", op)
        return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])

    logger.info("  -> IF条件满足 (%s)", op)
    return pc + 1


def _handle_goto_if_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if loops:
        raise RuntimeError("GOTO_IF 当前版本不允许在 LOOP 循环内部执行")

    var_val = _render_param(p, 'var_value', ctx)
    operator = p.get('operator', '==')
    expected = _render_param(p, 'expected_val', ctx)
    label = _render_param_stripped(p, 'label', ctx)
    res_bool = _compare_values(var_val, operator, expected)

    if res_bool:
        target = _resolve_goto_step(ctx, pc, goto_labels, 'GOTO_IF', label, p.get('max_jumps', GOTO_LABEL_DEFAULT_MAX_JUMPS))
        logger.info("  条件成立 (%s)，跳至 '%s'", operator, target['name'])
        return target['index']

    logger.info("  -> GOTO_IF 条件不满足 (%s)", operator)
    return pc + 1


def _handle_goto_label_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if loops:
        raise RuntimeError("GOTO_LABEL 当前版本不允许在 LOOP 循环内部执行")

    label = _render_param_stripped(p, 'label', ctx)
    target = _resolve_goto_step(
        ctx,
        pc,
        goto_labels,
        'GOTO_LABEL',
        label,
        p.get('max_jumps', GOTO_LABEL_DEFAULT_MAX_JUMPS),
        warn_on_default=True
    )
    next_pc = target['index']
    logger.info(f"  跳转到标签 '{target['name']}' -> 第 {next_pc + 1} 步")
    return next_pc


def _handle_else_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['END_IF'])


def _handle_foreach_line_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _handle_foreach_line_start(steps, pc, loops, p, ctx, status_callback)


def _handle_loop_start_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _handle_loop_start(steps, pc, loops, p, ctx, status_callback)


def _exit_active_loop(loops, ctx):
    """Pop the active loop and release its per-loop cache."""
    loops.pop()
    _get_loop_cache(ctx).exit()


def _handle_end_loop_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if not loops:
        logger.error("END_LOOP 缺少对应的 LOOP_START")
        return pc + 1

    top = loops[-1]
    if top.get('kind') == 'foreach_line':
        raise RuntimeError("END_LOOP 不能结束批量处理，请使用 END_FOREACH")

    mode = top.get('mode', 'fixed')
    if mode not in ('until_image', 'until_text'):
        return top['start']

    top['iteration'] += 1
    if status_callback:
        status_callback(f"🔄 循环第 {top['iteration']} 次 (最多 {top['max_iterations']} 次)")

    if top['iteration'] >= top['max_iterations']:
        _exit_active_loop(loops, ctx)
        if status_callback:
            status_callback(f"WARN 达到最大迭代 {top['max_iterations']} 次,强制退出")
        logger.warning("  达到最大迭代次数,强制退出")
        return pc + 1

    condition_met = _check_loop_condition(top, ctx)
    if condition_met:
        _exit_active_loop(loops, ctx)
        if status_callback:
            status_callback(f"OK 条件满足,循环结束 (共 {top['iteration']} 次)")
        logger.info("  条件满足,循环结束")
        return pc + 1

    logger.debug("  未找到目标,继续循环 (第 %s 次)", top['iteration'])
    time.sleep(LOOP_CHECK_INTERVAL)
    return top['start']


def _handle_end_foreach_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if loops and loops[-1].get('kind') == 'foreach_line':
        return loops[-1]['start']
    if loops:
        raise RuntimeError("END_FOREACH 不能结束普通循环，请使用 END_LOOP")

    logger.error("END_FOREACH 缺少对应的批量处理开始")
    return pc + 1


# ======================================================================
# Control-flow dispatch
# ======================================================================
_CONTROL_ACTION_HANDLERS = {
    'FOREACH_LINE': _handle_foreach_line_control,
    'IF_VAR': _handle_if_var_control,
    'GOTO_IF': _handle_goto_if_control,
    'GOTO_LABEL': _handle_goto_label_control,
    'ELSE': _handle_else_control,
    'LOOP_START': _handle_loop_start_control,
    'END_LOOP': _handle_end_loop_control,
    'END_FOREACH': _handle_end_foreach_control,
}


def _dispatch_control_action(act, steps, pc, loops, p, ctx, status_callback, goto_labels):
    handler = _CONTROL_ACTION_HANDLERS.get(act)
    if handler is None:
        return None
    return handler(steps, pc, loops, p, ctx, status_callback, goto_labels)


_EXECUTE_STEP_LOG_ACTIONS = {'LOOP_START', 'END_LOOP', 'FOREACH_LINE', 'END_FOREACH', 'ELSE', 'END_IF', 'RUN', 'NOTE', 'GOTO_LABEL'}


def _get_step_action_and_params(step):
    return step.get('action', ''), step.get('params', {})


def _should_log_step(act, loops, ctx):
    return ctx.get('debug_steps', False) or not loops or act in _EXECUTE_STEP_LOG_ACTIONS


def _is_find_or_condition_action(act):
    return (act.startswith('FIND_') or act.startswith('IF_')) and act != 'IF_VAR'


def _should_skip_disabled_step(step, act):
    return not step.get('enabled', True) and act not in MacroSchema.CONTROL_FLOW_ACTIONS


def _handle_ai_command_step(p, ctx):
    instruction = p.get('instruction', '')
    if not instruction:
        logger.error("  AI 指令为空")
        return 'break'

    region = None
    if p.get('region') is not None:
        bbox = _coerce_bbox(p.get('region'))
        if bbox is None:
            raise ValueError(f"AI 手动区域解析失败(格式错误): {p.get('region')}")
        region = tuple(bbox)
    elif p.get('cache_box') is not None:
        bbox = _coerce_bbox(p.get('cache_box'))
        if bbox:
            region = tuple(bbox)

    logger.info("  执行 AI 指令 (%s chars)", len(instruction))
    coords = _call_vlm_with_stop(instruction, region, ctx)

    if coords:
        target_x, target_y = coords
        logger.info(f"  返回坐标: ({target_x}, {target_y})")
        duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
        if not ok:
            _warn_param_default('AI_COMMAND', 'duration', duration)
        _move_to_with_stop(target_x, target_y, duration, ctx)
        _ctx_set_last_pos(ctx, target_x, target_y)
        return 'done'

    logger.info("  未找到目标位置")
    if p.get('fail_stop', True):
        return 'break'
    return 'done'


def _handle_find_or_condition_step(steps, pc, act, p, ctx):
    next_pc = pc + 1
    res = _handle_find_with_retries(act, p, ctx, _get_loop_cache(ctx).get_current_loop_id() is not None)

    if act.startswith('IF_'):
        if not res:
            logger.info("  -> IF condition not met, skipping block")
            next_pc = _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])
    elif not res:
        if p.get('ignore_fail', False):
            logger.info("  -> Target not found, skipped by ignore_fail")
            return next_pc, 'continue'
        logger.info("  -> Target not found, macro stopped")
        return next_pc, 'break'

    # FIND_ and IF_ both move the mouse after a match to preserve stable behavior.
    # A CLICK without coordinates inside an IF block can use the current mouse position.
    if res:
        target_x, target_y = res[0], res[1]
        _ctx_set_last_pos(ctx, target_x, target_y)
        pyautogui.moveTo(target_x, target_y)

    return next_pc, 'done'


# ======================================================================
# Core execution loop
# ======================================================================
def execute_steps(steps, run_context=None, status_callback=None):
    logger.info(f"\n--- 执行开始 (Core V{CORE_VERSION}) ---")
    perf.reset()
    loop_cache.reset()
    ctx = RunContext(run_context)
    ctx['_goto_counts'] = {}
    try:
        goto_labels = _build_goto_label_table(steps)
    except ValueError as e:
        logger.error(f"  标签配置错误: {e}")
        return False
    
    default_stop = "Ctrl+F2"
    try:
        s = ctx.get('stop_key_str', default_stop)
        stop_key_display = HotkeyUtils.format_hotkey_display(s)
    except Exception:
        stop_key_display = default_stop
    
    pc, loops = 0, []
    try:
        while pc < len(steps):

            if _ctx_check_stop(ctx): 
                logger.info(f"  用户请求停止 ({stop_key_display})")
                break
                
            step = steps[pc]
            ctx['_current_step_index'] = pc
            act, p = _get_step_action_and_params(step)
            if _should_log_step(act, loops, ctx):
                logger.info(f"[{pc+1}] {act}")
            next_pc = pc + 1

            # [新增] 处理被屏蔽的普通步骤
            if _should_skip_disabled_step(step, act):
                logger.info(f"  跳过步骤: {act}")
                pc = next_pc
                continue

            try:
                if _is_find_or_condition_action(act):
                    next_pc, flow = _handle_find_or_condition_step(steps, pc, act, p, ctx)
                    if flow == 'continue':
                        pc = next_pc
                        continue
                    if flow == 'break':
                        break

                elif act == 'AI_COMMAND':
                    if _handle_ai_command_step(p, ctx) == 'break':
                        break

                elif act in _LINEAR_ACTION_HANDLERS:
                    action_result = _dispatch_linear_action(act, p, ctx)
                    if action_result == _ACTION_BREAK:
                        break
                    if action_result == _ACTION_SKIP:
                        pc = next_pc
                        continue

                elif act in _CONTROL_ACTION_HANDLERS:
                    next_pc = _dispatch_control_action(act, steps, pc, loops, p, ctx, status_callback, goto_labels)

                elif act == 'END_IF':
                    pass

                else:
                    logger.warning(f"  未知动作: {act}")

            except pyautogui.FailSafeException as e:
                raise MacroStopException("PyAutoGUI failsafe triggered") from e
            except MacroStopException:
                raise  # 向上传播，不要吞掉
            except LoopConditionCheckError:
                raise
            except Exception as e:
                error_msg = f"  [执行异常] 步骤 {pc+1} ({act}): {e}"
                logger.error(error_msg)
                if status_callback:
                    status_callback(f"ERR {error_msg}")
                break
            pc = next_pc
        
        return pc >= len(steps)
    finally:
        cleanup_active_processes(ctx)
        ctx.loop_cache.reset()
        logger.info(f"--- 执行结束 ---\n[统计] {ctx.perf.get_stats()}\n")

def _handle_find_with_retries(act, p, ctx, in_loop):
    retry_count, ok = _safe_param_int(p, 'retry_count', 0, min_value=0)
    if not ok:
        _warn_param_default(act, 'retry_count', retry_count)
    retry_interval, ok = _safe_param_float(p, 'retry_interval', 0, min_value=0)
    if not ok:
        _warn_param_default(act, 'retry_interval', retry_interval)

    for attempt in range(retry_count + 1):
        _raise_if_stop_requested(ctx, "Stop requested before find retry")
        res = _handle_find(act, p, ctx, in_loop)
        if res or attempt >= retry_count:
            return res
        if retry_interval > 0:
            logger.debug("  %s not found; retrying in %ss (%s/%s)", act, retry_interval, attempt + 1, retry_count)
            _sleep_with_stop_check(retry_interval, ctx)
        else:
            logger.debug("  %s not found; retrying immediately (%s/%s)", act, attempt + 1, retry_count)
    return None


def _handle_find(act, p, ctx, in_loop):
    _raise_if_stop_requested(ctx, "Stop requested before find")
    is_img = 'IMAGE' in act
    final_engine = FORCE_OCR_ENGINE if (FORCE_OCR_ENGINE and FORCE_OCR_ENGINE != 'auto') else p.get('engine', 'auto')
    sig = f"{act}_{p.get('path', p.get('text',''))}"
    
    region = None
    is_manual_region = False
    runtime_cache_boxes = ctx.setdefault('_runtime_cache_boxes', {})
    if p.get('region') is not None:
        region = bbox_to_region(p.get('region'))
        if region is None:
            raise ValueError(f"手动查找区域解析失败(格式错误): {p.get('region')}")
        is_manual_region = True
    if not region:
        cb_raw = runtime_cache_boxes.get(sig, p.get('cache_box'))
        region = _padded_bbox_to_region(cb_raw, CACHE_BOX_PADDING)

    # 引入 try...finally 结构以保证 ss 得到显式释放
    ss = None
    try:
        ss, offset = smart_screenshot(region)
        _raise_if_stop_requested(ctx, "Stop requested after screenshot")

        # 增强模式：IF 动作和多缩放尝试（性能开销大，但更准确）
        enhanced_mode = ctx.get('enhanced_mode', False)

        if in_loop:
            cached = _get_loop_cache(ctx).get(sig)
            if cached and is_img:
                confidence, ok = _safe_param_float(p, 'confidence', 0.8, min_value=0, max_value=1)
                if not ok:
                    _warn_param_default(act, 'confidence', confidence)
                if quick_check_cv2(p.get('path',''), confidence, ss, offset, cached, enhanced_mode):
                    _get_perf(ctx).record_hit(True, False); logger.info(f"  {cached}"); _ctx_set_last_pos(ctx, cached); return cached

        ocr_cache_key = f"step:{ctx.get('_current_step_index', '?')}:{act}:{p.get('text', '')}:{p.get('region', p.get('cache_box', ''))}" if not is_img else None
        res = _do_find(is_img, p, ss, offset, final_engine, ctx, ocr_cache_key)
        _raise_if_stop_requested(ctx, "Stop requested after find")

        # IF 动作不使用全局 fallback（会大幅降低性能，因为它已经重复搜索了）
        # FIND_TEXT/FIND_IMAGE 才使用 fallback
        if not res and region and ENABLE_GLOBAL_FALLBACK and enhanced_mode and not act.startswith('IF_') and not is_manual_region:
            logger.info("  全局搜索...")
            # 释放旧的 ss，防止覆盖导致句柄丢失泄漏
            _close_image_quietly(ss)
            ss = None
            ss, offset = smart_screenshot(None)
            _raise_if_stop_requested(ctx, "Stop requested after fallback screenshot")
            res = _do_find(is_img, p, ss, offset, final_engine, ctx, ocr_cache_key)
            _raise_if_stop_requested(ctx, "Stop requested after fallback find")
            if res:
                if len(res) >= 2:
                    runtime_cache_boxes[sig] = [res[0]-20, res[1]-10, res[0]+20, res[1]+10]

        if res:
            pos = (res[0], res[1])
            if in_loop: _get_loop_cache(ctx).set(sig, pos)
            _ctx_set_last_pos(ctx, pos)
            if act in ('FIND_TEXT', 'IF_TEXT_FOUND') and len(res) >= 3:
                save_var = p.get('save_to_var', '').strip()
                if save_var:
                    _ctx_set_var(ctx, save_var, res[2])
                    logger.info("  %s 已保存 OCR 结果: %s", save_var, _log_value_summary(res[2]))
            return res

        _get_perf(ctx).record_miss(not is_img)
        return None
    finally:
        _close_image_quietly(ss)

def _do_find(is_img, p, ss, offset, engine='auto', ctx=None, ocr_cache_key=None):
    """执行查找（图像或文本）并返回统一格式坐标 (x, y)"""
    enhanced_mode = ctx.get('enhanced_mode', False) if ctx else False
    if is_img:
        # 图片查找返回: (cx, cy, w, h)
        confidence, ok = _safe_param_float(p, 'confidence', 0.8, min_value=0, max_value=1)
        if not ok:
            _warn_param_default('FIND_IMAGE', 'confidence', confidence)
        res_val = find_image_cv2(p.get('path',''), confidence, ss, offset, enhanced_mode, ctx)
        if res_val:
            _get_perf(ctx).record_hit(False, False)
            logger.info("  图 (%s,%s)，模板=%s", res_val[0][0], res_val[0][1], _log_display_text(os.path.basename(p.get('path', ''))))
            return (res_val[0][0], res_val[0][1]) 
    else:
        # OCR 查找返回: ((cx, cy), full_text)
        res = ocr_engine.find_text_location(
            p.get('text',''), 
            p.get('lang','eng'), 
            p.get('debug', False), 
            ss, offset, engine, enhanced_mode, cache_key=ocr_cache_key
        )
        
        if res:
            _get_perf(ctx).record_hit(False, True)
            
            # === [修复] 统一返回格式为扁平元组: (x, y, text) ===
            pos = (0, 0)
            text_content = ""

            # 解析 ocr_engine 的返回值
            if isinstance(res, tuple) and len(res) == 2:
                if isinstance(res[0], tuple) and len(res[0]) >= 2:
                    # 新格式: ((x, y), full_text)
                    pos = res[0]
                    text_content = res[1]
                else:
                    # 旧格式兼容: (x, y)
                    pos = res
                    text_content = p.get('text', '')
            else:
                pos = res
                text_content = p.get('text', '')

            # 打印调试信息
            logger.info("  文 (%s,%s)，目标=%s，识别原文=%s", pos[0], pos[1], _log_display_text(p.get('text', '')), _log_display_text(text_content))

            final_text = text_content
            extract_pattern = p.get('extract_pattern', '').strip()
            if extract_pattern and ctx and (p.get('save_to_clipboard', False) or p.get('save_to_var', '').strip()):
                try:
                    match = re.search(extract_pattern, text_content)
                    if match:
                        if match.lastindex:
                            final_text = match.group(1)
                        else:
                            final_text = match.group(0)
                        logger.info("  文本转换完成: %s", _log_display_text(final_text))
                    else:
                        logger.info("  未匹配，保留原文")
                except Exception as e:
                    logger.error(f"  {e}")

            # 处理剪贴板逻辑 (副作用)
            if ctx and p.get('save_to_clipboard', False):
                logger.info("  保留原始文本: %s", _log_value_summary(text_content))
                _ctx_set_clipboard_var(ctx, final_text)
                try:
                    if not _copy_to_clipboard_with_retry(final_text, ctx):
                        raise RuntimeError("clipboard is busy")
                    logger.info("  OK 已复制")
                except Exception as e:
                    logger.error(f"  失败: {e}")
            
            return (pos[0], pos[1], final_text)
    
    return None


def _handle_foreach_line_start(steps, pc, loops, p, ctx, cb):
    top = loops[-1] if loops else None

    if top and top.get('kind') == 'foreach_line' and top.get('start') == pc:
        time.sleep(LOOP_PHYSICAL_COOLDOWN)
        next_index = top.get('index', 0) + 1
        if next_index >= len(top.get('items', [])):
            loops.pop()
            if cb:
                cb("批量处理完成")
            logger.info("  所有行已处理完成")
            return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])

        top['index'] = next_index
        _set_foreach_line_vars(top, ctx)
        if cb:
            cb(f"批量处理第 {next_index + 1}/{len(top['items'])} 行")
        logger.info(f"  第 {next_index + 1}/{len(top['items'])} 行")
        return pc + 1

    if loops:
        raise RuntimeError("批量处理暂不支持嵌套在其他循环内部")

    file_path = _render_param_stripped(p, 'file_path', ctx)
    source_text = _render_param(p, 'source_text', ctx)
    encoding = p.get('encoding', 'utf-8')
    max_lines, ok = _safe_param_int(p, 'max_lines', 10000, min_value=1)
    if not ok:
        _warn_param_default('FOREACH_LINE', 'max_lines', max_lines)
    skip_empty = _parse_bool_param(p.get('skip_empty', True), True)

    if file_path:
        safe_path = _resolve_safe_file_path(file_path, ctx, 'FOREACH_LINE')
        lines = []
        with open(safe_path, 'r', encoding=encoding) as f:
            for raw_line in f:
                line = raw_line.rstrip('\r\n')
                if skip_empty and not line.strip():
                    continue
                lines.append(line)
                if len(lines) > max_lines:
                    raise RuntimeError(f"FOREACH_LINE exceeds max_lines {max_lines}")
    else:
        if source_text is None or source_text == '':
            logger.info("  data is empty, block skipped")
            return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])
        lines = source_text.splitlines()
        if skip_empty:
            lines = [line for line in lines if line.strip()]
        if len(lines) > max_lines:
            raise RuntimeError(f"FOREACH_LINE line count {len(lines)} exceeds max_lines {max_lines}")

    if not lines:
        logger.info("  没有可处理的行，跳过批量处理块")
        return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])

    _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])

    loop_data = {
        'kind': 'foreach_line',
        'start': pc,
        'index': 0,
        'items': lines,
        'current_line_var': str(p.get('current_line_var', 'current_line')).strip() or 'current_line',
        'index_var': str(p.get('index_var', 'loop_index')).strip() or 'loop_index',
        'total_var': str(p.get('total_var', 'loop_total')).strip() or 'loop_total',
        'delimiter': _normalize_split_delimiter(_render_param(p, 'split_delimiter', ctx)),
        'field_names': _split_field_names(p.get('field_names', '')),
        'strip_fields': _parse_bool_param(p.get('strip_fields', True), True),
    }
    loops.append(loop_data)
    _set_foreach_line_vars(loop_data, ctx)
    if cb:
        cb(f"批量处理第 1/{len(lines)} 行")
    logger.info(f"  开始处理 {len(lines)} 行")
    return pc + 1


def _handle_loop_start(steps, pc, loops, p, ctx, cb):
    top = loops[-1] if loops else None
    
    
    # 如果是已有循环的迭代检查
    if top and top['start'] == pc:
         # === [修复] 强制给循环加一个物理冷却，防止队列瞬间爆炸 ===
        time.sleep(LOOP_PHYSICAL_COOLDOWN)  # 使用常量 
        mode = top.get('mode', 'fixed')
        
        # 检查是否超过最大迭代次数 (所有模式通用)
        if top['iteration'] >= top['max_iterations']:
            _exit_active_loop(loops, ctx)
            if cb: cb(f"达到最大迭代 {top['max_iterations']} 次,循环结束")
            logger.warning(f"  警告:达到最大迭代次数 {top['max_iterations']}")
            return _find_jump_or_raise(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
        
        # 固定次数循环:检查剩余次数
        if mode == 'fixed':
            # [修复BUG-4] iteration 从初始化的 1 开始，每次回到 LOOP_START 时递增
            if top['remain'] > 0:
                top['remain'] -= 1
                top['iteration'] += 1
                total = top.get('total_count', top['iteration'] + top['remain'])
                if cb: cb(f"循环第 {top['iteration']} 次 (共 {total} 次)")
                return pc + 1
            else:
                _exit_active_loop(loops, ctx)
                return _find_jump_or_raise(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
        
        # === 关键修复: 条件循环不在此增加计数,交给 END_LOOP ===
        # 条件循环的迭代计数和退出判断统一在 END_LOOP 处理
        return pc + 1

    if loops and top and top.get('kind') == 'foreach_line':
        raise RuntimeError("普通循环暂不支持嵌套在批量处理内部")
    
    # 新循环初始化
    else:
        mode = p.get('mode', 'fixed')
        max_iter, ok = _safe_param_int(p, 'max_iterations', 1000, min_value=1)
        if not ok:
            _warn_param_default('LOOP_START', 'max_iterations', max_iter)
        
        if mode == 'fixed':
            count, ok = _safe_param_int(p, 'times', 1, min_value=1)
            if not ok:
                _error_param_skip('LOOP_START', 'times', 'positive integer')
                return _find_jump_or_raise(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
            remain = count - 1
        else:
            count = max_iter  # 条件循环用 max_iter 作为参考
            remain = max_iter
        
        loop_id = f"L{pc}_{len(loops)}"
        loop_data = {
            'start': pc,
            'remain': remain,
            'id': loop_id,
            'mode': mode,
            'iteration': 0 if mode in ('until_image', 'until_text') else 1,
            'total_count': count,  # [修复BUG-4] 保存总次数，供状态显示
            'max_iterations': max_iter
        }
        if p.get('region') is not None:
            loop_data['region'] = p['region']
        
        # 保存条件参数
        if mode == 'until_image':
            loop_data['condition_image'] = p.get('condition_image', '')
            confidence, ok = _safe_param_float(p, 'confidence', 0.8, min_value=0, max_value=1)
            if not ok:
                _warn_param_default('LOOP_START', 'confidence', confidence)
            loop_data['confidence'] = confidence
            # [优化] 保存搜索区域，加速条件检测
            if 'region' in p:
                logger.info("  目标模板=%s (区域: %s)", _log_display_text(_log_display_path(loop_data['condition_image'], ctx, basename_only=True)), p['region'])
            elif 'cache_box' in p:
                loop_data['cache_box'] = p['cache_box']
                logger.info("  目标模板=%s (区域: %s)", _log_display_text(_log_display_path(loop_data['condition_image'], ctx, basename_only=True)), p['cache_box'])
            else:
                logger.info("  目标模板=%s (全屏)", _log_display_text(_log_display_path(loop_data['condition_image'], ctx, basename_only=True)))
        elif mode == 'until_text':
            loop_data['condition_text'] = p.get('condition_text', '')
            loop_data['lang'] = p.get('lang', 'eng')
            # [优化] 保存搜索区域，加速条件检测
            if 'region' in p:
                logger.info(f"  目标: {loop_data['condition_text']} (区域: {p['region']})")
            elif 'cache_box' in p:
                loop_data['cache_box'] = p['cache_box']
                logger.info(f"  目标: {loop_data['condition_text']} (区域: {p['cache_box']})")
            else:
                logger.info(f"  目标: {loop_data['condition_text']} (全屏)")
        
        loops.append(loop_data)
        _get_loop_cache(ctx).enter(loop_id)
        
        if mode == 'fixed':
            if cb: cb(f"循环第 1 次 (共 {count} 次)")
        else:
            if cb: cb(f"🔄 条件循环第 1 次 (最多 {max_iter} 次)")
        
        return pc + 1

def _find_jump(steps, start, open_tag, close_tag, targets):
    lvl = 0
    for i in range(start + 1, len(steps)):
        a = steps[i].get('action','')
        if a.startswith(open_tag.rstrip('_')): lvl += 1
        elif a == close_tag:
            if lvl == 0 and a in targets: return i + 1
            lvl -= 1
        elif lvl == 0 and a in targets: return i + 1
    return -1

def _find_jump_or_raise(steps, start, open_tag, close_tag, targets):
    target = _find_jump(steps, start, open_tag, close_tag, targets)
    if target < 0:
        action = steps[start].get('action', '') if 0 <= start < len(steps) else ''
        raise RuntimeError(f"Control-flow structure is incomplete: step {start + 1} {action} cannot find {targets}")
    return target

def _check_loop_condition(loop_data, ctx):
    """检查循环退出条件是否满足
    
    返回值:
    - True: 找到了目标(应该退出循环)
    - False: 没找到(应该继续循环)
    """
    mode = loop_data.get('mode', 'fixed')
    
    # [优化] 统一构建截图区域（支持 cache_box 缩小截图范围）
    def _build_region(ld):
        if ld.get('region') is not None:
            region = bbox_to_region(ld.get('region'))
            if region is None:
                raise LoopConditionCheckError(f"Loop region is invalid: {ld.get('region')}")
            return region
        return _padded_bbox_to_region(ld.get('cache_box'), CACHE_BOX_PADDING)

    if mode == 'until_image':
        path = loop_data.get('condition_image', '')
        conf = loop_data.get('confidence', 0.8)
        
        if not path or not os.path.exists(path):
            logger.warning("  警告: 图像路径无效 %s", _log_display_text(_log_display_path(path, ctx, basename_only=True)))
            return False
        
        ss = None
        try:
            region = _build_region(loop_data)
            ss, offset = smart_screenshot(region)
            _raise_if_stop_requested(ctx, "Stop requested after loop screenshot")
            enhanced_mode = ctx.get('enhanced_mode', False) if ctx else False
            res_val = find_image_cv2(path, conf, ss, offset=offset, enhanced_mode=enhanced_mode, ctx=ctx)
            found = res_val is not None
            if found:
                logger.info(f"  OK 找到目标图像: {os.path.basename(path)}")
            return found
        except (ValueError, TypeError, AttributeError, IndexError) + ((cv2.error,) if OPENCV_AVAILABLE else ()) as e:
            logger.error(f"  图像检测错误: {e}")
            return False
        except Exception as e:
            logger.error(f"  严重错误 (退出循环): {e}")
            raise LoopConditionCheckError("Image loop condition check failed") from e
        finally:
            _close_image_quietly(ss)
    
    elif mode == 'until_text':
        text = loop_data.get('condition_text', '')
        lang = loop_data.get('lang', 'eng')
        
        if not text:
            logger.warning("  警告: 文本条件为空")
            return False
        
        ss = None
        try:
            region = _build_region(loop_data)
            ss, offset = smart_screenshot(region)
            _raise_if_stop_requested(ctx, "Stop requested after loop screenshot")
            enhanced_mode = ctx.get('enhanced_mode', False) if ctx else False
            cache_key = f"loop_until_text:{loop_data.get('start', '?')}:{text}:{loop_data.get('region', loop_data.get('cache_box', ''))}"
            res = ocr_engine.find_text_location(text, lang, False, ss, offset, 'auto', enhanced_mode, cache_key=cache_key)
            
            if res:
                found_txt = text
                if isinstance(res, tuple) and len(res) == 2 and isinstance(res[1], str):
                    found_txt = res[1]
                logger.info("  OK 找到目标文本: %s", _log_value_summary(found_txt))
                return True
            return False
        except (ValueError, TypeError, AttributeError) as e:
            logger.error(f"  文本检测错误: {e}")
            return False
        except Exception as e:
            logger.error(f"  严重错误 (退出循环): {e}")
            raise LoopConditionCheckError("Text loop condition check failed") from e
        finally:
            _close_image_quietly(ss)
    
    return False

core_engine_version = f"{CORE_VERSION} (Core) / OpenCV: {OPENCV_AVAILABLE}"

# ======================================================================
# RUN 处理函数：执行命令/脚本
# ======================================================================
def _split_command_line(text):
    """Split a command line without invoking a shell."""
    text = str(text or '').strip()
    if not text:
        return []
    if sys.platform == 'win32':
        argc = ctypes.c_int()
        argv = ctypes.windll.shell32.CommandLineToArgvW(text, ctypes.byref(argc))
        if argv:
            try:
                return [argv[i] for i in range(argc.value)]
            finally:
                ctypes.windll.kernel32.LocalFree(argv)
    return shlex.split(text, posix=(sys.platform != 'win32'))

def _build_run_command(command, args):
    cmd_list = _split_command_line(command)
    if args:
        cmd_list.extend(_split_command_line(args))
    return cmd_list

def _register_active_process(ctx, process):
    lock = ctx.setdefault('_active_process_lock', threading.RLock())
    active = ctx.setdefault('_active_processes', set())
    with lock:
        active.add(process)

def _unregister_active_process(ctx, process):
    lock = ctx.setdefault('_active_process_lock', threading.RLock())
    active = ctx.setdefault('_active_processes', set())
    with lock:
        active.discard(process)

def terminate_process_tree(process, wait_timeout=0.5):
    if process is None or process.poll() is not None:
        return

    if sys.platform == 'win32':
        try:
            subprocess.run(
                ['taskkill', '/PID', str(process.pid), '/T', '/F'],
                capture_output=True,
                text=True,
                timeout=max(2, wait_timeout + 1)
            )
            process.wait(timeout=wait_timeout)
            return
        except Exception:
            pass

    try:
        process.terminate()
        process.wait(timeout=wait_timeout)
        return
    except Exception:
        pass

    if process.poll() is None:
        try:
            process.kill()
            process.wait(timeout=wait_timeout)
        except Exception:
            pass

def cleanup_active_processes(ctx):
    if not ctx:
        return
    lock = ctx.setdefault('_active_process_lock', threading.RLock())
    active = ctx.setdefault('_active_processes', set())
    with lock:
        processes = list(active)
        active.clear()
    for process in processes:
        terminate_process_tree(process)

def _execute_subprocess(cmd_list, shell_mode, cwd, timeout, save_output, ctx, run_mode_name):
    """提取的通用子进程执行与输出处理逻辑"""
    process = None
    try:
        process = subprocess.Popen(
            cmd_list,
            shell=shell_mode,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=cwd if cwd else None
        )
        _register_active_process(ctx, process)

        deadline = time.monotonic() + timeout
        while True:
            if _ctx_check_stop(ctx):
                raise MacroStopException("Stop requested while RUN process is active")
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                terminate_process_tree(process)
                logger.error(f"  {run_mode_name}执行超时 ({timeout}秒)")
                return False
            try:
                stdout, stderr = process.communicate(timeout=min(0.1, remaining))
                break
            except subprocess.TimeoutExpired:
                continue
        
        output = stdout.strip() if stdout else stderr.strip()
        
        if process.returncode == 0:
            logger.info(f"  {run_mode_name}执行成功")
            if output:
                logger.info("        命令输出: %s chars", len(output))
            if save_output and output:
                _ctx_set_clipboard_var(ctx, output)
                try:
                    if not _copy_to_clipboard_with_retry(output, ctx):
                        raise RuntimeError("clipboard is busy")
                    logger.info("        已保存到剪贴板")
                except Exception:
                    pass
            return True
        else:
            logger.error(f"  {run_mode_name}执行失败 (退出码: {process.returncode})")
            if output:
                logger.error("        命令错误输出: %s chars", len(output))
            return False
            
    except MacroStopException:
        terminate_process_tree(process)
        raise
    except Exception as e:
        terminate_process_tree(process)
        logger.error(f"  {run_mode_name}执行错误: {e}")
        return False
    finally:
        if process is not None:
            _unregister_active_process(ctx, process)
            if process.poll() is not None:
                for stream in (process.stdout, process.stderr):
                    if stream and not stream.closed:
                        stream.close()

def _get_run_common_params(p):
    return {
        'args': p.get('args', ''),
        'timeout': _parse_run_timeout(p),
        'cwd': p.get('cwd', None),
        'save_output': p.get('save_output', False),
    }


def _handle_run(p, ctx):
    """执行命令或脚本
    
    参数:
        run_type: 类型 ("command" | "script")
        command: 命令 (run_type=command)
        script_path: 脚本路径 (run_type=script)
        interpreter: 解释器 (run_type=script, 默认 python)
        args: 命令/脚本参数
        timeout: 超时秒数 (默认 30)
        cwd: 工作目录
        save_output: 保存输出到剪贴板
    
    返回:
        bool | str:
            True: 成功
            False: 失败
            'SKIPPED': 被策略跳过（例如 run_enabled=False）
    """
    if not ctx.get('run_enabled', False):
        logger.info("  已跳过（执行外部命令默认已禁用，请在设置中手动开启）")
        return 'SKIPPED'

    p = _render_string_params(p, ctx)
    run_type = p.get('run_type', 'command')
    
    # === 文件写入模式 ===
    if run_type == 'file':
        logger.info("  file 写入模式已禁用，请改用 WRITE_FILE 写入文本结果")
        return False
    
    elif run_type == 'command':
        command = p.get('command', '')
        common = _get_run_common_params(p)
        args = common['args']
        timeout = common['timeout']
        cwd = common['cwd']
        save_output = common['save_output']
        shell_mode = bool(p.get('shell_mode', False))
        
        if not command:
            logger.error("  错误: 未指定命令")
            return False
        
        if shell_mode:
            if not ctx.get('allow_shell_mode', False):
                logger.error("  error: shell mode is disabled by default; use normal argument mode")
                return False
            logger.warning("  warning: shell mode is enabled for trusted macros only")
            run_cmd = f"{command} {args}" if args else command
        else:
            try:
                run_cmd = _build_run_command(command, args)
            except ValueError as e:
                logger.error(f"  命令参数解析失败: {e}")
                return False
            if not run_cmd:
                logger.error("  错误: 命令为空")
                return False
        
        return _execute_subprocess(run_cmd, shell_mode, cwd, timeout, save_output, ctx, "命令")
    
    # === 脚本执行模式 ===
    elif run_type == 'script':
        script_path = p.get('script_path', '')
        interpreter = p.get('interpreter', 'python')
        common = _get_run_common_params(p)
        args = common['args']
        timeout = common['timeout']
        cwd = common['cwd']
        save_output = common['save_output']
        
        if not script_path:
            logger.error("  错误: 未指定脚本路径")
            return False
        
        # Check script existence and enforce the file sandbox
        try:
            script_path = _resolve_safe_file_path(script_path, ctx, 'RUN script')
        except Exception as e:
            logger.error(f"  error: {e}")
            return False
        if not os.path.exists(script_path):
            logger.error("  error: script file does not exist: %s", _log_display_text(_log_display_path(script_path, ctx)))
            return False
        
        # 解释器映射
        INTERPRETERS = {
            'python': 'python',
            'python3': 'python',
            'node': 'node',
            'powershell': 'powershell',
            'cmd': 'cmd',
            'bat': 'cmd',
        }
        cmd = INTERPRETERS.get(interpreter)
        if not cmd:
            logger.error(f"  error: unsupported script interpreter: {interpreter}")
            return False
        
        cmd_list = [cmd, script_path]
        if args:
            try:
                cmd_list.extend(_split_command_line(args))
            except ValueError:
                cmd_list.append(args)
        
        return _execute_subprocess(cmd_list, False, cwd, timeout, save_output, ctx, "脚本")
    
    # 未知类型
    else:
        logger.error(f"  错误: 未知的 run_type: {run_type}")
        return False


# ======================================================================
# 宏数据校验（从 MacroMate.py 迁移，完成 1.7.0 Beta 中声明的迁移）
# ======================================================================
def validate_macro_data(data):
    """
    验证宏数据结构是否有效

    Args:
        data: 从 JSON 加载的数据

    Returns:
        bool: 数据是否有效
    """
    # 必须是列表
    if not isinstance(data, list):
        logger.error("根对象不是列表")
        return False

    # 验证每个步骤的基本结构
    for i, step in enumerate(data):
        # 必须是字典
        if not isinstance(step, dict):
            logger.error(f"步骤 {i+1} 不是字典对象")
            return False

        # 必须包含 'action' 字段
        if 'action' not in step:
            logger.error(f"步骤 {i+1} 缺少 'action' 字段")
            return False

        # 必须包含 'params' 字段且为字典
        if 'params' not in step or not isinstance(step['params'], dict):
            logger.error(f"步骤 {i+1} 缺少 'params' 字段或格式错误")
            return False

        # 验证 action 是否是已知的动作类型（仅警告，不阻止加载）
        if step['action'] not in MacroSchema.ACTION_TRANSLATIONS:
            logger.warning(f"步骤 {i+1} 包含未知的动作类型: {step['action']}")
            # 不返回 False，允许加载未知动作类型（向前兼容）

    return True
