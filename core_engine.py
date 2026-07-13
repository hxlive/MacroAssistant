# -*- coding: utf-8 -*-
# core_engine.py
# 描述:自动化宏的核心功能引擎
# Version: 1.8.2
CORE_VERSION = "1.8.2"

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
from PIL import ImageGrab
import re
import pyperclip
import json
import os
import sys
import subprocess
import shlex
import ctypes
import functools
import threading
import tempfile
import ast
import math
import operator as operator_module
from decimal import Decimal, InvalidOperation

if sys.platform == 'win32':
    ctypes.windll.shell32.CommandLineToArgvW.argtypes = (ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int))
    ctypes.windll.shell32.CommandLineToArgvW.restype = ctypes.POINTER(ctypes.c_wchar_p)
    ctypes.windll.kernel32.LocalFree.argtypes = (ctypes.c_void_p,)
    ctypes.windll.kernel32.LocalFree.restype = ctypes.c_void_p

# ======================================================================
# 7. 宏文件持久化工具 (从 MacroMate.py 迁移)
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
    print("[配置] FAIL 未找到 pygetwindow 库 (pip install pygetwindow)。'激活窗口' 功能将不可用。")

# ======================================================================
# 全局配置
# ======================================================================
FORCE_OCR_ENGINE = None 
ENABLE_GLOBAL_FALLBACK = True
# 条件循环检测间隔 (秒) - 平衡流畅度与准确率
LOOP_CHECK_INTERVAL = 0.2  # 优化: 从 0.5s 降低到 0.2s (平衡流畅度与准确率)
# 性能与缓存相关常量
LOOP_PHYSICAL_COOLDOWN = 0.05  # 循环物理冷却时间（秒），防止队列瞬间爆炸
CACHE_BOX_PADDING = 50  # 缓存区域扩展边距（像素）
TEMPLATE_CACHE_SIZE = 200  # 模板缓存大小，限制大图/多缩放场景的内存占用
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
    print("[严重错误] 未找到 'ocr_engine.py'。")
    class ocr_engine:
        def find_text_location(*args, **kwargs): return None
        WINOCR_AVAILABLE = False
        TESSERACT_AVAILABLE = False
        RAPIDOCR_AVAILABLE = False

try:
    import vlm_engine
except ImportError:
    print("[配置] FAIL 未找到 'vlm_engine.py'。AI 自然语言指令功能将不可用。")
    class vlm_engine:
        def find_location_by_vlm(*args, **kwargs): return None
        VLM_AVAILABLE = False

try:
    import cv2
    import numpy as np 
    OPENCV_AVAILABLE = True
    print("[CONFIG] OpenCV engine ready")
except ImportError:
    OPENCV_AVAILABLE = False
    print("[CONFIG] OpenCV not found, fallback to slower image matching")

# ======================================================================
# 快捷键工具模块
# ======================================================================
class HotkeyUtils:
    PYNPUT_TO_VK = {
        'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74, 'f6': 0x75,
        'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
        'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45, 'f': 0x46, 'g': 0x47,
        'h': 0x48, 'i': 0x49, 'j': 0x4A, 'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E,
        'o': 0x4F, 'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54, 'u': 0x55,
        'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
        '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35, '6': 0x36,
        '7': 0x37, '8': 0x38, '9': 0x39,
        'enter': 0x0D, 'space': 0x20, 'tab': 0x09, 'caps_lock': 0x14,
        'esc': 0x1B, 'page_up': 0x21, 'page_down': 0x22, 'end': 0x23, 'home': 0x24,
        'left': 0x25, 'up': 0x26, 'right': 0x27, 'down': 0x28, 'insert': 0x2D, 'delete': 0x2E,
        'backspace': 0x08,
    }
    VK_TO_PYNPUT = {v: k for k, v in PYNPUT_TO_VK.items()}
    
    if sys.platform == 'win32':
        PYNPUT_MOD_TO_WIN_MOD = {
            'ctrl': 0x0002,  # win32con.MOD_CONTROL
            'alt': 0x0001,   # win32con.MOD_ALT
            'shift': 0x0004, # win32con.MOD_SHIFT
            'cmd': 0x0008,   # win32con.MOD_WIN
        }
    else:
        PYNPUT_MOD_TO_WIN_MOD = {}
    
    @staticmethod
    def format_hotkey_display(hotkey_str):
        if not hotkey_str or "录制" in hotkey_str:
            return hotkey_str
        try:
            parts = hotkey_str.split('+')
            display_parts = []
            for part in parts:
                if part.lower() in {'ctrl', 'alt', 'shift', 'cmd'}:
                    display_parts.append(part.capitalize())
                else:
                    display_parts.append(part.upper())
            return "+".join(display_parts)
        except Exception:
            return hotkey_str.upper()

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
        'JSON_EXTRACT':    '22. JSON 提取',   # Json Extract
        'PROMPT_INPUT':    '23. 人工输入',   # Prompt Input
        'GOTO_LABEL':     '24. 跳转到标签',
        'GOTO_IF':        '25. 条件跳转',        # Goto If
        'READ_FILE':      '26. 读取文本文件到变量',
        'WRITE_FILE':     '27. 写入文本文件',     # Write File
        'FOREACH_LINE':    '28. 批量处理文本行',   # Batch Lines
        'END_FOREACH':     '29. 结束批量处理',
        'RUN':            '30. 执行命令/脚本',
        'AI_COMMAND':      '31. AI 自然语言指令',
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

    def clear_cache(self, loop_id):
        with self._lock:
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
    def __init__(self, raw_dict=None):
        if isinstance(raw_dict, RunContext):
            self._data = raw_dict._data
            self.perf = raw_dict.perf
            self.loop_cache = raw_dict.loop_cache
            self._data['_goto_counts'] = {}
            return

        self._data = raw_dict if isinstance(raw_dict, dict) else {}
        self._data.setdefault('vars', {})
        self._data.setdefault('stop_requested', False)
        self._data.setdefault('last_pos', (None, None))
        self._data.setdefault('clipboard_var', '')
        self._data.setdefault('_active_processes', set())
        self._data.setdefault('_active_process_lock', __import__('threading').RLock())
        self._data['_goto_counts'] = {}

        self.perf = PerformanceMonitor()
        self.loop_cache = LoopCacheManager()

    def check_stop(self):
        return self._data.get('stop_requested', False)

    @property
    def vars(self):
        return self._data.setdefault('vars', {})

    def get_var(self, name, default=None):
        return self.vars.get(name, default)

    def set_var(self, name, value):
        self.vars[name] = value

    def __getitem__(self, key): return self._data[key]
    def __setitem__(self, key, val): self._data[key] = val
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
    return bool(ctx.get('stop_requested', False))


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
        print(f"  [ERROR] {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_required_coords_or_break(p, action, x_key='x', y_key='y'):
    """解析必填的整数坐标，失败时返回 _ACTION_BREAK。"""
    try:
        x, y = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError):
        print(f"  [ERROR] {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_offset_or_break(p, action, x_key='x_offset', y_key='y_offset'):
    """解析整数偏移量，失败时返回 _ACTION_BREAK。"""
    try:
        ox, oy = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError):
        print(f"  [ERROR] {action} invalid offset")
        return _ACTION_BREAK
    return (ox, oy)


def _parse_positive_int_or_break(p, action, key, default=None):
    """解析正整数参数，失败时返回 _ACTION_BREAK。"""
    value, ok = _safe_param_int(p, key, default, min_value=1)
    if not ok:
        print(f"  [ERROR] {action} {key} must be a positive integer")
        return _ACTION_BREAK
    return value


def _parse_int_or_break(p, action, key, default=0):
    """解析整数参数，失败时返回 _ACTION_BREAK。"""
    value, ok = _safe_param_int(p, key, default)
    if not ok:
        print(f"  [ERROR] {action} {key} must be an integer")
        return _ACTION_BREAK
    return value


def _warn_param_default(action, name, default):
    print(f"  [WARN] {action} invalid parameter '{name}', using default: {default}")

def _error_param_skip(action, name, expected):
    print(f"  [ERROR] {action} invalid parameter '{name}' expected {expected}; step skipped")


def _run_with_fail_stop(action, p, fn, ctx=None, var_name=None, default_value=''):
    """执行可能抛异常的 IO/提取操作，统一处理 fail_stop 行为。"""
    try:
        return fn()
    except Exception as e:
        if ctx is not None and var_name is not None:
            _ctx_set_var(ctx, var_name, default_value)
        print(f"  [错误] {action} 失败: {e}")
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_DONE


def _close_image_quietly(image):
    if image:
        try:
            image.close()
        except Exception:
            pass


def _console_safe_text(value):
    """Return text that can be printed even when stdout uses a legacy code page."""
    text = str(value)
    encoding = getattr(sys.stdout, 'encoding', None) or 'utf-8'
    return text.encode(encoding, errors='replace').decode(encoding, errors='replace')

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

def smart_screenshot(region=None, pad=0):
    try:
        if sys.platform == 'win32':
            import ctypes
            SM_XVIRTUALSCREEN = 76
            SM_YVIRTUALSCREEN = 77
            SM_CXVIRTUALSCREEN = 78
            SM_CYVIRTUALSCREEN = 79
            user32 = ctypes.windll.user32
            vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
            
            if region:
                x1 = region[0] - pad
                y1 = region[1] - pad
                x2 = region[0] + region[2] + pad
                y2 = region[1] + region[3] + pad
                
                # 针对多屏幕或跨主屏坐标，采用 all_screens 抓取并裁剪
                screen_w = user32.GetSystemMetrics(0)
                screen_h = user32.GetSystemMetrics(1)
                if x1 < 0 or y1 < 0 or x2 > screen_w or y2 > screen_h or vx < 0 or vy < 0:
                    full_screen = ImageGrab.grab(all_screens=True)
                    try:
                        crop_box = (x1 - vx, y1 - vy, x2 - vx, y2 - vy)
                        # 确保不越界并精确计算钳位后的偏移量
                        cx1 = max(0, crop_box[0])
                        cy1 = max(0, crop_box[1])
                        cx2 = min(full_screen.width, crop_box[2])
                        cy2 = min(full_screen.height, crop_box[3])
                        
                        crop_box = (cx1, cy1, cx2, cy2)
                        if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
                            raise ValueError("Invalid crop box geometry")
                        cropped_img = full_screen.crop(crop_box)
                        full_screen.close()
                        # 使用 clamp 限制后的值反向修正物理绝对坐标作为 offset
                        return cropped_img, (cx1 + vx, cy1 + vy)
                    except Exception as e:
                        try: full_screen.close()
                        except Exception: pass
                        raise e
                
                return ImageGrab.grab(bbox=(x1, y1, x2, y2)), (x1, y1)
            else:
                primary_w = user32.GetSystemMetrics(0)
                primary_h = user32.GetSystemMetrics(1)
                if vw > 0 and vh > 0 and (vx != 0 or vy != 0 or vw > primary_w or vh > primary_h):
                    return ImageGrab.grab(all_screens=True), (vx, vy)
                return ImageGrab.grab(), (0, 0)
        else:
            if region:
                x = max(0, region[0] - pad)
                y = max(0, region[1] - pad)
                return ImageGrab.grab(bbox=(x, y, region[0]+region[2]+pad, region[1]+region[3]+pad)), (x, y)
            return ImageGrab.grab(), (0, 0)
    except OSError as e:
        raise ScreenshotUnavailableError("Screen capture is unavailable") from e

SCALES = (1.0, 0.9, 1.1, 0.8, 1.2)
@functools.lru_cache(maxsize=TEMPLATE_CACHE_SIZE)  # [优化] 增大缓存以减少文件读取
def _get_template(path, scale):
    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if img is None: return None, 0, 0
    if scale != 1.0:
        h, w = img.shape[:2]
        img = cv2.resize(img, (int(w*scale), int(h*scale)), interpolation=cv2.INTER_AREA)
    return img, img.shape[1], img.shape[0]

def find_image_cv2(path, conf, screenshot_pil, offset=(0,0), enhanced_mode=False, ctx=None):
    if not OPENCV_AVAILABLE: return None
    try:
        t0 = time.time()
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
            _get_perf(ctx).record_time(time.time()-t0, False)
            return (cx, cy, w, h), val
    except (cv2.error, ValueError, TypeError, AttributeError) as e:
        print(f"CV2找图错误: {e}")
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
        print(f"[quick_check_cv2] 异常: {e}")
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
    base_dir = ctx.get('macro_base_dir') if ctx else None
    if base_dir:
        roots.append(base_dir)
    for root in (ctx or {}).get('allowed_file_roots', []) or []:
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

    base_dir = (ctx or {}).get('macro_base_dir') or os.getcwd()
    expanded = os.path.expandvars(os.path.expanduser(file_path))
    if not os.path.isabs(expanded):
        expanded = os.path.join(base_dir, expanded)
    resolved = _normalize_path_for_compare(expanded)

    if (ctx or {}).get('allow_external_paths', False):
        return resolved

    for root in _get_allowed_file_roots(ctx or {}):
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

def _parse_json_path(path):
    path = str(path or '').strip()
    if not path or path == '$':
        return []
    if path.startswith('$'):
        path = path[1:]
    if path.startswith('.'):
        path = path[1:]

    tokens = []
    i = 0
    while i < len(path):
        if path[i] == '.':
            i += 1
            continue
        if path[i] == '[':
            end = path.find(']', i + 1)
            if end == -1:
                raise ValueError("JSON path bracket is not closed")
            raw = path[i + 1:end].strip()
            if not raw:
                raise ValueError("JSON path bracket is empty")
            if (raw[0:1] == raw[-1:] and raw[0:1] in ("'", '"')):
                tokens.append(raw[1:-1])
            else:
                try:
                    tokens.append(int(raw))
                except ValueError:
                    tokens.append(raw)
            i = end + 1
            continue

        start = i
        while i < len(path) and path[i] not in '.[':
            i += 1
        key = path[start:i].strip()
        if not key:
            raise ValueError("JSON path contains an empty key")
        tokens.append(key)
    return tokens

def _json_extract_value(source_text, json_path):
    data = json.loads(source_text)
    current = data
    for token in _parse_json_path(json_path):
        if isinstance(token, int):
            if not isinstance(current, list):
                raise KeyError(f"expected list before index [{token}]")
            current = current[token]
        else:
            if not isinstance(current, dict):
                raise KeyError(f"expected object before key '{token}'")
            current = current[token]
    if isinstance(current, (dict, list)):
        return json.dumps(current, ensure_ascii=False)
    if current is None:
        return ''
    return str(current)

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
    end_time = time.time() + max(0.0, float(seconds or 0))
    while True:
        _raise_if_stop_requested(ctx, "Stop requested during wait")
        remaining = end_time - time.time()
        if remaining <= 0:
            return
        time.sleep(min(interval, remaining))

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
_JSON_EXTRACT_ERRORS = (json.JSONDecodeError, ValueError, KeyError, IndexError, TypeError)


def _handle_click_action(p, ctx):
    btn = p.get('button', 'left').lower()
    clicks, ok = _safe_param_int(p, 'clicks', 1, min_value=1)
    if not ok:
        _error_param_skip('CLICK', 'clicks', 'positive integer')
        return _ACTION_SKIP
    interval, ok = _safe_param_float(p, 'interval', 0.0, min_value=0)
    if not ok:
        _warn_param_default('CLICK', 'interval', interval)
    duration, ok = _safe_param_float(p, 'duration', 0.0, min_value=0)
    if not ok:
        _warn_param_default('CLICK', 'duration', duration)
    coords = _parse_optional_coords_or_break(p, 'CLICK')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    pyautogui.click(x=x, y=y, button=btn, clicks=clicks, interval=interval, duration=duration)
    if x is not None and y is not None:
        ctx['last_pos'] = (x, y)
    return _ACTION_DONE


def _handle_move_to_action(p, ctx):
    coords = _parse_required_coords_or_break(p, 'MOVE_TO')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
    if not ok:
        _warn_param_default('MOVE_TO', 'duration', duration)
    pyautogui.moveTo(x, y, duration=duration)
    ctx['last_pos'] = (x, y)
    return _ACTION_DONE


def _handle_move_offset_action(p, ctx):
    if ctx['last_pos'][0] is None or ctx['last_pos'][1] is None:
        print("  [ERROR] MOVE_OFFSET has no last position")
        return _ACTION_BREAK
    offset = _parse_offset_or_break(p, 'MOVE_OFFSET')
    if offset == _ACTION_BREAK:
        return _ACTION_BREAK
    ox, oy = offset
    duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
    if not ok:
        _warn_param_default('MOVE_OFFSET', 'duration', duration)
    pyautogui.move(ox, oy, duration=duration)
    ctx['last_pos'] = (ctx['last_pos'][0] + ox, ctx['last_pos'][1] + oy)
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
    interrupted = False
    for elapsed_ms in range(0, total_ms, 100):
        if _ctx_check_stop(ctx):
            interrupted = True
            break
        time.sleep(min(100, total_ms - elapsed_ms) / 1000.0)
    if interrupted:
        raise MacroStopException("Stop requested during WAIT")
    return _ACTION_DONE


def _handle_type_text_action(p, ctx):
    interval, ok = _safe_param_float(p, 'interval', 0.0, min_value=0)
    if not ok:
        _warn_param_default('TYPE_TEXT', 'interval', interval)
    text = _render_param(p, 'text', ctx)
    if not text:
        print("  [WARN] TYPE_TEXT has no text; step skipped")
        return _ACTION_SKIP

    if '{CLIPBOARD}' in text:
        clipboard_content = ctx.get('clipboard_var', '')
        if not clipboard_content:
            try:
                clipboard_content = pyperclip.paste()
            except Exception:
                clipboard_content = ''
        text = text.replace('{CLIPBOARD}', clipboard_content)
        print(f"  [TYPE_TEXT] replaced clipboard placeholder: {text}")

    if interval > 0:
        pyautogui.write(text, interval=interval)
    else:
        copy_success = _copy_to_clipboard_with_retry(text, ctx)
        if copy_success:
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
        else:
            print("  [ERROR] clipboard copy failed; paste skipped")
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
            print(f"  [WARN] {msg}")
            return _ACTION_SKIP
        raise RuntimeError(msg)
    title = p.get('title')
    if not title:
        msg = "ACTIVATE_WINDOW is missing a window title"
        if p.get('ignore_fail', False):
            print(f"  [WARN] {msg}")
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
        print(f"  [成功] 已激活窗口: {target_win.title}")
        for _ in range(5):
            if _ctx_check_stop(ctx):
                raise MacroStopException("用户在窗口激活期间请求停止")
            time.sleep(0.1) 
    except Exception as e:
        msg = str(e)
        if p.get('ignore_fail', False) and not msg.startswith("No window title contains"):
            print(f"  [WARN] ACTIVATE_WINDOW failed and was ignored: {e}")
            return _ACTION_SKIP
        raise RuntimeError(f"ACTIVATE_WINDOW failed: {e}") from e

    return _ACTION_DONE

def _handle_note_action(p, ctx):
    note_text = p.get('text', '')
    if note_text:
        print(f"  [NOTE] {note_text}")
    return _ACTION_SKIP


def _handle_set_var_action(p, ctx):
    var_name = p.get('var_name', '').strip()
    var_value = _render_param(p, 'var_value', ctx)
    if var_name:
        _ctx_set_var(ctx, var_name, var_value)
        print(f"  [VAR] {var_name} = '{var_value}'")
    return _ACTION_DONE



def _handle_prompt_input_action(p, ctx):
    var_name = p.get('var_name', '').strip()
    title = _render_param(p, 'title', ctx, '智点助手输入')
    prompt = _render_param(p, 'prompt', ctx, '请输入内容:')
    default_value = _render_param(p, 'default_value', ctx)
    if not var_name:
        print("  [错误] PROMPT_INPUT 缺少变量名")
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
    print(f"  [变量] {var_name} = '{_ctx_get_var(ctx, var_name)}' (用户输入)")
         
    return _ACTION_DONE

def _handle_read_file_action(p, ctx):
    file_path = _render_param(p, 'file_path', ctx)
    var_name = p.get('var_name', '').strip()
    encoding = p.get('encoding', 'utf-8')
    if not file_path or not var_name:
        print("  [错误] READ_FILE 参数不完整")
        return _ACTION_BREAK

    def _do_read():
        safe_path = _resolve_safe_file_path(file_path, ctx, 'READ_FILE')
        with open(safe_path, 'r', encoding=encoding) as f:
            _ctx_set_var(ctx, var_name, f.read())
        print(f"  [变量] {var_name} 已读取文本文件 ({len(_ctx_get_var(ctx, var_name))} 字符)")
        return _ACTION_DONE

    return _run_with_fail_stop('READ_FILE', p, _do_read, ctx, var_name)

def _handle_extract_var_action(p, ctx):
    source = _render_param(p, 'source_text', ctx)
    pattern = str(p.get('regex', ''))
    var_name = p.get('var_name', '').strip()
    if not pattern or not var_name:
        print("  [错误] EXTRACT_VAR 参数不完整")
        return _ACTION_BREAK

    def _do_extract():
        match = re.search(pattern, source)
        val = match.group(1) if match and match.lastindex else (match.group(0) if match else '')
        _ctx_set_var(ctx, var_name, val)
        print(f"  [变量] {var_name} 提取为 '{val}'")
        return _ACTION_DONE

    return _run_with_fail_stop('EXTRACT_VAR', p, _do_extract, ctx, var_name)

def _handle_json_extract_action(p, ctx):
    source = _render_param(p, 'source_json', ctx)
    json_path = _render_param(p, 'json_path', ctx)
    var_name = p.get('var_name', '').strip()
    default_value = _render_param(p, 'default_value', ctx)
    has_default = (
        _parse_bool_param(p.get('use_default', False), False)
        or ('default_value' in p and str(p.get('default_value', '')) != '')
    )
    if not source or not var_name:
        print("  [错误] JSON_EXTRACT 参数不完整")
        return _ACTION_BREAK
    try:
        val = _json_extract_value(source, json_path)
        _ctx_set_var(ctx, var_name, val)
        print(f"  [变量] {var_name} = '{val}' (JSON)")
    except _JSON_EXTRACT_ERRORS as e:
        if has_default:
            _ctx_set_var(ctx, var_name, default_value)
            print(f"  [JSON] 提取失败，使用默认值 '{default_value}': {e}")
        else:
            print(f"  [错误] JSON_EXTRACT 提取失败: {e}")
            if p.get('fail_stop', False): return _ACTION_BREAK
    return _ACTION_DONE

def _handle_write_file_action(p, ctx):
    file_path = _render_param(p, 'file_path', ctx)
    content = _render_param(p, 'content', ctx)
    encoding = p.get('encoding', 'utf-8')
    append = _parse_bool_param(p.get('append', False), False)
    if not file_path:
        print("  [错误] WRITE_FILE 未指定路径")
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
        print(f"  [写入] {safe_path} (追加: {append})")
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
        print("  [错误] RUN 参数 'timeout' 必须是正整数，已使用默认值 30")
    return timeout

def _handle_run_action(p, ctx):
    run_result = _handle_run(p, ctx)
    if run_result == 'SKIPPED':
        print("  [RUN] \u5df2\u8df3\u8fc7\uff0c\u7ee7\u7eed\u6267\u884c\u540e\u7eed\u6b65\u9aa4")
        return _ACTION_SKIP
    if run_result is False:
        print("  [RUN] \u6267\u884c\u5931\u8d25")
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_DONE
    print("  [RUN] \u6267\u884c\u6210\u529f")
    return _ACTION_DONE


def _handle_calculate_action(p, ctx):
    expr = _render_param(p, 'expression', ctx)
    var_name = p.get('var_name', '').strip()
    if not expr or not var_name:
        print("  [ERROR] CALCULATE missing expression or var_name")
        return _ACTION_BREAK
    try:
        val = _safe_calculate_expression(expr)
        _ctx_set_var(ctx, var_name, str(val))
        print(f"  [VAR] {var_name} = {val} (calculated)")
    except (ValueError, TypeError, OverflowError, SyntaxError, ZeroDivisionError) as e:
        print(f"  [ERROR] CALCULATE failed '{expr}': {e}")
        return _ACTION_SKIP
    return _ACTION_DONE

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
    'JSON_EXTRACT': _handle_json_extract_action,
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
        print(f"  -> IF条件不满足: '{var_val}' {op} '{expected}'")
        return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])

    print(f"  -> IF条件满足: '{var_val}' {op} '{expected}'")
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
        print(f"  [GOTO] 条件成立 ({var_val} {operator} {expected})，跳至 '{target['name']}'")
        return target['index']

    print(f"  -> GOTO_IF 条件不满足 ({var_val} {operator} {expected})")
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
    print(f"  [GOTO] 跳转到标签 '{target['name']}' -> 第 {next_pc + 1} 步")
    return next_pc


def _handle_else_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['END_IF'])


def _handle_foreach_line_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _handle_foreach_line_start(steps, pc, loops, p, ctx, status_callback)


def _handle_loop_start_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    return _handle_loop_start(steps, pc, loops, p, ctx, status_callback)


def _handle_end_loop_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if not loops:
        print("[错误] END_LOOP 缺少对应的 LOOP_START")
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
        loop_id_to_exit = loops.pop()['id']
        _get_loop_cache(ctx).exit()
        _get_loop_cache(ctx).clear_cache(loop_id_to_exit)
        if status_callback:
            status_callback(f"WARN 达到最大迭代 {top['max_iterations']} 次,强制退出")
        print("  [Loop Until] WARN 达到最大迭代次数,强制退出")
        return pc + 1

    condition_met = _check_loop_condition(top, ctx)
    if condition_met:
        loop_id_to_exit = loops.pop()['id']
        _get_loop_cache(ctx).exit()
        _get_loop_cache(ctx).clear_cache(loop_id_to_exit)
        if status_callback:
            status_callback(f"OK 条件满足,循环结束 (共 {top['iteration']} 次)")
        print("  [Loop Until] OK 条件满足,循环结束")
        return pc + 1

    print(f"  [Loop Until] FAIL 未找到目标,继续循环 (第 {top['iteration']} 次)")
    time.sleep(LOOP_CHECK_INTERVAL)
    return top['start']


def _handle_end_foreach_control(steps, pc, loops, p, ctx, status_callback, goto_labels):
    if loops and loops[-1].get('kind') == 'foreach_line':
        return loops[-1]['start']
    if loops:
        raise RuntimeError("END_FOREACH 不能结束普通循环，请使用 END_LOOP")

    print("[错误] END_FOREACH 缺少对应的批量处理开始")
    return pc + 1


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
        print("  [错误] AI 指令为空")
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

    print(f"  [AI] 执行指令: {instruction}")
    coords = _call_vlm_with_stop(instruction, region, ctx)

    if coords:
        target_x, target_y = coords
        print(f"  [AI] 返回坐标: ({target_x}, {target_y})")
        duration, ok = _safe_param_float(p, 'duration', 0.25, min_value=0)
        if not ok:
            _warn_param_default('AI_COMMAND', 'duration', duration)
        pyautogui.moveTo(target_x, target_y, duration=duration)
        ctx['last_pos'] = (target_x, target_y)
        return 'done'

    print("  [AI] 未找到目标位置")
    if p.get('fail_stop', True):
        return 'break'
    return 'done'


def _handle_find_or_condition_step(steps, pc, act, p, ctx):
    next_pc = pc + 1
    res = _handle_find_with_retries(act, p, ctx, _get_loop_cache(ctx).get_current_loop_id() is not None)

    if act.startswith('IF_'):
        if not res:
            print("  -> IF condition not met, skipping block")
            next_pc = _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])
    elif not res:
        if p.get('ignore_fail', False):
            print("  -> Target not found, skipped by ignore_fail")
            return next_pc, 'continue'
        print("  -> Target not found, macro stopped")
        return next_pc, 'break'

    # FIND_ and IF_ both move the mouse after a match to preserve stable behavior.
    # A CLICK without coordinates inside an IF block can use the current mouse position.
    if res:
        target_x, target_y = res[0], res[1]
        ctx['last_pos'] = (target_x, target_y)
        pyautogui.moveTo(target_x, target_y)

    return next_pc, 'done'


def execute_steps(steps, run_context=None, status_callback=None):
    print(f"\n--- 执行开始 (Core V{CORE_VERSION}) ---")
    perf.reset()
    loop_cache.reset()
    _get_template.cache_clear()
    ctx = RunContext(run_context)
    try:
        goto_labels = _build_goto_label_table(steps)
    except ValueError as e:
        print(f"  [GOTO] 标签配置错误: {e}")
        return False
    
    default_stop = "Ctrl+F11"
    try:
        s = ctx.get('stop_key_str', default_stop)
        stop_key_display = HotkeyUtils.format_hotkey_display(s)
    except Exception:
        stop_key_display = default_stop
    
    pc, loops = 0, []
    try:
        while pc < len(steps):

            if _ctx_check_stop(ctx): 
                print(f"  [停止] 用户请求停止 ({stop_key_display})")
                break
                
            step = steps[pc]
            ctx['_current_step_index'] = pc
            act, p = _get_step_action_and_params(step)
            if _should_log_step(act, loops, ctx):
                print(f"[{pc+1}] {act}")
            next_pc = pc + 1

            # [新增] 处理被屏蔽的普通步骤
            if _should_skip_disabled_step(step, act):
                print(f"  [屏蔽] 跳过步骤: {act}")
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
                    print(f"  [警告] 未知动作: {act}")

            except pyautogui.FailSafeException as e:
                raise MacroStopException("PyAutoGUI failsafe triggered") from e
            except MacroStopException:
                raise  # 向上传播，不要吞掉
            except LoopConditionCheckError:
                raise
            except Exception as e:
                error_msg = f"  [执行异常] 步骤 {pc+1} ({act}): {e}"
                print(error_msg)
                if status_callback:
                    status_callback(f"ERR {error_msg}")
                break
            pc = next_pc
        
        return pc >= len(steps)
    finally:
        cleanup_active_processes(ctx)
        ctx.loop_cache.reset()
        print(f"--- 执行结束 ---\n[统计] {ctx.perf.get_stats()}\n")

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
            print(f"  [Retry] {act} not found; retrying in {retry_interval:g}s ({attempt + 1}/{retry_count})")
            _sleep_with_stop_check(retry_interval, ctx)
        else:
            print(f"  [Retry] {act} not found; retrying immediately ({attempt + 1}/{retry_count})")
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
                    _get_perf(ctx).record_hit(True, False); print(f"  [Loop缓存] {cached}"); ctx['last_pos'] = cached; return cached

        ocr_cache_key = f"step:{ctx.get('_current_step_index', '?')}:{act}:{p.get('text', '')}:{p.get('region', p.get('cache_box', ''))}" if not is_img else None
        res = _do_find(is_img, p, ss, offset, final_engine, ctx, ocr_cache_key)
        _raise_if_stop_requested(ctx, "Stop requested after find")

        # IF 动作不使用全局 fallback（会大幅降低性能，因为它已经重复搜索了）
        # FIND_TEXT/FIND_IMAGE 才使用 fallback
        if not res and region and ENABLE_GLOBAL_FALLBACK and enhanced_mode and not act.startswith('IF_') and not is_manual_region:
            print("  [缓存失效] 全局搜索...")
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
            ctx['last_pos'] = pos
            if act in ('FIND_TEXT', 'IF_TEXT_FOUND') and len(res) >= 3:
                save_var = p.get('save_to_var', '').strip()
                if save_var:
                    _ctx_set_var(ctx, save_var, res[2])
                    print(f"  [变量] {save_var} = '{res[2]}'")
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
            print(f"  [找到] 图 ({res_val[0][0]},{res_val[0][1]})")
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
            print(f"  [找到] 文 ({pos[0]},{pos[1]}) 内容: '{text_content}'")

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
                        print(f"  [正则提取] '{final_text}'")
                    else:
                        print("  [正则] 未匹配，保留原文")
                except Exception as e:
                    print(f"  [正则错误] {e}")

            # 处理剪贴板逻辑 (副作用)
            if ctx and p.get('save_to_clipboard', False):
                print(f"  [剪贴板] 原始文本: '{text_content}'")
                ctx['clipboard_var'] = final_text
                try:
                    if not _copy_to_clipboard_with_retry(final_text, ctx):
                        raise RuntimeError("clipboard is busy")
                    print("  [剪贴板] OK 已复制")
                except Exception as e:
                    print(f"  [剪贴板] 失败: {e}")
            
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
            print("  [批量] 所有行已处理完成")
            return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])

        top['index'] = next_index
        _set_foreach_line_vars(top, ctx)
        if cb:
            cb(f"批量处理第 {next_index + 1}/{len(top['items'])} 行")
        print(f"  [批量] 第 {next_index + 1}/{len(top['items'])} 行")
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
            print("  [FOREACH_LINE] data is empty, block skipped")
            return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])
        lines = source_text.splitlines()
        if skip_empty:
            lines = [line for line in lines if line.strip()]
        if len(lines) > max_lines:
            raise RuntimeError(f"FOREACH_LINE line count {len(lines)} exceeds max_lines {max_lines}")

    if not lines:
        print("  [批量] 没有可处理的行，跳过批量处理块")
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
    print(f"  [批量] 开始处理 {len(lines)} 行")
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
            loop_id_to_exit = loops.pop()['id']
            _get_loop_cache(ctx).exit()
            _get_loop_cache(ctx).clear_cache(loop_id_to_exit)
            if cb: cb(f"达到最大迭代 {top['max_iterations']} 次,循环结束")
            print(f"  [Loop] 警告:达到最大迭代次数 {top['max_iterations']}")
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
                loop_id_to_exit = loops.pop()['id']
                _get_loop_cache(ctx).exit()
                _get_loop_cache(ctx).clear_cache(loop_id_to_exit)
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
                print(f"  [Loop Until Image] 目标: {loop_data['condition_image']} (区域: {p['region']})")
            elif 'cache_box' in p:
                loop_data['cache_box'] = p['cache_box']
                print(f"  [Loop Until Image] 目标: {loop_data['condition_image']} (区域: {p['cache_box']})")
            else:
                print(f"  [Loop Until Image] 目标: {loop_data['condition_image']} (全屏)")
        elif mode == 'until_text':
            loop_data['condition_text'] = p.get('condition_text', '')
            loop_data['lang'] = p.get('lang', 'eng')
            # [优化] 保存搜索区域，加速条件检测
            if 'region' in p:
                print(f"  [Loop Until Text] 目标: {loop_data['condition_text']} (区域: {p['region']})")
            elif 'cache_box' in p:
                loop_data['cache_box'] = p['cache_box']
                print(f"  [Loop Until Text] 目标: {loop_data['condition_text']} (区域: {p['cache_box']})")
            else:
                print(f"  [Loop Until Text] 目标: {loop_data['condition_text']} (全屏)")
        
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
            print(f"  [Loop Until] 警告: 图像路径无效 '{path}'")
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
                print(f"  [Loop Until] OK 找到目标图像: {os.path.basename(path)}")
            return found
        except (ValueError, TypeError, AttributeError, IndexError) + ((cv2.error,) if OPENCV_AVAILABLE else ()) as e:
            print(f"  [Loop Until] 图像检测错误: {e}")
            return False
        except Exception as e:
            print(f"  [Loop Until] 严重错误 (退出循环): {e}")
            raise LoopConditionCheckError("Image loop condition check failed") from e
        finally:
            _close_image_quietly(ss)
    
    elif mode == 'until_text':
        text = loop_data.get('condition_text', '')
        lang = loop_data.get('lang', 'eng')
        
        if not text:
            print("  [Loop Until] 警告: 文本条件为空")
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
                print(f"  [Loop Until] OK 找到目标文本: '{found_txt[:50]}'")
                return True
            return False
        except (ValueError, TypeError, AttributeError) as e:
            print(f"  [Loop Until] 文本检测错误: {e}")
            return False
        except Exception as e:
            print(f"  [Loop Until] 严重错误 (退出循环): {e}")
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
                print(f"  [RUN] {run_mode_name}执行超时 ({timeout}秒)")
                return False
            try:
                stdout, stderr = process.communicate(timeout=min(0.1, remaining))
                break
            except subprocess.TimeoutExpired:
                continue
        
        output = stdout.strip() if stdout else stderr.strip()
        
        if process.returncode == 0:
            print(f"  [RUN] {run_mode_name}执行成功")
            if output:
                print(f"        输出: {_console_safe_text(output[:200])}")
            if save_output and output:
                ctx['clipboard_var'] = output
                try:
                    if not _copy_to_clipboard_with_retry(output, ctx):
                        raise RuntimeError("clipboard is busy")
                    print("        已保存到剪贴板")
                except Exception:
                    pass
            return True
        else:
            print(f"  [RUN] {run_mode_name}执行失败 (退出码: {process.returncode})")
            if output:
                print(f"        错误: {_console_safe_text(output[:200])}")
            return False
            
    except MacroStopException:
        terminate_process_tree(process)
        raise
    except Exception as e:
        terminate_process_tree(process)
        print(f"  [RUN] {run_mode_name}执行错误: {e}")
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
        print("  [RUN] 已跳过（执行外部命令默认已禁用，请在设置中手动开启）")
        return 'SKIPPED'

    p = _render_string_params(p, ctx)
    run_type = p.get('run_type', 'command')
    
    # === 文件写入模式 ===
    if run_type == 'file':
        print("  [RUN] file 写入模式已禁用，请改用 WRITE_FILE 写入文本结果")
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
            print("  [RUN] 错误: 未指定命令")
            return False
        
        if shell_mode:
            if not ctx.get('allow_shell_mode', False):
                print("  [RUN] error: shell mode is disabled by default; use normal argument mode")
                return False
            print("  [RUN] warning: shell mode is enabled for trusted macros only")
            run_cmd = f"{command} {args}" if args else command
        else:
            try:
                run_cmd = _build_run_command(command, args)
            except ValueError as e:
                print(f"  [RUN] 命令参数解析失败: {e}")
                return False
            if not run_cmd:
                print("  [RUN] 错误: 命令为空")
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
            print("  [RUN] 错误: 未指定脚本路径")
            return False
        
        # Check script existence and enforce the file sandbox
        try:
            script_path = _resolve_safe_file_path(script_path, ctx, 'RUN script')
        except Exception as e:
            print(f"  [RUN] error: {e}")
            return False
        if not os.path.exists(script_path):
            print(f"  [RUN] error: script file does not exist: {script_path}")
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
            print(f"  [RUN] error: unsupported script interpreter: {interpreter}")
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
        print(f"  [RUN] 错误: 未知的 run_type: {run_type}")
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
        print("[验证失败] 根对象不是列表")
        return False

    # 验证每个步骤的基本结构
    for i, step in enumerate(data):
        # 必须是字典
        if not isinstance(step, dict):
            print(f"[验证失败] 步骤 {i+1} 不是字典对象")
            return False

        # 必须包含 'action' 字段
        if 'action' not in step:
            print(f"[验证失败] 步骤 {i+1} 缺少 'action' 字段")
            return False

        # 必须包含 'params' 字段且为字典
        if 'params' not in step or not isinstance(step['params'], dict):
            print(f"[验证失败] 步骤 {i+1} 缺少 'params' 字段或格式错误")
            return False

        # 验证 action 是否是已知的动作类型（仅警告，不阻止加载）
        if step['action'] not in MacroSchema.ACTION_TRANSLATIONS:
            print(f"[警告] 步骤 {i+1} 包含未知的动作类型: {step['action']}")
            # 不返回 False，允许加载未知动作类型（向前兼容）

    return True
