# -*- coding: utf-8 -*-
# core_engine.py
# 功能说明：自动化宏核心执行引擎，负责运行上下文、动作分发、控制流和进程管理
# Version: 1.8.6
CORE_VERSION = "1.8.6"

import screen_locator
# ======================================================================
# 模块索引
# ======================================================================
# 1. 异常、导入、可选引擎与常量
# 2. 宏结构定义与单次运行上下文
# 3. 变量、表达式、文件安全与输入辅助
# 4. 线性动作处理器与只读分发表
# 5. 控制流处理器、执行循环与屏幕定位适配
# 6. 子进程生命周期与宏结构校验

# ======================================================================
# 即时中断异常
# ======================================================================
class MacroStopException(BaseException):
    """Raised in the execution thread to stop the macro immediately."""
    pass

class LoopConditionCheckError(RuntimeError):
    """Raised when a loop exit condition cannot be checked safely."""
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
import threading
import tempfile
import ast
import math
import operator as operator_module
from types import MappingProxyType
from decimal import Decimal, InvalidOperation
from sys_utils import HotkeyUtils

import logging
logger = logging.getLogger(__name__)

if sys.platform == 'win32':
    ctypes.windll.shell32.CommandLineToArgvW.argtypes = (ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int))
    ctypes.windll.shell32.CommandLineToArgvW.restype = ctypes.POINTER(ctypes.c_wchar_p)
    ctypes.windll.kernel32.LocalFree.argtypes = (ctypes.c_void_p,)
    ctypes.windll.kernel32.LocalFree.restype = ctypes.c_void_p

# ======================================================================
# Optional window activation dependency
# ======================================================================
try:
    import pygetwindow as gw
    PYGETWINDOW_AVAILABLE = True
except ImportError:
    PYGETWINDOW_AVAILABLE = False
    logger.error("FAIL pygetwindow is unavailable (pip install pygetwindow); window activation is disabled")

# ======================================================================
# 全局配置
# ======================================================================
# Conditional-loop polling interval
LOOP_CHECK_INTERVAL = 0.2
_MOUSE_MOVE_FRAME_INTERVAL = 1.0 / 60.0  # About 60 FPS
_CLICK_PROGRESS_UPDATE_INTERVAL = 0.1
# Performance and cache constants
LOOP_PHYSICAL_COOLDOWN = 0.05
CACHE_BOX_PADDING = 50  # 缓存区域扩展边距（像素）
GOTO_LABEL_DEFAULT_MAX_JUMPS = 100
APP_DIR = os.path.dirname(os.path.abspath(__file__))
_VAR_PATTERN = re.compile(r'\{([^{}]+)\}')
_GOTO_LABEL_PATTERN = re.compile(r'^\s*(?:LABEL|标签)\s*[:：]\s*(.+?)\s*$', re.IGNORECASE)
_FIELD_SPLIT_PATTERN = re.compile(r'[,，\t|]+')

TEXT_OUTPUT_EXTENSIONS = frozenset({
    '.txt', '.log', '.ini', '.cfg', '.csv', '.tsv',
    '.json', '.jsonl', '.md', '.yaml', '.yml', '.xml'
})


if screen_locator.OPENCV_AVAILABLE:
    logger.info("OpenCV engine ready")
else:
    logger.warning("OpenCV not found; image matching is unavailable")

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
        'DRAG_TO':        '08. 按住并拖动',
        'ACTIVATE_WINDOW': '09. 激活窗口 (按标题)',
        'NOTE':           "10. 备注",
        'FIND_IMAGE':     '11. 查找图像',
        'FIND_TEXT':      '12. 查找文本 (OCR)',
        'IF_IMAGE_FOUND':  '13. IF 找到图像',
        'IF_TEXT_FOUND':   '14. IF 找到文本',
        'IF_COLOR_MATCH':  '15. IF 颜色匹配',
        'IF_VAR':         '16. IF 变量比较',
        'ELSE':           '17. ELSE（否则）',
        'END_IF':         '18. END_IF（结束 IF）',
        'LOOP_START':     '19. 循环开始',          # Loop
        'END_LOOP':       '20. 结束循环',       # EndLoop
        'SET_VAR':        '21. 设置变量',        # Set Var
        'CALCULATE':      '22. 变量计算',      # Calculate
        'EXTRACT_VAR':    '23. 正则提取变量',    # Extract
        'PROMPT_INPUT':    '24. 人工输入',   # Prompt Input
        'GOTO_LABEL':     '25. 跳转到标签',
        'GOTO_IF':        '26. 条件跳转',        # Goto If
        'READ_FILE':      '27. 读取文本文件到变量',
        'WRITE_FILE':     '28. 写入文本文件',     # Write File
        'FOREACH_LINE':    '29. 批量处理文本行',   # Batch Lines
        'END_FOREACH':     '30. 结束批量处理',
        'RUN':            '31. 执行命令/脚本',
        'AI_COMMAND':      '32. AI 自然语言指令',
    }
    ACTION_KEYS_TO_NAME = {v: k for k, v in ACTION_TRANSLATIONS.items()}
    RUN_TYPE_OPTIONS = {
        'command (命令)': 'command',
        'script (脚本)': 'script',
    }
    RUN_TYPE_DISPLAY_BY_VALUE = {v: k for k, v in RUN_TYPE_OPTIONS.items()}
    CONTROL_FLOW_ACTIONS = {'IF_IMAGE_FOUND', 'IF_TEXT_FOUND', 'IF_COLOR_MATCH', 'IF_VAR', 'ELSE', 'END_IF', 'LOOP_START', 'END_LOOP', 'FOREACH_LINE', 'END_FOREACH'}

    LANG_OPTIONS = {'chi_sim (简体中文)': 'chi_sim', 'eng (英文)': 'eng'}
    LANG_VALUES_TO_NAME = {v: k for k, v in LANG_OPTIONS.items()}

    CLICK_OPTIONS = {'left (左键)': 'left', 'right (右键)': 'right', 'middle (中键)': 'middle'}
    CLICK_VALUES_TO_NAME = {v: k for k, v in CLICK_OPTIONS.items()}


class RunContext:
    """Own mutable state for one macro execution.

    ``execute_steps`` may wrap a caller-owned dictionary for public API
    compatibility. Private engine handlers receive this object only.
    """

    def __new__(cls, raw_dict=None):
        if isinstance(raw_dict, cls):
            return raw_dict
        return super().__new__(cls)

    def __init__(self, raw_dict=None):
        if isinstance(raw_dict, RunContext):
            return
        if raw_dict is not None and not isinstance(raw_dict, dict):
            raise TypeError('run_context must be a RunContext, dict, or None')

        self._data = raw_dict if isinstance(raw_dict, dict) else {}
        if not isinstance(self._data.get('vars'), dict):
            self._data['vars'] = {}
        self._data.setdefault('stop_requested', False)
        stop_event = self._data.get('stop_event')
        if not callable(getattr(stop_event, 'is_set', None)) or not callable(getattr(stop_event, 'set', None)):
            stop_event = threading.Event()
            self._data['stop_event'] = stop_event
        if self._data['stop_requested']:
            stop_event.set()
        self._data.setdefault('last_pos', (None, None))
        self._data.setdefault('last_locate_pos', (None, None))
        self._data.setdefault('clipboard_var', '')

        self._goto_counts = {}
        self._current_step_index = None
        self._active_processes = set()
        self._active_process_lock = threading.RLock()
        self._sync_internal_compatibility_view()
        self.locator = screen_locator.LocatorSession()

    def _sync_internal_compatibility_view(self):
        """Expose internal objects for existing observers, never accept them as input."""
        self._data['_goto_counts'] = self._goto_counts
        self._data['_current_step_index'] = self._current_step_index
        self._data['_active_processes'] = self._active_processes
        self._data['_active_process_lock'] = self._active_process_lock

    def reset_execution_state(self):
        self._goto_counts.clear()
        self.set_current_step_index(None)
        self.set_last_locate_pos(None, None)

    def set_current_step_index(self, index):
        self._current_step_index = index
        self._data['_current_step_index'] = index

    def get_current_step_index(self, default=None):
        return default if self._current_step_index is None else self._current_step_index

    def try_record_goto(self, program_counter, max_jumps):
        count = self._goto_counts.get(program_counter, 0)
        if count >= max_jumps:
            return False
        self._goto_counts[program_counter] = count + 1
        return True

    def register_process(self, process):
        with self._active_process_lock:
            self._active_processes.add(process)

    def unregister_process(self, process):
        with self._active_process_lock:
            self._active_processes.discard(process)

    def drain_active_processes(self):
        with self._active_process_lock:
            processes = list(self._active_processes)
            self._active_processes.clear()
        return processes

    def get_option(self, name, default=None):
        return self._data.get(name, default)

    def check_stop(self):
        stop_event = self._data.get('stop_event')
        return bool(
            self._data.get('stop_requested', False)
            or (callable(getattr(stop_event, 'is_set', None)) and stop_event.is_set())
        )

    def request_stop(self):
        self._data['stop_requested'] = True
        self._data['stop_event'].set()

    def clear_stop_request(self):
        self._data['stop_requested'] = False
        self._data['stop_event'].clear()

    @property
    def stop_event(self):
        return self._data['stop_event']

    def has_active_processes(self):
        with self._active_process_lock:
            return bool(self._active_processes)

    @property
    def vars(self):
        return self._data['vars']

    def get_var(self, name, default=None):
        return self.vars.get(name, default)

    def set_var(self, name, value):
        self.vars[name] = value

    def get_last_pos(self):
        return self._data.setdefault('last_pos', (None, None))

    def set_last_pos(self, x, y=None):
        self._data['last_pos'] = x if y is None else (x, y)

    def get_last_locate_pos(self):
        """Return the center of the latest successful image/OCR recognition."""
        return self._data.setdefault('last_locate_pos', (None, None))

    def set_last_locate_pos(self, x, y=None):
        if y is None and isinstance(x, (list, tuple)):
            self._data['last_locate_pos'] = tuple(x)
        else:
            self._data['last_locate_pos'] = (x, y)

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



def is_stop_requested(ctx: RunContext | None):
    """Return whether the current run has received a stop request."""
    return bool(ctx and ctx.check_stop())


def request_stop(ctx: RunContext | None):
    """Request cooperative cancellation for a live RunContext."""
    if ctx is not None:
        ctx.request_stop()

# ======================================================================
# 核心工具函数
# ======================================================================
def _safe_int(value, default=None, min_value=None, max_value=None):
    try:
        result = int(value)
    except (TypeError, ValueError, OverflowError):
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


def _time_param_seconds(
        p, ms_name, legacy_name, default_seconds, *, legacy_unit='seconds'):
    """Read a new millisecond field or a legacy time field as seconds."""
    if ms_name in p:
        milliseconds, ok = _safe_param_float(
            p, ms_name, default_seconds * 1000.0, min_value=0
        )
        return milliseconds / 1000.0, ok, ms_name

    legacy_default = (
        default_seconds * 1000.0
        if legacy_unit == 'milliseconds'
        else default_seconds
    )
    legacy_value, ok = _safe_param_float(
        p, legacy_name, legacy_default, min_value=0
    )
    if legacy_unit == 'milliseconds':
        legacy_value /= 1000.0
    return legacy_value, ok, legacy_name


def _parse_optional_coords_or_break(p, action, x_key='x', y_key='y'):
    """Parse optional integer coordinates; return _ACTION_BREAK on failure."""
    try:
        x = int(p.get(x_key, '')) if str(p.get(x_key, '')).strip() else None
        y = int(p.get(y_key, '')) if str(p.get(y_key, '')).strip() else None
    except (ValueError, TypeError, OverflowError):
        logger.error(f"  {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_required_coords_or_break(p, action, x_key='x', y_key='y'):
    """Parse required integer coordinates; return _ACTION_BREAK on failure."""
    try:
        x, y = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError, OverflowError):
        logger.error(f"  {action} invalid coordinates")
        return _ACTION_BREAK
    return (x, y)


def _parse_offset_or_break(p, action, x_key='x_offset', y_key='y_offset'):
    """Parse integer offsets; return _ACTION_BREAK on failure."""
    try:
        ox, oy = int(p.get(x_key, 0)), int(p.get(y_key, 0))
    except (ValueError, TypeError, OverflowError):
        logger.error(f"  {action} invalid offset")
        return _ACTION_BREAK
    return (ox, oy)


def _parse_positive_int_or_break(p, action, key, default=None):
    """Parse a positive integer parameter; return _ACTION_BREAK on failure."""
    value, ok = _safe_param_int(p, key, default, min_value=1)
    if not ok:
        logger.error(f"  {action} {key} must be a positive integer")
        return _ACTION_BREAK
    return value


def _parse_int_or_break(p, action, key, default=0):
    """Parse an integer parameter; return _ACTION_BREAK on failure."""
    value, ok = _safe_param_int(p, key, default)
    if not ok:
        logger.error(f"  {action} {key} must be an integer")
        return _ACTION_BREAK
    return value


def _warn_param_default(action, name, default):
    logger.warning(f"  {action} invalid parameter '{name}', using default: {default}")

def _error_param_skip(action, name, expected):
    logger.error(f"  {action} invalid parameter '{name}' expected {expected}; step skipped")


def _run_with_fail_stop(action, p, fn, ctx: RunContext | None = None, var_name=None, default_value=''):
    """Run an IO/extraction operation with consistent fail_stop handling."""
    try:
        return fn()
    except Exception as e:
        if ctx is not None and var_name is not None:
            ctx.set_var(var_name, default_value)
        logger.error(f"  {action} 失败: {e}")
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_DONE


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

def _log_display_path(path, ctx: RunContext | None = None, basename_only=False):
    """Return a useful path label without exposing macro-local absolute roots."""
    raw_path = os.fspath(path) if path else ''
    if not raw_path:
        return ''
    if basename_only:
        return os.path.basename(raw_path)

    absolute_path = os.path.abspath(raw_path)
    macro_base = ctx.get_macro_base_dir() if ctx else None
    if macro_base and _is_path_inside(absolute_path, macro_base):
        return os.path.relpath(absolute_path, os.path.abspath(macro_base))
    return absolute_path


def _copy_to_clipboard_with_retry(text, ctx: RunContext, retries=3, delay=0.2):
    for _ in range(retries):
        if ctx.check_stop():
            raise MacroStopException("Stop requested while writing to clipboard")
        try:
            pyperclip.copy(text)
            return True
        except Exception:
            time.sleep(delay)
    return False

def _locator_cancel_check(ctx: RunContext):
    return lambda: _raise_if_stop_requested(ctx, "Stop requested during image matching")


# Core execution engine
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
                    raise ValueError(f"Label '{label}' is inside a loop; jumping into loops is unsupported")
                key = _normalize_goto_label(label)
                if key in labels:
                    prev_idx = labels[key]['index']
                    raise ValueError(f"Duplicate label '{label}' at steps {prev_idx + 1} and {idx + 1}")
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

def _get_allowed_file_roots(ctx: RunContext):
    roots = [APP_DIR, os.getcwd(), tempfile.gettempdir()]
    base_dir = ctx.get_macro_base_dir()
    if base_dir:
        roots.append(base_dir)
    for root in ctx.get_allowed_file_roots():
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


def _clean_user_file_path(file_path):
    """Normalize paths pasted from Windows Explorer or returned by Tk dialogs."""
    cleaned = str(file_path or '').strip()
    quote_pairs = (("\"", "\""), ("'", "'"), ('“', '”'))
    for opening, closing in quote_pairs:
        if len(cleaned) >= 2 and cleaned.startswith(opening) and cleaned.endswith(closing):
            cleaned = cleaned[len(opening):-len(closing)].strip()
            break
    return os.path.normpath(cleaned) if cleaned else ''


def _resolve_safe_file_path(file_path, ctx: RunContext, purpose='file access'):
    file_path = _clean_user_file_path(file_path)
    if not file_path:
        raise ValueError('路径为空')

    base_dir = ctx.get_macro_base_dir() or os.getcwd()
    expanded = os.path.normpath(os.path.expandvars(os.path.expanduser(file_path)))
    if not os.path.isabs(expanded):
        expanded = os.path.join(base_dir, expanded)
    resolved = _normalize_path_for_compare(expanded)

    if ctx.allow_external_paths():
        return resolved

    for root in _get_allowed_file_roots(ctx):
        if _is_path_inside(resolved, root):
            return resolved

    raise PermissionError(f"{purpose} path is outside allowed directories: {file_path}")

def _ensure_text_output_path(safe_path, purpose='WRITE_FILE'):
    ext = os.path.splitext(str(safe_path))[1].lower()
    if ext not in TEXT_OUTPUT_EXTENSIONS:
        allowed = ', '.join(sorted(TEXT_OUTPUT_EXTENSIONS))
        raise ValueError(f"{purpose} only supports text output; allowed extensions: {allowed}")
    return safe_path


def _render_vars(param_str, ctx: RunContext):
    if not isinstance(param_str, str) or not param_str: return param_str
    if '{' not in param_str: return param_str
    vars_dict = ctx.vars
    def repl(match):
        var_name = match.group(1)
        return str(vars_dict.get(var_name, match.group(0)))
    return _VAR_PATTERN.sub(repl, param_str)


def _render_param(p, name, ctx: RunContext, default=''):
    return _render_vars(str(p.get(name, default)), ctx)


def _render_param_stripped(p, name, ctx: RunContext, default=''):
    return _render_param(p, name, ctx, default).strip()


def _render_string_params(params, ctx: RunContext):
    return {k: _render_vars(v, ctx) if isinstance(v, str) else v for k, v in params.items()}


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

def _set_foreach_line_vars(loop_data, ctx: RunContext):
    vars_dict = ctx.vars
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

_CALC_OPERATORS = MappingProxyType({
    ast.Add: operator_module.add,
    ast.Sub: operator_module.sub,
    ast.Mult: operator_module.mul,
    ast.Div: operator_module.truediv,
    ast.FloorDiv: operator_module.floordiv,
    ast.Mod: operator_module.mod,
})
_CALC_UNARY_OPERATORS = MappingProxyType({
    ast.UAdd: operator_module.pos,
    ast.USub: operator_module.neg,
})
_CALC_MAX_ABS_VALUE = 1e18

def _validate_calc_number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError("Only numeric constants are supported")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("Non-finite numeric values are not supported")
    if abs(value) > _CALC_MAX_ABS_VALUE:
        raise OverflowError("Calculation result is too large")
    return value

def _raise_if_stop_requested(ctx: RunContext, message="Stop requested"):
    if ctx.check_stop():
        raise MacroStopException(message)

def _sleep_with_stop_check(seconds, ctx: RunContext, interval=0.1):
    end_time = time.monotonic() + max(0.0, float(seconds or 0))
    while True:
        _raise_if_stop_requested(ctx, "Stop requested during wait")
        remaining = end_time - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(interval, remaining))


def _pause_after_pyautogui_call(ctx: RunContext):
    pause = max(0.0, float(getattr(pyautogui, 'PAUSE', 0.0) or 0.0))
    if pause > 0:
        _sleep_with_stop_check(pause, ctx, interval=0.02)


def _move_to_with_stop(x, y, duration=0.0, *, ctx: RunContext, pause_after=True):
    _raise_if_stop_requested(ctx, "Stop requested before mouse move")
    duration = max(0.0, float(duration or 0.0))
    if duration <= 0:
        pyautogui.moveTo(x, y, duration=0, _pause=False)
        _raise_if_stop_requested(ctx, "Stop requested after mouse move")
        if pause_after:
            _pause_after_pyautogui_call(ctx)
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

def _click_with_stop(
    x, y, button, clicks=1, interval=0.0, duration=0.0, *,
    ctx: RunContext, progress_callback=None,
):
    _raise_if_stop_requested(ctx, "Stop requested before click")
    duration = max(0.0, float(duration or 0.0))
    interval = max(0.0, float(interval or 0.0))
    clicks = max(1, int(clicks or 1))

    if clicks == 1 and interval == 0.0 and duration == 0.0:
        pyautogui.click(x=x, y=y, button=button, clicks=1, interval=0.0, duration=0.0)
        _raise_if_stop_requested(ctx, "Stop requested after click")
        return

    if x is not None and y is not None:
        _move_to_with_stop(x, y, duration, ctx=ctx, pause_after=False)
    last_progress_time = None
    for index in range(clicks):
        _raise_if_stop_requested(ctx, "Stop requested during click")
        pyautogui.click(button=button, clicks=1, interval=0.0, duration=0.0, _pause=False)
        current_click = index + 1
        if progress_callback is not None:
            now = time.monotonic()
            if (
                current_click == 1
                or current_click == clicks
                or last_progress_time is None
                or now - last_progress_time >= _CLICK_PROGRESS_UPDATE_INTERVAL
            ):
                progress_callback(current_click, clicks)
                last_progress_time = now
        if index < clicks - 1 and interval > 0:
            _sleep_with_stop_check(interval, ctx, interval=0.02)
    _pause_after_pyautogui_call(ctx)
    _raise_if_stop_requested(ctx, "Stop requested after click")


def _force_mouse_up(button):
    """Release a mouse button even when the pointer is on a failsafe corner."""
    try:
        pyautogui.mouseUp(button=button, _pause=False)
    except pyautogui.FailSafeException:
        failsafe_enabled = pyautogui.FAILSAFE
        try:
            pyautogui.FAILSAFE = False
            pyautogui.mouseUp(button=button, _pause=False)
        finally:
            pyautogui.FAILSAFE = failsafe_enabled
        raise


def _drag_to_with_stop(start_x, start_y, end_x, end_y, button, duration=0.0, *, ctx: RunContext):
    """Perform one interruptible drag and unconditionally release its button."""
    _move_to_with_stop(start_x, start_y, 0.0, ctx=ctx, pause_after=False)
    _raise_if_stop_requested(ctx, "Stop requested before drag")
    mouse_down_attempted = False
    try:
        mouse_down_attempted = True
        pyautogui.mouseDown(button=button, _pause=False)
        _raise_if_stop_requested(ctx, "Stop requested after mouse down")
        _move_to_with_stop(end_x, end_y, duration, ctx=ctx, pause_after=False)
    finally:
        if mouse_down_attempted:
            _force_mouse_up(button)
    _pause_after_pyautogui_call(ctx)
    _raise_if_stop_requested(ctx, "Stop requested after drag")

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
def _handle_click_action(p, ctx: RunContext, status_callback=None):
    btn = p.get('button', 'left').lower()
    clicks, ok = _safe_param_int(p, 'clicks', 1, min_value=1)
    if not ok:
        _error_param_skip('CLICK', 'clicks', 'positive integer')
        return _ACTION_SKIP
    interval, ok, interval_key = _time_param_seconds(
        p, 'interval_ms', 'interval', 0.0, legacy_unit='milliseconds'
    )
    if not ok:
        _warn_param_default('CLICK', interval_key, interval)
    duration, ok, duration_key = _time_param_seconds(
        p, 'duration_ms', 'duration', 0.0
    )
    if not ok:
        _warn_param_default('CLICK', duration_key, duration)
    coords = _parse_optional_coords_or_break(p, 'CLICK')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    progress_callback = None
    if clicks > 1 and status_callback is not None:
        progress_callback = lambda current, total: status_callback(f"点击 {current} / {total}")
    _click_with_stop(
        x, y, btn,
        clicks=clicks,
        interval=interval,
        duration=duration,
        ctx=ctx,
        progress_callback=progress_callback,
    )
    if x is not None and y is not None:
        ctx.set_last_pos(x, y)
    return _ACTION_DONE


def _handle_move_to_action(p, ctx: RunContext):
    coords = _parse_required_coords_or_break(p, 'MOVE_TO')
    if coords == _ACTION_BREAK:
        return _ACTION_BREAK
    x, y = coords
    duration, ok, duration_key = _time_param_seconds(
        p, 'duration_ms', 'duration', 0.25
    )
    if not ok:
        _warn_param_default('MOVE_TO', duration_key, duration)
    _move_to_with_stop(x, y, duration, ctx=ctx)
    ctx.set_last_pos(x, y)
    return _ACTION_DONE


def _handle_drag_to_action(p, ctx: RunContext):
    coord_keys = ('start_x', 'start_y', 'end_x', 'end_y')
    if any(key not in p or not str(p.get(key, '')).strip() for key in coord_keys):
        logger.error("  DRAG_TO requires start and end coordinates")
        return _ACTION_BREAK

    start = _parse_required_coords_or_break(
        p, 'DRAG_TO', x_key='start_x', y_key='start_y',
    )
    if start == _ACTION_BREAK:
        return _ACTION_BREAK
    end = _parse_required_coords_or_break(
        p, 'DRAG_TO', x_key='end_x', y_key='end_y',
    )
    if end == _ACTION_BREAK:
        return _ACTION_BREAK

    button = str(p.get('button', 'left')).strip().lower() or 'left'
    if button not in MacroSchema.CLICK_OPTIONS.values():
        logger.error("  DRAG_TO invalid mouse button: %s", button)
        return _ACTION_BREAK
    duration_ms, ok = _safe_param_int(p, 'duration_ms', 500, min_value=0)
    duration = duration_ms / 1000.0
    if not ok:
        logger.error("  DRAG_TO duration_ms must be a non-negative integer")
        return _ACTION_BREAK

    _drag_to_with_stop(
        start[0], start[1], end[0], end[1], button, duration, ctx=ctx,
    )
    ctx.set_last_pos(*end)
    return _ACTION_DONE


def _handle_move_offset_action(p, ctx: RunContext):
    last_x, last_y = ctx.get_last_pos()
    if last_x is None or last_y is None:
        logger.error("  MOVE_OFFSET has no last position")
        return _ACTION_BREAK
    offset = _parse_offset_or_break(p, 'MOVE_OFFSET')
    if offset == _ACTION_BREAK:
        return _ACTION_BREAK
    ox, oy = offset
    duration, ok, duration_key = _time_param_seconds(
        p, 'duration_ms', 'duration', 0.25
    )
    if not ok:
        _warn_param_default('MOVE_OFFSET', duration_key, duration)
    current_x, current_y = pyautogui.position()
    _move_to_with_stop(current_x + ox, current_y + oy, duration, ctx=ctx)
    ctx.set_last_pos(last_x + ox, last_y + oy)
    return _ACTION_DONE


def _handle_scroll_action(p, ctx: RunContext):
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


def _handle_wait_action(p, ctx: RunContext):
    total_ms = _parse_positive_int_or_break(p, 'WAIT', 'ms', 0)
    if total_ms == _ACTION_BREAK:
        return _ACTION_SKIP
    _sleep_with_stop_check(total_ms / 1000.0, ctx)
    return _ACTION_DONE


def _handle_type_text_action(p, ctx: RunContext):
    interval, ok, interval_key = _time_param_seconds(
        p, 'interval_ms', 'interval', 0.0
    )
    if not ok:
        _warn_param_default('TYPE_TEXT', interval_key, interval)
    text = _render_param(p, 'text', ctx)
    if not text:
        logger.warning("  TYPE_TEXT has no text; step skipped")
        return _ACTION_SKIP

    if '{CLIPBOARD}' in text:
        clipboard_content = ctx.get_clipboard_var()
        if not clipboard_content:
            try:
                clipboard_content = pyperclip.paste()
            except Exception:
                clipboard_content = ''
        text = text.replace('{CLIPBOARD}', clipboard_content)
        logger.info("  replaced clipboard placeholder (%s chars)", len(text))

    _raise_if_stop_requested(ctx, "Stop requested before typing")
    if interval > 0:
        for index, char in enumerate(text):
            _raise_if_stop_requested(ctx, "Stop requested during typing")
            pyautogui.write(char, interval=0, _pause=False)
            if index < len(text) - 1:
                _sleep_with_stop_check(interval, ctx, interval=0.02)
        _pause_after_pyautogui_call(ctx)
    else:
        copy_success = _copy_to_clipboard_with_retry(text, ctx)
        if copy_success:
            _sleep_with_stop_check(0.1, ctx, interval=0.02)
            _raise_if_stop_requested(ctx, "Stop requested before paste")
            pyautogui.hotkey('ctrl', 'v')
        else:
            logger.error("  clipboard copy failed; paste skipped")
    return _ACTION_DONE


def _handle_press_key_action(p, ctx: RunContext):
    keys = [k for k in p.get('key', '').lower().replace(' ', '').split('+') if k]
    if keys:
        pyautogui.hotkey(*keys)
    return _ACTION_DONE



def _handle_activate_window_action(p, ctx: RunContext):
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
        logger.info(f"  Window activated: {target_win.title}")
        for _ in range(5):
            if ctx.check_stop():
                raise MacroStopException("Stop requested while activating window")
            time.sleep(0.1)
    except MacroStopException:
        raise
    except Exception as e:
        msg = str(e)
        if p.get('ignore_fail', False):
            logger.warning(f"  ACTIVATE_WINDOW failed and was ignored: {e}")
            return _ACTION_SKIP
        raise RuntimeError(f"ACTIVATE_WINDOW failed: {e}") from e

    return _ACTION_DONE

def _handle_note_action(p, ctx: RunContext):
    note_text = p.get('text', '')
    if note_text:
        logger.info("  NOTE executed (%s chars)", len(note_text))
    return _ACTION_SKIP


def _handle_set_var_action(p, ctx: RunContext):
    var_name = p.get('var_name', '').strip()
    var_value = _render_param(p, 'var_value', ctx)
    if var_name:
        ctx.set_var(var_name, var_value)
        logger.info("  %s set to %s", var_name, _log_value_summary(var_value))
    return _ACTION_DONE



def _handle_prompt_input_action(p, ctx: RunContext):
    var_name = p.get('var_name', '').strip()
    title = _render_param(p, 'title', ctx, '智点助手输入')
    prompt = _render_param(p, 'prompt', ctx, '请输入内容:')
    default_value = _render_param(p, 'default_value', ctx)
    if not var_name:
        logger.error("  PROMPT_INPUT is missing a variable name")
        return _ACTION_BREAK
    if ctx.check_stop():
        raise MacroStopException("用户在输入前请求停止")
    callback = ctx.get_option('prompt_input_callback')
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
    ctx.set_var(var_name, str(value))
    logger.info("  %s received user input: %s", var_name, _log_value_summary(ctx.get_var(var_name)))

    return _ACTION_DONE

def _handle_read_file_action(p, ctx: RunContext):
    file_path = _render_param(p, 'file_path', ctx)
    var_name = p.get('var_name', '').strip()
    encoding = p.get('encoding', 'utf-8')
    if not file_path or not var_name:
        logger.error("  READ_FILE parameters are incomplete")
        return _ACTION_BREAK

    def _do_read():
        safe_path = _resolve_safe_file_path(file_path, ctx, 'READ_FILE')
        with open(safe_path, 'r', encoding=encoding) as f:
            ctx.set_var(var_name, f.read())
        logger.info(f"  {var_name} loaded text file ({len(ctx.get_var(var_name))} chars)")
        return _ACTION_DONE

    return _run_with_fail_stop('READ_FILE', p, _do_read, ctx, var_name)

def _handle_extract_var_action(p, ctx: RunContext):
    source = _render_param(p, 'source_text', ctx)
    pattern = str(p.get('regex', ''))
    var_name = p.get('var_name', '').strip()
    if not pattern or not var_name:
        logger.error("  EXTRACT_VAR parameters are incomplete")
        return _ACTION_BREAK

    def _do_extract():
        match = re.search(pattern, source)
        val = match.group(1) if match and match.lastindex else (match.group(0) if match else '')
        ctx.set_var(var_name, val)
        logger.info("  %s extracted: %s", var_name, _log_value_summary(val))
        return _ACTION_DONE

    return _run_with_fail_stop('EXTRACT_VAR', p, _do_extract, ctx, var_name)

def _handle_write_file_action(p, ctx: RunContext):
    file_path = _render_param(p, 'file_path', ctx)
    content = _render_param(p, 'content', ctx)
    encoding = p.get('encoding', 'utf-8')
    append = _parse_bool_param(p.get('append', False), False)
    if not file_path:
        logger.error("  WRITE_FILE path is missing")
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


def _resolve_goto_step(ctx: RunContext, pc, goto_labels, action_name, label, max_jumps_value, warn_on_default=False):
    if not label:
        raise RuntimeError(f"{action_name} is missing a label name")
    target = goto_labels.get(_normalize_goto_label(label))
    if not target:
        raise RuntimeError(f"{action_name} label not found: {label}")
    max_jumps, ok = _safe_int(
        max_jumps_value,
        GOTO_LABEL_DEFAULT_MAX_JUMPS,
        min_value=1
    )
    if warn_on_default and not ok:
        _warn_param_default(action_name, 'max_jumps', max_jumps)
    if not ctx.try_record_goto(pc, max_jumps):
        raise RuntimeError(f"{action_name} exceeded max jumps {max_jumps}: {label}")
    return target


def _parse_run_timeout(p):
    if 'timeout_ms' in p:
        timeout_ms, ok = _safe_param_int(p, 'timeout_ms', 30000, min_value=1)
        timeout = timeout_ms / 1000.0
        timeout_key = 'timeout_ms'
    else:
        timeout, ok = _safe_param_int(p, 'timeout', 30, min_value=1)
        timeout_key = 'timeout'
    if not ok:
        logger.warning(
            "  RUN %s must be a positive integer; using default 30000 ms",
            timeout_key,
        )
    return timeout

def _handle_run_action(p, ctx: RunContext):
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


def _handle_calculate_action(p, ctx: RunContext):
    expr = _render_param(p, 'expression', ctx)
    var_name = p.get('var_name', '').strip()
    if not expr or not var_name:
        logger.error("  CALCULATE missing expression or var_name")
        return _ACTION_BREAK
    try:
        val = _safe_calculate_expression(expr)
        ctx.set_var(var_name, str(val))
        logger.info("  %s calculated: %s", var_name, _log_value_summary(val))
    except (ValueError, TypeError, OverflowError, SyntaxError, ZeroDivisionError) as e:
        logger.error("  CALCULATE failed (expression length: %s): %s", len(expr), e)
        if p.get('fail_stop', False):
            return _ACTION_BREAK
        return _ACTION_SKIP
    return _ACTION_DONE

# ======================================================================
# Linear action dispatch
# ======================================================================
_LINEAR_ACTION_HANDLERS = MappingProxyType({
    'CLICK': _handle_click_action,
    'MOVE_TO': _handle_move_to_action,
    'DRAG_TO': _handle_drag_to_action,
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
})


def _dispatch_linear_action(act, p, ctx: RunContext, status_callback=None):
    handler = _LINEAR_ACTION_HANDLERS.get(act)
    if handler is None:
        return None
    if act == 'CLICK':
        return handler(p, ctx, status_callback=status_callback)
    return handler(p, ctx)


def _handle_if_var_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    var_val = _render_param(p, 'var_value', ctx)
    op = p.get('operator', '==')
    expected = _render_param(p, 'expected_val', ctx)
    res_bool = _compare_values(var_val, op, expected)

    if not res_bool:
        logger.info("  -> IF condition not met (%s)", op)
        return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])

    logger.info("  -> IF条件满足 (%s)", op)
    return pc + 1


def _handle_goto_if_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    if loops:
        raise RuntimeError("GOTO_IF 当前版本不允许在 LOOP 循环内部执行")

    var_val = _render_param(p, 'var_value', ctx)
    operator = p.get('operator', '==')
    expected = _render_param(p, 'expected_val', ctx)
    label = _render_param_stripped(p, 'label', ctx)
    res_bool = _compare_values(var_val, operator, expected)

    if res_bool:
        target = _resolve_goto_step(ctx, pc, goto_labels, 'GOTO_IF', label, p.get('max_jumps', GOTO_LABEL_DEFAULT_MAX_JUMPS))
        logger.info("  Condition met (%s), jumping to '%s'", operator, target['name'])
        return target['index']

    logger.info("  -> GOTO_IF condition not met (%s)", operator)
    return pc + 1


def _handle_goto_label_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
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
    logger.info(f"  Jump to label '{target['name']}' -> step {next_pc + 1}")
    return next_pc


def _handle_else_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    return _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['END_IF'])


def _handle_foreach_line_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    return _handle_foreach_line_start(steps, pc, loops, p, ctx, status_callback)


def _handle_loop_start_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    return _handle_loop_start(steps, pc, loops, p, ctx, status_callback)


def _exit_active_loop(loops, ctx: RunContext):
    """Pop the active loop and release its per-loop cache."""
    loops.pop()
    ctx.locator.exit_loop()


def _handle_end_loop_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    if not loops:
        logger.error("END_LOOP has no matching LOOP_START")
        return pc + 1

    top = loops[-1]
    if top.get('kind') == 'foreach_line':
        raise RuntimeError("END_LOOP 不能结束批量处理，请使用 END_FOREACH")

    mode = top.get('mode', 'fixed')
    if mode == 'fixed':
        return top['start']
    if mode not in ('until_image', 'until_text'):
        raise RuntimeError(f"LOOP_START mode is invalid: {mode!r}")

    top['iteration'] += 1
    if status_callback:
        status_callback(f"条件循环 {top['iteration']}（最多 {top['max_iterations']} 次）")

    if top['iteration'] >= top['max_iterations']:
        _exit_active_loop(loops, ctx)
        if status_callback:
            status_callback(f"警告：已达到最大循环次数 {top['max_iterations']}，循环结束")
        logger.warning("  Reached max iterations; exiting loop")
        return pc + 1

    condition_met = _check_loop_condition(top, ctx)
    if condition_met:
        _exit_active_loop(loops, ctx)
        if status_callback:
            status_callback(f"条件已满足，循环在第 {top['iteration']} 次结束")
        logger.info("  条件满足,循环结束")
        return pc + 1

    logger.debug("  Target not found; continuing loop (iteration %s)", top['iteration'])
    time.sleep(LOOP_CHECK_INTERVAL)
    return top['start']


def _handle_end_foreach_control(steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    if loops and loops[-1].get('kind') == 'foreach_line':
        return loops[-1]['start']
    if loops:
        raise RuntimeError("END_FOREACH cannot close a normal loop; use END_LOOP")

    logger.error("END_FOREACH has no matching FOREACH_LINE")
    return pc + 1


# ======================================================================
# Control-flow dispatch
# ======================================================================
_CONTROL_ACTION_HANDLERS = MappingProxyType({
    'FOREACH_LINE': _handle_foreach_line_control,
    'IF_VAR': _handle_if_var_control,
    'GOTO_IF': _handle_goto_if_control,
    'GOTO_LABEL': _handle_goto_label_control,
    'ELSE': _handle_else_control,
    'LOOP_START': _handle_loop_start_control,
    'END_LOOP': _handle_end_loop_control,
    'END_FOREACH': _handle_end_foreach_control,
})


def _dispatch_control_action(act, steps, pc, loops, p, ctx: RunContext, status_callback, goto_labels):
    handler = _CONTROL_ACTION_HANDLERS.get(act)
    if handler is None:
        return None
    return handler(steps, pc, loops, p, ctx, status_callback, goto_labels)


_EXECUTE_STEP_LOG_ACTIONS = frozenset({'LOOP_START', 'END_LOOP', 'FOREACH_LINE', 'END_FOREACH', 'ELSE', 'END_IF', 'RUN', 'NOTE', 'GOTO_LABEL'})


def _get_step_action_and_params(step):
    return step.get('action', ''), step.get('params', {})


def _should_log_step(act, loops, ctx: RunContext):
    return ctx.get_option('debug_steps', False) or not loops or act in _EXECUTE_STEP_LOG_ACTIONS


def _is_find_or_condition_action(act):
    return (act.startswith('FIND_') or act.startswith('IF_')) and act != 'IF_VAR'


def _should_skip_disabled_step(step, act):
    return not step.get('enabled', True) and act not in MacroSchema.CONTROL_FLOW_ACTIONS


def _handle_ai_command_step(p, ctx: RunContext):
    instruction = p.get('instruction', '')
    if not instruction:
        logger.error("  AI instruction is empty")
        return 'break'

    region_bbox = None
    if p.get('region') is not None:
        region_bbox = screen_locator.coerce_bbox(p.get('region'))
        if region_bbox is None:
            raise ValueError(
                f"AI manual region is invalid: {p.get('region')}"
            )
    elif p.get('cache_box') is not None:
        region_bbox = screen_locator.coerce_bbox(p.get('cache_box'))

    logger.info("  Executing AI instruction (%s chars)", len(instruction))
    request = screen_locator.LocateRequest(
        mode='ai',
        region_bbox=region_bbox,
        instruction=instruction,
    )
    result = screen_locator.locate(
        request, cancel_check=_locator_cancel_check(ctx)
    )

    if result.found:
        target_x, target_y = result.position
        logger.info("  AI returned position: (%s, %s)", target_x, target_y)
        duration, ok, duration_key = _time_param_seconds(
            p, 'duration_ms', 'duration', 0.25
        )
        if not ok:
            _warn_param_default('AI_COMMAND', duration_key, duration)
        _move_to_with_stop(target_x, target_y, duration, ctx=ctx)
        ctx.set_last_pos(target_x, target_y)
        return 'done'

    logger.info("  AI target was not found")
    if p.get('fail_stop', True):
        return 'break'
    return 'done'


def _parse_hex_rgb(value):
    """Parse a strict #RRGGBB color value."""
    text = str(value).strip()
    if re.fullmatch(r'#[0-9a-fA-F]{6}', text) is None:
        raise ValueError("target_color must use #RRGGBB format")
    return tuple(int(text[index:index + 2], 16) for index in (1, 3, 5))


def _color_matches(actual, target, tolerance):
    """Return whether every RGB channel is within the inclusive tolerance."""
    return all(abs(int(actual[index]) - int(target[index])) <= tolerance for index in range(3))


def _evaluate_color_condition(p, ctx: RunContext):
    x, x_ok = _safe_param_int(p, 'x')
    y, y_ok = _safe_param_int(p, 'y')
    tolerance, tolerance_ok = _safe_param_int(p, 'tolerance', 10, min_value=0, max_value=255)
    comparison = str(p.get('comparison', 'match')).strip()
    if not x_ok or not y_ok:
        raise ValueError('IF_COLOR_MATCH requires integer x and y coordinates')
    if not tolerance_ok:
        raise ValueError('IF_COLOR_MATCH tolerance must be an integer from 0 to 255')
    if comparison not in {'match', 'not_match'}:
        raise ValueError("IF_COLOR_MATCH comparison must be 'match' or 'not_match'")

    target = _parse_hex_rgb(p.get('target_color', ''))
    _raise_if_stop_requested(ctx, 'Stop requested before color sampling')
    actual = screen_locator.sample_screen_pixel(x, y)
    _raise_if_stop_requested(ctx, 'Stop requested after color sampling')
    matched = _color_matches(actual, target, tolerance)
    result = matched if comparison == 'match' else not matched
    logger.info(
        '  Color at (%s, %s): #%02X%02X%02X, target #%02X%02X%02X, tolerance %s -> %s',
        x, y, *actual, *target, tolerance, result,
    )
    return result


def _handle_find_or_condition_step(steps, pc, act, p, ctx: RunContext):
    next_pc = pc + 1
    if act == 'IF_COLOR_MATCH':
        if not _evaluate_color_condition(p, ctx):
            logger.info("  -> IF condition not met, skipping block")
            next_pc = _find_jump_or_raise(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])
        return next_pc, 'done'

    res = _handle_find_with_retries(act, p, ctx, ctx.locator.is_in_loop())

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
        ctx.set_last_pos(target_x, target_y)
        pyautogui.moveTo(target_x, target_y)

    return next_pc, 'done'


# ======================================================================
# Core execution loop
# ======================================================================
def execute_steps(steps, run_context: RunContext | dict | None = None, status_callback=None):
    logger.info(f"\n--- execution started (Core V{CORE_VERSION}) ---")
    ctx = RunContext(run_context)
    ctx.reset_execution_state()
    ctx.locator.begin_run()
    if not validate_macro_data(steps, allow_unknown_actions=False):
        return False
    try:
        goto_labels = _build_goto_label_table(steps)
    except ValueError as e:
        logger.error(f"  标签配置错误: {e}")
        return False

    default_stop = "Ctrl+F2"
    try:
        s = ctx.get_option('stop_key_str', default_stop)
        stop_key_display = HotkeyUtils.format_hotkey_display(s)
    except Exception:
        stop_key_display = default_stop

    pc, loops = 0, []
    try:
        while pc < len(steps):

            if ctx.check_stop():
                logger.info(f"  用户请求停止 ({stop_key_display})")
                raise MacroStopException("User requested macro stop")

            step = steps[pc]
            ctx.set_current_step_index(pc)
            act, p = _get_step_action_and_params(step)
            if _should_log_step(act, loops, ctx):
                logger.info(f"[{pc+1}] {act}")
            next_pc = pc + 1

            # Skip disabled non-control steps
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
                    action_result = _dispatch_linear_action(act, p, ctx, status_callback)
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
                raise  # Propagate stop immediately
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
        ctx.locator.clear_loop_state()
        logger.info(f"--- 执行结束 ---\n[统计] {ctx.locator.get_stats()}\n")

def _handle_find_with_retries(act, p, ctx: RunContext, in_loop):
    retry_count, ok = _safe_param_int(p, 'retry_count', 0, min_value=0)
    if not ok:
        _warn_param_default(act, 'retry_count', retry_count)
    retry_interval, ok, interval_key = _time_param_seconds(
        p, 'retry_interval_ms', 'retry_interval', 0.0
    )
    if not ok:
        _warn_param_default(act, interval_key, retry_interval)
    return _handle_find(
        act, p, ctx, in_loop, retry_count=retry_count,
        retry_interval=retry_interval,
    )


def _build_find_locate_request(
        act, p, ctx: RunContext, region_bbox=None, preferred_position=None, ocr_cache_key=None):
    is_img = 'IMAGE' in act
    enhanced_mode = ctx.get_option('enhanced_mode', False)
    if is_img:
        confidence, ok = _safe_param_float(
            p, 'confidence', 0.8, min_value=0, max_value=1
        )
        if not ok:
            _warn_param_default(act, 'confidence', confidence)
        return screen_locator.LocateRequest(
            mode='image',
            region_bbox=region_bbox,
            template_path=p.get('path', ''),
            confidence=confidence,
            enhanced_mode=enhanced_mode,
            preferred_position=preferred_position,
        )

    final_engine = (
        ctx.get_option('force_ocr_engine')
        if ctx.get_option('force_ocr_engine') not in (None, '', 'auto')
        else p.get('engine', 'auto')
    )
    return screen_locator.LocateRequest(
        mode='text',
        region_bbox=region_bbox,
        target_text=p.get('text', ''),
        lang=p.get('lang', 'eng'),
        ocr_engine=final_engine,
        ocr_debug=p.get('debug', False),
        enhanced_mode=enhanced_mode,
        cache_key=ocr_cache_key,
    )


def _build_find_signature(act, p, ctx: RunContext):
    """Return a step-scoped key for dynamic regions and loop fast paths."""
    target = p.get('path', p.get('text', ''))
    if p.get('region_mode') == 'relative':
        configured_region = {
            'mode': 'relative',
            'x_offset': p.get('region_x_offset'),
            'y_offset': p.get('region_y_offset'),
            'width': p.get('region_width'),
            'height': p.get('region_height'),
        }
    else:
        configured_region = (
            p.get('region') if p.get('region') is not None else p.get('cache_box')
        )
    try:
        region_key = json.dumps(
            configured_region, ensure_ascii=False, sort_keys=True,
            separators=(',', ':'),
        )
    except (TypeError, ValueError):
        region_key = repr(configured_region)
    return (ctx.get_current_step_index('?'), act, target, region_key)


def _resolve_find_region(p, ctx: RunContext):
    """Resolve an explicitly configured find scope into an absolute BBox.

    Returns ``(region_bbox, constrained)``. Legacy steps without
    ``region_mode`` keep treating ``region`` as an absolute region; steps
    without either field remain eligible for the existing runtime cache path.
    """
    raw_mode = p.get('region_mode')
    if raw_mode is None:
        if p.get('region') is None:
            return None, False
        mode = 'absolute'
    else:
        mode = str(raw_mode).strip().lower()

    if mode == 'full':
        return None, False

    if mode == 'absolute':
        region_bbox = screen_locator.coerce_bbox(p.get('region'))
        if region_bbox is None:
            raise ValueError(f"Manual find region is invalid: {p.get('region')}")
        return region_bbox, True

    if mode != 'relative':
        raise ValueError(f"Unsupported find region mode: {raw_mode}")

    anchor_x, anchor_y = ctx.get_last_locate_pos()
    if anchor_x is None or anchor_y is None:
        raise ValueError(
            'Relative find region requires a previous successful image or text recognition'
        )

    x_offset, x_ok = _safe_param_int(p, 'region_x_offset', 0)
    y_offset, y_ok = _safe_param_int(p, 'region_y_offset', 0)
    width, width_ok = _safe_param_int(p, 'region_width', min_value=1)
    height, height_ok = _safe_param_int(p, 'region_height', min_value=1)
    if not x_ok or not y_ok:
        raise ValueError('Relative find region offsets must be integers')
    if not width_ok or not height_ok:
        raise ValueError('Relative find region width and height must be positive integers')

    left = anchor_x + x_offset
    top = anchor_y + y_offset
    return (left, top, left + width, top + height), True


def _handle_find(
        act, p, ctx: RunContext, in_loop, retry_count=0, retry_interval=0.0):
    _raise_if_stop_requested(ctx, "Stop requested before find")
    is_img = 'IMAGE' in act
    sig = _build_find_signature(act, p, ctx)

    region_bbox, is_constrained_region = _resolve_find_region(p, ctx)
    uses_legacy_cache = p.get('region_mode') is None and p.get('region') is None
    if uses_legacy_cache:
        cache_bbox = ctx.locator.get_runtime_box(sig, p.get('cache_box'))
        region_bbox = screen_locator.expand_bbox(cache_bbox, CACHE_BOX_PADDING)

    cached = ctx.locator.get_preferred_position(sig) if in_loop and is_img else None
    ocr_cache_key = (
        f"step:{ctx.get_current_step_index('?')}:{act}:"
        f"{p.get('text', '')}:{p.get('region', p.get('cache_box', ''))}"
        if not is_img else None
    )
    request = _build_find_locate_request(
        act, p, ctx, region_bbox, cached, ocr_cache_key
    )

    enhanced_mode = ctx.get_option('enhanced_mode', False)
    fallback_request = None
    if (
            region_bbox is not None
            and ctx.get_option('enable_global_fallback', True)
            and enhanced_mode
            and not act.startswith('IF_')
            and not is_constrained_region
    ):
        fallback_request = _build_find_locate_request(
            act, p, ctx, None, None, ocr_cache_key
        )

    outcome = screen_locator.locate_with_policy(
        request,
        screen_locator.LocatePolicy(
            retry_count=retry_count,
            retry_interval=retry_interval,
            fallback_request=fallback_request,
        ),
        cancel_check=_locator_cancel_check(ctx),
    )
    if outcome.fallback_attempted:
        logger.info("  Global screen search...")

    ctx.locator.record_result(outcome.result)
    res = _process_find_result(is_img, p, ctx, outcome.result)
    if res and outcome.fallback_used:
        ctx.locator.set_runtime_box(
            sig, [res[0] - 20, res[1] - 10, res[0] + 20, res[1] + 10]
        )

    if res:
        pos = (res[0], res[1])
        if in_loop:
            ctx.locator.remember_position(sig, pos)
        ctx.set_last_locate_pos(pos)
        ctx.set_last_pos(pos)
        if act in ('FIND_TEXT', 'IF_TEXT_FOUND') and len(res) >= 3:
            save_var = p.get('save_to_var', '').strip()
            if save_var:
                ctx.set_var(save_var, res[2])
                logger.info(
                    "  %s saved OCR result: %s",
                    save_var, _log_value_summary(res[2]),
                )
        return res

    return None


def _process_find_result(is_img, p, ctx: RunContext, locator_result):
    """Apply core-owned statistics and OCR side effects to a locate result."""
    if is_img:
        if locator_result.found:
            x, y = locator_result.position
            logger.info(
                "  Image (%s,%s), template %s",
                x, y, _log_display_text(os.path.basename(p.get('path', ''))),
            )
            return (x, y)
        return None

    if locator_result.found:
        pos = locator_result.position
        text_content = locator_result.recognized_text or p.get('text', '')
        logger.info("  Text (%s,%s), target=%s, recognized=%s", pos[0], pos[1], _log_display_text(p.get('text', '')), _log_display_text(text_content))

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

        # Clipboard side effect
        if p.get('save_to_clipboard', False):
            logger.info("  保留原始文本: %s", _log_value_summary(text_content))
            ctx.set_clipboard_var(final_text)
            try:
                if not _copy_to_clipboard_with_retry(final_text, ctx):
                    raise RuntimeError("clipboard is busy")
                logger.info("  OK copied to clipboard")
            except Exception as e:
                logger.error(f"  失败: {e}")

        return (pos[0], pos[1], final_text)

    return None


def _handle_foreach_line_start(steps, pc, loops, p, ctx: RunContext, cb):
    top = loops[-1] if loops else None

    if top and top.get('kind') == 'foreach_line' and top.get('start') == pc:
        time.sleep(LOOP_PHYSICAL_COOLDOWN)
        next_index = top.get('index', 0) + 1
        if next_index >= len(top.get('items', [])):
            loops.pop()
            if cb:
                cb("批量处理完成")
            logger.info("  All lines processed")
            return _find_jump_or_raise(steps, pc, 'FOREACH_LINE', 'END_FOREACH', ['END_FOREACH'])

        top['index'] = next_index
        _set_foreach_line_vars(top, ctx)
        if cb:
            cb(f"批量处理 {next_index + 1} / {len(top['items'])}")
        logger.info(f"  Line {next_index + 1}/{len(top['items'])}")
        return pc + 1

    if loops:
        raise RuntimeError("FOREACH_LINE cannot be nested inside another loop")

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
        logger.info("  No lines to process; skipping FOREACH_LINE")
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
        cb(f"批量处理 1 / {len(lines)}")
    logger.info(f"  Processing {len(lines)} lines")
    return pc + 1


def _handle_loop_start(steps, pc, loops, p, ctx: RunContext, cb):
    top = loops[-1] if loops else None


    # Re-enter an active loop
    if top and top['start'] == pc:
         # === [修复] 强制给循环加一个物理冷却，防止队列瞬间爆炸 ===
        time.sleep(LOOP_PHYSICAL_COOLDOWN)  # 使用常量
        mode = top.get('mode', 'fixed')

        # Guard every loop mode with the maximum iteration count
        if top['iteration'] >= top['max_iterations']:
            _exit_active_loop(loops, ctx)
            if cb: cb(f"已达到最大循环次数 {top['max_iterations']}，循环结束")
            logger.warning(f"  Reached max iterations: {top['max_iterations']}")
            return _find_jump_or_raise(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])

        # Fixed-count loop
        if mode == 'fixed':
            # iteration starts at 1 and increments on re-entry
            if top['remain'] > 0:
                top['remain'] -= 1
                top['iteration'] += 1
                total = top.get('total_count', top['iteration'] + top['remain'])
                if cb: cb(f"循环 {top['iteration']} / {total}")
                return pc + 1
            else:
                _exit_active_loop(loops, ctx)
                return _find_jump_or_raise(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])

        # Conditional-loop counting and exit checks are handled by END_LOOP
        # Keep all conditional-loop exit decisions in one place
        return pc + 1

    if loops and top and top.get('kind') == 'foreach_line':
        raise RuntimeError("普通循环暂不支持嵌套在批量处理内部")

    # 新循环初始化
    else:
        mode = p.get('mode', 'fixed')
        if mode not in ('fixed', 'until_image', 'until_text'):
            raise ValueError(f"LOOP_START mode is invalid: {mode!r}")

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
            count = max_iter  # Reference count for conditional loops
            remain = max_iter

        loop_id = f"L{pc}_{len(loops)}"
        loop_data = {
            'start': pc,
            'remain': remain,
            'id': loop_id,
            'mode': mode,
            'iteration': 0 if mode in ('until_image', 'until_text') else 1,
            'total_count': count,
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
            # Preserve the configured search region
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
            # Preserve the configured search region
            if 'region' in p:
                logger.info(f"  目标: {loop_data['condition_text']} (区域: {p['region']})")
            elif 'cache_box' in p:
                loop_data['cache_box'] = p['cache_box']
                logger.info(f"  目标: {loop_data['condition_text']} (区域: {p['cache_box']})")
            else:
                logger.info(f"  目标: {loop_data['condition_text']} (全屏)")

        loops.append(loop_data)
        ctx.locator.enter_loop(loop_id)

        if mode == 'fixed':
            if cb: cb(f"循环 1 / {count}")
        else:
            if cb: cb(f"条件循环 1（最多 {max_iter} 次）")

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

def _locate_for_loop_condition(loop_data, ctx: RunContext):
    mode = loop_data.get('mode')
    if loop_data.get('region') is not None:
        region_bbox = screen_locator.coerce_bbox(loop_data.get('region'))
        if region_bbox is None:
            raise LoopConditionCheckError(
                f"Loop region is invalid: {loop_data.get('region')}"
            )
    else:
        region_bbox = screen_locator.expand_bbox(
            loop_data.get('cache_box'), CACHE_BOX_PADDING
        )

    text = loop_data.get('condition_text', '')
    cache_key = (
        f"loop_until_text:{loop_data.get('start', '?')}:{text}:"
        f"{loop_data.get('region', loop_data.get('cache_box', ''))}"
    ) if mode == 'until_text' else None
    try:
        result = screen_locator.locate_condition(
            mode,
            region_bbox=region_bbox,
            template_path=loop_data.get('condition_image', ''),
            target_text=text,
            confidence=loop_data.get('confidence', 0.8),
            lang=loop_data.get('lang', 'eng'),
            enhanced_mode=ctx.get_option('enhanced_mode', False),
            cache_key=cache_key,
            cancel_check=_locator_cancel_check(ctx),
        )
    except Exception as exc:
        kind = 'Image' if mode == 'until_image' else 'Text'
        logger.error("  %s loop condition check failed: %s", kind, exc)
        raise LoopConditionCheckError(
            f"{kind} loop condition check failed"
        ) from exc

    if result.found:
        ctx.set_last_locate_pos(result.position)
        if mode == 'until_image':
            logger.info(
                "  OK found target image: %s",
                os.path.basename(loop_data.get('condition_image', '')),
            )
        else:
            logger.info(
                "  OK found target text: %s",
                _log_value_summary(result.recognized_text or text),
            )
    return result.found


def _check_loop_condition(loop_data, ctx: RunContext):
    """Return whether a configured loop exit condition has been met."""
    mode = loop_data.get('mode', 'fixed')
    if mode in ('until_image', 'until_text'):
        return _locate_for_loop_condition(loop_data, ctx)
    return False

core_engine_version = f"{CORE_VERSION} (Core) / OpenCV: {screen_locator.OPENCV_AVAILABLE}"

# ======================================================================
# RUN handler: execute commands or scripts
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

def _register_active_process(ctx: RunContext, process):
    ctx.register_process(process)


def _unregister_active_process(ctx: RunContext, process):
    ctx.unregister_process(process)

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

def cleanup_active_processes(ctx: RunContext | None):
    if ctx is None:
        return
    for process in ctx.drain_active_processes():
        terminate_process_tree(process)

def _execute_subprocess(cmd_list, shell_mode, cwd, timeout, save_output, ctx: RunContext, run_mode_name):
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
            if ctx.check_stop():
                raise MacroStopException("Stop requested while RUN process is active")
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                terminate_process_tree(process)
                logger.error(f"  {run_mode_name} timed out ({timeout}s)")
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
                ctx.set_clipboard_var(output)
                try:
                    if not _copy_to_clipboard_with_retry(output, ctx):
                        raise RuntimeError("clipboard is busy")
                    logger.info("        Saved to clipboard")
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


def _handle_run(p, ctx: RunContext):
    """Execute a command or script.
    
    Returns True on success, False on failure, and SKIPPED when disabled.
    """
    if not ctx.get_option('run_enabled', False):
        logger.info("  已跳过（执行外部命令默认已禁用，请在设置中手动开启）")
        return 'SKIPPED'

    p = _render_string_params(p, ctx)
    run_type = p.get('run_type', 'command')

    # === 文件写入模式 ===
    if run_type == 'file':
        logger.info("  file mode is disabled; use WRITE_FILE for text output")
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
            logger.error("  Error: command is missing")
            return False

        if shell_mode:
            if not ctx.get_option('allow_shell_mode', False):
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
            logger.error("  Error: script path is missing")
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

        # Interpreter mapping
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
        logger.error(f"  Error: unknown run_type: {run_type}")
        return False


# ======================================================================
# Macro data validation
# ======================================================================
_IF_OPEN_ACTIONS = frozenset({'IF_IMAGE_FOUND', 'IF_TEXT_FOUND', 'IF_COLOR_MATCH', 'IF_VAR'})
_CONTROL_CLOSE_TO_OPEN = {
    'END_IF': _IF_OPEN_ACTIONS,
    'END_LOOP': frozenset({'LOOP_START'}),
    'END_FOREACH': frozenset({'FOREACH_LINE'}),
}
_LOOP_OPEN_ACTIONS = frozenset({'LOOP_START', 'FOREACH_LINE'})


def _validate_control_flow_structure(steps):
    stack = []
    for index, step in enumerate(steps):
        action = step['action']
        if action in _IF_OPEN_ACTIONS:
            stack.append({'action': action, 'step': index, 'else_seen': False})
            continue

        if action == 'LOOP_START':
            if any(frame['action'] == 'FOREACH_LINE' for frame in stack):
                logger.error(
                    "Step %s LOOP_START cannot be nested inside FOREACH_LINE",
                    index + 1,
                )
                return False
            stack.append({'action': action, 'step': index})
            continue

        if action == 'FOREACH_LINE':
            if any(frame['action'] in _LOOP_OPEN_ACTIONS for frame in stack):
                logger.error(
                    "Step %s FOREACH_LINE cannot be nested inside another loop",
                    index + 1,
                )
                return False
            stack.append({'action': action, 'step': index})
            continue

        if action == 'ELSE':
            if not stack or stack[-1]['action'] not in _IF_OPEN_ACTIONS:
                logger.error("Step %s ELSE has no matching IF", index + 1)
                return False
            if stack[-1]['else_seen']:
                logger.error(
                    "Step %s ELSE duplicates the ELSE for IF at step %s",
                    index + 1,
                    stack[-1]['step'] + 1,
                )
                return False
            stack[-1]['else_seen'] = True
            continue

        expected_openers = _CONTROL_CLOSE_TO_OPEN.get(action)
        if expected_openers is None:
            continue
        if not stack or stack[-1]['action'] not in expected_openers:
            logger.error(
                "Step %s %s has no correctly nested opener",
                index + 1,
                action,
            )
            return False
        stack.pop()

    if stack:
        frame = stack[-1]
        logger.error(
            "Step %s %s is missing its closing action",
            frame['step'] + 1,
            frame['action'],
        )
        return False
    return True


def validate_macro_data(data, allow_unknown_actions=True):
    """
    Validate the macro data structure.
    Args:
        data: Data loaded from JSON.
    Returns:
        bool: 数据是否有效
    """
    # Root value must be a list
    if not isinstance(data, list):
        logger.error("Macro root must be a list")
        return False

    # Validate each step
    for i, step in enumerate(data):
        # Every step must be a dictionary
        if not isinstance(step, dict):
            logger.error(f"步骤 {i+1} 不是字典对象")
            return False

        # 必须包含 'action' 字段
        if 'action' not in step:
            logger.error(f"步骤 {i+1} 缺少 'action' 字段")
            return False

        # 必须包含 'params' 字段且为字典
        action = step['action']
        if not isinstance(action, str) or not action.strip():
            logger.error(f"Step {i + 1} action must be a non-empty string")
            return False

        if 'params' not in step or not isinstance(step['params'], dict):
            logger.error(f"Step {i+1} is missing a valid params dictionary")
            return False

        # Validate known action-specific fields
        if action == 'LOOP_START':
            mode = step['params'].get('mode', 'fixed')
            if mode not in ('fixed', 'until_image', 'until_text'):
                logger.error(f"Step {i + 1} LOOP_START mode is invalid: {mode!r}")
                return False

        if action not in MacroSchema.ACTION_TRANSLATIONS:
            log = logger.warning if allow_unknown_actions else logger.error
            log(f"Step {i+1} contains unknown action: {action}")
            if not allow_unknown_actions:
                return False

    return _validate_control_flow_structure(data)
