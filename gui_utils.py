# gui_utils.py
# 描述：GUI 组件工厂与界面逻辑处理 (重构版)
# 版本：1.8.3

import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
import os
import re
import math
import logging
import base64
import queue

logger = logging.getLogger(__name__)

from sys_utils import capture_physical_bbox, center_child_window

# 引入核心库中的工具用于处理快捷键显示
try:
    from core_engine import MacroSchema
except ImportError:
    # Fallback
    class MacroSchema:
        LANG_OPTIONS = {}
        CLICK_OPTIONS = {}
        ACTION_TRANSLATIONS = {}
        RUN_TYPE_OPTIONS = {'command (命令)': 'command', 'script (脚本)': 'script'}
        RUN_TYPE_DISPLAY_BY_VALUE = {v: k for k, v in RUN_TYPE_OPTIONS.items()}

# =================================================================
# 1. 基础工具函数
# =================================================================
LOOP_MODE_OPTIONS = {
    '固定次数': 'fixed',
    '直到找到图像': 'until_image',
    '直到找到文本': 'until_text',
}
LOOP_MODE_DISPLAY_BY_VALUE = {v: k for k, v in LOOP_MODE_OPTIONS.items()}

_LOOP_PARAM_KEYS = ('times', 'condition_image', 'confidence', 'condition_text', 'lang', 'max_iterations', 'region')
_LOOP_VISIBLE_PARAMS = {
    'fixed': ('times',),
    'until_image': ('condition_image', 'confidence', 'max_iterations', 'region'),
    'until_text': ('condition_text', 'lang', 'max_iterations', 'region'),
}
_RUN_COMMAND_PARAMS = ('command', 'args', 'timeout', 'cwd', 'save_output', 'shell_mode', 'fail_stop')
_RUN_SCRIPT_PARAMS = ('script_path', 'interpreter', 'args', 'timeout', 'cwd', 'save_output', 'fail_stop')
_RUN_VISIBLE_PARAMS = {
    'command': _RUN_COMMAND_PARAMS,
    'script': _RUN_SCRIPT_PARAMS,
}
_RUN_PARAM_KEYS = tuple(sorted(set(_RUN_COMMAND_PARAMS + _RUN_SCRIPT_PARAMS)))

_OPTIONAL_PARAM_KEYS = {
    '*': {'region', 'extract_pattern', 'save_to_var'},
    'SET_VAR': {'var_value'},
    'WRITE_FILE': {'content', 'append'},
    'PROMPT_INPUT': {'default_value'},
    'FOREACH_LINE': {'file_path', 'source_text', 'split_delimiter', 'field_names'},
    'IF_VAR': {'var_value', 'expected_val'},
    'GOTO_IF': {'var_value', 'expected_val'},
    'SCROLL': {'x', 'y'},
    'CLICK': {'x', 'y', 'clicks', 'interval', 'duration'},
}
_EMPTY_SKIP_ACTIONS = {'ELSE', 'END_IF', 'END_LOOP', 'END_FOREACH', 'NOTE', 'RUN'}
_NUMERIC_PARAM_KEYS = {'x', 'y', 'ms', 'times', 'x_offset', 'y_offset', 'amount', 'max_iterations', 'max_jumps', 'max_lines', 'retry_count', 'timeout', 'clicks', 'interval', 'duration'}
_NON_NEGATIVE_INT_PARAM_KEYS = {'ms', 'times', 'max_iterations', 'max_jumps', 'max_lines', 'retry_count', 'clicks'}
_NON_NEGATIVE_FLOAT_PARAM_KEYS = {'interval', 'duration'}
_POSITIVE_INT_PARAM_KEYS = {'max_iterations', 'max_jumps', 'max_lines'}
_RUN_DEFAULT_OMIT_PARAMS = {
    'timeout': '30',
    'interpreter': 'python',
    'encoding': 'utf-8',
}
_RUN_FALSE_OMIT_PARAMS = {'append', 'save_output', 'shell_mode', 'fail_stop'}
_LOOP_REQUIRED_PARAMS = {
    'fixed': ('times', "参数 'times' 不能为空"),
    'until_image': ('condition_image', "参数 'condition_image' 不能为空"),
    'until_text': ('condition_text', "参数 'condition_text' 不能为空"),
}


_SIMPLE_ACTION_FORM_FIELDS = {
    'MOVE_OFFSET': (
        ('entry', 'x_offset', 'X 偏移:', '10', None),
        ('entry', 'y_offset', 'Y 偏移:', '0', None),
    ),
    'CLICK': (
        ('combobox', 'button', '按键:', None, lambda: list(MacroSchema.CLICK_OPTIONS.keys())),
        ('entry', 'x', 'X 坐标 (可选, 留空=当前位置):', '', None),
        ('entry', 'y', 'Y 坐标 (可选, 留空=当前位置):', '', None),
        ('entry', 'clicks', '点击次数 (可选, 默认1):', '', None),
        ('entry', 'interval', '点击间隔毫秒 (可选, 默认0):', '', None),
        ('entry', 'duration', '按下持续秒 (可选, 默认0):', '', None),
    ),
    'SCROLL': (
        ('entry', 'amount', '滚动量 (正数=上, 负数=下):', '100', None),
        ('entry', 'x', 'X 坐标 (可选):', '', None),
        ('entry', 'y', 'Y 坐标 (可选):', '', None),
    ),
    'WAIT': (
        ('entry', 'ms', '等待 (毫秒):', '500', None),
    ),
    'TYPE_TEXT': (
        ('entry', 'text', '输入文本:', '你好', None),
    ),
    'PRESS_KEY': (
        ('entry', 'key', '按键或组合键 (Enter, Ctrl+C):', 'Enter', None),
    ),
    'ACTIVATE_WINDOW': (
        ('entry', 'title', '窗口标题 (支持部分匹配):', '记事本', None),
    ),
    'NOTE': (
        ('entry', 'text', '备注内容:', '这里是需要备注的文本...', None),
    ),
    'GOTO_LABEL': (
        ('entry', 'label', '目标标签名:', '重试登录', None),
        ('entry', 'max_jumps', '最大跳转次数(安全阀):', '100', None),
    ),
    'SET_VAR': (
        ('entry', 'var_name', '变量名:', 'my_var', None),
        ('entry', 'var_value', '变量值:', '123', None),
    ),
    'CALCULATE': (
        ('entry', 'expression', '变量计算表达式:', '({price} * 1.5) + 10', None),
        ('entry', 'var_name', '保存至变量:', 'final_price', None),
        ('checkbox', 'fail_stop', '[警告] 计算失败时停止宏', False, None),
    ),
}

_SIMPLE_ACTION_HINTS = {
    'CLICK': '* 提示: X/Y 留空则在当前鼠标位置点击；clicks/interval/duration 留空则使用默认值（1次/0毫秒/0秒）。',
    'SCROLL': '* 提示: 如果 X, Y 为空，将在当前鼠标位置滚动。',
    'TYPE_TEXT': "* 此功能使用剪贴板 (Ctrl+V)，以支持中文及复杂文本输入。\n* 支持占位符: {CLIPBOARD} 将替换为剪贴板内容\n* 示例: '订单号: {CLIPBOARD}' → '订单号: 12345'",
    'ACTIVATE_WINDOW': '* 提示: 宏将查找标题中包含此文本的窗口，并将其激活到最前端。',
    'NOTE': '* 注意: 此步骤仅作为注释，不会执行任何操作。\n* 备注内容以 LABEL: 或 标签: 开头时，可作为“跳转到标签”的目标。\n* 示例: LABEL: 重试登录',
    'SET_VAR': '* 提示: 变量名无需大括号，变量值支持 {其他变量} 插值。',
    'CALCULATE': '* 提示: 用来把已识别或读取到的变量做加减乘除，例如价格、数量、序号；不适合当作普通计算器使用。',
}

_SIMPLE_ACTION_INSTRUCTIONS = {
    'GOTO_LABEL': ('使用说明', (
        '先添加一个“备注”步骤作为标签，例如: LABEL: 重试登录',
        '本步骤只填写标签名，例如: 重试登录',
        '标签名必须唯一；重复或找不到都会停止执行',
        '每个跳转步骤默认最多跳转 100 次，防止死循环',
    )),
}



def _is_optional_empty_param(action_key, param_key):
    return (
        param_key in _OPTIONAL_PARAM_KEYS.get('*', set())
        or param_key in _OPTIONAL_PARAM_KEYS.get(action_key, set())
    )


def _copy_present_params(params, keys):
    return {k: params[k] for k in keys if k in params}


def _is_default_or_empty_param(key, value):
    if key in _RUN_DEFAULT_OMIT_PARAMS and value == _RUN_DEFAULT_OMIT_PARAMS[key]:
        return True
    if key in _RUN_FALSE_OMIT_PARAMS and not value:
        return True
    return value is None or (not isinstance(value, bool) and not str(value).strip())


def _prune_run_params(params, keep_keys):
    return {
        k: v
        for k, v in _copy_present_params(params, keep_keys).items()
        if not _is_default_or_empty_param(k, v)
    }


def _validate_numeric_param(key, value):
    if key not in _NUMERIC_PARAM_KEYS or not value:
        return None

    text = str(value).strip()
    if key in _NON_NEGATIVE_INT_PARAM_KEYS:
        try:
            parsed_int = int(text)
        except (ValueError, TypeError):
            return f"parameter '{key}' must be a non-negative integer"
        if parsed_int < 0:
            return f"parameter '{key}' must be a non-negative integer"
        if key in _POSITIVE_INT_PARAM_KEYS and parsed_int <= 0:
            return f"parameter '{key}' must be greater than 0"
        return None

    if key == 'timeout':
        try:
            parsed_timeout = int(text)
        except (ValueError, TypeError):
            return "parameter 'timeout' must be a positive integer"
        if parsed_timeout <= 0:
            return "parameter 'timeout' must be greater than 0"
        return None

    if key in _NON_NEGATIVE_FLOAT_PARAM_KEYS:
        try:
            parsed_float = float(text)
        except (ValueError, TypeError):
            return f"parameter '{key}' must be a non-negative number"
        if not math.isfinite(parsed_float) or parsed_float < 0:
            return f"parameter '{key}' must be a finite non-negative number"
        return None

    try:
        int(text)
    except (ValueError, TypeError):
        return f"parameter '{key}' must be an integer"
    return None


def _validate_confidence(value):
    if not value:
        return None
    try:
        confidence = float(str(value).strip())
    except (ValueError, TypeError):
        return "参数 'confidence' 必须是数字（如 0.8）"
    if not math.isfinite(confidence) or not (0.0 < confidence <= 1.0):
        return "参数 'confidence' 必须在 0.0 ~ 1.0 之间"
    return None


def _validate_retry_interval(value):
    if not value:
        return None
    try:
        retry_interval = float(str(value).strip())
        if not math.isfinite(retry_interval) or retry_interval < 0:
            return "参数 'retry_interval' 必须是大于等于 0 的有限数字"
    except (ValueError, TypeError):
        return "参数 'retry_interval' 必须是数字（如 0.5）"
    return None


def _convert_param_value(key, value, engine_key_map=None):
    if key == 'mode':
        return LOOP_MODE_OPTIONS.get(value, 'fixed'), None, True
    if key == 'run_type':
        return MacroSchema.RUN_TYPE_OPTIONS.get(value, 'command'), None, True
    if key == 'lang':
        return MacroSchema.LANG_OPTIONS.get(value, 'eng'), None, True
    if key == 'button':
        return MacroSchema.CLICK_OPTIONS.get(value, 'left'), None, True
    if key == 'engine':
        if engine_key_map:
            return engine_key_map.get(value, 'auto'), None, True
        return value.split(' ')[0], None, True
    if key == 'region':
        if not value.strip():
            return None, None, False
        coords = parse_region_string(value)
        if not coords:
            return None, "参数 'region' 格式无效，应为 x1, y1, x2, y2", False
        if coords[2] <= coords[0] or coords[3] <= coords[1]:
            return None, "参数 'region' 必须满足 x2 > x1 且 y2 > y1", False
        return coords, None, True
    if key == 'extract_pattern' and not value.strip():
        return None, None, False
    return value, None, True


def _sanitize_run_params(params):
    run_type = params.get('run_type', 'command')
    if run_type == 'command' and not params.get('command'):
        return None, "RUN 命令类型必须填写 '命令'"
    if run_type == 'script' and not params.get('script_path'):
        return None, "RUN 脚本类型必须填写 '脚本路径'"
    if run_type == 'file':
        return None, 'RUN 文件写入模式已禁用，请改用 WRITE_FILE'

    keep_keys = ('run_type',) + _RUN_VISIBLE_PARAMS.get(run_type, ())
    return _prune_run_params(params, keep_keys), None


def _sanitize_goto_label_params(params):
    label = str(params.get('label', '')).strip()
    if not label:
        return None, "参数 'label' 不能为空"
    max_jumps = str(params.get('max_jumps', '100')).strip() or '100'
    return {'label': label, 'max_jumps': max_jumps}, None


def _sanitize_loop_start_params(params):
    mode = params.get('mode', 'fixed')
    keep_keys = ('mode',) + _LOOP_VISIBLE_PARAMS.get(mode, _LOOP_VISIBLE_PARAMS['fixed'])
    new_params = _copy_present_params(params, keep_keys)

    required_key, required_error = _LOOP_REQUIRED_PARAMS.get(mode, _LOOP_REQUIRED_PARAMS['fixed'])
    if not str(new_params.get(required_key, '')).strip():
        return None, required_error
    return new_params, None


def _sanitize_foreach_line_params(params):
    source_text = str(params.get('source_text', '')).strip()
    file_path = str(params.get('file_path', '')).strip()
    if not source_text and not file_path:
        return None, "批量处理必须填写数据内容/变量，或填写文本数据文件路径"

    max_lines = str(params.get('max_lines', '10000')).strip() or '10000'
    try:
        if int(max_lines) <= 0:
            return None, "参数 'max_lines' 必须大于 0"
    except (ValueError, TypeError):
        return None, "参数 'max_lines' 必须是正整数"

    params['max_lines'] = max_lines
    params['file_path'] = file_path
    params['source_text'] = params.get('source_text', '')
    params['current_line_var'] = str(params.get('current_line_var', 'current_line')).strip() or 'current_line'
    params['index_var'] = str(params.get('index_var', 'loop_index')).strip() or 'loop_index'
    params['total_var'] = str(params.get('total_var', 'loop_total')).strip() or 'loop_total'
    return params, None


def _validate_image_params(params):
    for path_key in ('path', 'condition_image'):
        if path_key not in params or not params[path_key]:
            continue
        image_path = params[path_key]
        if not os.path.exists(image_path):
            return f"文件不存在: {image_path}"
        if not image_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
            return f"文件格式错误 (仅支持图片): {os.path.basename(image_path)}"
    return None


def parse_region_string(region_str):
    if not region_str: return None
    try:
        parts = region_str.replace('，', ',').split(',')
        coords = [int(x.strip()) for x in parts if x.strip()]
        return coords if len(coords) == 4 else None
    except (ValueError, TypeError, IndexError, AttributeError):
        return None


# =================================================================
# 2. 自动换行标签 (AutoWrapLabel)
# =================================================================
class AutoWrapLabel(ttk.Label):
    def __init__(self, master, **kwargs):
        if 'wraplength' not in kwargs:
            kwargs['wraplength'] = 250
        super().__init__(master, **kwargs)
        self._last_width = None
        self.bind('<Configure>', self._on_configure)

    def _on_configure(self, event):
        width = event.width - 15
        if width > 0 and (self._last_width is None or abs(width - self._last_width) > 5):
            self._last_width = width
            self.configure(wraplength=width)





class MultiLineParamText(tk.Text):
    """Text widget with the Entry-like get/delete/insert calls used by form binding."""

    def get(self, *args):
        if not args:
            return super().get("1.0", "end-1c")
        return super().get(*args)

    def delete(self, *args):
        if args and args[0] == 0:
            return super().delete("1.0", "end")
        return super().delete(*args)

    def insert(self, index, chars, *args):
        if index == 0:
            index = "1.0"
        return super().insert(index, chars, *args)

# ======================================================================
# 参数控件工厂类（从 MacroMate.py 迁移）
# ======================================================================
class ParamWidgetFactory:
    """
    参数控件工厂类，封装各类参数输入控件的创建逻辑

    特性：
    - 无状态设计，所有方法都是纯函数
    - 通过构造函数传入必要的依赖（字体、回调函数等）
    - 保持与原 MacroMate 中的方法接口一致
    """

    def __init__(self, font_ui, font_code, ocr_name_map=None):
        """
        初始化参数控件工厂

        Args:
            font_ui: UI 字体配置
            font_code: 代码字体配置
            ocr_name_map: OCR 引擎名称映射字典 {key: name}
        """
        self.font_ui = font_ui
        self.font_code = font_code
        self.ocr_name_map = ocr_name_map or {}

    @staticmethod
    def _noop(*args, **kwargs):
        return None

    @staticmethod
    def _trace_add_for_widget(var, mode, callback):
        token = var.trace_add(mode, callback)
        frame = getattr(var, '_param_frame', None)
        if frame is not None:
            def _cleanup(_event=None, v=var, t=token):
                try:
                    v.trace_remove(mode, t)
                except Exception:
                    pass
            frame.bind('<Destroy>', _cleanup, add='+')
        return token

    def create_param_entry(self, parent, key, label_text, default_value):
        """Create a single-line parameter entry."""
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        entry = ttk.Entry(frame, width=25, font=self.font_ui)
        entry.insert(0, default_value)
        entry.pack(anchor="w", fill=tk.X)
        entry._param_frame = frame
        frame.pack(fill=tk.X, pady=8)
        return entry

    def create_param_text(self, parent, key, label_text, default_value, height=4):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        text = MultiLineParamText(frame, height=height, wrap="word", font=self.font_ui, relief="solid", borderwidth=1, undo=True)
        text.insert("1.0", default_value)
        text.pack(anchor="w", fill=tk.X)
        text._param_frame = frame
        frame.pack(fill=tk.X, pady=8)
        return text

    def create_file_path_entry(self, parent, key, label_text, default_value, mode="open", filetypes=None):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, expand=True)
        entry = ttk.Entry(input_frame, width=25, font=self.font_ui)
        entry.insert(0, default_value)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def browse():
            types = filetypes or [("All", "*.*")]
            if mode == "save":
                selected = filedialog.asksaveasfilename(filetypes=types)
            else:
                selected = filedialog.askopenfilename(filetypes=types)
            if selected:
                entry.delete(0, tk.END)
                entry.insert(0, selected)

        btn = ttk.Button(input_frame, text="浏览...", width=8, command=browse, bootstyle="info-outline")
        btn.pack(side=tk.RIGHT, padx=(5, 0))
        entry._param_frame = frame
        frame.pack(fill=tk.X, pady=8)
        return entry

    def create_param_checkbox(self, parent, key, label_text, default=False):
        """创建复选框"""
        frame = ttk.Frame(parent)
        var = tk.BooleanVar(value=default)
        checkbox = ttk.Checkbutton(frame, text=label_text, variable=var,
                                   bootstyle="primary-round-toggle")
        checkbox.pack(anchor="w")
        frame.pack(fill=tk.X, pady=8)
        var._param_frame = frame
        return var

    def create_param_combobox(self, parent, key, label_text, values, default=None):
        """创建下拉框"""
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=23, font=self.font_ui)
        if default and values and default in values:
            combo.set(default)
        elif values:
            combo.current(0)
        combo.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        return combo

    def create_compact_entry(self, parent, key, label_text, default_value):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui, width=15).pack(side=tk.LEFT, anchor="w")
        entry = ttk.Entry(frame, width=22, font=self.font_ui)
        entry.insert(0, default_value)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        entry._param_frame = frame
        frame.pack(fill=tk.X, pady=3)
        return entry

    def create_compact_combobox(self, parent, key, label_text, values, default=None):
        frame = ttk.Frame(parent)
        ttk.Label(frame, text=label_text, font=self.font_ui, width=15).pack(side=tk.LEFT, anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=20, font=self.font_ui)
        if default and values and default in values:
            combo.set(default)
        elif values:
            combo.current(0)
        combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        frame.pack(fill=tk.X, pady=3)
        return combo

    def create_collapsible_frame(self, parent, title, expanded=False):
        state = {'expanded': expanded}
        content = ttk.Frame(parent)

        def update_button():
            if state['expanded']:
                toggle.configure(text=f"隐藏{title}")
                content.pack(fill=tk.X, pady=(2, 4))
            else:
                toggle.configure(text=f"显示{title}")
                content.pack_forget()

        def toggle_content():
            state['expanded'] = not state['expanded']
            update_button()

        toggle = ttk.Button(
            parent,
            text=f"隐藏{title}" if expanded else f"显示{title}",
            command=toggle_content,
            bootstyle="info-outline",
            padding=(8, 4),
        )
        toggle.pack(anchor="w", fill=tk.X, pady=(6, 2))
        if expanded:
            content.pack(fill=tk.X, pady=(2, 4))
        return content

    def create_ocr_engine_combobox(self, parent, available_ocr_keys):
        """创建 OCR 引擎选择器"""
        combobox_values = ['自动选择 (Auto)']
        for key, name in self.ocr_name_map.items():
            if key in ('auto', 'none'):
                continue
            if key in available_ocr_keys:
                combobox_values.append(name)
            else:
                combobox_values.append(f"{name} (不可用)")

        return self.create_param_combobox(parent, "engine", "OCR 引擎:",
                                          combobox_values, default="自动选择 (Auto)")

    def create_region_selector(self, parent, default_val="", on_select_callback=None, on_preview_callback=None):
        """创建区域选择器"""
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="搜索范围 (x1,y1,x2,y2) [留空=全屏]:", font=self.font_ui).pack(anchor="w")

        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, expand=True)

        entry = ttk.Entry(input_frame, font=self.font_ui)
        entry.insert(0, str(default_val) if default_val else "")

        def on_click():
            if on_select_callback:
                on_select_callback(entry)

        btn = ttk.Button(input_frame, text="框选", width=5,
                         command=on_click,
                         bootstyle="info-outline")
        btn.pack(side=tk.RIGHT, padx=(4, 0))

        def on_preview():
            if on_preview_callback:
                on_preview_callback(entry)

        preview_btn = ttk.Button(
            input_frame,
            text="预览",
            width=5,
            command=on_preview,
            bootstyle="danger-outline",
        )
        preview_btn.pack(side=tk.RIGHT, padx=(4, 0))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        frame.pack(fill=tk.X, pady=8)
        return entry  # 返回 Entry 控件

    def create_browse_button(self, parent, callback):
        """创建浏览按钮"""
        btn = ttk.Button(parent, text="浏览...", command=callback,
                         bootstyle="info-outline", padding=(10, 6))
        btn.pack(anchor="w", fill=tk.X, pady=2)
        return btn

    def create_test_button(self, parent, text, command):
        """创建测试按钮"""
        ttk.Separator(parent, orient='horizontal').pack(fill='x', pady=(15, 5))
        btn = ttk.Button(parent, text=text, command=command,
                         bootstyle="info", padding=(10, 6))
        btn.pack(anchor="w", fill=tk.X, pady=2)
        return btn

    def _split_instruction_items(self, text):
        items = []
        for line in str(text or '').splitlines():
            line = line.strip()
            if not line:
                continue
            line = line.lstrip('*-').strip()
            line = re.sub(r'^\d+\s*[.、]\s*', '', line)
            if line:
                items.append(line)
        return items

    def create_hint_label(self, parent, text, bootstyle="secondary"):
        """创建统一样式的说明面板。"""
        title = "错误" if bootstyle == "danger" else "说明"
        items = self._split_instruction_items(text)
        return self.create_instruction_panel(parent, title, items, numbered=False, bootstyle=bootstyle)

    def create_instruction_panel(self, parent, title, items, numbered=True, bootstyle="secondary"):
        """创建多行说明面板，避免长提示文字在窄面板中挤成一团。"""
        panel = ttk.LabelFrame(parent, text=title, padding=(8, 6))
        panel._is_instruction_panel = True
        panel.pack(fill=tk.X, pady=(10, 6))
        panel.columnconfigure(1, weight=1)

        for idx, text in enumerate(items, start=1):
            row_pad_top = 2 if idx > 1 else 0
            marker = f"{idx}." if numbered else "•"
            number = ttk.Label(
                panel,
                text=marker,
                width=3,
                anchor="ne",
                font=self.font_ui,
                style=f"{bootstyle}.TLabel"
            )
            number.grid(row=idx - 1, column=0, sticky="ne", padx=(0, 4), pady=(row_pad_top, 2))

            item = AutoWrapLabel(
                panel,
                text=text,
                font=self.font_ui,
                style=f"{bootstyle}.TLabel",
                justify="left"
            )
            item.grid(row=idx - 1, column=1, sticky="ew", pady=(row_pad_top, 2))

        return panel


    def _add_find_retry_options(self, parent_frame, param_widgets, include_ignore_fail=False):
        param_widgets['retry_count'] = self.create_param_entry(parent_frame, "retry_count", "失败重试次数:", "0")
        param_widgets['retry_interval'] = self.create_param_entry(parent_frame, "retry_interval", "重试间隔(秒):", "0.5")
        if include_ignore_fail:
            param_widgets['ignore_fail'] = self.create_param_checkbox(parent_frame, "ignore_fail", "[OK] 找不到时继续执行", default=False)

    def _add_text_capture_options(self, parent_frame, param_widgets):
        param_widgets['save_to_clipboard'] = self.create_param_checkbox(parent_frame, "save_to_clipboard", "[OK] 保存识别结果到剪贴板", default=False)
        param_widgets['save_to_var'] = self.create_param_entry(parent_frame, "save_to_var", "保存至变量 (可选):", "")
        sub_frame = ttk.Frame(parent_frame)
        sub_frame.pack(fill=tk.X)
        extract_frame = ttk.Frame(sub_frame)
        ttk.Label(extract_frame, text="提取模式 (正则，可选):", font=self.font_ui).pack(anchor="w")
        extract_entry = ttk.Entry(extract_frame, width=25, font=self.font_ui)
        extract_entry.insert(0, r"\d+")
        extract_entry.pack(anchor="w", fill=tk.X)
        param_widgets['extract_pattern'] = extract_entry
        hint = AutoWrapLabel(sub_frame, text="提取模式: 用正则表达式过滤识别结果，如 \\d+ 只提取数字；留空则保存全部文本。", font=self.font_ui, style="secondary.TLabel")

        def toggle(var=param_widgets['save_to_clipboard'], ef=extract_frame, hint_label=hint):
            if var.get():
                ef.pack(fill=tk.X, pady=8); hint_label.pack(anchor="w", pady=5, fill=tk.X)
            else:
                ef.pack_forget(); hint_label.pack_forget()
        self._trace_add_for_widget(param_widgets['save_to_clipboard'], 'write', lambda *_: toggle())


    def _build_image_find_form(self, action_key, parent_frame, param_widgets, on_select_region,
                               browse_image_cb, on_test_find_image, on_preview_region=None):
        is_if_action = action_key == 'IF_IMAGE_FOUND'
        param_widgets['path'] = self.create_param_entry(parent_frame, "path", "图像路径:", "button.png")
        param_widgets['region'] = self.create_region_selector(parent_frame, "", on_select_region, on_preview_region)
        confidence_label = "置信度:" if is_if_action else "置信度(0.1-1.0):"
        param_widgets['confidence'] = self.create_param_entry(parent_frame, "confidence", confidence_label, "0.8")
        self._add_find_retry_options(parent_frame, param_widgets, include_ignore_fail=not is_if_action)
        if not is_if_action:
            self.create_hint_label(parent_frame, "* 提示：如果识别失败，请调低置信度")
        self.create_browse_button(parent_frame, browse_image_cb)
        test_label = "🧪 测试 IF 图像" if is_if_action else "🧪 测试查找图像"
        self.create_test_button(parent_frame, test_label, on_test_find_image)

    def _build_text_find_form(self, action_key, parent_frame, param_widgets, available_ocr_keys,
                              on_select_region, on_test_find_text, on_preview_region=None):
        is_if_action = action_key == 'IF_TEXT_FOUND'
        label_text = "查找文本:" if is_if_action else "查找的文本:"
        param_widgets['text'] = self.create_param_entry(parent_frame, "text", label_text, "确定")
        param_widgets['region'] = self.create_region_selector(parent_frame, "", on_select_region, on_preview_region)
        param_widgets['lang'] = self.create_param_combobox(parent_frame, "lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
        param_widgets['engine'] = self.create_ocr_engine_combobox(parent_frame, available_ocr_keys)
        self._add_find_retry_options(parent_frame, param_widgets, include_ignore_fail=not is_if_action)
        self._add_text_capture_options(parent_frame, param_widgets)
        test_label = "🧪 测试 IF 文本" if is_if_action else "🧪 测试查找文本 (OCR)"
        self.create_test_button(parent_frame, test_label, on_test_find_text)

    def _build_var_compare_form(self, parent_frame, param_widgets, left_default, right_default,
                                include_goto_fields=False):
        param_widgets['var_value'] = self.create_param_entry(parent_frame, "var_value", '比较左侧:', left_default)
        op_options = ['==', '!=', '>', '<', '>=', '<=', '包含', '不包含']
        param_widgets['operator'] = self.create_param_combobox(parent_frame, "operator", '操作符:', op_options, default="==")
        param_widgets['expected_val'] = self.create_param_entry(parent_frame, "expected_val", '比较右侧:', right_default)
        if include_goto_fields:
            param_widgets['label'] = self.create_param_entry(parent_frame, "label", '条件成立时跳转至标签:', "Next_Item")
            param_widgets['max_jumps'] = self.create_param_entry(parent_frame, "max_jumps", '最大跳转次数 (防死循环):', "100")

    def _build_simple_action_form(self, action_key, parent_frame, param_widgets):
        for field_type, key, label, default, options_factory in _SIMPLE_ACTION_FORM_FIELDS[action_key]:
            if field_type == 'entry':
                param_widgets[key] = self.create_param_entry(parent_frame, key, label, default)
            elif field_type == 'combobox':
                options = options_factory() if callable(options_factory) else options_factory
                param_widgets[key] = self.create_param_combobox(parent_frame, key, label, options, default=default)
            elif field_type == 'checkbox':
                param_widgets[key] = self.create_param_checkbox(parent_frame, key, label, default=default)

        hint_text = _SIMPLE_ACTION_HINTS.get(action_key)
        if hint_text:
            self.create_hint_label(parent_frame, hint_text)

        instruction = _SIMPLE_ACTION_INSTRUCTIONS.get(action_key)
        if instruction:
            title, lines = instruction
            self.create_instruction_panel(parent_frame, title, list(lines))


    def _build_ai_command_form(self, parent_frame, param_widgets, on_select_region, on_test_ai_command, on_preview_region=None):
        param_widgets['instruction'] = self.create_param_text(parent_frame, "instruction", "AI 指令:", "点击列表里价格最低的那个商品", height=3)
        param_widgets['region'] = self.create_region_selector(parent_frame, "", on_select_region, on_preview_region)
        self.create_hint_label(parent_frame, "* 提示: 输入自然语言指令，如 '点击确定按钮'\n* AI 会分析屏幕截图，理解指令并返回坐标\n* 支持: OpenAI, Anthropic, DeepSeek, 智谱, 通义千问等")
        self.create_test_button(parent_frame, "🧪 测试 AI 指令", on_test_ai_command)

    def _build_run_form(self, parent_frame, param_widgets, update_run_params_cb):
        run_type_options = MacroSchema.RUN_TYPE_OPTIONS
        param_widgets['run_type'] = self.create_param_combobox(parent_frame, "run_type", "类型:", list(run_type_options.keys()), default=MacroSchema.RUN_TYPE_DISPLAY_BY_VALUE.get('command', 'command (命令)'))
        param_widgets['command'] = self.create_param_entry(parent_frame, "command", "命令:", "curl")
        param_widgets['args'] = self.create_param_text(parent_frame, "args", "参数:", "", height=3)
        param_widgets['script_path'] = self.create_file_path_entry(parent_frame, "script_path", "脚本路径:", "process.py", filetypes=[("Scripts", "*.py *.js *.ps1"), ("All", "*.*")])
        param_widgets['interpreter'] = self.create_param_combobox(parent_frame, "interpreter", "解释器:", ["python", "node", "powershell"], default="python")
        param_widgets['timeout'] = self.create_param_entry(parent_frame, "timeout", "超时(秒):", "30")
        param_widgets['cwd'] = self.create_param_entry(parent_frame, "cwd", "工作目录:", "")
        param_widgets['save_output'] = self.create_param_checkbox(parent_frame, "save_output", "[OK] 保存输出到剪贴板", default=False)
        param_widgets['shell_mode'] = self.create_param_checkbox(parent_frame, "shell_mode", "[警告] shell 模式 (仅可信宏)", default=False)
        param_widgets['fail_stop'] = self.create_param_checkbox(parent_frame, "fail_stop", "[警告] RUN 失败时停止宏", default=False)
        self.create_hint_label(parent_frame, "* {CLIPBOARD} = 剪贴板内容, {DATETIME} = 当前时间")
        param_widgets['run_type'].bind("<<ComboboxSelected>>", update_run_params_cb)
        update_run_params_cb(None)

    def _build_move_to_form(self, parent_frame, param_widgets, mouse_tracker, mouse_pos_var):
        param_widgets['x'] = self.create_param_entry(parent_frame, "x", "X 坐标:", "100")
        param_widgets['y'] = self.create_param_entry(parent_frame, "y", "Y 坐标:", "100")
        ttk.Separator(parent_frame, orient='horizontal').pack(fill='x', pady=(15, 5))
        ttk.Label(parent_frame, text="当前鼠标位置 (参考):", font=self.font_ui, foreground='gray').pack(anchor="w", pady=(5, 0))
        if mouse_pos_var is not None:
            ttk.Label(parent_frame, textvariable=mouse_pos_var, font=self.font_code, bootstyle="info").pack(anchor="w")
        if mouse_tracker is not None:
            mouse_tracker.start()

    def _build_loop_start_form(self, parent_frame, param_widgets, on_select_region, update_loop_params_cb, on_preview_region=None):
        param_widgets['mode'] = self.create_param_combobox(parent_frame, "mode", '循环模式:', list(LOOP_MODE_OPTIONS.keys()), default=LOOP_MODE_DISPLAY_BY_VALUE.get('fixed', '固定次数'))
        param_widgets['times'] = self.create_param_entry(parent_frame, "times", "循环次数:", "10")
        param_widgets['max_iterations'] = self.create_param_entry(parent_frame, "max_iterations", "最大迭代次数 (安全阀):", "1000")
        param_widgets['condition_image'] = self.create_param_entry(parent_frame, "condition_image", "目标图像路径:", "target.png")
        param_widgets['confidence'] = self.create_param_entry(parent_frame, "confidence", "置信度:", "0.8")
        param_widgets['condition_text'] = self.create_param_entry(parent_frame, "condition_text", "目标文本:", "加载完成")
        param_widgets['lang'] = self.create_param_combobox(parent_frame, "lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
        param_widgets['region'] = self.create_region_selector(parent_frame, "", on_select_region, on_preview_region)
        self.create_hint_label(parent_frame, "* 提示:\n- 固定次数: 传统循环\n- 直到找到图像: 找到即停\n- 直到找到文本: 找到即停\n- 最大迭代: 安全机制")
        param_widgets['mode'].bind("<<ComboboxSelected>>", update_loop_params_cb)
        update_loop_params_cb(None)

    def build_action_form(self, action_key, parent_frame, param_widgets, available_ocr_keys, callbacks):
        """
        构建指定动作类型的参数表单界面
        
        Args:
            action_key: 动作类型键名 (如 'FIND_IMAGE')
            parent_frame: 容纳参数控件的父容器
            param_widgets: 用于存储控件引用的字典
            available_ocr_keys: 可用的 OCR 引擎列表
            callbacks: 包含各类回调函数的字典
        """
        # 快捷访问回调
        on_select_region = callbacks.get('on_select_region')
        on_preview_region = callbacks.get('on_preview_region')
        browse_image_cb = callbacks.get('browse_image')
        on_test_find_image = callbacks.get('on_test_find_image_click')
        on_test_find_text = callbacks.get('on_test_find_text_click')
        on_test_ai_command = callbacks.get('on_test_ai_command_click')
        update_loop_params_cb = callbacks.get('update_loop_params') or self._noop
        update_run_params_cb = callbacks.get('update_run_params') or self._noop
        mouse_tracker = callbacks.get('mouse_tracker')
        mouse_pos_var = callbacks.get('mouse_pos_var')

        if action_key in ('FIND_TEXT', 'IF_TEXT_FOUND'):
            if 'none' in available_ocr_keys:
                self.create_hint_label(parent_frame, 
                    "FAIL 错误: 未找到可用的OCR引擎。\n"
                    "请先安装 RapidOCR (推荐) 或 Tesseract，\n"
                    "然后重启本程序。",
                    bootstyle="danger")
                # 这里不能直接修改 action_type，因为它是 UI 层的变量。
                # 告知调用者需要切回图像模式。
                return "SWITCH_TO_FIND_IMAGE"

        if action_key == 'FIND_IMAGE':
            self._build_image_find_form(action_key, parent_frame, param_widgets, on_select_region, browse_image_cb, on_test_find_image, on_preview_region)

        elif action_key == 'FIND_TEXT':
            self._build_text_find_form(action_key, parent_frame, param_widgets, available_ocr_keys, on_select_region, on_test_find_text, on_preview_region)

        elif action_key in _SIMPLE_ACTION_FORM_FIELDS:
            self._build_simple_action_form(action_key, parent_frame, param_widgets)

        elif action_key == 'AI_COMMAND':
            self._build_ai_command_form(parent_frame, param_widgets, on_select_region, on_test_ai_command, on_preview_region)

        elif action_key == 'RUN':
            self._build_run_form(parent_frame, param_widgets, update_run_params_cb)

        elif action_key == 'MOVE_TO':
            self._build_move_to_form(parent_frame, param_widgets, mouse_tracker, mouse_pos_var)

        elif action_key == 'IF_IMAGE_FOUND':
            self._build_image_find_form(action_key, parent_frame, param_widgets, on_select_region, browse_image_cb, on_test_find_image, on_preview_region)

        elif action_key == 'IF_TEXT_FOUND':
            self._build_text_find_form(action_key, parent_frame, param_widgets, available_ocr_keys, on_select_region, on_test_find_text, on_preview_region)

        elif action_key == 'LOOP_START':
            self._build_loop_start_form(parent_frame, param_widgets, on_select_region, update_loop_params_cb, on_preview_region)

        elif action_key == 'ELSE':
            self.create_hint_label(parent_frame, "* 提示: 'ELSE' 必须与 'IF' 配合使用。")
        elif action_key == 'END_IF':
            self.create_hint_label(parent_frame, "* 提示: 'END_IF' 标志着逻辑块结束。")
        elif action_key == 'END_LOOP':
            self.create_hint_label(parent_frame, "* 提示: 'END_LOOP' 标志着循环结束。")
        elif action_key == 'END_FOREACH':
            self.create_hint_label(parent_frame, "* 提示: 'END_FOREACH' 标志着批量处理块结束。")
            
        elif action_key == 'READ_FILE':
            param_widgets['file_path'] = self.create_file_path_entry(parent_frame, "file_path", "文本文件路径:", "C:\\test.txt", filetypes=[("Text", "*.txt *.log *.csv *.json *.jsonl *.md *.xml *.yaml *.yml *.ini *.cfg"), ("All", "*.*")])
            param_widgets['var_name'] = self.create_param_entry(parent_frame, "var_name", "保存至变量:", "file_content")
            param_widgets['encoding'] = self.create_param_combobox(parent_frame, "encoding", "编码:", ["utf-8", "gbk", "gb2312"], default="utf-8")
            param_widgets['fail_stop'] = self.create_param_checkbox(parent_frame, "fail_stop", "[警告] 读取失败时停止宏", default=False)
            
        elif action_key == 'EXTRACT_VAR':
            param_widgets['source_text'] = self.create_param_text(parent_frame, "source_text", "源文本:", "{ocr_result}", height=3)
            param_widgets['regex'] = self.create_param_entry(parent_frame, "regex", "正则表达式:", r"\d+\.\d+")
            param_widgets['var_name'] = self.create_param_entry(parent_frame, "var_name", "保存至变量:", "price")
            param_widgets['fail_stop'] = self.create_param_checkbox(parent_frame, "fail_stop", "[警告] 提取失败时停止宏", default=False)
            self.create_hint_label(parent_frame, "* 提示: 使用正则表达式提取，例如 \\d+ 提取纯数字。")

        elif action_key == 'PROMPT_INPUT':
            param_widgets['title'] = self.create_param_entry(parent_frame, "title", "询问窗口标题:", "智点助手人工输入")
            param_widgets['prompt'] = self.create_param_entry(parent_frame, "prompt", "询问内容:", "请输入验证码")
            param_widgets['default_value'] = self.create_param_entry(parent_frame, "default_value", "默认值(可选):", "")
            param_widgets['var_name'] = self.create_param_entry(parent_frame, "var_name", "保存至变量:", "user_input")
            self.create_hint_label(parent_frame, "* 提示: 这是智点助手主动询问用户，不是识别其他软件或网页弹窗；取消输入会安全停止宏。")

        elif action_key == 'FOREACH_LINE':
            param_widgets['file_path'] = self.create_file_path_entry(parent_frame, "file_path", "文本数据文件:", "", filetypes=[("Text", "*.txt *.log *.csv *.tsv"), ("All", "*.*")])
            param_widgets['source_text'] = self.create_param_text(parent_frame, "source_text", "数据内容:", "{file_content}", height=3)
            param_widgets['current_line_var'] = self.create_compact_entry(parent_frame, "current_line_var", "当前行变量:", "current_line")
            param_widgets['split_delimiter'] = self.create_compact_entry(parent_frame, "split_delimiter", "分隔符:", ",")
            param_widgets['field_names'] = self.create_compact_entry(parent_frame, "field_names", "字段变量:", "account,password")
            param_widgets['skip_empty'] = self.create_param_checkbox(parent_frame, "skip_empty", "[OK] 跳过空行", default=True)
            self.create_hint_label(parent_frame, "* 每次取一行执行后续步骤；拆分后可直接使用 {account}、{password}。")

            advanced_frame = self.create_collapsible_frame(parent_frame, "高级选项", expanded=False)
            param_widgets['encoding'] = self.create_compact_combobox(advanced_frame, "encoding", "文本编码:", ["utf-8", "gbk", "gb2312"], default="utf-8")
            param_widgets['index_var'] = self.create_compact_entry(advanced_frame, "index_var", "序号变量:", "loop_index")
            param_widgets['total_var'] = self.create_compact_entry(advanced_frame, "total_var", "总数变量:", "loop_total")
            param_widgets['max_lines'] = self.create_compact_entry(advanced_frame, "max_lines", "最大行数:", "10000")
            param_widgets['strip_fields'] = self.create_param_checkbox(advanced_frame, "strip_fields", "[OK] 去掉字段前后空格", default=True)
             
        elif action_key == 'IF_VAR':
            self._build_var_compare_form(parent_frame, param_widgets, "{price}", "100")

        elif action_key == 'WRITE_FILE':
            param_widgets['file_path'] = self.create_file_path_entry(parent_frame, "file_path", "文本文件路径:", "C:\\log.txt", mode="save", filetypes=[("Text", "*.txt *.log *.csv *.json *.jsonl *.md *.xml *.yaml *.yml *.ini *.cfg"), ("All", "*.*")])
            param_widgets['content'] = self.create_param_text(parent_frame, "content", "写入文本:", "{date} - {result}", height=4)
            param_widgets['encoding'] = self.create_param_combobox(parent_frame, "encoding", "编码:", ["utf-8", "gbk", "gb2312"], default="utf-8")
            param_widgets['append'] = self.create_param_checkbox(parent_frame, "append", "[OK] 追加模式 (不覆盖原文本文件)", default=True)
            param_widgets['fail_stop'] = self.create_param_checkbox(parent_frame, "fail_stop", "[警告] 写入失败时停止宏", default=False)
            
        elif action_key == 'GOTO_IF':
            self._build_var_compare_form(parent_frame, param_widgets, "{status}", '缺货', include_goto_fields=True)

        return None

    def collect_step_data(self, action_key, param_widgets, engine_key_map=None):
        """
        从参数控件中收集并校验数据

        Returns:
            params: 整理后的参数字典
            error: 错误消息，若无错误则为 None
        """
        params = {}
        try:
            for k, w in param_widgets.items():
                if isinstance(w, tk.BooleanVar):
                    params[k] = w.get()
                    continue

                val = w.get()

                error = _validate_numeric_param(k, val)
                if error:
                    return None, error

                if k == 'confidence':
                    error = _validate_confidence(val)
                    if error:
                        return None, error

                if k == 'retry_interval':
                    error = _validate_retry_interval(val)
                    if error:
                        return None, error

                if not val:
                    if action_key in _EMPTY_SKIP_ACTIONS:
                        continue
                    if not _is_optional_empty_param(action_key, k):
                        return None, f"参数 '{k}' 不能为空"

                converted_value, error, include_param = _convert_param_value(k, val, engine_key_map)
                if error:
                    return None, error
                if include_param:
                    params[k] = converted_value

            if action_key == 'RUN':
                params, error = _sanitize_run_params(params)
                if error:
                    return None, error

            if action_key == 'GOTO_LABEL':
                params, error = _sanitize_goto_label_params(params)
                if error:
                    return None, error

            if action_key == 'LOOP_START':
                params, error = _sanitize_loop_start_params(params)
                if error:
                    return None, error

            if action_key == 'FOREACH_LINE':
                params, error = _sanitize_foreach_line_params(params)
                if error:
                    return None, error

            error = _validate_image_params(params)
            if error:
                return None, error

            return params, None
        except Exception as e:
            return None, str(e)


# ======================================================================
# 参数动态显示控制函数（从 MacroMate.py 迁移）
# ======================================================================
def update_loop_params(param_widgets, param_frame, mode_widget):
    """Show or hide loop parameters by mode."""
    if 'mode' not in param_widgets:
        return

    mode = LOOP_MODE_OPTIONS.get(mode_widget.get(), 'fixed')
    _set_param_visibility(
        param_widgets,
        param_frame,
        _LOOP_PARAM_KEYS,
        _LOOP_VISIBLE_PARAMS.get(mode, ()),
    )

def _get_widget_frame(widget):
    """Return the outer frame used to show or hide a parameter widget."""
    frame = getattr(widget, '_param_frame', None)
    if frame is not None:
        return frame
    if isinstance(widget, tk.BooleanVar):
        return None
    return widget.master

def _collect_instruction_widgets(param_frame):
    return [
        widget for widget in param_frame.winfo_children()
        if isinstance(widget, AutoWrapLabel) or getattr(widget, '_is_instruction_panel', False)
    ]


def _set_param_visibility(param_widgets, param_frame, all_keys, visible_keys):
    instruction_widgets = _collect_instruction_widgets(param_frame)

    for key in all_keys:
        if key in param_widgets:
            parent_frame = _get_widget_frame(param_widgets[key])
            if parent_frame:
                parent_frame.pack_forget()

    for key in visible_keys:
        if key in param_widgets:
            frame = _get_widget_frame(param_widgets[key])
            if frame:
                frame.pack(fill=tk.X, pady=8)

    for instruction_widget in instruction_widgets:
        instruction_widget.pack_forget()
        instruction_widget.pack(anchor="w", pady=(10, 6), fill=tk.X)


def update_run_params(param_widgets, param_frame, run_type_widget):
    """Show or hide RUN parameters by type."""
    if 'run_type' not in param_widgets:
        return

    run_type = MacroSchema.RUN_TYPE_OPTIONS.get(run_type_widget.get(), 'command')
    _set_param_visibility(param_widgets, param_frame, _RUN_PARAM_KEYS, _RUN_VISIBLE_PARAMS.get(run_type, ()))

# ======================================================================
# 参数转换工具函数（从 MacroMate.py 迁移）
# ======================================================================
def param_internal_to_display(key, internal_value, ocr_name_map, lang_values_to_name,
                              click_values_to_name, available_ocr_keys=None):
    """
    将内部存储值转换为UI显示值

    Args:
        key: 参数键名
        internal_value: 内部存储的值
        ocr_name_map: OCR引擎名称映射 {key: display_name}
        lang_values_to_name: 语言值到显示名的映射
        click_values_to_name: 点击值到显示名的映射
        available_ocr_keys: 可用的OCR引擎key列表

    Returns:
        UI中应该显示的值
    """
    reverse_mappings = {
        'lang': lang_values_to_name,
        'button': click_values_to_name,
        'engine': ocr_name_map
    }

    mapping = reverse_mappings.get(key)
    if mapping:
        display_val = mapping.get(internal_value, internal_value)

        if key == 'engine' and available_ocr_keys:
            if internal_value not in available_ocr_keys and internal_value != 'auto':
                display_val = f"{display_val} (不可用)"

        return display_val

    return internal_value


# =================================================================
# VLM settings dialog
# =================================================================
class VLMSettingsDialog:
    # [修复 BUG-3] 恢复完整功能：load_config、测试连接、timeout 设置
    def __init__(self, parent):
        self.result = None
        self._ui_callbacks = queue.Queue()
        try:
            import vlm_engine
            self.current_config = vlm_engine.load_config()
            self.providers = vlm_engine.get_providers()
        except Exception as e:
            logger.warning("load failed, using defaults: %s", e)
            self.current_config = {'provider': 'openai', 'api_key': '', 'model': '', 'timeout': 30, 'base_url': ''}
            self.providers = {}
        self.font_ui = ("Microsoft YaHei UI", 10)

        self.dialog = tk.Toplevel(parent)
        self.dialog.withdraw()  # 立即隐藏，防止闪烁
        self.dialog.title("🤖 AI 配置设置")
        self.dialog.geometry("520x660")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        center_child_window(parent, self.dialog)
        self.dialog.deiconify()  # 位置确定后再显示

        self._create_ui()
        self.dialog.after(50, self._drain_ui_callbacks)

    def _create_ui(self):
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="🤖 AI 大模型配置",
                  font=("Microsoft YaHei UI", 12, "bold")).pack(pady=(0, 15))

        # 提供商
        provider_frame = ttk.Labelframe(main_frame, text="AI 提供商", padding=10)
        provider_frame.pack(fill=tk.X, pady=(0, 10))
        self.provider_var = tk.StringVar(value=self.current_config.get('provider', 'openai'))
        provider_names = [f"{v['name']} ({k})" for k, v in self.providers.items()] if self.providers \
            else ['openai', 'anthropic', 'deepseek', 'zhipu', 'qianwen', 'openrouter', 'step']
        self.provider_combo = ttk.Combobox(provider_frame, values=provider_names,
                                            state="readonly", textvariable=self.provider_var,
                                            font=self.font_ui)
        self.provider_combo.pack(fill=tk.X)
        self.provider_combo.bind("<<ComboboxSelected>>", self._on_provider_change)

        # API Key
        key_frame = ttk.Labelframe(main_frame, text="API Key", padding=10)
        key_frame.pack(fill=tk.X, pady=(0, 10))
        self.api_key_var = tk.StringVar(value=self.current_config.get('api_key', ''))
        key_entry = ttk.Entry(key_frame, textvariable=self.api_key_var,
                              font=self.font_ui, show="*")
        key_entry.pack(fill=tk.X)
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(key_frame, text="显示 API Key", variable=self.show_key_var,
                        command=lambda: key_entry.config(
                            show="" if self.show_key_var.get() else "*")).pack(anchor="w", pady=(5, 0))

        # 模型
        model_frame = ttk.Labelframe(main_frame, text="模型 (可选)", padding=10)
        model_frame.pack(fill=tk.X, pady=(0, 10))
        self.model_var = tk.StringVar(value=self.current_config.get('model', ''))
        ttk.Entry(model_frame, textvariable=self.model_var, font=self.font_ui).pack(fill=tk.X)
        ttk.Label(model_frame, text="留空则使用默认值",
                  font=("Microsoft YaHei UI", 8), foreground="gray").pack(anchor="w")

        # 超时
        timeout_frame = ttk.Labelframe(main_frame, text="超时时间 (秒)", padding=10)
        timeout_frame.pack(fill=tk.X, pady=(0, 10))
        self.timeout_var = tk.IntVar(value=self.current_config.get('timeout', 30))
        ttk.Spinbox(timeout_frame, from_=10, to=120,
                    textvariable=self.timeout_var, font=self.font_ui).pack(fill=tk.X)

        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        ttk.Button(btn_frame, text="取消", command=self.dialog.destroy,
                   bootstyle="secondary", padding=(10, 8)).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ttk.Button(btn_frame, text="测试连接", command=self._test_connection,
                   bootstyle="info", padding=(10, 8)).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(btn_frame, text="保存", command=self._save,
                   bootstyle="primary", padding=(10, 8)).grid(row=0, column=2, sticky="ew", padx=(3, 0))

        ttk.Label(main_frame, text="输入 API Key，选择提供商，保存即可使用 AI 指令动作",
                  font=("Microsoft YaHei UI", 8), foreground="#666").pack(pady=(10, 0))

    def _selected_provider_key(self):
        selected = self.provider_var.get()
        return selected.split(" (")[-1].rstrip(")") if "(" in selected else selected

    def _build_current_vlm_config(self, system_prompt=None):
        import vlm_engine
        provider_key = self._selected_provider_key()
        config = vlm_engine.DEFAULT_CONFIG.copy()
        config.update({
            'provider': provider_key,
            'api_key': self.api_key_var.get().strip(),
            'timeout': self.timeout_var.get(),
        })
        if system_prompt is not None:
            config['system_prompt'] = system_prompt
        if self.model_var.get().strip():
            config['model'] = self.model_var.get().strip()
        if self.providers and provider_key in self.providers:
            config['base_url'] = self.providers[provider_key].get('base_url', '')
        return config

    def _on_provider_change(self, event):
        provider_key = self._selected_provider_key()
        if self.providers and provider_key in self.providers:
            default_model = self.providers[provider_key].get('model', '')
            if not self.model_var.get():
                self.model_var.set(default_model)

    def _test_connection(self):
        # [修复H-6] 改为后台线程执行，避免阻塞 UI 事件循环
        try:
            import vlm_engine, io, threading
            api_key = self.api_key_var.get().strip()
            if not api_key:
                messagebox.showwarning("提示", "请先输入 API Key", parent=self.dialog)
                return
            config = self._build_current_vlm_config(system_prompt="你是一个助手，直接回答用户问题即可。")

            # 禁用按钮并显示等待状态（UI 层立即响应）
            self.dialog.config(cursor="watch")
            original_states = {}
            for child in self.dialog.winfo_children():
                try:
                    original_states[child] = child.cget("state")
                    child.config(state="disabled")
                except Exception:
                    pass
            self.dialog.update()

            def _do_test():
                screenshot = None
                try:
                    screenshot, _offset = capture_physical_bbox()
                    buf = io.BytesIO()
                    screenshot.save(buf, format='JPEG', quality=85)
                    image_b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
                    vlm_engine.call_vlm_api("描述你看到了什么？", image_b64=image_b64, config=config, raise_on_error=True)
                    self._safe_dialog_after(lambda: messagebox.showinfo("成功", "API 连接成功！", parent=self.dialog))
                except Exception as e:
                    err = str(e)
                    self._safe_dialog_after(lambda msg=err: messagebox.showerror("错误", f"连接失败:\n{msg}", parent=self.dialog))
                finally:
                    if screenshot:
                        try: screenshot.close()
                        except Exception: pass
                    # 恢复 UI 状态
                    def _restore():
                        try:
                            self.dialog.config(cursor="")
                            for child in self.dialog.winfo_children():
                                try: child.config(state=original_states.get(child, "normal"))
                                except Exception: pass
                        except Exception:
                            pass
                    self._safe_dialog_after(_restore)

            threading.Thread(target=_do_test, daemon=True).start()

        except Exception as e:
            messagebox.showerror("错误", f"启动测试失败: {e}", parent=self.dialog)
            self.dialog.config(cursor="")

    def _safe_dialog_after(self, callback):
        """Queue a UI callback without touching Tk from a worker thread."""
        self._ui_callbacks.put(callback)

    def _drain_ui_callbacks(self):
        """Run queued callbacks on the Tk main thread."""
        while True:
            try:
                callback = self._ui_callbacks.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                logger.exception("UI callback failed")

        try:
            if self.dialog.winfo_exists():
                self.dialog.after(50, self._drain_ui_callbacks)
        except tk.TclError:
            pass

    def _save(self):
        try:
            import vlm_engine
            config = self._build_current_vlm_config()
            if vlm_engine.save_config(config):
                self.result = config
                self.dialog.destroy()
            else:
                messagebox.showerror("错误", "保存配置失败", parent=self.dialog)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}", parent=self.dialog)

def get_icon_path(icon_name="app_icon.ico", app_version="1.8.0"):
    """
    获取图标路径,打包后从临时目录提取
    返回可用于 iconbitmap() 的实际文件路径

    Args:
        icon_name: 图标文件名
        app_version: 应用版本号（用于临时文件命名）

    Returns:
        图标文件路径或 None
    """
    # 开发环境直接返回
    if not getattr(sys, 'frozen', False):
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_name)
        if os.path.exists(icon_path):
            return icon_path
        return None

    # 打包环境:提取到临时目录
    try:
        import tempfile
        import shutil

        # 从 _MEIPASS 获取图标
        source_icon = os.path.join(sys._MEIPASS, icon_name)

        if not os.path.exists(source_icon):
            logger.warning(f"未找到打包的图标文件: {source_icon}")
            return None

        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        temp_icon = os.path.join(temp_dir, f"macromate_{app_version}.ico")

        # 复制图标到临时目录
        shutil.copy2(source_icon, temp_icon)
        logger.info(f"图标已提取到: {temp_icon}")

        return temp_icon
    except Exception as e:
        logger.error(f"提取图标失败: {e}")
        return None

