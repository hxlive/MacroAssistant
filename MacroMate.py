# -*- coding: utf-8 -*-
# MacroMate.py
# 描述: 自动化宏的 GUI 界面
# Version: 1.8.2


# 使用: 
#   - GUI 模式: python MacroMate.py
#   - 命令行: python MacroMate.py script.json
#             python MacroMate.py --run script.json
#             python MacroMate.py --theme darkly (指定主题)

import os
import sys

# 允许在最早期通过命令行覆写日志编码（必须在 init_system_runtime 前）
for i, arg in enumerate(sys.argv):
    if arg.startswith('--log-encoding='):
        _stdio_encoding = arg.split('=', 1)[1].strip()
        os.environ['MACROMATE_STDIO_ENCODING'] = _stdio_encoding
        os.environ['MACROASSISTANT_STDIO_ENCODING'] = _stdio_encoding
    elif arg == '--log-encoding' and i + 1 < len(sys.argv):
        _stdio_encoding = sys.argv[i + 1].strip()
        os.environ['MACROMATE_STDIO_ENCODING'] = _stdio_encoding
        os.environ['MACROASSISTANT_STDIO_ENCODING'] = _stdio_encoding

import sys_utils  # [新增] 系统底层工具与初始化
sys_utils.init_system_runtime() # [新增] 初始化 DPI 感知与流重定向

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import pyautogui
import threading
import copy
import ttkbootstrap as tb
import queue
import ctypes
ctypes.pythonapi.PyThreadState_SetAsyncExc.argtypes = [ctypes.c_ulong, ctypes.py_object]
ctypes.pythonapi.PyThreadState_SetAsyncExc.restype = ctypes.c_int


# =================================================================
# 全局配置
# =================================================================
APP_VERSION = "1.8.2"
APP_TITLE = f"智点助手 (MacroMate) v{APP_VERSION}"
APP_ICON = "app_icon.ico" 
def _get_program_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _get_program_dir()
CONFIG_FILE = os.path.join(APP_DIR, "macro_settings.json")
MAX_RECENT_FILES = 5

DEFAULT_HOTKEY_RUN = "ctrl+f10"
DEFAULT_HOTKEY_STOP = "ctrl+f11"
# =================================================================
# 性能优化常量
STATUS_QUEUE_CHECK_INTERVAL_IDLE = 500  # 空闲时状态队列检查间隔（毫秒）
STATUS_QUEUE_CHECK_INTERVAL_RUNNING = 50  # 运行时状态队列检查间隔（毫秒）
STATUS_QUEUE_MAX_BATCH = 50  # 状态队列单次最大处理数
OCR_PRELOAD_DELAY = 100  # OCR引擎预热延迟（毫秒）


# resource_path 和 get_icon_path 已迁移到 gui_utils.py

import logging
logging.getLogger('rapidocr').setLevel(logging.WARNING)

# [重构] 导入核心模块与工具类
try:
    import core_engine as macro_engine
    import ocr_engine
    import vlm_engine
    import gui_utils
    
    from sys_utils import (
        GlobalHotkeyManager, MouseTracker, RegionSelector, 
        HotkeySettingsDialog, VLMSettingsDialog, ImageTooltipManager, MiniStatusWindow,
        AboutDialog
    )
    from gui_utils import (
        ParamWidgetFactory, parse_region_string, get_icon_path,
        update_loop_params, update_run_params, param_internal_to_display
    )
    from core_engine import HotkeyUtils, MacroSchema, validate_macro_data, MacroPersistence
except ImportError as e:
    messagebox.showerror("导入错误", f"缺少必要的模块文件或导入失败: {e}\n请确保所有 py 文件都在同一目录。")
    exit()

def capitalize_hotkey_str(s): return HotkeyUtils.format_hotkey_display(s)


def _region_box_to_screenshot_region(region_box):
    """Convert an optional bbox into the screenshot region used by test actions."""
    if region_box is None:
        return None

    region = macro_engine.bbox_to_region(region_box)
    if region is None:
        raise ValueError(f"手动查找区域解析失败(格式错误): {region_box}")
    return region


def _parse_optional_test_region(region_value):
    raw_region = (region_value or "").strip()
    if not raw_region:
        return None

    region_box = parse_region_string(raw_region)
    if region_box is None:
        raise ValueError("手动查找区域格式无效，应为 x1, y1, x2, y2")

    _region_box_to_screenshot_region(region_box)
    return region_box




def _preview_goto_label(p):
    label = p.get('label', '')
    max_jumps = p.get('max_jumps', 100)
    return f"-> {label}  [最多 {max_jumps} 次]"


def _preview_json_extract(p):
    suffix = "；失败用默认值" if p.get('use_default') or str(p.get('default_value', '')) != '' else ""
    return f"JSON 路径 '{p.get('json_path', '') or '$'}' -> 变量 {p.get('var_name', '')}{suffix}"


def _preview_foreach_line(p):
    source = p.get('file_path') or p.get('source_text', '')
    line_var = p.get('current_line_var', 'current_line')
    fields = p.get('field_names', '')
    suffix = f"；拆分为 {fields}" if fields else ""
    return f"批量处理文本行 '{source}' -> {{{line_var}}}{suffix}"


_STEP_PARAM_PREVIEW_FORMATTERS = {
    'GOTO_LABEL': _preview_goto_label,
    'SET_VAR': lambda p: f"变量 {p.get('var_name', '')} = '{p.get('var_value', '')}'",
    'READ_FILE': lambda p: f"读取文本 '{p.get('file_path', '')}' -> 变量 {p.get('var_name', '')}",
    'EXTRACT_VAR': lambda p: f"'{p.get('source_text', '')}' 提取 '{p.get('regex', '')}' -> 变量 {p.get('var_name', '')}",
    'JSON_EXTRACT': _preview_json_extract,
    'PROMPT_INPUT': lambda p: f"人工输入 '{p.get('prompt', '')}' -> 变量 {p.get('var_name', '')}",
    'FOREACH_LINE': _preview_foreach_line,
    'END_FOREACH': lambda _p: '结束批量处理',
    'IF_VAR': lambda p: f"如果 '{p.get('var_value', '')}' {p.get('operator', '==')} '{p.get('expected_val', '')}'",
    'CALCULATE': lambda p: f"变量计算 '{p.get('expression', '')}' -> 变量 {p.get('var_name', '')}",
    'WRITE_FILE': lambda p: f"写入文本至 '{p.get('file_path', '')}'",
    'GOTO_IF': lambda p: f"如果 '{p.get('var_value', '')}' {p.get('operator', '==')} '{p.get('expected_val', '')}' -> 跳转至 {p.get('label', '')}",
}
_LIST_DEDENT_ACTIONS = {'ELSE', 'END_IF', 'END_LOOP', 'END_FOREACH'}
_LIST_BLOCK_START_ACTIONS = {'LOOP_START', 'FOREACH_LINE'}
_LIST_BLOCK_END_ACTIONS = {'END_IF', 'END_LOOP', 'END_FOREACH'}


def _is_list_block_start(action):
    return action.startswith('IF_') or action in _LIST_BLOCK_START_ACTIONS


class MacroApp:
    def __init__(self, root):
        self.root = root
        self._setup_window()
        self._setup_app_state()
        self._setup_hotkeys()
        self._setup_app_options()
        self._setup_mouse_tracker()
        self._setup_ocr_state()
        self._setup_widget_factory()
        self._show_loading_then_defer_ui()

    def _setup_window(self):
        self.root.title(APP_TITLE)
        self.root.geometry("1160x820")  # 稍微加宽以适应优化后的列宽

        self.font_ui = ("Microsoft YaHei UI", 10)
        self.font_code = ("Consolas", 10)

        self.root.style.configure(".", font=self.font_ui)
        self.root.style.configure("Treeview", font=self.font_code, rowheight=25)
        self.root.style.configure("Treeview.Heading", font=self.font_ui)

        self.is_app_running = True
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

        icon_path = get_icon_path(APP_ICON, APP_VERSION)
        if icon_path and os.path.exists(icon_path):
            try:
                # Set AppUserModelID before iconbitmap for taskbar icon stability.
                sys_utils.set_windows_app_id(APP_VERSION)
                print(f"[Info] AppUserModelID set: {APP_VERSION}")
                self.root.iconbitmap(icon_path)
                print(f"[Info] icon set: {icon_path}")
            except Exception as e:
                print(f"[错误] 设置图标失败: {e}")
        else:
            print("[警告] 未找到图标文件，使用默认图标")

    def _setup_app_state(self):
        self.steps = []
        self.editing_index = None
        self.is_macro_running = False
        self.last_test_location = None
        self.current_run_context = None
        self.current_filepath = None
        self._macro_thread = None
        self._stop_in_progress = False
        self._run_pending = False
        self._pending_run_id = None
        self.mini_status_window = None
        self.recent_files = []
        self.status_queue = queue.Queue()
        self._last_mini_status = (None, None)

    def _setup_hotkeys(self):
        self.hotkey_run_str = tb.StringVar(value=DEFAULT_HOTKEY_RUN)
        self.hotkey_stop_str = tb.StringVar(value=DEFAULT_HOTKEY_STOP)
        self.hotkey_manager = GlobalHotkeyManager(
            self.root,
            get_run_str_cb=self.hotkey_run_str.get,
            get_stop_str_cb=self.hotkey_stop_str.get,
            trigger_run_cb=self.safe_run_macro,
            trigger_stop_cb=self.safe_stop_macro
        )

    def _setup_app_options(self):
        self.current_theme = tb.StringVar(value=self.root.style.theme_use())
        self.skip_confirm_var = tb.BooleanVar(value=False)
        self.dont_minimize_var = tb.BooleanVar(value=False)
        self.enhanced_mode_var = tb.BooleanVar(value=False)
        self.run_enabled_var = tb.BooleanVar(value=False)

    def _setup_mouse_tracker(self):
        self.mouse_pos_var = tb.StringVar()
        self.mouse_tracker = MouseTracker(self.root, self.mouse_pos_var)

    def _setup_ocr_state(self):
        self.FULL_OCR_NAME_MAP = {
            'auto': '自动选择 (Auto)',
            'winocr': 'Windows 10/11 OCR',
            'rapidocr': 'RapidOCR',
            'tesseract': 'Tesseract OCR',
            'none': '无可用OCR引擎'
        }
        self.FULL_OCR_KEY_MAP = {name: key for key, name in self.FULL_OCR_NAME_MAP.items()}
        # OCR 引擎检测将在后台线程运行，先用占位值保证主线程快速进入 mainloop
        self.available_ocr_engines = []
        self.available_ocr_keys = ['auto']

    def _setup_widget_factory(self):
        self.widget_factory = ParamWidgetFactory(
            font_ui=self.font_ui,
            font_code=self.font_code,
            ocr_name_map=self.FULL_OCR_NAME_MAP
        )

    def _show_loading_then_defer_ui(self):
        # 窗口出现后再构建重型 UI，避免启动期看起来未响应。
        self.root.update_idletasks()
        self._splash_label = tk.Label(
            self.root,
            text="正在加载界面...",
            font=("Microsoft YaHei UI", 14),
            fg="#666666",
            bg="#FFFFFF"
        )
        self._splash_label.place(relx=0.5, rely=0.5, anchor="center")
        self.root.update_idletasks()
        self.root.after(10, self._deferred_ui_init)
    def _deferred_ui_init(self):
        """mainloop 已启动后才执行 UI 构建，避免窗口冻结。"""
        if hasattr(self, '_splash_label') and self._splash_label:
            self._splash_label.destroy()
            self._splash_label = None

        if not self._build_deferred_ui():
            return

        self._start_runtime_services()

    def _build_deferred_ui(self):
        try:
            self._init_menu()
            self._init_ui()
        except Exception as e:
            self.root.deiconify()
            self.root.update()
            messagebox.showerror("初始化失败", f"UI 构建出错:\n{str(e)}")
            self.root.quit()
            return False
        return True

    def _start_runtime_services(self):
        self.tooltip_manager = ImageTooltipManager(self.steps_tree, lambda: self.steps)
        self.load_app_settings()
        self.update_recent_files_menu()
        self.update_status_bar_hotkeys()
        self.root.after(500, self.hotkey_manager.check_conflicts)
        self.hotkey_manager.start_listener()
        threading.Thread(target=self._detect_ocr_engines_bg, daemon=True).start()
        self._check_status_queue()
    # ------------------------------------------------------------------
    # OCR 引擎异步检测（避免阻塞主线程）
    # ------------------------------------------------------------------
    def _detect_ocr_engines_bg(self):
        """后台线程：检测可用 OCR 引擎并预热，完成后回调主线程更新状态。"""
        try:
            engines = ocr_engine.get_available_engines()
            # 检测完成后顺手预热（合并两次后台任务）
            ocr_engine.preload_engines()
        except Exception as e:
            print(f"[OCR] 后台检测异常: {e}")
            engines = [('none', '无可用OCR引擎')]
        self.root.after(0, self._on_ocr_engines_ready, engines)

    def _on_ocr_engines_ready(self, engines):
        """主线程回调：OCR 引擎检测完成，更新状态。"""
        # [修复H-7] 应用可能在检测期间已关闭，需先检查生命周期
        if not self.is_app_running:
            return
        self.available_ocr_engines = engines
        self.available_ocr_keys = [e[0] for e in engines]
        if 'none' in self.available_ocr_keys:
            print("[警告] 未找到任何可用的OCR引擎 (RapidOCR, Tesseract, WinOCR)。")
            try:
                self.status_var.set("WARN 未找到可用 OCR 引擎，文本查找功能不可用。")
            except Exception:
                pass
        else:
            engine_names = ' / '.join(e[1] for e in engines)
            print(f"[OCR] 引擎就绪: {engine_names}")

    def _init_menu(self):
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        self._build_file_menu()
        self._build_settings_menu()
        self._build_theme_menu()
        self._build_about_menu()

    def _build_file_menu(self):
        file_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  文件  ", menu=file_menu)
        file_menu.add_command(label="📄 新建宏", accelerator="Ctrl+N", command=self.new_macro)
        file_menu.add_command(label="📂 打开宏...", accelerator="Ctrl+O", command=self.load_macro)
        file_menu.add_command(label="💾 保存宏...", accelerator="Ctrl+S", command=self.save_macro)
        file_menu.add_separator()
        self.recent_files_menu = tk.Menu(file_menu, tearoff=0, font=self.font_ui)
        file_menu.add_cascade(label="最近加载", menu=self.recent_files_menu)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_exit)

        self.root.bind('<Control-n>', lambda e: self.new_macro())
        self.root.bind('<Control-o>', lambda e: self.load_macro())
        self.root.bind('<Control-s>', lambda e: self.save_macro())

    def _build_settings_menu(self):
        settings_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  设置  ", menu=settings_menu)
        settings_menu.add_command(label="⌨ 快捷键设置...", command=self.open_hotkey_settings)
        settings_menu.add_separator()
        settings_menu.add_command(label="🤖 AI 设置...", command=self.open_vlm_settings)

    def _build_theme_menu(self):
        theme_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  主题  ", menu=theme_menu)
        
        light_themes = ['litera', 'cosmo', 'flatly', 'journal', 'lumen', 'minty', 'pulse', 'sandstone', 'united', 'yeti']
        for theme in light_themes:
            theme_menu.add_radiobutton(label=f"亮 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)
        theme_menu.add_separator()
        dark_themes = ['superhero', 'cyborg', 'darkly', 'solar']
        for theme in dark_themes:
            theme_menu.add_radiobutton(label=f"暗 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)

    def _build_about_menu(self):
        about_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  关于  ", menu=about_menu)
        about_menu.add_command(label="关于", command=self.show_about_dialog)

    def _init_ui(self):
        self._build_status_bar()
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self._build_step_list_panel(main_frame)
        self._build_step_editor_panel(main_frame)
        self.update_param_fields(None)

    def _build_status_bar(self):
        status_bar_frame = ttk.Frame(self.root, bootstyle="primary")
        status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var = tk.StringVar()
        self.status_label_left = ttk.Label(status_bar_frame, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W, padding=5, bootstyle="primary-inverse", font=self.font_ui)
        self.status_label_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.loop_status_var = tk.StringVar()
        self.loop_status_label_right = ttk.Label(status_bar_frame, textvariable=self.loop_status_var, relief=tk.FLAT, anchor=tk.E, padding=(0, 5, 5, 5), bootstyle="primary-inverse", font=self.font_ui)
        self.loop_status_label_right.pack(side=tk.RIGHT)

    def _build_step_list_panel(self, main_frame):
        list_frame = ttk.Frame(main_frame, padding=10)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_steps_tree(list_frame)
        self._build_step_list_controls(list_frame)

    def _build_steps_tree(self, list_frame):
        title_frame = ttk.Frame(list_frame)
        title_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(title_frame, text="宏步骤序列:", font=("Microsoft YaHei UI", 11, "bold")).pack(side=tk.LEFT)

        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("id", "action", "params")
        self.steps_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

        self.steps_tree.heading("id", text="#")
        self.steps_tree.heading("action", text="动作")
        self.steps_tree.heading("params", text="参数详情 / 备注")

        self.steps_tree.column("id", width=45, minwidth=40, stretch=False, anchor="center")
        self.steps_tree.column("action", width=220, minwidth=200, stretch=False)
        self.steps_tree.column("params", width=320, minwidth=280, stretch=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.steps_tree.yview)
        self.steps_tree.configure(yscrollcommand=scrollbar.set)

        self.steps_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.steps_tree.bind("<Double-1>", lambda e: self.load_step_for_edit())

        self.tree_menu = tk.Menu(self.root, tearoff=0, font=self.font_ui)
        self.tree_menu.add_command(label="屏蔽/启用选中步骤", command=self.toggle_step_enabled)
        self.steps_tree.bind("<Button-3>", self.show_tree_menu)

        self.steps_tree.tag_configure('editing', background='#FFF3CD')
        self.steps_tree.tag_configure('disabled', foreground='#999999')

    def _build_step_list_controls(self, list_frame):
        left_bottom_frame = ttk.Frame(list_frame)
        left_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        left_bottom_frame.columnconfigure(0, weight=1); left_bottom_frame.columnconfigure(1, weight=1)
        left_bottom_frame.columnconfigure(2, weight=1); left_bottom_frame.columnconfigure(3, weight=1)

        self.move_up_btn = ttk.Button(left_bottom_frame, text="↑ 上移", command=lambda: self.move_step("up"), bootstyle="primary-outline", padding=(10, 6))
        self.move_up_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2), pady=(0, 5))
        self.move_down_btn = ttk.Button(left_bottom_frame, text="↓ 下移", command=lambda: self.move_step("down"), bootstyle="primary-outline", padding=(10, 6))
        self.move_down_btn.grid(row=0, column=1, sticky="nsew", padx=2, pady=(0, 5))
        self.remove_btn = ttk.Button(left_bottom_frame, text="🗑 删除选中", command=self.remove_step, bootstyle="danger-outline", padding=(10, 6))
        self.remove_btn.grid(row=0, column=2, sticky="nsew", padx=2, pady=(0, 5))
        self.load_step_btn = ttk.Button(left_bottom_frame, text="✎ 修改步骤", command=self.load_step_for_edit, bootstyle="info-outline", padding=(10, 6))
        self.load_step_btn.grid(row=0, column=3, sticky="nsew", padx=(2, 0), pady=(0, 5))

        self.run_btn = ttk.Button(left_bottom_frame, text="", command=self.run_macro, bootstyle="success", padding=(15, 10))
        self.run_btn.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=(0, 0), pady=5)

        check_frame = ttk.Frame(left_bottom_frame)
        check_frame.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=(10, 0))
        check_frame.columnconfigure(0, weight=1); check_frame.columnconfigure(1, weight=1)

        skip_check = ttk.Checkbutton(check_frame, text="跳过运行前的确认提示", variable=self.skip_confirm_var, bootstyle="primary-round-toggle")
        skip_check.grid(row=0, column=0, sticky="w", padx=2, pady=(0, 5))
        minimize_check = ttk.Checkbutton(check_frame, text="运行时主界面不最小化", variable=self.dont_minimize_var, bootstyle="primary-round-toggle")
        minimize_check.grid(row=0, column=1, sticky="w", padx=2, pady=(0, 5))

        enhanced_check = ttk.Checkbutton(check_frame, text="开启增强模式 (识别不到小字时可开启)", variable=self.enhanced_mode_var, bootstyle="success-round-toggle")
        enhanced_check.grid(row=1, column=0, sticky="w", padx=2, pady=(0, 5))

        run_enabled_check = ttk.Checkbutton(check_frame, text="启用 RUN 步骤 (注意安全风险)", variable=self.run_enabled_var, bootstyle="danger-round-toggle")
        run_enabled_check.grid(row=1, column=1, sticky="w", padx=2, pady=(0, 5))

    def _build_step_editor_panel(self, main_frame):
        add_frame = ttk.Labelframe(main_frame, text="添加新步骤", padding=10)
        add_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)

        add_frame.pack_propagate(False)
        add_frame.configure(width=380)

        right_bottom_frame = ttk.Frame(add_frame)
        right_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        right_bottom_frame.columnconfigure(0, weight=2); right_bottom_frame.columnconfigure(1, weight=1)

        self.add_step_btn = ttk.Button(right_bottom_frame, text="＋ 添加到序列 >>", command=self.add_or_update_step, bootstyle="success", padding=(12, 8))
        self.add_step_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2), columnspan=2)
        self.cancel_edit_btn = ttk.Button(right_bottom_frame, text="✕ 取消修改", command=self.cancel_edit_mode, bootstyle="secondary", padding=(10, 6))

        ttk.Label(add_frame, text="选择动作:").pack(anchor="w")
        self.action_type = ttk.Combobox(add_frame, state="readonly", font=self.font_ui, height=16)
        self.action_type['values'] = list(MacroSchema.ACTION_TRANSLATIONS.values())
        self.action_type.current(0)

        self.action_type.pack(anchor="w", fill=tk.X, pady=5)
        self.action_type.bind("<<ComboboxSelected>>", self._on_action_type_selected)
        self.action_type.bind("<FocusIn>", self._clear_combobox_text_selection, add="+")
        self.action_type.bind("<ButtonRelease-1>", self._clear_combobox_text_selection, add="+")

        param_area = ttk.Frame(add_frame)
        param_area.pack(fill=tk.BOTH, expand=True, pady=5)
        self._build_scrollable_param_area(param_area)

        self.param_widgets = {}

    def _clear_combobox_text_selection(self, event=None):
        widget = getattr(event, 'widget', None) if event is not None else getattr(self, 'action_type', None)
        if widget is None:
            return

        def clear_selection():
            try:
                widget.selection_clear()
                widget.icursor(tk.END)
            except tk.TclError:
                pass

        try:
            widget.after_idle(clear_selection)
        except tk.TclError:
            pass

    def _on_action_type_selected(self, event):
        self.update_param_fields(event)
        self._clear_combobox_text_selection(event)

    def _build_scrollable_param_area(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        self.param_canvas = tk.Canvas(parent, highlightthickness=0, borderwidth=0)
        self.param_scrollbar = tk.Canvas(self.param_canvas, width=8, highlightthickness=0, borderwidth=0)
        self.param_frame = ttk.Frame(self.param_canvas)
        self.param_window = self.param_canvas.create_window((0, 0), window=self.param_frame, anchor="nw")
        self._param_scrollbar_visible = False
        self._param_scrollbar_update_pending = False
        self._param_scroll_drag_offset = 0
        self._param_scroll_thumb = (0, 0)

        self.param_canvas.configure(yscrollcommand=self._on_param_yview_changed)
        self.param_canvas.grid(row=0, column=0, sticky="nsew")
        self._style_floating_param_scrollbar()

        self.param_frame.bind("<Configure>", self._on_param_frame_configure)
        self.param_canvas.bind("<Configure>", self._on_param_canvas_configure)
        self.param_scrollbar.bind("<ButtonPress-1>", self._on_param_scrollbar_press)
        self.param_scrollbar.bind("<B1-Motion>", self._on_param_scrollbar_drag)
        self._bind_param_mousewheel()
        self._update_param_scrollbar_visibility()

    def _style_floating_param_scrollbar(self):
        try:
            bg = self.param_canvas.cget("background")
        except Exception:
            bg = "#f5f5f5"
        self.param_scrollbar.configure(background=bg)
        self._param_scroll_thumb_color = "#9aa0a6"
        self._param_scroll_thumb_active_color = "#6f767d"

    def _on_param_frame_configure(self, event=None):
        if hasattr(self, 'param_canvas'):
            self.param_canvas.configure(scrollregion=self.param_canvas.bbox("all"))
            self._schedule_param_scrollbar_update()

    def _on_param_canvas_configure(self, event):
        if hasattr(self, 'param_window'):
            self.param_canvas.itemconfigure(self.param_window, width=event.width)
            self._schedule_param_scrollbar_update()

    def _on_param_yview_changed(self, first, last):
        self._schedule_param_scrollbar_update()

    def _schedule_param_scrollbar_update(self):
        if getattr(self, '_param_scrollbar_update_pending', False):
            return
        self._param_scrollbar_update_pending = True
        self.root.after_idle(self._update_param_scrollbar_visibility)

    def _update_param_scrollbar_visibility(self):
        if not hasattr(self, 'param_canvas'):
            return
        self._param_scrollbar_update_pending = False
        bbox = self.param_canvas.bbox("all")
        content_height = (bbox[3] - bbox[1]) if bbox else 0
        canvas_height = max(self.param_canvas.winfo_height(), 1)
        should_show = content_height > canvas_height + 1
        self._set_param_scrollbar_visible(should_show)
        if not should_show:
            self.param_canvas.yview_moveto(0)
        self._draw_param_scrollbar_thumb()

    def _set_param_scrollbar_visible(self, visible):
        if getattr(self, '_param_scrollbar_visible', False) == visible:
            return
        self._param_scrollbar_visible = visible
        if visible:
            self.param_scrollbar.place(in_=self.param_canvas, relx=1.0, rely=0.0, relheight=1.0, anchor="ne", width=8)
        else:
            self.param_scrollbar.place_forget()

    def _draw_param_scrollbar_thumb(self, active=False):
        if not getattr(self, '_param_scrollbar_visible', False):
            self.param_scrollbar.delete("all")
            return
        self.param_scrollbar.update_idletasks()
        height = max(self.param_scrollbar.winfo_height(), self.param_canvas.winfo_height(), 1)
        first, last = self.param_canvas.yview()
        visible_fraction = max(last - first, 0.05)
        thumb_height = max(int(height * visible_fraction), 24)
        thumb_height = min(thumb_height, height)
        track_height = max(height - thumb_height, 1)
        y1 = int(track_height * first / max(1.0 - visible_fraction, 0.0001)) if visible_fraction < 1 else 0
        y1 = max(0, min(y1, height - thumb_height))
        y2 = y1 + thumb_height
        self._param_scroll_thumb = (y1, y2)
        color = self._param_scroll_thumb_active_color if active else self._param_scroll_thumb_color
        self.param_scrollbar.delete("all")
        self.param_scrollbar.create_rectangle(2, y1, 6, y2, fill=color, outline="")

    def _on_param_scrollbar_press(self, event):
        if not getattr(self, '_param_scrollbar_visible', False):
            return "break"
        y1, y2 = self._param_scroll_thumb
        if y1 <= event.y <= y2:
            self._param_scroll_drag_offset = event.y - y1
        else:
            self._param_scroll_drag_offset = max((y2 - y1) // 2, 0)
            self._move_param_scrollbar_thumb_to(event.y)
        self._draw_param_scrollbar_thumb(active=True)
        return "break"

    def _on_param_scrollbar_drag(self, event):
        if not getattr(self, '_param_scrollbar_visible', False):
            return "break"
        self._move_param_scrollbar_thumb_to(event.y)
        self._draw_param_scrollbar_thumb(active=True)
        return "break"

    def _move_param_scrollbar_thumb_to(self, y):
        height = max(self.param_scrollbar.winfo_height(), 1)
        y1, y2 = self._param_scroll_thumb
        thumb_height = max(y2 - y1, 1)
        track_height = max(height - thumb_height, 1)
        target = max(0, min(y - self._param_scroll_drag_offset, track_height))
        self.param_canvas.yview_moveto(target / track_height)

    def _bind_param_mousewheel(self):
        self._bind_param_mousewheel_target(self.param_canvas)
        self._bind_param_mousewheel_target(self.param_frame)

    def _bind_param_mousewheel_target(self, widget):
        if getattr(widget, '_macromate_param_scroll_bound', False):
            return
        widget.bind("<MouseWheel>", self._on_param_mousewheel, add="+")
        widget.bind("<Button-4>", self._on_param_mousewheel, add="+")
        widget.bind("<Button-5>", self._on_param_mousewheel, add="+")
        widget._macromate_param_scroll_bound = True

    def _bind_param_mousewheel_to_children(self, widget=None):
        widget = widget or self.param_frame
        self._bind_param_mousewheel_target(widget)
        for child in widget.winfo_children():
            self._bind_param_mousewheel_to_children(child)

    def _on_param_mousewheel(self, event):
        if not hasattr(self, 'param_canvas'):
            return None
        if getattr(event, 'num', None) == 4:
            direction = -1
        elif getattr(event, 'num', None) == 5:
            direction = 1
        else:
            direction = -1 if event.delta > 0 else 1
        self.param_canvas.yview_scroll(direction, "units")
        return "break"

    def _refresh_param_scroll_region(self):
        if not hasattr(self, 'param_canvas'):
            return
        self.root.update_idletasks()
        self._bind_param_mousewheel_to_children()
        self.param_canvas.configure(scrollregion=self.param_canvas.bbox("all"))
        self.param_canvas.yview_moveto(0)
        self._update_param_scrollbar_visibility()

    # --- Treeview 辅助方法 ---
    def _get_selected_index(self):
        """获取当前选中项的索引"""
        selected_items = self.steps_tree.selection()
        if not selected_items: return None
        return self.steps_tree.index(selected_items[0])

    def _get_hotkey_display_pair(self):
        run_display = capitalize_hotkey_str(self.hotkey_run_str.get())
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        return run_display, stop_display

    def _format_status_bar_text(self, run_display, stop_display):
        return f"准备就绪...  |  [{run_display}] 启动宏  |  [{stop_display}] 停止宏"

    def _format_run_button_text(self, run_display):
        return f"▶ 运行宏 ({run_display})"

    def update_status_bar_hotkeys(self):
        """更新状态栏和运行按钮上的快捷键提示"""
        run_display, stop_display = self._get_hotkey_display_pair()
        self.status_var.set(self._format_status_bar_text(run_display, stop_display))
        self.run_btn.config(text=self._format_run_button_text(run_display))

    def _format_window_title(self):
        if not self.current_filepath:
            return APP_TITLE

        filename = os.path.basename(self.current_filepath)
        return f"{APP_TITLE}  ---  {filename}"

    def update_title(self):
        """更新窗口标题栏，额外加上当前宏文件的文件名"""
        self.root.title(self._format_window_title())

    def open_hotkey_settings(self):
        """打开快捷键设置对话框"""
        dialog = HotkeySettingsDialog(
            self.root, 
            self.hotkey_run_str.get(), 
            self.hotkey_stop_str.get(),
            default_run=DEFAULT_HOTKEY_RUN,
            default_stop=DEFAULT_HOTKEY_STOP
        )
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            new_run, new_stop = dialog.result
            self.hotkey_run_str.set(new_run)
            self.hotkey_stop_str.set(new_stop)
            
            self.on_save_hotkeys()
            
            messagebox.showinfo(
                "设置已保存",
                f"快捷键已更新:\n\n"
                f"运行宏: {capitalize_hotkey_str(new_run)}\n"
                f"停止宏: {capitalize_hotkey_str(new_stop)}",
                parent=self.root
            )
            
    def on_save_hotkeys(self):
        """保存并重启监听器"""
        self.save_app_settings()
        
        if not self.hotkey_manager.check_conflicts(show_success=False):
            messagebox.showwarning("冲突警告", "快捷键已保存，但检测到冲突。\n请确保没有其他程序占用它。", parent=self.root)
        
        self.hotkey_manager.restart_listener()
        self.update_status_bar_hotkeys()

    def open_vlm_settings(self):
        """打开 VLM (AI) 设置对话框"""
        dialog = VLMSettingsDialog(self.root)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            messagebox.showinfo("设置已保存", "AI 配置已更新", parent=self.root)

    def show_about_dialog(self):
        """显示关于对话框"""
        if hasattr(self, '_about_dialog_ref') and self._about_dialog_ref and self._about_dialog_ref.dialog.winfo_exists():
            self._about_dialog_ref.dialog.focus_force()
            return
            
        icon_path = get_icon_path(APP_ICON, APP_VERSION)
        self._about_dialog_ref = AboutDialog(self.root, APP_VERSION, icon_path)


    def on_exit(self):
        self.is_app_running = False
        self.safe_stop_macro()
        if self.current_run_context:
            macro_engine.cleanup_active_processes(self.current_run_context)
        
        # [变更] 使用 MouseTracker 类停止
        self.mouse_tracker.stop()
            
        if self.hotkey_manager and self.hotkey_manager.listener:
            print("[Info] 正在停止快捷键监听器...")
            try:
                self.hotkey_manager.listener.stop()
                self.hotkey_manager.listener.join(timeout=0.5) 
            except Exception as e:
                print(f"[警告] 停止监听器时出错: {e}")
                
        try:
            self.root.quit()
            self.root.destroy()
        except Exception: 
            pass

    def update_param_fields(self, event):
        self.last_test_location = None
        
        # [变更] 停止鼠标追踪
        self.mouse_tracker.stop()
        self.mouse_pos_var.set("")
        
        for widget in self.param_frame.winfo_children(): widget.destroy()
        self.param_widgets = {}
        action_key = MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action_key: return
        
        # [重构] 准备回调函数字典
        callbacks = {
            'on_select_region': self.on_select_region,
            'browse_image': self.browse_image,
            'on_test_find_image_click': self.on_test_find_image_click,
            'on_test_find_text_click': self.on_test_find_text_click,
            'on_test_ai_command_click': self.on_test_ai_command_click,
            'update_loop_params': lambda event: update_loop_params(self.param_widgets, self.param_frame, self.param_widgets.get('mode')) if self.param_widgets.get('mode') is not None else None,
            'update_run_params': lambda event: update_run_params(self.param_widgets, self.param_frame, self.param_widgets.get('run_type')) if self.param_widgets.get('run_type') is not None else None,
            'mouse_tracker': self.mouse_tracker,
            'mouse_pos_var': self.mouse_pos_var
        }

        # [重构] 使用工厂模式构建参数表单
        res = self.widget_factory.build_action_form(
            action_key, 
            self.param_frame, 
            self.param_widgets, 
            self.available_ocr_keys, 
            callbacks
        )

        # 处理特殊返回 (如 OCR 不可用时自动切回图像模式)
        if res == "SWITCH_TO_FIND_IMAGE":
            self.action_type.set(MacroSchema.ACTION_TRANSLATIONS['FIND_IMAGE'])
            self.update_param_fields(None)
            return

        self._refresh_param_scroll_region()


    # update_loop_params 和 update_run_params 已迁移到 gui_utils.py

    def on_select_region(self, entry_widget):
        self.root.iconify()
        self.root.after(300, lambda: self._do_select_region(entry_widget))
        
    def _do_select_region(self, entry_widget):
        
        try:
            # [变更] 使用 gui_utils 中的 RegionSelector
            region = RegionSelector(self.root).get_region()
            self.root.deiconify()
            
            if region:
                val_str = f"{region[0]}, {region[1]}, {region[2]}, {region[3]}"
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, val_str)
        except Exception as e:
            self.root.deiconify()
            messagebox.showerror("错误", f"选区失败: {e}")

    # create_browse_button, create_test_button, _create_hint_label 已迁移到 gui_utils.py

    def on_test_find_image_click(self):
        try:
            path = self.param_widgets['path'].get()
            conf = float(self.param_widgets['confidence'].get())
            if not os.path.exists(path): raise FileNotFoundError
            
            # <--- 读取搜索范围
            region_box = None
            if 'region' in self.param_widgets:
                region_box = _parse_optional_test_region(self.param_widgets['region'].get())

            self.status_var.set("测试中...")
            self.root.iconify()
            # 将 region_box 传给线程
            self._run_test_after_iconify(self._test_find_image, (path, conf, region_box))
        except Exception as e:
            messagebox.showerror('错误', f"参数无效: {e}")

    def on_test_find_text_click(self):
        try:
            text = self.param_widgets['text'].get()
            lang = MacroSchema.LANG_OPTIONS.get(self.param_widgets['lang'].get(), 'eng')
            
            # 获取下拉框的原始值
            engine_name = self.param_widgets['engine'].get()
            
            # <--- 优先检查是否不可用
            if engine_name.endswith(" (不可用)"):
                messagebox.showwarning(
                    "引擎不可用", 
                    f"您选择的引擎 '{engine_name}' 在当前环境中未安装或无法加载。\n\n请选择其他引擎，或安装相应组件后重启程序。",
                    parent=self.root
                )
                return # 直接阻断测试，不再往下执行
            
            engine = self.FULL_OCR_KEY_MAP.get(engine_name, 'auto')
            
            region_box = None
            if 'region' in self.param_widgets:
                region_box = _parse_optional_test_region(self.param_widgets['region'].get())
            
            if not text: raise ValueError
            self.status_var.set("测试中...")
            self.root.iconify()
            # 将 region_box 传给线程
            self._run_test_after_iconify(self._test_find_text, (text, lang, engine, region_box))
        except Exception as e:
            messagebox.showerror('错误', f"参数无效: {e}")

    def on_test_ai_command_click(self):
        """测试 AI 自然语言指令"""
        try:
            instruction = self.param_widgets['instruction'].get()
            if not instruction.strip():
                messagebox.showwarning("输入错误", "请输入 AI 指令")
                return
            
            # 读取区域
            region_box = None
            if 'region' in self.param_widgets:
                region_box = _parse_optional_test_region(self.param_widgets['region'].get())
            
            self.status_var.set("AI 分析中...")
            self.root.iconify()
            self._run_test_after_iconify(self._test_ai_command, (instruction, region_box))
        except Exception as e:
            messagebox.showerror('错误', f"参数无效: {e}")

    def _run_test_after_iconify(self, func, args, attempts=0):
        if self.root.state() == 'iconic' or attempts >= 15:
            self.root.after(250, lambda: self._run_test_thread(func, args))
            return
        self.root.after(100, lambda: self._run_test_after_iconify(func, args, attempts + 1))

    def _test_ai_command(self, instruction, region_box=None):
        """后台执行 AI 指令测试"""
        try:
            _region_box_to_screenshot_region(region_box)
            coords = vlm_engine.find_location_by_vlm(instruction, region=region_box)
            self.root.after(0, lambda: self._on_test_ai_complete(coords))
        except Exception as e:
            self.root.after(0, lambda err=e: self._on_test_error(err))

    def _on_test_ai_complete(self, coords):
        """AI 测试完成回调"""
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        if coords and len(coords) >= 2:
            self.last_test_location = (coords[0], coords[1])
            pyautogui.moveTo(coords[0], coords[1])
            messagebox.showinfo("AI 成功", f"找到坐标: {self.last_test_location}\n\nAI 已移动鼠标到该位置")
        else:
            messagebox.showwarning("AI 失败", "未能从 AI 获取有效坐标\n\n请检查:\n1. API Key 是否正确配置\n2. 网络是否正常\n3. 指令是否清晰")
        self.update_status_bar_hotkeys()
        self.root.attributes('-topmost', False)

    def _run_test_thread(self, func, args):
        threading.Thread(target=func, args=args, daemon=True).start()

    def _test_find_image(self, path, conf, region_box=None):
        screenshot = None
        try:
            screenshot_region = _region_box_to_screenshot_region(region_box)
            screenshot, offset = macro_engine.smart_screenshot(screenshot_region)
                
            res_val = macro_engine.find_image_cv2(path, conf, screenshot_pil=screenshot, offset=offset)
            loc = res_val[0] if res_val else None
            self.root.after(0, lambda: self._on_test_complete(loc))
        except Exception as e: 
            self.root.after(0, lambda err=e: self._on_test_error(err))
        finally:
            if screenshot:
                try: screenshot.close()
                except Exception: pass

    def _test_find_text(self, text, lang, engine, region_box=None):
        screenshot = None
        try:
            screenshot_region = _region_box_to_screenshot_region(region_box)
            screenshot, offset = macro_engine.smart_screenshot(screenshot_region)
            
            loc = ocr_engine.find_text_location(text, lang, True, screenshot_pil=screenshot, offset=offset, engine=engine)
            self.root.after(0, lambda: self._on_test_complete(loc))
        except Exception as e: 
            self.root.after(0, lambda err=e: self._on_test_error(err))
        finally:
            if screenshot:
                try: screenshot.close()
                except Exception: pass

    def _on_test_complete(self, loc):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        
        x, y = None, None
        
        if loc and len(loc) >= 2:
            # 兼容两种格式:
            # 1. (x, y) - 图像查找返回
            # 2. ((x, y), text) - OCR返回
            first = loc[0]
            if isinstance(first, tuple) and len(first) >= 2:
                # 格式2: ((x, y), text)
                x, y = first[0], first[1]
            else:
                # 格式1: (x, y)
                x, y = loc[0], loc[1]
        
        if x is not None and y is not None:
            self.last_test_location = (x, y)
            pyautogui.moveTo(x, y)
            messagebox.showinfo("成功", f"找到于 {self.last_test_location}")
        else:
            messagebox.showwarning("失败", "未找到目标")
        
        self.update_status_bar_hotkeys()
        self.root.attributes('-topmost', False)

    def _on_test_error(self, e):
        self.root.deiconify()
        error_msg = str(e)
        messagebox.showerror("错误", error_msg)
        self.update_status_bar_hotkeys()

    def browse_image(self):
        """浏览图片文件（保持向后兼容）"""
        f = filedialog.askopenfilename(filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if f: 
            f = os.path.abspath(f) 
            self.param_widgets['path'].delete(0, tk.END)
            self.param_widgets['path'].insert(0, f)

    def add_or_update_step(self):
        """添加或更新步骤 (已重构：使用工厂收集数据)"""
        action = MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action: return
        
        # [重构] 使用工厂模式收集并验证数据
        params, error = self.widget_factory.collect_step_data(action, self.param_widgets, self.FULL_OCR_KEY_MAP)
        if error:
            messagebox.showwarning("输入错误", error)
            return
            
        # 如果原来是 cache_box，且保存时仍收集到了 region，则自动将其转回 cache_box 以保持旧宏语义
        if getattr(self, 'editing_step_has_cache_box', False) and 'region' in params:
            params['cache_box'] = params.pop('region')
            
        step = {"action": action, "params": params}
        
        # 仅在没有手动指定区域时，才询问是否使用测试结果作为缓存
        if action in ('FIND_TEXT', 'FIND_IMAGE', 'IF_TEXT_FOUND', 'IF_IMAGE_FOUND') \
           and self.editing_index is None \
           and self.last_test_location \
           and 'region' not in step['params'] \
           and 'cache_box' not in step['params']:
            if messagebox.askyesno("缓存", "使用测试坐标作为缓存？"):
                step["params"]["cache_box"] = [self.last_test_location[0], self.last_test_location[1], self.last_test_location[0]+1, self.last_test_location[1]+1]

        # ============================================================
        # [核心修改] 插入逻辑优化
        # ============================================================
        target_index = -1 # 记录新位置用于滚动
        
        if self.editing_index is not None:
            # 修改模式：原地更新
            self.steps[self.editing_index] = step
            target_index = self.editing_index
            self.cancel_edit_mode()
        else:
            # 新增模式：检查当前是否有选中行
            selected_idx = self._get_selected_index()
            
            if selected_idx is not None:
                # 有选中：插入到选中行的下一行
                target_index = selected_idx + 1
                self.steps.insert(target_index, step)
            else:
                # 无选中：追加到末尾
                self.steps.append(step)
                target_index = len(self.steps) - 1
                
            self.update_listbox_display()
        
        # ============================================================
        # [UI优化] 自动滚动并选中新添加/修改的行
        # ============================================================
        children = self.steps_tree.get_children()
        if 0 <= target_index < len(children):
            item_id = children[target_index]
            self.steps_tree.see(item_id)           # 滚动到可见
            self.steps_tree.selection_set(item_id) # 自动选中
            
        self.last_test_location = None

    def load_step_for_edit(self):
        """加载选中步骤到编辑区 (修复：循环模式回显问题)"""
        idx = self._get_selected_index()
        if idx is None: return
        
        step = self.steps[idx]
        
        # 1. 设置动作类型 (这将重置右侧面板为默认状态)
        self.action_type.set(MacroSchema.ACTION_TRANSLATIONS.get(step['action']))
        self.update_param_fields(None)
        self._clear_combobox_text_selection()
        
        # ============================================================
        # [关键修复] 优先强制处理 LOOP_START 的模式
        # ============================================================
        if step['action'] == 'LOOP_START':
            # 获取保存的模式 (默认 fixed)
            saved_mode = step['params'].get('mode', 'fixed')
            
            # 翻译模式为中文
            display_mode = gui_utils.LOOP_MODE_DISPLAY_BY_VALUE.get(saved_mode, '固定次数')
            
            # 1. 强行修改下拉框的值
            if 'mode' in self.param_widgets:
                self.param_widgets['mode'].set(display_mode)
            
            # 2. 强行触发界面刷新 (这一步会让"目标文本"输入框从隐藏变为显示)
            if 'mode' in self.param_widgets:
                update_loop_params(self.param_widgets, self.param_frame, self.param_widgets['mode'])

        # ============================================================
        # [新增] 优先强制处理 RUN 的类型
        # ============================================================
        elif step['action'] == 'RUN':
            # 获取保存的类型 (默认 command)
            saved_run_type = step['params'].get('run_type', 'command')
            
            # 翻译类型为中文
            display_run_type = MacroSchema.RUN_TYPE_DISPLAY_BY_VALUE.get(saved_run_type, 'command (命令)')
            
            # 1. 强行修改下拉框的值
            if 'run_type' in self.param_widgets:
                self.param_widgets['run_type'].set(display_run_type)
            
            if 'run_type' in self.param_widgets:
                update_run_params(self.param_widgets, self.param_frame, self.param_widgets['run_type'])

        # ============================================================
        # 常规参数填充 (此时输入框已经显示出来了，可以安全填值了)
        # ============================================================
        
        # 预处理 Region 显示
        if 'region' in self.param_widgets:
            cb = step['params'].get('region', step['params'].get('cache_box'))
            # 记录被编辑步骤在加载前是否是 cache_box 而非 region
            self.editing_step_has_cache_box = ('cache_box' in step['params'] and 'region' not in step['params'])
            if isinstance(cb, list) and len(cb) == 4:
                self.param_widgets['region'].delete(0, tk.END)
                self.param_widgets['region'].insert(0, f"{cb[0]}, {cb[1]}, {cb[2]}, {cb[3]}")
        
        # 遍历并填充所有参数
        for k, v in step['params'].items():
            # 跳过 mode, run_type (前面处理了) 和 cache_box (前面处理了)
            if k in ('mode', 'run_type', 'cache_box', 'region'): continue
            
            if k in self.param_widgets:
                w = self.param_widgets[k]
                
                if k in ('lang', 'button', 'engine'):
                    display_val = param_internal_to_display(
                        k, v,
                        self.FULL_OCR_NAME_MAP,
                        MacroSchema.LANG_VALUES_TO_NAME,
                        MacroSchema.CLICK_VALUES_TO_NAME,
                        self.available_ocr_keys
                    )
                else:
                    display_val = v
                
                # 赋值
                if isinstance(w, tk.BooleanVar): w.set(bool(v))
                elif isinstance(w, ttk.Combobox): w.set(display_val)
                else: 
                    w.delete(0, tk.END)
                    w.insert(0, str(display_val))

        # 更新编辑状态
        self.editing_index = idx
        self.add_step_btn.config(text="[OK] 更新步骤", bootstyle="warning")
        self.add_step_btn.grid_configure(columnspan=1)
        self.cancel_edit_btn.grid(row=0, column=1, sticky="nsew", padx=(2,0))
        self.update_listbox_display()

    def cancel_edit_mode(self):
        self.editing_index = None
        self.editing_step_has_cache_box = False  # 重置标志
        self.add_step_btn.config(text="＋ 添加到序列 >>", bootstyle="success")
        self.cancel_edit_btn.grid_remove()
        self.add_step_btn.grid_configure(columnspan=2)
        self.update_listbox_display()

    def _format_step_params(self, step, act):
        # 参数预览文本
        display_params = step['params'].copy()
        
        cache_str = ""
        if 'region' in display_params or 'cache_box' in display_params:
            box = display_params.pop('region', display_params.pop('cache_box', None))
            if isinstance(box, (list, tuple)) and len(box) >= 4:
                cache_str = f"[区域: {box[0]},{box[1]},{box[2]},{box[3]}] "
            elif box is not None:
                cache_str = "[区域: 无效] "

        if 'engine' in display_params:
            # <--- 列表显示时也使用完整映射
            display_params['engine'] = self.FULL_OCR_NAME_MAP.get(display_params['engine'], display_params['engine'])
            
        # 格式化参数列字符串
        param_text = f"{cache_str}{display_params}" if display_params else ""
        
        # 备注动作特殊处理：显示为注释格式
        if act == 'NOTE':
            note_text = step['params'].get('text', '')
            param_text = f"// {note_text}" if note_text else "// (空备注)"

        formatter = _STEP_PARAM_PREVIEW_FORMATTERS.get(act)
        if formatter:
            param_text = formatter(step['params'])
        
        # 插入行 (Values对应: id, action, params)
        return param_text

    def _get_step_display_indent(self, action, block_stack):
        return max(0, len(block_stack) - (1 if action in _LIST_DEDENT_ACTIONS else 0))

    def _update_display_block_stack(self, action, block_stack):
        if _is_list_block_start(action):
            block_stack.append(action)
        elif action in _LIST_BLOCK_END_ACTIONS and block_stack:
            block_stack.pop()

    def _get_step_tree_tags(self, index, is_enabled):
        tags = []
        if index == self.editing_index:
            tags.append('editing')
        if not is_enabled:
            tags.append('disabled')
        return tuple(tags)

    def _build_step_tree_row(self, index, step, block_stack):
        act = step['action']
        indent_str = "    " * self._get_step_display_indent(act, block_stack)
        param_text = self._format_step_params(step, act)
        action_label = MacroSchema.ACTION_TRANSLATIONS.get(act, act)
        is_enabled = step.get('enabled', True)

        display_action = f"{indent_str}{action_label}"
        if not is_enabled:
            display_action = f"{indent_str}[屏蔽] {action_label}"

        values = (index + 1, display_action, param_text)
        tags = self._get_step_tree_tags(index, is_enabled)
        return act, values, tags

    def _focus_step_tree_item_if_editing(self, index, item_id):
        if index == self.editing_index:
            self.steps_tree.see(item_id)
            self.steps_tree.selection_set(item_id)

    def _is_step_toggle_allowed(self, index):
        action = self.steps[index].get('action', '')
        return action not in MacroSchema.CONTROL_FLOW_ACTIONS

    def _set_tree_menu_toggle_state(self, index):
        state = "normal" if self._is_step_toggle_allowed(index) else "disabled"
        self.tree_menu.entryconfig(0, state=state)

    def _get_context_menu_step_index(self, event):
        item = self.steps_tree.identify_row(event.y)
        if not item:
            return None

        self.steps_tree.selection_set(item)
        return self._get_selected_index()

    def _warn_control_flow_toggle_blocked(self):
        messagebox.showwarning("提示", "不可屏蔽流程控制节点（条件、循环），以防止引发严重 BUG。", parent=self.root)

    def _toggle_step_enabled_at(self, index):
        step = self.steps[index]
        step['enabled'] = not step.get('enabled', True)

    def update_listbox_display(self):
        """Refresh the Treeview display."""
        for item in self.steps_tree.get_children():
            self.steps_tree.delete(item)

        block_stack = []
        for i, step in enumerate(self.steps):
            act, values, tags = self._build_step_tree_row(i, step, block_stack)
            item_id = self.steps_tree.insert("", "end", values=values)
            if tags:
                self.steps_tree.item(item_id, tags=tags)

            self._focus_step_tree_item_if_editing(i, item_id)
            self._update_display_block_stack(act, block_stack)

    def show_tree_menu(self, event):
        """Show the step list context menu."""
        idx = self._get_context_menu_step_index(event)
        if idx is None:
            return

        self._set_tree_menu_toggle_state(idx)
        self.tree_menu.post(event.x_root, event.y_root)

    def toggle_step_enabled(self):
        """切换选中步骤的启用/屏蔽状态"""
        idx = self._get_selected_index()
        if idx is None:
            return

        if not self._is_step_toggle_allowed(idx):
            self._warn_control_flow_toggle_blocked()
            return

        self._toggle_step_enabled_at(idx)
        self.update_listbox_display()

    def remove_step(self):
        # --- 升级: 适配 Treeview ---
        idx = self._get_selected_index()
        if idx is None: return
        
        # [修复] 使用 elif 确保逻辑互斥
        # 原代码问题: cancel_edit_mode 会将 editing_index 设为 None，
        # 导致后续的 if 判断永远为 False，索引调整失效
        if self.editing_index == idx:
            self.cancel_edit_mode()
        elif self.editing_index is not None and self.editing_index > idx:
            self.editing_index -= 1
            
        del self.steps[idx]
        self.update_listbox_display()
        
        # 尝试选中下一行
        children = self.steps_tree.get_children()
        if idx < len(children):
             self.steps_tree.selection_set(children[idx])
        elif children:
             self.steps_tree.selection_set(children[-1])

    def move_step(self, d):
        # --- 升级: 适配 Treeview ---
        idx = self._get_selected_index()
        if idx is None: return
        
        i = idx
        new_i = i - 1 if d == "up" else i + 1
        
        if 0 <= new_i < len(self.steps):
            self.steps.insert(new_i, self.steps.pop(i))
            
            # 同步更新 editing_index
            if self.editing_index == i: self.editing_index = new_i
            elif self.editing_index == new_i: self.editing_index = i
            self.update_listbox_display()
            
            # 保持选中移动后的项
            children = self.steps_tree.get_children()
            if 0 <= new_i < len(children):
                self.steps_tree.selection_set(children[new_i])
                self.steps_tree.see(children[new_i])

    def safe_run_macro(self):
        # [修复BUG-5] 步骤为空时给出明确提示，而非静默无响应
        if not self.is_macro_running and not self._run_pending and self.editing_index is None:
            if not self.steps:
                self.root.after(0, self.status_var.set, '提示: 宏为空，请先添加步骤再运行')
                return
            self.root.after(0, self.run_macro, True)
        
    def safe_stop_macro(self):
        """Request a cooperative macro stop; force-inject only as a delayed fallback."""
        if self._stop_in_progress:
            return
        if self._run_pending:
            self._run_pending = False
            if self._pending_run_id is not None:
                self.root.after_cancel(self._pending_run_id)
                self._pending_run_id = None
            self.status_var.set("已取消待执行的宏")
            self._restore_macro_idle_ui()
            return
        if not self.is_macro_running:
            return
        self._stop_in_progress = True
        self.root.after(0, self.status_var.set, "正在停止...")
        if self.current_run_context:
            self.current_run_context['stop_requested'] = True
            macro_engine.cleanup_active_processes(self.current_run_context)
        self.root.after(2500, self._force_stop_macro_if_needed)

    def _force_stop_macro_if_needed(self):
        """Last-resort stop for code paths that do not reach cooperative checks."""
        if not self._stop_in_progress or not self.is_macro_running:
            return
        t = self._macro_thread
        if not (t and t.is_alive()):
            return
        tid = t.ident
        if not tid:
            print("中断: thread ID invalid; exception not injected")
            return
        if not (self.current_run_context or {}).get('allow_force_thread_stop', False):
            print("Stop: cooperative stop timed out; force thread injection is disabled")
            if self.current_run_context:
                macro_engine.cleanup_active_processes(self.current_run_context)
            return
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(
            ctypes.c_ulong(tid),
            ctypes.py_object(macro_engine.MacroStopException)
        )
        if res == 0:
            print("Stop: thread ID invalid; exception not injected")
        elif res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_ulong(tid), None)
            print("Stop: exception affected multiple threads and was reverted")
        else:
            print("Stop: cooperative stop timed out; MacroStopException injected")
        
    def run_macro(self, hotkey=False):
        if self.is_macro_running or self._run_pending or not self.steps: return
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        
        if not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno("运行", f"是否立即开始？(按 {stop_display} 停止)"): return

        run_steps = [s for s in self.steps if s.get('action') == 'RUN' and s.get('enabled', True)]
        if run_steps and self.run_enabled_var.get() and not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno(
                "安全警告",
                f"此宏包含 {len(run_steps)} 个执行外部命令的步骤（RUN）。\n\n"
                "执行外部命令可能存在安全风险，请确保来源可信。\n\n"
                "是否继续运行？\n"
                "(可在左下角开关中永久禁用 RUN 步骤)"
            ): return
            
        self.loop_status_var.set("") 
        
        # [核心修复] 暴力清空之前的状态队列，防止积压
        while not self.status_queue.empty():
            try: self.status_queue.get_nowait()
            except queue.Empty: break
            
        self.run_btn.config(state="disabled")
        self.status_var.set(f"宏正在运行... [{stop_display}] 停止")
        
        # [新增] 创建迷你状态栏窗口（在最小化前）
        if not self.dont_minimize_var.get():
            if getattr(self, 'mini_status_window', None):
                try:
                    self.mini_status_window.destroy()
                except Exception:
                    pass
                self.mini_status_window = None
            self.mini_status_window = MiniStatusWindow(self.root, self.safe_stop_macro)
            self.mini_status_window.update_status(
                f"宏正在运行... [点击停止 或 {stop_display}]",
                ""
            )
            self.root.iconify()
        else:
            self.root.attributes('-topmost', True) 
        self._run_pending = True
        self._pending_run_id = self.root.after(600, self._start_macro_thread)

    def _start_macro_thread(self):
        self._run_pending = False
        self._pending_run_id = None
        self.is_macro_running = True
        macro_base_dir = os.path.dirname(self.current_filepath) if self.current_filepath else os.getcwd()
        self.current_run_context = {
            'stop_requested': False,
            'stop_key_str': self.hotkey_stop_str.get(),
            'enhanced_mode': self.enhanced_mode_var.get(),
            'run_enabled': self.run_enabled_var.get(),
            'macro_base_dir': macro_base_dir,
            'allowed_file_roots': [macro_base_dir, os.getcwd(), APP_DIR],
            'prompt_input_callback': self._prompt_input_for_macro,
        }
        self._macro_thread = threading.Thread(target=self._run, args=(copy.deepcopy(self.steps),), daemon=True)
        self._macro_thread.start()

    def _prompt_input_for_macro(self, title, prompt, default_value='', ctx=None):
        done = threading.Event()
        result = {'value': None}

        def ask():
            try:
                if ctx and ctx.get('stop_requested'):
                    done.set()
                    return
                if self.root.winfo_exists():
                    self.root.deiconify()
                    self.root.attributes('-topmost', True)
                    self.root.lift()
                result['value'] = simpledialog.askstring(
                    title or "智点助手输入",
                    prompt or "请输入内容:",
                    initialvalue=default_value or "",
                    parent=self.root
                )
            except Exception as e:
                result['error'] = e
            finally:
                try:
                    if self.root.winfo_exists() and self.dont_minimize_var.get():
                        self.root.attributes('-topmost', True)
                except Exception:
                    pass
                done.set()

        self.root.after(0, ask)
        while not done.wait(0.1):
            if ctx and ctx.get('stop_requested'):
                raise macro_engine.MacroStopException("用户在输入期间请求停止")

        if 'error' in result:
            raise result['error']
        return result.get('value')
        
    def _run(self, steps):
        try:
            macro_engine.execute_steps(steps, run_context=self.current_run_context, status_callback=self.update_loop_status)
        except macro_engine.MacroStopException:
            print("[宏] 已将循环强制中断")
        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("错误", str(err)))
        finally:
            self.root.after(0, self._on_macro_complete)

    def _restore_macro_idle_ui(self):
        if self.mini_status_window:
            self.mini_status_window.destroy()
            self.mini_status_window = None

        self.root.deiconify()
        self.root.attributes('-topmost', False)
        self.run_btn.config(state="normal")

    def _on_macro_complete(self):
        self.is_macro_running = False
        self._stop_in_progress = False
        if self.current_run_context:
            macro_engine.cleanup_active_processes(self.current_run_context)
        self.current_run_context = None

        self._restore_macro_idle_ui()
        self.update_status_bar_hotkeys() 

    def update_loop_status(self, text):
        self.status_queue.put(text)

    def _check_status_queue(self):
        """
        [补丁优化] 动态调整状态队列检查频率
        
        优化:
        - 运行时: 50ms (快速响应)
        - 空闲时: 500ms (节省CPU)
        """
        if not self.is_app_running: return
        
        # [补丁优化] 根据运行状态动态调整检查频率
        interval = STATUS_QUEUE_CHECK_INTERVAL_RUNNING if self.is_macro_running else STATUS_QUEUE_CHECK_INTERVAL_IDLE
        
        try:
            text = None
            count = 0
            while not self.status_queue.empty() and count < STATUS_QUEUE_MAX_BATCH:
                text = self.status_queue.get_nowait()
                count += 1
            
            if text:
                self.loop_status_var.set(text)
            
            # [新增] 同步更新迷你窗口（仅在内容变化时刷新）
            if self.mini_status_window:
                stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
                current_loop_status = self.loop_status_var.get()
                new_status = (f"宏正在运行... [点击停止 或 {stop_display}]", current_loop_status)
                if new_status != self._last_mini_status:
                    self._last_mini_status = new_status
                    self.mini_status_window.update_status(new_status[0], new_status[1])
        except queue.Empty:
            pass
        except Exception as e:
            print(f"[StatusQueue] 错误: {e}")
            
        self.root.after(interval, self._check_status_queue)

    def new_macro(self):
        if self.steps:
            if not messagebox.askyesno("新建", "清空当前宏？"): return
        self.steps = []
        self.editing_index = None
        self.last_test_location = None
        self.current_filepath = None
        self.cancel_edit_mode()
        self.update_listbox_display()
        self.update_title()
        self.status_var.set("已新建空白宏。")

    def load_macro(self):
        f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if f: self._load_file(f)

    def save_macro(self):
        """保存宏 (已重构)"""
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f:
            try:
                MacroPersistence.save(f, self.steps)
                self.current_filepath = f
                self.update_title()
                messagebox.showinfo("成功", "宏已保存！")
                self.add_to_recent_files(f)
            except Exception as e: messagebox.showerror("失败", str(e))

    def _load_file(self, f):
        """从文件加载宏 (已重构)"""
        if not os.path.exists(f):
            messagebox.showerror("失败", "文件不存在")
            if f in self.recent_files: 
                self.recent_files.remove(f); self.save_app_settings(); self.update_recent_files_menu()
            return
        try:
            self.cancel_edit_mode()
            data = MacroPersistence.load(f)
            
            # 验证JSON数据结构
            if not validate_macro_data(data):
                messagebox.showerror("加载失败", f"文件格式无效或损坏:\n{os.path.basename(f)}")
                return
            
            self.steps = data
            self.current_filepath = f
            self.update_listbox_display()
            self.update_title()
            self.status_var.set(f"已加载: {os.path.basename(f)}")
            self.add_to_recent_files(f)
        except Exception as e: 
            messagebox.showerror("加载失败", f"无法加载文件:\n{str(e)}")
    
    # _validate_macro_data 已迁移到 core_engine.py

    def add_to_recent_files(self, f):
        f = os.path.abspath(f)
        if f in self.recent_files: self.recent_files.remove(f)
        self.recent_files.insert(0, f)
        self.recent_files = self.recent_files[:MAX_RECENT_FILES]
        self.update_recent_files_menu()
        self.save_app_settings()

    def update_recent_files_menu(self):
        self.recent_files_menu.delete(0, tk.END)
        for i, f in enumerate(self.recent_files):
            self.recent_files_menu.add_command(label=f"{i+1}. {os.path.basename(f)}", command=lambda p=f: self._load_file(p))

    def load_app_settings(self):
        """加载应用设置"""
        try:
            config_path = CONFIG_FILE
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                    self.recent_files = d.get('recent_files', [])
                    self.current_theme.set(d.get('theme', 'litera'))
                    self.hotkey_run_str.set(d.get('hotkey_run', DEFAULT_HOTKEY_RUN))
                    self.hotkey_stop_str.set(d.get('hotkey_stop', DEFAULT_HOTKEY_STOP))
                    self.enhanced_mode_var.set(d.get('enhanced_mode', False))
                    self.run_enabled_var.set(d.get('run_enabled', False))
                    self.skip_confirm_var.set(d.get('skip_confirm', False))
                    self.dont_minimize_var.set(d.get('dont_minimize', False))
        except (OSError, json.JSONDecodeError, TypeError) as e:
            print(f"[设置] 加载应用设置失败: {e}")
        self.root.style.theme_use(self.current_theme.get())

    def save_app_settings(self):
        """保存应用设置"""
        try:
            settings = {}
            read_path = CONFIG_FILE
            if os.path.exists(read_path):
                try:
                    with open(read_path, 'r', encoding='utf-8') as f:
                        settings = json.load(f)
                    if not isinstance(settings, dict):
                        settings = {}
                except (OSError, json.JSONDecodeError, TypeError):
                    settings = {}
            settings.update({
                'recent_files': self.recent_files,
                'theme': self.current_theme.get(),
                'hotkey_run': self.hotkey_run_str.get(),
                'hotkey_stop': self.hotkey_stop_str.get(),
                'enhanced_mode': self.enhanced_mode_var.get(),
                'run_enabled': self.run_enabled_var.get(),
                'skip_confirm': self.skip_confirm_var.get(),
                'dont_minimize': self.dont_minimize_var.get()
            })
            os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
            tmp_path = CONFIG_FILE + '.tmp'
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, CONFIG_FILE)
        except (OSError, TypeError) as e:
            print(f"[设置] 保存应用设置失败: {e}")

    def change_theme(self):
        self.root.style.theme_use(self.current_theme.get())
        self.root.style.configure(".", font=self.font_ui)
        self.save_app_settings()
        



if __name__ == "__main__":
    import argparse
    
    # 命令行参数解析
    parser = argparse.ArgumentParser(description='MacroMate - 智点助手，智能桌面自动化工具')
    parser.add_argument('script_file', nargs='?', help='要执行的脚本文件 (.json)')
    parser.add_argument('--run', dest='run', help='执行指定脚本文件 (效果同直接传参)')
    parser.add_argument('--enable-run', action='store_true', help='允许命令行模式执行 RUN 步骤；默认禁用')
    parser.add_argument('--theme', dest='theme', default='litera', help='指定主题')
    parser.add_argument('--log-encoding', dest='log_encoding', default='', help='指定日志输出编码（如 utf-8 或 gbk）')
    args = parser.parse_args()
    
    # 确定要执行的脚本
    script_file = args.script_file or args.run
    
    if script_file:
        # 命令行模式：执行脚本
        if not os.path.exists(script_file):
            print(f"[CLI] ERROR: Script file not found: {script_file}")
            sys.exit(1)
        
        print(f"[CLI] Start script: {script_file}")
        
        try:
            # 加载脚本
            print("[CLI] Loading script...")
            with open(script_file, 'r', encoding='utf-8-sig') as f:
                script_data = json.load(f)
            
            # 支持两种格式:
            # 1. {"steps": [...]} - GUI 导出的格式
            # 2. [...] - 直接是步骤列表
            if isinstance(script_data, list):
                steps = script_data
            else:
                steps = script_data.get('steps', [])
            
            if not steps:
                print("[CLI] ERROR: No steps in script")
                sys.exit(1)
            
            # 执行脚本
            print(f"[CLI] Total steps: {len(steps)}, running...")
            macro_base_dir = os.path.dirname(os.path.abspath(script_file))
            run_context = {
                'run_enabled': args.enable_run,
                'macro_base_dir': macro_base_dir,
                'allowed_file_roots': [macro_base_dir, os.getcwd(), APP_DIR],
            }
            if not args.enable_run:
                print("[CLI] RUN steps are disabled by default. Use --enable-run to allow RUN actions.")
            result = macro_engine.execute_steps(steps, run_context=run_context)
            
            if result:
                print("[CLI] Script finished successfully")
            else:
                print("[CLI] Script failed")
                sys.exit(1)
                
        except macro_engine.MacroStopException as e:
            print(f"\n[CLI] 宏执行已被用户或系统安全机制中断: {e}")
            sys.exit(1)
        except Exception as e:
            import traceback
            error_msg = str(e)
            traceback_str = traceback.format_exc()
            print(f"[CLI] ERROR: {error_msg}")
            print(f"[CLI] TRACEBACK:\n{traceback_str}")
            sys.exit(1)
    else:
        # GUI 模式
        pyautogui.FAILSAFE = True
        try:
            theme = args.theme
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    theme_config = json.load(f)
                if isinstance(theme_config, dict):
                    theme = theme_config.get('theme', 'litera')
        except (OSError, json.JSONDecodeError, TypeError):
            pass
        main_window = tb.Window(themename=theme)
        app = MacroApp(main_window)
        main_window.mainloop()





