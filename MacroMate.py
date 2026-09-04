# -*- coding: utf-8 -*-
# MacroMate.py
# 功能说明：应用入口与 GUI 外壳，负责界面生命周期、宏运行控制和配置持久化
# Version: 1.8.6
APP_VERSION = "1.8.6"
MINI_STATUS_POSITION_MODES = frozenset({'above_taskbar', 'inside_taskbar'})

# 使用:
#   - GUI 模式: python MacroMate.py
#   - 命令行: python MacroMate.py script.json
#             python MacroMate.py --run script.json
#             python MacroMate.py --theme darkly (指定主题)

import os
import sys
import logging
from logging.handlers import RotatingFileHandler


_LOG_FORMAT = '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
_LOG_DATEFMT = '%Y-%m-%d %H:%M:%S'
_LOG_HANDLER_MARKER = '_macromate_owned_handler'


def _get_program_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _get_program_dir()


def _has_owned_log_handler(root_logger, handler_kind):
    return any(
        getattr(handler, _LOG_HANDLER_MARKER, None) == handler_kind
        for handler in root_logger.handlers
    )


def _configure_file_logging():
    """Install the bounded file handler early enough to capture startup logs."""
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    if _has_owned_log_handler(root_logger, 'file'):
        return None

    preferred_dir = os.environ.get('MACROMATE_LOG_DIR')
    if not preferred_dir:
        local_app_data = os.environ.get('LOCALAPPDATA')
        preferred_dir = (
            os.path.join(local_app_data, 'MacroMate', 'logs')
            if local_app_data else APP_DIR
        )

    formatter = logging.Formatter(_LOG_FORMAT, _LOG_DATEFMT)
    errors = []
    log_names = ('macromate.log', f'macromate-{os.getpid()}.log')
    for log_dir in dict.fromkeys((preferred_dir, APP_DIR)):
        try:
            os.makedirs(log_dir, exist_ok=True)
        except OSError as exc:
            errors.append(f'{log_dir}: {exc}')
            continue

        for log_name in log_names:
            log_path = os.path.join(log_dir, log_name)
            try:
                # Open immediately so a locked/unwritable file is detected here,
                # where the process-specific fallback can still be selected.
                handler = RotatingFileHandler(
                    log_path,
                    maxBytes=5 * 1024 * 1024,
                    backupCount=3,
                    encoding='utf-8',
                    delay=False,
                )
                handler.setLevel(logging.INFO)
                handler.setFormatter(formatter)
                setattr(handler, _LOG_HANDLER_MARKER, 'file')
                root_logger.addHandler(handler)
                return log_path
            except OSError as exc:
                errors.append(f'{log_path}: {exc}')

    root_logger.warning('文件日志不可用: %s', '; '.join(errors))
    return None


class _DynamicStderrHandler(logging.StreamHandler):
    """Resolve stderr at emit time so redirects and windowed builds stay safe."""
    def emit(self, record):
        if sys.stderr is None:
            return
        self.stream = sys.stderr
        super().emit(record)

def _configure_console_logging():
    """Install one stderr handler after sys_utils has configured stream encoding."""
    root_logger = logging.getLogger()
    if _has_owned_log_handler(root_logger, 'console'):
        return
    handler = _DynamicStderrHandler()
    level_name = os.environ.get('MACROMATE_CONSOLE_LOG_LEVEL', 'INFO').upper()
    handler.setLevel(getattr(logging, level_name, logging.INFO))
    handler.setFormatter(logging.Formatter(_LOG_FORMAT, _LOG_DATEFMT))
    setattr(handler, _LOG_HANDLER_MARKER, 'console')
    root_logger.addHandler(handler)


LOG_FILE = _configure_file_logging()

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
_configure_console_logging()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import re
import pyautogui
import threading
import ttkbootstrap as tb
import queue
import ctypes
ctypes.pythonapi.PyThreadState_SetAsyncExc.argtypes = [ctypes.c_ulong, ctypes.py_object]
ctypes.pythonapi.PyThreadState_SetAsyncExc.restype = ctypes.c_int


# =================================================================
# 全局配置
# =================================================================

APP_TITLE = f"智点助手 (MacroMate) v{APP_VERSION}"
APP_ICON = "app_icon.ico"
APP_USER_MODEL_ID = "hxlive.macromate"

CONFIG_FILE = os.path.join(APP_DIR, "macro_settings.mmcfg")
LEGACY_CONFIG_FILE = os.path.join(APP_DIR, "macro_settings.json")
_DEFAULT_CONFIG_FILE = CONFIG_FILE
MAX_RECENT_FILES = 8

DEFAULT_HOTKEY_RUN = "ctrl+f1"
DEFAULT_HOTKEY_STOP = "ctrl+f2"
LIGHT_THEMES = ('litera', 'cosmo', 'flatly', 'journal', 'lumen', 'minty', 'pulse', 'sandstone', 'united', 'yeti')
DARK_THEMES = ('superhero', 'cyborg', 'darkly', 'solar')
KNOWN_THEMES = frozenset((
    *LIGHT_THEMES, *DARK_THEMES,
    'cerculean', 'morph', 'simplex', 'vapor',
))
# =================================================================
# 性能优化常量
STATUS_QUEUE_CHECK_INTERVAL_IDLE = 500  # 空闲时状态队列检查间隔（毫秒）
STATUS_QUEUE_CHECK_INTERVAL_RUNNING = 50  # 运行时状态队列检查间隔（毫秒）
STATUS_QUEUE_MAX_BATCH = 50  # 状态队列单次最大处理数
FORCE_STOP_DELAY_MS = 2500
FORCE_STOP_VERIFY_MS = 1000



# 压制 noisy 第三方库
logging.getLogger('rapidocr').setLevel(logging.WARNING)

logger = logging.getLogger(__name__)

# 导入核心模块与工具类
try:
    import core_engine as macro_engine
    from step_controller import StepController, StepControllerServices
    import ocr_engine

    from sys_utils import (
        GlobalHotkeyManager, HotkeyUtils, MouseTracker,
        HotkeySettingsDialog, MiniStatusWindow,
        AboutDialog
    )
    from gui_utils import (
        ParamWidgetFactory, VLMSettingsDialog, get_icon_path
    )
    from core_engine import validate_macro_data
except ImportError as e:
    messagebox.showerror("导入错误", f"缺少必要的模块文件或导入失败: {e}\n请确保所有 py 文件都在同一目录。")
    exit()

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
        file_path = os.fspath(file_path)
        native_steps = MacroPersistence.convert_to_native(steps)
        tmp_path = file_path + '.tmp'
        try:
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
        except BaseException:
            try:
                os.remove(tmp_path)
            except FileNotFoundError:
                pass
            except OSError:
                logger.debug('Failed to clean macro temp file', exc_info=True)
            raise

    @staticmethod
    def load(file_path):
        """从 JSON 文件加载宏"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data


def capitalize_hotkey_str(s): return HotkeyUtils.format_hotkey_display(s)


def _resolve_app_config_path():
    """Return the readable settings path, migrating the production legacy file."""
    if os.path.normcase(os.path.abspath(CONFIG_FILE)) != os.path.normcase(os.path.abspath(_DEFAULT_CONFIG_FILE)):
        return CONFIG_FILE
    return sys_utils.migrate_legacy_config_file(CONFIG_FILE, LEGACY_CONFIG_FILE)


def _normalized_file_path(path):
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


def _is_reserved_config_path(path):
    candidate = _normalized_file_path(path)
    return candidate in {
        _normalized_file_path(CONFIG_FILE),
        _normalized_file_path(LEGACY_CONFIG_FILE),
    }


class MacroApp:

    # ================================================================
    # Lifecycle And Setup
    # ================================================================

    def __init__(self, root, icon_path=None):
        self.root = root
        self.app_icon_path = icon_path
        self._setup_window()
        self._setup_app_state()
        self._setup_hotkeys()
        self._setup_app_options()
        self._setup_mouse_tracker()
        self._setup_ocr_state()
        self._setup_widget_factory()
        self._setup_step_controller()
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

    def _setup_app_state(self):
        # === Macro Run State Machine ===
        self.is_macro_running = False
        self.current_run_context = None
        self._macro_thread = None
        self._stop_in_progress = False
        self._stop_request_id = 0
        self._run_pending = False
        self._pending_run_id = None

        # === Mini Status Window ===
        self.mini_status_window = None
        self._last_mini_status = (None, None)

        # === File ===
        self.recent_files = []
        self.current_filepath = None

        # === Main-thread Queues ===
        self.status_queue = queue.Queue(maxsize=1)
        self._ui_callback_queue = queue.Queue()

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
        self.mini_status_position_var = tb.StringVar(value='above_taskbar')

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


    def _setup_step_controller(self):
        services = StepControllerServices(
            post_to_ui=self._queue_ui_callback,
            is_window_alive=lambda: self.is_app_running,
            set_status=lambda text: self.root.after(
                0, lambda value=text: (
                    self.status_var.set(value)
                    if value else self.update_status_bar_hotkeys()
                )
            ),
            get_enhanced_mode=lambda: bool(self.enhanced_mode_var.get()),
            set_key_recording_active=self.hotkey_manager.set_key_recording_active,
            set_coordinate_capture=self.hotkey_manager.set_coordinate_capture,
            get_reserved_hotkeys=lambda: {
                "运行宏": self.hotkey_run_str.get(),
                "停止宏": self.hotkey_stop_str.get(),
            },
        )
        self.step_controller = StepController(
            root=self.root,
            widget_factory=self.widget_factory,
            mouse_tracker=self.mouse_tracker,
            ocr_engine_mapping=self.FULL_OCR_NAME_MAP,
            services=services,
        )
        self.step_controller.set_available_ocr_keys(self.available_ocr_keys)




    # ================================================================
    # Deferred Startup And Runtime Services
    # ================================================================

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
        self.step_controller.start_ui_services()
        self.load_app_settings()
        self.update_recent_files_menu()
        self.update_status_bar_hotkeys()
        self.root.after(500, self.hotkey_manager.check_conflicts)
        self.hotkey_manager.start_listener()
        threading.Thread(target=self._detect_ocr_engines_bg, daemon=True).start()
        self._check_status_queue()

    def _detect_ocr_engines_bg(self):
        """后台线程：检测可用 OCR 引擎并预热，完成后回调主线程更新状态。"""
        try:
            engines = ocr_engine.get_available_engines()
            # 检测完成后顺手预热（合并两次后台任务）
            ocr_engine.preload_engines()
        except Exception:
            logger.exception("后台检测异常")
            engines = [('none', '无可用OCR引擎')]
        self._queue_ui_callback(lambda engines=engines: self._on_ocr_engines_ready(engines))

    def _on_ocr_engines_ready(self, engines):
        """主线程回调：OCR 引擎检测完成，更新状态。"""
        # [修复H-7] 应用可能在检测期间已关闭，需先检查生命周期
        if not self.is_app_running:
            return
        self.available_ocr_engines = engines
        self.available_ocr_keys = [e[0] for e in engines]
        self.step_controller.set_available_ocr_keys(self.available_ocr_keys)
        if 'none' in self.available_ocr_keys:
            logger.warning("未找到任何可用的OCR引擎 (RapidOCR, Tesseract, WinOCR)。")
            try:
                self.status_var.set("WARN 未找到可用 OCR 引擎，文本查找功能不可用。")
            except Exception:
                pass
        else:
            engine_names = ' / '.join(e[1] for e in engines)
            logger.info(f"引擎就绪: {engine_names}")


    # ================================================================
    # Menu Construction
    # ================================================================

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

        self._bind_file_shortcuts()

    def _bind_file_shortcuts(self):
        """Bind file shortcuts without relying on a keyboard-layout-specific keysym."""
        self.root.bind('<Control-KeyPress>', self._handle_file_shortcut, add='+')

    def _handle_file_shortcut(self, event):
        shortcut_actions = {
            'n': 'new_macro',
            'o': 'load_macro',
            's': 'save_macro',
        }
        key = str(getattr(event, 'keysym', '') or '').lower()
        action_name = shortcut_actions.get(key)
        if action_name is None and sys.platform == 'win32':
            action_name = {
                ord('N'): 'new_macro',
                ord('O'): 'load_macro',
                ord('S'): 'save_macro',
            }.get(getattr(event, 'keycode', None))
        if action_name is None:
            return None
        return self._run_file_shortcut(action_name)

    def _run_file_shortcut(self, action_name):
        getattr(self, action_name)()
        return 'break'

    def _build_settings_menu(self):
        settings_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  设置  ", menu=settings_menu)
        settings_menu.add_command(label="⌨ 快捷键设置...", command=self.open_hotkey_settings)
        settings_menu.add_separator()
        settings_menu.add_command(label="🤖 AI 设置...", command=self.open_vlm_settings)
        settings_menu.add_separator()
        mini_status_menu = tk.Menu(settings_menu, tearoff=0, font=self.font_ui)
        settings_menu.add_cascade(label="悬浮条位置", menu=mini_status_menu)
        mini_status_menu.add_radiobutton(
            label="紧贴任务栏上沿（推荐）",
            variable=self.mini_status_position_var,
            value='above_taskbar',
            command=self.save_app_settings,
        )
        mini_status_menu.add_radiobutton(
            label="任务栏内部（实验）",
            variable=self.mini_status_position_var,
            value='inside_taskbar',
            command=self.save_app_settings,
        )

    def _build_theme_menu(self):
        theme_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  主题  ", menu=theme_menu)

        for theme in LIGHT_THEMES:
            theme_menu.add_radiobutton(label=f"亮 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)
        theme_menu.add_separator()
        for theme in DARK_THEMES:
            theme_menu.add_radiobutton(label=f"暗 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)

    def _build_about_menu(self):
        about_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  关于  ", menu=about_menu)
        about_menu.add_command(label="关于", command=self.show_about_dialog)


    # ================================================================
    # Main UI Construction
    # ================================================================

    def _init_ui(self):
        self._build_status_bar()
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self._build_step_list_panel(main_frame)
        self.step_controller.build_step_form(main_frame)
        self.step_controller.update_param_fields(None)

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
        return self.step_controller.build_step_tree(list_frame)

    def _build_step_list_controls(self, list_frame):
        left_bottom_frame = ttk.Frame(list_frame)
        left_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        left_bottom_frame.columnconfigure(0, weight=1); left_bottom_frame.columnconfigure(1, weight=1)
        left_bottom_frame.columnconfigure(2, weight=1); left_bottom_frame.columnconfigure(3, weight=1)

        self.step_controller.build_step_controls(left_bottom_frame)
        self._build_runtime_controls(left_bottom_frame)

    def _build_runtime_controls(self, left_bottom_frame):
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

    # ================================================================
    # Window And Status Helpers
    # ================================================================

    def _get_hotkey_display_pair(self):
        run_display = capitalize_hotkey_str(self.hotkey_run_str.get())
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        return run_display, stop_display

    def _format_status_bar_text(self, run_display, stop_display):
        return f"准备就绪...  |  [{run_display}] 启动宏  |  [{stop_display}] 停止宏"

    def _format_run_button_text(self, run_display):
        return f"▶ 运行宏 ({run_display})"

    def _format_mini_run_status(self, stop_display, loop_status=""):
        """Return compact left/right text while keeping the stop hotkey visible."""
        progress = loop_status.strip() if isinstance(loop_status, str) else ""
        progress = re.sub(r'\s*/\s*', '/', progress)
        left_text = f"MacroMate：{progress or '运行中'}"
        return left_text, f"｜{stop_display} 停止"

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


    # ================================================================
    # Dialogs And Exit
    # ================================================================

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
            conflicts = self.step_controller.get_press_key_hotkey_conflicts({
                "运行宏": new_run,
                "停止宏": new_stop,
            })
            if conflicts:
                details = "\n".join(
                    f"步骤 {index}: {capitalize_hotkey_str(key_chord)} "
                    f"与“{purpose}”快捷键冲突"
                    for index, purpose, key_chord in conflicts
                )
                messagebox.showwarning(
                    "步骤快捷键冲突",
                    "新的全局快捷键与已有“按下按键”步骤冲突，设置未保存。\n\n"
                    f"{details}\n\n请先修改冲突步骤或选择其他全局快捷键。",
                    parent=self.root,
                )
                return
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
        self.step_controller.stop_input_capture()
        self.save_app_settings()

        if not self.hotkey_manager.check_conflicts(show_success=False):
            messagebox.showwarning("冲突警告", "快捷键已保存，但检测到冲突。\n请确保没有其他程序占用它。", parent=self.root)

        self.hotkey_manager.restart_listener()
        self.step_controller.refresh_coordinate_capture()
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

        icon_path = self.app_icon_path or get_icon_path(APP_ICON, APP_VERSION)
        self._about_dialog_ref = AboutDialog(self.root, APP_VERSION, icon_path)

    def on_exit(self):
        self.is_app_running = False
        self.safe_stop_macro()
        if self.current_run_context:
            macro_engine.cleanup_active_processes(self.current_run_context)

        # [变更] 使用 MouseTracker 类停止
        self.mouse_tracker.stop()

        if self.hotkey_manager:
            logger.info("正在停止快捷键监听器...")
            try:
                self.hotkey_manager.shutdown()
            except Exception:
                logger.exception("停止监听器时出错")

        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass


    # ================================================================
    # Action Parameter Form
    # ================================================================

    # ================================================================
    # Region Selection
    # ================================================================

    # ================================================================
    # Macro Run And Stop
    # ================================================================

    def safe_run_macro(self):
        # 步骤为空时给出明确提示，而非静默无响应
        if not self.is_macro_running and not self._run_pending and not self.step_controller.is_editing():
            if not self.step_controller.has_steps():
                self.root.after(0, self.status_var.set, '提示: 宏为空，请先添加步骤再运行')
                return
            self.root.after(0, self.run_macro, True)

    def _clear_status_queue(self):
        cleared = 0
        while not self.status_queue.empty():
            try:
                self.status_queue.get_nowait()
                cleared += 1
            except queue.Empty:
                break
        return cleared
    def run_macro(self, hotkey=False):
        if self.is_macro_running or self._run_pending or not self.step_controller.has_steps(): return
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())

        if not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno("运行", f"是否立即开始？(按 {stop_display} 停止)"): return

        run_step_count = self.step_controller.count_enabled_action('RUN')
        if run_step_count and self.run_enabled_var.get() and not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno(
                "安全警告",
                f"此宏包含 {run_step_count} 个执行外部命令的步骤（RUN）。\n\n"
                "执行外部命令可能存在安全风险，请确保来源可信。\n\n"
                "是否继续运行？\n"
                "(可在左下角开关中永久禁用 RUN 步骤)"
            ): return

        self.step_controller.stop_input_capture()
        self.loop_status_var.set("")

        # 清空之前的状态队列，防止积压
        self._clear_status_queue()

        self.run_btn.config(state="disabled")
        self.status_var.set(f"宏正在运行... [{stop_display}] 停止")

        # [新增] 创建迷你状态栏窗口（在最小化前）
        if not self.dont_minimize_var.get():
            self._show_mini_status_window(stop_display)
        else:
            self.root.attributes('-topmost', True)
        self._run_pending = True
        self._pending_run_id = self.root.after(600, self._start_macro_thread)

    def _show_mini_status_window(self, stop_display):
        """销毁旧迷你窗并创建新的，主窗最小化。"""
        if getattr(self, 'mini_status_window', None):
            try:
                self.mini_status_window.destroy()
            except Exception:
                pass
            self.mini_status_window = None
        self.mini_status_window = MiniStatusWindow(
            self.root,
            position_mode=self.mini_status_position_var.get(),
            icon_path=self.app_icon_path,
        )
        self.mini_status_window.update_status(
            *self._format_mini_run_status(stop_display)
        )
        self.root.iconify()

    def _start_macro_thread(self):
        self._run_pending = False
        self._pending_run_id = None
        self.is_macro_running = True
        self.current_run_context = self._build_run_context()
        steps_snapshot = self.step_controller.get_steps_snapshot()
        self._macro_thread = threading.Thread(
            target=self._run, args=(steps_snapshot,), daemon=True
        )
        self._macro_thread.start()

    def _build_run_context(self):
        """Build the RunContext passed to core_engine.execute_steps."""
        macro_base_dir = os.path.dirname(self.current_filepath) if self.current_filepath else os.getcwd()
        return macro_engine.RunContext({
            'stop_requested': False,
            'stop_event': threading.Event(),
            'stop_key_str': self.hotkey_stop_str.get(),
            'enhanced_mode': self.enhanced_mode_var.get(),
            'run_enabled': self.run_enabled_var.get(),
            'macro_base_dir': macro_base_dir,
            'allowed_file_roots': [macro_base_dir, os.getcwd(), APP_DIR],
            'prompt_input_callback': self._prompt_input_for_macro,
            'allow_force_thread_stop': True,
        })

    def _prompt_input_for_macro(self, title, prompt, default_value='', ctx=None):
        done = threading.Event()
        result = {'value': None}

        def ask():
            try:
                if ctx is not None and macro_engine.is_stop_requested(ctx):
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

        self._queue_ui_callback(ask)
        while not done.wait(0.1):
            if ctx is not None and macro_engine.is_stop_requested(ctx):
                raise macro_engine.MacroStopException("用户在输入期间请求停止")

        if 'error' in result:
            raise result['error']
        return result.get('value')

    def _run(self, steps):
        try:
            succeeded = macro_engine.execute_steps(
                steps,
                run_context=self.current_run_context,
                status_callback=self.update_loop_status,
            )
            if not succeeded:
                self._queue_ui_callback(lambda: messagebox.showwarning(
                    "执行未完成",
                    "宏未完整执行。请检查状态信息和日志以确定失败步骤。",
                ))
        except macro_engine.MacroStopException:
            logger.info("已将循环强制中断")
        except Exception as e:
            self._queue_ui_callback(lambda err=e: messagebox.showerror("错误", str(err)))
        finally:
            self._queue_ui_callback(self._on_macro_complete)

    def _cancel_pending_macro_start(self):
        if not self._run_pending:
            return False
        self._run_pending = False
        if self._pending_run_id is not None:
            self.root.after_cancel(self._pending_run_id)
            self._pending_run_id = None
        return True
    def safe_stop_macro(self):
        """Request a cooperative macro stop; force-inject only as a delayed fallback."""
        if self._stop_in_progress:
            return
        if self._cancel_pending_macro_start():
            self.status_var.set("已取消待执行的宏")
            self._restore_macro_idle_ui()
            return
        if not self.is_macro_running:
            return
        self._stop_in_progress = True
        self._stop_request_id = getattr(self, '_stop_request_id', 0) + 1
        request_id = self._stop_request_id
        self.root.after(0, self.status_var.set, "正在停止...")
        if self.current_run_context:
            macro_engine.request_stop(self.current_run_context)
            macro_engine.cleanup_active_processes(self.current_run_context)
        self.root.after(FORCE_STOP_DELAY_MS, self._force_stop_macro_if_needed, request_id)

    def _allow_stop_retry(self, request_id, message):
        if request_id != getattr(self, '_stop_request_id', 0) or not self.is_macro_running:
            return
        self._stop_in_progress = False
        self.status_var.set(message)

    def _verify_force_stop_result(self, request_id):
        """Release the stop latch if async injection did not finish the worker."""
        if request_id != getattr(self, '_stop_request_id', 0):
            return
        thread = self._macro_thread
        if self.is_macro_running and thread and thread.is_alive():
            self._allow_stop_retry(request_id, "仍在停止，可再次按停止快捷键重试")

    def _force_stop_macro_if_needed(self, request_id=None):
        """Last-resort stop for code paths that do not reach cooperative checks."""
        if request_id is None:
            request_id = getattr(self, '_stop_request_id', 0)
        if request_id != getattr(self, '_stop_request_id', 0):
            return
        if not self._stop_in_progress or not self.is_macro_running:
            return
        t = self._macro_thread
        if not (t and t.is_alive()):
            self._stop_in_progress = False
            return
        tid = t.ident
        if not tid:
            logger.error("中断: thread ID invalid; exception not injected")
            self._allow_stop_retry(request_id, "停止未完成，可再次按停止快捷键重试")
            return
        force_stop_enabled = bool(
            self.current_run_context
            and self.current_run_context.get_option('allow_force_thread_stop', False)
        )
        if not force_stop_enabled:
            logger.warning("Stop: cooperative stop timed out; force thread injection is disabled")
            if self.current_run_context:
                macro_engine.cleanup_active_processes(self.current_run_context)
            self._allow_stop_retry(request_id, "停止未完成，可再次按停止快捷键重试")
            return
        try:
            res = ctypes.pythonapi.PyThreadState_SetAsyncExc(
                ctypes.c_ulong(tid),
                ctypes.py_object(macro_engine.MacroStopException)
            )
        except Exception:
            logger.exception("Stop: async exception injection failed")
            self._allow_stop_retry(request_id, "停止未完成，可再次按停止快捷键重试")
            return
        if res == 0:
            logger.error("Stop: thread ID invalid; exception not injected")
            self._allow_stop_retry(request_id, "停止未完成，可再次按停止快捷键重试")
        elif res > 1:
            try:
                ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_ulong(tid), None)
                logger.warning("Stop: exception affected multiple threads and was reverted")
            except Exception:
                logger.exception("Stop: failed to revert multi-thread async exception")
            finally:
                self._allow_stop_retry(request_id, "停止未完成，可再次按停止快捷键重试")
        else:
            logger.info("Stop: cooperative stop timed out; MacroStopException injected")
            self.root.after(FORCE_STOP_VERIFY_MS, self._verify_force_stop_result, request_id)

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
        self._stop_request_id = getattr(self, '_stop_request_id', 0) + 1
        if self.current_run_context:
            macro_engine.cleanup_active_processes(self.current_run_context)
        self.current_run_context = None

        self._restore_macro_idle_ui()
        self.step_controller.refresh_coordinate_capture()
        self.update_status_bar_hotkeys()


    # ================================================================
    # Status Queue
    # ================================================================

    def update_loop_status(self, text):
        while True:
            try:
                self.status_queue.put_nowait(text)
                return
            except queue.Full:
                try:
                    self.status_queue.get_nowait()
                except queue.Empty:
                    pass

    def _queue_ui_callback(self, callback):
        if self.is_app_running:
            self._ui_callback_queue.put(callback)

    def _drain_ui_callbacks(self):
        count = 0
        while count < STATUS_QUEUE_MAX_BATCH:
            try:
                callback = self._ui_callback_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception as e:
                logger.error(f"error: {e}")
            count += 1
        return count

    def _check_status_queue(self):
        """
        [补丁优化] 动态调整状态队列检查频率

        优化:
        - 运行时: 50ms (快速响应)
        - 空闲时: 500ms (节省CPU)
        """
        if not self.is_app_running: return

        self._drain_ui_callbacks()

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
                new_status = self._format_mini_run_status(stop_display, current_loop_status)
                if new_status != self._last_mini_status:
                    self._last_mini_status = new_status
                    self.mini_status_window.update_status(*new_status)
        except queue.Empty:
            pass
        except Exception as e:
            logger.error(f"错误: {e}")

        self.root.after(interval, self._check_status_queue)


    # ================================================================
    # File Operations
    # ================================================================

    def new_macro(self):
        if self.step_controller.has_steps():
            if not messagebox.askyesno("新建", "清空当前宏？"): return
        self.step_controller.clear_steps()
        self.current_filepath = None
        self.update_title()
        self.status_var.set("已新建空白宏。")

    def load_macro(self):
        f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if f: self._load_file(f)

    def save_macro(self):
        """保存宏到 JSON 文件。"""
        steps = self.step_controller.get_steps_snapshot()
        if not validate_macro_data(steps):
            messagebox.showerror(
                "保存失败",
                "当前步骤的控制结构不完整或嵌套顺序错误，请修正后再保存。",
            )
            return
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f:
            if _is_reserved_config_path(f):
                messagebox.showerror(
                    "保存失败",
                    "不能使用 MacroMate 的配置文件路径保存宏，请选择其他文件名。",
                )
                return
            try:
                MacroPersistence.save(f, steps)
                self.current_filepath = f
                self.update_title()
                messagebox.showinfo("成功", "宏已保存！")
                self.add_to_recent_files(f)
            except Exception as e: messagebox.showerror("失败", str(e))

    def _load_file(self, f):
        """从 JSON 文件加载宏。"""
        if not os.path.exists(f):
            messagebox.showerror("失败", "文件不存在")
            if f in self.recent_files:
                self.recent_files.remove(f); self.save_app_settings(); self.update_recent_files_menu()
            return
        try:
            data = MacroPersistence.load(f)

            # 验证JSON数据结构
            if not validate_macro_data(data):
                messagebox.showerror("加载失败", f"文件格式无效或损坏:\n{os.path.basename(f)}")
                return

            self.step_controller.load_steps(data)
            self.current_filepath = f
            self.update_title()
            self.status_var.set(f"已加载: {os.path.basename(f)}")
            self.add_to_recent_files(f)
        except Exception as e:
            messagebox.showerror("加载失败", f"无法加载文件:\n{str(e)}")

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


    # ================================================================
    # App Settings And Theme
    # ================================================================

    def load_app_settings(self):
        """加载应用设置"""
        self._apply_app_settings(self._read_app_settings())
        self.root.style.theme_use(self.current_theme.get())

    def _read_app_settings(self):
        return self._read_app_settings_file(log_errors=True)

    def _read_app_settings_file(self, log_errors=False):
        read_path = _resolve_app_config_path()
        if not os.path.exists(read_path):
            return {}
        try:
            with open(read_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            return settings if isinstance(settings, dict) else {}
        except (OSError, json.JSONDecodeError, TypeError) as e:
            if log_errors:
                logger.warning("加载应用设置失败，使用默认设置: %s", e)
            return {}

    def _available_theme_names(self):
        try:
            names = self.root.style.theme_names()
        except Exception:
            names = ()
        return set(names) or set(KNOWN_THEMES)

    def _normalize_app_settings(self, settings):
        normalized = dict(settings) if isinstance(settings, dict) else {}

        recent_files = normalized.get('recent_files', [])
        if not isinstance(recent_files, list):
            recent_files = []
        normalized['recent_files'] = [
            path for path in recent_files
            if isinstance(path, str) and path.strip()
        ][:MAX_RECENT_FILES]

        available_themes = self._available_theme_names()
        theme = normalized.get('theme')
        if not isinstance(theme, str) or theme not in available_themes:
            theme = 'litera' if 'litera' in available_themes else sorted(available_themes)[0]
        normalized['theme'] = theme

        run_hotkey = normalized.get('hotkey_run')
        stop_hotkey = normalized.get('hotkey_stop')
        run_hotkey = run_hotkey.strip().lower() if isinstance(run_hotkey, str) else ''
        stop_hotkey = stop_hotkey.strip().lower() if isinstance(stop_hotkey, str) else ''
        if not HotkeyUtils.is_valid_hotkey(run_hotkey):
            run_hotkey = DEFAULT_HOTKEY_RUN
        if not HotkeyUtils.is_valid_hotkey(stop_hotkey):
            stop_hotkey = DEFAULT_HOTKEY_STOP
        if run_hotkey == stop_hotkey:
            run_hotkey, stop_hotkey = DEFAULT_HOTKEY_RUN, DEFAULT_HOTKEY_STOP
        normalized['hotkey_run'] = run_hotkey
        normalized['hotkey_stop'] = stop_hotkey

        mini_status_position = normalized.get('mini_status_position', 'above_taskbar')
        if mini_status_position not in MINI_STATUS_POSITION_MODES:
            mini_status_position = 'above_taskbar'
        normalized['mini_status_position'] = mini_status_position

        for key in ('enhanced_mode', 'run_enabled', 'skip_confirm', 'dont_minimize'):
            value = normalized.get(key, False)
            normalized[key] = value if isinstance(value, bool) else False
        return normalized

    def _apply_app_settings(self, settings):
        settings = self._normalize_app_settings(settings)
        self.recent_files = settings['recent_files']
        self.current_theme.set(settings['theme'])
        self.hotkey_run_str.set(settings['hotkey_run'])
        self.hotkey_stop_str.set(settings['hotkey_stop'])
        self.enhanced_mode_var.set(settings['enhanced_mode'])
        self.run_enabled_var.set(settings['run_enabled'])
        self.skip_confirm_var.set(settings['skip_confirm'])
        self.dont_minimize_var.set(settings['dont_minimize'])
        self.mini_status_position_var.set(settings['mini_status_position'])

    def save_app_settings(self):
        """保存应用设置"""
        try:
            with sys_utils.get_shared_file_lock(CONFIG_FILE):
                settings = self._read_existing_app_settings_for_save()
                settings.update(self._collect_app_settings())
                self._write_app_settings(settings)
        except (OSError, TypeError):
            logger.exception("保存应用设置失败")

    def _read_existing_app_settings_for_save(self):
        return self._read_app_settings_file()

    def _collect_app_settings(self):
        return {
            'recent_files': self.recent_files,
            'theme': self.current_theme.get(),
            'hotkey_run': self.hotkey_run_str.get(),
            'hotkey_stop': self.hotkey_stop_str.get(),
            'enhanced_mode': self.enhanced_mode_var.get(),
            'run_enabled': self.run_enabled_var.get(),
            'skip_confirm': self.skip_confirm_var.get(),
            'dont_minimize': self.dont_minimize_var.get(),
            'mini_status_position': self.mini_status_position_var.get(),
        }

    def _write_app_settings(self, settings):
        import vlm_engine
        stored_settings = vlm_engine.prepare_app_config_for_storage(settings)
        sys_utils.write_json_file_atomically(CONFIG_FILE, stored_settings)

    def change_theme(self):
        self.root.style.theme_use(self.current_theme.get())
        self.root.style.configure(".", font=self.font_ui)
        self.save_app_settings()




def _run_cli_script(script_file, enable_run):
    """CLI 模式：加载并执行脚本，执行完退出。"""
    if not os.path.exists(script_file):
        logger.error(f"ERROR: Script file not found: {script_file}")
        sys.exit(1)

    logger.info(f"Start script: {script_file}")

    try:
        steps = _load_cli_steps(script_file)
        if not steps:
            logger.error("ERROR: No steps in script")
            sys.exit(1)

        logger.info(f"Total steps: {len(steps)}, running...")
        run_context = _build_cli_run_context(script_file, enable_run)
        if not enable_run:
            logger.info("RUN steps are disabled by default. Use --enable-run to allow RUN actions.")
        _finish_cli_run(macro_engine.execute_steps(steps, run_context=run_context))

    except macro_engine.MacroStopException as e:
        logger.error(f"\n宏执行已被用户或系统安全机制中断: {e}")
        sys.exit(1)
    except Exception as e:
        import traceback
        logger.error(f"ERROR: {e}")
        logger.error(f"TRACEBACK:\n{traceback.format_exc()}")
        sys.exit(1)


def _load_cli_steps(script_file):
    logger.info("Loading script...")
    with open(script_file, 'r', encoding='utf-8-sig') as f:
        return _extract_cli_steps(json.load(f))


def _extract_cli_steps(script_data):
    # 支持 GUI 导出的 {"steps": [...]}，也支持直接保存的步骤列表。
    if isinstance(script_data, list):
        return script_data
    if isinstance(script_data, dict):
        return script_data.get('steps', [])
    return []


def _build_cli_run_context(script_file, enable_run):
    macro_base_dir = os.path.dirname(os.path.abspath(script_file))
    return macro_engine.RunContext({
        'run_enabled': enable_run,
        'stop_requested': False,
        'stop_event': threading.Event(),
        'macro_base_dir': macro_base_dir,
        'allowed_file_roots': [macro_base_dir, os.getcwd(), APP_DIR],
    })


def _finish_cli_run(result):
    if result:
        logger.info("Script finished successfully")
        return
    logger.error("Script failed")
    sys.exit(1)


def _run_gui(theme):
    """GUI 模式：创建主窗口并进入 mainloop。"""
    pyautogui.FAILSAFE = True
    if sys_utils.set_windows_app_id(APP_USER_MODEL_ID):
        logger.info("AppUserModelID set: %s", APP_USER_MODEL_ID)
    else:
        logger.warning("AppUserModelID could not be set before window creation")
    icon_path = get_icon_path(APP_ICON, APP_VERSION)
    main_window = tb.Window(themename=theme)
    main_window.withdraw()
    if sys_utils.apply_window_icon(main_window, icon_path, set_default=True):
        logger.info("icon set before first window display: %s", os.path.basename(icon_path))
    else:
        logger.warning("未找到或无法应用图标文件，使用默认图标")
    MacroApp(main_window, icon_path=icon_path)
    main_window.deiconify()
    main_window.mainloop()


def _run_cli_mode(args):
    """命令行模式：执行指定脚本。"""
    script_file = args.script_file or args.run
    _run_cli_script(script_file, args.enable_run)


def _run_gui_mode(args):
    """桌面 GUI 模式：解析主题并启动主窗口。"""
    _run_gui(_resolve_initial_theme(args.theme))

def _resolve_initial_theme(cli_theme):
    """从配置文件解析初始主题，失败时回退到 CLI 指定主题。"""
    fallback_theme = cli_theme if cli_theme in KNOWN_THEMES else 'litera'
    try:
        read_path = _resolve_app_config_path()
        if os.path.exists(read_path):
            with open(read_path, 'r', encoding='utf-8') as f:
                theme_config = json.load(f)
            if isinstance(theme_config, dict):
                configured_theme = theme_config.get('theme')
                if configured_theme in KNOWN_THEMES:
                    return configured_theme
    except (OSError, json.JSONDecodeError, TypeError):
        pass
    return fallback_theme


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='MacroMate - 智点助手，智能桌面自动化工具')
    parser.add_argument('script_file', nargs='?', help='要执行的脚本文件 (.json)')
    parser.add_argument('--run', dest='run', help='执行指定脚本文件 (效果同直接传参)')
    parser.add_argument('--enable-run', action='store_true', help='允许命令行模式执行 RUN 步骤；默认禁用')
    parser.add_argument('--theme', dest='theme', default='litera', help='指定主题')
    parser.add_argument('--log-encoding', dest='log_encoding', default='', help='指定日志输出编码（如 utf-8 或 gbk）')
    args = parser.parse_args()

    if args.script_file or args.run:
        _run_cli_mode(args)
    else:
        _run_gui_mode(args)
