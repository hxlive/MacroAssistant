# -*- coding: utf-8 -*-
# MacroAssistant.py
# 描述: 自动化宏的 GUI 界面
# 版本: 1.56.0
# 变更: (升级) 列表控件升级为 Treeview (分列显示)。
#       (新增) 增加悬浮图片预览功能 (鼠标悬停在步骤上自动显示)。
#       (依赖) 需要 Pillow 库 (PIL) 支持图片显示。

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import pyautogui
import time
import threading
import ttkbootstrap as tb
from pynput import keyboard
import os
import sys
import queue
from PIL import Image, ImageGrab, ImageTk
import functools

# 强制启用 DPI 感知，解决 125%/150% 缩放下的坐标偏移问题
try:
    if sys.platform == 'win32':
        import ctypes
        # 设置 DPI 感知级别为 "PerMonitorV2" (Awareness 2)
        # 这会让 ImageGrab 和 pyautogui 的坐标系强制对齐
        ctypes.windll.shcore.SetProcessDpiAwareness(2) 
except Exception:
    try:
        # 回退旧版 API (兼容 Win7/8)
        ctypes.windll.user32.SetProcessDPIAware()
    except: pass
    
# 依赖：快捷键冲突检测
try:
    if sys.platform == 'win32':
        import ctypes
        import ctypes.wintypes
        import win32con
        HOTKEY_CHECK_AVAILABLE = True
except ImportError:
    HOTKEY_CHECK_AVAILABLE = False
    print("[配置] ✗ 未找到 pywin32 库 (pip install pywin32)。将跳过快捷键冲突检测。")

# =================================================================
# 全局配置
# =================================================================
APP_VERSION = "1.56.0"
APP_TITLE = f"宏助手 (Macro Assistant) V{APP_VERSION}"
APP_ICON = "app_icon.ico" 
CONFIG_FILE = "macro_settings.json"
MAX_RECENT_FILES = 5

DEFAULT_HOTKEY_RUN = "Ctrl+F10"
DEFAULT_HOTKEY_STOP = "Ctrl+F11"
# =================================================================
# 性能优化常量
STATUS_QUEUE_CHECK_INTERVAL_IDLE = 500  # 空闲时状态队列检查间隔（毫秒）
STATUS_QUEUE_CHECK_INTERVAL_RUNNING = 50  # 运行时状态队列检查间隔（毫秒）
STATUS_QUEUE_MAX_BATCH = 50  # 状态队列单次最大处理数
OCR_PRELOAD_DELAY = 100  # OCR引擎预热延迟（毫秒）


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

try:
    import core_engine as macro_engine
    import ocr_engine
    from core_engine import HotkeyUtils, MacroSchema
    # [变更] 导入重构后的 gui_utils 组件
    import gui_utils
    from gui_utils import (
        RegionSelector, 
        HotkeyEntry, 
        HotkeySettingsDialog, 
        ImageTooltipManager, 
        MouseTracker, 
        AutoWrapLabel, 
        parse_region_string
    )
except ImportError as e:
    messagebox.showerror("导入错误", f"缺少必要的模块文件或导入失败: {e}\n请确保 core_engine.py, ocr_engine.py, gui_utils.py 都在同一目录。")
    exit()

# -----------------------------------------------------------------
# 快捷键录制与冲突检测
# -----------------------------------------------------------------
PYNPUT_TO_VK = HotkeyUtils.PYNPUT_TO_VK
VK_TO_PYNPUT = HotkeyUtils.VK_TO_PYNPUT

if HOTKEY_CHECK_AVAILABLE:
    PYNPUT_MOD_TO_WIN_MOD = {
        'ctrl': win32con.MOD_CONTROL,
        'alt': win32con.MOD_ALT,
        'shift': win32con.MOD_SHIFT,
        'cmd': win32con.MOD_WIN,
    }

def capitalize_hotkey_str(s):
    """辅助函数：将 ctrl+f10 转换为 Ctrl+F10"""
    return HotkeyUtils.format_hotkey_display(s)


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1140x730")  # 稍微加宽以适应优化后的列宽 
        
        self.font_ui = ("Microsoft YaHei UI", 10)
        self.font_code = ("Consolas", 10)
        
        self.root.style.configure(".", font=self.font_ui)
        # <--- Treeview 样式配置
        self.root.style.configure("Treeview", font=self.font_code, rowheight=25)
        self.root.style.configure("Treeview.Heading", font=self.font_ui)
        
        self.is_app_running = True
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        
        icon_path = resource_path(APP_ICON) 
        if os.path.exists(icon_path):
            try: self.root.iconbitmap(icon_path)
            except tk.TclError: pass
        
        self.steps = []
        self.editing_index = None
        self.is_macro_running = False
        self.last_test_location = None 
        self.current_run_context = None 
        self.held_keys = set()
        
        self.hotkey_run_str = tb.StringVar(value=DEFAULT_HOTKEY_RUN)
        self.hotkey_stop_str = tb.StringVar(value=DEFAULT_HOTKEY_STOP)
        self.hotkey_listener = None
        
        self.current_theme = tb.StringVar(value=self.root.style.theme_use())
        self.skip_confirm_var = tb.BooleanVar(value=False)
        self.dont_minimize_var = tb.BooleanVar(value=False)
        self.recent_files = []
        self.status_queue = queue.Queue()
        
        # [变更] 使用 MouseTracker 类替代原有的 job 和 func
        self.mouse_pos_var = tb.StringVar()
        self.mouse_tracker = MouseTracker(self.root, self.mouse_pos_var)
        
        # OCR 引擎健康检查与映射
        self.FULL_OCR_NAME_MAP = {
            'auto': '自动选择 (Auto)',
            'winocr': 'Windows 10/11 OCR',
            'rapidocr': 'RapidOCR',
            'tesseract': 'Tesseract OCR',
            'none': '无可用OCR引擎'
        }
        self.FULL_OCR_KEY_MAP = {name: key for key, name in self.FULL_OCR_NAME_MAP.items()}
        self.available_ocr_engines = ocr_engine.get_available_engines()
        self.available_ocr_keys = [e[0] for e in self.available_ocr_engines]
        
        if 'none' in self.available_ocr_keys:
            print("[警告] 未找到任何可用的OCR引擎 (RapidOCR, Tesseract, WinOCR)。")

        self._init_menu()
        self._init_ui()
        
        # [变更] 初始化悬浮预览管理器 (使用 lambda 动态获取 steps)
        self.tooltip_manager = ImageTooltipManager(self.steps_tree, lambda: self.steps)
        
        self.load_app_settings()
        self.update_recent_files_menu()
        self.update_status_bar_hotkeys() 
        self.root.after(500, self.check_hotkey_conflicts)
        self.start_hotkey_listener() 
        # [补丁优化] 提前预热OCR引擎，改善首次使用体验
        self.root.after(OCR_PRELOAD_DELAY, lambda: threading.Thread(target=ocr_engine.preload_engines, daemon=True).start())
        self._check_status_queue()

    def _init_menu(self):
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
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

        settings_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  设置  ", menu=settings_menu)
        settings_menu.add_command(label="⌨️ 快捷键设置...", command=self.open_hotkey_settings)

        theme_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  主题  ", menu=theme_menu)
        
        light_themes = ['litera', 'cosmo', 'flatly', 'journal', 'lumen', 'minty', 'pulse', 'sandstone', 'united', 'yeti']
        for theme in light_themes:
            theme_menu.add_radiobutton(label=f"亮 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)
        theme_menu.add_separator()
        dark_themes = ['superhero', 'cyborg', 'darkly', 'solar']
        for theme in dark_themes:
            theme_menu.add_radiobutton(label=f"暗 - {theme.capitalize()}", variable=self.current_theme, value=theme, command=self.change_theme)

    def _init_ui(self):
        status_bar_frame = ttk.Frame(self.root, bootstyle="primary")
        status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var = tk.StringVar()
        self.status_label_left = ttk.Label(status_bar_frame, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W, padding=5, bootstyle="primary-inverse", font=self.font_ui)
        self.status_label_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.loop_status_var = tk.StringVar()
        self.loop_status_label_right = ttk.Label(status_bar_frame, textvariable=self.loop_status_var, relief=tk.FLAT, anchor=tk.E, padding=(0, 5, 5, 5), bootstyle="primary-inverse", font=self.font_ui)
        self.loop_status_label_right.pack(side=tk.RIGHT)

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # =====================================================================
        # 左侧面板 (Treeview + Preview)
        # =====================================================================
        list_frame = ttk.Frame(main_frame, padding=10)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 标题栏
        title_frame = ttk.Frame(list_frame)
        title_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(title_frame, text="宏步骤序列:", font=("Microsoft YaHei UI", 11, "bold")).pack(side=tk.LEFT)
        
        # --- Treeview 替换 Listbox ---
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("id", "action", "params")
        self.steps_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        self.steps_tree.heading("id", text="#")
        self.steps_tree.heading("action", text="动作")
        self.steps_tree.heading("params", text="参数详情 / 备注")
        
        # 优化列宽：缩小序号列，适当缩小动作列，扩大参数列
        self.steps_tree.column("id", width=40, minwidth=35, stretch=False, anchor="center")
        self.steps_tree.column("action", width=300, minwidth=120, stretch=False)
        self.steps_tree.column("params", width=280, minwidth=250, stretch=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.steps_tree.yview)
        self.steps_tree.configure(yscrollcommand=scrollbar.set)
        
        self.steps_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定事件
        self.steps_tree.bind("<Double-1>", lambda e: self.load_step_for_edit())
        
        # 配置编辑行的样式
        self.steps_tree.tag_configure('editing', background='#FFF3CD')

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
        skip_check.grid(row=0, column=0, sticky="w", padx=2) 
        minimize_check = ttk.Checkbutton(check_frame, text="运行时主界面不最小化", variable=self.dont_minimize_var, bootstyle="primary-round-toggle")
        minimize_check.grid(row=0, column=1, sticky="w", padx=2)
        
        # =====================================================================
        # 右侧面板
        # =====================================================================
        add_frame = ttk.Labelframe(main_frame, text="添加新步骤", padding=10)
        add_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)
        
        add_frame.pack_propagate(False)  # 禁止子控件影响父容器
        add_frame.configure(width=380)   # 固定宽度
        
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
        self.action_type.bind("<<ComboboxSelected>>", self.update_param_fields)
        self.param_frame = ttk.Frame(add_frame)
        self.param_frame.pack(fill=tk.X, expand=True, pady=5)
        
        # [变更] 不再需要绑定 Configure 事件，AutoWrapLabel 会自动处理
        self.param_widgets = {}
        self.update_param_fields(None)

    # --- Treeview 辅助方法 ---
    def _param_display_to_internal(self, key, display_value):
        """
        将UI显示值转换为内部存储值
        
        解决重复代码问题：
        - add_or_update_step 中的转换逻辑
        - load_step_for_edit 中的转换逻辑
        
        Args:
            key: 参数键名 ('lang', 'button', 'engine' 等)
            display_value: UI中显示的值
            
        Returns:
            内部存储的实际值
        """
        # 定义映射表
        mappings = {
            'lang': MacroSchema.LANG_OPTIONS,
            'button': MacroSchema.CLICK_OPTIONS,
            'engine': self.FULL_OCR_KEY_MAP
        }
        
        # 特殊处理: engine 可能带 "(不可用)" 后缀
        if key == 'engine' and display_value.endswith(" (不可用)"):
            display_value = display_value.replace(" (不可用)", "")
        
        # 查找映射
        mapping = mappings.get(key)
        if mapping:
            return mapping.get(display_value, display_value)
        
        return display_value
    
    def _param_internal_to_display(self, key, internal_value):
        """
        将内部存储值转换为UI显示值
        
        Args:
            key: 参数键名
            internal_value: 内部存储的值
            
        Returns:
            UI中应该显示的值
        """
        # 定义反向映射表
        reverse_mappings = {
            'lang': MacroSchema.LANG_VALUES_TO_NAME,
            'button': MacroSchema.CLICK_VALUES_TO_NAME,
            'engine': self.FULL_OCR_NAME_MAP
        }
        
        mapping = reverse_mappings.get(key)
        if mapping:
            display_val = mapping.get(internal_value, internal_value)
            
            # 特殊处理: engine 不可用标记
            if key == 'engine' and internal_value not in self.available_ocr_keys and internal_value != 'auto':
                display_val = f"{display_val} (不可用)"
            
            return display_val
        
        return internal_value

    def _get_selected_index(self):
        """获取当前选中项的索引"""
        selected_items = self.steps_tree.selection()
        if not selected_items: return None
        return self.steps_tree.index(selected_items[0])

    def update_status_bar_hotkeys(self):
        """更新状态栏和运行按钮上的快捷键提示"""
        run_display = capitalize_hotkey_str(self.hotkey_run_str.get())
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        self.status_var.set(f"准备就绪...  |  [{run_display}] 启动宏  |  [{stop_display}] 停止宏")
        self.run_btn.config(text=f"▶ 运行宏 ({run_display})")

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
        
        if not self.check_hotkey_conflicts(show_success=False):
            messagebox.showwarning("冲突警告", "快捷键已保存，但检测到冲突。\n请确保没有其他程序占用它。", parent=self.root)
        
        self.restart_hotkey_listener()
        self.update_status_bar_hotkeys()

    def on_exit(self):
        self.is_app_running = False
        self.held_keys.clear()
        
        # [变更] 使用 MouseTracker 类停止
        self.mouse_tracker.stop()
            
        if self.hotkey_listener:
            print("[Info] 正在停止快捷键监听器...")
            try:
                self.hotkey_listener.stop()
                self.hotkey_listener.join(timeout=0.5) 
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
        
        if action_key in ('FIND_TEXT', 'IF_TEXT_FOUND'):
            if 'none' in self.available_ocr_keys:
                self._create_hint_label(self.param_frame, 
                    "✗ 错误: 未找到可用的OCR引擎。\n"
                    "请先安装 RapidOCR (推荐) 或 Tesseract，\n"
                    "然后重启本程序。",
                    bootstyle="danger")
                self.action_type.set(MacroSchema.ACTION_TRANSLATIONS['FIND_IMAGE'])
                self.update_param_fields(None)
                return
        
        if action_key == 'FIND_IMAGE':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_region_selector() # <--- 新增: 区域选择
            self.create_param_entry("confidence", "置信度(0.1-1.0):", "0.8")
            self._create_hint_label(self.param_frame, "* 提示：如果识别失败，请调低置信度")
            self.create_browse_button()
            self.create_test_button("🧪 测试查找图像", self.on_test_find_image_click)
            
        elif action_key == 'FIND_TEXT':
            self.create_param_entry("text", "查找的文本:", "确定")
            self.create_region_selector()
            self.create_param_combobox("lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
            self.create_ocr_engine_combobox()
            
            # === 新增：保存到剪贴板选项 ===
            self.create_param_checkbox("save_to_clipboard", "✓ 保存识别结果到剪贴板", default=False)
            self.create_param_entry("extract_pattern", "提取模式 (正则，可选):", r"\d+")
            self._create_hint_label(self.param_frame, 
                "* 提示: 勾选后，识别到的文本将保存到剪贴板"
                "* 提取模式: 用正则表达式过滤，如 \\d+ 提取数字")
            
            self.create_test_button("🧪 测试查找文本 (OCR)", self.on_test_find_text_click)
            
        elif action_key == 'MOVE_OFFSET':
            self.create_param_entry("x_offset", "X 偏移:", "10")
            self.create_param_entry("y_offset", "Y 偏移:", "0")
        elif action_key == 'CLICK':
            self.create_param_combobox("button", "按键:", list(MacroSchema.CLICK_OPTIONS.keys()))
        
        elif action_key == 'SCROLL':
            self.create_param_entry("amount", "滚动量 (正数=上, 负数=下):", "100")
            self.create_param_entry("x", "X 坐标 (可选):", "")
            self.create_param_entry("y", "Y 坐标 (可选):", "")
            self._create_hint_label(self.param_frame, "* 提示: 如果 X, Y 为空，将在当前鼠标位置滚动。")

        elif action_key == 'WAIT':
            self.create_param_entry("ms", "等待 (毫秒):", "500")
        elif action_key == 'TYPE_TEXT':
            self.create_param_entry("text", "输入文本:", "你好")
            self._create_hint_label(self.param_frame, 
                "* 此功能使用剪贴板 (Ctrl+V)，以支持中文及复杂文本输入。\n"
                "* 支持占位符: {CLIPBOARD} 将替换为剪贴板内容\n"
                "* 示例: '订单号: {CLIPBOARD}' → '订单号: 12345'")
        elif action_key == 'PRESS_KEY':
            self.create_param_entry("key", "按键或组合键 (Enter, Ctrl+C):", "Enter")
        
        elif action_key == 'ACTIVATE_WINDOW':
            self.create_param_entry("title", "窗口标题 (支持部分匹配):", "记事本")
            self._create_hint_label(self.param_frame, "* 提示: 宏将查找标题中包含此文本的窗口，并将其激活到最前端。")

        elif action_key == 'MOVE_TO':
            self.create_param_entry("x", "X 坐标:", "100")
            self.create_param_entry("y", "Y 坐标:", "100")
            
            ttk.Separator(self.param_frame, orient='horizontal').pack(fill='x', pady=(15, 5))
            ttk.Label(self.param_frame, text="当前鼠标位置 (参考):", font=self.font_ui, foreground='gray').pack(anchor="w", pady=(5,0))
            ttk.Label(self.param_frame, textvariable=self.mouse_pos_var, font=self.font_code, bootstyle="info").pack(anchor="w")
            # [变更] 启动鼠标追踪
            self.mouse_tracker.start()
            
        elif action_key == 'IF_IMAGE_FOUND':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_region_selector() 
            self.create_param_entry("confidence", "置信度:", "0.8")
            self.create_browse_button()
            self.create_test_button("🧪 测试 IF 图像", self.on_test_find_image_click)
            
        elif action_key == 'IF_TEXT_FOUND':
            self.create_param_entry("text", "查找文本:", "确定")
            self.create_region_selector() 
            self.create_param_combobox("lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
            self.create_ocr_engine_combobox()
            
            # === 新增：保存到剪贴板选项 ===
            self.create_param_checkbox("save_to_clipboard", "✓ 保存识别结果到剪贴板", default=False)
            self.create_param_entry("extract_pattern", "提取模式 (正则，可选):", r"\d+")
            
            self.create_test_button("🧪 测试 IF 文本", self.on_test_find_text_click)
            
        elif action_key == 'LOOP_START':
            # 循环模式选择
            mode_options = {
                '固定次数': 'fixed',
                '直到找到图像': 'until_image',
                '直到找到文本': 'until_text'
            }
            self.create_param_combobox("mode", "循环模式:", list(mode_options.keys()), default='固定次数')
            
            # 根据模式动态显示参数
            # 这里先创建所有可能的控件，后续通过 update_loop_params 动态显示/隐藏
            self.create_param_entry("times", "循环次数:", "10")
            self.create_param_entry("max_iterations", "最大迭代次数 (安全阀):", "1000")
            
            # 条件：图像
            self.create_param_entry("condition_image", "目标图像路径:", "target.png")
            self.create_param_entry("confidence", "置信度:", "0.8")
            
            # 条件：文本
            self.create_param_entry("condition_text", "目标文本:", "加载完成")
            self.create_param_combobox("lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
            
            self._create_hint_label(self.param_frame, 
                "* 提示:"
                "- 固定次数: 传统循环，执行指定次数"
                "- 直到找到图像: 找到图像即停止"
                "- 直到找到文本: 找到文本即停止"
                "- 最大迭代: 防止无限循环的安全机制")
            
            # 绑定模式切换事件
            if 'mode' in self.param_widgets:
                self.param_widgets['mode'].bind("<<ComboboxSelected>>", self.update_loop_params)
            
            # 初始化显示
            self.update_loop_params(None)
        elif action_key == 'ELSE':
            self._create_hint_label(self.param_frame, "* 提示: 'ELSE' 必须与 'IF' 配合使用。它将执行 'IF' 条件不满足时的逻辑。")
        elif action_key == 'END_IF':
            self._create_hint_label(self.param_frame, "* 提示: 'END_IF' 必须与 'IF' 配合使用。它标志着 'IF' 或 'ELSE' 逻辑块的结束。")
        elif action_key == 'END_LOOP':
            self._create_hint_label(self.param_frame, "* 提示: 'END_LOOP' 必须与 'LOOP_START' 配合使用。它标志着循环体的结束。")



    def update_loop_params(self, event):
        """根据循环模式动态显示/隐藏参数"""
        if 'mode' not in self.param_widgets:
            return
        
        mode_map = {
            '固定次数': 'fixed',
            '直到找到图像': 'until_image',
            '直到找到文本': 'until_text'
        }
        
        selected_mode = self.param_widgets['mode'].get()
        mode = mode_map.get(selected_mode, 'fixed')
        
        # === 改进：记住提示标签的位置 ===
        hint_labels = []
        for widget in self.param_frame.winfo_children():
            if isinstance(widget, AutoWrapLabel): # 检查是否是新的 AutoWrapLabel
                hint_labels.append(widget)
        
        # 隐藏所有条件参数
        for key in ['times', 'condition_image', 'confidence', 'condition_text', 'lang', 'max_iterations']:
            if key in self.param_widgets:
                widget = self.param_widgets[key]
                # 获取父 frame
                parent_frame = widget.master
                if parent_frame:
                    parent_frame.pack_forget()
        
        # 根据模式显示对应参数（在提示之前插入）
        params_to_show = []
        if mode == 'fixed':
            params_to_show = ['times']
        elif mode == 'until_image':
            params_to_show = ['condition_image', 'confidence', 'max_iterations']
        elif mode == 'until_text':
            params_to_show = ['condition_text', 'lang', 'max_iterations']
        
        # 显示参数
        for key in params_to_show:
            if key in self.param_widgets:
                self.param_widgets[key].master.pack(fill=tk.X, pady=8)
        
        # === 确保提示标签始终在最后 ===
        for hint_label in hint_labels:
            hint_label.pack_forget()
            hint_label.pack(anchor="w", pady=5, fill=tk.X)

    def create_param_entry(self, key, label_text, default_value):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        entry = ttk.Entry(frame, width=25, font=self.font_ui)  # 缩小宽度
        entry.insert(0, default_value)
        entry.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = entry
        

    def create_param_checkbox(self, key, label_text, default=False):
        frame = ttk.Frame(self.param_frame)
        var = tk.BooleanVar(value=default)
        checkbox = ttk.Checkbutton(frame, text=label_text, variable=var, 
                                   bootstyle="primary-round-toggle")
        checkbox.pack(anchor="w")
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = var  # 注意：存储的是 BooleanVar

    def create_param_combobox(self, key, label_text, values, default=None):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=23, font=self.font_ui)  # 缩小宽度
        if default and default in values:
            combo.set(default)
        else:
            combo.current(0)
        combo.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = combo
    
    def create_ocr_engine_combobox(self):
        combobox_values = ['自动选择 (Auto)']
        # 遍历 *所有* 引擎，而不仅仅是可用的引擎
        for key, name in self.FULL_OCR_NAME_MAP.items():
            if key in ('auto', 'none'): continue
            
            if key in self.available_ocr_keys:
                combobox_values.append(name) 
            else:
                combobox_values.append(f"{name} (不可用)") 
                
        self.create_param_combobox("engine", "OCR 引擎:", combobox_values, default="自动选择 (Auto)")

    def create_region_selector(self, default_val=""):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text="搜索范围 (x1,y1,x2,y2) [留空=全屏]:", font=self.font_ui).pack(anchor="w")
        
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, expand=True)
        
        entry = ttk.Entry(input_frame, font=self.font_ui)
        entry.insert(0, str(default_val) if default_val else "")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn = ttk.Button(input_frame, text="🎯 框选", width=8, 
                         command=lambda: self.on_select_region(entry),
                         bootstyle="info-outline")
        btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets['region'] = entry # 注意：这里键名用 'region'，保存时会转为 'cache_box'

    def on_select_region(self, entry_widget):
        self.root.iconify()
        time.sleep(0.3) # 等待最小化动画完成
        
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

    def create_browse_button(self):
        btn = ttk.Button(self.param_frame, text="浏览...", command=self.browse_image, bootstyle="info-outline", padding=(10, 6))
        btn.pack(anchor="w", fill=tk.X, pady=2)

    def create_test_button(self, text, command):
        ttk.Separator(self.param_frame, orient='horizontal').pack(fill='x', pady=(15, 5))
        ttk.Button(self.param_frame, text=text, command=command, bootstyle="info", padding=(10, 6)).pack(anchor="w", fill=tk.X, pady=2)

    def _create_hint_label(self, parent, text, bootstyle="secondary"):
        # [变更] 使用 AutoWrapLabel 替代原有的复杂逻辑
        label_style = f"{bootstyle}.TLabel"
        # 使用 fill=tk.X 以便 Label 知道父容器宽度
        label = AutoWrapLabel(parent, text=text, font=self.font_ui, style=label_style)
        label.pack(anchor="w", pady=5, fill=tk.X)
        return label

    def on_test_find_image_click(self):
        try:
            path = self.param_widgets['path'].get()
            conf = float(self.param_widgets['confidence'].get())
            if not os.path.exists(path): raise FileNotFoundError
            
            # <--- 读取搜索范围
            region_box = None
            if 'region' in self.param_widgets:
                val = self.param_widgets['region'].get().strip()
                # [变更] 使用 gui_utils.parse_region_string
                region_box = parse_region_string(val)

            self.status_var.set("测试中...")
            self.root.iconify()
            # 将 region_box 传给线程
            self.root.after(2000, lambda: self._run_test_thread(self._test_find_image, (path, conf, region_box)))
        except: messagebox.showerror("错误", "参数无效")

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
                val = self.param_widgets['region'].get().strip()
                # [变更] 使用 gui_utils.parse_region_string
                region_box = parse_region_string(val)
            
            if not text: raise ValueError
            self.status_var.set("测试中...")
            self.root.iconify()
            # 将 region_box 传给线程
            self.root.after(2000, lambda: self._run_test_thread(self._test_find_text, (text, lang, engine, region_box)))
        except: messagebox.showerror("错误", "参数无效")

    def _run_test_thread(self, func, args):
        threading.Thread(target=func, args=args, daemon=True).start()

    def _test_find_image(self, path, conf, region_box=None):
        try:
            # <--- 根据区域截图
            if region_box:
                screenshot = ImageGrab.grab(bbox=tuple(region_box))
                offset = (region_box[0], region_box[1])
            else:
                screenshot = ImageGrab.grab()
                offset = (0, 0)
                
            res_val = macro_engine.find_image_cv2(path, conf, screenshot_pil=screenshot, offset=offset)
            loc = res_val[0] if res_val else None
            self.root.after(0, lambda: self._on_test_complete(loc))
        except Exception as e: 
            self.root.after(0, lambda err=e: self._on_test_error(err))

    def _test_find_text(self, text, lang, engine, region_box=None):
        try:
            # <--- 根据区域截图
            if region_box:
                screenshot = ImageGrab.grab(bbox=tuple(region_box))
                offset = (region_box[0], region_box[1])
            else:
                screenshot = ImageGrab.grab()
                offset = (0, 0)
            
            loc = ocr_engine.find_text_location(text, lang, True, screenshot_pil=screenshot, offset=offset, engine=engine)
            self.root.after(0, lambda: self._on_test_complete(loc))
        except Exception as e: 
            self.root.after(0, lambda err=e: self._on_test_error(err))

    def _on_test_complete(self, loc):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        if loc and len(loc) >= 2:
            self.last_test_location = (loc[0], loc[1])
            pyautogui.moveTo(loc[0], loc[1])
            messagebox.showinfo("成功", f"找到于 {self.last_test_location}")
        else:
            messagebox.showwarning("失败", "未找到目标")
        self.update_status_bar_hotkeys()
        self.root.attributes('-topmost', False)

    def _on_test_error(self, e):
        self.root.deiconify()
        messagebox.showerror("错误", str(e))
        self.update_status_bar_hotkeys()

    def browse_image(self):
        f = filedialog.askopenfilename(filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if f: 
            f = os.path.abspath(f) 
            self.param_widgets['path'].delete(0, tk.END); self.param_widgets['path'].insert(0, f)

    def add_or_update_step(self):
        """添加或更新步骤 (已优化：支持插入到选中行下方)"""
        action = MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action: return
        params = {}
        try:
            for k, w in self.param_widgets.items():
                # === 新增：处理 BooleanVar (复选框) ===
                if isinstance(w, tk.BooleanVar):
                    val = w.get()
                    params[k] = val
                    continue
                
                val = w.get()
                
                # 数字校验
                if k in ['x', 'y', 'ms', 'times', 'x_offset', 'y_offset', 'amount', 'max_iterations']:
                    if val and not val.strip().lstrip('-').isdigit():
                        messagebox.showwarning("输入错误", f"参数 '{k}' 必须是整数")
                        return
                
                if action == 'SCROLL' and k in ['x', 'y'] and not val:
                    continue
                
                if not val:
                    if k == 'region': pass # region 允许为空
                    elif k == 'extract_pattern': pass # 正则允许为空
                    elif action in ['ELSE', 'END_IF', 'END_LOOP']: continue
                    elif action == 'SCROLL' and k in ['x', 'y']: continue
                    else: return
                
                # 参数转换
                elif k == 'mode':
                    mode_map = {
                        '固定次数': 'fixed',
                        '直到找到图像': 'until_image',
                        '直到找到文本': 'until_text'
                    }
                    params[k] = mode_map.get(val, 'fixed')
                # [重构] 使用统一的参数映射函数
                elif k in ('lang', 'button', 'engine'):
                    params[k] = self._param_display_to_internal(k, val)
                
                # [变更] 使用通用函数解析 region
                elif k == 'region':
                    if val.strip():
                        coords = parse_region_string(val)
                        if coords: params['cache_box'] = coords
                    continue
                
                # === 新增：处理 extract_pattern，为空时不保存 ===
                elif k == 'extract_pattern':
                    if val and val.strip():
                        params[k] = val.strip()
                    continue

                else:
                    params[k] = val
        except Exception as e: 
            print(f"参数解析错误: {e}")
            return
        
        # [补丁优化] 验证图片文件的有效性
        if action in ('FIND_IMAGE', 'IF_IMAGE_FOUND'):
            img_path = params.get('path', '')
            if img_path:
                if not os.path.exists(img_path):
                    messagebox.showwarning(
                        "文件不存在", 
                        f"图片文件不存在:\n{img_path}\n\n请确认文件路径是否正确。",
                        parent=self.root
                    )
                    return
                if not img_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                    messagebox.showwarning(
                        "文件格式错误",
                        f"仅支持常见图片格式 (PNG, JPG, BMP, GIF)\n\n当前文件: {os.path.basename(img_path)}",
                        parent=self.root
                    )
                    return
        
        # [补丁优化] 验证循环条件图片
        if action == 'LOOP_START':
            mode = params.get('mode', 'fixed')
            if mode == 'until_image':
                img_path = params.get('condition_image', '')
                if img_path:
                    if not os.path.exists(img_path):
                        messagebox.showwarning(
                            "文件不存在",
                            f"循环条件图片不存在:\n{img_path}\n\n请确认文件路径是否正确。",
                            parent=self.root
                        )
                        return
                    if not img_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                        messagebox.showwarning(
                            "文件格式错误",
                            f"仅支持常见图片格式\n\n当前文件: {os.path.basename(img_path)}",
                            parent=self.root
                        )
                        return
        
        step = {"action": action, "params": params}
        
        # 仅在没有手动指定区域时，才询问是否使用测试结果作为缓存
        if action in ('FIND_TEXT', 'FIND_IMAGE', 'IF_TEXT_FOUND', 'IF_IMAGE_FOUND') \
           and not self.editing_index \
           and self.last_test_location \
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
        
        # ============================================================
        # [关键修复] 优先强制处理 LOOP_START 的模式
        # ============================================================
        if step['action'] == 'LOOP_START':
            # 获取保存的模式 (默认 fixed)
            saved_mode = step['params'].get('mode', 'fixed')
            
            # 翻译模式为中文
            mode_map_rev = {
                'fixed': '固定次数',
                'until_image': '直到找到图像',
                'until_text': '直到找到文本'
            }
            display_mode = mode_map_rev.get(saved_mode, '固定次数')
            
            # 1. 强行修改下拉框的值
            if 'mode' in self.param_widgets:
                self.param_widgets['mode'].set(display_mode)
            
            # 2. 强行触发界面刷新 (这一步会让"目标文本"输入框从隐藏变为显示)
            # 必须在填入"沙发"等文字之前完成这一步！
            self.update_loop_params(None)

        # ============================================================
        # 常规参数填充 (此时输入框已经显示出来了，可以安全填值了)
        # ============================================================
        
        # 预处理 Region 显示
        if 'cache_box' in step['params'] and 'region' in self.param_widgets:
            cb = step['params']['cache_box']
            if isinstance(cb, list) and len(cb) == 4:
                self.param_widgets['region'].delete(0, tk.END)
                self.param_widgets['region'].insert(0, f"{cb[0]}, {cb[1]}, {cb[2]}, {cb[3]}")
        
        # 遍历并填充所有参数
        for k, v in step['params'].items():
            # 跳过 mode (前面处理了) 和 cache_box (前面处理了)
            if k in ('mode', 'cache_box', 'region'): continue
            
            if k in self.param_widgets:
                w = self.param_widgets[k]
                
                # [重构] 使用统一的参数映射函数
                if k in ('lang', 'button', 'engine'):
                    display_val = self._param_internal_to_display(k, v)
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
        self.add_step_btn.config(text="✓ 更新步骤", bootstyle="warning")
        self.add_step_btn.grid_configure(columnspan=1)
        self.cancel_edit_btn.grid(row=0, column=1, sticky="nsew", padx=(2,0))
        self.update_listbox_display()

    def cancel_edit_mode(self):
        self.editing_index = None
        self.add_step_btn.config(text="＋ 添加到序列 >>", bootstyle="success")
        self.cancel_edit_btn.grid_remove()
        self.add_step_btn.grid_configure(columnspan=2)
        self.update_listbox_display()

    def update_listbox_display(self):
        """更新 Treeview 显示"""
        for item in self.steps_tree.get_children():
            self.steps_tree.delete(item)
            
        block_stack = []
        for i, step in enumerate(self.steps):
            act = step['action']
            
            # 缩进逻辑
            current_indent_level = max(0, len(block_stack) - (1 if act in ['ELSE', 'END_IF', 'END_LOOP'] else 0))
            indent_str = "    " * current_indent_level
            
            # 参数预览文本
            display_params = step['params'].copy()
            
            cache_str = ""
            if 'cache_box' in display_params:
                box = display_params.pop('cache_box')
                cache_str = f"[区域: {box[0]},{box[1]},{box[2]},{box[3]}] "

            if 'engine' in display_params:
                # <--- 列表显示时也使用完整映射
                display_params['engine'] = self.FULL_OCR_NAME_MAP.get(display_params['engine'], display_params['engine'])
                
            # 格式化参数列字符串
            param_text = f"{cache_str}{display_params}" if display_params else ""
            
            action_label = MacroSchema.ACTION_TRANSLATIONS.get(act, act)
            
            # 插入行 (Values对应: id, action, params)
            item_id = self.steps_tree.insert("", "end", values=(
                i + 1,
                f"{indent_str}{action_label}",
                param_text
            ))
            
            # 如果是编辑行，高亮显示 (Tag: editing)
            if i == self.editing_index:
                self.steps_tree.item(item_id, tags=('editing',))
                # 确保滚动可见
                self.steps_tree.see(item_id)
                # 保持选中状态 (可选)
                self.steps_tree.selection_set(item_id)

            if act.startswith('IF_') or act == 'LOOP_START':
                block_stack.append(act)
            elif act in ['END_IF', 'END_LOOP'] and block_stack:
                block_stack.pop()

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

    def start_hotkey_listener(self):
        """切换回 Listener 模式"""
        if self.hotkey_listener:
            try:
                self.hotkey_listener.stop()
            except:
                pass
        threading.Thread(target=self._hotkey_listener_thread, daemon=True).start()

    def _hotkey_listener_thread(self):
        """快捷键监听线程"""
        try:
            self.hotkey_listener = keyboard.Listener(
                on_press=self.on_hotkey_press, 
                on_release=self.on_hotkey_release
            )
            self.hotkey_listener.start()
            self.hotkey_listener.join()
        except Exception as e: 
            msg = f"热键监听器启动失败: {e}\n\n快捷键将无法工作。请尝试重启程序。"
            self.root.after(0, messagebox.showerror, "严重错误", msg)

    def _get_key_name_from_key(self, key):
        """辅助函数：优先使用 vk 获取按键名称"""
        try:
            if hasattr(key, 'vk') and key.vk in VK_TO_PYNPUT:
                return VK_TO_PYNPUT[key.vk]
            if hasattr(key, 'name') and key.name:
                return key.name.lower()
            if hasattr(key, 'char') and key.char:
                return key.char.lower()
            return str(key).lower()
        except:
            return None

    def on_hotkey_press(self, key):
        """ 按键按下事件"""
        try:
            key_name = self._get_key_name_from_key(key)
            if not key_name: return
                
            if key_name in ['ctrl_l', 'ctrl_r']: key_name = 'ctrl'
            elif key_name in ['alt_l', 'alt_r', 'alt_gr']: key_name = 'alt'
            elif key_name in ['shift_l', 'shift_r']: key_name = 'shift'
            elif key_name in ['cmd_l', 'cmd_r', 'cmd']: key_name = 'cmd'
            
            if key_name not in self.held_keys:
                self.held_keys.add(key_name)
                
                run_mods, run_key = self._parse_hotkey(self.hotkey_run_str.get())
                if key_name == run_key and run_mods.issubset(self.held_keys):
                    self.root.after(0, self.safe_run_macro)
                
                stop_mods, stop_key = self._parse_hotkey(self.hotkey_stop_str.get())
                if key_name == stop_key and stop_mods.issubset(self.held_keys):
                    self.root.after(0, self.safe_stop_macro)
        except (AttributeError, KeyError) as e:
            print(f"[Hotkey] 按键解析错误: {e}")
        except Exception as e:
            print(f"[Hotkey] 未知错误 (press): {e}")

    def on_hotkey_release(self, key):
        """按键释放事件"""
        try:
            key_name = self._get_key_name_from_key(key)
            if not key_name: return
                
            if key_name in ['ctrl_l', 'ctrl_r']: key_name = 'ctrl'
            elif key_name in ['alt_l', 'alt_r', 'alt_gr']: key_name = 'alt'
            elif key_name in ['shift_l', 'shift_r']: key_name = 'shift'
            elif key_name in ['cmd_l', 'cmd_r', 'cmd']: key_name = 'cmd'
            
            if key_name in self.held_keys:
                self.held_keys.remove(key_name)
        except (AttributeError, KeyError) as e:
            print(f"[Hotkey] 按键解析错误: {e}")
        except Exception as e:
            print(f"[Hotkey] 未知错误 (release): {e}")

    @functools.lru_cache(maxsize=16)
    def _parse_hotkey(self, hotkey_str):
        """ 解析快捷键字符串（小写），返回 (modifiers, key)"""
        parts = [p.strip() for p in hotkey_str.lower().split('+')]
        key = parts[-1]
        modifiers = set(parts[:-1])
        return modifiers, key

    def restart_hotkey_listener(self):
        """停止并重新启动监听器"""
        if self.hotkey_listener:
            self.hotkey_listener.stop()
        self.start_hotkey_listener()

    def safe_run_macro(self):
        if not self.is_macro_running and self.editing_index is None:
            self.root.after(0, self.run_macro, True)
        
    def safe_stop_macro(self):
        if self.is_macro_running:
            self.root.after(0, self.status_var.set, "正在停止...")
            if self.current_run_context: 
                self.current_run_context['stop_requested'] = True
        
    def run_macro(self, hotkey=False):
        if self.is_macro_running or not self.steps: return
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        
        if not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno("运行", f"是否立即开始？(按 {stop_display} 停止)"): return
            
        self.loop_status_var.set("") 
        
        # [核心修复] 暴力清空之前的状态队列，防止积压
        while not self.status_queue.empty():
            try: self.status_queue.get_nowait()
            except queue.Empty: break
            
        self.run_btn.config(state="disabled")
        self.status_var.set(f"宏正在运行... [{stop_display}] 停止")
        if not self.dont_minimize_var.get(): self.root.iconify()
        else: self.root.attributes('-topmost', True) 
        self.root.after(1500, self._start_macro_thread)

    def _start_macro_thread(self):
        self.is_macro_running = True
        self.current_run_context = {
            'stop_requested': False,
            'stop_key_str': self.hotkey_stop_str.get()
        }
        threading.Thread(target=self._run, args=(self.steps.copy(),), daemon=True).start()
        
    def _run(self, steps):
        try:
            macro_engine.execute_steps(steps, run_context=self.current_run_context, status_callback=self.update_loop_status)
        except Exception as e: self.root.after(0, lambda err=e: messagebox.showerror("错误", str(err)))
        finally: self.root.after(0, self._on_macro_complete)

    def _on_macro_complete(self):
        self.is_macro_running = False
        self.current_run_context = None
        self.root.deiconify()
        self.root.attributes('-topmost', False)
        self.run_btn.config(state="normal")
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
            
            if text: self.loop_status_var.set(text)
        except queue.Empty:
            pass
        except Exception as e:
            print(f"[StatusQueue] 错误: {e}")
            import traceback
            traceback.print_exc()  # [补丁优化] 记录完整堆栈
            
        self.root.after(interval, self._check_status_queue)

    def new_macro(self):
        if self.steps:
            if not messagebox.askyesno("新建", "清空当前宏？"): return
        self.steps = []
        self.editing_index = None
        self.last_test_location = None
        self.cancel_edit_mode()
        self.update_listbox_display()
        self.status_var.set("已新建空白宏。")

    def load_macro(self):
        f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if f: self._load_file(f)

    def save_macro(self):
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f:
            try:
                with open(f, 'w', encoding='utf-8') as file: json.dump(self.steps, file, indent=4)
                messagebox.showinfo("成功", "宏已保存！")
                self.add_to_recent_files(f)
            except Exception as e: messagebox.showerror("失败", str(e))

    def _load_file(self, f):
        if not os.path.exists(f):
            messagebox.showerror("失败", "文件不存在")
            if f in self.recent_files: self.recent_files.remove(f); self.save_app_settings(); self.update_recent_files_menu()
            return
        try:
            self.cancel_edit_mode()
            with open(f, 'r', encoding='utf-8') as file: 
                data = json.load(file)
            
            # 验证JSON数据结构
            if not self._validate_macro_data(data):
                messagebox.showerror(
                    "加载失败", 
                    f"文件格式无效或损坏:\n{os.path.basename(f)}\n\n"
                    "可能原因:\n"
                    "• 不是有效的宏文件\n"
                    "• 文件被手动编辑导致格式错误\n"
                    "• 文件损坏"
                )
                return
            
            self.steps = data
            self.update_listbox_display()
            self.status_var.set(f"已加载: {os.path.basename(f)}")
            self.add_to_recent_files(f)
        except json.JSONDecodeError as e:
            messagebox.showerror(
                "JSON解析错误", 
                f"文件不是有效的JSON格式:\n{os.path.basename(f)}\n\n"
                f"错误详情: {str(e)}\n\n"
                "请检查文件是否被意外修改。"
            )
        except Exception as e: 
            messagebox.showerror("加载失败", f"无法加载文件:\n{str(e)}")
    
    def _validate_macro_data(self, data):
        """
        [补丁新增] 验证宏数据结构是否有效
        
        Args:
            data: 从JSON加载的数据
            
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
            
            # 验证 action 是否是已知的动作类型 (仅警告，不阻止)
            if step['action'] not in MacroSchema.ACTION_TRANSLATIONS:
                print(f"[警告] 步骤 {i+1} 包含未知的动作类型: {step['action']}")
                # 不返回 False，允许加载未知动作类型（向前兼容）
        
        return True

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
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                    self.recent_files = d.get('recent_files', [])
                    self.current_theme.set(d.get('theme', 'litera'))
                    self.hotkey_run_str.set(d.get('hotkey_run', DEFAULT_HOTKEY_RUN))
                    self.hotkey_stop_str.set(d.get('hotkey_stop', DEFAULT_HOTKEY_STOP))
        except:
            pass
        self.root.style.theme_use(self.current_theme.get())

    def save_app_settings(self):
        """保存应用设置"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'recent_files': self.recent_files,
                    'theme': self.current_theme.get(),
                    'hotkey_run': self.hotkey_run_str.get(),
                    'hotkey_stop': self.hotkey_stop_str.get()
                }, f, indent=2)
        except:
            pass

    def change_theme(self):
        self.root.style.theme_use(self.current_theme.get())
        self.root.style.configure(".", font=self.font_ui)
        self.save_app_settings()
        
    def check_hotkey_conflicts(self, show_success=True):
        if not HOTKEY_CHECK_AVAILABLE:
            print("[警告] 跳过快捷键冲突检测 (pywin32 未安装或非 Windows 系统)")
            return True 

        conflicts = []
        
        if not self._test_register_hotkey(self.hotkey_run_str.get(), 1):
            conflicts.append(f"运行快捷键 '{capitalize_hotkey_str(self.hotkey_run_str.get())}'")
        
        if not self._test_register_hotkey(self.hotkey_stop_str.get(), 2):
            conflicts.append(f"停止快捷键 '{capitalize_hotkey_str(self.hotkey_stop_str.get())}'")
            
        if conflicts:
            msg = "检测到快捷键冲突：\n\n" + "\n".join(conflicts) + "\n\n可能已被其他程序 (如 NVIDIA, QQ, 微信) 占用。\n请在设置中修改快捷键，否则热键可能无法工作。"
            self.root.after(0, messagebox.showwarning, "快捷键冲突", msg)
            return False
        elif show_success:
            pass 
        return True

    def _parse_hotkey_string_to_win32(self, hotkey_str):
        parts = hotkey_str.lower().split('+')
        modifiers = 0
        vk_key = None
        
        for part in parts:
            part = part.strip()
            if part in PYNPUT_MOD_TO_WIN_MOD:
                modifiers |= PYNPUT_MOD_TO_WIN_MOD[part]
            elif part in PYNPUT_TO_VK:
                vk_key = PYNPUT_TO_VK[part]
                
        return modifiers, vk_key

    def _test_register_hotkey(self, hotkey_str, hotkey_id):
        if not hotkey_str: return True
        try:
            modifiers, vk = self._parse_hotkey_string_to_win32(hotkey_str)
            if vk is None:
                print(f"无法解析快捷键进行冲突检测: {hotkey_str}")
                return True 
                
            hwnd = None 
            if ctypes.windll.user32.RegisterHotKey(hwnd, hotkey_id, modifiers, vk) == 0:
                return False
            else:
                ctypes.windll.user32.UnregisterHotKey(hwnd, hotkey_id)
                return True
        except Exception as e:
            print(f"快捷键检测时发生错误: {e}")
            return True


if __name__ == "__main__":
    pyautogui.FAILSAFE = False
    try:
        theme = "litera"
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f: theme = json.load(f).get('theme', 'litera')
    except: pass
    main_window = tb.Window(themename=theme)
    app = MacroApp(main_window)
    main_window.mainloop()
