# -*- coding: utf-8 -*-
# MacroAssistant.py
# 描述: 自动化宏的 GUI 界面
# 版本: 1.53.2
# 变更: (修复#B) 优化 OCR 引擎下拉框逻辑，正确处理不可用引擎的加载和保存。

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
from PIL import ImageGrab 
import functools

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
APP_VERSION = "1.53.2" # <--- 版本更新
APP_TITLE = f"宏助手 (Macro Assistant) V{APP_VERSION}"
APP_ICON = "app_icon.ico" 
CONFIG_FILE = "macro_settings.json"
MAX_RECENT_FILES = 5

DEFAULT_HOTKEY_RUN = "ctrl+f10"
DEFAULT_HOTKEY_STOP = "ctrl+f11"
# =================================================================

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
except ImportError:
    messagebox.showerror("导入错误", "未找到 'core_engine.py' 或 'ocr_engine.py'。\n请确保它们与 'MacroAssistant.py' 位于同一目录。")
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

class HotkeyEntry(ttk.Entry):
    """一个用于捕获和显示 pynput 快捷键的输入框"""
    def __init__(self, master=None, **kwargs):
        self.string_var = kwargs.pop("textvariable", None)
        super().__init__(master, **kwargs)
        
        self.current_keys = set()
        self.bind("<KeyPress>", self._on_key_press)
        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self["font"] = ("Consolas", 10)
        self.config(justify="center")
        
    def set_hotkey(self, hotkey_str):
        """设置快捷键 (存小写, 显大写)"""
        display_str = capitalize_hotkey_str(hotkey_str) if hotkey_str else "点击 [捕获] 录制"
        self.configure(state="normal")
        self.delete(0, tk.END)
        self.insert(0, display_str)
        self.configure(state="readonly")
        if self.string_var:
            self.string_var.set(hotkey_str)

    def _on_focus_in(self, event):
        self.configure(state="normal")
        self.delete(0, tk.END)
        self.insert(0, "录制中...")
        self.configure(state="readonly")
        
    def _on_focus_out(self, event):
        if not self.current_keys and self.string_var:
             self.set_hotkey(self.string_var.get())
        self.current_keys.clear()

    def _on_key_press(self, event):
        self.configure(state="normal")
        self.delete(0, tk.END)
        key_name = self._get_key_name(event)
        if key_name:
            self.current_keys.add(key_name)
            self._format_hotkey_string(update_var=False) 
        self.configure(state="readonly")
        return "break"

    def _on_key_release(self, event):
        key_name = self._get_key_name(event)
        if key_name and key_name not in {'ctrl', 'alt', 'shift', 'cmd'}:
            self._format_hotkey_string(update_var=True)
            self.current_keys.clear()
            self.master.focus()
        return "break"

    def _format_hotkey_string(self, update_var=False):
        """手动构建快捷键字符串"""
        if not self.current_keys:
            self.configure(state="normal")
            self.delete(0, tk.END)
            self.insert(0, "录制中...")
            self.configure(state="readonly")
            return

        mods = []
        key = None
        
        if 'ctrl' in self.current_keys: mods.append('ctrl')
        if 'alt' in self.current_keys: mods.append('alt')
        if 'shift' in self.current_keys: mods.append('shift')
        if 'cmd' in self.current_keys: mods.append('cmd')
        
        for k in self.current_keys:
            if k not in {'ctrl', 'alt', 'shift', 'cmd'}:
                key = k
                break
        
        if key:
            hotkey_str_value = "+".join(mods + [key])
        else:
            hotkey_str_value = "+".join(mods)
        
        hotkey_str_display = capitalize_hotkey_str(hotkey_str_value)

        self.configure(state="normal")
        self.delete(0, tk.END)
        self.insert(0, hotkey_str_display)
        self.configure(state="readonly")
        
        if update_var and key and self.string_var:
            self.string_var.set(hotkey_str_value)

    def _get_key_name(self, event):
        name = event.keysym.lower()
        if "control" in name: return "ctrl"
        if "alt" in name: return "alt"
        if "shift" in name: return "shift"
        if "win" in name or "super" in name: return "cmd"
        if name.startswith("f") and name[1:].isdigit(): return name
        if len(name) == 1 and ('a' <= name <= 'z' or '0' <= name <= '9'):
            return name
            
        special_keys_map = {
            'return': 'enter', 'space': 'space', 'tab': 'tab',
            'capital': 'caps_lock', 'escape': 'esc',
            'prior': 'page_up', 'next': 'page_down', 'end': 'end', 'home': 'home',
            'left': 'left', 'up': 'up', 'right': 'right', 'down': 'down',
            'insert': 'insert', 'delete': 'delete', 'backspace': 'backspace'
        }
        return special_keys_map.get(name, None)


class HotkeySettingsDialog:
    """快捷键设置对话框"""
    def __init__(self, parent, current_run, current_stop):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("快捷键设置")
        self.dialog.geometry("450x480") 
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="⌨️ 自定义快捷键", 
                  font=("Microsoft YaHei UI", 12, "bold")).pack(pady=(0, 15))
        
        run_frame = ttk.Labelframe(main_frame, text="运行/继续 快捷键", padding=15)
        run_frame.pack(fill=tk.X, pady=(0, 15))
        run_inner = ttk.Frame(run_frame)
        run_inner.pack(fill=tk.X)
        run_inner.columnconfigure(0, weight=1)

        self.run_var = tk.StringVar(value=current_run)
        self.run_display = HotkeyEntry(run_inner, textvariable=self.run_var)
        self.run_display.set_hotkey(current_run)
        self.run_display.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        
        self.run_capture_btn = ttk.Button(run_inner, text="🎯 录制", 
                                          command=self.run_display.focus_set,
                                          bootstyle="info", width=12)
        self.run_capture_btn.grid(row=0, column=1, ipady=3)
        
        stop_frame = ttk.Labelframe(main_frame, text="停止宏快捷键", padding=15)
        stop_frame.pack(fill=tk.X, pady=(0, 15))
        stop_inner = ttk.Frame(stop_frame)
        stop_inner.pack(fill=tk.X)
        stop_inner.columnconfigure(0, weight=1)
        
        self.stop_var = tk.StringVar(value=current_stop)
        self.stop_display = HotkeyEntry(stop_inner, textvariable=self.stop_var)
        self.stop_display.set_hotkey(current_stop)
        self.stop_display.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        
        self.stop_capture_btn = ttk.Button(stop_inner, text="🎯 录制", 
                                           command=self.stop_display.focus_set,
                                           bootstyle="info", width=12)
        self.stop_capture_btn.grid(row=0, column=1, ipady=3)
        
        hint_frame = ttk.Frame(main_frame)
        hint_frame.pack(fill=tk.X, pady=(20, 20))
        
        hint_text = "💡 支持: Ctrl, Alt, Shift, F1-F12, A-Z, 0-9等"
        ttk.Label(hint_frame, text=hint_text, font=("Microsoft YaHei UI", 9), 
                 foreground="#666", justify=tk.LEFT).pack()
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        
        ttk.Button(btn_frame, text="✕ 取消", command=self.cancel, 
                  bootstyle="secondary", padding=(10, 10)).grid(row=0, column=0, sticky="ew", padx=(5, 0))
        ttk.Button(btn_frame, text="🔄 恢复默认", command=self.reset_default, 
                  bootstyle="warning-outline", padding=(10, 10)).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(btn_frame, text="✓ 保存", command=self.save, 
                  bootstyle="success", padding=(10, 10)).grid(row=0, column=2, sticky="ew", padx=(0, 5))
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
    def reset_default(self):
        self.run_var.set(DEFAULT_HOTKEY_RUN)
        self.run_display.set_hotkey(DEFAULT_HOTKEY_RUN)
        self.stop_var.set(DEFAULT_HOTKEY_STOP)
        self.stop_display.set_hotkey(DEFAULT_HOTKEY_STOP)
        
    def save(self):
        run_hotkey = self.run_var.get().strip().lower()
        stop_hotkey = self.stop_var.get().strip().lower()
        
        if not run_hotkey or not stop_hotkey or "录制" in run_hotkey or "录制" in stop_hotkey:
            messagebox.showerror("错误", "快捷键不能为空", parent=self.dialog)
            return
            
        if run_hotkey == stop_hotkey:
            messagebox.showerror("错误", "运行和停止快捷键不能相同", parent=self.dialog)
            return
        
        if not self._validate_hotkey(run_hotkey):
            messagebox.showerror("错误", f"运行快捷键格式无效: {run_hotkey}", parent=self.dialog)
            return
            
        if not self._validate_hotkey(stop_hotkey):
            messagebox.showerror("错误", f"停止快捷键格式无效: {stop_hotkey}", parent=self.dialog)
            return
        
        self.result = (run_hotkey, stop_hotkey)
        self.dialog.destroy()
        
    def _validate_hotkey(self, hotkey):
        parts = hotkey.split('+')
        if len(parts) == 0: return False

        if len(parts) == 1:
            part = parts[0]
            if part.startswith('f') and part[1:].isdigit():
                 return int(part[1:]) in range(1, 13)
            return False
        
        modifiers = {'ctrl', 'alt', 'shift', 'cmd'}
        valid_keys = set('abcdefghijklmnopqrstuvwxyz0123456789')
        valid_keys.update([f'f{i}' for i in range(1, 13)])
        valid_keys.update(['space', 'enter', 'tab', 'esc', 'backspace', 'delete'])
        
        for i, part in enumerate(parts):
            part = part.strip()
            if i < len(parts) - 1:
                if part not in modifiers:
                    return False
            else:
                if part not in valid_keys:
                    return False
        return True
        
    def cancel(self):
        self.result = None
        self.dialog.destroy()


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("960x700")
        
        self.font_ui = ("Microsoft YaHei UI", 10)
        self.font_code = ("Consolas", 10)
        
        self.root.style.configure(".", font=self.font_ui)
        
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
        
        self.mouse_tracker_job = None
        self.mouse_pos_var = tb.StringVar()
        
        self.dynamic_wrap_labels = []
        
        # <--- 重构 OCR 引擎映射
        # 1. 创建一个包含 *所有* 可能引擎的完整映射 (用于显示和解析)
        self.FULL_OCR_NAME_MAP = {
            'auto': '自动选择 (Auto)',
            'rapidocr': 'RapidOCR (推荐)',
            'tesseract': 'Tesseract OCR',
            'winocr': 'Windows 10/11 OCR',
            'none': '无可用OCR引擎'
        }
        # 2. 创建反向映射 (用于保存)
        self.FULL_OCR_KEY_MAP = {name: key for key, name in self.FULL_OCR_NAME_MAP.items()}
        
        # 3. 获取当前环境 *实际可用* 的引擎
        self.available_ocr_engines = ocr_engine.get_available_engines()
        self.available_ocr_keys = [e[0] for e in self.available_ocr_engines]
        
        if 'none' in self.available_ocr_keys:
             print("[警告] 未找到任何可用的OCR引擎 (RapidOCR, Tesseract, WinOCR)。")


        self.menu_bar = tk.Menu(root)
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

        status_bar_frame = ttk.Frame(root, bootstyle="primary")
        status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var = tk.StringVar()
        self.status_label_left = ttk.Label(status_bar_frame, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W, padding=5, bootstyle="primary-inverse", font=self.font_ui)
        self.status_label_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.loop_status_var = tk.StringVar()
        self.loop_status_label_right = ttk.Label(status_bar_frame, textvariable=self.loop_status_var, relief=tk.FLAT, anchor=tk.E, padding=(0, 5, 5, 5), bootstyle="primary-inverse", font=self.font_ui)
        self.loop_status_label_right.pack(side=tk.RIGHT)

        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        list_frame = ttk.Frame(main_frame, padding=10)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(list_frame, text="宏步骤序列:", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")

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
        
        self.steps_listbox = tk.Listbox(list_frame, width=55, font=self.font_code)
        self.steps_listbox.pack(fill=tk.BOTH, expand=True, pady=5) 

        add_frame = ttk.Labelframe(main_frame, text="添加新步骤", padding=10)
        add_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10, expand=True)
        right_bottom_frame = ttk.Frame(add_frame)
        right_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        right_bottom_frame.columnconfigure(0, weight=2); right_bottom_frame.columnconfigure(1, weight=1) 
        
        self.add_step_btn = ttk.Button(right_bottom_frame, text="＋ 添加到序列 >>", command=self.add_or_update_step, bootstyle="success", padding=(12, 8))
        self.add_step_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2), columnspan=2)
        self.cancel_edit_btn = ttk.Button(right_bottom_frame, text="✕ 取消修改", command=self.cancel_edit_mode, bootstyle="secondary", padding=(10, 6))
        
        ttk.Label(add_frame, text="选择动作:").pack(anchor="w")
        self.action_type = ttk.Combobox(add_frame, state="readonly", width=30, font=self.font_ui, height=16)
        self.action_type['values'] = list(MacroSchema.ACTION_TRANSLATIONS.values())
        self.action_type.current(0)
        self.action_type.pack(anchor="w", fill=tk.X, pady=5)
        self.action_type.bind("<<ComboboxSelected>>", self.update_param_fields)
        self.param_frame = ttk.Frame(add_frame)
        self.param_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.param_frame.bind("<Configure>", self._on_param_frame_configure)
        
        self.param_widgets = {}
        self.update_param_fields(None)
        
        self.load_app_settings()
        self.update_recent_files_menu()
        self.update_status_bar_hotkeys() 
        
        self.root.after(500, self.check_hotkey_conflicts)
        self.start_hotkey_listener() 
        
        self.root.after(2000, lambda: threading.Thread(target=ocr_engine.preload_engines, daemon=True).start())
        self._check_status_queue()

    def update_status_bar_hotkeys(self):
        """更新状态栏和运行按钮上的快捷键提示"""
        run_display = capitalize_hotkey_str(self.hotkey_run_str.get())
        stop_display = capitalize_hotkey_str(self.hotkey_stop_str.get())
        self.status_var.set(f"准备就绪...  |  [{run_display}] 启动宏  |  [{stop_display}] 停止宏")
        self.run_btn.config(text=f"▶ 运行宏 ({run_display})")

    def open_hotkey_settings(self):
        """打开快捷键设置对话框"""
        dialog = HotkeySettingsDialog(self.root, self.hotkey_run_str.get(), self.hotkey_stop_str.get())
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
        
        if self.mouse_tracker_job:
            try:
                self.root.after_cancel(self.mouse_tracker_job)
            except tk.TclError:
                pass
            self.mouse_tracker_job = None
            
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
        
        if self.mouse_tracker_job:
            try:
                self.root.after_cancel(self.mouse_tracker_job)
            except tk.TclError:
                pass # 已经取消
            finally:
                self.mouse_tracker_job = None # 确保被清除
        self.mouse_pos_var.set("")
        
        self.dynamic_wrap_labels.clear()
        
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
                    bootstyle="danger") # 使用红色提示
                # 自动切换回一个安全选项
                self.action_type.set(MacroSchema.ACTION_TRANSLATIONS['FIND_IMAGE'])
                # 递归调用以刷新界面
                self.update_param_fields(None)
                return
        
        if action_key == 'FIND_IMAGE':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_param_entry("confidence", "置信度(0.1-1.0):", "0.8")
            self._create_hint_label(self.param_frame, "* 提示：如果识别失败，请尝试调低置信度 (如 0.7)")
            self.create_browse_button()
            self.create_test_button("🧪 测试查找图像", self.on_test_find_image_click)
            
        elif action_key == 'FIND_TEXT':
            self.create_param_entry("text", "查找的文本:", "确定")
            self.create_param_combobox("lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
            # <--- 动态构建引擎下拉框
            self.create_ocr_engine_combobox()
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
            self._create_hint_label(self.param_frame, "* 此功能使用剪贴板 (Ctrl+V)，以支持中文及复杂文本输入。")
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
            self._start_mouse_tracker()
            
        elif action_key == 'IF_IMAGE_FOUND':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_param_entry("confidence", "置信度:", "0.8")
            self.create_browse_button()
            self.create_test_button("🧪 测试 IF 图像", self.on_test_find_image_click)
            
        elif action_key == 'IF_TEXT_FOUND':
            self.create_param_entry("text", "查找文本:", "确定")
            self.create_param_combobox("lang", "语言:", list(MacroSchema.LANG_OPTIONS.keys()))
            # <--- 动态构建引擎下拉框
            self.create_ocr_engine_combobox()
            self.create_test_button("🧪 测试 IF 文本", self.on_test_find_text_click)
            
        elif action_key == 'LOOP_START':
            self.create_param_entry("times", "循环次数:", "10")
        elif action_key == 'ELSE':
            self._create_hint_label(self.param_frame, "* 提示: 'ELSE' 必须与 'IF' 配合使用。它将执行 'IF' 条件不满足时的逻辑。")
        elif action_key == 'END_IF':
            self._create_hint_label(self.param_frame, "* 提示: 'END_IF' 必须与 'IF' 配合使用。它标志着 'IF' 或 'ELSE' 逻辑块的结束。")
        elif action_key == 'END_LOOP':
            self._create_hint_label(self.param_frame, "* 提示: 'END_LOOP' 必须与 'LOOP_START' 配合使用。它标志着循环体的结束。")


    def create_param_entry(self, key, label_text, default_value):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        entry = ttk.Entry(frame, width=30, font=self.font_ui)
        entry.insert(0, default_value)
        entry.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = entry
        
    def create_param_combobox(self, key, label_text, values, default=None):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=28, font=self.font_ui)
        if default and default in values:
            combo.set(default)
        else:
            combo.current(0)
        combo.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = combo
    
    # <--- 专用函数创建引擎下拉框
    def create_ocr_engine_combobox(self):
        """动态构建 OCR 引擎下拉框，标记不可用"""
        combobox_values = ['自动选择 (Auto)']
        # 遍历 *所有* 引擎，而不仅仅是可用的引擎
        for key, name in self.FULL_OCR_NAME_MAP.items():
            if key in ('auto', 'none'): continue
            
            if key in self.available_ocr_keys:
                combobox_values.append(name) # "RapidOCR (推荐)"
            else:
                combobox_values.append(f"{name} (不可用)") # "RapidOCR (推荐) (不可用)"
                
        self.create_param_combobox("engine", "OCR 引擎:", combobox_values, default="自动选择 (Auto)")

    def create_browse_button(self):
        btn = ttk.Button(self.param_frame, text="浏览...", command=self.browse_image, bootstyle="info-outline", padding=(10, 6))
        btn.pack(anchor="w", fill=tk.X, pady=2)

    def create_test_button(self, text, command):
        ttk.Separator(self.param_frame, orient='horizontal').pack(fill='x', pady=(15, 5))
        ttk.Button(self.param_frame, text=text, command=command, bootstyle="info", padding=(10, 6)).pack(anchor="w", fill=tk.X, pady=2)

    def _create_hint_label(self, parent, text, bootstyle="secondary"):
        parent_width = parent.winfo_width()
        initial_wrap = max(250, parent_width - 15) 
        
        label_style = f"{bootstyle}.TLabel"
        label = ttk.Label(parent, text=text, wraplength=initial_wrap, font=self.font_ui, style=label_style)
        
        # 兼容旧的 bootstyle (如果 secondary.TLabel 不存在)
        try:
            label.pack(anchor="w", pady=5)
        except tk.TclError:
            label.config(style="TLabel", foreground='gray') # 回退
            label.pack(anchor="w", pady=5)
            
        self.dynamic_wrap_labels.append(label)
        return label

    def _on_param_frame_configure(self, event):
        width = event.width - 15 
        if width > 0:
            for label in self.dynamic_wrap_labels:
                try:
                    label.config(wraplength=width)
                except tk.TclError:
                    pass

    def _start_mouse_tracker(self):
        if not self.is_app_running: return
        self._update_mouse_pos()
        self.mouse_tracker_job = self.root.after(100, self._start_mouse_tracker)

    def _update_mouse_pos(self):
        try:
            x, y = pyautogui.position()
            self.mouse_pos_var.set(f"X: {x}, Y: {y}")
        except Exception:
            self.mouse_pos_var.set("无法获取坐标")

    def on_test_find_image_click(self):
        try:
            path = self.param_widgets['path'].get()
            conf = float(self.param_widgets['confidence'].get())
            if not os.path.exists(path): raise FileNotFoundError
            self.status_var.set("测试中...")
            self.root.iconify()
            self.root.after(2000, lambda: self._run_test_thread(self._test_find_image, (path, conf)))
        except: messagebox.showerror("错误", "参数无效")

    def on_test_find_text_click(self):
        try:
            text = self.param_widgets['text'].get()
            lang = MacroSchema.LANG_OPTIONS.get(self.param_widgets['lang'].get(), 'eng')
            
            # <--- 解析引擎名称
            engine_name = self.param_widgets['engine'].get()
            if engine_name.endswith(" (不可用)"):
                engine_name = engine_name.replace(" (不可用)", "")
            engine = self.FULL_OCR_KEY_MAP.get(engine_name, 'auto')
            
            if not text: raise ValueError
            self.status_var.set("测试中...")
            self.root.iconify()
            self.root.after(2000, lambda: self._run_test_thread(self._test_find_text, (text, lang, engine)))
        except: messagebox.showerror("错误", "参数无效")

    def _run_test_thread(self, func, args):
        threading.Thread(target=func, args=args, daemon=True).start()

    def _test_find_image(self, path, conf):
        try:
            screenshot = ImageGrab.grab()
            res_val = macro_engine.find_image_cv2(path, conf, screenshot_pil=screenshot)
            loc = res_val[0] if res_val else None
            self.root.after(0, lambda: self._on_test_complete(loc))
        except Exception as e: 
            self.root.after(0, lambda err=e: self._on_test_error(err))

    def _test_find_text(self, text, lang, engine):
        try:
            screenshot = ImageGrab.grab()
            loc = ocr_engine.find_text_location(text, lang, True, screenshot_pil=screenshot, offset=(0,0), engine=engine)
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
            f = os.path.abspath(f) # <--- 建议的修复 (路径规范化)
            self.param_widgets['path'].delete(0, tk.END); self.param_widgets['path'].insert(0, f)

    def add_or_update_step(self):
        action = MacroSchema.ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action: return
        params = {}
        try:
            for k, w in self.param_widgets.items():
                val = w.get()
                
                if action == 'SCROLL' and k in ['x', 'y'] and not val:
                    continue
                
                if not val:
                    if action in ['ELSE', 'END_IF', 'END_LOOP']:
                        continue
                    if action == 'SCROLL' and k in ['x', 'y']:
                        continue
                    
                    return
                
                if k == 'lang':
                    params[k] = MacroSchema.LANG_OPTIONS.get(val, val)
                elif k == 'button':
                    params[k] = MacroSchema.CLICK_OPTIONS.get(val, val)
                elif k == 'engine':
                    if val.endswith(" (不可用)"):
                        val = val.replace(" (不可用)", "")
                    params[k] = self.FULL_OCR_KEY_MAP.get(val, 'auto')
                else:
                    params[k] = val
        except: return
        
        step = {"action": action, "params": params}
        if action in ('FIND_TEXT', 'FIND_IMAGE', 'IF_TEXT_FOUND', 'IF_IMAGE_FOUND') and not self.editing_index and self.last_test_location:
            if messagebox.askyesno("缓存", "使用测试坐标作为缓存？"):
                step["params"]["cache_box"] = [self.last_test_location[0], self.last_test_location[1], self.last_test_location[0]+1, self.last_test_location[1]+1]

        if self.editing_index is not None: self.steps[self.editing_index] = step; self.cancel_edit_mode()
        else: self.steps.append(step); self.update_listbox_display(); self.steps_listbox.see(tk.END)
        self.last_test_location = None

    def load_step_for_edit(self):
        sel = self.steps_listbox.curselection()
        if not sel: return
        idx = sel[0]
        step = self.steps[idx]
        self.action_type.set(MacroSchema.ACTION_TRANSLATIONS.get(step['action']))
        self.update_param_fields(None)
        
        for k, v in step['params'].items():
            if k in self.param_widgets:
                
                if k=='lang':
                    val = MacroSchema.LANG_VALUES_TO_NAME.get(v, v)
                elif k=='button':
                    val = MacroSchema.CLICK_VALUES_TO_NAME.get(v, v)
                # <--- 加载时反向映射引擎名称
                elif k=='engine':
                    # 检查保存的 key (v) 是否在 *当前可用* 列表中
                    if v not in self.available_ocr_keys and v != 'auto':
                        # 不可用，显示 (不可用)
                        name = self.FULL_OCR_NAME_MAP.get(v, v) # 获取友好名称
                        val = f"{name} (不可用)"
                    else:
                        # 可用，或为 auto
                        val = self.FULL_OCR_NAME_MAP.get(v, "自动选择 (Auto)")
                else:
                    val = v
                
                w = self.param_widgets[k]
                if isinstance(w, ttk.Combobox): w.set(val)
                else: w.delete(0, tk.END); w.insert(0, str(val))
        
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
        display_texts = []
        block_stack = []
        for i, step in enumerate(self.steps):
            act = step['action']
            current_indent_level = max(0, len(block_stack) - (1 if act in ['ELSE', 'END_IF', 'END_LOOP'] else 0))
            indent_str = "    " * current_indent_level
            
            display_params = step['params'].copy()
            cache_str = ""
            if 'cache_box' in display_params:
                 box = display_params.pop('cache_box')
                 cache_str = f" [Cache: {box[0]}, {box[1]}]"

            if 'engine' in display_params:
                # <--- 列表显示时也使用完整映射
                display_params['engine'] = self.FULL_OCR_NAME_MAP.get(display_params['engine'], display_params['engine'])
                
            prefix = "[编辑] -> " if i == self.editing_index else f"步骤 {i+1}: "
            
            action_label = MacroSchema.ACTION_TRANSLATIONS.get(act, act)
            
            param_str = f"| {display_params}" if display_params else ""
            display_texts.append(f"{indent_str}{prefix}{action_label} {param_str}{cache_str}")
            
            if act.startswith('IF_') or act == 'LOOP_START': block_stack.append(act)
            elif act in ['END_IF', 'END_LOOP'] and block_stack: block_stack.pop()

        self.steps_listbox.delete(0, tk.END)
        if display_texts: self.steps_listbox.insert(tk.END, *display_texts)
        
        if self.editing_index is not None and self.editing_index < len(display_texts):
             self.steps_listbox.itemconfig(self.editing_index, {'bg':'#fff9e1', 'fg':'#e6a23c'})
             self.steps_listbox.see(self.editing_index)
             self.steps_listbox.selection_clear(0, tk.END)
             self.steps_listbox.selection_set(self.editing_index)
        elif self.steps_listbox.curselection() and self.steps_listbox.curselection()[0] < len(display_texts):
             self.steps_listbox.see(self.steps_listbox.curselection()[0])

    def remove_step(self):
        sel = self.steps_listbox.curselection()
        if not sel: return
        if self.editing_index in sel: self.cancel_edit_mode()
        for i in reversed(sel): del self.steps[i]
        self.update_listbox_display()

    def move_step(self, d):
        sel = self.steps_listbox.curselection()
        if not sel: return
        i = sel[0]
        new_i = i - 1 if d == "up" else i + 1
        if 0 <= new_i < len(self.steps):
            self.steps.insert(new_i, self.steps.pop(i))
            if self.editing_index == i: self.editing_index = new_i
            elif self.editing_index == new_i: self.editing_index = i
            self.update_listbox_display()
            self.steps_listbox.selection_set(new_i)

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
        if not self.is_app_running: return
        try:
            text = None
            max_updates = 10 
            count = 0
            while not self.status_queue.empty() and count < max_updates:
                text = self.status_queue.get_nowait()
                count += 1
            
            if text: self.loop_status_var.set(text)
        except queue.Empty:
            pass
        except Exception as e:
            print(f"[StatusQueue] 错误: {e}")
            
        self.root.after(100, self._check_status_queue)

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
            with open(f, 'r', encoding='utf-8') as file: self.steps = json.load(file)
            self.update_listbox_display()
            self.status_var.set(f"已加载: {os.path.basename(f)}")
            self.add_to_recent_files(f)
        except Exception as e: messagebox.showerror("失败", str(e))

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
