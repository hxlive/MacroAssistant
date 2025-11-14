# MacroAssistant.py
# 描述：自动化宏的 GUI 界面
# 版本：1.50.0
# 变更：应用最终的动作列表（两位数前缀），实现完美对齐

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

# =================================================================
# 全局配置
# =================================================================
APP_VERSION = "1.50.0"
APP_TITLE = f"宏助手 (Macro Assistant) V{APP_VERSION}"
APP_ICON = "app_icon.ico" 
CONFIG_FILE = "macro_settings.json"
MAX_RECENT_FILES = 5
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
except ImportError:
    messagebox.showerror("导入错误", "未找到 'core_engine.py' 或 'ocr_engine.py'。\n请确保它们与 'MacroAssistant.py' 位于同一目录。")
    exit()

# 【变更】使用两位数前缀实现完美对齐
ACTION_TRANSLATIONS = {
    'FIND_IMAGE':     '01.  查找图像',
    'FIND_TEXT':      '02.  查找文本 (OCR)',
    'MOVE_OFFSET':    '03.  相对移动',
    'MOVE_TO':        '04.  移动到 (绝对坐标)',
    'CLICK':          '05.  点击鼠标',
    'SCROLL':         '06.  滚动滚轮',
    'WAIT':           '07.  等待',
    'TYPE_TEXT':      '08.  输入文本',
    'PRESS_KEY':      '09.  按下按键',
    'IF_IMAGE_FOUND': '10.  IF 找到图像',
    'IF_TEXT_FOUND':  '11.  IF 找到文本',
    'ELSE':           '12.  ELSE',
    'END_IF':         '13.  END_IF',
    'LOOP_START':     '14.  循环开始 (Loop)',
    'END_LOOP':       '15.  结束循环 (EndLoop)',
}
LANG_OPTIONS = {'chi_sim (简体中文)': 'chi_sim', 'eng (英文)': 'eng'}
CLICK_OPTIONS = {'left (左键)': 'left', 'right (右键)': 'right', 'middle (中键)': 'middle'}
ACTION_KEYS_TO_NAME = {v: k for k, v in ACTION_TRANSLATIONS.items()}
LANG_VALUES_TO_NAME = {v: k for k, v in LANG_OPTIONS.items()}
CLICK_VALUES_TO_NAME = {v: k for k, v in CLICK_OPTIONS.items()}


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("950x700")
        
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
        self.current_theme = tb.StringVar(value=self.root.style.theme_use())
        self.skip_confirm_var = tb.BooleanVar(value=False)
        self.dont_minimize_var = tb.BooleanVar(value=False)
        self.recent_files = []
        self.status_queue = queue.Queue()
        
        self.mouse_tracker_job = None
        self.mouse_pos_var = tb.StringVar()
        
        self.dynamic_wrap_labels = []
        
        self.menu_bar = tk.Menu(root)
        self.root.config(menu=self.menu_bar)
        
        file_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  文件  ", menu=file_menu)
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

        theme_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.font_ui)
        self.menu_bar.add_cascade(label="  主题  ", menu=theme_menu)
        
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
        self.status_var.set("准备就绪...   |   [Ctrl+F10] 启动宏   |   [Ctrl+F11] 停止宏")
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

        self.run_btn = ttk.Button(left_bottom_frame, text="▶ 运行宏 (Ctrl+F10)", command=self.run_macro, bootstyle="success", padding=(15, 10))
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
        self.action_type = ttk.Combobox(add_frame, state="readonly", width=30, font=self.font_ui, height=15)
        self.action_type['values'] = list(ACTION_TRANSLATIONS.values())
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
        
        self.root.after(1000, self.start_hotkey_listener)
        self.root.after(2000, lambda: threading.Thread(target=ocr_engine.preload_engines, daemon=True).start())
        self._check_status_queue()

    def on_exit(self):
        self.is_app_running = False
        if self.mouse_tracker_job:
            self.root.after_cancel(self.mouse_tracker_job)
        try:
            self.root.quit()
            self.root.destroy()
        except Exception: pass

    def update_param_fields(self, event):
        self.last_test_location = None
        
        if self.mouse_tracker_job:
            self.root.after_cancel(self.mouse_tracker_job)
            self.mouse_tracker_job = None
        self.mouse_pos_var.set("")
        
        self.dynamic_wrap_labels.clear()
        
        for widget in self.param_frame.winfo_children(): widget.destroy()
        self.param_widgets = {}
        action_key = ACTION_KEYS_TO_NAME.get(self.action_type.get())
        if not action_key: return
        
        if action_key == 'FIND_IMAGE':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_param_entry("confidence", "置信度(0.1-1.0):", "0.8")
            self._create_hint_label(self.param_frame, "* 提示：如果识别失败，请尝试调低置信度 (如 0.7)")
            self.create_browse_button()
            self.create_test_button("🧪 测试查找图像", self.on_test_find_image_click)
        elif action_key == 'FIND_TEXT':
            self.create_param_entry("text", "查找的文本:", "确定")
            self.create_param_combobox("lang", "语言:", list(LANG_OPTIONS.keys()))
            ocr_status = f"V{ocr_engine.ocr_engine_version}"
            self._create_hint_label(self.param_frame, f"* OCR 引擎: {ocr_status}")
            self.create_test_button("🧪 测试查找文本 (OCR)", self.on_test_find_text_click)
        elif action_key == 'MOVE_OFFSET':
            self.create_param_entry("x_offset", "X 偏移:", "10")
            self.create_param_entry("y_offset", "Y 偏移:", "0")
        elif action_key == 'CLICK':
            self.create_param_combobox("button", "按钮:", list(CLICK_OPTIONS.keys()))
        
        elif action_key == 'SCROLL':
            self.create_param_entry("amount", "滚动量 (正数=上, 负数=下):", "100")
            self.create_param_entry("x", "X 坐标 (可选):", "")
            self.create_param_entry("y", "Y 坐标 (可选):", "")
            self._create_hint_label(self.param_frame, "* 提示: 如果 X, Y 为空, \n  将在当前鼠标位置滚动。")

        elif action_key == 'WAIT':
            self.create_param_entry("ms", "等待 (毫秒):", "500")
        elif action_key == 'TYPE_TEXT':
            self.create_param_entry("text", "输入文本:", "你好")
            self._create_hint_label(self.param_frame, "* 此功能使用剪贴板 (Ctrl+V) \n  以支持中文及复杂文本输入。")
        elif action_key == 'PRESS_KEY':
            self.create_param_entry("key", "按键或组合键 (如 enter, ctrl+c):", "enter")
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
            self.create_param_combobox("lang", "语言:", list(LANG_OPTIONS.keys()))
            self.create_test_button("🧪 测试 IF 文本", self.on_test_find_text_click)
        elif action_key == 'LOOP_START':
            self.create_param_entry("times", "循环次数:", "10")

    def create_param_entry(self, key, label_text, default_value):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        entry = ttk.Entry(frame, width=30, font=self.font_ui)
        entry.insert(0, default_value)
        entry.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = entry
        
    def create_param_combobox(self, key, label_text, values):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=self.font_ui).pack(anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=28, font=self.font_ui)
        combo.current(0)
        combo.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=8)
        self.param_widgets[key] = combo
        
    def create_browse_button(self):
        btn = ttk.Button(self.param_frame, text="浏览...", command=self.browse_image, bootstyle="info-outline", padding=(10, 6))
        btn.pack(anchor="w", fill=tk.X, pady=2)

    def create_test_button(self, text, command):
        ttk.Separator(self.param_frame, orient='horizontal').pack(fill='x', pady=(15, 5))
        ttk.Button(self.param_frame, text=text, command=command, bootstyle="info", padding=(10, 6)).pack(anchor="w", fill=tk.X, pady=2)

    def _create_hint_label(self, parent, text):
        label = ttk.Label(parent, text=text, wraplength=200, font=self.font_ui, foreground='gray')
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
            lang = LANG_OPTIONS.get(self.param_widgets['lang'].get(), 'eng')
            engine = macro_engine.FORCE_OCR_ENGINE or 'auto'
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
        self.status_var.set("准备就绪...")
        self.root.attributes('-topmost', False)

    def _on_test_error(self, e):
        self.root.deiconify()
        messagebox.showerror("错误", str(e))
        self.status_var.set("准备就绪...")

    def browse_image(self):
        f = filedialog.askopenfilename(filetypes=[("PNG", "*.png"), ("All", "*.*")])
        if f: self.param_widgets['path'].delete(0, tk.END); self.param_widgets['path'].insert(0, f)

    def add_or_update_step(self):
        action = ACTION_KEYS_TO_NAME.get(self.action_type.get())
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

                params[k] = LANG_OPTIONS.get(val, val) if k=='lang' else CLICK_OPTIONS.get(val, val) if k=='button' else val
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
        self.action_type.set(ACTION_TRANSLATIONS.get(step['action']))
        self.update_param_fields(None)
        
        for k, v in step['params'].items():
            if k in self.param_widgets:
                val = LANG_VALUES_TO_NAME.get(v, v) if k=='lang' else CLICK_VALUES_TO_NAME.get(v, v) if k=='button' else v
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

            prefix = "[编辑] -> " if i == self.editing_index else f"步骤 {i+1}: "
            
            action_label = ACTION_TRANSLATIONS.get(act, act)
            
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
        threading.Thread(target=self._hotkey_listener_thread, daemon=True).start()

    def _hotkey_listener_thread(self):
        try:
            with keyboard.Listener(on_press=self.on_hotkey_press, on_release=self.on_hotkey_release) as l: l.join()
        except Exception as e: 
            self.root.after(0, lambda err=e: messagebox.showerror("热键启动失败", f"错误: {err}"))

    def on_hotkey_press(self, key):
        try:
            k = None
            if key == keyboard.Key.f10: k = 'f10'
            elif key == keyboard.Key.f11: k = 'f11'
            elif key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r): k = 'ctrl'
            if k and k not in self.held_keys:
                self.held_keys.add(k)
                if 'f11' in self.held_keys and 'ctrl' in self.held_keys:
                    self.root.after(0, self.safe_stop_macro)
                elif 'f10' in self.held_keys and 'ctrl' in self.held_keys:
                    self.root.after(0, self.safe_run_macro)
        except: pass

    def on_hotkey_release(self, key):
        k = None
        try:
            if key == keyboard.Key.f10: k = 'f10'
            elif key == keyboard.Key.f11: k = 'f11'
            elif key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r): k = 'ctrl'
        finally:
            if k and k in self.held_keys: self.held_keys.remove(k)
        
    def safe_run_macro(self):
        if not self.is_macro_running and self.editing_index is None: self.run_macro(True)
        
    def safe_stop_macro(self):
        if self.is_macro_running:
            self.status_var.set("正在停止...")
            if self.current_run_context: self.current_run_context['stop_requested'] = True
        
    def run_macro(self, hotkey=False):
        if self.is_macro_running or not self.steps: return
        if not hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno("运行", "是否立即开始？(按 Ctrl+F11 停止)"): return
        self.loop_status_var.set("") 
        while not self.status_queue.empty():
            try: self.status_queue.get_nowait()
            except queue.Empty: break
        self.run_btn.config(state="disabled")
        self.status_var.set("宏正在运行... [Ctrl+F11] 停止")
        if not self.dont_minimize_var.get(): self.root.iconify()
        else: self.root.attributes('-topmost', True) 
        self.root.after(1500, self._start_macro_thread)

    def _start_macro_thread(self):
        self.is_macro_running = True
        threading.Thread(target=self._run, args=(self.steps.copy(),), daemon=True).start()
        
    def _run(self, steps):
        try:
            self.current_run_context = {'stop_requested': False}
            macro_engine.execute_steps(steps, run_context=self.current_run_context, status_callback=self.update_loop_status)
        except Exception as e: self.root.after(0, lambda err=e: messagebox.showerror("错误", str(err)))
        finally: self.root.after(0, self._on_macro_complete)

    def _on_macro_complete(self):
        self.is_macro_running = False
        self.current_run_context = None
        self.root.deiconify()
        self.root.attributes('-topmost', False)
        self.run_btn.config(state="normal")
        self.status_var.set("准备就绪...")

    def update_loop_status(self, text):
        self.status_queue.put(text)

    def _check_status_queue(self):
        if not self.is_app_running: return
        try:
            text = None
            while not self.status_queue.empty(): text = self.status_queue.get_nowait()
            if text: self.loop_status_var.set(text)
        except: pass
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
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                    self.recent_files = d.get('recent_files', [])
                    self.current_theme.set(d.get('theme', 'litera'))
        except: pass
        self.root.style.theme_use(self.current_theme.get())

    def save_app_settings(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({'recent_files': self.recent_files, 'theme': self.current_theme.get()}, f)
        except: pass

    def change_theme(self):
        self.root.style.theme_use(self.current_theme.get())
        self.root.style.configure(".", font=self.font_ui)
        self.save_app_settings()

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