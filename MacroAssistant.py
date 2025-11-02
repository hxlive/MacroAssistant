# MacroAssistant.py
# 描述：自动化宏的 GUI 界面 (V1.29.2 - 修复循环计数)
# 修复内容：正确传递 run_context 以支持循环计数显示

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

# =================================================================
# 【V1.29.2 优化】更新版本号
# =================================================================
APP_VERSION = "1.29.2 (修复循环计数)"
APP_TITLE = f"宏助手 (Macro Assistant) V{APP_VERSION}"
APP_ICON = "app_icon.ico" 
CONFIG_FILE = "macro_settings.json"
MAX_RECENT_FILES = 5
# =================================================================

# --- (V1.19 的 resource_path 函数保持不变) ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# =================================================================
# 【V39 重构】导入新的核心引擎和 OCR 引擎
# =================================================================
try:
    import core_engine as macro_engine
    import ocr_engine                  
except ImportError:
    messagebox.showerror("导入错误", "未找到 'core_engine.py' 或 'ocr_engine.py'。\n请确保它们与 'MacroAssistant.py' 位于同一目录。")
    exit()
# =================================================================

# --- (V1.03 的字典定义保持不变) ---
ACTION_TRANSLATIONS = {
    'FIND_IMAGE': '1. 查找图像',
    'FIND_TEXT': '2. 查找文本 (OCR)',
    'MOVE_OFFSET': '3. 相对移动',
    'CLICK': '4. 点击鼠标',
    'WAIT': '5. 等待',
    'TYPE_TEXT': '6. 输入文本 (中文/粘贴)',
    'PRESS_KEY': '7. 按下按键',
    'MOVE_TO': '8. 移动到 (绝对坐标)',
    'IF_IMAGE_FOUND': '9. IF 找到图像',
    'IF_TEXT_FOUND': '10. IF 找到文本',
    'ELSE': '11. ELSE',
    'END_IF': '12. END_IF',
    'LOOP_START': '13. 循环开始 (Loop)',
    'END_LOOP': '14. 结束循环 (EndLoop)',
}
LANG_OPTIONS = {
    'chi_sim (简体中文)': 'chi_sim',
    'eng (英文)': 'eng',
}
CLICK_OPTIONS = {
    'left (左键)': 'left',
    'right (右键)': 'right',
    'middle (中键)': 'middle'
}
ACTION_KEYS_TO_NAME = {v: k for k, v in ACTION_TRANSLATIONS.items()}
LANG_VALUES_TO_NAME = {v: k for k, v in LANG_OPTIONS.items()}
CLICK_VALUES_TO_NAME = {v: k for k, v in CLICK_OPTIONS.items()}


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("950x700")
        
        icon_path = resource_path(APP_ICON) 
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
                print(f"[配置] 成功加载本地图标: {icon_path}")
            except tk.TclError:
                print(f"[警告] 找到图标文件 {icon_path}，但无法加载。")
        else:
            print(f"[配置] 未找到图标文件 {icon_path}。将使用默认图标。")
        
        self.steps = []
        self.editing_index = None
        self.is_macro_running = False
        
        # 【V1.27】用于存储上一次成功测试的坐标
        self.last_test_location = None 
        
        self.skip_confirm_var = tb.BooleanVar(value=False)
        self.dont_minimize_var = tb.BooleanVar(value=False)
        
        # 【V1.29】最近文件列表
        self.recent_files = []
        
        # =================================================================
        # 【V1.29】创建菜单栏
        # =================================================================
        self.menu_bar = tk.Menu(root)
        self.root.config(menu=self.menu_bar)
        
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="文件", menu=file_menu)
        
        file_menu.add_command(label="加载宏...", command=self.load_macro)
        file_menu.add_command(label="保存宏...", command=self.save_macro)
        file_menu.add_separator()
        
        # "最近加载"子菜单
        self.recent_files_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="最近加载", menu=self.recent_files_menu)
        
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)

        # =================================================================
        # 【V1.29】创建双层状态栏
        # =================================================================
        status_bar_frame = ttk.Frame(root, bootstyle="primary")
        status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪... | [F10] 启动宏 | [F11] 停止宏")
        self.status_label_left = ttk.Label(
            status_bar_frame, 
            textvariable=self.status_var, 
            relief=tk.FLAT, 
            anchor=tk.W, 
            padding=5,
            bootstyle="primary-inverse"
        )
        self.status_label_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.loop_status_var = tk.StringVar()
        self.loop_status_label_right = ttk.Label(
            status_bar_frame, 
            textvariable=self.loop_status_var, 
            relief=tk.FLAT, 
            anchor=tk.E, 
            padding=(0, 5, 5, 5),
            bootstyle="primary-inverse"
        )
        self.loop_status_label_right.pack(side=tk.RIGHT)
        # =================================================================

        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 左侧布局 (V1.23 的并排开关布局) ---
        list_frame = ttk.Frame(main_frame, padding=10)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(list_frame, text="宏步骤序列:", font=("微软雅黑", 12, "bold")).pack(anchor="w")

        left_bottom_frame = ttk.Frame(list_frame)
        left_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        left_bottom_frame.columnconfigure(0, weight=1)
        left_bottom_frame.columnconfigure(1, weight=1)
        left_bottom_frame.columnconfigure(2, weight=1)
        left_bottom_frame.columnconfigure(3, weight=1)

        self.move_up_btn = ttk.Button(left_bottom_frame, text="上移", command=lambda: self.move_step("up"), bootstyle="secondary-outline")
        self.move_up_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2), pady=(0, 5))
        self.move_down_btn = ttk.Button(left_bottom_frame, text="下移", command=lambda: self.move_step("down"), bootstyle="secondary-outline")
        self.move_down_btn.grid(row=0, column=1, sticky="nsew", padx=2, pady=(0, 5))
        self.remove_btn = ttk.Button(left_bottom_frame, text="删除选中", command=self.remove_step, bootstyle="danger-outline")
        self.remove_btn.grid(row=0, column=2, sticky="nsew", padx=2, pady=(0, 5))
        self.load_step_btn = ttk.Button(left_bottom_frame, text="加载步骤 [修改]", command=self.load_step_for_edit, bootstyle="info-outline")
        self.load_step_btn.grid(row=0, column=3, sticky="nsew", padx=(2, 0), pady=(0, 5))

        # 【V1.29】修改：保存/加载按钮移至菜单栏，这里只保留运行按钮
        self.run_btn = ttk.Button(left_bottom_frame, text="运行宏 (F10)", command=self.run_macro, bootstyle="primary")
        self.run_btn.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=(0, 0), pady=5) 
        
        check_frame = ttk.Frame(left_bottom_frame)
        check_frame.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=(10, 0))
        check_frame.columnconfigure(0, weight=1)
        check_frame.columnconfigure(1, weight=1) 
        skip_check = ttk.Checkbutton(check_frame, 
                                     text="跳过运行前的确认提示", 
                                     variable=self.skip_confirm_var, 
                                     bootstyle="primary-round-toggle")
        skip_check.grid(row=0, column=0, sticky="w", padx=2) 
        minimize_check = ttk.Checkbutton(check_frame, 
                                         text="运行时主界面不最小化", 
                                         variable=self.dont_minimize_var, 
                                         bootstyle="primary-round-toggle")
        minimize_check.grid(row=0, column=1, sticky="w", padx=2)
        
        self.steps_listbox = tk.Listbox(list_frame, width=55, font=("Consolas", 10))
        self.steps_listbox.pack(fill=tk.BOTH, expand=True, pady=5) 

        # --- 右侧布局 ---
        add_frame = ttk.Labelframe(main_frame, text="添加新步骤", padding=10)
        add_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10, expand=True)
        right_bottom_frame = ttk.Frame(add_frame)
        right_bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10,0))
        right_bottom_frame.columnconfigure(0, weight=2) 
        right_bottom_frame.columnconfigure(1, weight=1) 
        self.add_step_btn = ttk.Button(right_bottom_frame, text="添加到序列 >>", command=self.add_or_update_step, bootstyle="success")
        self.add_step_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 2))
        self.cancel_edit_btn = ttk.Button(right_bottom_frame, text="[ 取消修改 ]", command=self.cancel_edit_mode, bootstyle="secondary")
        ttk.Label(add_frame, text="选择动作:").pack(anchor="w")
        self.action_type = ttk.Combobox(add_frame, state="readonly", width=30, font=("微软雅黑", 9), height=15)
        self.action_type['values'] = list(ACTION_TRANSLATIONS.values())
        self.action_type.current(0)
        self.action_type.pack(anchor="w", fill=tk.X, pady=5)
        self.action_type.bind("<<ComboboxSelected>>", self.update_param_fields)
        self.param_frame = ttk.Frame(add_frame)
        self.param_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.param_widgets = {}
        self.update_param_fields(None)
        
        # --- 【V1.29】启动时加载配置并启动监听 ---
        self.load_app_settings()
        self.update_recent_files_menu()
        self.start_hotkey_listener()

    # --- (update_fields, create_... 等函数) ---
    def update_param_fields(self, event):
        # 【V1.27】切换动作时，清除上一次的测试坐标
        self.last_test_location = None
        
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        self.param_widgets = {}
        selected_action_chinese = self.action_type.get()
        action_key = ACTION_KEYS_TO_NAME.get(selected_action_chinese)
        if not action_key: return
        
        if action_key == 'FIND_IMAGE':
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_param_entry("confidence", "置信度(0.1-1.0):", "0.8")
            ttk.Label(self.param_frame, text="* 提示：如果识别失败，请尝试调低置信度 (如 0.7)", wraplength=200).pack(anchor="w", pady=5)
            self.create_browse_button()
            self.create_test_button("🧪 测试查找图像", self.on_test_find_image_click)

        elif action_key == 'FIND_TEXT':
            self.create_param_entry("text", "查找的文本:", "确定")
            self.create_param_combobox("lang", "语言:", list(LANG_OPTIONS.keys()))
            # 【V1.28】更新提示
            ttk.Label(self.param_frame, text="* OCR 引擎: WinOCR (快速) + Tesseract (兜底)", wraplength=200).pack(anchor="w", pady=5)
            self.create_test_button("🧪 测试查找文本 (OCR)", self.on_test_find_text_click)

        elif action_key == 'MOVE_OFFSET':
            self.create_param_entry("x_offset", "X 偏移量 (右为+, 左为-):", "10")
            self.create_param_entry("y_offset", "Y 偏移量 (下为+, 上为-):", "0")
            
        elif action_key == 'CLICK':
            self.create_param_combobox("button", "按钮:", list(CLICK_OPTIONS.keys()))
            
        elif action_key == 'WAIT':
            self.create_param_entry("ms", "等待 (毫秒):", "500")
            
        elif action_key == 'TYPE_TEXT':
            self.create_param_entry("text", "输入的文本:", "你好")
            ttk.Label(self.param_frame, text="* 此功能使用剪贴板 (Ctrl+V) \n  以支持中文及复杂文本输入。", wraplength=200).pack(anchor="w", pady=5)
            
        elif action_key == 'PRESS_KEY':
            self.create_param_entry("key", "按键名称 (例如: enter, tab, f1):", "enter")
            
        elif action_key == 'MOVE_TO':
            self.create_param_entry("x", "X 绝对坐标:", "100")
            self.create_param_entry("y", "Y 绝对坐标:", "100")
            
        elif action_key == 'IF_IMAGE_FOUND':
            ttk.Label(self.param_frame, text="[条件] 如果找到这个图像:").pack(anchor="w")
            self.create_param_entry("path", "图像路径:", "button.png")
            self.create_param_entry("confidence", "置信度(0.1-1.0):", "0.8")
            self.create_browse_button()
            self.create_test_button("🧪 测试 IF 图像", self.on_test_find_image_click)

        elif action_key == 'IF_TEXT_FOUND':
            ttk.Label(self.param_frame, text="[条件] 如果找到这段文本:").pack(anchor="w")
            self.create_param_entry("text", "查找的文本:", "确定")
            self.create_param_combobox("lang", "语言:", list(LANG_OPTIONS.keys()))
            self.create_test_button("🧪 测试 IF 文本 (OCR)", self.on_test_find_text_click)

        elif action_key == 'ELSE':
            ttk.Label(self.param_frame, text="[逻辑] 否则... (如果 IF 条件为假)").pack(anchor="w")
            
        elif action_key == 'END_IF':
            ttk.Label(self.param_frame, text="[逻辑] 结束 IF/ELSE 块").pack(anchor="w")
            
        elif action_key == 'LOOP_START':
            ttk.Label(self.param_frame, text="[逻辑] 开始一个循环:").pack(anchor="w")
            self.create_param_entry("times", "循环次数:", "10")
            
        elif action_key == 'END_LOOP':
            ttk.Label(self.param_frame, text="[逻辑] 结束循环块").pack(anchor="w")

    def create_param_entry(self, key, label_text, default_value):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=("微软雅黑", 9)).pack(anchor="w")
        entry = ttk.Entry(frame, width=30)
        entry.insert(0, default_value)
        entry.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=5)
        self.param_widgets[key] = entry
        
    def create_param_combobox(self, key, label_text, values):
        frame = ttk.Frame(self.param_frame)
        ttk.Label(frame, text=label_text, font=("微软雅黑", 9)).pack(anchor="w")
        combo = ttk.Combobox(frame, values=values, state="readonly", width=28)
        combo.current(0)
        combo.pack(anchor="w", fill=tk.X)
        frame.pack(fill=tk.X, pady=5)
        self.param_widgets[key] = combo
        
    def create_browse_button(self):
        btn = ttk.Button(self.param_frame, text="浏览...", command=self.browse_image, bootstyle="info-outline")
        btn.pack(anchor="w", fill=tk.X, pady=2)

    # =================================================================
    # 【V1.28】测试按钮相关功能 (V1.27 逻辑)
    # =================================================================
    def create_test_button(self, text, command):
        """在参数框底部创建一个测试按钮"""
        sep = ttk.Separator(self.param_frame, orient='horizontal')
        sep.pack(fill='x', pady=(15, 5))
        btn = ttk.Button(self.param_frame, text=text, command=command, bootstyle="info")
        btn.pack(anchor="w", fill=tk.X, pady=2)

    def on_test_find_image_click(self):
        """处理"测试查找图像"按钮点击"""
        try:
            path = self.param_widgets['path'].get()
            confidence = float(self.param_widgets['confidence'].get())
        except KeyError:
            messagebox.showerror("测试错误", "无法找到 'path' 或 'confidence' 控件。")
            return
        except ValueError:
            messagebox.showerror("测试错误", "置信度必须是一个数字 (例如 0.8)。")
            return
            
        if not path or not os.path.exists(path):
            messagebox.showerror("测试错误", f"图像路径无效或文件不存在:\n{path}")
            return

        self.status_var.set("测试中... 2秒后开始查找图像，请切换窗口。")
        self.root.iconify()
        self.root.after(2000, lambda: self._run_test_thread(
            self._test_find_image, (path, confidence)
        ))

    def on_test_find_text_click(self):
        """处理"测试查找文本"按钮点击"""
        try:
            text = self.param_widgets['text'].get()
            lang_key = self.param_widgets['lang'].get()
            lang = LANG_OPTIONS.get(lang_key, 'eng')
        except KeyError:
            messagebox.showerror("测试错误", "无法找到 'text' 或 'lang' 控件。")
            return
        except Exception as e:
            messagebox.showerror("测试错误", f"获取参数时出错: {e}")
            return

        if not text:
            messagebox.showerror("测试错误", "查找的文本不能为空。")
            return
            
        self.status_var.set(f"测试中... 2秒后开始 OCR (查找 '{text}')...")
        self.root.iconify()
        self.root.after(2000, lambda: self._run_test_thread(
            self._test_find_text, (text, lang)
        ))

    def _run_test_thread(self, test_function, args):
        """启动一个新线程来运行测试，防止 GUI 冻结"""
        print(f"[测试线程] 启动测试: {test_function.__name__} {args}")
        self.last_test_location = None
        threading.Thread(target=test_function, args=args, daemon=True).start()

    def _test_find_image(self, path, confidence):
        """在线程中执行图像查找"""
        try:
            self.status_var.set(f"正在查找图像 '{os.path.basename(path)}'...")
            location = macro_engine.find_image_location(path, confidence)
            self.root.after(0, lambda: self._on_test_complete(location))
        except Exception as e:
            self.root.after(0, lambda: self._on_test_error(e))

    def _test_find_text(self, text, lang):
        """在线程中执行文本查找"""
        try:
            self.status_var.set(f"正在查找文本 '{text}' (OCR)...")
            location = ocr_engine.find_text_location(text, lang, debug=True)
            self.root.after(0, lambda: self._on_test_complete(location))
        except Exception as e:
            self.root.after(0, lambda: self._on_test_error(e))

    def _on_test_complete(self, location):
        """[主线程] 测试完成的回调函数"""
        if not self.root.state() == 'normal': self.root.deiconify()
        self.root.attributes('-topmost', True) 
        
        if location:
            # 【V1.27】测试成功！存储这个坐标
            self.last_test_location = location
            print(f"[测试成功] 缓存坐标: {location}")
            
            self.status_var.set(f"测试成功！找到于 {location}，正在移动鼠标...")
            pyautogui.moveTo(location[0], location[1], duration=0.25)
            # 【V1.28】更新提示
            messagebox.showinfo("测试成功", f"已找到目标于 {location}\n鼠标已移动。\n\n点击\"添加到序列\"时，可选择将此坐标添加为\"缓存提示\"。")
        else:
            self.last_test_location = None 
            self.status_var.set("测试失败。未找到目标。")
            messagebox.showwarning("测试失败", "未找到目标。\n请检查控制台日志获取详细的 OCR 调试信息。")
            
        self.status_var.set("准备就绪... | [F10] 启动宏 | [F11] 停止宏")
        self.root.attributes('-topmost', False) 

    def _on_test_error(self, e):
        """[主线程] 测试发生错误的回调函数"""
        self.last_test_location = None 
        if not self.root.state() == 'normal': self.root.deiconify()
        self.root.attributes('-topmost', True)
        messagebox.showerror("测试出错", f"测试时发生意外错误:\n{e}\n\n请检查控制台日志。")
        self.status_var.set("测试出错。| [F10] 启动 | [F11] 停止")
        self.root.attributes('-topmost', False)
    # =================================================================

    def browse_image(self):
        if 'path' in self.param_widgets:
            filepath = filedialog.askopenfilename(
                title="选择图像文件",
                filetypes=(("PNG files", "*.png"), ("All files", "*.*"))
            )
            if filepath:
                self.param_widgets['path'].delete(0, tk.END)
                self.param_widgets['path'].insert(0, filepath)
                
    # =================================================================
    # 【V1.28】优化 add_or_update_step 以支持"坐标提示"
    # =================================================================
    def add_or_update_step(self):
        selected_action_chinese = self.action_type.get()
        action_key = ACTION_KEYS_TO_NAME.get(selected_action_chinese)
        if not action_key:
            messagebox.showwarning("错误", "未选择有效的动作。"); return

        # 1. 收集所有参数
        params = {}
        no_param_actions = ['ELSE', 'END_IF', 'END_LOOP']
        try:
            for key, widget in self.param_widgets.items():
                value = widget.get()
                if not value and action_key not in no_param_actions:
                    messagebox.showwarning("输入错误", f"参数 '{key}' 不能为空。"); return
                if key == 'lang': params[key] = LANG_OPTIONS.get(value, value)
                elif key == 'button': params[key] = CLICK_OPTIONS.get(value, value)
                else: params[key] = value
        except Exception as e:
            messagebox.showerror("参数错误", f"获取参数时出错: {e}"); return
            
        step_to_add = { "action": action_key, "params": params }

        # 2. 检查是否可以应用"坐标提示"
        # (仅适用于"添加新步骤"，不适用于"更新步骤")
        # (仅适用于 FIND 和 IF_FIND 动作)
        if (action_key in ('FIND_TEXT', 'FIND_IMAGE', 'IF_TEXT_FOUND', 'IF_IMAGE_FOUND') and 
            self.editing_index is None and 
            self.last_test_location is not None):
            
            # 弹窗询问
            msg = f"测试成功，已在 {self.last_test_location} 找到目标。\n\n" \
                  "您想将这个坐标添加为\"缓存提示\"吗？\n\n" \
                  "[是] = 添加步骤并包含缓存 (推荐, 运行时更快)\n" \
                  "[否] = 添加步骤但不含缓存 (每次都全局搜索)\n" \
                  "[取消] = 不添加"
            
            result = messagebox.askyesnocancel("使用测试坐标？", msg)
            
            if result is None: # 取消
                self.last_test_location = None # 清除缓存
                return
            elif result is True: # 是 (Yes) -> 添加坐标到 params
                step_to_add["params"]["cache_x"] = self.last_test_location[0]
                step_to_add["params"]["cache_y"] = self.last_test_location[1]
                print(f"坐标提示 {self.last_test_location} 已添加到步骤。")
            # else: (否) -> 不做任何事，使用不带缓存的 step_to_add

        # 3. 添加或更新步骤
        if self.editing_index is not None:
            self.steps[self.editing_index] = step_to_add
            print(f"步骤 {self.editing_index + 1} 已更新: {step_to_add}")
            self.cancel_edit_mode() # cancel_edit_mode 会清除缓存和更新列表
        else:
            self.steps.append(step_to_add)
            print(f"步骤已添加: {step_to_add}")
            self.steps_listbox.see(tk.END)
            # 【V1.28】清除缓存并更新列表
            self.last_test_location = None
            self.update_listbox_display()
        
    def load_step_for_edit(self):
        try:
            selected_indices = self.steps_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("提示", "请先在左侧列表中选中一个要修改的步骤。"); return
            index = selected_indices[0]
            step_data = self.steps[index]
            print(f"正在加载步骤 {index + 1} 进行编辑...")
            action_name = ACTION_TRANSLATIONS.get(step_data['action'])
            if not action_name:
                messagebox.showerror("加载错误", f"未知的动作: {step_data['action']}"); return
            
            self.action_type.set(action_name)
            self.update_param_fields(None) # 这会清除 last_test_location
            
            params = step_data.get('params', {})
            for key, widget in self.param_widgets.items():
                value = params.get(key)
                if value is None: continue
                
                # 【V1.28】加载时，不要加载 cache_x/y 到 GUI 控件中
                if key in ('cache_x', 'cache_y'): continue 
                
                if key == 'lang': value = LANG_VALUES_TO_NAME.get(value, value)
                elif key == 'button': value = CLICK_VALUES_TO_NAME.get(value, value)
                
                if isinstance(widget, ttk.Combobox): widget.set(value)
                elif isinstance(widget, ttk.Entry):
                    widget.delete(0, tk.END); widget.insert(0, str(value))
                    
            self.editing_index = index
            self.add_step_btn.config(text="✓ 更新步骤", bootstyle="warning")
            self.cancel_edit_btn.grid(row=0, column=1, sticky="nsew", padx=(2, 0))
            self.update_listbox_display() # 高亮
        except Exception as e: messagebox.showerror("加载失败", f"加载步骤时出错: {e}")
        
    def cancel_edit_mode(self):
        self.editing_index = None
        self.last_test_location = None # 【V1.27】清除缓存
        self.add_step_btn.config(text="添加到序列 >>", bootstyle="success")
        self.cancel_edit_btn.grid_remove()
        self.action_type.current(0)
        self.update_param_fields(None)
        self.update_listbox_display(); print("修改已取消。")
        
    def update_listbox_display(self):
        self.steps_listbox.delete(0, tk.END)
        editing_item_index = self.editing_index
        indent_level = 0; block_stack = []
        for i, step in enumerate(self.steps):
            action_key = step.get('action', '')
            chinese_action = ACTION_TRANSLATIONS.get(action_key, action_key)
            display_params = step.get('params', {}).copy()
            current_indent_level = len(block_stack)
            if action_key in ['ELSE', 'END_IF', 'END_LOOP']:
                if block_stack: current_indent_level = max(0, len(block_stack) - 1)
            indent_str = "    " * current_indent_level
            
            # 【V1.28】在列表中显示缓存坐标
            cache_str = ""
            if 'cache_x' in display_params:
                cache_str = f" [Cache: {display_params['cache_x']}, {display_params['cache_y']}]"
                # 不在列表中显示
                del display_params['cache_x']
                if 'cache_y' in display_params: del display_params['cache_y']
            
            if 'lang' in display_params:
                lang_key = [k for k, v in LANG_OPTIONS.items() if v == display_params['lang']]
                if lang_key: display_params['lang'] = lang_key[0]
            if 'button' in display_params:
                button_key = [k for k, v in CLICK_OPTIONS.items() if v == display_params['button']]
                if button_key: display_params['button'] = button_key[0]
                
            prefix = "[编辑中] -> " if i == editing_item_index else f"步骤 {i+1}: "
            param_str = f"| {display_params}" if display_params else ""
            display_text = f"{indent_str}{prefix}{chinese_action} {param_str}{cache_str}"
            
            self.steps_listbox.insert(tk.END, display_text)
            if i == editing_item_index:
                self.steps_listbox.itemconfig(i, {'bg':'#fff9e1', 'fg':'#e6a23c'})
            if action_key.startswith('IF_') or action_key == 'LOOP_START': block_stack.append(action_key)
            elif action_key == 'ELSE': pass
            elif action_key == 'END_IF':
                if block_stack and block_stack[-1].startswith('IF_'): block_stack.pop()
            elif action_key == 'END_LOOP':
                if block_stack and block_stack[-1] == 'LOOP_START': block_stack.pop()
                
    def remove_step(self):
        try:
            selected_indices = self.steps_listbox.curselection()
            if not selected_indices: return
            if self.editing_index in selected_indices: self.cancel_edit_mode()
            for index in reversed(selected_indices): del self.steps[index]
            self.update_listbox_display(); print("步骤已删除。")
        except Exception as e: messagebox.showerror("错误", f"删除失败: {e}")
        
    def move_step(self, direction):
        try:
            selected_indices = self.steps_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("提示", "请先在列表中选中一个步骤。"); return
            index = selected_indices[0]
            if direction == "up":
                if index == 0: return
                new_index = index - 1
            elif direction == "down":
                if index == len(self.steps) - 1: return
                new_index = index + 1
            else: return
            step_to_move = self.steps.pop(index); self.steps.insert(new_index, step_to_move)
            if self.editing_index == index: self.editing_index = new_index
            elif self.editing_index == new_index: self.editing_index = index
            self.update_listbox_display(); self.steps_listbox.selection_set(new_index)
        except Exception as e: messagebox.showerror("错误", f"移动步骤时出错: {e}")

    # --- (热键和多线程执行函数保持 V1.01 不变) ---
    def start_hotkey_listener(self):
        listener_thread = threading.Thread(target=self._hotkey_listener_thread, daemon=True)
        listener_thread.start(); print("全局热键监听器已启动...")
        
    def _hotkey_listener_thread(self):
        try:
            with keyboard.Listener(on_press=self.on_hotkey_press) as listener:
                listener.join()
        except Exception as e: print(f"热键监听线程出错: {e}")
        
    def on_hotkey_press(self, key):
        try:
            if key == keyboard.Key.f10:
                print("[热键] 检测到 F10 (启动)"); self.root.after(0, self.safe_run_macro)
            elif key == keyboard.Key.f11:
                print("[热键] 检测到 F11 (停止)"); self.root.after(0, self.safe_stop_macro)
        except Exception as e: print(f"热键回调处理失败: {e}")
        
    def safe_run_macro(self):
        if self.is_macro_running: print("[主线程] 宏已在运行，F10 被忽略。"); return
        if self.editing_index is not None:
            messagebox.showwarning("提示", "您正处于编辑模式。\n请先\"更新步骤\"或\"取消修改\"。"); return
        print("[主线程] 热键 F10 触发 run_macro()"); self.run_macro(from_hotkey=True)
        
    def safe_stop_macro(self):
        if not self.is_macro_running: print("[主线程] 宏未在运行，F11 被忽略。"); return
        print("[主线程] 热键 F11 触发紧急停止 (moveTo 0,0)...")
        self.status_var.set("正在停止... | [F10] 启动 | [F11] 停止"); pyautogui.moveTo(0, 0)
        
    # (V1.21 的 run_macro 逻辑保持不变)
    def run_macro(self, from_hotkey=False):
        if self.is_macro_running: return
        if self.editing_index is not None:
            messagebox.showwarning("提示", "您正处于编辑模式。\n请先\"更新步骤\"或\"取消修改\"。"); return
        if not self.steps:
            messagebox.showinfo("提示", "宏序列为空，请先添加步骤。"); return
            
        if not from_hotkey and not self.skip_confirm_var.get():
            if not messagebox.askyesno("准备运行",
                    f"将按顺序执行 {len(self.steps)} 个步骤。\n\n"
                    f"【重要】要中途紧急停止，请将鼠标快速移动到屏幕的【左上角】(0, 0)。\n\n"
                    f"是否立即开始？"):
                return
                
        print("--- 准备运行宏 ---")
        
        # =========================================================
        # 【V1.29.2】清除上一次的循环状态 (新运行开始时)
        self.loop_status_var.set("") 
        # =========================================================
        
        self.run_btn.config(state="disabled")
        self.status_var.set("宏正在运行... [F11] 停止")

        if not self.dont_minimize_var.get():
            print("GUI 窗口将最小化... 1.5秒后开始执行...")
            self.root.iconify()
        else:
            print("GUI 窗口将保持可见... 1.5秒后开始执行...")
            self.root.attributes('-topmost', True) 

        self.root.after(1500, self._start_macro_thread)

    def _start_macro_thread(self):
        print("...延迟结束，正在启动新的工作线程。")
        self.is_macro_running = True
        macro_thread = threading.Thread(target=self._run_macro_in_thread,
                                        args=(self.steps.copy(),),
                                        daemon=True)
        macro_thread.start()
        
    def _run_macro_in_thread(self, steps_copy):
        try:
            # =========================================================
            # 【V1.29.2 修复】创建 run_context 并传递
            print("...工作线程已启动，开始调用核心引擎。")
            
            # 创建 run_context 字典
            run_context = {
                'stop_requested': False,
                'loops_executed': 0
            }
            
            # 传递 run_context 和 status_callback
            macro_engine.execute_steps(
                steps_copy, 
                run_context=run_context,  # ← 修复：添加这个参数
                status_callback=self.update_loop_status
            )
            # =========================================================
            
            self.root.after(0, self._on_macro_complete)
        except pyautogui.FailSafeException:
            print("--- 宏被用户紧急停止！(来自工作线程) ---")
            self.root.after(0, self._on_macro_failsafe)
        except Exception as e:
            print(f"--- 宏执行出错(来自工作线程): {e} ---")
            self.root.after(0, lambda: self._on_macro_error(e))

    # =================================================================
    # (V1.24 的弹窗逻辑保持不变)
    # =================================================================
    def _on_macro_failsafe(self):
        print("[主线程] _on_macro_failsafe (紧急停止) 回调")
        self.is_macro_running = False
        if not self.root.state() == 'normal': self.root.deiconify()
        self.root.attributes('-topmost', False) 
        messagebox.showwarning("紧急停止", "宏已被用户（或 F11 热键）紧急停止。")
        self.run_btn.config(state="normal")
        self.status_var.set("宏已停止。| [F10] 启动 | [F11] 停止")

    def _on_macro_error(self, error):
        print("[主线程] _on_macro_error (执行出错) 回调")
        self.is_macro_running = False
        if not self.root.state() == 'normal': self.root.deiconify()
        self.root.attributes('-topmost', False) 
        messagebox.showerror("执行出错", f"执行宏时发生意外错误:\n{error}")
        self.run_btn.config(state="normal")
        self.status_var.set("宏因错误停止。| [F10] 启动 | [F11] 停止")
        
    def _on_macro_complete(self):
        print("[主线程] _on_macro_complete (正常完成) 回调")
        self.is_macro_running = False
        
        if not self.root.state() == 'normal':
            self.root.deiconify()
        self.root.attributes('-topmost', False)
        
        if self.dont_minimize_var.get():
            messagebox.showinfo("完成", "宏执行完毕。")
        else:
            pass 
            
        self.run_btn.config(state="normal")
        self.status_var.set("准备就绪... | [F10] 启动宏 | [F11] 停止宏")
    # =================================================================

    # =================================================================
    # 【V1.29 新增】文件加载/保存 和 最近列表 逻辑
    # =================================================================
    
    def update_loop_status(self, text):
        """[跨线程回调] 更新循环状态标签"""
        # 使用 root.after 确保在主 GUI 线程中更新
        self.root.after(0, self.loop_status_var.set, text)

    def save_macro(self):
        filepath = filedialog.asksaveasfilename(
            title="保存宏文件",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.steps, f, indent=4)
            messagebox.showinfo("成功", "宏已保存！")
            # 【V1.29】添加到最近列表
            self.add_to_recent_files(filepath)
        except Exception as e: 
            messagebox.showerror("保存失败", f"无法保存文件: {e}")
        
    def load_macro(self):
        filepath = filedialog.askopenfilename(
            title="加载宏文件",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not filepath: return
        # 【V1.29】调用新的加载函数
        self._load_file_path(filepath)
        
    def _load_file_path(self, filepath):
        """【V1.29】从指定路径加载宏文件"""
        if not os.path.exists(filepath):
            messagebox.showerror("加载失败", f"文件不存在: {filepath}")
            # 从最近列表中移除无效路径
            if filepath in self.recent_files:
                self.recent_files.remove(filepath)
                self.save_app_settings()
                self.update_recent_files_menu()
            return
            
        try:
            self.cancel_edit_mode()
            with open(filepath, 'r', encoding='utf-8') as f:
                self.steps = json.load(f)
            self.update_listbox_display()
            
            # 【V1.29】不弹窗，改用状态栏提示
            filename = os.path.basename(filepath)
            self.status_var.set(f"已加载: {filename} | [F10] 启动 | [F11] 停止")
            
            # 【V1.29】添加到最近列表
            self.add_to_recent_files(filepath)
            
        except Exception as e: 
            messagebox.showerror("加载失败", f"无法加载文件: {e}")

    def load_app_settings(self):
        """【V1.29】启动时加载配置"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.recent_files = data.get('recent_files', [])
                    print(f"[配置] 加载了 {len(self.recent_files)} 个最近文件。")
        except Exception as e:
            print(f"[配置] 加载 {CONFIG_FILE} 失败: {e}")
            self.recent_files = []

    def save_app_settings(self):
        """【V1.29】保存配置"""
        try:
            data = {'recent_files': self.recent_files}
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"[配置] 保存 {CONFIG_FILE} 失败: {e}")
            
    def add_to_recent_files(self, filepath):
        """【V1.29】添加一个文件到最近列表"""
        # 确保路径是绝对路径
        abs_path = os.path.abspath(filepath)
        
        # 移除已存在的
        if abs_path in self.recent_files:
            self.recent_files.remove(abs_path)
        
        # 插入到最前面
        self.recent_files.insert(0, abs_path)
        
        # 保持列表长度
        self.recent_files = self.recent_files[:MAX_RECENT_FILES]
        
        # 更新菜单并保存
        self.update_recent_files_menu()
        self.save_app_settings()
        
    def update_recent_files_menu(self):
        """【V1.29】更新"最近加载"菜单的显示"""
        self.recent_files_menu.delete(0, tk.END)
        
        if not self.recent_files:
            self.recent_files_menu.add_command(label="(无)", state="disabled")
            return
            
        for i, path in enumerate(self.recent_files):
            filename = os.path.basename(path)
            # 使用 lambda p=path: ... 来正确捕获循环中的 path 变量
            self.recent_files_menu.add_command(
                label=f"{i+1}. {filename}", 
                command=lambda p=path: self._load_file_path(p)
            )

# --- 程序入口 ---
if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    main_window = tb.Window(themename="litera")
    app = MacroApp(main_window)
    main_window.mainloop()