# -*- coding: utf-8 -*-
# gui_utils.py
# 描述：GUI 辅助工具库 (重构版 - 样式完美还原)
# 版本：1.2.6

import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
from PIL import Image, ImageTk
import os
import time

# 引入核心库中的工具用于处理快捷键显示
try:
    from core_engine import HotkeyUtils
except ImportError:
    # Fallback
    class HotkeyUtils:
        @staticmethod
        def format_hotkey_display(s): return s.upper()

# =================================================================
# 1. 基础工具函数
# =================================================================
def parse_region_string(region_str):
    """
    解析 "x1,y1,x2,y2" 字符串为整数列表。
    
    Note:
        此函数仅用于 UI 输入解析。
        宏数据持久化时会自动转换为 'cache_box' 字段。
        
    Args:
        region_str (str): 格式为 "x1,y1,x2,y2" 的坐标字符串
        
    Returns:
        list[int] | None: [x1, y1, x2, y2] 或 None (解析失败时)
    """
    if not region_str: return None
    try:
        parts = region_str.replace('，', ',').split(',')
        coords = [int(x.strip()) for x in parts if x.strip()]
        return coords if len(coords) == 4 else None
    except (ValueError, TypeError, IndexError, AttributeError):
        return None

# =================================================================
# 2. 鼠标位置追踪器 (MouseTracker)
# =================================================================
class MouseTracker:
    def __init__(self, root, tk_var):
        self.root = root
        self.var = tk_var
        self.job = None
        self.is_running = False

    def start(self):
        if not self.is_running:
            self.is_running = True
            self._update()

    def stop(self):
        self.is_running = False
        if self.job:
            try: self.root.after_cancel(self.job)
            except: pass
            self.job = None
            self.var.set("")

    def _update(self):
        if not self.is_running: return
        try:
            x, y = pyautogui.position()
            self.var.set(f"X: {x}, Y: {y}")
        except Exception as e:
            self.var.set("未知")
            # 打印异常但不中断，防止刷屏
            # print(f"[MouseTracker] Error: {e}") 
        self.job = self.root.after(100, self._update)

# =================================================================
# 3. 自动换行标签 (AutoWrapLabel)
# =================================================================
class AutoWrapLabel(ttk.Label):
    def __init__(self, master, **kwargs):
        # [优化] 给一个合理的初始换行宽度，防止布局抖动
        if 'wraplength' not in kwargs:
            kwargs['wraplength'] = 250
        super().__init__(master, **kwargs)
        self.bind('<Configure>', self._on_configure)

    def _on_configure(self, event):
        width = event.width - 15
        if width > 0:
            self.configure(wraplength=width)

# =================================================================
# 4. 区域选择器 (RegionSelector)
# =================================================================
class RegionSelector:
    def __init__(self, master):
        self.master = master
        self.selection = None
        self.is_selecting = False
        self.start_x = 0; self.start_y = 0; self.cur_x = 0; self.cur_y = 0
        
        self.top = tk.Toplevel(self.master)
        self.top.attributes('-fullscreen', True)
        self.top.attributes('-alpha', 0.3)
        self.top.attributes('-topmost', True)
        self.top.configure(cursor="cross")
        self.top.overrideredirect(True)
        
        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.top.bind("<Escape>", self.on_cancel)
        self.top.bind("<Return>", self.on_confirm)
        
        w, h = self.top.winfo_screenwidth(), self.top.winfo_screenheight()
        self.canvas.create_text(w//2, h//2, text="按住左键拖拽 | Enter确认 | ESC取消", 
                                fill="white", font=("Arial", 20, "bold"), tag="hint")

    def on_mouse_down(self, event):
        self.is_selecting = True
        self.start_x = self.top.winfo_pointerx() - self.top.winfo_rootx()
        self.start_y = self.top.winfo_pointery() - self.top.winfo_rooty()
        self.canvas.delete("hint")

    def on_mouse_move(self, event):
        if not self.is_selecting: return
        self.cur_x = self.top.winfo_pointerx() - self.top.winfo_rootx()
        self.cur_y = self.top.winfo_pointery() - self.top.winfo_rooty()
        self.canvas.delete("rect")
        self.canvas.create_rectangle(self.start_x, self.start_y, self.cur_x, self.cur_y, outline="red", width=2, tag="rect")

    def on_mouse_up(self, event):
        self.is_selecting = False
        self._finish_selection()

    def on_confirm(self, event):
        self._finish_selection()

    def _finish_selection(self):
        x1, y1 = min(self.start_x, self.cur_x), min(self.start_y, self.cur_y)
        x2, y2 = max(self.start_x, self.cur_x), max(self.start_y, self.cur_y)
        if (x2 - x1) > 5 and (y2 - y1) > 5:
            self.selection = [x1, y1, x2, y2]
            self.top.destroy()

    def on_cancel(self, event):
        self.selection = None
        self.top.destroy()

    def get_region(self):
        self.master.wait_window(self.top)
        return self.selection

# =================================================================
# 5. 图片悬浮预览 (ImageTooltipManager)
# =================================================================
class ImageTooltipManager:
    def __init__(self, treeview, app_steps_getter):
        self.tree = treeview
        self.get_steps = app_steps_getter
        self.tooltip_window = None
        self.timer = None
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def on_select(self, event):
        if self.timer: self.tree.after_cancel(self.timer)
        self.hide_tooltip()
        sel = self.tree.selection()
        if not sel: return
        self.timer = self.tree.after(500, lambda: self.show_tooltip(sel[0]))

    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def show_tooltip(self, row_id):
        try:
            steps = self.get_steps()
            if not steps: return
            idx = self.tree.index(row_id)
            if idx >= len(steps): return
            
            path = steps[idx].get('params', {}).get('path')
            if not path or not os.path.exists(path): return
            
            # [优化] 使用 try-finally 确保资源清理
            window = None
            try:
                window = tk.Toplevel(self.tree)
                window.withdraw()
                window.wm_overrideredirect(True)
                
                frame = ttk.Frame(window, relief='solid', borderwidth=1)
                frame.pack()
                
                img = Image.open(path)
                img.thumbnail((300, 300), Image.Resampling.LANCZOS)
                self.tk_img = ImageTk.PhotoImage(img)
                
                ttk.Label(frame, image=self.tk_img).pack(padx=2, pady=2)
                ttk.Label(frame, text=os.path.basename(path), font=('Arial', 8), foreground='#666').pack()
                
                window.update_idletasks()
                x, y = self.tree.winfo_pointerx() + 15, self.tree.winfo_pointery() + 10
                window.wm_geometry(f'+{x}+{y}')
                window.attributes('-topmost', True)
                window.deiconify()
                
                self.tooltip_window = window
                window = None
            finally:
                if window: window.destroy()
        except (OSError, IOError, tk.TclError, AttributeError) as e:
            print(f"[Tooltip] 显示失败: {e}")
            self.hide_tooltip()

# =================================================================
# 6. 快捷键输入控件 (HotkeyEntry) - 保持样式一致
# =================================================================
class HotkeyEntry(ttk.Entry):
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
        disp = HotkeyUtils.format_hotkey_display(hotkey_str) if hotkey_str else "点击 [捕获] 录制"
        self._update_text(disp)
        if self.string_var and hotkey_str: self.string_var.set(hotkey_str)

    def _update_text(self, text):
        self.configure(state="normal")
        self.delete(0, tk.END)
        self.insert(0, text)
        self.configure(state="readonly")

    def _on_focus_in(self, e): self._update_text("录制中...")
    def _on_focus_out(self, e): 
        if not self.current_keys and self.string_var: self.set_hotkey(self.string_var.get())
        self.current_keys.clear()

    def _on_key_press(self, e):
        k = self._get_key_name(e)
        if k: 
            self.current_keys.add(k)
            self._display_current_keys()
        return "break"

    def _on_key_release(self, e):
        k = self._get_key_name(e)
        if k and k not in ('ctrl','alt','shift','cmd'):
            self._display_current_keys(final=True)
            self.current_keys.clear()
            self.master.focus()
        return "break"

    def _display_current_keys(self, final=False):
        mods = [k for k in ['ctrl','alt','shift','cmd'] if k in self.current_keys]
        key = next((k for k in self.current_keys if k not in mods), None)
        res = "+".join(mods + [key]) if key else "+".join(mods)
        self._update_text(HotkeyUtils.format_hotkey_display(res))
        if final and key and self.string_var: self.string_var.set(res)

    def _get_key_name(self, event):
        n = event.keysym.lower()
        if "control" in n: return "ctrl"
        if "alt" in n: return "alt"
        if "shift" in n: return "shift"
        if "win" in n or "super" in n: return "cmd"
        if n.startswith("f") and n[1:].isdigit(): return n
        if len(n)==1 and n.isalnum(): return n
        return {'return':'enter','space':'space','tab':'tab','capital':'caps_lock','escape':'esc','backspace':'backspace','delete':'delete','prior':'page_up','next':'page_down','end':'end','home':'home','left':'left','up':'up','right':'right','down':'down','insert':'insert'}.get(n, None)

# =================================================================
# 7. 快捷键设置弹窗 (HotkeySettingsDialog) - 完美还原版
# =================================================================
class HotkeySettingsDialog:
    """快捷键设置对话框 (完全还原旧版样式与位置逻辑)"""
    def __init__(self, parent, current_run, current_stop, default_run="ctrl+f10", default_stop="ctrl+f11"):
        self.result = None
        self.default_run = default_run
        self.default_stop = default_stop
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("快捷键设置")
        self.dialog.geometry("450x480") 
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # === 核心修复：还原窗口居中逻辑 (相对于父窗口) ===
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="⌨️ 自定义快捷键", 
                  font=("Microsoft YaHei UI", 12, "bold")).pack(pady=(0, 15))
        
        # --- 还原 Run 区域布局 ---
        run_frame = ttk.Labelframe(main_frame, text="运行/继续 快捷键", padding=15)
        run_frame.pack(fill=tk.X, pady=(0, 15))
        run_inner = ttk.Frame(run_frame)
        run_inner.pack(fill=tk.X)
        run_inner.columnconfigure(0, weight=1)

        self.run_var = tk.StringVar(value=current_run)
        self.run_display = HotkeyEntry(run_inner, textvariable=self.run_var)
        self.run_display.set_hotkey(current_run)
        self.run_display.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        
        ttk.Button(run_inner, text="🎯 录制", 
                   command=self.run_display.focus_set,
                   bootstyle="info", width=12).grid(row=0, column=1, ipady=3)
        
        # --- 还原 Stop 区域布局 ---
        stop_frame = ttk.Labelframe(main_frame, text="停止宏快捷键", padding=15)
        stop_frame.pack(fill=tk.X, pady=(0, 15))
        stop_inner = ttk.Frame(stop_frame)
        stop_inner.pack(fill=tk.X)
        stop_inner.columnconfigure(0, weight=1)
        
        self.stop_var = tk.StringVar(value=current_stop)
        self.stop_display = HotkeyEntry(stop_inner, textvariable=self.stop_var)
        self.stop_display.set_hotkey(current_stop)
        self.stop_display.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        
        ttk.Button(stop_inner, text="🎯 录制", 
                   command=self.stop_display.focus_set,
                   bootstyle="info", width=12).grid(row=0, column=1, ipady=3)
        
        # --- 还原提示文字 ---
        hint_frame = ttk.Frame(main_frame)
        hint_frame.pack(fill=tk.X, pady=(20, 20))
        hint_text = "💡 支持: Ctrl, Alt, Shift, F1-F12, A-Z, 0-9等"
        ttk.Label(hint_frame, text=hint_text, font=("Microsoft YaHei UI", 9), 
                  foreground="#666", justify=tk.LEFT).pack()
        
        # --- 还原底部按钮布局 (Grid 3列) ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        
        ttk.Button(btn_frame, text="✕ 取消", command=self.dialog.destroy, 
                bootstyle="secondary", padding=(10, 10)).grid(row=0, column=0, sticky="ew", padx=(5, 0))
        # 还原“恢复默认”按钮
        ttk.Button(btn_frame, text="🔄 恢复默认", command=self.reset_default, 
                bootstyle="warning-outline", padding=(10, 10)).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(btn_frame, text="✓ 保存", command=self.save, 
                bootstyle="success", padding=(10, 10)).grid(row=0, column=2, sticky="ew", padx=(0, 5))
        
    def reset_default(self):
        self.run_var.set(self.default_run)
        self.run_display.set_hotkey(self.default_run)
        self.stop_var.set(self.default_stop)
        self.stop_display.set_hotkey(self.default_stop)
        
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
