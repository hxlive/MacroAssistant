# sys_utils.py
# 描述: 系统底层工具、全局热键管理及稳定工具类集
# 版本: 1.8.1

import sys
import os
import threading
import functools
import base64
import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import ctypes
from PIL import Image, ImageTk
from pynput import keyboard

# 引入核心库中的工具
from core_engine import HotkeyUtils

# ======================================================================
# 1. 系统底层初始化 (DPI, 流, AppID)
# ======================================================================

def init_system_runtime():
    """初始化系统运行环境（流重定向、DPI感知等）"""
    if sys.platform == 'win32':
        # 1. 重构标准输出编码
        #    优先级：
        #    1) MACROMATE_STDIO_ENCODING（--log-encoding 显式覆盖）
        #    2) UTF-8（兼容现代终端与工具输出捕获，Windows 10+ 控制台原生支持）
        #    Python 默认编码在 Windows 管道场景下常为 GBK，导致 UTF-8 终端显示乱码
        try:
            stdio_encoding = (
                os.environ.get('MACROMATE_STDIO_ENCODING')
                or os.environ.get('MACROASSISTANT_STDIO_ENCODING')
                or 'utf-8'
            )
            sys.stdout.reconfigure(encoding=stdio_encoding, errors='replace')
            sys.stderr.reconfigure(encoding=stdio_encoding, errors='replace')
            print(f"[CONFIG] STDIO encoding: {stdio_encoding}")
        except AttributeError:
            pass
            
        # 2. 强制启用 DPI 感知 (解决 125%/150% 缩放下的坐标偏移)
        try:
            # 设置 DPI 感知级别为 "PerMonitorV2" (Awareness 2)
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                # 回退旧版 API (兼容 Win7/8)
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

def set_windows_app_id(app_version):
    """设置 Windows AppUserModelID 以确保任务栏图标显示正确"""
    if sys.platform == 'win32':
        try:
            myappid = f'hxlive.macromate.{app_version}'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            return True
        except Exception as e:
            print(f"[警告] 设置 AppUserModelID 失败: {e}")
    return False

# ======================================================================
# 2. 快捷键冲突检测支持
# ======================================================================
HOTKEY_CHECK_AVAILABLE = False
_WIN_MODIFIER_VK_MAP = {
    'ctrl': (0x11,),       # VK_CONTROL
    'alt': (0x12,),        # VK_MENU
    'shift': (0x10,),      # VK_SHIFT
    'cmd': (0x5B, 0x5C),   # VK_LWIN, VK_RWIN
}
if sys.platform == 'win32':
    HOTKEY_CHECK_AVAILABLE = True

# ======================================================================
# 3. 鼠标位置追踪器 (MouseTracker)
# ======================================================================
def _center_child_window(parent, dialog):
    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - dialog.winfo_width()) // 2
    y = parent.winfo_y() + (parent.winfo_height() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")

class MouseTracker:
    def __init__(self, root, tk_var):
        self.root = root
        self.var = tk_var
        self.job = None
        self.is_running = False
        self._lock = threading.RLock()

    def start(self):
        with self._lock:
            if self.is_running:
                return
            self.is_running = True
        self._update()

    def stop(self):
        with self._lock:
            self.is_running = False
            job = self.job
            self.job = None
        if job:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self.var.set("")

    def _update(self):
        with self._lock:
            if not self.is_running:
                return
        try:
            x, y = pyautogui.position()
            self.var.set(f"X: {x}, Y: {y}")
        except Exception:
            self.var.set("未知")
        with self._lock:
            if self.is_running:
                self.job = self.root.after(100, self._update)

# ======================================================================
# 4. 区域选择器 (RegionSelector)
# ======================================================================
class RegionSelector:
    def __init__(self, master):
        self.master = master
        self.selection = None
        self.is_selecting = False
        self.has_dragged = False
        self.rect = None
        self.start_x = 0; self.start_y = 0; self.cur_x = 0; self.cur_y = 0

        # 获取虚拟显示器的联合（Virtual Screen）坐标，以完美覆盖多屏
        self.offset_x = 0
        self.offset_y = 0
        w = self.master.winfo_screenwidth()
        h = self.master.winfo_screenheight()
        if sys.platform == 'win32':
            try:
                SM_XVIRTUALSCREEN = 76
                SM_YVIRTUALSCREEN = 77
                SM_CXVIRTUALSCREEN = 78
                SM_CYVIRTUALSCREEN = 79
                user32 = ctypes.windll.user32
                x_val = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                y_val = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                w_val = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                h_val = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
                if w_val > 0 and h_val > 0:
                    self.offset_x, self.offset_y, w, h = x_val, y_val, w_val, h_val
            except Exception as e:
                print(f"[RegionSelector] 获取多屏几何信息失败: {e}")

        self.top = tk.Toplevel(self.master)
        ox = f"+{self.offset_x}" if self.offset_x >= 0 else str(self.offset_x)
        oy = f"+{self.offset_y}" if self.offset_y >= 0 else str(self.offset_y)
        self.top.geometry(f"{w}x{h}{ox}{oy}")
        self.top.attributes('-alpha', 0.3)
        self.top.attributes('-topmost', True)
        self.top.configure(cursor="cross")
        self.top.overrideredirect(True)

        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self._bind_events()

    def _bind_events(self):
        self.top.bind("<ButtonPress-1>", self._on_press)
        self.top.bind("<B1-Motion>", self._on_drag)
        self.top.bind("<ButtonRelease-1>", self._on_release)
        self.top.bind("<Escape>", self._on_cancel)
        self.top.bind("<Return>", self._on_confirm)

    def _on_confirm(self, event=None):
        self._finalize_selection()

    def _finalize_selection(self):
        if getattr(self, 'has_dragged', self.cur_x != 0 or self.cur_y != 0):
            x1, y1 = min(self.start_x, self.cur_x) + self.offset_x, min(self.start_y, self.cur_y) + self.offset_y
            x2, y2 = max(self.start_x, self.cur_x) + self.offset_x, max(self.start_y, self.cur_y) + self.offset_y
            if abs(x2 - x1) > 5 and abs(y2 - y1) > 5:
                self.selection = (x1, y1, x2, y2)
        self.top.destroy()

    def _on_press(self, event):
        self.is_selecting = True
        self.start_x, self.start_y = event.x, event.y
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="red", width=2, fill="white", stipple="gray50"
        )

    def _on_drag(self, event):
        if self.is_selecting:
            self.has_dragged = True
            self.cur_x, self.cur_y = event.x, event.y
            self.canvas.coords(self.rect, self.start_x, self.start_y, self.cur_x, self.cur_y)

    def _on_release(self, event):
        if self.is_selecting:
            self.is_selecting = False
            self.cur_x, self.cur_y = event.x, event.y
            self._finalize_selection()

    def _on_cancel(self, event=None):
        self.is_selecting = False
        self.top.destroy()

    def get_region(self):
        self.master.wait_window(self.top)
        return self.selection

# ======================================================================
# 5. 全局热键管理器 (GlobalHotkeyManager)
# ======================================================================
class GlobalHotkeyManager:
    def __init__(self, root, get_run_str_cb, get_stop_str_cb, trigger_run_cb, trigger_stop_cb):
        self.root = root
        self.get_run_str = get_run_str_cb
        self.get_stop_str = get_stop_str_cb
        self.trigger_run = trigger_run_cb
        self.trigger_stop = trigger_stop_cb
        
        # 缓存快捷键字符串，避免多线程直接调用 Tk 控件
        self.run_hotkey_cache = ""
        self.stop_hotkey_cache = ""
        
        self.held_keys = {}
        self.listener = None
        self._listener_lock = threading.RLock()
        self._listener_generation = 0
        
    def start_listener(self):
        """Start or restart the global hotkey listener."""
        # Read Tk-backed hotkey values on the main thread before listener callbacks use them.
        try:
            run_cache = self.get_run_str()
        except Exception:
            run_cache = ""
        try:
            stop_cache = self.get_stop_str()
        except Exception:
            stop_cache = ""

        with self._listener_lock:
            self._listener_generation += 1
            generation = self._listener_generation
            self.run_hotkey_cache = run_cache
            self.stop_hotkey_cache = stop_cache
            old_listener = self.listener
            if self.listener:
                try:
                    self.listener.stop()
                    self.listener.join(timeout=0.5)
                except Exception as e:
                    print(f"[Hotkey] stop old listener failed: {e}")
            if self.listener is old_listener:
                self.listener = None
            self.held_keys.clear()
            threading.Thread(target=self._listener_thread, args=(generation,), daemon=True).start()
        
    def _listener_thread(self, generation):
        try:
            listener = keyboard.Listener(
                on_press=lambda key: self.on_press(key, generation),
                on_release=lambda key: self.on_release(key, generation),
            )
            with self._listener_lock:
                if generation != self._listener_generation:
                    return
                self.listener = listener
            listener.start()
            listener.join()
        except Exception as e:
            msg = f"热键监听器启动失败: {e}\n\n快捷键将无法工作。请尝试重启程序。"
            self.root.after(0, messagebox.showerror, "严重错误", msg)
            
    def restart_listener(self): self.start_listener()

    def _get_key_name(self, key):
        try:
            if hasattr(key, 'vk') and key.vk in HotkeyUtils.VK_TO_PYNPUT:
                return HotkeyUtils.VK_TO_PYNPUT[key.vk]
            if hasattr(key, 'name') and key.name:
                return key.name.lower()
            if hasattr(key, 'char') and key.char:
                return key.char.lower()
            return str(key).lower()
        except Exception:
            return None

    def _normalize_key(self, key_name):
        if key_name in ('ctrl_l', 'ctrl_r'): return 'ctrl'
        if key_name in ('alt_l', 'alt_r', 'alt_gr'): return 'alt'
        if key_name in ('shift_l', 'shift_r'): return 'shift'
        if key_name in ('cmd_l', 'cmd_r', 'cmd'): return 'cmd'
        return key_name

    def _modifiers_satisfied(self, required_mods):
        if sys.platform == 'win32':
            try:
                for mod, vks in _WIN_MODIFIER_VK_MAP.items():
                    is_pressed = False
                    for vk in vks:
                        if (ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000) != 0:
                            is_pressed = True
                            break
                    if not is_pressed:
                        self.held_keys.pop(mod, None)
                    else:
                        if mod not in self.held_keys:
                            self.held_keys[mod] = 1
            except Exception as e:
                print(f"[Hotkey] check physical keys failed: {e}")
        return all(self.held_keys.get(m, 0) > 0 for m in required_mods)

    def on_press(self, key, generation=None):
        try:
            if generation is not None and generation != self._listener_generation:
                return
            key_name = self._normalize_key(self._get_key_name(key))
            if not key_name: return
            
            # [优化] 长按重复触发拦截，防止多次累加造成计数器残留
            if self.held_keys.get(key_name, 0) > 0:
                return
                
            self.held_keys[key_name] = 1
            # 从本地缓存读取快捷键配置，避免跨线程直接调用 Tk 控件
            run_mods, run_key = self._parse_hotkey(self.run_hotkey_cache)
            if key_name == run_key and self._modifiers_satisfied(run_mods):
                self.root.after(0, self.trigger_run)
            stop_mods, stop_key = self._parse_hotkey(self.stop_hotkey_cache)
            if key_name == stop_key and self._modifiers_satisfied(stop_mods):
                self.root.after(0, self.trigger_stop)
        except Exception as e:
            print(f"[Hotkey] press error: {e}")

    def on_release(self, key, generation=None):
        try:
            if generation is not None and generation != self._listener_generation:
                return
            key_name = self._normalize_key(self._get_key_name(key))
            if not key_name: return
            # [优化] 松开按键时直接清空字典中该键的状态，根治所有残留假死
            self.held_keys.pop(key_name, None)
        except Exception as e:
            print(f"[Hotkey] release error: {e}")

    @functools.lru_cache(maxsize=16)
    def _parse_hotkey(self, hotkey_str):
        if not hotkey_str: return set(), ""
        parts = [p.strip() for p in hotkey_str.lower().split('+')]
        if not parts: return set(), ""
        return set(parts[:-1]), parts[-1]

    def check_conflicts(self, show_success=True):
        if not HOTKEY_CHECK_AVAILABLE: return True
        conflicts = []
        run_str = self.get_run_str()
        if not self._test_register(run_str, 1):
            conflicts.append(f"运行快捷键 '{HotkeyUtils.format_hotkey_display(run_str)}'")
        stop_str = self.get_stop_str()
        if not self._test_register(stop_str, 2):
            conflicts.append(f"停止快捷键 '{HotkeyUtils.format_hotkey_display(stop_str)}'")
        if conflicts:
            msg = "检测到快捷键冲突：\n\n" + "\n".join(conflicts) + "\n\n可能已被其他程序占用。\n请修改快捷键，否则热键可能无法工作。"
            self.root.after(0, messagebox.showwarning, "快捷键冲突", msg)
            return False
        return True

    def _test_register(self, hotkey_str, hotkey_id):
        if not hotkey_str: return True
        try:
            parts = hotkey_str.lower().split('+')
            modifiers, vk = 0, None
            for part in [p.strip() for p in parts]:
                if part in HotkeyUtils.PYNPUT_MOD_TO_WIN_MOD: modifiers |= HotkeyUtils.PYNPUT_MOD_TO_WIN_MOD[part]
                elif part in HotkeyUtils.PYNPUT_TO_VK: vk = HotkeyUtils.PYNPUT_TO_VK[part]
            if vk is None: return True
            if ctypes.windll.user32.RegisterHotKey(None, hotkey_id, modifiers, vk) == 0: return False
            ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
            return True
        except Exception: return True

# ======================================================================
# 6. 快捷键输入控件 (HotkeyEntry)
# ======================================================================
class HotkeyEntry(ttk.Entry):
    def __init__(self, master, hotkey_var, **kwargs):
        super().__init__(master, **kwargs)
        self.hotkey_var = hotkey_var
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Key>", self._on_key)
        self._placeholder = "点击此处，按下快捷键..."
        self._is_recording = False
        self._pressed_keys = set()
        self._display_text = tk.StringVar()
        self.config(textvariable=self._display_text)
        self.refresh_display()
        
        # [终极方案] 在 Windows 下彻底禁用此 Entry 的输入法 (IME)
        # 这会强制该 Entry 接收原始英文按键，绕过所有输入法拦截和乱码问题
        if sys.platform == 'win32':
            self.after(100, self._disable_ime)

    def _disable_ime(self):
        """调用 Windows API 禁用当前组件的输入法上下文"""
        try:
            hwnd = self.winfo_id()
            ctypes.windll.imm32.ImmAssociateContext(hwnd, 0)
        except Exception:
            pass

    def refresh_display(self):
        """外部修改 hotkey_var 后，调用此方法同步显示文本"""
        current = self.hotkey_var.get()
        if current:
            self._display_text.set(HotkeyUtils.format_hotkey_display(current))
            self.config(bootstyle="default")
        else:
            self._display_text.set(self._placeholder)
            self.config(bootstyle="secondary")

    def _on_focus_in(self, event):
        self._is_recording = True
        self._pressed_keys.clear()
        self._display_text.set("按下快捷键组合...")
        self.config(bootstyle="info")

    def _on_focus_out(self, event):
        self._is_recording = False
        self._pressed_keys.clear()
        current = self.hotkey_var.get()
        if current:
            self._display_text.set(HotkeyUtils.format_hotkey_display(current))
            self.config(bootstyle="default")
        else:
            self._display_text.set(self._placeholder)
            self.config(bootstyle="secondary")

    def _on_key(self, event):
        if not self._is_recording: return
        key = event.keysym.lower()

        # [终极杀手锏 2.0] 彻底解决中文输入法拦截问题
        # 在 Windows 中文输入法下，所有按键的 keycode 会被系统统一接管为 229 (VK_PROCESSKEY)
        if sys.platform == 'win32' and getattr(event, 'keycode', None):
            if event.keycode != 229:
                # 英文状态下，直接使用底层硬件码，100% 准确
                vk_key = HotkeyUtils.VK_TO_PYNPUT.get(event.keycode)
                if vk_key: key = vk_key

        # 统一修饰键名称
        if key in ('shift_l', 'shift_r'): key = 'shift'
        elif key in ('control_l', 'control_r'): key = 'ctrl'
        elif key in ('alt_l', 'alt_r', 'alt_gr'): key = 'alt'
        elif key in ('command', 'command_l', 'command_r', 'win', 'win_l', 'win_r'): key = 'cmd'
        
        # [输入法抢救逻辑] 针对输入法拦截 (keycode=229) 或 Tkinter 解析失败 (??) 的情况
        char = getattr(event, 'char', '').lower()
        if key == '??' or getattr(event, 'keycode', None) == 229:
            # 全角/半角符号反向映射表 (包含中文特有符号)
            CHAR_TO_BASE = {
                '!': '1', '@': '2', '#': '3', '$': '4', '%': '5',
                '^': '6', '&': '7', '*': '8', '(': '9', ')': '0',
                '_': '-', '+': '=', '{': '[', '}': ']', '|': '\\',
                ':': ';', '"': "'", '<': ',', '>': '.', '?': '/', '~': '`',
                '！': '1', '＠': '2', '＃': '3', '￥': '4', '％': '5',
                '……': '6', '…': '6', '＆': '7', '＊': '8', '（': '9', '）': '0',
                '—': '-', '＋': '=', '【': '[', '】': ']', '、': '\\', '｜': '\\',
                '；': ';', '：': ';', '‘': "'", '’': "'", '“': "'", '”': "'",
                '，': ',', '《': ',', '。': '.', '》': '.', '？': '/',
                '·': '`', '～': '`'
            }
            if char in CHAR_TO_BASE:
                key = CHAR_TO_BASE[char]
            elif char and len(char) == 1 and (char.isalnum() or char in "`-=[]\\;',./"):
                key = char

        # 兜底：处理 Tkinter 在英文下解析出的特定符号名
        SHIFT_MAP = {
            'exclam': '1', 'at': '2', 'numbersign': '3', 'dollar': '4', 'percent': '5',
            'asciicircum': '6', 'ampersand': '7', 'asterisk': '8', 'parenleft': '9', 'parenright': '0',
            'underscore': '-', 'plus': '=', 'braceleft': '[', 'braceright': ']', 'bar': '\\',
            'colon': ';', 'quotedbl': "'", 'less': ',', 'greater': '.', 'question': '/',
            'yen': '4'  # 特殊补充
        }
        if key in SHIFT_MAP: key = SHIFT_MAP[key]

        self._pressed_keys.clear()

        # 1. 添加当前按下的主键 (绝对排除掉输入法引发的 '??' 乱码)
        if key and key not in ('caps_lock', 'num_lock', 'scroll_lock', 'next', 'prior', '??'):
             self._pressed_keys.add(key)

        # 2. 根据 event.state 添加当前正被按住的修饰键
        is_mac = sys.platform == 'darwin'
        if event.state & 0x0001: self._pressed_keys.add('shift')
        if event.state & 0x0004: self._pressed_keys.add('ctrl')
        
        if is_mac:
            if event.state & 0x0008: self._pressed_keys.add('cmd')
            if event.state & 0x0010: self._pressed_keys.add('alt')
        else:
            if event.state & 0x20000: self._pressed_keys.add('alt')

        # 3. 排序并显示
        order = {'ctrl': 0, 'alt': 1, 'shift': 2, 'cmd': 3}
        sorted_keys = sorted(list(self._pressed_keys), key=lambda k: (order.get(k, 4), k))
        
        if sorted_keys:
            hotkey_str = '+'.join(sorted_keys)
            self.hotkey_var.set(hotkey_str)
            self._display_text.set(HotkeyUtils.format_hotkey_display(hotkey_str))
        return 'break'

# ======================================================================
# 7. 设置对话框 (Hotkey/VLM)
# ======================================================================
class HotkeySettingsDialog:
    # [修复 BUG-4] 恢复快捷键格式校验；默认值修正为 ctrl+f10/ctrl+f11
    def __init__(self, parent, run_hotkey, stop_hotkey,
                 default_run='ctrl+f10', default_stop='ctrl+f11'):
        self.parent = parent
        self.default_run = default_run
        self.default_stop = default_stop
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.withdraw()  # 立即隐藏，防止闪烁
        self.dialog.title("快捷键设置")
        self.dialog.geometry("450x480")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        _center_child_window(parent, self.dialog)
        self.dialog.deiconify()  # 位置确定后再显示
        self.dialog.deiconify()  # 位置确定后再显示

        self._create_ui(run_hotkey, stop_hotkey)
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_ui(self, run_hotkey, stop_hotkey):
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="⌨️ 自定义快捷键",
                  font=("Microsoft YaHei UI", 12, "bold")).pack(pady=(0, 15))

        self.run_var = tk.StringVar(value=run_hotkey)
        run_frame = ttk.Labelframe(main_frame, text="运行/继续 快捷键", padding=15)
        run_frame.pack(fill=tk.X, pady=(0, 15))
        run_inner = ttk.Frame(run_frame)
        run_inner.pack(fill=tk.X)
        run_inner.columnconfigure(0, weight=1)
        self.run_entry = HotkeyEntry(run_inner, self.run_var, width=25)
        self.run_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        ttk.Button(run_inner, text="🎯 录制", command=self.run_entry.focus_set,
                   bootstyle="info", width=12).grid(row=0, column=1, ipady=3)

        self.stop_var = tk.StringVar(value=stop_hotkey)
        stop_frame = ttk.Labelframe(main_frame, text="停止宏快捷键", padding=15)
        stop_frame.pack(fill=tk.X, pady=(0, 15))
        stop_inner = ttk.Frame(stop_frame)
        stop_inner.pack(fill=tk.X)
        stop_inner.columnconfigure(0, weight=1)
        self.stop_entry = HotkeyEntry(stop_inner, self.stop_var, width=25)
        self.stop_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=5)
        ttk.Button(stop_inner, text="🎯 录制", command=self.stop_entry.focus_set,
                   bootstyle="info", width=12).grid(row=0, column=1, ipady=3)

        ttk.Label(main_frame, text="💡 支持: Ctrl, Alt, Shift, F1-F12, A-Z, 0-9等",
                  font=("Microsoft YaHei UI", 9), foreground="#666").pack(pady=(20, 20))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        ttk.Button(btn_frame, text="✕ 取消", command=self._on_close,
                   bootstyle="secondary", padding=(10, 10)).grid(row=0, column=0, sticky="ew", padx=(5, 0))
        ttk.Button(btn_frame, text="🔄 恢复默认", command=self._reset_default,
                   bootstyle="warning-outline", padding=(10, 10)).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(btn_frame, text="✓ 保存", command=self._on_save,
                   bootstyle="success", padding=(10, 10)).grid(row=0, column=2, sticky="ew", padx=(0, 5))

    def _reset_default(self):
        self.run_var.set(self.default_run)
        self.stop_var.set(self.default_stop)
        self.run_entry.refresh_display()
        self.stop_entry.refresh_display()

    def _on_save(self):
        run_hk = self.run_var.get().strip().lower()
        stop_hk = self.stop_var.get().strip().lower()
        if not run_hk or not stop_hk:
            messagebox.showerror("错误", "快捷键不能为空", parent=self.dialog)
            return
        if run_hk == stop_hk:
            messagebox.showerror("错误", "运行和停止快捷键不能相同", parent=self.dialog)
            return
        if not self._validate_hotkey(run_hk):
            messagebox.showerror("错误", f"运行快捷键格式无效: {run_hk}", parent=self.dialog)
            return
        if not self._validate_hotkey(stop_hk):
            messagebox.showerror("错误", f"停止快捷键格式无效: {stop_hk}", parent=self.dialog)
            return
        self.result = (run_hk, stop_hk)
        self.dialog.destroy()

    def _validate_hotkey(self, hotkey):
        parts = hotkey.split('+')
        if not parts:
            return False
        if len(parts) == 1:
            p = parts[0].strip().lower()
            if len(p) == 1 and 'a' <= p <= 'z':
                return True
            if p.startswith('f') and p[1:].isdigit():
                return int(p[1:]) in range(1, 13)
            return False
        modifiers = {'ctrl', 'alt', 'shift', 'cmd'}
        valid_keys = {name for name in HotkeyUtils.PYNPUT_TO_VK.keys()
                      if name not in ('ctrl_l', 'ctrl_r', 'alt_l', 'alt_r', 'alt_gr',
                                      'shift_l', 'shift_r', 'cmd_l', 'cmd_r', 'cmd')}
        for i, part in enumerate(parts):
            part = part.strip()
            if i < len(parts) - 1:
                if part not in modifiers:
                    return False
            else:
                if part not in valid_keys:
                    return False
        return True

    def _on_close(self):
        self.dialog.destroy()

class VLMSettingsDialog:
    # [修复 BUG-3] 恢复完整功能：load_config、测试连接、timeout 设置
    def __init__(self, parent):
        self.result = None
        try:
            import vlm_engine
            self.current_config = vlm_engine.load_config()
            self.providers = vlm_engine.get_providers()
        except Exception:
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
        _center_child_window(parent, self.dialog)
        self.dialog.deiconify()  # 位置确定后再显示
        self.dialog.deiconify()  # 位置确定后再显示

        self._create_ui()

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

    def _on_provider_change(self, event):
        selected = self.provider_var.get()
        provider_key = selected.split(" (")[-1].rstrip(")") if "(" in selected else selected
        if self.providers and provider_key in self.providers:
            default_model = self.providers[provider_key].get('model', '')
            if not self.model_var.get():
                self.model_var.set(default_model)

    def _test_connection(self):
        # [修复H-6] 改为后台线程执行，避免阻塞 UI 事件循环
        try:
            import vlm_engine, io, threading
            selected = self.provider_var.get()
            provider_key = selected.split(" (")[-1].rstrip(")") if "(" in selected else selected
            api_key = self.api_key_var.get().strip()
            if not api_key:
                messagebox.showwarning("提示", "请先输入 API Key", parent=self.dialog)
                return
            config = vlm_engine.DEFAULT_CONFIG.copy()
            config.update({'provider': provider_key, 'api_key': api_key,
                           'timeout': self.timeout_var.get(),
                           'system_prompt': "你是一个助手，直接回答用户问题即可。"})
            if self.model_var.get().strip():
                config['model'] = self.model_var.get().strip()
            if self.providers and provider_key in self.providers:
                config['base_url'] = self.providers[provider_key].get('base_url', '')

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
                    from PIL import ImageGrab
                    screenshot = ImageGrab.grab()
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
        try:
            if self.dialog.winfo_exists():
                self.dialog.after(0, callback)
        except Exception:
            pass

    def _save(self):
        try:
            import vlm_engine
            selected = self.provider_var.get()
            provider_key = selected.split(" (")[-1].rstrip(")") if "(" in selected else selected
            config = vlm_engine.DEFAULT_CONFIG.copy()
            config['provider'] = provider_key
            config['api_key'] = self.api_key_var.get().strip()
            config['timeout'] = self.timeout_var.get()
            if self.model_var.get().strip():
                config['model'] = self.model_var.get().strip()
            if self.providers and provider_key in self.providers:
                config['base_url'] = self.providers[provider_key].get('base_url', '')
            if vlm_engine.save_config(config):
                self.result = config
                self.dialog.destroy()
            else:
                messagebox.showerror("错误", "保存配置失败", parent=self.dialog)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}", parent=self.dialog)

# ======================================================================
# 8. 悬浮提示与迷你窗口 (Tooltip/MiniWindow)
# ======================================================================
class ImageTooltipManager:
    """[修复 BUG-1] 接受 steps 列表或 getter 函数，兼容两种调用方式"""
    def __init__(self, treeview, steps_or_getter):
        self.tree = treeview
        # 支持传入 lambda: self.steps 或直接传入 list
        self._getter = steps_or_getter if callable(steps_or_getter) else lambda: steps_or_getter
        self.tooltip = None
        self.current_item = None
        self._bind_events()

    def _bind_events(self):
        self.tree.bind('<Motion>', self._on_motion)
        self.tree.bind('<Leave>', self._on_leave)

    def _on_motion(self, event):
        item = self.tree.identify_row(event.y)
        if item != self.current_item:
            self.current_item = item
            self._hide_tooltip()
            if item:
                self._show_tooltip(item, event.x_root, event.y_root)

    def _on_leave(self, event):
        self._hide_tooltip()
        self.current_item = None

    def _show_tooltip(self, item, x, y):
        try:
            steps = self._getter()
            if not steps:
                return
            idx = self.tree.index(item)
            if idx < 0 or idx >= len(steps):
                return

            step = steps[idx]
            action = step.get('action', '')
            params = step.get('params', {})

            if action not in ('FIND_IMAGE', 'IF_IMAGE_FOUND'):
                return

            img_path = params.get('path', '')
            if not img_path or not os.path.exists(img_path):
                return

            # [修复BUG-2] 使用 load() 强制解码全部像素，copy() 创建独立副本
            # 避免 with 块关闭文件后 ImageTk.PhotoImage 持有悬空引用
            with Image.open(img_path) as img:
                img.load()  # 强制解码，防止懒加载在文件关闭后失败
                orig_size = img.size  # 在 with 块内读取尺寸
                img_copy = img.copy()

            img_copy.thumbnail((200, 150), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img_copy)

            self.tooltip = tk.Toplevel(self.tree)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x+15}+{y+15}")

            label = ttk.Label(self.tooltip, image=photo)
            label.image = photo
            label.pack()

            info_text = f"{os.path.basename(img_path)}\n{orig_size[0]}x{orig_size[1]}"
            ttk.Label(self.tooltip, text=info_text, font=("Microsoft YaHei UI", 8)).pack()

        except Exception as e:
            print(f"图片提示加载失败: {e}")

    def _hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class MiniStatusWindow:
    """
    宏执行时的迷你悬浮状态栏窗口
    - 无边框、始终置顶，显示于屏幕左下角
    - 点击可停止宏
    """
    def __init__(self, parent, stop_callback):
        self.parent = parent
        self.stop_callback = stop_callback

        self.window = tk.Toplevel(parent)
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)

        window_width = 500
        window_height = 35
        px, py = self.window.winfo_pointerxy()
        screen_height = self.window.winfo_screenheight()
        monitor_x = self.window.winfo_vrootx()
        monitor_y = self.window.winfo_vrooty()
        monitor_h = self.window.winfo_vrootheight() or screen_height
        if sys.platform == 'win32':
            try:
                user32 = ctypes.windll.user32
                vx = user32.GetSystemMetrics(76)
                vy = user32.GetSystemMetrics(77)
                vw = user32.GetSystemMetrics(78)
                vh = user32.GetSystemMetrics(79)
                if vw > 0 and vh > 0 and vx <= px < vx + vw and vy <= py < vy + vh:
                    monitor_x, monitor_y, monitor_h = vx, vy, vh
            except Exception:
                pass
        x = monitor_x + 10
        y = monitor_y + max(0, monitor_h - window_height - 50)
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        main_frame = ttk.Frame(self.window, bootstyle="primary", padding=0)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(
            main_frame, text="运行中...",
            relief=tk.FLAT, anchor=tk.W, padding=(8, 5),
            bootstyle="primary-inverse", font=("Microsoft YaHei UI", 9)
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.loop_label = ttk.Label(
            main_frame, text="",
            relief=tk.FLAT, anchor=tk.E, padding=(0, 5, 8, 5),
            bootstyle="primary-inverse", font=("Microsoft YaHei UI", 9)
        )
        self.loop_label.pack(side=tk.RIGHT)

        for w in (main_frame, self.status_label, self.loop_label):
            w.bind("<Button-1>", self._on_click)
            w.bind("<Enter>", lambda e: self.window.config(cursor="hand2"))
            w.bind("<Leave>", lambda e: self.window.config(cursor=""))

    def _on_click(self, event):
        if self.stop_callback:
            self.stop_callback()

    def update_status(self, status_text, loop_text=""):
        """更新状态栏显示内容"""
        try:
            if self.window.winfo_exists():
                self.status_label.config(text=status_text)
                self.loop_label.config(text=loop_text)
        except tk.TclError:
            pass

    def destroy(self):
        try:
            if self.window.winfo_exists():
                self.window.destroy()
        except tk.TclError:
            pass

class AboutDialog:
    """智点助手关于对话框"""
    def __init__(self, parent, app_version, icon_path=None):
        self.parent = parent
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.withdraw()  # 立即隐藏，防止闪烁
        self.dialog.title("关于")
        self.dialog.geometry("500x400")  # 足够宽度显示完整链接
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 设置窗口图标
        if icon_path and os.path.exists(icon_path):
            try:
                self.dialog.iconbitmap(icon_path)
            except (OSError, tk.TclError) as e:
                print(f"[警告] 设置关于对话框图标失败: {e}")
                
        # 居中显示
        _center_child_window(parent, self.dialog)
        self.dialog.deiconify()  # 位置确定后再显示

        self._create_ui(app_version, icon_path)
        
        # 绑定事件
        self.dialog.bind("<Escape>", lambda e: self.dialog.destroy())
        self.dialog.bind("<Return>", lambda e: self.dialog.destroy())

    def _create_ui(self, app_version, icon_path):
        import webbrowser
        from PIL import Image, ImageTk
        
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 顶部：图标和软件标题区域 ==========
        top_outer = ttk.Frame(main_frame)
        top_outer.pack(fill=tk.X, pady=(5, 18))
        
        top_frame = ttk.Frame(top_outer)
        top_frame.pack(anchor="center")
        
        # 左侧：图标
        icon_container = ttk.Frame(top_frame)
        icon_container.pack(side=tk.LEFT, padx=(0, 28))
        
        if icon_path and os.path.exists(icon_path):
            try:
                with Image.open(icon_path) as _raw:
                    _raw.load()
                    _icon_copy = _raw.copy()
                resized_img = _icon_copy.resize((96, 96), Image.Resampling.LANCZOS)
                icon_photo = ImageTk.PhotoImage(resized_img)
                
                icon_label = ttk.Label(icon_container, image=icon_photo)
                icon_label.image = icon_photo
                icon_label.pack()
            except (OSError, IOError) as e:
                print(f"[警告] 加载图标图像失败: {e}")
                ttk.Label(icon_container, text="🔧", font=("Microsoft YaHei UI", 48)).pack()
        else:
            ttk.Label(icon_container, text="🔧", font=("Microsoft YaHei UI", 48)).pack()
            
        # 右侧：软件标题和版本
        title_container = ttk.Frame(top_frame)
        title_container.pack(side=tk.LEFT, pady=10)
        
        ttk.Label(title_container, text="智点助手", font=("Microsoft YaHei UI", 17, "bold")).pack(anchor="w", pady=(0, 2))
        ttk.Label(title_container, text="MacroMate", font=("Microsoft YaHei UI", 10), foreground="#666666").pack(anchor="w", pady=(0, 6))
        
        version_frame = ttk.Frame(title_container)
        version_frame.pack(anchor="w")
        ttk.Label(version_frame, text=f" v{app_version} ", font=("Consolas", 9, "bold"), bootstyle="info", padding=(6, 2)).pack(side=tk.LEFT)
        
        # ========== 分隔线 ==========
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=(0, 18))
        
        # ========== 中部：详细信息区域 ==========
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 18), padx=5)
        info_frame.columnconfigure(1, weight=1)
        
        ttk.Label(info_frame, text="软件作者", font=("Microsoft YaHei UI", 10, "bold"), foreground="#777777").grid(row=0, column=0, sticky="w", padx=(0, 20), pady=6)
        ttk.Label(info_frame, text="寒星", font=("Microsoft YaHei UI", 10)).grid(row=0, column=1, sticky="w", pady=6)
        
        ttk.Label(info_frame, text="项目主页", font=("Microsoft YaHei UI", 10, "bold"), foreground="#777777").grid(row=1, column=0, sticky="w", padx=(0, 20), pady=6)
        link_label = ttk.Label(info_frame, text="github.com/hxlive/MacroMate", font=("Microsoft YaHei UI", 10), foreground="#0066CC", cursor="hand2")
        link_label.grid(row=1, column=1, sticky="w", pady=6)
        
        link_label.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/hxlive/MacroMate/"))
        link_label.bind("<Enter>", lambda e: link_label.config(font=("Microsoft YaHei UI", 10, "underline"), foreground="#0052A3"))
        link_label.bind("<Leave>", lambda e: link_label.config(font=("Microsoft YaHei UI", 10), foreground="#0066CC"))
        
        # ========== 分隔线 ==========
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=(0, 18))
        
        # ========== 底部：操作按钮区域 ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(button_frame, text="确  定", command=self.dialog.destroy, bootstyle="primary", width=18, padding=(15, 8)).pack(anchor="center")



