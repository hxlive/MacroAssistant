# sys_utils.py
# 描述: 系统底层工具、全局热键管理及稳定工具类集
# 版本: 1.8.3

import sys
import os
import threading
import queue
import functools
import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import ctypes
from PIL import Image, ImageGrab, ImageTk
from pynput import keyboard

import logging
logger = logging.getLogger(__name__)


_SHARED_FILE_LOCKS = {}
_SHARED_FILE_LOCKS_GUARD = threading.Lock()

def get_shared_file_lock(path):
    """Return the process-wide lock used for read-modify-write file transactions."""
    normalized = os.path.normcase(os.path.abspath(os.fspath(path)))
    with _SHARED_FILE_LOCKS_GUARD:
        return _SHARED_FILE_LOCKS.setdefault(normalized, threading.RLock())

class HotkeyUtils:
    PYNPUT_TO_VK = {
        'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74, 'f6': 0x75,
        'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
        'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45, 'f': 0x46, 'g': 0x47,
        'h': 0x48, 'i': 0x49, 'j': 0x4A, 'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E,
        'o': 0x4F, 'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54, 'u': 0x55,
        'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
        '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35, '6': 0x36,
        '7': 0x37, '8': 0x38, '9': 0x39,
        'enter': 0x0D, 'space': 0x20, 'tab': 0x09, 'caps_lock': 0x14,
        'esc': 0x1B, 'page_up': 0x21, 'page_down': 0x22, 'end': 0x23, 'home': 0x24,
        'left': 0x25, 'up': 0x26, 'right': 0x27, 'down': 0x28, 'insert': 0x2D, 'delete': 0x2E,
        'backspace': 0x08,
    }
    VK_TO_PYNPUT = {v: k for k, v in PYNPUT_TO_VK.items()}
    
    if sys.platform == 'win32':
        PYNPUT_MOD_TO_WIN_MOD = {
            'ctrl': 0x0002,  # win32con.MOD_CONTROL
            'alt': 0x0001,   # win32con.MOD_ALT
            'shift': 0x0004, # win32con.MOD_SHIFT
            'cmd': 0x0008,   # win32con.MOD_WIN
        }
    else:
        PYNPUT_MOD_TO_WIN_MOD = {}
    
    @staticmethod
    def format_hotkey_display(hotkey_str):
        if not hotkey_str or "录制" in hotkey_str:
            return hotkey_str
        try:
            parts = hotkey_str.split('+')
            display_parts = []
            for part in parts:
                if part.lower() in {'ctrl', 'alt', 'shift', 'cmd'}:
                    display_parts.append(part.capitalize())
                else:
                    display_parts.append(part.upper())
            return "+".join(display_parts)
        except Exception:
            return hotkey_str.upper()


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
            logger.info(f"STDIO encoding: {stdio_encoding}")
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
            logger.warning(f"设置 AppUserModelID 失败: {e}")
    return False

def _get_virtual_screen_rect():
    """Return the physical virtual-screen rectangle as (x, y, width, height)."""
    if sys.platform != 'win32':
        return (0, 0, 0, 0)

    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    user32 = ctypes.windll.user32
    return (
        user32.GetSystemMetrics(SM_XVIRTUALSCREEN),
        user32.GetSystemMetrics(SM_YVIRTUALSCREEN),
        user32.GetSystemMetrics(SM_CXVIRTUALSCREEN),
        user32.GetSystemMetrics(SM_CYVIRTUALSCREEN),
    )


def capture_physical_bbox(bbox=None):
    """Capture a physical screen bbox and return (PIL.Image, absolute_offset)."""
    if bbox is not None:
        if len(bbox) != 4:
            raise ValueError("bbox must be a 4-item (x1, y1, x2, y2) tuple")
        x1, y1, x2, y2 = (int(v) for v in bbox)
        if x2 <= x1 or y2 <= y1:
            raise ValueError("Invalid crop box geometry")
        bbox = (x1, y1, x2, y2)

    if sys.platform != 'win32':
        if bbox is None:
            return ImageGrab.grab(), (0, 0)
        return ImageGrab.grab(bbox=bbox), (bbox[0], bbox[1])

    vx, vy, vw, vh = _get_virtual_screen_rect()
    user32 = ctypes.windll.user32

    if bbox is None:
        primary_w = user32.GetSystemMetrics(0)
        primary_h = user32.GetSystemMetrics(1)
        if vw > 0 and vh > 0 and (vx != 0 or vy != 0 or vw > primary_w or vh > primary_h):
            return ImageGrab.grab(all_screens=True), (vx, vy)
        return ImageGrab.grab(), (0, 0)

    x1, y1, x2, y2 = bbox
    screen_w = user32.GetSystemMetrics(0)
    screen_h = user32.GetSystemMetrics(1)
    needs_virtual_capture = x1 < 0 or y1 < 0 or x2 > screen_w or y2 > screen_h or vx < 0 or vy < 0

    if not needs_virtual_capture:
        return ImageGrab.grab(bbox=bbox), (x1, y1)

    full_screen = ImageGrab.grab(all_screens=True)
    try:
        crop_x1 = max(0, x1 - vx)
        crop_y1 = max(0, y1 - vy)
        crop_x2 = min(full_screen.width, x2 - vx)
        crop_y2 = min(full_screen.height, y2 - vy)
        if crop_x2 <= crop_x1 or crop_y2 <= crop_y1:
            raise ValueError("Invalid crop box geometry")
        return full_screen.crop((crop_x1, crop_y1, crop_x2, crop_y2)), (crop_x1 + vx, crop_y1 + vy)
    finally:
        try:
            full_screen.close()
        except Exception:
            pass
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
def center_child_window(parent, dialog):
    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - dialog.winfo_width()) // 2
    y = parent.winfo_y() + (parent.winfo_height() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")

# Backward-compatible private alias for existing system dialogs.
_center_child_window = center_child_window

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
                x_val, y_val, w_val, h_val = _get_virtual_screen_rect()
                if w_val > 0 and h_val > 0:
                    self.offset_x, self.offset_y, w, h = x_val, y_val, w_val, h_val
            except Exception as e:
                logger.warning("获取多屏几何信息失败: %s", e)

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

def _region_preview_geometry(region, virtual_rect):
    """Map an absolute region to clipped canvas coordinates on the virtual desktop."""
    if not isinstance(region, (list, tuple)) or len(region) != 4:
        raise ValueError("region must contain x1, y1, x2, y2")

    x1, y1, x2, y2 = (int(value) for value in region)
    vx, vy, vw, vh = (int(value) for value in virtual_rect)
    if x2 <= x1 or y2 <= y1:
        raise ValueError("region must satisfy x2 > x1 and y2 > y1")
    if vw <= 0 or vh <= 0:
        raise ValueError("virtual screen size is invalid")

    left = max(x1, vx)
    top = max(y1, vy)
    right = min(x2, vx + vw)
    bottom = min(y2, vy + vh)
    if right <= left or bottom <= top:
        raise ValueError("region is outside the current virtual screen")

    return (left - vx, top - vy, right - vx, bottom - vy)


def _region_preview_border_bounds(region, virtual_rect, thickness=4):
    """Return absolute bounds for four border-only preview windows."""
    x1, y1, x2, y2 = _region_preview_geometry(region, virtual_rect)
    vx, vy, _vw, _vh = (int(value) for value in virtual_rect)
    left, top = vx + x1, vy + y1
    right, bottom = vx + x2, vy + y2
    width, height = right - left, bottom - top
    border = max(1, min(int(thickness), width, height))
    return (
        (left, top, width, border),
        (left, bottom - border, width, border),
        (left, top, border, height),
        (right - border, top, border, height),
    )


def _set_preview_window_bounds(window, bounds):
    """Place a border window using physical coordinates, including negative monitors."""
    x, y, width, height = (int(value) for value in bounds)
    x_pos = f'+{x}' if x >= 0 else str(x)
    y_pos = f'+{y}' if y >= 0 else str(y)
    window.geometry(f'{max(1, width)}x{max(1, height)}{x_pos}{y_pos}')

    if sys.platform != 'win32':
        return
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        ctypes.windll.user32.SetWindowPos(
            hwnd,
            -1,
            x,
            y,
            max(1, width),
            max(1, height),
            0x0010 | 0x0040,
        )
    except (AttributeError, OSError, tk.TclError):
        pass


class RegionPreviewOverlay:
    """Show only a red border and coordinates over the unchanged desktop."""

    BORDER_COLOR = '#ff2d2d'
    BORDER_WIDTH = 4

    def __init__(self, master, region, duration_ms=1500):
        self.master = master
        self.region = tuple(int(value) for value in region)
        self.duration_ms = max(250, int(duration_ms))
        self.windows = []

        virtual_rect = _get_virtual_screen_rect()
        border_bounds = _region_preview_border_bounds(
            self.region,
            virtual_rect,
            self.BORDER_WIDTH,
        )
        for bounds in border_bounds:
            self._create_border_window(bounds)

        self._create_coordinate_labels(border_bounds[0], border_bounds[1], virtual_rect)
        self.top = self.windows[0]
        self.top.after(self.duration_ms, self.close)

    def _new_window(self, background):
        window = tk.Toplevel(self.master)
        window.withdraw()
        window.overrideredirect(True)
        window.configure(bg=background)
        window.attributes('-topmost', True)
        window.bind('<Escape>', self.close)
        window.bind('<Button-1>', self.close)
        self.windows.append(window)
        return window

    def _create_border_window(self, bounds):
        window = self._new_window(self.BORDER_COLOR)
        _set_preview_window_bounds(window, bounds)
        window.deiconify()

    def _create_coordinate_labels(self, top_border, bottom_border, virtual_rect):
        left, top, width, border = top_border
        _bottom_left, bottom_top, _bottom_width, _bottom_border = bottom_border
        self._create_coordinate_label(
            f'左上 ({self.region[0]}, {self.region[1]})',
            left + border,
            top + border,
            virtual_rect,
        )
        self._create_coordinate_label(
            f'右下 ({self.region[2]}, {self.region[3]})',
            left + width - border,
            bottom_top,
            virtual_rect,
            align_right=True,
            align_bottom=True,
        )

    def _create_coordinate_label(
        self,
        text,
        anchor_x,
        anchor_y,
        virtual_rect,
        align_right=False,
        align_bottom=False,
    ):
        vx, vy, vw, vh = virtual_rect
        label_window = self._new_window(self.BORDER_COLOR)
        label = tk.Label(
            label_window,
            text=text,
            bg=self.BORDER_COLOR,
            fg='white',
            font=('Consolas', 10, 'bold'),
            padx=6,
            pady=2,
        )
        label.pack()
        label_window.update_idletasks()
        width = max(1, label_window.winfo_reqwidth())
        height = max(1, label_window.winfo_reqheight())
        label_x = anchor_x - width if align_right else anchor_x
        label_y = anchor_y - height if align_bottom else anchor_y
        label_x = max(vx, min(label_x, vx + vw - width))
        label_y = max(vy, min(label_y, vy + vh - height))
        _set_preview_window_bounds(label_window, (label_x, label_y, width, height))
        label_window.deiconify()

    def close(self, _event=None):
        windows, self.windows = self.windows, []
        for window in windows:
            try:
                if window.winfo_exists():
                    window.destroy()
            except tk.TclError:
                pass

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
        self._callback_queue = queue.Queue()
        self._callback_pump_active = False
        self._callback_pump_after_id = None
        
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
                    logger.error(f"stop old listener failed: {e}")
            if self.listener is old_listener:
                self.listener = None
            self.held_keys.clear()
            threading.Thread(target=self._listener_thread, args=(generation,), daemon=True).start()
        self._start_callback_pump()
        
    def _start_callback_pump(self):
        if self._callback_pump_active:
            return
        self._callback_pump_active = True
        self._schedule_callback_drain()

    def _schedule_callback_drain(self):
        try:
            self._callback_pump_after_id = self.root.after(50, self._drain_callbacks)
        except Exception:
            self._callback_pump_active = False
            logger.exception("schedule callback drain failed")

    def _enqueue_callback(self, callback):
        self._callback_queue.put(callback)

    def _drain_callbacks(self):
        while True:
            try:
                callback = self._callback_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                logger.exception("callback failed")
        if self._callback_pump_active:
            self._schedule_callback_drain()
        else:
            self._callback_pump_after_id = None

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
            self._enqueue_callback(lambda msg=msg: messagebox.showerror("严重错误", msg))
            
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
                logger.error(f"check physical keys failed: {e}")
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
                self._enqueue_callback(self.trigger_run)
            stop_mods, stop_key = self._parse_hotkey(self.stop_hotkey_cache)
            if key_name == stop_key and self._modifiers_satisfied(stop_mods):
                self._enqueue_callback(self.trigger_stop)
        except Exception as e:
            logger.error(f"press error: {e}")

    def on_release(self, key, generation=None):
        try:
            if generation is not None and generation != self._listener_generation:
                return
            key_name = self._normalize_key(self._get_key_name(key))
            if not key_name: return
            # [优化] 松开按键时直接清空字典中该键的状态，根治所有残留假死
            self.held_keys.pop(key_name, None)
        except Exception as e:
            logger.error(f"release error: {e}")

    @functools.lru_cache(maxsize=16)
    def _parse_hotkey(self, hotkey_str):
        if not hotkey_str: return set(), ""
        parts = [p.strip() for p in hotkey_str.lower().split('+')]
        if not parts: return set(), ""
        return set(parts[:-1]), parts[-1]

    def check_conflicts(self, show_success=True):
        if not HOTKEY_CHECK_AVAILABLE: return True
        conflicts = []
        unavailable = []
        run_str = self.get_run_str()
        run_result = self._test_register(run_str, 1)
        if run_result is False:
            conflicts.append(f"运行快捷键 '{HotkeyUtils.format_hotkey_display(run_str)}'")
        elif run_result is None:
            unavailable.append(f"运行快捷键 '{HotkeyUtils.format_hotkey_display(run_str)}'")
        stop_str = self.get_stop_str()
        stop_result = self._test_register(stop_str, 2)
        if stop_result is False:
            conflicts.append(f"停止快捷键 '{HotkeyUtils.format_hotkey_display(stop_str)}'")
        elif stop_result is None:
            unavailable.append(f"停止快捷键 '{HotkeyUtils.format_hotkey_display(stop_str)}'")
        if conflicts:
            msg = "检测到快捷键冲突：\n\n" + "\n".join(conflicts) + "\n\n可能已被其他程序占用。\n请修改快捷键，否则热键可能无法工作。"
            if unavailable:
                msg += "\n\n另有快捷键无法完成占用检测：\n" + "\n".join(unavailable)
            self.root.after(0, messagebox.showwarning, "快捷键冲突", msg)
            return False
        if unavailable:
            msg = "无法完成以下快捷键的占用检测：\n\n" + "\n".join(unavailable) + "\n\n快捷键仍会保存并尝试启用。"
            self.root.after(0, messagebox.showwarning, "快捷键检测不可用", msg)
        return True

    def _test_register(self, hotkey_str, hotkey_id):
        if not hotkey_str: return True
        try:
            parts = hotkey_str.lower().split('+')
            modifiers, vk = 0, None
            for part in [p.strip() for p in parts]:
                if part in HotkeyUtils.PYNPUT_MOD_TO_WIN_MOD: modifiers |= HotkeyUtils.PYNPUT_MOD_TO_WIN_MOD[part]
                elif part in HotkeyUtils.PYNPUT_TO_VK: vk = HotkeyUtils.PYNPUT_TO_VK[part]
            if vk is None: return None
            if ctypes.windll.user32.RegisterHotKey(None, hotkey_id, modifiers, vk) == 0: return False
            ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
            return True
        except Exception as e:
            logger.error(f"conflict check unavailable for '{hotkey_str}': {e}")
            return None

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
# 7. 快捷键设置对话框 (Hotkey)
# ======================================================================
class HotkeySettingsDialog:
    # [修复 BUG-4] 恢复快捷键格式校验；默认值修正为 ctrl+f1/ctrl+f2
    def __init__(self, parent, run_hotkey, stop_hotkey,
                 default_run='ctrl+f1', default_stop='ctrl+f2'):
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
            logger.error(f"图片提示加载失败: {e}")

    def _hide_tooltip(self):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

def _calculate_mini_status_position(pointer, fallback_rect, virtual_rect, window_height):
    """Return the mini window's lower-left position for the active desktop bounds."""
    selected = fallback_rect
    if virtual_rect is not None:
        vx, vy, vw, vh = virtual_rect
        px, py = pointer
        if vw > 0 and vh > 0 and vx <= px < vx + vw and vy <= py < vy + vh:
            selected = virtual_rect

    monitor_x, monitor_y, _monitor_w, monitor_h = selected
    return monitor_x + 10, monitor_y + max(0, monitor_h - window_height - 50)

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
        pointer = self.window.winfo_pointerxy()
        screen_height = self.window.winfo_screenheight()
        fallback_rect = (
            self.window.winfo_vrootx(),
            self.window.winfo_vrooty(),
            self.window.winfo_vrootwidth(),
            self.window.winfo_vrootheight() or screen_height,
        )
        virtual_rect = _get_virtual_screen_rect() if sys.platform == 'win32' else None
        x, y = _calculate_mini_status_position(
            pointer,
            fallback_rect,
            virtual_rect,
            window_height,
        )
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
                logger.warning(f"设置关于对话框图标失败: {e}")
                
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
                logger.warning(f"加载图标图像失败: {e}")
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

