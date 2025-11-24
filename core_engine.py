# -*- coding: utf-8 -*-
# core_engine.py
# 描述：自动化宏的核心功能引擎
# 版本：1.53.2
# 变更：(修复#A) 移除了 LoopCacheManager 中未被使用的 next_iteration 方法和调用。
#       (修复#C) 添加了 UTF-8 编码声明。

import pyautogui
import time
from PIL import Image, ImageGrab, ImageStat
import re
import pyperclip
import os
import sys
from collections import defaultdict
import functools 

try:
    import pygetwindow as gw
    PYGETWINDOW_AVAILABLE = True
except ImportError:
    PYGETWINDOW_AVAILABLE = False
    print("[配置] ✗ 未找到 pygetwindow 库 (pip install pygetwindow)。'激活窗口' 功能将不可用。")

# ======================================================================
# 全局配置
# ======================================================================
FORCE_OCR_ENGINE = None 
ENABLE_GLOBAL_FALLBACK = True # 控制是否启用缓存失效后的全局搜索

# 导入 OCR 引擎
try:
    import ocr_engine
except ImportError:
    print("[严重错误] 未找到 'ocr_engine.py'。")
    class ocr_engine:
        def find_text_location(*args, **kwargs): return None
        WINOCR_AVAILABLE = False
        TESSERACT_AVAILABLE = False
        RAPIDOCR_AVAILABLE = False

# 导入 OpenCV
try:
    import cv2
    import numpy as np 
    OPENCV_AVAILABLE = True
    print("[配置] ✓ OpenCV 引擎就绪 (极速找图内核已启用)")
except ImportError:
    OPENCV_AVAILABLE = False
    print("[配置] ✗ 未找到 OpenCV。将回退到慢速找图模式。")

# ======================================================================
# 快捷键工具模块
# ======================================================================
class HotkeyUtils:
    """快捷键解析、验证工具类"""
    
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
    
    try:
        if sys.platform == 'win32':
            import win32con
            PYNPUT_MOD_TO_WIN_MOD = {
                'ctrl': win32con.MOD_CONTROL, 'alt': win32con.MOD_ALT,
                'shift': win32con.MOD_SHIFT, 'cmd': win32con.MOD_WIN,
            }
        else:
            PYNPUT_MOD_TO_WIN_MOD = {}
    except ImportError:
        PYNPUT_MOD_TO_WIN_MOD = {}
    
    @staticmethod
    def format_hotkey_display(hotkey_str):
        """格式化快捷键显示 (ctrl+f10 -> Ctrl+F10)"""
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
        except:
            return hotkey_str.upper()

# ======================================================================
# 宏定义元数据
# ======================================================================
class MacroSchema:
    """宏系统的元数据定义"""
    
    ACTION_TRANSLATIONS = {
        'FIND_IMAGE':     '01. 查找图像',
        'FIND_TEXT':      '02. 查找文本 (OCR)',
        'MOVE_OFFSET':    '03. 相对移动',
        'MOVE_TO':        '04. 移动到 (绝对坐标)',
        'CLICK':          '05. 点击鼠标',
        'SCROLL':         '06. 滚动滚轮',
        'WAIT':           '07. 等待',
        'TYPE_TEXT':      '08. 输入文本',
        'PRESS_KEY':      '09. 按下按键',
        'ACTIVATE_WINDOW':'10. 激活窗口 (按标题)',
        'IF_IMAGE_FOUND': '11. IF 找到图像',
        'IF_TEXT_FOUND':  '12. IF 找到文本',
        'ELSE':           '13. ELSE',
        'END_IF':         '14. END_IF',
        'LOOP_START':     '15. 循环开始 (Loop)',
        'END_LOOP':       '16. 结束循环 (EndLoop)',
    }
    ACTION_KEYS_TO_NAME = {v: k for k, v in ACTION_TRANSLATIONS.items()}
    
    LANG_OPTIONS = {'chi_sim (简体中文)': 'chi_sim', 'eng (英文)': 'eng'}
    LANG_VALUES_TO_NAME = {v: k for k, v in LANG_OPTIONS.items()}
    
    CLICK_OPTIONS = {'left (左键)': 'left', 'right (右键)': 'right', 'middle (中键)': 'middle'}
    CLICK_VALUES_TO_NAME = {v: k for k, v in CLICK_OPTIONS.items()}

# ======================================================================
# 性能监控
# ======================================================================
class PerformanceMonitor:
    def __init__(self): self.reset()
    def reset(self):
        self.image_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}
        self.ocr_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}
    def _get_stats_for(self, stats_dict):
        total = stats_dict['hits'] + stats_dict['misses']
        if total == 0: return "(无记录)"
        unique_hits = stats_dict['hits'] - stats_dict['loop_hits']
        total_valid = unique_hits + stats_dict['misses']
        hit_rate = (unique_hits / total_valid * 100) if total_valid > 0 else 0
        avg_ms = (sum(stats_dict['times']) / len(stats_dict['times']) * 1000) if stats_dict['times'] else 0
        return f"(命中{hit_rate:.0f}% | 循环{stats_dict['loop_hits']} | 均耗{avg_ms:.0f}ms)"
    def record_hit(self, is_loop, is_ocr):
        s = self.ocr_stats if is_ocr else self.image_stats
        s['hits'] += 1
        if is_loop: s['loop_hits'] += 1
    def record_miss(self, is_ocr): (self.ocr_stats if is_ocr else self.image_stats)['misses'] += 1
    def record_time(self, dt, is_ocr): (self.ocr_stats if is_ocr else self.image_stats)['times'].append(dt)
    def get_stats(self): return f"图像{self._get_stats_for(self.image_stats)} | OCR{self._get_stats_for(self.ocr_stats)}"

perf = PerformanceMonitor()

# ======================================================================
# 循环缓存管理器
# ======================================================================
class LoopCacheManager:
    def __init__(self): self.reset()
    
    def reset(self):
        """清空所有缓存和循环堆栈"""
        self.caches = {}
        self.stack = [] # 使用堆栈来管理嵌套循环
        
    def get_current_loop_id(self):
        """获取当前 (最内层) 循环的ID"""
        return self.stack[-1] if self.stack else None

    def enter(self, loop_id):
        """进入一个新循环 (压栈)"""
        if loop_id not in self.caches:
            self.caches[loop_id] = {}
        self.stack.append(loop_id)

    def exit(self):
        """退出一个循环 (弹栈)"""
        if self.stack:
            self.stack.pop()

    def clear_cache(self, loop_id):
        """显式清除指定循环的缓存 (当循环结束时)"""
        if loop_id in self.caches:
            del self.caches[loop_id]

    def get(self, sig): 
        """从当前循环获取缓存"""
        loop_id = self.get_current_loop_id()
        return self.caches.get(loop_id, {}).get(sig) if loop_id else None

    def set(self, sig, loc): 
        """向当前循环设置缓存"""
        loop_id = self.get_current_loop_id()
        if loop_id:
            if loop_id not in self.caches:
                 self.caches[loop_id] = {}
            self.caches[loop_id][sig] = loc

loop_cache = LoopCacheManager()

# ======================================================================
# 核心工具函数
# ======================================================================
def smart_screenshot(region=None):
    if region:
        pad = 0
        x = max(0, region[0] - pad)
        y = max(0, region[1] - pad)
        return ImageGrab.grab(bbox=(x, y, region[0]+region[2]+pad, region[1]+region[3]+pad)), (x, y)
    return ImageGrab.grab(), (0, 0)

SCALES = [1.0, 0.9, 1.1, 0.8, 1.2]
@functools.lru_cache(maxsize=100)
def _get_template(path, scale):
    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if img is None: return None, 0, 0
    if scale != 1.0:
        h, w = img.shape[:2]
        img = cv2.resize(img, (int(w*scale), int(h*scale)), interpolation=cv2.INTER_AREA)
    return img, img.shape[1], img.shape[0]

def find_image_cv2(path, conf, screenshot_pil, offset=(0,0)):
    if not OPENCV_AVAILABLE: return None
    try:
        t0 = time.time()
        screen_gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
        best = (-1, None, 0, 0)
        for scale in SCALES:
            tmpl, tw, th = _get_template(path, scale)
            if tmpl is None or th > screen_gray.shape[0] or tw > screen_gray.shape[1]: continue
            res = cv2.matchTemplate(screen_gray, tmpl, cv2.TM_CCOEFF_NORMED)
            min_v, max_v, min_l, max_l = cv2.minMaxLoc(res)
            if max_v > best[0]: best = (max_v, max_l, tw, th)
            if best[0] >= 0.95 and best[0] >= conf: break 
        val, loc, w, h = best
        if val >= conf and loc:
            cx, cy = offset[0] + loc[0] + w//2, offset[1] + loc[1] + h//2
            perf.record_time(time.time()-t0, False)
            return (cx, cy, w, h), val
    except Exception as e: print(f"CV2找图错误: {e}")
    return None

def quick_check_cv2(path, conf, screenshot_pil, offset, target_loc):
    if not OPENCV_AVAILABLE: return False
    try:
        tmpl, tw, th = _get_template(path, 1.0)
        if tmpl is None: return False
        pad_w, pad_h = tw//2 + 15, th//2 + 15
        rel_x, rel_y = target_loc[0] - offset[0], target_loc[1] - offset[1]
        l, t = max(0, rel_x - pad_w), max(0, rel_y - pad_h)
        r, b = min(screenshot_pil.width, rel_x + pad_w), min(screenshot_pil.height, rel_y + pad_h)
        if r <= l or b <= t: return False
        crop = cv2.cvtColor(np.array(screenshot_pil.crop((l, t, r, b))), cv2.COLOR_RGB2GRAY)
        _, max_v, _, _ = cv2.minMaxLoc(cv2.matchTemplate(crop, tmpl, cv2.TM_CCOEFF_NORMED))
        return max_v >= conf
    except: return False

# ======================================================================
# 主执行引擎
# ======================================================================
def execute_steps(steps, run_context=None, status_callback=None):
    print(f"\n--- 宏执行开始 (Core V1.53.2) ---")
    perf.reset(); loop_cache.reset()
    ctx = run_context if run_context else {}
    ctx.setdefault('last_pos', (None, None))
    ctx.setdefault('stop_requested', False)
    
    default_stop = "Ctrl+F11"
    try:
        s = run_context.get('stop_key_str', default_stop)
        stop_key_display = HotkeyUtils.format_hotkey_display(s)
    except:
        stop_key_display = default_stop
    
    pc, loops = 0, []
    try:
        while pc < len(steps):
            if ctx.get('stop_requested', False): 
                print(f"  [停止] 用户请求停止 ({stop_key_display})")
                break
                
            step = steps[pc]; act = step.get('action',''); p = step.get('params',{})
            print(f"[{pc+1}] {act}")
            next_pc = pc + 1

            try:
                if act.startswith('FIND_') or act.startswith('IF_'):
                    res = _handle_find(act, p, ctx, loop_cache.get_current_loop_id() is not None)
                    if act.startswith('IF_'):
                        if not res:
                            print("  -> IF条件不满足，跳过")
                            next_pc = _find_jump(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])
                    elif not res: print("  -> 没找到目标，宏停止"); break
                    if res: pyautogui.moveTo(res[0], res[1])
                
                elif act == 'CLICK':
                    btn = p.get('button', 'left').lower()
                    clicks = int(p.get('clicks', 1))
                    interval = float(p.get('interval', 0.0))
                    duration = float(p.get('duration', 0.0))
                    x = int(p['x']) if 'x' in p else None
                    y = int(p['y']) if 'y' in p else None
                    pyautogui.click(x=x, y=y, button=btn, clicks=clicks, interval=interval, duration=duration)
                    if x and y: ctx['last_pos'] = (x, y)
                
                elif act == 'MOVE_TO':
                    x, y = int(p['x']), int(p['y'])
                    pyautogui.moveTo(x, y, duration=float(p.get('duration', 0.25)))
                    ctx['last_pos'] = (x, y)
                
                elif act == 'MOVE_OFFSET':
                    if not ctx['last_pos'][0]: print("  [错误] 无上次坐标"); break
                    ox, oy = int(p['x_offset']), int(p['y_offset'])
                    pyautogui.move(ox, oy, duration=float(p.get('duration', 0.25)))
                    ctx['last_pos'] = (ctx['last_pos'][0]+ox, ctx['last_pos'][1]+oy)
                
                elif act == 'SCROLL':
                    clicks = int(p.get('amount', 0))
                    if 'x' in p and 'y' in p: pyautogui.moveTo(int(p['x']), int(p['y']))
                    pyautogui.scroll(clicks) 
                
                elif act == 'WAIT': time.sleep(int(p['ms']) / 1000.0)
                
                elif act == 'TYPE_TEXT':
                    interval = float(p.get('interval', 0.0))
                    if interval > 0: pyautogui.write(p['text'], interval=interval)
                    else: pyperclip.copy(p['text']); time.sleep(0.1); pyautogui.hotkey('ctrl', 'v')
                
                elif act == 'PRESS_KEY':
                    keys = p.get('key', '').lower().replace(' ', '').split('+')
                    if keys: pyautogui.hotkey(*keys)
                
                elif act == 'ACTIVATE_WINDOW':
                    if not PYGETWINDOW_AVAILABLE:
                        print("  [错误] pygetwindow 库未安装，无法激活窗口。")
                        break
                    title = p.get('title')
                    if not title:
                        print("  [错误] 未提供窗口标题。")
                        break
                    
                    try:
                        wins = gw.getWindowsWithTitle(title)
                        if not wins:
                            print(f"  [失败] 未找到标题包含 '{title}' 的窗口。")
                            break
                        
                        target_win = wins[0]
                        if target_win.isMinimized:
                            target_win.restore()
                        target_win.activate()
                        print(f"  [成功] 已激活窗口: {target_win.title}")
                        time.sleep(0.5) 
                    except Exception as e:
                        print(f"  [错误] 激活窗口时出错: {e}")
                        break

                elif act == 'ELSE': 
                    next_pc = _find_jump(steps, pc, 'IF_', 'END_IF', ['END_IF'])
                
                elif act == 'LOOP_START':
                    next_pc = _handle_loop_start(steps, pc, loops, p, ctx, status_callback)
                
                elif act == 'END_LOOP': 
                    next_pc = loops[-1]['start'] # 返回循环开始处

            except Exception as e:
                print(f"  [执行异常] {e}"); import traceback; traceback.print_exc(); break
            pc = next_pc
    finally:
        loop_cache.reset()
        print(f"--- 执行结束 ---\n[统计] {perf.get_stats()}\n")

def _handle_find(act, p, ctx, in_loop):
    is_img = 'IMAGE' in act
    final_engine = FORCE_OCR_ENGINE if (FORCE_OCR_ENGINE and FORCE_OCR_ENGINE != 'auto') else p.get('engine', 'auto')
    
    region = None
    if 'cache_box' in p:
        cb = p['cache_box']
        if isinstance(cb, list) and len(cb) == 2: cb = [cb[0], cb[1], cb[0], cb[1]]; p['cache_box'] = cb
        if isinstance(cb, list) and len(cb) >= 4:
            pad = 50 
            region = (max(0, cb[0]-pad), max(0, cb[1]-pad), (cb[2]-cb[0])+pad*2, (cb[3]-cb[1])+pad*2)
        else:
            if 'cache_box' in p: del p['cache_box']

    ss, offset = smart_screenshot(region)
    sig = f"{act}_{p.get('path', p.get('text',''))}"

    if in_loop:
        cached = loop_cache.get(sig)
        if cached and is_img and quick_check_cv2(p['path'], float(p.get('confidence',0.8)), ss, offset, cached):
            perf.record_hit(True, False); print(f"  [Loop缓存] {cached}"); ctx['last_pos'] = cached; return cached

    res = _do_find(is_img, p, ss, offset, final_engine)
    
    if not res and region and ENABLE_GLOBAL_FALLBACK:
        print("  [缓存失效] 全局搜索...")
        ss, offset = smart_screenshot(None)
        res = _do_find(is_img, p, ss, offset, final_engine)
        if res:
            w, h = (res[2], res[3]) if len(res) == 4 else (0, 0)
            p['cache_box'] = [res[0]-w//2, res[1]-h//2, res[0]+w//2, res[1]+h//2]

    if res:
        pos = (res[0], res[1])
        if in_loop: loop_cache.set(sig, pos)
        ctx['last_pos'] = pos
        return pos
    
    perf.record_miss(not is_img)
    return None

def _do_find(is_img, p, ss, offset, engine='auto'):
    if is_img:
        res_val = find_image_cv2(p['path'], float(p.get('confidence', 0.8)), ss, offset)
        if res_val:
            perf.record_hit(False, False); print(f"  [找到] 图 ({res_val[0][0]},{res_val[0][1]})")
            return res_val[0]
    else:
        res = ocr_engine.find_text_location(p['text'], p.get('lang','eng'), p.get('debug',True), ss, offset, engine)
        if res:
            perf.record_hit(False, True)
            return res
    return None

def _handle_loop_start(steps, pc, loops, p, ctx, cb):
    top = loops[-1] if loops else None
    if top and top['start'] == pc:
        top['remain'] -= 1
        if top['remain'] > 0:
            if cb: cb(f"循环剩余: {top['remain']}")
            return pc + 1
        else:
            loop_id_to_exit = loops.pop()['id']
            loop_cache.exit() # 弹栈
            loop_cache.clear_cache(loop_id_to_exit) # 清理缓存
            return _find_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
    else:
        count = int(p.get('times', 1))
        if count <= 0: 
            return _find_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
        
        loop_id = f"L{pc}_{len(loops)}" 
        loops.append({'start': pc, 'remain': count, 'id': loop_id})
        loop_cache.enter(loop_id) # 压栈
        if cb: cb(f"循环剩余: {count}")
        return pc + 1

def _find_jump(steps, start, open_tag, close_tag, targets):
    lvl = 0
    for i in range(start + 1, len(steps)):
        a = steps[i].get('action','')
        if a.startswith(open_tag.rstrip('_')): lvl += 1
        elif a == close_tag:
            if lvl == 0 and a in targets: return i + 1
            lvl -= 1
        elif lvl == 0 and a in targets: return i + 1
    return len(steps)

core_engine_version = f"1.53.2 (Core) / OpenCV: {OPENCV_AVAILABLE}"
