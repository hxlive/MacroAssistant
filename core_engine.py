# core_engine.py
# 描述：自动化宏的核心功能引擎
# 版本：1.49.1

import pyautogui
import time
from PIL import Image, ImageGrab, ImageStat
import re
import pyperclip
import os
import sys
from collections import defaultdict
import functools 

# ======================================================================
# 全局配置
# ======================================================================
FORCE_OCR_ENGINE = None

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

# 导入 OpenCV (用于高速找图)
try:
    import cv2
    import numpy as np 
    OPENCV_AVAILABLE = True
    print("[配置] ✓ OpenCV 引擎就绪 (极速找图内核已启用)")
except ImportError:
    OPENCV_AVAILABLE = False
    print("[配置] ✗ 未找到 OpenCV。将回退到慢速找图模式。")

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
        self.caches = {}; self.cur_loop = None; self.loop_iteration = 0
    def enter(self, loop_id):
        self.cur_loop = loop_id; self.loop_iteration = 0
        if loop_id not in self.caches: self.caches[loop_id] = {}
    def next_iteration(self): self.loop_iteration += 1
    def exit(self):
        if self.cur_loop in self.caches: del self.caches[self.cur_loop]
        self.cur_loop = None; self.loop_iteration = 0
    def get(self, sig): return self.caches.get(self.cur_loop, {}).get(sig) if self.cur_loop else None
    def set(self, sig, loc): 
        if self.cur_loop: 
            if self.cur_loop not in self.caches: self.caches[self.cur_loop] = {}
            self.caches[self.cur_loop][sig] = loc

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
    print(f"\n--- 宏执行开始 (Core V1.49.1) ---")
    perf.reset(); loop_cache.reset()
    ctx = run_context if run_context else {}
    ctx.setdefault('last_pos', (None, None))
    ctx.setdefault('stop_requested', False)
    
    pc, loops = 0, []
    try:
        while pc < len(steps):
            if ctx.get('stop_requested', False): 
                print("  [停止] 用户请求停止 (Ctrl+F11)"); break
                
            step = steps[pc]; act = step.get('action',''); p = step.get('params',{})
            print(f"[{pc+1}] {act}")
            next_pc = pc + 1

            try:
                if act.startswith('FIND_') or act.startswith('IF_'):
                    res = _handle_find(act, p, ctx, loop_cache.cur_loop is not None)
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
                
                elif act == 'ELSE': 
                    next_pc = _find_jump(steps, pc, 'IF_', 'END_IF', ['END_IF'])
                
                elif act == 'LOOP_START':
                    next_pc, is_looping = _handle_loop_start(steps, pc, loops, p, ctx, status_callback)
                    if not is_looping: loop_cache.exit() 
                
                elif act == 'END_LOOP': 
                    loop_cache.next_iteration()
                    next_pc = loops[-1]['start']

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
    
    if not res and region:
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
            return pc + 1, True
        else:
            loops.pop()
            return _find_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP']), False
    else:
        count = int(p.get('times', 1))
        if count <= 0: return _find_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP']), False
        loops.append({'start': pc, 'remain': count})
        loop_cache.enter(f"L{pc}")
        return pc + 1, True

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

macro_engine_version = f"1.49.1 (Core) / OpenCV: {OPENCV_AVAILABLE}"