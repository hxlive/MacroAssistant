# core_engine.py
# 描述：自动化宏的核心功能引擎 (V45.1 - 多尺度重构版)
# 优化内容:
# 1. (Grok P1) find_image_location 和 _quick_image_check_V44
#    已重构，支持多尺度(Multi-Scale)图像匹配，提高鲁棒性。
# 2. (Grok P3, P5) 保持 V45.0 的性能监控粒度和缓存清理修复。

import pyautogui
import time
from PIL import Image, ImageGrab, ImageStat
import re
import pyperclip
import os
import sys
from collections import defaultdict
import functools # 【V45.0 修复】 导入 functools 用于 LRU 缓存

# ======================================================================
# 【V43.2】全局 OCR 调试开关 (保持不变)
# ======================================================================
FORCE_OCR_ENGINE = None
# --- 调试示例 ---
# FORCE_OCR_ENGINE = 'rapidocr'
# ======================================================================


# 【V43】导入 OCR 模块
try:
    import ocr_engine
except ImportError:
    print("[严重错误] 未找到 'ocr_engine.py'。")
    class ocr_engine:
        def find_text_location(*args, **kwargs): return None
        WINOCR_AVAILABLE = False
        TESSERACT_AVAILABLE = False
        RAPIDOCR_AVAILABLE = False

# 【V43】OpenCV 检测
try:
    import cv2
    import numpy as np # cv2 强依赖 numpy
    OPENCV_AVAILABLE = True
    print("[配置] ✓ OpenCV 引擎就绪 (图像查找 `confidence` 已启用)")
except ImportError:
    OPENCV_AVAILABLE = False
    print("[配置] ✗ 未找到 OpenCV。`FIND_IMAGE` (模糊查找) 将禁用。")

# ======================================================================
# 【V45.0 修复】性能监控粒度 (Grok P3)
# ======================================================================
class PerformanceMonitor:
    def __init__(self): self.reset()

    def reset(self):
        self.image_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}
        self.ocr_stats = {'hits': 0, 'misses': 0, 'times': [], 'loop_hits': 0}

    def _get_stats_for(self, stats_dict):
        total_finds = stats_dict['hits'] + stats_dict['misses']
        if total_finds == 0:
            return " (0 查找)"
        total_cacheable = stats_dict['hits'] - stats_dict['loop_hits'] + stats_dict['misses']
        hit_rate = 0
        if total_cacheable > 0:
            hit_rate = ((stats_dict['hits'] - stats_dict['loop_hits']) / total_cacheable) * 100
        avg_time = (sum(stats_dict['times']) / len(stats_dict['times'])) * 1000 if stats_dict['times'] else 0
        return (f" (命中: {hit_rate:.0f}% [{stats_dict['hits'] - stats_dict['loop_hits']}/{total_cacheable}] | "
                f"循环: {stats_dict['loop_hits']} | 均耗: {avg_time:.0f}ms)")

    def record_cache_hit(self, is_loop_cache=False, is_ocr=False):
        stats = self.ocr_stats if is_ocr else self.image_stats
        stats['hits'] += 1
        if is_loop_cache:
            stats['loop_hits'] += 1
    
    def record_cache_miss(self, is_ocr=False):
        stats = self.ocr_stats if is_ocr else self.image_stats
        stats['misses'] += 1

    def record_search_time(self, duration, is_ocr=False):
        stats = self.ocr_stats if is_ocr else self.image_stats
        stats['times'].append(duration)

    def get_stats(self):
        img_stats = self._get_stats_for(self.image_stats)
        ocr_stats = self._get_stats_for(self.ocr_stats)
        return f"图像:{img_stats} | OCR:{ocr_stats}"

perf_monitor = PerformanceMonitor()

# ======================================================================
# 【V44 修复】循环缓存管理器 (保持不变)
# ======================================================================
class LoopCacheManager:
    def __init__(self):
        self.loop_caches = {}; self.current_loop_id = None
        self.loop_iteration = 0
    def reset(self):
        self.loop_caches = {}; self.current_loop_id = None
        self.loop_iteration = 0
    def enter_loop(self, loop_id):
        self.current_loop_id = loop_id; self.loop_iteration = 0
        if loop_id not in self.loop_caches: self.loop_caches[loop_id] = {}
    def next_iteration(self): self.loop_iteration += 1
    def exit_loop(self):
        if self.current_loop_id in self.loop_caches:
            del self.loop_caches[self.current_loop_id]
        self.current_loop_id = None; self.loop_iteration = 0
    def get_cache(self, step_signature):
        if self.current_loop_id is None or self.loop_iteration == 0: return None
        cache = self.loop_caches.get(self.current_loop_id, {})
        return cache.get(step_signature)
    def set_cache(self, step_signature, location):
        if self.current_loop_id is None: return
        if self.current_loop_id not in self.loop_caches:
            self.loop_caches[self.current_loop_id] = {}
        self.loop_caches[self.current_loop_id][step_signature] = location

loop_cache = LoopCacheManager()

# ======================================================================
# 【V44】智能区域截图 (保持不变)
# ======================================================================
def smart_screenshot(region=None, expand_ratio=1.5):
    if region:
        x, y, w, h = region
        expand_w = int(w * expand_ratio); expand_h = int(h * expand_ratio)
        offset_x = (expand_w - w) // 2; offset_y = (expand_h - h) // 2
        screen_w, screen_h = pyautogui.size()
        new_x = max(0, x - offset_x); new_y = max(0, y - offset_y)
        new_w = min(expand_w, screen_w - new_x); new_h = min(expand_h, screen_h - new_y)
        return ImageGrab.grab(bbox=(new_x, new_y, new_x + new_w, new_y + new_h)), (new_x, new_y)
    else:
        return ImageGrab.grab(), (0, 0)

# ======================================================================
# 【V45.1 修复】图像查找 (Grok P1 - 多尺度)
# ======================================================================
# (P1) 定义缩放比例，1.0 (原图) 优先
SCALES = [1.0, 0.9, 1.1, 0.8, 1.2] 

@functools.lru_cache(maxsize=150) # 缓存 50 张图 * 3 个尺度 = 150
def _load_and_scale_template(path, scale):
    """(V45.1) 带缓存的模板加载和缩放"""
    template = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print(f"  [错误] 无法读取模板图像: {path}")
        return None, (0, 0)
    
    if scale == 1.0:
        return template, template.shape[::-1] # (W, H)
    
    # 缩放
    tH, tW = template.shape[:2]
    new_w = int(tW * scale)
    new_h = int(tH * scale)
    if new_w <= 0 or new_h <= 0:
        return None, (0, 0) # 缩放后太小
        
    scaled_template = cv2.resize(template, (new_w, new_h), interpolation=cv2.INTER_AREA)
    return scaled_template, (new_w, new_h)

def find_image_location(image_path, confidence=0.8, 
                        screenshot=None, offset=(0,0), 
                        region=None, 
                        use_loop_cache=False, step_signature=None):
    """
    V45.1: 使用 cv2.matchTemplate 和多尺度匹配。
    """
    start_time = time.time()
    
    if not OPENCV_AVAILABLE:
        print("  [错误] 图像查找失败: OpenCV 未安装。")
        return None

    try:
        # 1. 获取要搜索的图像 (V44 vs V43)
        search_image_pil = None
        search_offset = (0,0)
        
        if screenshot is not None:
            # V44 架构: 使用传入的截图
            search_image_pil = screenshot
            search_offset = offset
        else:
            # V43 兼容/测试模式: 自己截图
            print("  [警告] find_image_location 被直接调用 (无截图)，回退到全屏 locateOnScreen 模式。")
            search_image_pil = ImageGrab.grab(bbox=region)
            search_offset = (region[0], region[1]) if region else (0,0)

        search_image_gray = cv2.cvtColor(np.array(search_image_pil), cv2.COLOR_RGB2GRAY)
        
        # 2. 【V45.1 修复】多尺度匹配
        best_val = -1.0
        best_loc = None
        best_size = (0,0)
        
        for scale in SCALES:
            template_gray, (tW, tH) = _load_and_scale_template(image_path, scale)
            if template_gray is None: continue
            
            # 确保模板不会大于搜索图像
            if tH > search_image_gray.shape[0] or tW > search_image_gray.shape[1]:
                continue # 模板比截图还大，跳过

            # 3. 执行模板匹配
            result = cv2.matchTemplate(search_image_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val > best_val and max_val >= confidence:
                best_val = max_val
                best_loc = max_loc
                best_size = (tW, tH)
        
        # 4. 处理结果
        if best_loc is not None:
            (startX, startY) = best_loc
            (tW, tH) = best_size
            
            abs_center_x = search_offset[0] + startX + tW // 2
            abs_center_y = search_offset[1] + startY + tH // 2
            
            result_tuple = (abs_center_x, abs_center_y, tW, tH)
            
            if use_loop_cache and step_signature:
                loop_cache.set_cache(step_signature, (abs_center_x, abs_center_y))
            
            perf_monitor.record_cache_miss(is_ocr=False) 
            print(f"  [查找✓] 图像于 ({abs_center_x}, {abs_center_y}) (C:{best_val:.2f})")
            
            perf_monitor.record_search_time(time.time() - start_time, is_ocr=False)
            return result_tuple
        
        perf_monitor.record_cache_miss(is_ocr=False)
        perf_monitor.record_search_time(time.time() - start_time, is_ocr=False)
        return None
        
    except Exception as e:
        print(f"  [错误] 图像查找 (cv2-multi): {e}")
        import traceback
        traceback.print_exc()
        perf_monitor.record_search_time(time.time() - start_time, is_ocr=False)
        return None

# ======================================================================
# 【V45.1 修复】_quick_image_check (Grok P1/P2 - 多尺度/动态区域)
# ======================================================================
def _quick_image_check_V44(image_path, confidence, 
                          screenshot, offset, cached_loc):
    """
    V45.1: 在已有的 'screenshot' (PIL Image) 上验证 'cached_loc'。
    使用 cv2.matchTemplate 和多尺度/动态区域。
    """
    if not OPENCV_AVAILABLE: return False
    
    try:
        # 1. 加载 1.0 尺度的模板以获取基础尺寸
        template_gray_base, (tW, tH) = _load_and_scale_template(image_path, 1.0)
        if template_gray_base is None: return False

        # 2. 计算缓存点在截图中的相对坐标
        rel_x = cached_loc[0] - offset[0]
        rel_y = cached_loc[1] - offset[1]
        
        # 3.【Grok P2 修复】根据模板大小创建动态验证区域
        padding_x = tW // 2 + 20 
        padding_y = tH // 2 + 20
        img_w, img_h = screenshot.size
        
        rel_region_x1 = max(0, rel_x - padding_x)
        rel_region_y1 = max(0, rel_y - padding_y)
        rel_region_x2 = min(img_w, rel_x + padding_x)
        rel_region_y2 = min(img_h, rel_y + padding_y)
        
        search_image_pil_crop = screenshot.crop((rel_region_x1, rel_region_y1, rel_region_x2, rel_region_y2))
        search_image_crop_gray = cv2.cvtColor(np.array(search_image_pil_crop), cv2.COLOR_RGB2GRAY)
        
        # 4. 【Grok P1 修复】在小区域上进行多尺度匹配
        for scale in SCALES:
            scaled_template, (stW, stH) = _load_and_scale_template(image_path, scale)
            if scaled_template is None: continue
            
            if stH > search_image_crop_gray.shape[0] or stW > search_image_crop_gray.shape[1]:
                continue # 模板比截图还大

            result = cv2.matchTemplate(search_image_crop_gray, scaled_template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

            if max_val >= confidence:
                return True # 只要任一尺度匹配成功，即视为命中

        return False # 所有尺度均未匹配
        
    except Exception as e:
        print(f"  [错误] _quick_image_check_V44 失败: {e}")
        return False

# ======================================================================
# 【V45.0 修复】execute_steps (Grok P5 - 缓存清理)
# ======================================================================
def execute_steps(steps, run_context=None, status_callback=None):
    """V45.1: 修复 last_position, LoopCache, 和 _quick_image_check"""
    
    print(f"\n--- 开始执行宏 (V45.1 / OCR V43.3) ---")
    if FORCE_OCR_ENGINE and FORCE_OCR_ENGINE != 'auto':
        print(f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"!! [调试] 全局 OCR 引擎已强制为: {FORCE_OCR_ENGINE}")
        print(f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    
    perf_monitor.reset()
    loop_cache.reset()
    
    if run_context is None:
        run_context = {}
    
    if 'last_position' not in run_context:
        run_context['last_position'] = (None, None)
    if 'stop_requested' not in run_context:
        run_context['stop_requested'] = False
    
    pc = 0
    loop_stack = []

    try:
        while pc < len(steps):
            
            if run_context.get('stop_requested', False):
                print(f"  [停止] 检测到 F11 软停止请求。")
                break
                
            step = steps[pc]
            action = step.get('action', '')
            params = step.get('params', {})
            print(f"步骤 {pc + 1}/{len(steps)}: {action}")
            default_pc_increment = True

            try:
                if action in ('FIND_IMAGE', 'FIND_TEXT', 'IF_IMAGE_FOUND', 'IF_TEXT_FOUND'):
                    step_signature = f"{action}_{params.get('path', params.get('text', ''))}"
                    use_loop_cache = loop_cache.current_loop_id is not None
                    
                    location_center = _execute_find_action(action, params, run_context, use_loop_cache, step_signature)
                    
                    if action.startswith('IF_'):
                        if not location_center:
                            print("  [IF✗] 跳转 ELSE/END_IF")
                            pc = find_matching_jump(steps, pc, 'IF_', 'END_IF', ['ELSE', 'END_IF'])
                            default_pc_increment = False
                        else:
                            print("  [IF✓] 继续")
                            pyautogui.moveTo(location_center[0], location_center[1], duration=0.25)
                    else:
                        if not location_center:
                            print("  [停止] 未找到目标")
                            break
                        pyautogui.moveTo(location_center[0], location_center[1], duration=0.25)
                
                elif action == 'MOVE_TO':
                    x, y = int(params['x']), int(params['y'])
                    pyautogui.moveTo(x, y, duration=0.25)
                    run_context['last_position'] = (x, y)
                
                elif action == 'MOVE_OFFSET':
                    last_pos = run_context['last_position']
                    if last_pos[0] is None:
                        print("  [错误] 无法相对移动")
                        break
                    x_off, y_off = int(params['x_offset']), int(params['y_offset'])
                    pyautogui.move(x_off, y_off, duration=0.25)
                    run_context['last_position'] = (last_pos[0] + x_off, last_pos[1] + y_off)
                
                elif action == 'CLICK':
                    pyautogui.click(button=params.get('button', 'left').lower())
                elif action == 'SCROLL':
                    amount = int(params.get('amount', 0))
                    pyautogui.scroll(amount)
                    print(f"  [动作] 滚轮滚动 {amount}")
                elif action == 'WAIT':
                    time.sleep(int(params['ms']) / 1000.0)
                elif action == 'TYPE_TEXT':
                    pyperclip.copy(params['text'])
                    pyautogui.hotkey('ctrl', 'v')
                elif action == 'PRESS_KEY':
                    key = params.get('key', 'enter').lower()
                    if key in pyautogui.KEYBOARD_KEYS:
                        pyautogui.keyDown(key)
                        time.sleep(0.05)
                        pyautogui.keyUp(key)
                    else:
                        pyperclip.copy(key)
                        pyautogui.hotkey('ctrl', 'v') 
                elif action == 'ELSE':
                    pc = find_matching_jump(steps, pc, 'IF_', 'END_IF', ['END_IF'])
                    default_pc_increment = False
                elif action == 'END_IF':
                    pass
                elif action == 'LOOP_START':
                    pc, default_pc_increment = _execute_loop_start(steps, pc, loop_stack, params, run_context, status_callback)
                elif action == 'END_LOOP':
                    if loop_stack:
                        loop_cache.next_iteration()
                        pc = loop_stack[-1]['start_pc']
                        default_pc_increment = False
                
            except Exception as e:
                print(f"  [严重错误] {e}")
                import traceback
                traceback.print_exc()
                break
                
            if default_pc_increment:
                pc += 1
    finally:
        # 【V45.0 修复】 (Grok P5) 确保任何退出都重置缓存
        loop_cache.reset()
        print(f"\n--- 宏执行完毕 ---")
        print(f"[性能] {perf_monitor.get_stats()}")
        print()

# ======================================================================
# 【V45.0 修复】_execute_find_action (V44.2 逻辑 + P2/P3 修复)
# ======================================================================
def _execute_find_action(action, params, run_context, use_loop_cache=False, step_signature=None):
    """(V45.0) V44 架构: 使用 cv2 + 修复 P2/P3"""
    
    cache_box = params.get('cache_box') 
    if not cache_box:
        try:
            cache_x, cache_y = int(params.get('cache_x')), int(params.get('cache_y'))
            if cache_x is not None:
                cache_box = [cache_x, cache_y, cache_x+1, cache_y+1]
                print(f"  [缓存] 检测到旧版 cache_x/y，已转换为 {cache_box}")
        except (TypeError, ValueError):
            cache_box = None
            
    is_image = 'IMAGE' in action
    is_ocr = not is_image # 【V45.0 修复】(P3)
    location_result = None
    
    final_engine = 'auto'
    if is_ocr:
        step_engine = params.get('engine', 'auto')
        final_engine = FORCE_OCR_ENGINE if (FORCE_OCR_ENGINE and FORCE_OCR_ENGINE != 'auto') else step_engine
        if final_engine != 'auto' and final_engine != step_engine:
             print(f"  [调试] 全局覆盖引擎: {final_engine} (脚本原设: {step_engine})")

    # 1. 确定第一次截图区域
    region = None
    if cache_box and isinstance(cache_box, list) and len(cache_box) == 4:
        x1, y1, x2, y2 = cache_box
        padding = 20
        region = (
            max(0, x1 - padding), 
            max(0, y1 - padding), 
            (x2 - x1) + (padding * 2), 
            (y2 - y1) + (padding * 2)
        )
        print(f"  [缓存] 验证区域 {region}...")

    # 2. 执行第一次截图 (V44 架构)
    screenshot, offset = smart_screenshot(region)

    # 3. 优先检查循环缓存
    if use_loop_cache and step_signature:
        cached_loc = loop_cache.get_cache(step_signature)
        if cached_loc:
            if is_image:
                # 【V45.1 修复】(Grok P1/P2) 调用新的 _quick_image_check_V44
                if _quick_image_check_V44(params['path'], 
                                      float(params.get('confidence', 0.8)),
                                      screenshot, 
                                      offset,
                                      cached_loc):
                    perf_monitor.record_cache_hit(is_loop_cache=True, is_ocr=False) # (P3)
                    print(f"  [循环缓存✓] {cached_loc}")
                    run_context['last_position'] = cached_loc
                    return cached_loc
    
    # 4. 执行第一次查找 (使用已有的截图)
    if is_image:
        location_result = find_image_location(
            params['path'], 
            float(params.get('confidence', 0.8)), 
            screenshot=screenshot, offset=offset,
            use_loop_cache=False, 
            step_signature=step_signature
        )
    else:
        location_result = ocr_engine.find_text_location(
            params['text'], 
            params.get('lang', 'eng'), 
            params.get('debug', True), 
            screenshot_pil=screenshot, offset=offset,
            engine=final_engine
        )
    
    # 5. 第二次截图 (全局搜索，仅当第一次是区域搜索且失败时)
    if not location_result and cache_box:
        print("  [缓存✗] 全局搜索...")
        screenshot, offset = smart_screenshot(None) # 全局截图

        # 6. 第二次查找
        if is_image:
            location_result = find_image_location(
                params['path'], 
                float(params.get('confidence', 0.8)), 
                screenshot=screenshot, offset=offset,
                use_loop_cache=False,
                step_signature=step_signature
            )
        else:
            location_result = ocr_engine.find_text_location(
                params['text'], 
                params.get('lang', 'eng'), 
                params.get('debug', True), 
                screenshot_pil=screenshot, offset=offset,
                engine=final_engine
            )
        
        # 7. 更新缓存
        if location_result:
            if is_image:
                cx, cy, w, h = location_result
                new_box = [cx - w//2, cy - h//2, cx + w//2, cy + h//2]
            else:
                if isinstance(location_result, (list, tuple)) and len(location_result) == 2:
                    new_box = [location_result[0], location_result[1], location_result[0]+1, location_result[1]+1]
                elif isinstance(location_result, (list, tuple)) and len(location_result) == 4:
                    new_box = location_result
                else:
                    new_box = None
            
            if new_box:
                params['cache_box'] = new_box
                if 'cache_x' in params: del params['cache_x']
                if 'cache_y' in params: del params['cache_y']
                print(f"  [缓存] 更新为 {new_box}")
            
    # 8. 处理结果
    if location_result:
        if is_image:
            center_point = (location_result[0], location_result[1])
        else:
            if isinstance(location_result, (list, tuple)) and len(location_result) == 2:
                center_point = location_result
            elif isinstance(location_result, (list, tuple)) and len(location_result) == 4:
                box = location_result
                center_point = ((box[0] + box[2]) // 2, (box[1] + box[3]) // 2)
            else:
                print(f"  [警告] OCR/图像 返回了意外的格式: {location_result}")
                perf_monitor.record_cache_miss(is_ocr=is_ocr) # (P3)
                return None
        
        if use_loop_cache and step_signature:
            loop_cache.set_cache(step_signature, center_point)
        
        run_context['last_position'] = center_point
        return center_point
    
    perf_monitor.record_cache_miss(is_ocr=is_ocr) # (P3)
    return None

def _execute_loop_start(steps, pc, loop_stack, params, run_context, status_callback=None):
    """(V43) 处理 LOOP_START 逻辑，并使用循环缓存"""
    loop_top = loop_stack[-1] if loop_stack else None
    
    if loop_top and loop_top['start_pc'] == pc:
        loop_top['remaining'] -= 1
        if run_context is not None:
            if 'loops_executed' not in run_context: 
                run_context['loops_executed'] = 0
            run_context['loops_executed'] += 1
        
        loops_count = run_context.get('loops_executed', 0) if run_context else 0
        if status_callback:
            status_callback(f"循环: {loops_count} 次")
        
        if loop_top['remaining'] > 0:
            print(f"  [LOOP] 剩余 {loop_top['remaining']} (已执行 {loops_count} 次)")
            return pc, True
        else:
            # 【V45.1 修复】 G5 BUG 清理。移除多余的 exit_loop()。
            # 'finally' 块会调用 reset()，这里不需要清理。
            # loop_cache.exit_loop() 
            loop_stack.pop()
            new_pc = find_matching_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
            return new_pc, False
    else:
        times = int(params.get('times', 1))
        if times <= 0:
            new_pc = find_matching_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
            return new_pc, False
        else:
            loop_id = f"loop_{pc}_{id(loop_stack)}"
            loop_cache.enter_loop(loop_id)
            
            loop_stack.append({'start_pc': pc, 'remaining': times})
            if run_context is not None:
                if 'loops_executed' not in run_context: 
                    run_context['loops_executed'] = 0
                run_context['loops_executed'] += 1
            
            loops_count = run_context.get('loops_executed', 0) if run_context else 0
            if status_callback:
                status_callback(f"循环: {loops_count} 次")
            
            print(f"  [LOOP] 开始 {times} 次 (已执行 {loops_count} 次)")
            return pc, True

def find_matching_jump(steps, start_index, open_block_prefix, close_block, targets):
    """(保持 V39 不变) 查找匹配的跳转位置"""
    nest_level = 0
    pc = start_index + 1
    while pc < len(steps):
        action = steps[pc].get('action', '')
        is_open = (action.startswith(open_block_prefix) if open_block_prefix.endswith('_') 
                   else (action == open_block_prefix))

        if is_open:
            nest_level += 1
        elif action in targets and nest_level == 0:
            return pc + 1
        elif action == close_block:
            if nest_level > 0:
                nest_level -= 1
            elif nest_level == 0 and close_block in targets:
                return pc + 1
        pc += 1
    return len(steps)

# 【V45.1】版本号
macro_engine_version = f"V45.1 (Core - 多尺度CV2) / OCR (V43.2) / OpenCV: {OPENCV_AVAILABLE}"