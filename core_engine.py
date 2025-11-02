# core_engine.py
# 描述：自动化宏的核心功能引擎 (V43 - 性能优化版)
# 优化内容:
# 1. 循环内智能缓存 (Loop Cache) - 大幅提升循环执行速度
# 2. 区域截图优化 - 减少不必要的全屏截图
# 3. 缓存命中率统计 - 实时监控性能
# 4. 预处理图像复用 - 减少重复计算

import pyautogui
import time
from PIL import Image, ImageGrab, ImageStat
import re
import pyperclip
import os
import sys
from collections import defaultdict

# 【V43】导入 OCR 模块
try:
    import ocr_engine
except ImportError:
    print("[严重错误] 未找到 'ocr_engine.py'。")
    class ocr_engine:
        def find_text_location(*args, **kwargs): return None
        WINOCR_AVAILABLE = False
        TESSERACT_AVAILABLE = False

# 【V43】OpenCV 检测
try:
    import cv2
    OPENCV_AVAILABLE = True
    print("[配置] ✓ OpenCV 引擎就绪 (图像查找 `confidence` 已启用)")
except ImportError:
    OPENCV_AVAILABLE = False
    print("[配置] ✗ 未找到 OpenCV。`FIND_IMAGE` (模糊查找) 将禁用。")

# ======================================================================
# 【V43 新增】性能监控和缓存统计
# ======================================================================
class PerformanceMonitor:
    """性能监控器 - 跟踪缓存命中率和执行时间"""
    def __init__(self):
        self.reset()
    
    def reset(self):
        self.cache_hits = 0
        self.cache_misses = 0
        self.loop_cache_hits = 0
        self.total_searches = 0
        self.search_times = []
    
    def record_cache_hit(self, is_loop_cache=False):
        self.cache_hits += 1
        if is_loop_cache:
            self.loop_cache_hits += 1
    
    def record_cache_miss(self):
        self.cache_misses += 1
    
    def record_search_time(self, duration):
        self.search_times.append(duration)
        self.total_searches += 1
    
    def get_stats(self):
        total = self.cache_hits + self.cache_misses
        if total == 0:
            return "无缓存数据"
        
        hit_rate = (self.cache_hits / total) * 100
        avg_time = sum(self.search_times) / len(self.search_times) if self.search_times else 0
        
        return (f"缓存命中率: {hit_rate:.1f}% ({self.cache_hits}/{total}) | "
                f"循环缓存: {self.loop_cache_hits} | "
                f"平均查找: {avg_time*1000:.0f}ms")

# 全局性能监控器
perf_monitor = PerformanceMonitor()

# ======================================================================
# 【V43 新增】循环缓存管理器
# ======================================================================
class LoopCacheManager:
    """循环缓存管理器 - 在循环内保持目标位置缓存"""
    def __init__(self):
        self.loop_caches = {}  # {loop_id: {step_signature: location}}
        self.current_loop_id = None
        self.loop_iteration = 0
    
    def enter_loop(self, loop_id):
        """进入循环"""
        self.current_loop_id = loop_id
        self.loop_iteration = 0
        if loop_id not in self.loop_caches:
            self.loop_caches[loop_id] = {}
    
    def next_iteration(self):
        """下一次迭代"""
        self.loop_iteration += 1
    
    def exit_loop(self):
        """退出循环并清理缓存"""
        if self.current_loop_id in self.loop_caches:
            del self.loop_caches[self.current_loop_id]
        self.current_loop_id = None
        self.loop_iteration = 0
    
    def get_cache(self, step_signature):
        """获取循环内缓存"""
        if self.current_loop_id is None or self.loop_iteration == 0:
            return None
        
        cache = self.loop_caches.get(self.current_loop_id, {})
        return cache.get(step_signature)
    
    def set_cache(self, step_signature, location):
        """设置循环内缓存"""
        if self.current_loop_id is None:
            return
        
        if self.current_loop_id not in self.loop_caches:
            self.loop_caches[self.current_loop_id] = {}
        
        self.loop_caches[self.current_loop_id][step_signature] = location

# 全局循环缓存管理器
loop_cache = LoopCacheManager()

# ======================================================================
# 内部变量
# ======================================================================
last_position = (None, None)

# ======================================================================
# 【V43 优化】智能区域截图
# ======================================================================
def smart_screenshot(region=None, expand_ratio=1.5):
    """
    智能截图 - 根据需要扩展区域以提高容错性
    expand_ratio: 区域扩展比例 (默认1.5倍)
    """
    if region:
        x, y, w, h = region
        # 扩展区域
        expand_w = int(w * expand_ratio)
        expand_h = int(h * expand_ratio)
        offset_x = (expand_w - w) // 2
        offset_y = (expand_h - h) // 2
        
        # 确保不超出屏幕边界
        screen_w, screen_h = pyautogui.size()
        new_x = max(0, x - offset_x)
        new_y = max(0, y - offset_y)
        new_w = min(expand_w, screen_w - new_x)
        new_h = min(expand_h, screen_h - new_y)
        
        return ImageGrab.grab(bbox=(new_x, new_y, new_x + new_w, new_y + new_h)), (new_x, new_y)
    else:
        return ImageGrab.grab(), (0, 0)

# ======================================================================
# 【V43 优化】图像查找 (增强缓存逻辑)
# ======================================================================
def find_image_location(image_path, confidence=0.8, region=None, use_loop_cache=False, step_signature=None):
    """V43: 增强的图像查找，支持循环缓存"""
    start_time = time.time()
    
    try:
        # 【V43】检查循环缓存
        if use_loop_cache and step_signature:
            cached_loc = loop_cache.get_cache(step_signature)
            if cached_loc:
                # 快速验证缓存是否仍然有效
                verify_region = (
                    cached_loc[0] - 50, cached_loc[1] - 50,
                    100, 100
                )
                quick_check = _quick_image_check(image_path, verify_region, confidence)
                if quick_check:
                    perf_monitor.record_cache_hit(is_loop_cache=True)
                    perf_monitor.record_search_time(time.time() - start_time)
                    print(f"  [循环缓存✓] 命中 {cached_loc}")
                    return cached_loc
        
        # 【V43】标准查找逻辑
        pyautogui_region = None
        if region:
            pyautogui_region = (region[0], region[1], region[2], region[3])
        
        kwargs = {'region': pyautogui_region}
        if OPENCV_AVAILABLE:
            kwargs['confidence'] = confidence
        else:
            if confidence < 0.99:
                print("  [警告] 未安装 OpenCV, 'confidence' (模糊查找) 将被忽略。")

        box = pyautogui.locateOnScreen(image_path, **kwargs)
        
        if box:
            center = pyautogui.center(box)
            result = (center.x, center.y, box.width, box.height)
            
            # 【V43】更新循环缓存
            if use_loop_cache and step_signature:
                loop_cache.set_cache(step_signature, (center.x, center.y))
            
            if region:
                perf_monitor.record_cache_hit(is_loop_cache=False)
                print(f"  [区域缓存✓] 图像于 ({center.x}, {center.y})")
            else:
                perf_monitor.record_cache_miss()
                print(f"  [全局搜索] 图像于 ({center.x}, {center.y})")
            
            perf_monitor.record_search_time(time.time() - start_time)
            return result
        
        perf_monitor.record_cache_miss()
        perf_monitor.record_search_time(time.time() - start_time)
        return None
        
    except Exception as e:
        print(f"  [错误] 图像查找: {e}")
        perf_monitor.record_search_time(time.time() - start_time)
        return None

def _quick_image_check(image_path, region, confidence):
    """快速检查图像是否在指定小区域内"""
    try:
        kwargs = {'region': region}
        if OPENCV_AVAILABLE:
            kwargs['confidence'] = confidence
        
        box = pyautogui.locateOnScreen(image_path, **kwargs)
        return box is not None
    except:
        return False

# ======================================================================
# 【V43 优化】execute_steps (增强循环性能)
# ======================================================================
def execute_steps(steps, run_context=None, status_callback=None):
    """V43: 优化的宏执行引擎 (循环缓存 + 性能监控)"""
    print(f"\n--- 开始执行宏 (V43 性能优化版 / OCR V41.1) ---")
    
    # 重置性能监控
    perf_monitor.reset()
    
    global last_position
    last_position = (None, None)
    pc = 0
    loop_stack = []

    while pc < len(steps):
        
        if run_context and run_context.get('stop_requested', False):
            print(f"  [停止] 检测到 F11 软停止请求。")
            break
            
        step = steps[pc]
        action = step.get('action', '')
        params = step.get('params', {})
        print(f"步骤 {pc + 1}/{len(steps)}: {action}")
        default_pc_increment = True

        try:
            if action in ('FIND_IMAGE', 'FIND_TEXT', 'IF_IMAGE_FOUND', 'IF_TEXT_FOUND'):
                # 【V43】生成步骤签名用于循环缓存
                step_signature = f"{action}_{params.get('path', params.get('text', ''))}"
                use_loop_cache = loop_cache.current_loop_id is not None
                
                location_center = _execute_find_action(action, params, use_loop_cache, step_signature)
                
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
                last_position = (x, y) 
            
            elif action == 'MOVE_OFFSET':
                if last_position[0] is None:
                    print("  [错误] 无法相对移动")
                    break
                x_off, y_off = int(params['x_offset']), int(params['y_offset'])
                pyautogui.move(x_off, y_off, duration=0.25)
                last_position = (last_position[0] + x_off, last_position[1] + y_off) 
            
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
                    # 【V43】循环迭代计数
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
    
    # 【V43】输出性能统计
    print(f"\n--- 宏执行完毕 ---")
    print(f"[性能] {perf_monitor.get_stats()}")
    print()

# ======================================================================
# 【V43 优化】辅助函数
# ======================================================================
def _execute_find_action(action, params, use_loop_cache=False, step_signature=None):
    """(V43) 统一处理 FIND 和 IF_FIND (智能缓存 + 循环缓存)"""
    global last_position
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
    location_result = None
    
    # 【V43】优先检查循环缓存
    if use_loop_cache and step_signature:
        cached_loc = loop_cache.get_cache(step_signature)
        if cached_loc:
            # 快速验证
            if is_image:
                verify_region = (cached_loc[0] - 30, cached_loc[1] - 30, 60, 60)
                if _quick_image_check(params['path'], verify_region, float(params.get('confidence', 0.8))):
                    perf_monitor.record_cache_hit(is_loop_cache=True)
                    print(f"  [循环缓存✓] {cached_loc}")
                    last_position = cached_loc
                    return cached_loc
    
    # 【V43】检查常规缓存
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
        
        if is_image:
            location_result = find_image_location(
                params['path'], 
                float(params.get('confidence', 0.8)), 
                region,
                use_loop_cache,
                step_signature
            )
        else:
            location_result = ocr_engine.find_text_location(
                params['text'], 
                params.get('lang', 'eng'), 
                params.get('debug', True), 
                region
            )
    
    # 【V43】全局搜索
    if not location_result:
        if cache_box:
            print("  [缓存✗] 全局搜索...")
        
        if is_image:
            location_result = find_image_location(
                params['path'], 
                float(params.get('confidence', 0.8)), 
                None,
                use_loop_cache,
                step_signature
            )
        else:
            location_result = ocr_engine.find_text_location(
                params['text'], 
                params.get('lang', 'eng'), 
                params.get('debug', True), 
                None
            )
        
        if location_result:
            if is_image:
                cx, cy, w, h = location_result
                new_box = [cx - w//2, cy - h//2, cx + w//2, cy + h//2]
            else:
                new_box = location_result
            
            params['cache_box'] = new_box
            if 'cache_x' in params: del params['cache_x']
            if 'cache_y' in params: del params['cache_y']
            print(f"  [缓存] 更新为 {new_box}")
            
    if location_result:
        if is_image:
            center_point = (location_result[0], location_result[1])
        else:
            # 【V43 修复】OCR 可能返回 (x, y) 或 [x1, y1, x2, y2]
            if isinstance(location_result, (list, tuple)) and len(location_result) == 2:
                # 直接返回的坐标
                center_point = location_result
            elif isinstance(location_result, (list, tuple)) and len(location_result) == 4:
                # 边界框格式
                box = location_result
                center_point = ((box[0] + box[2]) // 2, (box[1] + box[3]) // 2)
            else:
                print(f"  [警告] OCR 返回了意外的格式: {location_result}")
                return None
        
        # 【V43】更新循环缓存
        if use_loop_cache and step_signature:
            loop_cache.set_cache(step_signature, center_point)
        
        last_position = center_point
        return center_point
            
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
        
        # 【V43】调用状态回调更新循环计数
        loops_count = run_context.get('loops_executed', 0) if run_context else 0
        if status_callback:
            status_callback(f"循环: {loops_count} 次")
        
        if loop_top['remaining'] > 0:
            print(f"  [LOOP] 剩余 {loop_top['remaining']} (已执行 {loops_count} 次)")
            return pc, True
        else:
            print(f"  [LOOP] 循环完成 (共 {loops_count} 次)")
            # 【V43】退出循环，清理缓存
            loop_cache.exit_loop()
            loop_stack.pop()
            new_pc = find_matching_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
            return new_pc, False
    else:
        times = int(params.get('times', 1))
        if times <= 0:
            new_pc = find_matching_jump(steps, pc, 'LOOP_START', 'END_LOOP', ['END_LOOP'])
            return new_pc, False
        else:
            # 【V43】进入新循环，初始化缓存
            loop_id = f"loop_{pc}_{id(loop_stack)}"
            loop_cache.enter_loop(loop_id)
            
            loop_stack.append({'start_pc': pc, 'remaining': times})
            if run_context is not None:
                if 'loops_executed' not in run_context: 
                    run_context['loops_executed'] = 0
                run_context['loops_executed'] += 1
            
            # 【V43】调用状态回调更新循环计数
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

# 【V43】版本号
macro_engine_version = f"V43 (性能优化 - 循环缓存) / OCR (V41.1) / OpenCV: {OPENCV_AVAILABLE}"