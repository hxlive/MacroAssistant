# -*- coding: utf-8 -*-
# ocr_engine.py
# 描述:自动化宏的 OCR 功能引擎
# 版本:1.8.2

from PIL import Image, ImageGrab
import re
import os
import subprocess
import time
import sys
import threading
import traceback
import unicodedata

RAPIDOCR_CLASS = None
NUMPY_CV2_AVAILABLE = False

try:
    import numpy as np
    import cv2
    NUMPY_CV2_AVAILABLE = True
    from rapidocr import RapidOCR
    RAPIDOCR_CLASS = RapidOCR 
except Exception as e:
    print(f"[OCR] RapidOCR 依赖加载失败: {e}")
    traceback.print_exc()
    pass 

_RAPID_OCR_INSTANCE = None
_RAPID_OCR_INIT_FAILED = False
_RAPID_OCR_LOCK = threading.Lock() 
_ENGINE_ERROR_LOGGED = {}
_OCR_ERROR_LOG_LOCK = threading.Lock()

_TESSERACT_CMD = None
_TESSERACT_TESSDATA = None
_TESSERACT_CHECKED = False
_TESSERACT_LOCK = threading.Lock()

_AUTO_ENGINE_CACHE = {}
_AUTO_ENGINE_CACHE_LOCK = threading.Lock()
_OCR_POSITION_CACHE = {}
_OCR_POSITION_CACHE_LOCK = threading.Lock()

def preload_engines():
    print("[OCR] 后台预热开始...")
    if NUMPY_CV2_AVAILABLE:
        get_rapid_ocr_engine()
    get_tesseract_cmd()

def get_rapid_ocr_engine():
    global _RAPID_OCR_INSTANCE, _RAPID_OCR_INIT_FAILED
    if _RAPID_OCR_INSTANCE: return _RAPID_OCR_INSTANCE
    if not RAPIDOCR_CLASS: return None
    
    with _RAPID_OCR_LOCK: 
        if _RAPID_OCR_INSTANCE: return _RAPID_OCR_INSTANCE
        if _RAPID_OCR_INIT_FAILED: return None
        try:
            print("[OCR] 正在加载 RapidOCR 模型...")
            t0 = time.time()
            _RAPID_OCR_INSTANCE = RAPIDOCR_CLASS()
            print(f"[OCR] RapidOCR 就绪 ({time.time()-t0:.2f}s)")
            return _RAPID_OCR_INSTANCE
        except Exception as e:
            print(f"[严重错误] RapidOCR 初始化失败: {e}")
            traceback.print_exc()
            _RAPID_OCR_INIT_FAILED = True
            return None

def get_tesseract_cmd():
    global _TESSERACT_CMD, _TESSERACT_TESSDATA, _TESSERACT_CHECKED
    if _TESSERACT_CHECKED: return _TESSERACT_CMD
    
    with _TESSERACT_LOCK:
        if _TESSERACT_CHECKED: return _TESSERACT_CMD
        
        search_roots = [
            getattr(sys, '_MEIPASS', None),
            os.path.dirname(os.path.abspath(__file__)),
            os.path.dirname(sys.executable),
            os.path.join(os.path.dirname(sys.executable), '_internal')
        ]
        for root in search_roots:
            if not root: continue
            exe = os.path.join(root, 'tesseract_local', 'tesseract.exe')
            if os.path.exists(exe):
                _TESSERACT_CMD = exe
                data = os.path.join(root, 'tesseract_local', 'tessdata')
                if os.path.exists(data):
                    _TESSERACT_TESSDATA = os.path.abspath(data)
                break
        
        if not _TESSERACT_CMD:
            try:
                cflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                find_cmd = 'where' if os.name == 'nt' else 'which'
                res = subprocess.run([find_cmd, 'tesseract'], capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=cflags)
                if res.returncode == 0: _TESSERACT_CMD = res.stdout.strip().split('\n')[0]
            except Exception:
                pass

        if _TESSERACT_CMD:
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
            except Exception as e:
                print(f"[OCR] pytesseract 加载失败: {e}")
                _TESSERACT_CMD = None
        _TESSERACT_CHECKED = True
        return _TESSERACT_CMD

LANG_MAP = {
    'winocr': {'eng': 'en-US', 'chi_sim': 'zh-Hans'},
    'rapidocr': {'eng': 'en', 'chi_sim': 'ch'},
    'tesseract': {'eng': 'eng', 'chi_sim': 'chi_sim'}
}

MAX_MERGE_WORDS = 20
TESSERACT_MIN_CONFIDENCE = 30
TESSERACT_PSM_ORDER = (6, 11, 3)
AUTO_ENGINE_ORDER = ('winocr', 'rapidocr', 'tesseract')
OCR_POSITION_CACHE_MAX_MISSES = 2
OCR_POSITION_CACHE_PAD_X = 90
OCR_POSITION_CACHE_PAD_Y = 45
OCR_POSITION_CACHE_MAX_PAD_X = 260
OCR_POSITION_CACHE_MAX_PAD_Y = 90
OCR_POSITION_CACHE_PAD_SCALES = (1, 2)
RAPIDOCR_ENHANCED_MAX_PIXELS = 300000
AUTO_ENGINE_CACHE_MAX_SIZE = 256
OCR_POSITION_CACHE_MAX_SIZE = 512

def _normalize_text(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(text))).casefold()

def _dict_box_to_tuple(box):
    try:
        return (box['x'], box['y'], box['x'] + box['width'], box['y'] + box['height'])
    except (KeyError, TypeError):
        return None

def _quad_box_to_tuple(box, scale=1):
    xs = [p[0] for p in box]
    ys = [p[1] for p in box]
    return (min(xs) // scale, min(ys) // scale, max(xs) // scale, max(ys) // scale)

def _rect_box_to_tuple(left, top, right, bottom, scale=1):
    return (left // scale, top // scale, right // scale, bottom // scale)

def _ocr_box_to_tuple(box, scale=1):
    if box is None:
        return None
    if NUMPY_CV2_AVAILABLE and isinstance(box, np.ndarray):
        box = box.tolist()
    if not isinstance(box, (list, tuple)):
        return None

    numeric_types = (int, float)
    if NUMPY_CV2_AVAILABLE:
        numeric_types = numeric_types + (np.integer, np.floating)

    if len(box) == 4 and all(isinstance(v, numeric_types) for v in box):
        return _rect_box_to_tuple(box[0], box[1], box[2], box[3], scale)
    if len(box) >= 4 and all(isinstance(p, (list, tuple)) and len(p) >= 2 for p in box[:4]):
        return _quad_box_to_tuple(box[:4], scale)
    return None

def _center_word_boxes(boxes, offset):
    left = min(b[0] for b in boxes)
    top = min(b[1] for b in boxes)
    right = max(b[2] for b in boxes)
    bottom = max(b[3] for b in boxes)
    return (offset[0] + (left + right) // 2, offset[1] + (top + bottom) // 2)

def _recognition_pass(words, full_text, engine_name='OCR', pass_label=''):
    return {
        'words': words,
        'full_text': full_text,
        'engine_name': engine_name,
        'pass_label': pass_label,
    }

def _remember_bounded_cache(cache, key, value, max_size):
    if key in cache:
        cache.pop(key, None)
    cache[key] = value
    while len(cache) > max_size:
        oldest_key = next(iter(cache))
        cache.pop(oldest_key, None)

def _get_bounded_cache(cache, key):
    if key not in cache:
        return None
    value = cache.pop(key)
    cache[key] = value
    return value

def _auto_cache_key(target_norm, lang, enhanced_mode, cache_key=None):
    return (cache_key, target_norm, lang, bool(enhanced_mode))

def _get_auto_engine_order(target_norm, lang, enhanced_mode, cache_key=None):
    key = _auto_cache_key(target_norm, lang, enhanced_mode, cache_key)
    with _AUTO_ENGINE_CACHE_LOCK:
        preferred = _get_bounded_cache(_AUTO_ENGINE_CACHE, key)
    if preferred in AUTO_ENGINE_ORDER:
        return (preferred,) + tuple(e for e in AUTO_ENGINE_ORDER if e != preferred)
    return AUTO_ENGINE_ORDER

def _remember_auto_engine(target_norm, lang, enhanced_mode, engine_name, cache_key=None):
    if engine_name not in AUTO_ENGINE_ORDER:
        return
    key = _auto_cache_key(target_norm, lang, enhanced_mode, cache_key)
    with _AUTO_ENGINE_CACHE_LOCK:
        _remember_bounded_cache(_AUTO_ENGINE_CACHE, key, engine_name, AUTO_ENGINE_CACHE_MAX_SIZE)

def _position_cache_key(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key=None):
    return (cache_key, target_norm, lang, bool(enhanced_mode), int(offset[0]), int(offset[1]), screenshot_pil.width, screenshot_pil.height)

def _get_ocr_position_cache(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key=None):
    key = _position_cache_key(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key)
    with _OCR_POSITION_CACHE_LOCK:
        cached = _get_bounded_cache(_OCR_POSITION_CACHE, key)
        return dict(cached) if cached else None

def _remember_ocr_position(target_norm, lang, enhanced_mode, screenshot_pil, offset, position, engine_name, cache_key=None):
    key = _position_cache_key(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key)
    with _OCR_POSITION_CACHE_LOCK:
        _remember_bounded_cache(_OCR_POSITION_CACHE, key, {'pos': tuple(position), 'engine': engine_name, 'misses': 0}, OCR_POSITION_CACHE_MAX_SIZE)

def _record_ocr_position_cache_miss(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key=None):
    key = _position_cache_key(target_norm, lang, enhanced_mode, screenshot_pil, offset, cache_key)
    with _OCR_POSITION_CACHE_LOCK:
        cached = _OCR_POSITION_CACHE.get(key)
        if not cached:
            return
        cached['misses'] = cached.get('misses', 0) + 1
        if cached['misses'] >= OCR_POSITION_CACHE_MAX_MISSES:
            _OCR_POSITION_CACHE.pop(key, None)

def _crop_around_cached_position(screenshot_pil, offset, position, target_norm, pad_scale=1):
    rel_x = int(position[0] - offset[0])
    rel_y = int(position[1] - offset[1])
    if rel_x < 0 or rel_y < 0 or rel_x >= screenshot_pil.width or rel_y >= screenshot_pil.height:
        return None, None
    text_len = len(target_norm)
    if text_len <= 4:
        pad_x = OCR_POSITION_CACHE_PAD_X
        pad_y = OCR_POSITION_CACHE_PAD_Y
    else:
        pad_x = OCR_POSITION_CACHE_PAD_X + text_len * 5
        pad_y = OCR_POSITION_CACHE_PAD_Y + text_len // 4
    pad_x = min(OCR_POSITION_CACHE_MAX_PAD_X, int(pad_x * pad_scale))
    pad_y = min(OCR_POSITION_CACHE_MAX_PAD_Y, int(pad_y * pad_scale))
    left = max(0, rel_x - pad_x)
    top = max(0, rel_y - pad_y)
    right = min(screenshot_pil.width, rel_x + pad_x)
    bottom = min(screenshot_pil.height, rel_y + pad_y)
    if right <= left or bottom <= top:
        return None, None
    return screenshot_pil.crop((left, top, right, bottom)), (offset[0] + left, offset[1] + top)

class _OcrContext:
    def __init__(self, screenshot_pil, offset, should_close=False):
        self.screenshot = screenshot_pil
        self.offset = offset
        self.should_close = should_close
        self._bgr_cache = {}

    def get_bgr(self, image_pil=None):
        if not NUMPY_CV2_AVAILABLE:
            return None
        image_pil = image_pil or self.screenshot
        cache_key = id(image_pil)
        cached = self._bgr_cache.get(cache_key)
        if cached and cached[0] is image_pil:
            return cached[1]
        try:
            bgr = cv2.cvtColor(np.array(image_pil), cv2.COLOR_RGB2BGR)
        except Exception as e:
            print(f"  [OCR] 截图转换失败: {e}")
            return None
        self._bgr_cache[cache_key] = (image_pil, bgr)
        return bgr

    def close(self):
        self._bgr_cache.clear()
        if self.should_close and self.screenshot is not None:
            try:
                self.screenshot.close()
            except Exception:
                pass

def _capture_ocr_context(screenshot_pil, offset):
    if screenshot_pil is not None:
        return _OcrContext(screenshot_pil, offset, False)

    try:
        if sys.platform == 'win32':
            import ctypes
            SM_XVIRTUALSCREEN = 76
            SM_YVIRTUALSCREEN = 77
            SM_CXVIRTUALSCREEN = 78
            SM_CYVIRTUALSCREEN = 79
            user32 = ctypes.windll.user32
            vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
            primary_w = user32.GetSystemMetrics(0)
            primary_h = user32.GetSystemMetrics(1)
            if vw > 0 and vh > 0 and (vx != 0 or vy != 0 or vw > primary_w or vh > primary_h):
                return _OcrContext(ImageGrab.grab(all_screens=True), (vx, vy), True)
            return _OcrContext(ImageGrab.grab(), (0, 0), True)
        return _OcrContext(ImageGrab.grab(), (0, 0), True)
    except OSError as e:
        print(f"  [OCR] 截图失败 (锁屏或无显示器): {e}")
        return None

_ENGINE_SKIPPED = object()

def _debug_engine_name(engine_name, ocr_pass):
    return ocr_pass.get('engine_name') or engine_name

def _print_match_debug(engine_name, ocr_pass, match):
    cx, cy = match['position']
    display_name = _debug_engine_name(engine_name, ocr_pass)
    label = ocr_pass.get('pass_label') or ''
    label_part = f" {label}" if label else ''
    merge_part = ' 合并' if match.get('kind') in ('merged', 'chars') else ''
    word = match.get('word') or {}
    score = word.get('score', 0.0)
    if match.get('kind') in ('single', 'single_exact', 'single_substring') and score > 0:
        print(f"  [{display_name} OK]{label_part}{merge_part} ({cx}, {cy}) @ {score:.2f}")
    else:
        print(f"  [{display_name} OK]{label_part}{merge_part} ({cx}, {cy})")


def _result_from_match(match, full_text):
    return (match['position'], full_text)

def _run_ocr_engine(engine_name, runner, target_norm, offset, debug, attempts, tried_engines=None, record_stats=True):
    t0 = time.time()
    passes = None
    result = None
    skipped = False
    try:
        passes = runner()
        if passes is _ENGINE_SKIPPED:
            skipped = True
            return None
        for ocr_pass in passes or ():
            words = ocr_pass.get('words') or []
            full_text = ocr_pass.get('full_text', '')
            match = _MATCHER.match(words, target_norm, offset)
            if match:
                if debug:
                    _print_match_debug(engine_name, ocr_pass, match)
                result = _result_from_match(match, full_text)
                break
    except Exception as e:
        _log_engine_error(engine_name, e, debug)
        return None
    finally:
        if passes is not None and passes is not _ENGINE_SKIPPED and hasattr(passes, 'close'):
            try:
                passes.close()
            except Exception:
                pass
        if not skipped:
            duration = time.time() - t0
            if tried_engines is not None:
                tried_engines.add(engine_name)
            success = result is not None
            if record_stats:
                ocr_stats.record(engine_name, success, duration)
            attempts.append((engine_name, success, duration))
    return result

def _format_ocr_attempts(attempts):
    return ' -> '.join(f"{name} {'hit' if ok else 'miss'} {dt*1000:.0f}ms" for name, ok, dt in attempts)

def _log_engine_error(engine_name, error, debug):
    if debug:
        print(f"  [{engine_name}] 识别异常: {error}")
        return

    with _OCR_ERROR_LOG_LOCK:
        if _ENGINE_ERROR_LOGGED.get(engine_name):
            return
        _ENGINE_ERROR_LOGGED[engine_name] = True
    print(f"  [{engine_name}] 识别异常，已自动跳过该次结果: {error}")

def _make_ocr_engine_runners(ctx, target_norm, lang, debug, enhanced_mode):
    def try_winocr(image_pil=ctx.screenshot, image_offset=ctx.offset):
        if lang not in LANG_MAP['winocr']:
            return _ENGINE_SKIPPED
        try:
            import winocr
        except ImportError:
            return _ENGINE_SKIPPED
        return _recognize_winocr_words(winocr, LANG_MAP['winocr'][lang], debug, image_pil)

    def try_rapidocr(image_pil=ctx.screenshot, image_offset=ctx.offset):
        if lang not in LANG_MAP['rapidocr']:
            return _ENGINE_SKIPPED
        rapid_inst = get_rapid_ocr_engine()
        img_bgr = ctx.get_bgr(image_pil)
        if not rapid_inst or img_bgr is None:
            return _ENGINE_SKIPPED
        return _recognize_rapidocr_words(rapid_inst, debug, img_bgr, enhanced_mode)

    def try_tesseract(image_pil=ctx.screenshot, image_offset=ctx.offset):
        if lang not in LANG_MAP['tesseract'] or not get_tesseract_cmd():
            return _ENGINE_SKIPPED
        return _recognize_tesseract_words(LANG_MAP['tesseract'][lang], debug, image_pil, enhanced_mode)

    return {
        'winocr': try_winocr,
        'rapidocr': try_rapidocr,
        'tesseract': try_tesseract,
    }

def _try_cached_position(ctx, target_norm, lang, debug, enhanced_mode, cache_key, engine_runners, attempts, tried_engines):
    cached = _get_ocr_position_cache(target_norm, lang, enhanced_mode, ctx.screenshot, ctx.offset, cache_key)
    if not cached:
        ocr_stats.record_cache('position_absent')
        return None

    cached_engine = cached.get('engine')
    verification_order = ['winocr']
    if cached_engine and cached_engine != 'winocr':
        verification_order.append(cached_engine)

    for pad_scale in OCR_POSITION_CACHE_PAD_SCALES:
        crop_img, crop_offset = _crop_around_cached_position(ctx.screenshot, ctx.offset, cached.get('pos', (0, 0)), target_norm, pad_scale)
        if crop_img is None:
            ocr_stats.record_cache('position_miss')
            _record_ocr_position_cache_miss(target_norm, lang, enhanced_mode, ctx.screenshot, ctx.offset, cache_key)
            return None

        try:
            if debug:
                print(f"  [OCR cache] verify scale {pad_scale} {crop_img.width}x{crop_img.height} @ {crop_offset}")
            for engine_name in verification_order:
                runner = engine_runners.get(engine_name)
                if not runner:
                    continue
                result = _run_ocr_engine(engine_name, lambda r=runner: r(crop_img, crop_offset), target_norm, crop_offset, debug, attempts, tried_engines, record_stats=False)
                if result:
                    ocr_stats.record_cache('position_hit')
                    if pad_scale != OCR_POSITION_CACHE_PAD_SCALES[0]:
                        ocr_stats.record_cache('position_expand_hit')
                    ocr_stats.record_cache('position_winocr_hit' if engine_name == 'winocr' else 'position_cached_engine_hit')
                    _remember_auto_engine(target_norm, lang, enhanced_mode, engine_name, cache_key)
                    _remember_ocr_position(target_norm, lang, enhanced_mode, ctx.screenshot, ctx.offset, result[0], engine_name, cache_key)
                    if debug and attempts:
                        print(f"  [OCR auto] cache {_format_ocr_attempts(attempts)}")
                    return result
        finally:
            try:
                crop_img.close()
            except Exception:
                pass

    ocr_stats.record_cache('position_miss')
    _record_ocr_position_cache_miss(target_norm, lang, enhanced_mode, ctx.screenshot, ctx.offset, cache_key)
    return None

def _defer_tried_engines(engine_order, tried_engines):
    if not tried_engines:
        return engine_order
    untried = tuple(name for name in engine_order if name not in tried_engines)
    tried = tuple(name for name in engine_order if name in tried_engines)
    return untried + tried

class OCRPerformanceStats:
    def __init__(self):
        self._lock = threading.Lock()
        self.reset()
    def reset(self):
        with self._lock:
            self.stats = {'winocr': [0,0], 'rapidocr': [0,0], 'tesseract': [0,0]}
            self.cache_stats = {
                'position_hit': 0,
                'position_miss': 0,
                'position_absent': 0,
                'position_winocr_hit': 0,
                'position_cached_engine_hit': 0,
                'position_expand_hit': 0,
                'full_search': 0,
            }
            self.total_time = 0; self.call_count = 0
    def record(self, engine, success, duration):
        with self._lock:
            self.call_count += 1; self.total_time += duration
            self.stats.setdefault(engine, [0, 0])
            self.stats[engine][0 if success else 1] += 1
    def record_cache(self, event):
        with self._lock:
            if event in self.cache_stats:
                self.cache_stats[event] += 1
    def get_stats(self):
        with self._lock:
            if self.call_count == 0 and not any(self.cache_stats.values()): return "无 OCR 统计"
            avg = (self.total_time / self.call_count) * 1000 if self.call_count else 0
            parts = []
            for eng, (succ, fail) in self.stats.items():
                if succ + fail > 0: parts.append(f"{eng}({succ/(succ+fail)*100:.0f}%)")
            cache_parts = []
            if self.cache_stats['position_hit'] or self.cache_stats['position_miss']:
                cache_parts.append(f"位置缓存{self.cache_stats['position_hit']}/{self.cache_stats['position_hit'] + self.cache_stats['position_miss']}")
            if self.cache_stats.get('position_absent'):
                cache_parts.append(f"无缓存{self.cache_stats['position_absent']}")
            if self.cache_stats['position_winocr_hit']:
                cache_parts.append(f"WinOCR复核{self.cache_stats['position_winocr_hit']}")
            if self.cache_stats['position_cached_engine_hit']:
                cache_parts.append(f"缓存引擎复核{self.cache_stats['position_cached_engine_hit']}")
            if self.cache_stats['position_expand_hit']:
                cache_parts.append(f"expand{self.cache_stats['position_expand_hit']}")
            if self.cache_stats['full_search']:
                cache_parts.append(f"全量{self.cache_stats['full_search']}")
            suffix = f" | 缓存: {' '.join(cache_parts)}" if cache_parts else ""
            engine_part = ' | '.join(parts) if parts else '无引擎调用'
            return f"OCR统计 (均{avg:.0f}ms): {engine_part}{suffix}"

ocr_stats = OCRPerformanceStats()

def find_text_location(target_text, lang='eng', debug=False, screenshot_pil=None, offset=(0,0), engine='auto', enhanced_mode=False, cache_key=None):
    """
    查找文本在屏幕上的位置
    
    返回值:
    - 成功: ((x, y), full_text) - 坐标和完整识别文本
    - 失败: None
    """
    if not target_text or not isinstance(target_text, str): return None
    target_norm = _normalize_text(target_text)
    if not target_norm: return None

    ctx = _capture_ocr_context(screenshot_pil, offset)
    if ctx is None:
        return None

    try:
        attempts = []
        tried_engines = set()
        engine_runners = _make_ocr_engine_runners(ctx, target_norm, lang, debug, enhanced_mode)
        engine_order = _get_auto_engine_order(target_norm, lang, enhanced_mode, cache_key) if engine == 'auto' else (engine,)

        if engine == 'auto':
            cached_result = _try_cached_position(ctx, target_norm, lang, debug, enhanced_mode, cache_key, engine_runners, attempts, tried_engines)
            if cached_result:
                return cached_result
            engine_order = _defer_tried_engines(engine_order, tried_engines)

        if engine == 'auto' and engine_order:
            ocr_stats.record_cache('full_search')

        for engine_name in engine_order:
            runner = engine_runners.get(engine_name)
            if not runner:
                continue
            result = _run_ocr_engine(engine_name, runner, target_norm, ctx.offset, debug, attempts)
            if result:
                if engine == 'auto':
                    _remember_auto_engine(target_norm, lang, enhanced_mode, engine_name, cache_key)
                    _remember_ocr_position(target_norm, lang, enhanced_mode, ctx.screenshot, ctx.offset, result[0], engine_name, cache_key)
                    if debug and attempts:
                        print(f"  [OCR auto] {_format_ocr_attempts(attempts)}")
                return result

        if debug:
            print(f"  [失败] 未能找到 '{target_text}' (模式: {engine})")
        if debug and engine == 'auto' and attempts:
            print(f"  [OCR auto] {_format_ocr_attempts(attempts)}")
        if debug: print(f"  [统计] {ocr_stats.get_stats()}")
        return None
    finally:
        ctx.close()

def _text_range_box(word, start, end):
    text = word.get('text') or ''
    box = word.get('box')
    if not text or not box:
        return box
    length = len(text)
    start = max(0, min(start, length))
    end = max(start + 1, min(end, length))
    left, top, right, bottom = box
    width = max(1, right - left)
    span_left = left + (width * start) // length
    span_right = left + (width * end) // length
    if span_right <= span_left:
        span_right = min(right, span_left + 1)
    return (span_left, top, span_right, bottom)

def _build_char_index(words):
    text_parts = []
    char_boxes = []
    for word in words:
        text = word.get('text') or ''
        box = word.get('box')
        if not text or not box:
            continue
        for index, char in enumerate(text):
            text_parts.append(char)
            char_boxes.append(_text_range_box(word, index, index + 1))
    return ''.join(text_parts), char_boxes

class _Matcher:
    def match(self, words, target_norm, offset):
        valid_words = [word for word in words if word.get('text') and word.get('box')]

        for word in valid_words:
            if word['text'] == target_norm:
                return {
                    'position': _center_word_boxes([word['box']], offset),
                    'kind': 'single_exact',
                    'word': word,
                }

        for i, word in enumerate(valid_words):
            merged = word['text']
            if not target_norm.startswith(merged):
                continue
            boxes = [word['box']]
            for j in range(i + 1, min(i + MAX_MERGE_WORDS, len(valid_words))):
                merged += valid_words[j]['text']
                boxes.append(valid_words[j]['box'])
                if target_norm == merged:
                    return {
                        'position': _center_word_boxes(boxes, offset),
                        'kind': 'merged',
                        'word': None,
                    }

        for word in valid_words:
            start = word['text'].find(target_norm)
            if start >= 0:
                end = start + len(target_norm)
                return {
                    'position': _center_word_boxes([_text_range_box(word, start, end)], offset),
                    'kind': 'single_substring',
                    'word': word,
                }

        if len(valid_words) < 2:
            return None
        indexed_text, char_boxes = _build_char_index(valid_words)
        start = indexed_text.find(target_norm)
        if start >= 0:
            end = start + len(target_norm)
            return {
                'position': _center_word_boxes(char_boxes[start:end], offset),
                'kind': 'chars',
                'word': None,
            }
        return None

_MATCHER = _Matcher()

def _recognize_winocr_words(winocr_module, lang_code, debug, screenshot_pil):
    try:
        res = winocr_module.recognize_pil_sync(screenshot_pil, lang=lang_code)
        if not isinstance(res, dict):
            return []

        words = []
        line_candidates = []
        all_texts = []

        for line in res.get('lines', []):
            line_words = []
            line_originals = []
            for w in line.get('words', []):
                if 'text' in w and 'bounding_rect' in w:
                    text_clean = _normalize_text(w['text'])
                    word = {
                        'text': text_clean,
                        'box': _dict_box_to_tuple(w['bounding_rect']),
                        'original': w['text'],
                        'score': 0.0,
                    }
                    words.append(word)
                    line_words.append(word)
                    line_originals.append(w['text'])
                    all_texts.append(w['text'])
            if line_words:
                line_candidates.append({
                    'text': ''.join(w['text'] for w in line_words),
                    'box': (
                        min(w['box'][0] for w in line_words),
                        min(w['box'][1] for w in line_words),
                        max(w['box'][2] for w in line_words),
                        max(w['box'][3] for w in line_words),
                    ),
                    'original': ' '.join(line_originals),
                    'score': 0.0,
                })

        full_text = ' '.join(all_texts)
        passes = [_recognition_pass(words, full_text, 'WinOCR')]
        if line_candidates:
            passes.append(_recognition_pass(line_candidates, full_text, 'WinOCR', '行匹配'))
        return passes
    except Exception as e:
        _log_engine_error('WinOCR', e, debug)
        return []

def _recognize_rapidocr_words(inst, debug, img_bgr, enhanced_mode=False):
    def run_at_scale(scale):
        h, w = img_bgr.shape[:2]
        if scale != 1:
            img_scaled = cv2.resize(img_bgr, (w * scale, h * scale), interpolation=cv2.INTER_CUBIC)
        else:
            img_scaled = img_bgr

        res = inst(img_scaled)
        all_boxes, all_texts, all_scores = [], [], []

        if isinstance(res, tuple):
            res_list = res[0]
            if res_list:
                for item in res_list:
                    if isinstance(item, (list, tuple)) and len(item) >= 2:
                        all_boxes.append(item[0])
                        all_texts.append(item[1])
                        all_scores.append(item[2] if len(item)>2 else 0.0)
        elif isinstance(res, list):
            for item in res:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    all_boxes.append(item[0])
                    all_texts.append(item[1])
                    all_scores.append(item[2] if len(item)>2 else 0.0)
        else:
            all_boxes = getattr(res, 'boxes', [])
            all_texts = getattr(res, 'txts', [])
            all_scores = getattr(res, 'scores', [])
            if all_boxes is None: all_boxes = getattr(res, 'dt_boxes', [])
            if all_texts is None:
                rec_res = getattr(res, 'rec_res', [])
                if rec_res: all_texts, all_scores = zip(*rec_res)

        if not all_texts or len(all_texts) == 0:
            return None
        if len(all_scores) != len(all_texts):
            all_scores = [0.0] * len(all_texts)

        full_text = ' '.join(all_texts)
        words = []
        for box, text, score in zip(all_boxes, all_texts, all_scores):
            box_tuple = _ocr_box_to_tuple(box, scale)
            if not box_tuple:
                continue
            words.append({
                'text': _normalize_text(text),
                'box': box_tuple,
                'score': score,
                'original': text,
            })
        return _recognition_pass(words, full_text, 'RapidOCR')

    def iter_passes():
        try:
            pixels = img_bgr.shape[0] * img_bgr.shape[1]
            scales = (1, 2) if enhanced_mode and pixels <= RAPIDOCR_ENHANCED_MAX_PIXELS else (1,)
            if debug and enhanced_mode and scales == (1,):
                print(f"  [RapidOCR] 图像过大({pixels} px)，跳过 2x 放大")

            for scale in scales:
                ocr_pass = run_at_scale(scale)
                if ocr_pass:
                    if debug and scale != 1:
                        print(f"  [RapidOCR] 放大 {scale}x 后识别到文本")
                    yield ocr_pass
        except Exception as e:
            _log_engine_error('RapidOCR', e, debug)

    return iter_passes()

def _recognize_tesseract_words(lang, debug, screenshot_pil, enhanced_mode=False):
    def iter_passes():
        img_processed = None
        g = None
        try:
            import pytesseract
            if _TESSERACT_CMD: pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD

            s = 2 if enhanced_mode else 1
            if NUMPY_CV2_AVAILABLE:
                gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
                h, w = gray.shape[:2]
                if s != 1:
                    scaled = cv2.resize(gray, (w*s, h*s), interpolation=cv2.INTER_CUBIC)
                else:
                    scaled = gray
                bw = cv2.adaptiveThreshold(scaled, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                img_processed = Image.fromarray(bw)
            else:
                g = screenshot_pil.convert('L')
                if s != 1:
                    img_processed = g.resize((g.size[0]*s, g.size[1]*s), resample=Image.LANCZOS)
                else:
                    img_processed = g

            config = f'-l {lang}'
            if _TESSERACT_TESSDATA:
                config += f' --tessdata-dir "{_TESSERACT_TESSDATA}"'

            for psm in TESSERACT_PSM_ORDER:
                data = pytesseract.image_to_data(img_processed, config=config + f' --psm {psm}', output_type=pytesseract.Output.DICT)
                words = []
                all_texts = []

                for i in range(len(data['text'])):
                    try:
                        confidence = float(data['conf'][i])
                    except (ValueError, TypeError):
                        confidence = -1

                    if confidence > TESSERACT_MIN_CONFIDENCE and data['text'][i].strip():
                        words.append({
                            'text': _normalize_text(data['text'][i]),
                            'box': _rect_box_to_tuple(data['left'][i], data['top'][i], data['left'][i]+data['width'][i], data['top'][i]+data['height'][i], s),
                            'original': data['text'][i],
                            'score': confidence,
                        })
                        all_texts.append(data['text'][i])

                full_text = ' '.join(all_texts)
                if debug: print(f"  [Tesseract] PSM {psm} 识别 {len(words)} 词")
                yield _recognition_pass(words, full_text, 'Tesseract', f'PSM {psm}')
        except Exception as e:
            _log_engine_error('Tesseract', e, debug)
        finally:
            if img_processed is not None:
                try: img_processed.close()
                except Exception: pass
            if g is not None and g is not screenshot_pil and g is not img_processed:
                try: g.close()
                except Exception: pass

    return iter_passes()

def _is_winocr_available():
    try:
        import winocr
        return True
    except ImportError:
        return False

def _is_rapidocr_available_lightweight():
    return NUMPY_CV2_AVAILABLE and RAPIDOCR_CLASS is not None and not _RAPID_OCR_INIT_FAILED

def get_available_engines():
    engines = []

    if _is_winocr_available() and 'eng' in LANG_MAP['winocr']:
        engines.append(('winocr', 'Windows 10/11 OCR'))

    if _is_rapidocr_available_lightweight() and 'eng' in LANG_MAP['rapidocr']:
        engines.append(('rapidocr', 'RapidOCR (推荐)'))

    if get_tesseract_cmd() and 'eng' in LANG_MAP['tesseract']:
        engines.append(('tesseract', 'Tesseract OCR'))

    if not engines:
        engines.append(('none', '无可用OCR引擎'))

    return engines
