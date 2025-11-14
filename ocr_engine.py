# ocr_engine.py
# 描述：自动化宏的 OCR 功能引擎
# 版本：1.49.1
# 变更：版本号同步

from PIL import Image, ImageGrab
import re
import os
import subprocess
import io
import time
import sys
import threading

# ======================================================================
# 依赖库预加载
# ======================================================================
RAPIDOCR_CLASS = None
NUMPY_CV2_AVAILABLE = False

try:
    import numpy as np
    import cv2
    NUMPY_CV2_AVAILABLE = True
    from rapidocr import RapidOCR
    RAPIDOCR_CLASS = RapidOCR 
except Exception as e:
    pass # 预加载失败不打印，等到真正使用时再报错

# ======================================================================
# 全局状态缓存
# ======================================================================
_RAPID_OCR_INSTANCE = None
_RAPID_OCR_INIT_FAILED = False
_RAPID_OCR_LOCK = threading.Lock() # 确保预热线程安全

_TESSERACT_CMD = None
_TESSERACT_TESSDATA = None
_TESSERACT_CHECKED = False
_TESSERACT_LOCK = threading.Lock()

# ======================================================================
# 懒加载与预热实现
# ======================================================================
def preload_engines():
    """静默预热所有可用的 OCR 引擎"""
    print("[OCR] 后台预热开始...")
    if NUMPY_CV2_AVAILABLE:
        get_rapid_ocr_engine()
    get_tesseract_cmd()
    # WinOCR 不需要预热

def get_rapid_ocr_engine():
    global _RAPID_OCR_INSTANCE, _RAPID_OCR_INIT_FAILED
    if _RAPID_OCR_INSTANCE: return _RAPID_OCR_INSTANCE
    if _RAPID_OCR_INIT_FAILED or not RAPIDOCR_CLASS: return None
    
    with _RAPID_OCR_LOCK: # 加锁防止多线程同时初始化
        if _RAPID_OCR_INSTANCE: return _RAPID_OCR_INSTANCE
        try:
            print("[OCR] 正在加载 RapidOCR 模型...")
            t0 = time.time()
            _RAPID_OCR_INSTANCE = RAPIDOCR_CLASS()
            print(f"[OCR] RapidOCR 就绪 ({time.time()-t0:.2f}s)")
            return _RAPID_OCR_INSTANCE
        except Exception as e:
            print(f"[严重错误] RapidOCR 加载失败: {e}")
            _RAPID_OCR_INIT_FAILED = True
            return None

def get_tesseract_cmd():
    global _TESSERACT_CMD, _TESSERACT_TESSDATA, _TESSERACT_CHECKED
    if _TESSERACT_CHECKED: return _TESSERACT_CMD
    
    with _TESSERACT_LOCK:
        if _TESSERACT_CHECKED: return _TESSERACT_CMD
        _TESSERACT_CHECKED = True
        
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
                if os.path.exists(data): _TESSERACT_TESSDATA = data
                break
        
        if not _TESSERACT_CMD:
            try:
                cflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                find_cmd = 'where' if os.name == 'nt' else 'which'
                res = subprocess.run([find_cmd, 'tesseract'], capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=cflags)
                if res.returncode == 0: _TESSERACT_CMD = res.stdout.strip().split('\n')[0]
            except: pass

        if _TESSERACT_CMD:
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
                # print(f"[OCR] Tesseract 就绪")
            except: _TESSERACT_CMD = None
            
        return _TESSERACT_CMD

# ======================================================================
# 引擎状态
# ======================================================================
LANG_MAP = {
    'winocr': {'eng': 'en-US', 'chi_sim': 'zh-Hans'},
    'rapidocr': {'eng': 'en', 'chi_sim': 'ch'},
    'tesseract': {'eng': 'eng', 'chi_sim': 'chi_sim'}
}

class OCRPerformanceStats:
    def __init__(self): self.reset()
    def reset(self):
        self.stats = {'winocr': [0,0], 'rapidocr': [0,0], 'tesseract': [0,0]}
        self.total_time = 0; self.call_count = 0
    def record(self, engine, success, duration):
        self.call_count += 1; self.total_time += duration
        self.stats[engine][0 if success else 1] += 1
    def get_stats(self):
        if self.call_count == 0: return "无 OCR 统计"
        avg = (self.total_time / self.call_count) * 1000
        parts = []
        for eng, (succ, fail) in self.stats.items():
            if succ + fail > 0: parts.append(f"{eng}({succ/(succ+fail)*100:.0f}%)")
        return f"OCR统计 (均{avg:.0f}ms): {' | '.join(parts)}"

ocr_stats = OCRPerformanceStats()

# ======================================================================
# 统一查找入口
# ======================================================================
def find_text_location(target_text, lang='eng', debug=False, screenshot_pil=None, offset=(0,0), engine='auto'):
    target_norm = re.sub(r'\s+', '', target_text).lower()
    if not target_norm: return None
    if screenshot_pil is None: screenshot_pil = ImageGrab.grab(); offset = (0, 0)
    img_bgr_cache = None 
    def get_img_bgr():
        nonlocal img_bgr_cache
        if img_bgr_cache is None: img_bgr_cache = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
        return img_bgr_cache

    if engine == 'auto':
        try:
            import winocr
            if lang in LANG_MAP['winocr']:
                t0 = time.time()
                loc = _find_text_winocr(winocr, target_norm, LANG_MAP['winocr'][lang], debug, screenshot_pil, offset)
                ocr_stats.record('winocr', loc is not None, time.time() - t0)
                if loc: return loc
        except ImportError: pass
        
        rapid_inst = get_rapid_ocr_engine()
        if rapid_inst and lang in LANG_MAP['rapidocr']:
            t0 = time.time()
            loc = _find_text_rapidocr_internal(rapid_inst, target_norm, debug, get_img_bgr(), offset)
            ocr_stats.record('rapidocr', loc is not None, time.time() - t0)
            if loc: return loc

        if get_tesseract_cmd() and lang in LANG_MAP['tesseract']:
            t0 = time.time()
            loc = _find_text_tesseract(target_norm, LANG_MAP['tesseract'][lang], debug, screenshot_pil, offset)
            ocr_stats.record('tesseract', loc is not None, time.time() - t0)
            if loc: return loc

    elif engine == 'rapidocr':
        rapid_inst = get_rapid_ocr_engine()
        if rapid_inst and lang in LANG_MAP['rapidocr']:
            t0 = time.time()
            loc = _find_text_rapidocr_internal(rapid_inst, target_norm, debug, get_img_bgr(), offset)
            ocr_stats.record('rapidocr', loc is not None, time.time() - t0)
            if loc: return loc

    elif engine == 'winocr':
        try:
            import winocr
            if lang in LANG_MAP['winocr']:
                t0 = time.time()
                loc = _find_text_winocr(winocr, target_norm, LANG_MAP['winocr'][lang], debug, screenshot_pil, offset)
                ocr_stats.record('winocr', loc is not None, time.time() - t0)
                if loc: return loc
        except ImportError: pass

    elif engine == 'tesseract':
        if get_tesseract_cmd() and lang in LANG_MAP['tesseract']:
            t0 = time.time()
            loc = _find_text_tesseract(target_norm, LANG_MAP['tesseract'][lang], debug, screenshot_pil, offset)
            ocr_stats.record('tesseract', loc is not None, time.time() - t0)
            if loc: return loc

    print(f"  [失败] 未能找到 '{target_text}' (模式: {engine})")
    if debug: print(f"  [统计] {ocr_stats.get_stats()}")
    return None

# --- 具体实现函数 ---
def _find_text_winocr(winocr_module, target_norm, lang_code, debug, screenshot_pil, offset):
    try:
        res = winocr_module.recognize_pil_sync(screenshot_pil, lang=lang_code)
        if not isinstance(res, dict): return None
        words = []
        for line in res.get('lines', []):
            for w in line.get('words', []):
                if 'text' in w and 'bounding_rect' in w:
                    words.append({'text': re.sub(r'\s+','',w['text']).lower(), 'box': w['bounding_rect']})
        for w in words:
            if target_norm in w['text']:
                b = w['box']; cx, cy = offset[0]+b['x']+b['width']//2, offset[1]+b['y']+b['height']//2
                if debug: print(f"  [WinOCR✓] ({cx}, {cy})")
                return (cx, cy)
        for i in range(len(words)):
            merged = words[i]['text']
            if not target_norm.startswith(merged): continue
            b_list = [words[i]['box']]
            for j in range(i+1, min(i+5, len(words))):
                merged += words[j]['text']; b_list.append(words[j]['box'])
                if target_norm == merged:
                    cx = offset[0] + sum(b['x']+b['width']//2 for b in b_list)//len(b_list)
                    cy = offset[1] + sum(b['y']+b['height']//2 for b in b_list)//len(b_list)
                    if debug: print(f"  [WinOCR✓] 合并 ({cx}, {cy})")
                    return (cx, cy)
        return None
    except: return None

def _find_text_rapidocr_internal(inst, target_norm, debug, img_bgr, offset):
    try:
        res = inst(img_bgr)
        all_boxes, all_texts, all_scores = [], [], []
        
        if isinstance(res, tuple):
            res_list = res[0]
            if res_list:
                for item in res_list:
                    if len(item) >= 2:
                        all_boxes.append(item[0]); all_texts.append(item[1])
                        all_scores.append(item[2] if len(item)>2 else 0.0)
        elif isinstance(res, list):
             for item in res:
                if len(item) >= 2:
                    all_boxes.append(item[0]); all_texts.append(item[1])
                    all_scores.append(item[2] if len(item)>2 else 0.0)
        else:
            all_boxes = getattr(res, 'boxes', [])
            all_texts = getattr(res, 'txts', [])
            all_scores = getattr(res, 'scores', [])
            if all_boxes is None: all_boxes = getattr(res, 'dt_boxes', [])
            if all_texts is None:
                 rec_res = getattr(res, 'rec_res', [])
                 if rec_res: all_texts, all_scores = zip(*rec_res)

        if not all_texts or len(all_texts) == 0: return None
        if len(all_scores) != len(all_texts): all_scores = [0.0] * len(all_texts)

        words = []
        for box, text, score in zip(all_boxes, all_texts, all_scores):
            if not isinstance(box, (list, np.ndarray)): continue
            xs = [p[0] for p in box]; ys = [p[1] for p in box]
            words.append({'text': re.sub(r'\s+','',text).lower(), 'box': [min(xs), min(ys), max(xs), max(ys)], 'score': score})

        for w in words:
            if target_norm in w['text']:
                b = w['box']; cx, cy = offset[0]+(b[0]+b[2])//2, offset[1]+(b[1]+b[3])//2
                if debug: print(f"  [RapidOCR✓] ({cx}, {cy}) @ {w['score']:.2f}")
                return (cx, cy)
        for i in range(len(words)):
            merged = words[i]['text']
            if not target_norm.startswith(merged): continue
            b_list = [words[i]['box']]
            for j in range(i+1, min(i+5, len(words))):
                merged += words[j]['text']; b_list.append(words[j]['box'])
                if target_norm == merged:
                    cx = offset[0] + sum((b[0]+b[2])//2 for b in b_list)//len(b_list)
                    cy = offset[1] + sum((b[1]+b[3])//2 for b in b_list)//len(b_list)
                    if debug: print(f"  [RapidOCR✓] 合并 ({cx}, {cy})")
                    return (cx, cy)
        return None
    except Exception as e:
        if debug: print(f"  [RapidOCR] 解析错误: {e}")
        return None

def _find_text_tesseract(target_norm, lang, debug, screenshot_pil, offset):
    try:
        import pytesseract
        if _TESSERACT_CMD: pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
        if NUMPY_CV2_AVAILABLE:
            gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
            h, w = gray.shape[:2]; s = 2
            scaled = cv2.resize(gray, (w*s, h*s), interpolation=cv2.INTER_CUBIC)
            bw = cv2.adaptiveThreshold(scaled, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
            img_processed = Image.fromarray(bw)
        else:
            s = 2; g = screenshot_pil.convert('L')
            img_processed = g.resize((g.size[0]*s, g.size[1]*s), resample=Image.LANCZOS)
        config = f'-l {lang}'
        if _TESSERACT_TESSDATA: config += f' --tessdata-dir "{_TESSERACT_TESSDATA}"'
        for psm in [6, 11, 3]:
            data = pytesseract.image_to_data(img_processed, config=config + f' --psm {psm}', output_type=pytesseract.Output.DICT)
            words = []
            for i in range(len(data['text'])):
                if int(data['conf'][i]) > 30 and data['text'][i].strip():
                    words.append({'text': re.sub(r'\s+','',data['text'][i]).lower(), 'box': [data['left'][i]//s, data['top'][i]//s, (data['left'][i]+data['width'][i])//s, (data['top'][i]+data['height'][i])//s]})
            for w in words:
                if target_norm in w['text']:
                    b = w['box']; cx, cy = offset[0]+(b[0]+b[2])//2, offset[1]+(b[1]+b[3])//2
                    if debug: print(f"  [Tesseract✓] ({cx}, {cy})")
                    return (cx, cy)
            for i in range(len(words)):
                if not target_norm.startswith(words[i]['text']): continue
                m = words[i]['text']; b_list = [words[i]['box']]
                for j in range(i+1, min(i+5, len(words))):
                    m += words[j]['text']; b_list.append(words[j]['box'])
                    if target_norm == m:
                        ax = sum((b[0]+b[2])//2 for b in b_list)//len(b_list)
                        ay = sum((b[1]+b[3])//2 for b in b_list)//len(b_list)
                        print(f"  [Tesseract✓] 合并{len(b_list)}词 ({offset[0]+ax}, {offset[1]+ay})")
                        return (offset[0]+ax, offset[1]+ay)
        return None
    except Exception as e: return None

ocr_engine_version = "1.49.1"