# ocr_engine.py
# 描述：自动化宏的 OCR 功能引擎 (V43.3 - Tesseract 多-PSM 修复版)
# 修复内容:
# 1. (Grok P1) _find_text_tesseract 现在会尝试多种 PSM (6, 11, 3) 以提高准确率。
# 2. (V43.2) 保持 Tesseract CV2/BMP 优化 和 RapidOCR 健壮性修复。
# 3. (V43.1) 保持 V43.1 的所有修复 (懒加载, V44 架构, 语法修复)。

from PIL import Image, ImageGrab, ImageStat
import re
import os
import subprocess
import io
import time

# 导入 RapidOCR/CV2 所需的库
try:
    import numpy as np
    import cv2
    NUMPY_CV2_AVAILABLE = True
except ImportError:
    NUMPY_CV2_AVAILABLE = False
    print("[配置] ✗ 缺少 numpy 或 opencv-python，RapidOCR 将被禁用。")


# ======================================================================
# 【V43 修复】RapidOCR 懒加载 (Lazy Loading)
# ======================================================================

_rapid_ocr_engine = None # 默认为 None，在第一次使用时才初始化

def init_rapid_ocr():
    """初始化 RapidOCR 引擎（懒加载）"""
    global _rapid_ocr_engine
    global RAPIDOCR_AVAILABLE 

    if _rapid_ocr_engine is None and RAPIDOCR_AVAILABLE:
        try:
            from rapidocr import RapidOCR
            _rapid_ocr_engine = RapidOCR()
            print("[配置] ✓ RapidOCR 引擎懒加载... 成功")
            return True
        except (ImportError, OSError) as e:
            RAPIDOCR_AVAILABLE = False
            print(f"[配置] ✗ RapidOCR 懒加载失败: {e}")
            print("[配置]   请尝试以下步骤：")
            print("     1. 确保已安装: pip install rapidocr onnxruntime")
            print("     2. 下载 VC++ 运行库: https://aka.ms/vs/17/release/vc_redist.x64.exe")
            print("     3. 重启电脑后重试")
            return False
    if not RAPIDOCR_AVAILABLE:
        return False
    return _rapid_ocr_engine is not None

# 2. RapidOCR (深度学习模型 - 优先级 2)
if NUMPY_CV2_AVAILABLE:
    RAPIDOCR_AVAILABLE = True
    print("[配置] ✓ RapidOCR 引擎已准备 (懒加载模式)")
else:
    RAPIDOCR_AVAILABLE = False
    print("[配置] ✗ RapidOCR 禁用 (缺少 numpy 或 opencv-python)")

# ======================================================================
# 【V42】加载其他引擎
# ======================================================================

# 1. WinOCR (Windows 10/11 内置 - 优先级 1)
try:
    import winocr
    WINOCR_AVAILABLE = True
    print("[配置] ✓ WinOCR 引擎就绪 (优先级 1 - 极快)")
except (ImportError, OSError) as e:
    WINOCR_AVAILABLE = False
    print(f"[配置] ✗ WinOCR 不可用: {e}")

# 3. Tesseract (兜底方案 - 优先级 3)
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
    print("[配置] ✓ Tesseract 引擎就绪 (优先级 3 - 兜底)")
except ImportError as e:
    TESSERACT_AVAILABLE = False
    print(f"[配置] ✗ Tesseract 不可用: {e}")


# 检查引擎可用性
if not (WINOCR_AVAILABLE or RAPIDOCR_AVAILABLE or TESSERACT_AVAILABLE):
    print("[严重错误] 没有任何可用的 OCR 引擎！")
else:
    engines = []
    if WINOCR_AVAILABLE: engines.append("WinOCR(0MB)")
    if RAPIDOCR_AVAILABLE: engines.append("RapidOCR(10MB)")
    if TESSERACT_AVAILABLE: engines.append("Tesseract(5MB)")
    total_size = (10 if RAPIDOCR_AVAILABLE else 0) + (5 if TESSERACT_AVAILABLE else 0)
    print(f"[配置] 可用引擎: {' + '.join(engines)} (总计: {total_size}MB)")

# 语言映射
LANG_MAP = {
    'winocr': {'eng': 'en-US', 'chi_sim': 'zh-Hans'},
    'rapidocr': {'eng': 'en', 'chi_sim': 'ch'},
    'tesseract': {'eng': 'eng', 'chi_sim': 'chi_sim'}
}

# ======================================================================
# Tesseract 配置 (与 V42 相同)
# ======================================================================
_subprocess_creationflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
_basedir = os.path.dirname(__file__)
_local_tesseract_path = os.path.join(_basedir, 'tesseract_local', 'tesseract.exe')
_local_tessdata_path = os.path.join(_basedir, 'tesseract_local', 'tessdata')
_tesseract_cmd = ''

if TESSERACT_AVAILABLE:
    if os.path.exists(_local_tesseract_path):
        _tesseract_cmd = _local_tesseract_path
        pytesseract.pytesseract.tesseract_cmd = _local_tesseract_path
    else:
        try:
            find_cmd = 'where' if os.name == 'nt' else 'which'
            result = subprocess.run([find_cmd, 'tesseract'], capture_output=True, 
                                  text=True, check=True, encoding='utf-8', 
                                  errors='ignore', creationflags=_subprocess_creationflags)
            _tesseract_cmd = result.stdout.strip().split('\n')[0]
            pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd
        except:
            _tesseract_cmd = ''
            TESSERACT_AVAILABLE = False

# ======================================================================
# 【V42】性能统计 (与 V42 相同)
# ======================================================================
class OCRPerformanceStats:
    def __init__(self): self.reset()
    def reset(self):
        self.winocr_success = 0; self.winocr_fail = 0
        self.rapidocr_success = 0; self.rapidocr_fail = 0
        self.tesseract_success = 0; self.tesseract_fail = 0
        self.total_time = 0; self.call_count = 0
    def record(self, engine, success, duration):
        self.call_count += 1; self.total_time += duration
        if engine == 'winocr':
            if success: self.winocr_success += 1
            else: self.winocr_fail += 1
        elif engine == 'rapidocr':
            if success: self.rapidocr_success += 1
            else: self.rapidocr_fail += 1
        elif engine == 'tesseract':
            if success: self.tesseract_success += 1
            else: self.tesseract_fail += 1
    def get_stats(self):
        if self.call_count == 0: return "无 OCR 统计"
        avg_time = (self.total_time / self.call_count) * 1000
        stats = f"OCR统计 (平均{avg_time:.0f}ms): "
        parts = []
        if self.winocr_success + self.winocr_fail > 0:
            rate = self.winocr_success / (self.winocr_success + self.winocr_fail) * 100
            parts.append(f"WinOCR({rate:.0f}%)")
        if self.rapidocr_success + self.rapidocr_fail > 0:
            rate = self.rapidocr_success / (self.rapidocr_success + self.rapidocr_fail) * 100
            parts.append(f"RapidOCR({rate:.0f}%)")
        if self.tesseract_success + self.tesseract_fail > 0:
            rate = self.tesseract_success / (self.tesseract_success + self.tesseract_fail) * 100
            parts.append(f"Tesseract({rate:.0f}%)")
        return stats + " | ".join(parts)

ocr_stats = OCRPerformanceStats()
# ======================================================================
# 【V43 修复】统一截图流程 (V44 架构)
# ======================================================================
def find_text_location(target_text, lang='eng', debug=False, 
                       screenshot_pil=None, offset=(0,0), engine='auto'):
    """V43.0: V44 架构，接收截图 (screenshot_pil) 和偏移量 (offset)"""
    target_norm = re.sub(r'\s+', '', target_text).lower()
    if not target_norm:
        print("  [错误] 搜索文本为空")
        return None
        
    if screenshot_pil is None:
        print("  [警告] ocr_engine 进行了紧急截图，这不符合 V44 架构。")
        screenshot_pil = ImageGrab.grab()
        offset = (0, 0)

    img_bgr_cache = None 

    if engine == 'auto':
        if WINOCR_AVAILABLE and lang in LANG_MAP.get('winocr', {}):
            start_time = time.time()
            location = _find_text_winocr(target_norm, LANG_MAP['winocr'][lang], debug, screenshot_pil, offset)
            duration = time.time() - start_time
            if location:
                ocr_stats.record('winocr', True, duration)
                return location
            ocr_stats.record('winocr', False, duration)
            if debug: print(f"  [OCR] WinOCR 未找到...")

        if RAPIDOCR_AVAILABLE and lang in LANG_MAP.get('rapidocr', {}):
            if img_bgr_cache is None:
                img_bgr_cache = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
            start_time = time.time()
            location = _find_text_rapidocr(target_norm, LANG_MAP['rapidocr'][lang], debug, img_bgr_cache, offset)
            duration = time.time() - start_time
            if location:
                ocr_stats.record('rapidocr', True, duration)
                return location
            ocr_stats.record('rapidocr', False, duration)
            if debug: print(f"  [OCR] RapidOCR 未找到...")

        if TESSERACT_AVAILABLE and lang in LANG_MAP.get('tesseract', {}):
            start_time = time.time()
            location = _find_text_tesseract(target_norm, LANG_MAP['tesseract'][lang], debug, screenshot_pil, offset)
            duration = time.time() - start_time
            if location:
                ocr_stats.record('tesseract', True, duration)
                return location
            ocr_stats.record('tesseract', False, duration)

    elif engine == 'winocr':
        if WINOCR_AVAILABLE and lang in LANG_MAP.get('winocr', {}):
            print("  [引擎] 强制使用 WinOCR...")
            start_time = time.time()
            location = _find_text_winocr(target_norm, LANG_MAP['winocr'][lang], debug, screenshot_pil, offset)
            duration = time.time() - start_time
            ocr_stats.record('winocr', location is not None, duration)
            if location: return location
        else: print(f"  [错误] WinOCR 不可用或不支持语言 {lang}")

    elif engine == 'rapidocr':
        if RAPIDOCR_AVAILABLE and lang in LANG_MAP.get('rapidocr', {}):
            print("  [引擎] 强制使用 RapidOCR...")
            if img_bgr_cache is None:
                img_bgr_cache = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2BGR)
            start_time = time.time()
            location = _find_text_rapidocr(target_norm, LANG_MAP['rapidocr'][lang], debug, img_bgr_cache, offset)
            duration = time.time() - start_time
            ocr_stats.record('rapidocr', location is not None, duration)
            if location: return location
        else: print(f"  [错误] RapidOCR 不可用或不支持语言 {lang}")

    elif engine == 'tesseract':
        if TESSERACT_AVAILABLE and lang in LANG_MAP.get('tesseract', {}):
            print("  [引擎] 强制使用 Tesseract...")
            start_time = time.time()
            location = _find_text_tesseract(target_norm, LANG_MAP['tesseract'][lang], debug, screenshot_pil, offset)
            duration = time.time() - start_time
            ocr_stats.record('tesseract', location is not None, duration)
            if location: return location
        else: print(f"  [错误] Tesseract 不可用或不支持语言 {lang}")
    
    else:
        print(f"  [错误] 未知的引擎名称: '{engine}'。")

    print(f"  [失败] 未能找到 '{target_text}' (模式: {engine})")
    if debug: print(f"  [统计] {ocr_stats.get_stats()}")
    return None

# ======================================================================
# 【V43 修复】WinOCR 实现 (接收截图)
# ======================================================================
def _find_text_winocr(target_norm, lang_code, debug, screenshot_pil, offset):
    try:
        result = winocr.recognize_pil_sync(screenshot_pil, lang=lang_code)
        if not isinstance(result, dict): return None

        all_words = []
        for line in result.get('lines', []):
            for word in line.get('words', []):
                if isinstance(word, dict):
                    text = word.get('text', '').strip()
                    box = word.get('bounding_rect')
                    if text and box:
                        all_words.append({
                            'text': re.sub(r'\s+', '', text).lower(),
                            'box': box
                        })
        if debug: print(f"  [WinOCR] 识别 {len(all_words)} 词")

        for word in all_words:
            if target_norm in word['text']:
                box = word['box']
                x = offset[0] + box['x'] + box['width'] // 2
                y = offset[1] + box['y'] + box['height'] // 2
                print(f"  [WinOCR✓] ({x}, {y})")
                return (x, y)

        for i in range(len(all_words)):
            merged = all_words[i]['text']
            if not target_norm.startswith(merged): continue
            boxes = [all_words[i]['box']]
            for j in range(i + 1, min(i + 3, len(all_words))):
                merged += all_words[j]['text']
                boxes.append(all_words[j]['box'])
                if target_norm == merged:
                    avg_x = sum(b['x'] + b['width']//2 for b in boxes) // len(boxes)
                    avg_y = sum(b['y'] + b['height']//2 for b in boxes) // len(boxes)
                    x = offset[0] + avg_x
                    y = offset[1] + avg_y
                    print(f"  [WinOCR✓] 合并{len(boxes)}词 ({x}, {y})")
                    return (x, y)
        return None
    except Exception as e:
        if debug: print(f"  [WinOCR] 错误: {e}")
        return None

# ======================================================================
# 【V43.2 修复】RapidOCR 实现 (Grok P4 - 健壮性)
# ======================================================================
def _find_text_rapidocr(target_norm, lang_code, debug, img_bgr, offset):
    """V43.2: 修复 RapidOCR 输出格式健壮性"""
    
    try:
        if not init_rapid_ocr():
             if debug: print(f"  [RapidOCR] 错误: 引擎初始化失败。")
             return None

        ocr_output = _rapid_ocr_engine(img_bgr)
        
        try:
            all_boxes = getattr(ocr_output, 'boxes', [])
            all_texts = getattr(ocr_output, 'txts', [])
            all_scores = getattr(ocr_output, 'scores', [])
            elapse = getattr(ocr_output, 'elapse', 0.0)
        except Exception as e:
            if debug:
                print(f"  [RapidOCR] 错误: 无法解析输出对象. {e}")
                print(f"  [RapidOCR] 对象属性: {dir(ocr_output)}")
            return None
        
        if not (all_texts and all_boxes and len(all_texts) == len(all_boxes)):
            if debug:
                print(f"  [RapidOCR] 警告: 输出列表长度不匹配或为空。")
                print(f"    Boxes: {len(all_boxes)}, Texts: {len(all_texts)}")
            return None
        
        if not all_scores or len(all_scores) != len(all_texts):
            all_scores = [0.0] * len(all_texts)
            
        all_words = []
        for box, text, score in zip(all_boxes, all_texts, all_scores):
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            x1, x2 = min(xs), max(xs)
            y1, y2 = min(ys), max(ys)
            
            all_words.append({
                'text': re.sub(r'\s+', '', text).lower(),
                'box': [x1, y1, x2, y2],
                'score': score
            })

        if debug:
            print(f"  [RapidOCR] 识别 {len(all_words)} 词 (耗时 {elapse:.2f}s)")

        for word in all_words:
            if target_norm in word['text']:
                box = word['box']
                x = offset[0] + (box[0] + box[2]) // 2
                y = offset[1] + (box[1] + box[3]) // 2
                print(f"  [RapidOCR✓] ({x}, {y}) @ {word['score']:.2f}")
                return (x, y)

        for i in range(len(all_words)):
            merged = all_words[i]['text']
            if not target_norm.startswith(merged):
                continue
            boxes = [all_words[i]['box']]
            for j in range(i + 1, min(i + 5, len(all_words))):
                merged += all_words[j]['text']
                boxes.append(all_words[j]['box'])
                
                if target_norm == merged:
                    avg_x = sum((b[0] + b[2]) // 2 for b in boxes) // len(boxes)
                    avg_y = sum((b[1] + b[3]) // 2 for b in boxes) // len(boxes)
                    x = offset[0] + avg_x
                    y = offset[1] + avg_y
                    print(f"  [RapidOCR✓] 合并{len(boxes)}词 ({x}, {y})")
                    return (x, y)

        return None
    except Exception as e:
        if debug:
            print(f"  [RapidOCR] 错误: {e}")
        return None

# ======================================================================
# 【V43.3 修复】Tesseract 实现 (Grok P1 - 多 PSM)
# ======================================================================
def _find_text_tesseract(target_norm, lang, debug, screenshot_pil, offset):
    if not _tesseract_cmd:
        return None
    try:
        if NUMPY_CV2_AVAILABLE:
            img_np_gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
            scale_factor = 2
            w, h = img_np_gray.shape[1], img_np_gray.shape[0]
            img_scaled = cv2.resize(img_np_gray, (w * scale_factor, h * scale_factor), 
                                    interpolation=cv2.INTER_CUBIC)
            img_bw = cv2.adaptiveThreshold(img_scaled, 255, 
                                           cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                           cv2.THRESH_BINARY, 11, 2)
            screenshot_bw = Image.fromarray(img_bw)
        else:
            # (V43.2) 回退到旧的 PIL 逻辑
            screenshot_gray = screenshot_pil.convert('L')
            scale_factor = 2
            w, h = screenshot_gray.size
            resample = Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.ANTIALIAS
            screenshot_scaled = screenshot_gray.resize((w * scale_factor, h * scale_factor), 
                                                    resample=resample)
            try:
                median = ImageStat.Stat(screenshot_scaled).median[0]
                threshold = 130
                if median >= 128:
                    screenshot_bw = screenshot_scaled.point(lambda p: 255 if p > threshold else 0, '1')
                else:
                    screenshot_bw = screenshot_scaled.point(lambda p: 0 if p > threshold else 255, '1')
            except:
                screenshot_bw = screenshot_scaled
        
        img_byte_arr = io.BytesIO()
        screenshot_bw.save(img_byte_arr, format='BMP') # (V43.2) 
        img_bytes = img_byte_arr.getvalue()
        
        # 【V43.3 修复】 (Grok P1) 尝试多种 PSM 模式
        # 6: 假设为单个统一的文本块。
        # 11: 查找稀疏文本。
        # 3: 全自动页面分割 (默认)。
        for psm in [6, 11, 3]:
            if debug: print(f"  [Tesseract] 尝试 PSM 模式: {psm}")
            
            command = [
                _tesseract_cmd, 'stdin', 'stdout',
                '--tessdata-dir', _local_tessdata_path,
                '-l', lang, '--psm', str(psm), 'tsv'
            ]
            command = [c for c in command if c]
            result = subprocess.run(command, input=img_bytes, capture_output=True, 
                                  check=True, creationflags=_subprocess_creationflags)
            output = result.stdout.decode('utf-8', errors='ignore')

            all_words = []
            for line in output.strip().split('\n')[1:]:
                parts = line.split('\t')
                if len(parts) != 12: continue
                try:
                    level = int(parts[0])
                    if level != 5: continue
                    text = parts[11].strip()
                    conf = float(parts[10])
                    if not text or conf <= 30: continue
                    x, y, w, h = int(parts[6]), int(parts[7]), int(parts[8]), int(parts[9])
                    if w <= 0 or h <= 0: continue
                    
                    ox = offset[0] + x // scale_factor
                    oy = offset[1] + y // scale_factor
                    ow = w // scale_factor
                    oh = h // scale_factor
                    
                    all_words.append({
                        'text': re.sub(r'\s+', '', text).lower(),
                        'box': [ox, oy, ox + ow, oy + oh]
                    })
                except (ValueError, IndexError):
                    continue
            
            if debug and not all_words:
                print(f"  [Tesseract] PSM {psm} 未识别到单词。")
                continue # 尝试下一种 PSM 模式
                
            if debug: print(f"  [Tesseract] PSM {psm} 识别 {len(all_words)} 词")
            
            location = _merge_and_match(all_words, target_norm, debug)
            if location:
                return location # 匹配成功，立即返回
        
        return None # 所有 PSM 模式均失败
        
    except Exception as e:
        if debug: print(f"  [Tesseract] 错误: {e}")
        return None

# ======================================================================
# 【V43 修复】通用多词合并匹配 (使用绝对坐标)
# ======================================================================
def _merge_and_match(all_words, target_norm, debug):
    """V43.0: Tesseract 的合并函数，现在 all_words 包含绝对坐标"""
    for word in all_words:
        if target_norm in word['text']:
            box = word['box']
            x = (box[0] + box[2]) // 2
            y = (box[1] + box[3]) // 2
            print(f"  [Tesseract✓] ({x}, {y})")
            return (x, y)

    for i in range(len(all_words)):
        if not target_norm.startswith(all_words[i]['text']):
            continue
        merged = all_words[i]['text']
        boxes = [all_words[i]['box']]
        for j in range(i + 1, min(i + 5, len(all_words))):
            merged += all_words[j]['text']
            boxes.append(all_words[j]['box'])
            if target_norm == merged:
                avg_x = sum((b[0] + b[2]) // 2 for b in boxes) // len(boxes)
                avg_y = sum((b[1] + b[3]) // 2 for b in boxes) // len(boxes)
                print(f"  [Tesseract✓] 合并{len(boxes)}词 ({avg_x}, {avg_y})")
                return (avg_x, avg_y)
    
    return None

# 【V43.3】版本号
ocr_engine_version = "V43.3 (V44 架构 - 多PSM/CV2/BMP 优化 - RapidOCR 健壮性修复)"