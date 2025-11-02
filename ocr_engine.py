# ocr_engine.py
# 描述：自动化宏的 OCR 功能引擎 (V39 - 拆分版)
# 1. WinOCR (快速, 0MB)
# 2. Tesseract (准确, 5MB)

from PIL import Image, ImageGrab, ImageStat
import re
import os
import subprocess
import io

# ======================================================================
# 【V39】双引擎 OCR 系统 (移除 EasyOCR)
# ======================================================================

# 1. WinOCR (Windows 10/11 内置 - 优先级 1)
try:
    import winocr
    WINOCR_AVAILABLE = True
    print("[配置] ✓ WinOCR 引擎就绪 (优先级 1 - 快速)")
except (ImportError, OSError) as e:
    WINOCR_AVAILABLE = False
    print(f"[配置] ✗ WinOCR 不可用: {e}")

# 2. Tesseract (兜底方案 - 优先级 2)
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
    print("[配置] ✓ Tesseract 引擎就绪 (优先级 2 - 准确)")
except ImportError as e:
    TESSERACT_AVAILABLE = False
    print(f"[配置] ✗ Tesseract 不可用: {e}")

# 检查引擎可用性
if not (WINOCR_AVAILABLE or TESSERACT_AVAILABLE):
    print("[严重错误] 没有任何可用的 OCR 引擎！")
else:
    engines = []
    if WINOCR_AVAILABLE: engines.append("WinOCR")
    if TESSERACT_AVAILABLE: engines.append("Tesseract")
    print(f"[配置] 可用引擎: {' + '.join(engines)} (体积: 0MB)")

# 语言映射
LANG_MAP = {
    'winocr': {'eng': 'en-US', 'chi_sim': 'zh-Hans'},
    'tesseract': {'eng': 'eng', 'chi_sim': 'chi_sim'}
}

# ======================================================================
# Tesseract 配置
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
# 【V39 优化】双引擎文本查找 (高效调度)
# ======================================================================
def find_text_location(target_text, lang='eng', debug=False, region=None):
    """V39: 双引擎调度，WinOCR 优先，失败回退 Tesseract"""
    target_norm = re.sub(r'\s+', '', target_text).lower()
    if not target_norm:
        print("  [错误] 搜索文本为空")
        return None

    # 【优化】优先使用 WinOCR (快速)
    if WINOCR_AVAILABLE and lang in LANG_MAP.get('winocr', {}):
        location = _find_text_winocr(target_norm, LANG_MAP['winocr'][lang], debug, region)
        if location:
            return location
        if debug:
            print(f"  [OCR] WinOCR 未找到，回退 Tesseract...")

    # 【优化】回退 Tesseract (准确)
    if TESSERACT_AVAILABLE and lang in LANG_MAP.get('tesseract', {}):
        return _find_text_tesseract(target_norm, LANG_MAP['tesseract'][lang], debug, region)

    print(f"  [失败] 无可用引擎或未找到 '{target_text}'")
    return None

# ======================================================================
# 【V39 优化】WinOCR 实现 (减少不必要的处理)
# ======================================================================
def _find_text_winocr(target_norm, lang_code, debug, region):
    """V39: WinOCR 优化版，直接处理 words"""
    # 【重构】在函数内部导入 core_engine 以设置全局变量
    import core_engine

    try:
        # 【优化】截图逻辑简化
        if region:
            screenshot = ImageGrab.grab(bbox=(region[0], region[1], 
                                             region[0] + region[2], 
                                             region[1] + region[3]))
            offset = (region[0], region[1])
        else:
            screenshot = ImageGrab.grab()
            offset = (0, 0)

        result = winocr.recognize_pil_sync(screenshot, lang=lang_code)
        if not isinstance(result, dict):
            return None

        # 【优化】直接提取 words (跳过 lines)
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

        if debug:
            print(f"  [WinOCR] 识别 {len(all_words)} 词")

        # 【优化】单词直接匹配 (最常见情况)
        for word in all_words:
            if target_norm in word['text']:
                box = word['box']
                x = offset[0] + box['x'] + box['width'] // 2
                y = offset[1] + box['y'] + box['height'] // 2
                print(f"  [WinOCR✓] ({x}, {y})")
                
                # 【重构】更新 core_engine 的全局变量
                core_engine.last_position = (x, y)
                return (x, y)

        # 【优化】多词合并 (限制为 3 词，减少计算)
        for i in range(len(all_words)):
            merged = all_words[i]['text']
            if not target_norm.startswith(merged):
                continue
            
            if merged == target_norm:
                continue  # 已在上面单词匹配中处理
            
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
                    
                    # 【重构】更新 core_engine 的全局变量
                    core_engine.last_position = (x, y)
                    return (x, y)

        return None
    except Exception as e:
        if debug:
            print(f"  [WinOCR] 错误: {e}")
        return None

# ======================================================================
# 【V39 优化】Tesseract 实现 (智能二值化 + PSM 优化)
# ======================================================================
def _find_text_tesseract(target_norm, lang, debug, region):
    """V39: Tesseract 优化版，智能预处理 + 高效解析"""
    if not _tesseract_cmd:
        return None

    try:
        # 【优化】截图
        if region:
            screenshot = ImageGrab.grab(bbox=(region[0], region[1], 
                                             region[0] + region[2], 
                                             region[1] + region[3]))
            offset = (region[0], region[1])
        else:
            screenshot = ImageGrab.grab()
            offset = (0, 0)

        # 【优化】图像预处理
        screenshot_gray = screenshot.convert('L')
        scale_factor = 2
        w, h = screenshot_gray.size
        resample = Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.ANTIALIAS
        screenshot_scaled = screenshot_gray.resize((w * scale_factor, h * scale_factor), 
                                                   resample=resample)

        # 【优化】智能二值化 (根据亮度自适应)
        try:
            median = ImageStat.Stat(screenshot_scaled).median[0]
            threshold = 130
            if median >= 128:
                screenshot_bw = screenshot_scaled.point(lambda p: 255 if p > threshold else 0, '1')
            else:
                screenshot_bw = screenshot_scaled.point(lambda p: 0 if p > threshold else 255, '1')
        except:
            screenshot_bw = screenshot_scaled

        # 【优化】转 bytes
        img_byte_arr = io.BytesIO()
        screenshot_bw.save(img_byte_arr, format='PNG')
        img_bytes = img_byte_arr.getvalue()

        # 【优化】Tesseract 命令 (PSM 6 = 统一文本块)
        command = [
            _tesseract_cmd, 'stdin', 'stdout',
            '--tessdata-dir', _local_tessdata_path,
            '-l', lang, '--psm', '6', 'tsv'
        ]
        command = [c for c in command if c]

        result = subprocess.run(command, input=img_bytes, capture_output=True, 
                              check=True, creationflags=_subprocess_creationflags)
        output = result.stdout.decode('utf-8', errors='ignore')

        # 【优化】快速解析 TSV (只提取 level=5 的单词)
        all_words = []
        for line in output.strip().split('\n')[1:]:  # 跳过标题行
            parts = line.split('\t')
            if len(parts) != 12:
                continue
            
            try:
                level = int(parts[0])
                if level != 5:  # 只要单词级别
                    continue
                
                text = parts[11].strip()
                conf = float(parts[10])
                if not text or conf <= 30:  # 降低阈值从 40 到 30
                    continue
                
                x, y, w, h = int(parts[6]), int(parts[7]), int(parts[8]), int(parts[9])
                if w <= 0 or h <= 0:
                    continue
                
                # 坐标转换 (缩放 + 偏移)
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

        if debug:
            print(f"  [Tesseract] 识别 {len(all_words)} 词")

        # 【优化】使用通用合并函数
        return _merge_and_match(all_words, target_norm, debug)

    except Exception as e:
        if debug:
            print(f"  [Tesseract] 错误: {e}")
        return None

# ======================================================================
# 【V39 优化】通用多词合并匹配 (减少重复代码)
# ======================================================================
def _merge_and_match(all_words, target_norm, debug):
    """V39: 高效的多词合并算法"""
    # 【重构】在函数内部导入 core_engine 以设置全局变量
    import core_engine

    # 单词直接匹配
    for word in all_words:
        if target_norm == word['text']:
            box = word['box']
            x = (box[0] + box[2]) // 2
            y = (box[1] + box[3]) // 2
            print(f"  [Tesseract✓] ({x}, {y})")

            # 【重构】更新 core_engine 的全局变量
            core_engine.last_position = (x, y)
            return (x, y)

    # 多词合并 (最多 5 个词)
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
                
                # 【重构】更新 core_engine 的全局变量
                core_engine.last_position = (avg_x, avg_y)
                return (avg_x, avg_y)
    
    return None