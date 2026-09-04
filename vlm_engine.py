# -*- coding: utf-8 -*-
# vlm_engine.py
# 功能说明：视觉语言模型接口，负责配置管理、屏幕编码、API 适配与坐标解析
# Version: 1.8.6
# 处理流程：将屏幕截图编码为 Base64，随自然语言指令发送给模型并解析目标坐标

import base64
import copy
import ctypes
import json
import os
import sys
import time
import re
import threading
import io
import logging

logger = logging.getLogger(__name__)

# 依赖库
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    logger.error("未找到 requests 库 (pip install requests)")

# ======================================================================
# 全局配置
# ======================================================================
def _get_program_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _get_program_dir()
APP_CONFIG_FILE = os.path.join(APP_DIR, "macro_settings.mmcfg")
LEGACY_APP_CONFIG_FILE = os.path.join(APP_DIR, "macro_settings.json")
_DEFAULT_APP_CONFIG_FILE = APP_CONFIG_FILE
VLM_CONFIG_KEY = "vlm"
_PROTECTED_API_KEY_PREFIX = 'dpapi:'
_DPAPI_ENTROPY = b'MacroMate/VLM/API-Key/v1'

# 默认配置
DEFAULT_CONFIG = {
    "provider": "openai",  # openai, anthropic, deepseek, zhipu
    "api_key": "",
    "base_url": "https://api.openai.com/v1",
    "model": "gpt-4o",
    "timeout": 30,
    "system_prompt": "你是一个自动化助手。请分析用户指令和屏幕截图，返回目标位置的坐标。只返回 X, Y 坐标数字，用英文逗号分隔，例如: 123,456。如果找不到目标，返回 none"
}

_COORDINATE_PATTERNS = tuple(re.compile(pattern) for pattern in (
    r'(-?\d+)\s*[,，]\s*(-?\d+)\s*[,，]\s*(-?\d+)\s*[,，]\s*(-?\d+)',
    r'(-?\d+)\s*[,，]\s*(-?\d+)',
    r'x\s*[:=]\s*(-?\d+)\s*[,，]?\s*y\s*[:=]\s*(-?\d+)',
    r'(-?\d+)\s*[,，]\s*(-?\d+).*(?:坐标|location)',
    r'位于\s*[（(]?\s*(-?\d+)\s*[,，]\s*(-?\d+)\s*[）)]?',
    r'(-?\d+)px?\s*[,，]\s*(-?\d+)px?',
    r'^\s*(?:coordinate|position|location|point)?\s*[:：]?\s*(-?\d+)\s+(-?\d+)\s*$',
))

# 提供商配置
PROVIDER_CONFIGS = {
    "openai": {
        "name": "OpenAI (GPT-4o)",
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o",
        "supports_vision": True
    },
    "anthropic": {
        "name": "Anthropic (Claude)",
        "base_url": "https://api.anthropic.com/v1",
        "model": "claude-3-5-sonnet-20241022",
        "supports_vision": True
    },
    "deepseek": {
        "name": "DeepSeek",
        "base_url": "https://api.deepseek.com/v1",
        "model": "deepseek-chat",
        "supports_vision": False
    },
    "zhipu": {
        "name": "智谱清言 (GLM-4V)",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-4v-plus",
        "supports_vision": True
    },
    "qianwen": {
        "name": "阿里通义千问 (Qwen-VL)",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-plus",
        "supports_vision": True
    },
    "openrouter": {
        "name": "OpenRouter (聚合AI)",
        "base_url": "https://openrouter.ai/api/v1",
        "model": "google/gemma-3-4b-it:free",
        "supports_vision": True
    },
    "step": {
        "name": "阶跃星辰 (Step)",
        "base_url": "https://api.stepfun.com/v1",
        "model": "step-1v-8k",
        "supports_vision": True
    }
}

# ======================================================================
# 引擎状态
# ======================================================================
_vlm_config = None
_vlm_lock = threading.Lock()


def _read_json_file(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError, TypeError) as e:
        logger.error("加载配置失败 (%s): %s", os.path.basename(path), e)
        return {}


def _merge_user_config(default, user_config):
    merged = default.copy()
    if not isinstance(user_config, dict):
        return merged

    for k, v in user_config.items():
        if v is not None:
            merged[k] = v
    user_has_base_url = bool(user_config.get('base_url'))
    user_has_model = bool(user_config.get('model'))
    if merged['provider'] in PROVIDER_CONFIGS:
        pc = PROVIDER_CONFIGS[merged['provider']]
        if not user_has_base_url:
            merged['base_url'] = pc.get('base_url', DEFAULT_CONFIG['base_url'])
        if not user_has_model:
            merged['model'] = pc.get('model', DEFAULT_CONFIG['model'])
    return merged


class _DATA_BLOB(ctypes.Structure):
    _fields_ = (
        ('cbData', ctypes.c_ulong),
        ('pbData', ctypes.POINTER(ctypes.c_ubyte)),
    )


def _make_data_blob(data):
    buffer = ctypes.create_string_buffer(data)
    blob = _DATA_BLOB(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
    )
    return blob, buffer


def _protect_api_key(api_key):
    """Protect an API key for the current Windows user with DPAPI."""
    if not api_key:
        return ''
    if sys.platform != 'win32':
        logger.error("API Key 加密仅支持 Windows，已拒绝明文保存")
        return None
    try:
        raw_blob, raw_buffer = _make_data_blob(str(api_key).encode('utf-8'))
        entropy_blob, entropy_buffer = _make_data_blob(_DPAPI_ENTROPY)
        output_blob = _DATA_BLOB()
        crypt_protect = ctypes.windll.crypt32.CryptProtectData
        crypt_protect.argtypes = (
            ctypes.POINTER(_DATA_BLOB), ctypes.c_wchar_p,
            ctypes.POINTER(_DATA_BLOB), ctypes.c_void_p, ctypes.c_void_p,
            ctypes.c_ulong, ctypes.POINTER(_DATA_BLOB),
        )
        crypt_protect.restype = ctypes.c_bool
        if not crypt_protect(
                ctypes.byref(raw_blob), None, ctypes.byref(entropy_blob),
                None, None, 0x1, ctypes.byref(output_blob)):
            raise OSError("CryptProtectData failed")
        try:
            protected = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        finally:
            ctypes.windll.kernel32.LocalFree(output_blob.pbData)
        # Keep input buffers alive until the native call has completed.
        _ = raw_buffer, entropy_buffer
        return _PROTECTED_API_KEY_PREFIX + base64.b64encode(protected).decode('ascii')
    except Exception:
        logger.exception("API Key 加密失败，已拒绝明文保存")
        return None


def _unprotect_api_key(stored_value):
    """Decode a DPAPI value; return None when it cannot be safely recovered."""
    if not isinstance(stored_value, str) or not stored_value.startswith(_PROTECTED_API_KEY_PREFIX):
        return stored_value
    if sys.platform != 'win32':
        return None
    try:
        protected = base64.b64decode(
            stored_value[len(_PROTECTED_API_KEY_PREFIX):], validate=True,
        )
        protected_blob, protected_buffer = _make_data_blob(protected)
        entropy_blob, entropy_buffer = _make_data_blob(_DPAPI_ENTROPY)
        output_blob = _DATA_BLOB()
        crypt_unprotect = ctypes.windll.crypt32.CryptUnprotectData
        crypt_unprotect.argtypes = (
            ctypes.POINTER(_DATA_BLOB), ctypes.POINTER(ctypes.c_wchar_p),
            ctypes.POINTER(_DATA_BLOB), ctypes.c_void_p, ctypes.c_void_p,
            ctypes.c_ulong, ctypes.POINTER(_DATA_BLOB),
        )
        crypt_unprotect.restype = ctypes.c_bool
        if not crypt_unprotect(
                ctypes.byref(protected_blob), None, ctypes.byref(entropy_blob),
                None, None, 0x1, ctypes.byref(output_blob)):
            raise OSError("CryptUnprotectData failed")
        try:
            raw = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        finally:
            ctypes.windll.kernel32.LocalFree(output_blob.pbData)
        _ = protected_buffer, entropy_buffer
        return raw.decode('utf-8')
    except Exception:
        logger.exception("API Key 解密失败，请重新输入并保存")
        return None


def _prepare_vlm_config_for_storage(config, allow_existing_protected=False):
    stored = copy.deepcopy(config)
    stored.pop('_api_key_source', None)
    api_key = stored.get('api_key', '')
    if not api_key:
        stored['api_key'] = ''
        return stored
    if allow_existing_protected and isinstance(api_key, str) \
            and api_key.startswith(_PROTECTED_API_KEY_PREFIX):
        return stored
    protected = _protect_api_key(api_key)
    if protected is None:
        raise OSError("API Key protection failed")
    stored['api_key'] = protected
    return stored


def prepare_app_config_for_storage(app_config):
    """Return an app-config copy whose VLM secret is safe for disk storage."""
    stored = copy.deepcopy(app_config)
    vlm_config = stored.get(VLM_CONFIG_KEY)
    if isinstance(vlm_config, dict):
        stored[VLM_CONFIG_KEY] = _prepare_vlm_config_for_storage(
            vlm_config,
            allow_existing_protected=True,
        )
    return stored


def _load_user_config():
    app_config = _read_json_file(_resolve_app_config_path())
    vlm_config = app_config.get(VLM_CONFIG_KEY)
    if not isinstance(vlm_config, dict):
        return {}
    loaded = copy.deepcopy(vlm_config)
    stored_api_key = loaded.get('api_key', '')
    if isinstance(stored_api_key, str) and stored_api_key.startswith(_PROTECTED_API_KEY_PREFIX):
        loaded['api_key'] = _unprotect_api_key(stored_api_key) or ''
    return loaded


def _resolve_app_config_path():
    """Return the readable app config path without migrating patched test paths."""
    if os.path.normcase(os.path.abspath(APP_CONFIG_FILE)) != os.path.normcase(os.path.abspath(_DEFAULT_APP_CONFIG_FILE)):
        return APP_CONFIG_FILE
    from sys_utils import migrate_legacy_config_file
    return migrate_legacy_config_file(APP_CONFIG_FILE, LEGACY_APP_CONFIG_FILE)


def _get_env_api_key(provider):
    provider_key = re.sub(r'[^A-Za-z0-9]+', '_', str(provider or '')).upper().strip('_')
    env_names = []
    if provider_key:
        env_names.append(f"MACROMATE_{provider_key}_API_KEY")
    env_names.append("MACROMATE_VLM_API_KEY")
    for name in env_names:
        value = os.environ.get(name, '').strip()
        if value:
            return value
    return ''


def _apply_runtime_secret_config(cfg):
    runtime_cfg = cfg.copy()
    if not runtime_cfg.get('api_key'):
        env_key = _get_env_api_key(runtime_cfg.get('provider', 'openai'))
        if env_key:
            runtime_cfg['api_key'] = env_key
            runtime_cfg['_api_key_source'] = 'env'
    return runtime_cfg


# ======================================================================
# 配置管理
# ======================================================================
def load_config():
    """加载 VLM 配置并返回与全局缓存隔离的副本。"""
    global _vlm_config
    with _vlm_lock:
        if _vlm_config is None:
            default = DEFAULT_CONFIG.copy()
            provider = default.get('provider', 'openai')
            if provider in PROVIDER_CONFIGS:
                pc = PROVIDER_CONFIGS[provider]
                default['base_url'] = pc.get('base_url', DEFAULT_CONFIG['base_url'])
                default['model'] = pc.get('model', DEFAULT_CONFIG['model'])
            _vlm_config = _merge_user_config(default, _load_user_config())
        return copy.deepcopy(_vlm_config)


def save_config(config):
    """保存 VLM 配置（原子写入），不保留调用方的可变对象。"""
    global _vlm_config
    from sys_utils import get_shared_file_lock

    if not isinstance(config, dict):
        logger.error("保存配置失败: VLM config must be a dictionary")
        return False

    with _vlm_lock, get_shared_file_lock(APP_CONFIG_FILE):
        try:
            config_copy = copy.deepcopy(config)
            read_path = _resolve_app_config_path()
            app_config = _read_json_file(read_path)
            app_config[VLM_CONFIG_KEY] = _prepare_vlm_config_for_storage(config_copy)
            from sys_utils import write_json_file_atomically
            write_json_file_atomically(APP_CONFIG_FILE, app_config)
            _vlm_config = copy.deepcopy(config_copy)
            return True
        except Exception as e:
            logger.error(f"保存配置失败: {e}")
            return False


def get_providers():
    """获取支持的提供商列表副本。"""
    return copy.deepcopy(PROVIDER_CONFIGS)


# ======================================================================
# 截图与编码
# ======================================================================
def capture_screen(region=None):
    """
    截取屏幕并转为 Base64

    Args:
        region: 可选的区域坐标 (x1, y1, x2, y2)，None 表示全屏

    Returns:
        base64_str: Base64 编码的图片字符串
        offset: (x_offset, y_offset) 区域左上角坐标
    """
    screenshot = None
    try:
        from sys_utils import capture_physical_bbox

        screenshot, offset = capture_physical_bbox(region)

        return _encode_screenshot_pil(screenshot), offset
    except Exception as e:
        logger.error(f"截图失败: {e}")
        raise RuntimeError(f"VLM screen capture failed: {e}") from e
    finally:
        if screenshot:
            try:
                screenshot.close()
            except Exception:
                pass

# ======================================================================
# 坐标解析
# ======================================================================
def parse_coordinates(response_text):
    """
    解析 API 返回的坐标文本
    
    支持格式:
    - "123,456"
    - "X: 123, Y: 456"  
    - "x=123, y=456"
    - "坐标: 123, 456"
    - "位于 (123, 456)"
    
    Returns:
        (x, y) 或 None
    """
    if not response_text:
        return None
    
    # 清理文本
    text = response_text.strip().lower()
    
    # 检查无结果标记
    if re.search(r'\bnone\b', text) or '找不到' in text or '未找到' in text or '无法' in text:
        return None
    
    for pattern in _COORDINATE_PATTERNS:
        match = pattern.search(text)
        if match:
            try:
                groups = match.groups()
                if len(groups) == 4:
                    x1, y1, x2, y2 = int(groups[0]), int(groups[1]), int(groups[2]), int(groups[3])
                    if -10000 <= x1 <= 10000 and -10000 <= y1 <= 10000 and -10000 <= x2 <= 10000 and -10000 <= y2 <= 10000:
                        cx = (x1 + x2) // 2
                        cy = (y1 + y2) // 2
                        return (cx, cy)
                else:
                    # 2个坐标
                    x = int(groups[0])
                    y = int(groups[1])
                    # 合理性检查 (屏幕坐标通常在 -10000 到 10000 范围内，放宽以支持多屏和负坐标)
                    if -10000 <= x <= 10000 and -10000 <= y <= 10000:
                        return (x, y)
            except (ValueError, IndexError):
                continue
    
    return None


# ======================================================================
# API 调用
# ======================================================================
# API call adapters
# ======================================================================
class _OpenAICompatibleAdapter:
    endpoint = "/chat/completions"
    include_system = True
    max_tokens = 100
    image_order = "text_first"
    extra_headers = None

    @classmethod
    def build_request(cls, instruction, image_b64, cfg):
        headers = {
            "Authorization": f"Bearer {cfg.get('api_key', '')}",
            "Content-Type": "application/json",
        }
        if cls.extra_headers:
            headers.update(cls.extra_headers)

        messages = []
        system_prompt = cfg.get('system_prompt', DEFAULT_CONFIG['system_prompt'])
        if cls.include_system:
            messages.append({"role": "system", "content": system_prompt})

        if image_b64:
            text_part = {"type": "text", "text": instruction}
            image_part = {
                "type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"},
            }
            content = [text_part, image_part] if cls.image_order == "text_first" else [image_part, text_part]
            messages.append({"role": "user", "content": content})
        else:
            messages.append({"role": "user", "content": instruction})

        payload = {"model": cfg.get('model', ''), "messages": messages}
        if cls.max_tokens is not None:
            payload["max_tokens"] = cls.max_tokens
        return f"{cfg.get('base_url', '')}{cls.endpoint}", headers, payload

    @staticmethod
    def parse_response(result):
        if not isinstance(result, dict):
            return ""
        choices = result.get('choices', [])
        if not isinstance(choices, list) or not choices or not isinstance(choices[0], dict):
            return ""
        msg = choices[0].get('message', {})
        if not isinstance(msg, dict):
            return ""
        content = msg.get('content', '')
        return content if isinstance(content, str) else ""


class _AnthropicAdapter:
    @staticmethod
    def build_request(instruction, image_b64, cfg):
        headers = {
            "x-api-key": cfg.get('api_key', ''),
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }
        payload = {
            "model": cfg.get('model', ''),
            "max_tokens": 100,
            "system": cfg.get('system_prompt', DEFAULT_CONFIG['system_prompt']),
        }
        if image_b64:
            payload["messages"] = [{
                "role": "user",
                "content": [
                    {"type": "text", "text": instruction},
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": image_b64,
                        },
                    },
                ],
            }]
        else:
            payload["messages"] = [{"role": "user", "content": instruction}]
        return f"{cfg.get('base_url', '')}/messages", headers, payload

    @staticmethod
    def parse_response(result):
        if not isinstance(result, dict):
            return ""
        content = result.get('content', [])
        if not isinstance(content, list) or not content or not isinstance(content[0], dict):
            return ""
        text = content[0].get('text', '')
        return text if isinstance(text, str) else ""


class _DeepSeekAdapter(_OpenAICompatibleAdapter):
    image_order = "image_first"


class _ZhipuAdapter(_OpenAICompatibleAdapter):
    include_system = False
    max_tokens = None
    image_order = "image_first"


class _QianwenAdapter(_OpenAICompatibleAdapter):
    image_order = "image_first"


class _StepAdapter(_OpenAICompatibleAdapter):
    include_system = False
    image_order = "image_first"


class _OpenRouterAdapter(_OpenAICompatibleAdapter):
    extra_headers = {
        "HTTP-Referer": "https://github.com/hxlive/MacroMate",
        "X-Title": "MacroMate",
    }


VLM_ADAPTERS = {
    "openai": _OpenAICompatibleAdapter,
    "anthropic": _AnthropicAdapter,
    "deepseek": _DeepSeekAdapter,
    "zhipu": _ZhipuAdapter,
    "qianwen": _QianwenAdapter,
    "step": _StepAdapter,
    "openrouter": _OpenRouterAdapter,
}


def _encode_screenshot_pil(screenshot_pil):
    buffer = io.BytesIO()
    screenshot_pil.save(buffer, format='JPEG', quality=85)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def _resolve_vlm_image_b64(image_b64, screenshot_pil, raise_on_error):
    if image_b64 or not screenshot_pil:
        return image_b64
    try:
        return _encode_screenshot_pil(screenshot_pil)
    except Exception as e:
        err_msg = f"[VLM] FAIL image encoding failed: {e}"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg) from e
        return None


def _validate_vlm_capability(provider, model, image_b64, raise_on_error):
    if not image_b64:
        logger.info("text-only mode")
        return True
    if PROVIDER_CONFIGS.get(provider, {}).get('supports_vision') is False:
        msg = f"current model does not support image input: {provider}/{model}"
        logger.error(msg)
        if raise_on_error:
            raise ValueError(msg)
        return False
    return True


def _parse_vlm_response_json(response, raise_on_error):
    try:
        return response.json()
    except ValueError as e:
        err_msg = f"[VLM] FAIL API returned invalid JSON: {e}"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg) from e
        return None


def call_vlm_api(instruction, image_b64=None, screenshot_pil=None, config=None, raise_on_error=False):
    """Call the configured VLM API and return an (x, y) coordinate or None."""
    if not REQUESTS_AVAILABLE:
        err_msg = "[VLM] FAIL requests package is unavailable"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg)
        return None

    try:
        source_config = config if config is not None else load_config()
        if not isinstance(source_config, dict):
            raise TypeError("VLM config must be a dictionary")
        cfg = _apply_runtime_secret_config(source_config)
        provider = cfg.get('provider', 'openai')
        adapter = VLM_ADAPTERS.get(provider)
    except (AttributeError, KeyError, TypeError, IndexError, ValueError) as e:
        err_msg = f"[VLM] FAIL invalid configuration: {e}"
        logger.error(err_msg)
        if raise_on_error:
            raise RuntimeError(err_msg) from e
        return None

    if not cfg.get('api_key'):
        err_msg = "[VLM] FAIL API key is not configured"
        logger.error(err_msg)
        if raise_on_error: raise ValueError(err_msg)
        return None

    model = cfg.get('model', '')
    timeout = cfg.get('timeout', 30)
    if not adapter:
        err_msg = f"unsupported provider: {provider}"
        logger.error(err_msg)
        if raise_on_error:
            raise ValueError(err_msg)
        return None

    image_b64 = _resolve_vlm_image_b64(image_b64, screenshot_pil, raise_on_error)
    if screenshot_pil and not image_b64:
        return None
    if not _validate_vlm_capability(provider, model, image_b64, raise_on_error):
        return None

    try:
        url, headers, payload = adapter.build_request(instruction, image_b64, cfg)
    except (AttributeError, KeyError, TypeError, IndexError, ValueError) as e:
        err_msg = f"[VLM] FAIL invalid request configuration: {e}"
        logger.error(err_msg)
        if raise_on_error:
            raise RuntimeError(err_msg) from e
        return None

    try:
        t0 = time.time()
        response = requests.post(url, headers=headers, json=payload, timeout=timeout)
        elapsed = time.time() - t0

        if response.status_code != 200:
            reason = getattr(response, 'reason', '') or 'request failed'
            err_msg = f"[VLM] API returned error: {response.status_code} - {reason}"
            logger.error(err_msg)
            if raise_on_error: raise RuntimeError(err_msg)
            return None

        result = _parse_vlm_response_json(response, raise_on_error)
        if result is None:
            return None

        text_content = adapter.parse_response(result)
        if not text_content:
            err_msg = "[VLM] FAIL API returned empty or malformed content"
            logger.warning(err_msg)
            if raise_on_error:
                raise RuntimeError(err_msg)
            return None

        logger.info("API response received (%.2fs, %s chars)", elapsed, len(text_content))
        return parse_coordinates(text_content)

    except requests.Timeout as e:
        err_msg = f"[VLM] FAIL request timed out ({timeout}s)"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg) from e
        return None
    except requests.RequestException as e:
        err_msg = f"[VLM] FAIL request failed: {e}"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg) from e
        return None
    except (AttributeError, KeyError, TypeError, IndexError, ValueError) as e:
        err_msg = f"[VLM] FAIL malformed API response: {e}"
        logger.error(err_msg)
        if raise_on_error: raise RuntimeError(err_msg) from e
        return None


def find_location_by_vlm(instruction, region=None, config=None):
    """
    同步兼容定位入口。

    新的宏执行流程应通过 screen_locator.locate() 调用，以获得统一的取消检测、
    截图复用和定位策略。本函数保留给独立同步调用及旧接口兼容使用。
    
    Args:
        instruction: 自然语言指令
        region: 可选的搜索区域 (x1, y1, x2, y2)
        config: 可选的配置
        
    Returns:
        (x, y) 坐标或 None (如果指定了 region，会自动转换为绝对坐标)
    """
    # 截图 (带偏移量)
    b64, offset = capture_screen(region)
    if not b64:
        return None
    
    # 调用 API
    coords = call_vlm_api(instruction, image_b64=b64, config=config)
    
    # 截图坐标以截图左上角为原点；区域截图和虚拟屏全屏截图都可能带 offset。
    if coords and offset != (0, 0):
        abs_x = coords[0] + offset[0]
        abs_y = coords[1] + offset[1]
        return (abs_x, abs_y)
    
    return coords
