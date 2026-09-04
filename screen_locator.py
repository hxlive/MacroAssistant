# -*- coding: utf-8 -*-
# screen_locator.py
# 功能说明：统一屏幕定位服务，负责截图、图像匹配、OCR/VLM 调用适配与定位状态管理
# Version: 1.8.6

"""屏幕目标定位的统一契约与后端实现。

本模块统一管理屏幕截图、图像匹配、OCR/VLM 适配及相关资源生命周期，
并与宏执行引擎、GUI 和控制器保持独立。调用方通过轻量的取消检查函数
传递停止信号，无需传入应用对象或执行上下文。
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, replace
import logging
import math
from numbers import Real
import os
import threading
import time
from typing import Callable, Literal, TypeAlias

from sys_utils import capture_physical_bbox

import ocr_engine
import vlm_engine


logger = logging.getLogger(__name__)

try:
    import cv2
    import numpy as np
    OPENCV_AVAILABLE = True
except ImportError:
    cv2 = None
    np = None
    OPENCV_AVAILABLE = False

SCALES = (1.0, 0.9, 1.1, 0.8, 1.2)
QUICK_CHECK_SCALES = (1.0, 0.9, 1.1)
TEMPLATE_CACHE_SIZE = 200
TEMPLATE_CACHE_MAX_BYTES = 128 * 1024 * 1024

_TEMPLATE_CACHE = OrderedDict()
_TEMPLATE_CACHE_BYTES = 0
_TEMPLATE_CACHE_LOCK = threading.RLock()
_VLM_REQUEST_LOCK = threading.Lock()


Point: TypeAlias = tuple[int, int]
BBox: TypeAlias = tuple[int, int, int, int]
CancelCheck: TypeAlias = Callable[[], None]


class ScreenLocatorError(RuntimeError):
    """Base error raised by screen locating infrastructure."""


class ScreenCaptureError(ScreenLocatorError):
    """Screen capture could not produce an image."""


class LocatorBackendError(ScreenLocatorError):
    """A configured locating backend failed unexpectedly."""


@dataclass(frozen=True)
class LocateRequest:
    mode: Literal['image', 'text', 'ai']
    region_bbox: BBox | None = None
    template_path: str | None = None
    target_text: str | None = None
    instruction: str | None = None
    confidence: float = 0.8
    lang: str = 'eng'
    ocr_engine: str = 'auto'
    ocr_debug: bool = False
    enhanced_mode: bool = False
    cache_key: str | None = None
    preferred_position: Point | None = None


@dataclass(frozen=True)
class LocateResult:
    position: Point | None
    source: Literal['image', 'ocr', 'vlm']
    recognized_text: str = ''
    score: float | None = None
    elapsed_seconds: float = 0.0
    fast_path_hit: bool = False

    @property
    def found(self) -> bool:
        return self.position is not None


# Per-run locating state. These classes live with the locating subsystem so the
# macro engine does not own image/OCR cache or locating performance details.
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

class LoopCacheManager:
    def __init__(self):
        self._lock = threading.RLock()
        self.reset()

    def reset(self):
        with self._lock:
            self.caches = {}
            self.stack = []

    def get_current_loop_id(self):
        with self._lock:
            return self.stack[-1] if self.stack else None

    def enter(self, loop_id):
        with self._lock:
            if loop_id not in self.caches:
                self.caches[loop_id] = {}
            self.stack.append(loop_id)

    def exit(self):
        with self._lock:
            if self.stack:
                loop_id = self.stack.pop()
                # Clear this loop cache here; execute_steps finally also resets as a fallback.
                if loop_id in self.caches:
                    del self.caches[loop_id]


    def get(self, sig):
        with self._lock:
            loop_id = self.stack[-1] if self.stack else None
            return self.caches.get(loop_id, {}).get(sig) if loop_id else None

    def set(self, sig, loc):
        with self._lock:
            loop_id = self.stack[-1] if self.stack else None
            if loop_id:
                if loop_id not in self.caches:
                     self.caches[loop_id] = {}
                self.caches[loop_id][sig] = loc


class LocatorSession:
    """Own all mutable locating state for one macro execution."""

    def __init__(self):
        self.performance = PerformanceMonitor()
        self._loop_cache = LoopCacheManager()
        self._runtime_boxes = {}
        self._lock = threading.RLock()

    def begin_run(self):
        """Start with clean statistics and caches, even when a context is reused."""
        self.performance.reset()
        self._loop_cache.reset()
        with self._lock:
            self._runtime_boxes.clear()

    def enter_loop(self, loop_id):
        self._loop_cache.enter(loop_id)

    def exit_loop(self):
        self._loop_cache.exit()

    def clear_loop_state(self):
        self._loop_cache.reset()

    def is_in_loop(self):
        return self._loop_cache.get_current_loop_id() is not None

    def get_preferred_position(self, signature):
        return self._loop_cache.get(signature)

    def remember_position(self, signature, position):
        self._loop_cache.set(signature, position)

    def get_runtime_box(self, signature, default=None):
        with self._lock:
            return self._runtime_boxes.get(signature, default)

    def set_runtime_box(self, signature, bbox):
        with self._lock:
            self._runtime_boxes[signature] = list(bbox)

    def record_result(self, result):
        """Record a final image/OCR result without exposing counter details to core."""
        is_ocr = result.source == 'ocr'
        if not result.found:
            self.performance.record_miss(is_ocr)
            return
        self.performance.record_hit(bool(result.fast_path_hit), is_ocr)
        if not result.fast_path_hit and result.elapsed_seconds:
            self.performance.record_time(result.elapsed_seconds, is_ocr)

    def get_stats(self):
        return self.performance.get_stats()

@dataclass(frozen=True)
class LocatePolicy:
    retry_count: int = 0
    retry_interval: float = 0.0
    fallback_request: LocateRequest | None = None


@dataclass(frozen=True)
class LocateOutcome:
    result: LocateResult
    attempt_count: int
    fallback_attempted: bool = False
    fallback_used: bool = False


def _validate_policy(policy: LocatePolicy) -> None:
    if not isinstance(policy, LocatePolicy):
        raise TypeError('policy must be a LocatePolicy')
    if isinstance(policy.retry_count, bool) or not isinstance(policy.retry_count, int):
        raise ValueError('retry_count must be a non-negative integer')
    if policy.retry_count < 0:
        raise ValueError('retry_count must be a non-negative integer')
    try:
        interval = float(policy.retry_interval)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('retry_interval must be a finite non-negative number') from exc
    if not math.isfinite(interval) or interval < 0:
        raise ValueError('retry_interval must be a finite non-negative number')


def _wait_with_cancel(seconds: float, cancel_check: CancelCheck) -> None:
    deadline = time.monotonic() + seconds
    while True:
        cancel_check()
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(0.1, remaining))


def locate_with_policy(
        request: LocateRequest,
        policy: LocatePolicy | None = None,
        cancel_check: CancelCheck = lambda: None,
) -> LocateOutcome:
    """Locate with cancellable retries and an optional full-screen fallback."""
    policy = policy or LocatePolicy()
    _validate_policy(policy)
    fallback_attempted = False
    last_result = None

    for attempt in range(policy.retry_count + 1):
        cancel_check()
        last_result = locate(request, cancel_check=cancel_check)
        if last_result.found:
            return LocateOutcome(
                last_result, attempt + 1,
                fallback_attempted=fallback_attempted,
            )

        if policy.fallback_request is not None:
            fallback_attempted = True
            last_result = locate(policy.fallback_request, cancel_check=cancel_check)
            if last_result.found:
                return LocateOutcome(
                    last_result, attempt + 1,
                    fallback_attempted=True, fallback_used=True,
                )

        if attempt < policy.retry_count and policy.retry_interval > 0:
            _wait_with_cancel(float(policy.retry_interval), cancel_check)

    return LocateOutcome(
        last_result, policy.retry_count + 1,
        fallback_attempted=fallback_attempted,
    )


def locate_condition(
        mode: str,
        *,
        region_bbox: BBox | None = None,
        template_path: str = '',
        target_text: str = '',
        confidence: float = 0.8,
        lang: str = 'eng',
        enhanced_mode: bool = False,
        cache_key: str | None = None,
        cancel_check: CancelCheck = lambda: None,
) -> LocateResult:
    """Build and execute an image/text condition request."""
    if mode == 'until_image':
        if not template_path or not os.path.exists(template_path):
            return LocateResult(position=None, source='image')
        try:
            confidence = float(confidence)
        except (TypeError, ValueError, OverflowError):
            return LocateResult(position=None, source='image')
        if not math.isfinite(confidence) or not 0.0 <= confidence <= 1.0:
            return LocateResult(position=None, source='image')
        request = LocateRequest(
            mode='image', region_bbox=region_bbox, template_path=template_path,
            confidence=confidence, enhanced_mode=enhanced_mode,
        )
    elif mode == 'until_text':
        if not target_text:
            return LocateResult(position=None, source='ocr')
        request = LocateRequest(
            mode='text', region_bbox=region_bbox, target_text=target_text,
            lang=lang, ocr_engine='auto', enhanced_mode=enhanced_mode,
            cache_key=cache_key,
        )
    else:
        raise ValueError(f'unsupported condition mode: {mode!r}')
    return locate(request, cancel_check=cancel_check)


def coerce_bbox(raw_bbox) -> BBox | None:
    if not isinstance(raw_bbox, (list, tuple)):
        return None
    if len(raw_bbox) not in (2, 4):
        return None
    try:
        bbox = tuple(int(value) for value in raw_bbox)
    except (TypeError, ValueError, OverflowError):
        return None
    if len(bbox) == 2:
        return (bbox[0], bbox[1], bbox[0] + 1, bbox[1] + 1)
    if bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
        return None
    return bbox


def bbox_to_region(raw_bbox):
    """Convert an absolute BBox into a ``(left, top, width, height)`` region."""
    bbox = coerce_bbox(raw_bbox)
    if bbox is None:
        return None
    return (bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1])


def expand_bbox(raw_bbox, padding: int) -> BBox | None:
    """Expand an absolute BBox without clipping virtual-screen coordinates."""
    bbox = coerce_bbox(raw_bbox)
    if bbox is None:
        return None
    try:
        padding = int(padding)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('padding must be a non-negative integer') from exc
    if padding < 0:
        raise ValueError('padding must be a non-negative integer')
    return (
        bbox[0] - padding,
        bbox[1] - padding,
        bbox[2] + padding,
        bbox[3] + padding,
    )


def smart_screenshot(region=None, pad: int = 0):
    """Capture a full screen or a ``(left, top, width, height)`` region.

    The returned image belongs to the caller and must be closed by the caller.
    """
    try:
        pad = int(pad)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('pad must be a non-negative integer') from exc
    if pad < 0:
        raise ValueError('pad must be a non-negative integer')

    capture_bbox = None
    if region is not None:
        if not isinstance(region, (list, tuple)) or len(region) != 4:
            raise ValueError(f'invalid screenshot region: {region!r}')
        try:
            left, top, width, height = (int(value) for value in region)
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValueError(f'invalid screenshot region: {region!r}') from exc
        if width <= 0 or height <= 0:
            raise ValueError(f'invalid screenshot region: {region!r}')
        capture_bbox = (
            left - pad,
            top - pad,
            left + width + pad,
            top + height + pad,
        )

    try:
        return capture_physical_bbox(capture_bbox)
    except ScreenCaptureError:
        raise
    except Exception as exc:
        raise ScreenCaptureError('Screen capture is unavailable') from exc


def sample_screen_pixel(x: int, y: int) -> tuple[int, int, int]:
    """Return one physical-screen pixel as an RGB tuple.

    Coordinates use the virtual desktop coordinate space, so negative values on
    monitors placed left of or above the primary display are supported.
    """
    try:
        x = int(x)
        y = int(y)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('pixel coordinates must be integers') from exc

    screenshot = None
    rgb_image = None
    try:
        screenshot, _offset = capture_physical_bbox((x, y, x + 1, y + 1))
        rgb_image = screenshot.convert('RGB')
        pixel = rgb_image.getpixel((0, 0))
        return int(pixel[0]), int(pixel[1]), int(pixel[2])
    except ScreenCaptureError:
        raise
    except Exception as exc:
        raise ScreenCaptureError(f'Unable to sample screen pixel at ({x}, {y})') from exc
    finally:
        if rgb_image is not None and rgb_image is not screenshot:
            try:
                rgb_image.close()
            except Exception:
                pass
        if screenshot is not None:
            try:
                screenshot.close()
            except Exception:
                pass


def _clear_template_cache():
    global _TEMPLATE_CACHE_BYTES
    with _TEMPLATE_CACHE_LOCK:
        _TEMPLATE_CACHE.clear()
        _TEMPLATE_CACHE_BYTES = 0


def _get_template_cached(path, scale, modified_ns, file_size):
    global _TEMPLATE_CACHE_BYTES
    key = (path, scale, modified_ns, file_size)
    with _TEMPLATE_CACHE_LOCK:
        cached = _TEMPLATE_CACHE.pop(key, None)
        if cached is not None:
            _TEMPLATE_CACHE[key] = cached
            return cached[:3]
        stale_keys = [cached_key for cached_key in _TEMPLATE_CACHE
                      if cached_key[:2] == key[:2]]
        for stale_key in stale_keys:
            _TEMPLATE_CACHE_BYTES -= _TEMPLATE_CACHE.pop(stale_key)[3]

    image = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if image is None:
        return None, 0, 0
    if scale != 1.0:
        height, width = image.shape[:2]
        image = cv2.resize(
            image,
            (int(width * scale), int(height * scale)),
            interpolation=cv2.INTER_AREA,
        )
    result = (image, image.shape[1], image.shape[0])
    item_bytes = int(getattr(image, 'nbytes', 0))
    if item_bytes <= 0 or item_bytes > TEMPLATE_CACHE_MAX_BYTES:
        return result

    with _TEMPLATE_CACHE_LOCK:
        existing = _TEMPLATE_CACHE.pop(key, None)
        if existing is not None:
            _TEMPLATE_CACHE_BYTES -= existing[3]
        _TEMPLATE_CACHE[key] = (*result, item_bytes)
        _TEMPLATE_CACHE_BYTES += item_bytes
        while (len(_TEMPLATE_CACHE) > TEMPLATE_CACHE_SIZE or
               _TEMPLATE_CACHE_BYTES > TEMPLATE_CACHE_MAX_BYTES):
            _, removed = _TEMPLATE_CACHE.popitem(last=False)
            _TEMPLATE_CACHE_BYTES -= removed[3]
    return result


_get_template_cached.cache_clear = _clear_template_cache


def _get_template_cache_info():
    with _TEMPLATE_CACHE_LOCK:
        return len(_TEMPLATE_CACHE), _TEMPLATE_CACHE_BYTES


def _get_template(path, scale):
    try:
        stat = os.stat(path)
        signature = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        signature = (None, None)
    return _get_template_cached(path, scale, *signature)


def find_image_cv2(
        path,
        confidence,
        screenshot_pil,
        offset=(0, 0),
        enhanced_mode=False,
        cancel_check: CancelCheck = lambda: None,
):
    if not OPENCV_AVAILABLE:
        return None
    try:
        screen_gray = cv2.cvtColor(np.array(screenshot_pil), cv2.COLOR_RGB2GRAY)
        best = (-1, None, 0, 0)
        scales_to_try = SCALES if enhanced_mode else (1.0,)
        for scale in scales_to_try:
            cancel_check()
            template, width, height = _get_template(path, scale)
            if (template is None or height > screen_gray.shape[0] or
                    width > screen_gray.shape[1]):
                continue
            matched = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
            _, max_value, _, max_location = cv2.minMaxLoc(matched)
            if max_value > best[0]:
                best = (max_value, max_location, width, height)
            if best[0] >= 0.95 and best[0] >= confidence:
                break
        value, location, width, height = best
        if value >= confidence and location is not None:
            center_x = offset[0] + location[0] + width // 2
            center_y = offset[1] + location[1] + height // 2
            return (center_x, center_y, width, height), value
    except (cv2.error, ValueError, TypeError, AttributeError) as exc:
        logger.error('CV2 image match error: %s', exc)
    return None


def quick_check_cv2(
        path,
        confidence,
        screenshot_pil,
        offset,
        target_location,
        enhanced_mode=False,
        cancel_check: CancelCheck = lambda: None,
):
    if not OPENCV_AVAILABLE:
        return False
    try:
        scales_to_try = QUICK_CHECK_SCALES if enhanced_mode else (1.0,)
        for scale in scales_to_try:
            cancel_check()
            template, width, height = _get_template(path, scale)
            if template is None:
                continue

            pad_width, pad_height = width // 2 + 15, height // 2 + 15
            relative_x = target_location[0] - offset[0]
            relative_y = target_location[1] - offset[1]
            left = max(0, relative_x - pad_width)
            top = max(0, relative_y - pad_height)
            right = min(screenshot_pil.width, relative_x + pad_width)
            bottom = min(screenshot_pil.height, relative_y + pad_height)
            if right <= left or bottom <= top:
                continue

            crop = cv2.cvtColor(
                np.array(screenshot_pil.crop((left, top, right, bottom))),
                cv2.COLOR_RGB2GRAY,
            )
            if crop.shape[0] < height or crop.shape[1] < width:
                continue
            matched = cv2.matchTemplate(crop, template, cv2.TM_CCOEFF_NORMED)
            _, max_value, _, _ = cv2.minMaxLoc(matched)
            if max_value >= confidence:
                return True
        return False
    except (cv2.error, ValueError, TypeError, AttributeError, IndexError) as exc:
        logger.error('CV2 quick match error: %s', exc)
        return False


def _validate_request(request: LocateRequest) -> None:
    if not isinstance(request, LocateRequest):
        raise TypeError('request must be a LocateRequest')
    if request.mode not in ('image', 'text', 'ai'):
        raise ValueError(f'unsupported locate mode: {request.mode!r}')

    try:
        confidence = float(request.confidence)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('confidence must be a finite number between 0 and 1') from exc
    if not math.isfinite(confidence) or not 0.0 <= confidence <= 1.0:
        raise ValueError('confidence must be a finite number between 0 and 1')

    if request.region_bbox is not None and coerce_bbox(request.region_bbox) is None:
        raise ValueError(f'invalid region_bbox: {request.region_bbox!r}')

    if request.preferred_position is not None:
        if request.mode != 'image':
            raise ValueError('preferred_position is only valid for image mode')
        if (not isinstance(request.preferred_position, (list, tuple)) or
                len(request.preferred_position) != 2):
            raise ValueError('preferred_position must be a two-element point')
        if _coerce_point(request.preferred_position) is None:
            raise ValueError('preferred_position must contain integer coordinates')

    if request.mode == 'image':
        if not isinstance(request.template_path, str) or not request.template_path.strip():
            raise ValueError('image mode requires a non-empty template_path')
        if request.target_text is not None or request.instruction is not None:
            raise ValueError('image mode cannot include text or AI fields')
    elif request.mode == 'text':
        if not isinstance(request.target_text, str) or not request.target_text:
            raise ValueError('text mode requires non-empty target_text')
        if request.template_path is not None or request.instruction is not None:
            raise ValueError('text mode cannot include image or AI fields')
    else:
        if not isinstance(request.instruction, str) or not request.instruction.strip():
            raise ValueError('ai mode requires a non-empty instruction')
        if request.template_path is not None or request.target_text is not None:
            raise ValueError('ai mode cannot include image or text fields')



def _coerce_point(raw_point) -> Point | None:
    if not isinstance(raw_point, (list, tuple)) or len(raw_point) != 2:
        return None

    normalized = []
    for value in raw_point:
        if isinstance(value, bool) or not isinstance(value, Real):
            return None
        numeric_value = float(value)
        if not math.isfinite(numeric_value) or not numeric_value.is_integer():
            return None
        normalized.append(int(numeric_value))

    return normalized[0], normalized[1]


def _locate_text(
        request: LocateRequest,
        screenshot_pil,
        offset: Point,
        cancel_check: CancelCheck,
) -> LocateResult:
    """Run OCR against a caller-owned screenshot and normalize its result."""
    cancel_check()
    try:
        raw_result = ocr_engine.find_text_location(
            request.target_text,
            lang=request.lang,
            debug=request.ocr_debug,
            screenshot_pil=screenshot_pil,
            offset=offset,
            engine=request.ocr_engine,
            enhanced_mode=request.enhanced_mode,
            cache_key=request.cache_key,
        )
    except Exception as exc:
        raise LocatorBackendError('OCR screen location failed') from exc
    cancel_check()

    if raw_result is None:
        result = LocateResult(position=None, source='ocr')
    else:
        point = _coerce_point(raw_result)
        recognized_text = ''
        if point is None and isinstance(raw_result, (list, tuple)) and len(raw_result) == 2:
            point = _coerce_point(raw_result[0])
            if point is not None and isinstance(raw_result[1], str):
                recognized_text = raw_result[1]
            else:
                point = None
        if point is None:
            raise LocatorBackendError(
                f'OCR backend returned malformed location: {raw_result!r}'
            )
        result = LocateResult(
            position=point,
            source='ocr',
            recognized_text=recognized_text,
        )

    cancel_check()
    return result



def _locate_ai(
        request: LocateRequest,
        cancel_check: CancelCheck,
) -> LocateResult:
    """Capture once and wait cancellably for a single VLM API request."""
    cancel_check()
    if not _VLM_REQUEST_LOCK.acquire(blocking=False):
        raise LocatorBackendError('another VLM request is still running')
    try:
        try:
            image_b64, offset = vlm_engine.capture_screen(request.region_bbox)
        except Exception as exc:
            raise ScreenCaptureError('VLM screen capture failed') from exc
        cancel_check()
    except BaseException:
        _VLM_REQUEST_LOCK.release()
        raise

    instruction = request.instruction
    completed = threading.Event()
    outcome = {}

    def call_backend():
        try:
            outcome['position'] = vlm_engine.call_vlm_api(
                instruction,
                image_b64=image_b64,
                raise_on_error=True,
            )
        except BaseException as exc:
            outcome['error'] = exc
        finally:
            _VLM_REQUEST_LOCK.release()
            completed.set()

    try:
        worker = threading.Thread(
            target=call_backend,
            name='MacroMateVLM',
            daemon=True,
        )
        worker.start()
    except BaseException:
        _VLM_REQUEST_LOCK.release()
        raise

    while not completed.wait(0.1):
        cancel_check()
    cancel_check()

    error = outcome.get('error')
    if error is not None:
        if isinstance(error, Exception):
            raise LocatorBackendError('VLM screen location failed') from error
        raise error

    raw_position = outcome.get('position')
    if raw_position is None:
        return LocateResult(position=None, source='vlm')

    position = _coerce_point(raw_position)
    offset_point = _coerce_point(offset)
    if position is None or offset_point is None:
        raise LocatorBackendError(
            f'VLM backend returned malformed location: {raw_position!r}'
        )
    return LocateResult(
        position=(position[0] + offset_point[0], position[1] + offset_point[1]),
        source='vlm',
    )



def _locate_image(
        request: LocateRequest,
        screenshot_pil,
        offset: Point,
        cancel_check: CancelCheck,
) -> LocateResult:
    """Run the preferred-position fast path before a full template match."""
    if request.preferred_position is not None:
        cancel_check()
        try:
            fast_path_hit = quick_check_cv2(
                request.template_path,
                request.confidence,
                screenshot_pil,
                offset,
                request.preferred_position,
                enhanced_mode=request.enhanced_mode,
                cancel_check=cancel_check,
            )
        except Exception as exc:
            raise LocatorBackendError('image quick check failed') from exc
        if fast_path_hit:
            return LocateResult(
                position=_coerce_point(request.preferred_position),
                source='image',
                fast_path_hit=True,
            )

    cancel_check()
    try:
        raw_result = find_image_cv2(
            request.template_path,
            request.confidence,
            screenshot_pil,
            offset=offset,
            enhanced_mode=request.enhanced_mode,
            cancel_check=cancel_check,
        )
    except Exception as exc:
        raise LocatorBackendError('image screen location failed') from exc
    cancel_check()
    if raw_result is None:
        return LocateResult(position=None, source='image')

    if not isinstance(raw_result, (list, tuple)) or len(raw_result) != 2:
        raise LocatorBackendError(
            f'image backend returned malformed result: {raw_result!r}'
        )
    raw_bbox, raw_score = raw_result
    if not isinstance(raw_bbox, (list, tuple)) or len(raw_bbox) < 2:
        raise LocatorBackendError(
            f'image backend returned malformed result: {raw_result!r}'
        )
    position = _coerce_point(raw_bbox[:2])
    try:
        score = float(raw_score)
    except (TypeError, ValueError, OverflowError) as exc:
        raise LocatorBackendError(
            f'image backend returned malformed score: {raw_score!r}'
        ) from exc
    if position is None or not math.isfinite(score):
        raise LocatorBackendError(
            f'image backend returned malformed result: {raw_result!r}'
        )
    return LocateResult(position=position, source='image', score=score)


def _with_elapsed(result: LocateResult, started_at: float) -> LocateResult:
    return replace(result, elapsed_seconds=max(0.0, time.perf_counter() - started_at))


def locate(
        request: LocateRequest,
        cancel_check: CancelCheck = lambda: None,
) -> LocateResult:
    """Validate, dispatch, normalize timing, and own capture lifecycle."""
    started_at = time.perf_counter()
    _validate_request(request)
    request = replace(request, confidence=float(request.confidence))

    if request.mode == 'ai':
        return _with_elapsed(_locate_ai(request, cancel_check), started_at)

    cancel_check()
    screenshot_pil = None
    try:
        region = bbox_to_region(request.region_bbox)
        screenshot_pil, offset = smart_screenshot(region)
        cancel_check()
        if request.mode == 'image':
            result = _locate_image(request, screenshot_pil, offset, cancel_check)
        else:
            result = _locate_text(request, screenshot_pil, offset, cancel_check)
        return _with_elapsed(result, started_at)
    finally:
        if screenshot_pil is not None:
            try:
                screenshot_pil.close()
            except Exception:
                logger.debug('Failed to close locator screenshot', exc_info=True)
