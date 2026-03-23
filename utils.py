import os
import cv2
import numpy as np
import json
import sys
import difflib
import pytesseract
import re
import win32api
import tempfile
import csv
import io
import subprocess
import math

# Tesseract Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TESSERACT_CMD = os.path.join(BASE_DIR, 'third_party', 'tesseract', 'tesseract.exe')
TESSDATA_DIR = os.path.join(BASE_DIR, 'third_party', 'tesseract', 'tessdata')

if os.path.exists(TESSERACT_CMD):
    try:
        # Use short path to avoid encoding issues with Chinese characters in path
        short_cmd = win32api.GetShortPathName(TESSERACT_CMD)
        pytesseract.pytesseract.tesseract_cmd = short_cmd
    except Exception as e:
        print(f"Warning: Could not get short path for tesseract: {e}")
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    
    # Set TESSDATA_PREFIX env var
    tess_prefix = os.path.join(BASE_DIR, 'third_party', 'tesseract')
    try:
        short_prefix = win32api.GetShortPathName(tess_prefix)
        os.environ['TESSDATA_PREFIX'] = short_prefix
    except:
        os.environ['TESSDATA_PREFIX'] = tess_prefix
else:
    # Fallback or warning
    pass

def log(message):
    """Helper to print with flush for real-time GUI updates"""
    print(message, flush=True)

def load_config(config_path):
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        log(f"Error loading config: {e}")
        return {}

def get_media_size(file_path):
    """
    Get original width and height of media file (image or video).
    Returns: (width, height) or None
    """
    if not os.path.exists(file_path):
        return None
    
    # Try as video
    if file_path.lower().endswith(('.mp4', '.avi', '.mov', '.wmv')):
        cap = cv2.VideoCapture(file_path)
        if cap.isOpened():
            w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            cap.release()
            if w > 0 and h > 0:
                return (w, h)
    
    # Try as image
    try:
        # Use imdecode to handle paths with non-ascii characters
        img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), -1)
        if img is not None:
            h, w = img.shape[:2]
            return (w, h)
    except Exception:
        pass

    return None

def calculate_centered_rect(container, content_size, bias_top=False):
    """
    Calculate new rect to fit content inside container while maintaining aspect ratio and centering.
    container: (left, top, width, height)
    content_size: (content_width, content_height)
    bias_top: if True, align vertically with 0.3 bias (for append report)
    Returns: (left, top, width, height)
    """
    c_left, c_top, c_width, c_height = container
    m_width, m_height = content_size
    
    if m_width == 0 or m_height == 0:
        return container

    # Calculate scale to fit
    scale_w = c_width / m_width
    scale_h = c_height / m_height
    scale = min(scale_w, scale_h)
    
    new_width = m_width * scale
    new_height = m_height * scale
    
    # Center horizontally
    new_left = c_left + (c_width - new_width) / 2
    
    if bias_top:
        # Vertically align with bias (0.3 = top 30%, bottom 70%)
        new_top = c_top + (c_height - new_height) * 0.3
    else:
        # Center vertically
        new_top = c_top + (c_height - new_height) / 2
    
    return (new_left, new_top, new_width, new_height)

def extract_last_frame(video_path, output_image_path):
    """Extracts the last frame of a video."""
    if not os.path.exists(video_path):
        log(f"Warning: Video not found at {video_path}")
        return False
    
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        log(f"Warning: Could not open video {video_path}")
        return False
    
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if frame_count == 0:
        log(f"Warning: Video {video_path} has 0 frames")
        cap.release()
        return False

    # Set to last frame
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_count - 1)
    ret, frame = cap.read()
    
    if ret:
        is_success, im_buf_arr = cv2.imencode(".png", frame)
        if is_success:
            im_buf_arr.tofile(output_image_path)
            log(f"Extracted last frame to {output_image_path}")
            cap.release()
            return True
        else:
             log(f"Warning: Could not encode frame for {video_path}")
    else:
        log(f"Warning: Could not read last frame of {video_path}")
    
    cap.release()
    return False

def split_rect(left, top, width, height, gap):
    try:
        gap = float(gap)
    except Exception:
        gap = 0.0

    if gap < 0:
        gap = 0.0

    if width <= gap:
        gap = 0.0

    half_width = (width - gap) / 2.0
    a = (left, top, half_width, height)
    b = (left + half_width + gap, top, half_width, height)
    return a, b

import difflib

def find_media_file(directory, key):
    """
    Finds a media file in the directory that matches the key.
    Prioritizes exact matches of name (ignoring extension).
    Fallbacks to fuzzy matching if exact match not found.
    """
    if not os.path.exists(directory):
        return None
        
    # Common media extensions
    extensions = ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.png', '.jpg', '.jpeg', '.bmp', '.gif']
    
    # 1. Exact Match Strategy
    # search for key.ext
    for ext in extensions:
        path = os.path.join(directory, f"{key}{ext}")
        if os.path.exists(path):
            return path

    # 2. Fuzzy Match Strategy
    # Collect all valid media files in directory
    candidates = []
    try:
        for f in os.listdir(directory):
            name, ext = os.path.splitext(f)
            if ext.lower() in extensions:
                candidates.append(name)
    except:
        return None
    
    if not candidates:
        return None

    # Find closest match
    # cutoff=0.6 means 60% similarity required
    matches = difflib.get_close_matches(key, candidates, n=1, cutoff=0.6)
    
    if matches:
        best_match_name = matches[0]
        # Reconstruct path (we need to find the extension again)
        for f in os.listdir(directory):
            if f.startswith(best_match_name) and os.path.splitext(f)[0] == best_match_name:
                 log(f"Fuzzy match: '{key}' -> '{f}'")
                 return os.path.join(directory, f)
            
    return None

FLOAT_PATTERN = re.compile(r"-?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?")
FULL_SCI_PATTERN = re.compile(r"-?(?:\d+\.\d*|\.\d+|\d+)[eE][-+]?\d+$")
EXPONENT_FRAGMENT_PATTERN = re.compile(r"[eE]\s*([+-]?\d+)")
DEFAULT_CROP_RATIO = 0.475
DEFAULT_AXIS_RATIOS = (0.25, 0.22, 0.20, 0.18, 0.155)

def extract_float(text):
    """Extract all float numbers from text, including decimals like .015 or -.015."""
    normalized = (text or "").replace(",", ".").replace("−", "-").replace("—", "-")
    return [float(x) for x in FLOAT_PATTERN.findall(normalized)]

def run_tesseract_custom(image, psm=6, whitelist='0123456789.eE+-'):
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False, mode='wb') as f:
        temp_name = f.name
        f.close()
        
    try:
        cv2.imwrite(temp_name, image)
        
        # Construct command
        cmd = [pytesseract.pytesseract.tesseract_cmd, temp_name, 'stdout']
        
        # Explicitly pass tessdata-dir to avoid env var encoding issues
        try:
            tess_base = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'third_party', 'tesseract', 'tessdata')
            if os.path.exists(tess_base):
                try:
                    short_tess = win32api.GetShortPathName(tess_base)
                    cmd.extend(['--tessdata-dir', short_tess])
                except:
                    cmd.extend(['--tessdata-dir', tess_base])
        except:
            pass
            
        cmd.extend(['--oem', '3'])
        cmd.extend(['--psm', str(psm)])
        cmd.extend(['-c', f'tessedit_char_whitelist={whitelist}'])
        cmd.extend(['-c', 'debug_file=NUL'])
        cmd.append('tsv') 
        
        # Run
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)
        stdout, stderr = proc.communicate()
        
        if proc.returncode != 0:
            log(f"Tesseract Error: {stderr.decode('utf-8', errors='replace')}")
            return None
            
        text = stdout.decode('utf-8', errors='replace')
        
        reader = csv.DictReader(io.StringIO(text), delimiter='\t', quoting=csv.QUOTE_NONE)
        data = {'text': [], 'top': [], 'height': [], 'left': [], 'width': [], 'conf': []}
        
        for row in reader:
            data['text'].append(row.get('text', ''))
            try:
                data['top'].append(int(row.get('top', 0)))
                data['height'].append(int(row.get('height', 0)))
                data['left'].append(int(row.get('left', 0)))
                data['width'].append(int(row.get('width', 0)))
                data['conf'].append(float(row.get('conf', -1)))
            except:
                data['top'].append(0)
                data['height'].append(0)
                data['left'].append(0)
                data['width'].append(0)
                data['conf'].append(-1.0)
                 
        return data
        
    except Exception as e:
        log(f"Custom Tesseract Failed: {e}")
        return None
    finally:
        if os.path.exists(temp_name):
            try: os.remove(temp_name)
            except: pass

def _normalize_numeric_text(text):
    cleaned = (text or "").strip().replace(" ", "").replace(",", ".")
    cleaned = cleaned.replace("−", "-").replace("—", "-")
    if cleaned.startswith("."):
        cleaned = f"0{cleaned}"
    elif cleaned.startswith("-."):
        cleaned = cleaned.replace("-.", "-0.", 1)
    return cleaned

def _prepare_axis_ocr_variants(axis_roi, scale=4):
    roi_large = cv2.resize(axis_roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    roi_blur = cv2.GaussianBlur(roi_large, (3, 3), 0)

    adaptive = cv2.adaptiveThreshold(
        roi_blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
    )
    adaptive = cv2.erode(adaptive, np.ones((2, 2), np.uint8), iterations=1)

    _, otsu = cv2.threshold(roi_blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return scale, [("adaptive", adaptive), ("otsu", otsu), ("gray", roi_large)]

def _dedupe_axis_candidates(candidates):
    deduped = []
    for candidate in sorted(candidates, key=lambda item: item['conf'], reverse=True):
        duplicate = False
        for kept in deduped:
            same_position = (
                abs(candidate['x_center'] - kept['x_center']) <= 8
                and abs(candidate['y_center'] - kept['y_center']) <= 6
            )
            if same_position:
                duplicate = True
                break
        if not duplicate:
            deduped.append(candidate)
    return deduped

def _infer_scientific_exponent(raw_tokens):
    magnitude_scores = {}
    negative_scores = {}
    positive_scores = {}
    loose_negative_tokens = 0

    for token in raw_tokens:
        text = token.get('text', '')
        conf = max(0.0, float(token.get('conf', 0.0)))
        weight = 1.0 + (conf / 100.0)

        if '-7' in text:
            magnitude_scores[7] = magnitude_scores.get(7, 0.0) + weight
            negative_scores[7] = negative_scores.get(7, 0.0) + 2.0 + weight
        if '+' in text:
            for match in re.findall(r"\+(\d+)", text):
                try:
                    exp_mag = abs(int(match))
                except ValueError:
                    continue
                magnitude_scores[exp_mag] = magnitude_scores.get(exp_mag, 0.0) + weight
                positive_scores[exp_mag] = positive_scores.get(exp_mag, 0.0) + 1.5 + weight
        if '-' in text:
            loose_negative_tokens += 1

        for match in EXPONENT_FRAGMENT_PATTERN.finditer(text):
            exp_text = match.group(1)
            try:
                exp = int(exp_text)
            except ValueError:
                continue

            exp_mag = abs(exp)
            magnitude_scores[exp_mag] = magnitude_scores.get(exp_mag, 0.0) + weight
            if exp_text.startswith('-'):
                negative_scores[exp_mag] = negative_scores.get(exp_mag, 0.0) + 2.0 + weight
            elif exp_text.startswith('+'):
                positive_scores[exp_mag] = positive_scores.get(exp_mag, 0.0) + 1.5 + weight

    if not magnitude_scores:
        return None

    exp_mag, score = max(magnitude_scores.items(), key=lambda item: (item[1], -item[0]))
    if score < 2.0:
        return None

    neg_score = negative_scores.get(exp_mag, 0.0) + (1.5 * loose_negative_tokens)
    pos_score = positive_scores.get(exp_mag, 0.0)
    sign = -1 if neg_score >= pos_score else 1
    return sign * exp_mag

def _build_scientific_notation_candidates(raw_tokens, exponent_hint, crop_x_start):
    if exponent_hint is None or exponent_hint > -3:
        return []

    synthesized = []
    for token in raw_tokens:
        text = token.get('text', '')
        conf = float(token.get('conf', 0.0))
        if conf < 15:
            continue

        normalized = _normalize_numeric_text(text)
        if not normalized or 'e' in normalized.lower():
            continue

        trimmed = normalized[:-1] if normalized.endswith('.') else normalized
        if any(ch == '.' for ch in trimmed):
            continue

        digits = ''.join(ch for ch in trimmed if ch.isdigit())
        if len(digits) < 4:
            continue

        significant = digits.lstrip('0')
        if len(significant) < 4:
            continue

        mantissa = int(significant) / (10 ** (len(significant) - 1))
        value = mantissa * (10 ** exponent_hint)
        synth_text = f"{mantissa:.{len(significant) - 1}f}e{exponent_hint}"

        x_rel = token['left'] / token['scale']
        y_rel = token['top'] / token['scale']
        w_box = max(1, int(round(token['width'] / token['scale'])))
        h_box = max(1, int(round(token['height'] / token['scale'])))

        synthesized.append({
            'val': value,
            'text': synth_text,
            'conf': conf + 8.0,
            'source': f"{token['source']}/sci-hint",
            'x_center': x_rel + w_box / 2.0,
            'y_center': y_rel + h_box / 2.0,
            'rect': (int(round(crop_x_start + x_rel)), int(round(y_rel)), w_box, h_box),
        })

    return synthesized

def _axis_model_is_confident(axis_model, roi_height):
    if not axis_model:
        return False

    inliers = axis_model.get('inliers', [])
    if len(inliers) < 4:
        return False

    span = axis_model['y_bottom_px'] - axis_model['y_top_px']
    if span < max(120.0, roi_height * 0.38):
        return False

    mean_conf = float(np.mean([item['conf'] for item in inliers])) if inliers else -1.0
    if mean_conf < 35.0:
        return False

    distinct_values = {
        round(float(item['val']), 6)
        for item in inliers
        if math.isfinite(float(item['val']))
    }
    return len(distinct_values) >= 4

def _axis_model_is_usable(axis_model, roi_height):
    if not axis_model:
        return False

    inliers = axis_model.get('inliers', [])
    if len(inliers) < 2:
        return False

    span = axis_model['y_bottom_px'] - axis_model['y_top_px']
    if span < max(260.0, roi_height * 0.55):
        return False

    mean_conf = float(np.mean([item['conf'] for item in inliers])) if inliers else -1.0
    if mean_conf < 40.0:
        return False

    distinct_values = {
        round(float(item['val']), 6)
        for item in inliers
        if math.isfinite(float(item['val']))
    }
    return len(distinct_values) >= 2

def _collect_axis_candidates(axis_roi, crop_x_start, scale=4, conf_threshold=25, positive_only=False):
    scale, variants = _prepare_axis_ocr_variants(axis_roi, scale=scale)
    candidates = []
    raw_tokens = []
    best_candidates = []
    best_model = None

    for source_name, processed in variants:
        for psm in (6, 11):
            data = run_tesseract_custom(processed, psm=psm)
            if not data:
                continue

            num_boxes = len(data['text'])
            for i in range(num_boxes):
                text = _normalize_numeric_text(data['text'][i])
                conf = data.get('conf', [-1] * num_boxes)[i]
                raw_tokens.append({
                    'text': text,
                    'conf': conf,
                    'left': data['left'][i],
                    'top': data['top'][i],
                    'width': data['width'][i],
                    'height': data['height'][i],
                    'scale': scale,
                    'source': f"{source_name}/psm{psm}",
                })
                if not text or conf < conf_threshold:
                    continue

                nums = extract_float(text)
                if not nums:
                    continue

                val = nums[0]
                x_rel = data['left'][i] / scale
                y_rel = data['top'][i] / scale
                w_box = max(1, int(round(data['width'][i] / scale)))
                h_box = max(1, int(round(data['height'][i] / scale)))

                candidates.append({
                    'val': val,
                    'text': text,
                    'conf': conf,
                    'source': f"{source_name}/psm{psm}",
                    'x_center': x_rel + w_box / 2.0,
                    'y_center': y_rel + h_box / 2.0,
                    'rect': (int(round(crop_x_start + x_rel)), int(round(y_rel)), w_box, h_box),
                })

            deduped = _dedupe_axis_candidates(candidates)
            axis_model = _select_axis_model(deduped, positive_only=positive_only)
            if axis_model:
                best_candidates = deduped
                best_model = axis_model
                if _axis_model_is_confident(axis_model, axis_roi.shape[0]):
                    return best_candidates, best_model

    exponent_hint = _infer_scientific_exponent(raw_tokens)
    if exponent_hint is not None:
        candidates.extend(_build_scientific_notation_candidates(raw_tokens, exponent_hint, crop_x_start))

    if best_model:
        deduped = _dedupe_axis_candidates(candidates)
        hinted_model = _select_axis_model(deduped, positive_only=positive_only)
        if hinted_model and _score_axis_model(hinted_model) >= _score_axis_model(best_model):
            return deduped, hinted_model
        return best_candidates, best_model

    deduped = _dedupe_axis_candidates(candidates)
    return deduped, _select_axis_model(deduped, positive_only=positive_only)

def _fit_axis_model(aligned, positive_only=False):
    if len(aligned) < 2:
        return None

    aligned = sorted(aligned, key=lambda item: item['y_center'])
    best = None

    for i in range(len(aligned) - 1):
        for j in range(i + 1, len(aligned)):
            top = aligned[i]
            bottom = aligned[j]
            y_span = bottom['y_center'] - top['y_center']
            if y_span < 12:
                continue

            slope = (bottom['val'] - top['val']) / y_span
            if abs(slope) < 1e-9 or slope >= 0:
                continue

            value_scale = max(abs(top['val']), abs(bottom['val']), abs(bottom['val'] - top['val']), 1e-9)
            tol = max(abs(bottom['val'] - top['val']) * 0.08, abs(slope) * 8, value_scale * 0.02)
            inliers = []
            residual_sum = 0.0

            for item in aligned:
                predicted = top['val'] + slope * (item['y_center'] - top['y_center'])
                residual = abs(item['val'] - predicted)
                if residual <= tol:
                    if positive_only and item['val'] < 0:
                        continue
                    inliers.append(item)
                    residual_sum += residual

            if len(inliers) < 2:
                continue

            y_values = [item['y_center'] for item in inliers]
            score = (
                len(inliers),
                max(y_values) - min(y_values),
                -residual_sum,
                np.mean([item['conf'] for item in inliers]),
            )

            if (best is None) or (score > best['score']):
                best = {'score': score, 'inliers': sorted(inliers, key=lambda item: item['y_center'])}

    if not best:
        return None

    inliers = best['inliers']
    y = np.array([item['y_center'] for item in inliers], dtype=np.float64)
    v = np.array([item['val'] for item in inliers], dtype=np.float64)

    if len(inliers) >= 3:
        slope, intercept = np.polyfit(y, v, 1)
    else:
        slope = (v[-1] - v[0]) / (y[-1] - y[0])
        intercept = v[0] - slope * y[0]

    if abs(slope) < 1e-9 or slope >= 0:
        return None

    top = inliers[0]
    bottom = inliers[-1]
    if positive_only and (top['val'] < 0 or bottom['val'] < 0):
        return None

    return {
        'inliers': inliers,
        'slope': float(slope),
        'intercept': float(intercept),
        'y_top_px': float(top['y_center']),
        'val_top': float(top['val']),
        'y_bottom_px': float(bottom['y_center']),
        'val_bottom': float(bottom['val']),
    }

def _select_axis_model(axis_candidates, positive_only=False):
    if len(axis_candidates) < 2:
        return None

    bin_width = 20
    bins = {}
    for candidate in axis_candidates:
        bin_idx = int(candidate['x_center'] / bin_width)
        bins.setdefault(bin_idx, []).append(candidate)

    best = None
    for bin_idx, items in bins.items():
        median_x = float(np.median([item['x_center'] for item in items]))
        aligned = [item for item in axis_candidates if abs(item['x_center'] - median_x) < 15]
        model = _fit_axis_model(aligned, positive_only=positive_only)
        if not model:
            continue

        score = _score_axis_model(model)
        if (best is None) or (score > best['score']):
            best = {
                'score': score,
                'bin_idx': bin_idx,
                'median_x': median_x,
                **model,
            }

    return best

def _score_axis_model(axis_model):
    if not axis_model:
        return None

    inliers = axis_model.get('inliers', [])
    full_sci = sum(1 for item in inliers if FULL_SCI_PATTERN.fullmatch(item.get('text', '')))
    truncated_sci = sum(1 for item in inliers if 'e' in item.get('text', '').lower() and not FULL_SCI_PATTERN.fullmatch(item.get('text', '')))
    return (
        full_sci,
        len(inliers),
        axis_model['y_bottom_px'] - axis_model['y_top_px'],
        -truncated_sci,
        np.mean([item['conf'] for item in inliers]) if inliers else -1,
    )

def _resolve_axis_model(gray, crop_x_start, axis_ratio, positive_only=False):
    ratios = [axis_ratio] if axis_ratio is not None else list(DEFAULT_AXIS_RATIOS)
    best = None

    for ratio in ratios:
        axis_width = int(gray.shape[1] * ratio)
        axis_roi = gray[:, :axis_width]
        axis_candidates, axis_model = _collect_axis_candidates(
            axis_roi,
            crop_x_start,
            positive_only=positive_only,
        )
        if len(axis_candidates) < 2 or not axis_model:
            continue

        if not axis_model:
            continue

        if positive_only:
            axis_model = _maybe_rebase_positive_axis(axis_model, gray.shape[0])

        score = _score_axis_model(axis_model)
        if (best is None) or (score > best['score']):
            best = {
                'score': score,
                'axis_ratio': ratio,
                'axis_width': axis_width,
                'axis_candidates': axis_candidates,
                'axis_model': axis_model,
            }
        if axis_ratio is None and _axis_model_is_confident(axis_model, gray.shape[0]):
            return best
        if axis_ratio is None and ratio == ratios[0] and _axis_model_is_usable(axis_model, gray.shape[0]):
            return best

    return best

def _maybe_rebase_positive_axis(axis_model, roi_height):
    if not axis_model or len(axis_model.get('inliers', [])) < 4:
        return axis_model

    vals = [item['val'] for item in axis_model['inliers']]
    if any(val <= 0 for val in vals):
        return axis_model

    decimal_vals = [val for val in vals if abs(val - round(val)) > 1e-6 and abs(val) >= 1]
    if len(decimal_vals) < max(4, len(vals) // 2):
        return axis_model

    integer_parts = [int(np.floor(val)) for val in decimal_vals]
    if len(set(integer_parts)) != 1:
        return axis_model

    common_integer = integer_parts[0]
    if common_integer <= 0:
        return axis_model

    current_zero_y = -axis_model['intercept'] / axis_model['slope']
    target_zero_y = roi_height - 1

    adjusted_items = []
    for item in axis_model['inliers']:
        adjusted = dict(item)
        adjusted['val'] = item['val'] - common_integer
        adjusted_items.append(adjusted)

    adjusted_model = _fit_axis_model(adjusted_items, positive_only=True)
    if not adjusted_model:
        return axis_model

    adjusted_zero_y = -adjusted_model['intercept'] / adjusted_model['slope']
    current_distance = abs(current_zero_y - target_zero_y)
    adjusted_distance = abs(adjusted_zero_y - target_zero_y)

    if adjusted_distance + 40 < current_distance:
        return {
            **axis_model,
            **adjusted_model,
            'rebased_by': common_integer,
        }

    return axis_model

def _mask_curve_noise(mask):
    clean = mask.copy()
    h_chart, w_chart = clean.shape[:2]

    mask_h = int(h_chart * 0.10)
    mask_w = int(w_chart * 0.45)
    clean[0:mask_h, (w_chart - mask_w):] = 0

    mask_bottom_h = int(h_chart * 0.12)
    clean[(h_chart - mask_bottom_h):, :] = 0

    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))
    clean = cv2.subtract(clean, cv2.morphologyEx(clean, cv2.MORPH_OPEN, kernel_h))
    clean = cv2.subtract(clean, cv2.morphologyEx(clean, cv2.MORPH_OPEN, kernel_v))

    return clean

def _select_curve_component(mask):
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    h_chart, w_chart = mask.shape[:2]
    best = None

    for idx in range(1, num_labels):
        x, y, w_comp, h_comp, area = stats[idx]
        if area < 18 or w_comp < 4:
            continue

        aspect_ratio = h_comp / float(max(w_comp, 1))
        if aspect_ratio > 10 and h_comp > (h_chart * 0.2):
            continue

        x_end = x + w_comp - 1
        score = (
            x_end,
            min(w_comp, int(w_chart * 0.5)),
            area,
            -abs((y + h_comp / 2.0) - (h_chart / 2.0)),
        )

        if (best is None) or (score > best['score']):
            best = {
                'score': score,
                'label': idx,
                'bbox': (x, y, w_comp, h_comp),
                'x_end': x_end,
                'area': area,
                'labels': labels,
            }

    return best

def _detect_curve_endpoint(chart_bgr):
    hsv = cv2.cvtColor(chart_bgr, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(chart_bgr, cv2.COLOR_BGR2GRAY)

    color_mask = np.zeros_like(gray)
    color_mask[(hsv[:, :, 1] > 35) & (hsv[:, :, 2] < 250)] = 255

    _, dark_mask = cv2.threshold(gray, 210, 255, cv2.THRESH_BINARY_INV)

    best = None
    for mask_name, base_mask, bonus in (("color", color_mask, 5), ("dark", dark_mask, 0)):
        clean_mask = _mask_curve_noise(base_mask)
        component = _select_curve_component(clean_mask)
        if not component:
            continue

        label_mask = component['labels'] == component['label']
        ys, xs = np.where(label_mask)
        if xs.size == 0:
            continue

        x_max = int(xs.max())
        band = xs >= max(0, x_max - 4)
        if band.sum() < 3:
            band = xs >= max(0, x_max - 8)
        if band.sum() == 0:
            continue

        band_points = {}
        for x_val, y_val in zip(xs[band], ys[band]):
            band_points.setdefault(int(x_val), []).append(float(y_val))

        y_candidates = [float(np.median(values)) for values in band_points.values()]
        y_end = float(np.median(y_candidates))

        score = (
            component['x_end'] + bonus,
            component['bbox'][2],
            component['area'],
        )

        if (best is None) or (score > best['score']):
            best = {
                'score': score,
                'mask_name': mask_name,
                'mask': clean_mask,
                'bbox': component['bbox'],
                'x_end': x_max,
                'y_end': y_end,
            }

    return best

def analyze_curve_image(image_path, positive_only=False, crop_ratio=DEFAULT_CROP_RATIO, axis_ratio=None):
    """
    Analyze a curve image and return OCR/curve debug details plus the mapped value.
    """
    if not os.path.exists(image_path):
        return {'ok': False, 'error': f'Image not found: {image_path}'}

    try:
        img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        img = None

    if img is None:
        return {'ok': False, 'error': f'Failed to decode image: {image_path}'}

    h, w = img.shape[:2]
    crop_x_start = int(w * crop_ratio)
    roi = img[:, crop_x_start:]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

    axis_result = _resolve_axis_model(gray, crop_x_start, axis_ratio, positive_only=positive_only)
    if not axis_result:
        return {
            'ok': False,
            'error': 'OCR failed to find enough Y-axis labels.',
            'img': img,
            'crop_x_start': crop_x_start,
            'axis_width': int(roi.shape[1] * (axis_ratio or DEFAULT_AXIS_RATIOS[0])),
            'axis_candidates': [],
        }

    axis_width = axis_result['axis_width']
    axis_candidates = axis_result['axis_candidates']
    axis_model = axis_result['axis_model']
    axis_ratio = axis_result['axis_ratio']

    chart_start_x = crop_x_start + axis_width
    chart_bgr = roi[:, axis_width:]
    curve = _detect_curve_endpoint(chart_bgr)
    if not curve:
        return {
            'ok': False,
            'error': 'No curve detected.',
            'img': img,
            'crop_x_start': crop_x_start,
            'axis_width': axis_width,
            'axis_candidates': axis_candidates,
            'axis_model': axis_model,
            'chart_start_x': chart_start_x,
        }

    y_end_roi = curve['y_end']
    value = axis_model['slope'] * y_end_roi + axis_model['intercept']
    if positive_only and axis_model['val_top'] >= 0 and axis_model['val_bottom'] >= 0 and value < 0:
        return {
            'ok': False,
            'error': 'Mapped value became negative under positive-only mode.',
            'img': img,
            'crop_x_start': crop_x_start,
            'axis_width': axis_width,
            'axis_candidates': axis_candidates,
            'axis_model': axis_model,
            'chart_start_x': chart_start_x,
            'curve': curve,
        }

    return {
        'ok': True,
        'value': float(value),
        'img': img,
        'crop_x_start': crop_x_start,
        'axis_ratio': axis_ratio,
        'axis_width': axis_width,
        'axis_candidates': axis_candidates,
        'axis_model': axis_model,
        'chart_start_x': chart_start_x,
        'curve': curve,
    }

def detect_curve_end_value(image_path, positive_only=False, crop_ratio=DEFAULT_CROP_RATIO, axis_ratio=None):
    """
    Detect the value of the curve end point on the right-side chart.
    """
    result = analyze_curve_image(
        image_path,
        positive_only=positive_only,
        crop_ratio=crop_ratio,
        axis_ratio=axis_ratio,
    )
    if not result.get('ok'):
        log(result.get('error', f"Curve analysis failed for {image_path}"))
        return None

    axis_model = result['axis_model']
    curve = result['curve']
    log(
        "DEBUG: Curve end from "
        f"{curve['mask_name']} mask at y={curve['y_end']:.1f}, x={curve['x_end']}"
    )
    log(
        "DEBUG: Axis inliers "
        f"{[(round(item['y_center'], 1), item['val']) for item in axis_model['inliers']]}"
    )
    log(
        "DEBUG: Top/Bottom "
        f"{axis_model['val_top']}@{axis_model['y_top_px']:.1f}px -> "
        f"{axis_model['val_bottom']}@{axis_model['y_bottom_px']:.1f}px"
    )

    return result['value']
