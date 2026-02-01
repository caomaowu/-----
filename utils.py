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

def extract_float(text):
    """Extract all float numbers from text"""
    # Pattern for scientific notation (e.g. 1.23e-04) and standard floats
    return [float(x) for x in re.findall(r"-?\d+\.?\d*(?:[eE][-+]?\d+)?", text)]

def run_tesseract_custom(image):
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
        cmd.extend(['--psm', '6'])
        cmd.extend(['-c', 'tessedit_char_whitelist=0123456789.eE+-'])
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
        data = {'text': [], 'top': [], 'height': [], 'left': [], 'width': []}
        
        for row in reader:
            data['text'].append(row.get('text', ''))
            try:
                data['top'].append(int(row.get('top', 0)))
                data['height'].append(int(row.get('height', 0)))
                data['left'].append(int(row.get('left', 0)))
                data['width'].append(int(row.get('width', 0)))
            except:
                data['top'].append(0)
                data['height'].append(0)
                data['left'].append(0)
                data['width'].append(0)
                
        return data
        
    except Exception as e:
        log(f"Custom Tesseract Failed: {e}")
        return None
    finally:
        if os.path.exists(temp_name):
            try: os.remove(temp_name)
            except: pass

def detect_curve_end_value(image_path):
    """
    Detects the Y-coordinate value of the curve's end point in the right half of the image.
    """
    if not os.path.exists(image_path):
        return None
    
    # Use imdecode to handle paths with non-ascii characters
    try:
        img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        return None
        
    if img is None:
        return None
        
    h, w = img.shape[:2]
    
    # 1. Crop Right Half (Optimized based on user feedback)
    # Crop Start X: 47.5%
    crop_x_start = int(w * 0.475)
    roi = img[:, crop_x_start:]
    roi_h, roi_w = roi.shape[:2]
    
    # 2. Preprocess
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    
    # 3. OCR Y-Axis (Left side of ROI)
    # Axis Zone Width: 15.5%
    axis_width = int(roi_w * 0.155)
    axis_roi = gray[:, :axis_width]
    
    # Upscale for better OCR
    scale = 2
    roi_large = cv2.resize(axis_roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    
    data = run_tesseract_custom(roi_large)
    if not data:
        return None
        
    y_values = []
    num_boxes = len(data['text'])
    for i in range(num_boxes):
        text = data['text'][i].strip()
        if not text: continue
        
        nums = extract_float(text)
        if nums:
            val = nums[0]
            # data['top'] and ['height'] are in scaled coordinates
            # Center Y of the text box relative to original ROI
            y_center = (data['top'][i] + data['height'][i]/2) / scale
            
            # X center relative to ROI (for alignment check)
            x_center = (data['left'][i] + data['width'][i]/2) / scale
            
            y_values.append({
                'val': val,
                'y_center': y_center,
                'x_center': x_center
            })
            
    if len(y_values) < 2:
        log(f"OCR failed to find enough Y-axis labels in {image_path}")
        return None
        
    # Improved Axis Separation Logic
    # Goal: Separate Y-axis labels (vertically stacked) from X-axis labels (horizontally spread at bottom)
    
    # 1. Cluster by X-coordinate
    # Y-axis labels should have very similar X-coordinates (low variance)
    # X-axis labels will have different X-coordinates (high variance)
    
    # We use a histogram/binning approach to find the "main vertical column"
    x_coords = [v['x_center'] for v in y_values]
    
    # Bin width of 20 pixels
    bin_width = 20
    bins = {}
    
    for v in y_values:
        bin_idx = int(v['x_center'] / bin_width)
        if bin_idx not in bins:
            bins[bin_idx] = []
        bins[bin_idx].append(v)
        
    # Find the bin with the most items (assuming Y-axis has more labels than X-axis labels that accidentally fell into this zone)
    # Or simply, the Y-axis is usually the leftmost dominant column in the search zone
    
    best_bin = None
    max_count = 0
    
    for bin_idx, items in bins.items():
        if len(items) > max_count:
            max_count = len(items)
            best_bin = items
        elif len(items) == max_count:
            # If tie, prefer the leftmost one (smaller bin_idx)
            if best_bin and bin_idx < int(best_bin[0]['x_center'] / bin_width):
                best_bin = items
    
    if not best_bin or len(best_bin) < 2:
        log(f"Could not identify a valid Y-axis column in {image_path}")
        return None
        
    # 2. Refine the selected column
    # Calculate median X of the best bin
    selected_x = [v['x_center'] for v in best_bin]
    median_x = np.median(selected_x)
    
    # Strict filter: keep points within +/- 15px of median X
    aligned_y_values = []
    for v in y_values: # Check ALL points again, not just bin, to catch edge cases
        if abs(v['x_center'] - median_x) < 15:
            aligned_y_values.append((v['y_center'], v['val']))
            
    if len(aligned_y_values) < 2:
        log(f"Not enough aligned Y-axis labels found in {image_path}")
        return None
        
    y_values = aligned_y_values
        
    # Sort by Y position (top to bottom)
    y_values.sort(key=lambda x: x[0])
    
    # Filter outliers using slope consistency
    if len(y_values) > 2:
        # Calculate slopes between adjacent points
        slopes = []
        for i in range(len(y_values) - 1):
            dy = y_values[i+1][0] - y_values[i][0]
            dv = y_values[i+1][1] - y_values[i][1]
            if abs(dv) > 1e-9:
                slopes.append(dy / dv)
        
        if slopes:
            median_slope = np.median(slopes)
            
            # Filter based on consistency with median slope from a pivot point
            mid_idx = len(y_values) // 2
            pivot = y_values[mid_idx]
            
            filtered_y_values = []
            for p in y_values:
                # Always keep pivot
                if p == pivot:
                    filtered_y_values.append(p)
                    continue
                    
                dy = p[0] - pivot[0]
                dv = p[1] - pivot[1]
                
                if abs(dv) < 1e-9:
                    continue # Ignore points with same value as pivot (unless very close in Y)
                    
                slope = dy / dv
                
                # Check deviation from median slope
                # If slope has same sign and magnitude is within factor of 3
                if (slope * median_slope > 0) and (0.33 < abs(slope / median_slope) < 3.0):
                    filtered_y_values.append(p)
            
            if len(filtered_y_values) >= 2:
                y_values = filtered_y_values
                y_values.sort(key=lambda x: x[0])
                
    # Debug logs
    log(f"DEBUG: OCR found {len(y_values)} values: {y_values}")
    
    # Use top-most (max value) and bottom-most (min value) for mapping
    y_top_px, val_top = y_values[0]
    y_bottom_px, val_bottom = y_values[-1]
    
    log(f"DEBUG: Top: {val_top} at {y_top_px}px, Bottom: {val_bottom} at {y_bottom_px}px")
    
    pixel_range = y_bottom_px - y_top_px
    value_range = val_top - val_bottom # Top pixel corresponds to higher value
    
    if abs(pixel_range) < 10 or abs(value_range) == 0:
        log("Invalid axis detected (range too small)")
        return None
        
    # 4. Find Curve End (Rightmost dark pixel)
    # Exclude axis area to avoid detecting text as curve
    chart_area = gray[:, axis_width:] 
    
    # Threshold: assume curve is dark (< 200) on light background
    _, binary = cv2.threshold(chart_area, 200, 255, cv2.THRESH_BINARY_INV)
    
    # Remove grid lines (thin horizontal/vertical lines)
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
    detected_lines_h = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_h)
    
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 20))
    detected_lines_v = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_v)
    
    clean_binary = cv2.subtract(binary, detected_lines_h)
    clean_binary = cv2.subtract(clean_binary, detected_lines_v)
    
    points = cv2.findNonZero(clean_binary)
    if points is None:
        log(f"No curve detected in {image_path}")
        return None
        
    # points is (N, 1, 2) array of (x, y) relative to chart_area
    points = points[:, 0, :]
    
    # Sort by x descending (rightmost first)
    points = points[points[:, 0].argsort()[::-1]]
    
    # Take average of top 10 rightmost points
    top_n = min(10, len(points))
    rightmost_points = points[:top_n]
    
    y_end_avg_chart = np.mean(rightmost_points[:, 1])
    log(f"DEBUG: Curve End Y (relative to chart): {y_end_avg_chart}")
    
    # Map back to ROI coordinates (chart_area is offset by axis_width in X, but Y is same)
    y_end_roi = y_end_avg_chart
    
    # 5. Map y_end to value
    # Linear interpolation
    # ratio = (current_y - top_y) / (bottom_y - top_y)
    # val = top_val - ratio * (top_val - bottom_val)
    # Note: top_y is smaller pixel value (higher in image), top_val is usually larger number
    # bottom_y is larger pixel value (lower in image), bottom_val is usually smaller number
    
    # Check if axis is inverted (Top pixel val > Bottom pixel val)
    # Standard chart: Top Y (small px) -> Max Val, Bottom Y (large px) -> Min Val
    
    # Let's use a more robust linear fit if we have many points, but for now 2 points is enough
    # slope m = (val_bottom - val_top) / (y_bottom_px - y_top_px)
    # val = val_top + m * (y_current - y_top_px)
    
    slope = (val_bottom - val_top) / (y_bottom_px - y_top_px)
    value = val_top + slope * (y_end_roi - y_top_px)
    
    return value
