import os
import cv2
import numpy as np
import json
import sys

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
