import os
import cv2
import win32com.client
import argparse
import time
import datetime
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Mapping from Key to Filename
VIDEO_MAPPING = {
    "wendu": {"file": "wendu.mp4", "slide": 4, "img_slide": 5},
    "juanqi": {"file": "juanqi.mp4", "slide": 6, "img_slide": 8},
    "tijuanqi": {"file": "tijijuanqi.mp4", "slide": 7, "img_slide": 9},
    "qipao": {"file": "qipao.mp4", "slide": 10, "img_slide": 11},
    "yanghuazha": {"file": "yanghuazha.mp4", "slide": 12, "img_slide": 14},
    "tijiyanghua": {"file": "tijiyanghua.mp4", "slide": 13, "img_slide": 15},
    "liaoliuzhuizong": {"file": "liaoliuzhuizong.mp4", "slide": 16, "img_slide": 17},
    "suokong": {"file": "suokong.mp4", "slide": 18, "img_slide": 21},
    "suokonglv": {"file": "suokonglv.mp4", "slide": 19, "img_slide": 22},
    "suokongtiji": {"file": "suokongtiji.mp4", "slide": 20, "img_slide": 22},
    "rejie": {"file": "rejie.mp4", "slide": 23, "img_slide": 24},
    "qiya": {"file": "qiya.mp4", "slide": 25, "img_slide": 26},
}

def log(message):
    print(message, flush=True)

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
        # Use imdecode to handle paths with non-ascii characters if needed, 
        # though standard imread usually works if locale is set right.
        # Safe way for windows paths:
        img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), -1)
        if img is not None:
            h, w = img.shape[:2]
            return (w, h)
    except Exception:
        pass

    return None

def calculate_centered_rect(container, content_size):
    """
    Calculate new rect to fit content inside container while maintaining aspect ratio and centering.
    container: (left, top, width, height)
    content_size: (content_width, content_height)
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
    
    # Vertically align with bias (0.3 = top 30%, bottom 70%)
    # This moves content up closer to the title as requested
    new_top = c_top + (c_height - new_height) * 0.3
    
    return (new_left, new_top, new_width, new_height)

def extract_last_frame(video_path, output_image_path):
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

    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_count - 1)
    ret, frame = cap.read()

    if ret:
        is_success, im_buf_arr = cv2.imencode(".png", frame)
        if is_success:
            im_buf_arr.tofile(output_image_path)
            log(f"Extracted last frame to {output_image_path}")
            cap.release()
            return True
        log(f"Warning: Could not encode frame for {video_path}")
    else:
        log(f"Warning: Could not read last frame of {video_path}")

    cap.release()
    return False

def _split_rect(left, top, width, height, gap):
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

def find_largest_media(slide):
    """Finds the largest shape on the slide, assuming it's the main content."""
    max_area = 0
    target_shape = None
    
    # Iterate all shapes
    for shape in slide.Shapes:
        # Filter out obvious non-content shapes
        # 1. Skip Title Placeholders
        if shape.Type == 14: # msoPlaceholder
            try:
                if shape.PlaceholderFormat.Type in [1, 3]: # Title or Center Title
                    continue
            except:
                pass
        
        # 2. Skip very small shapes (likely page numbers, logos)
        # Assuming slide size is roughly 960x540 or similar
        if shape.Width < 50 or shape.Height < 50:
            continue
            
        area = shape.Width * shape.Height
        if area > max_area:
            max_area = area
            target_shape = shape
            
    return target_shape

def get_content_area(pres, slide):
    """
    Calculates the safe content area on the slide, excluding the title.
    Returns: (left, top, width, height)
    """
    slide_w = pres.PageSetup.SlideWidth
    slide_h = pres.PageSetup.SlideHeight
    
    # Standard margins
    margin_x = 20.0
    margin_bottom = 20.0
    
    # Determine top margin based on Title
    top_margin = 60.0 # Default fallback
    
    try:
        if slide.Shapes.HasTitle:
            title = slide.Shapes.Title
            # Ensure title is at the top
            if title.Top < slide_h / 2: 
                top_margin = title.Top + title.Height + 10.0
    except:
        pass
        
    # Define rect
    left = margin_x
    top = top_margin
    width = slide_w - (2 * margin_x)
    height = slide_h - top - margin_bottom
    
    # Safety check: if height is too small, fallback to a generous default
    if height < slide_h * 0.5: 
        top = 60.0
        height = slide_h - top - margin_bottom
        
    return (left, top, width, height)

def generate_append_report(existing_ppt_path, video_dir, output_path, gap):
    log("Starting append report generation...")
    log(f"Existing PPT: {existing_ppt_path}")
    log(f"New Video Directory: {video_dir}")
    log(f"Output Path: {output_path}")

    if not os.path.exists(existing_ppt_path):
        log(f"Error: Existing PPT does not exist: {existing_ppt_path}")
        return 2
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return 2

    log("Step 1: Extracting frames for new videos...")
    temp_dir = BASE_DIR
    temp_images = {}

    for key, info in VIDEO_MAPPING.items():
        filename = info["file"]
        video_path = os.path.join(video_dir, filename)
        out_path = os.path.join(temp_dir, f"temp_{key}_new.png")
        if os.path.exists(video_path) and extract_last_frame(video_path, out_path):
            temp_images[key] = out_path
        else:
             if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except: pass

    log("Step 2: Opening PPT...")
    try:
        try:
            ppt_app = win32com.client.Dispatch("Kwpp.Application")
            log("Using WPS Presentation.")
        except Exception:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            log("Using Microsoft PowerPoint.")
    except Exception:
        log("Error: Could not open WPS or PowerPoint.")
        return 3

    try:
        ppt_app.Visible = True
    except Exception:
        pass

    try:
        pres = ppt_app.Presentations.Open(existing_ppt_path)
    except Exception as e:
        log(f"Error opening PPT: {e}")
        return 4

    log("Step 3: Processing slides...")
    
    # Process Video Slides
    for key, info in VIDEO_MAPPING.items():
        slide_idx = info["slide"]
        filename = info["file"]
        video_path = os.path.join(video_dir, filename)
        
        if not os.path.exists(video_path):
            log(f"Skipping {key}: New video not found.")
            continue
            
        try:
            # Slides is 1-based
            if slide_idx > pres.Slides.Count:
                log(f"Warning: Slide {slide_idx} out of range.")
                continue
                
            slide = pres.Slides(slide_idx)
            
            # Find existing video/content
            existing_shape = find_largest_media(slide)
            
            if existing_shape:
                log(f"Found existing content on Slide {slide_idx} for {key}")
                
                # Optimization: Maximize the container width to use full slide width
                # This helps reduce white space for side-by-side layout
                container_left, container_top, container_width, container_height = get_content_area(pres, slide)
                
                # Calculate split layout using the maximized container
                (a_left, a_top, a_width, a_height), (b_left, b_top, b_width, b_height) = _split_rect(
                    container_left, container_top, container_width, container_height, gap
                )
                
                # --- Update for Aspect Ratio Fit (Left Side - Existing) ---
                # Note: Existing shape might already have a specific aspect ratio. 
                # We want to fit it into the left slot (a_rect) while keeping its ratio if possible.
                # Or we assume the existing shape *is* the content and we just want to shrink it to fit A?
                # The user requirement: "原始视频/图片保持原有尺寸不变（正确）" 
                # But also "左右内容应等比例平分屏幕空间".
                # To be consistent with compare report, we should treat the existing shape as content 
                # that needs to be fitted into the Left Box (A).
                
                # 1. Move existing shape to fit into Left Box (A) while keeping aspect ratio
                rect_a = (a_left, a_top, a_width, a_height)
                
                # We can get current size of existing shape as "content size"
                # But wait, existing shape might be huge. We should scale it down to fit rect_a.
                existing_w = existing_shape.Width
                existing_h = existing_shape.Height
                
                # Calculate new pos/size for existing shape in A
                new_a_left, new_a_top, new_a_width, new_a_height = calculate_centered_rect(rect_a, (existing_w, existing_h))
                
                existing_shape.Left = new_a_left
                existing_shape.Top = new_a_top
                existing_shape.Width = new_a_width
                existing_shape.Height = new_a_height
                
                # --- Update for Aspect Ratio Fit (Right Side - New Video) ---
                rect_b = (b_left, b_top, b_width, b_height)
                final_b_left, final_b_top, final_b_width, final_b_height = rect_b
                
                if os.path.exists(video_path):
                    size_b = get_media_size(video_path)
                    if size_b:
                        final_b_left, final_b_top, final_b_width, final_b_height = calculate_centered_rect(rect_b, size_b)

                # Insert New Video to Right (B)
                log(f"Inserting new video {filename} to right side...")
                try:
                    s = slide.Shapes.AddMediaObject2(video_path, False, True, final_b_left, final_b_top, final_b_width, final_b_height)
                except Exception:
                    s = slide.Shapes.AddMediaObject(video_path, final_b_left, final_b_top, final_b_width, final_b_height)
                
                s.Left = final_b_left
                s.Top = final_b_top
                s.Width = final_b_width
                s.Height = final_b_height
                
            else:
                log(f"Warning: No existing content found on Slide {slide_idx} to compare against.")
                # Optional: Just insert new video centered? For now, skipping to avoid mess.
                
        except Exception as e:
            log(f"Error processing video slide {slide_idx} for {key}: {e}")

    # Process Image Slides
    for key, info in VIDEO_MAPPING.items():
        slide_idx = info["img_slide"]
        img_path = temp_images.get(key)
        
        if not img_path or not os.path.exists(img_path):
            continue
            
        try:
            if slide_idx > pres.Slides.Count:
                continue
            
            slide = pres.Slides(slide_idx)
            existing_shape = find_largest_media(slide)
            
            if existing_shape:
                log(f"Found existing image on Slide {slide_idx} for {key} (Shape: {existing_shape.Name})")
                
                # Force aspect ratio lock off to ensure we can resize freely if needed
                # though we are calculating ratio ourselves.
                try:
                    existing_shape.LockAspectRatio = 0 # msoFalse
                except:
                    pass

                # Optimization: Maximize the container width to use full slide width
                # This helps reduce white space for side-by-side layout
                container_left, container_top, container_width, container_height = get_content_area(pres, slide)
                
                # Calculate split layout using the maximized container
                (a_left, a_top, a_width, a_height), (b_left, b_top, b_width, b_height) = _split_rect(
                    container_left, container_top, container_width, container_height, gap
                )
                
                # --- Update for Aspect Ratio Fit (Left Side - Existing) ---
                rect_a = (a_left, a_top, a_width, a_height)
                existing_w = existing_shape.Width
                existing_h = existing_shape.Height
                
                new_a_left, new_a_top, new_a_width, new_a_height = calculate_centered_rect(rect_a, (existing_w, existing_h))
                
                existing_shape.Left = new_a_left
                existing_shape.Top = new_a_top
                existing_shape.Width = new_a_width
                existing_shape.Height = new_a_height
                
                # --- Update for Aspect Ratio Fit (Right Side - New Image) ---
                rect_b = (b_left, b_top, b_width, b_height)
                final_b_left, final_b_top, final_b_width, final_b_height = rect_b
                
                if os.path.exists(img_path):
                     size_b = get_media_size(img_path)
                     if size_b:
                         final_b_left, final_b_top, final_b_width, final_b_height = calculate_centered_rect(rect_b, size_b)
                
                log(f"Inserting new image for {key} to right side...")
                slide.Shapes.AddPicture(img_path, False, True, final_b_left, final_b_top, final_b_width, final_b_height)
            else:
                log(f"Warning: No existing content found on Slide {slide_idx} for {key}. Skipping move.")
                # If we don't find existing content, we might still want to insert the new one?
                # For now, let's just insert the new one on the right side of the *page*?
                # No, that's risky. Let's stick to the current logic but log the failure clearly.
                pass
                
        except Exception as e:
            log(f"Error processing image slide {slide_idx} for {key}: {e}")

    log("Step 4: Saving report...")
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception as e:
            log(f"Warning: Could not remove existing output file: {e}")

    try:
        pres.SaveAs(output_path)
        log(f"Saved to {output_path}")
    except Exception as e:
        log(f"Error saving: {e}")
        try:
            pres.Close()
            ppt_app.Quit()
        except: pass
        return 5

    try:
        pres.Close()
        ppt_app.Quit()
    except: pass

    log("Step 5: Cleaning up...")
    for p in temp_images.values():
        if p and os.path.exists(p):
            try: os.remove(p)
            except: pass

    log("Done!")
    return 0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Append new videos to existing PPT for comparison")
    parser.add_argument("--existing_ppt", required=True, help="Path to existing PPT file")
    parser.add_argument("--video_dir", required=True, help="Path to new video directory")
    parser.add_argument("--output_path", help="Path to output PPT")
    parser.add_argument("--gap", default=10, type=float, help="Gap between A and B in points")
    
    args = parser.parse_args()

    output_path = args.output_path
    if not output_path:
        video_folder_name = os.path.basename(os.path.normpath(args.video_dir))
        current_date = time.strftime("%Y.%m.%d")
        base_name = os.path.splitext(os.path.basename(args.existing_ppt))[0]
        output_path = os.path.join(BASE_DIR, f"{base_name}-vs-{video_folder_name}-{current_date}.pptx")

    raise SystemExit(
        generate_append_report(
            existing_ppt_path=args.existing_ppt,
            video_dir=args.video_dir,
            output_path=output_path,
            gap=args.gap,
        )
    )
