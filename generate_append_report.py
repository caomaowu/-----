import os
import cv2
import win32com.client
import argparse
import time
import datetime

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
    """Finds the largest media object (video) or picture on the slide."""
    max_area = 0
    target_shape = None
    for shape in slide.Shapes:
        # Check for Media (Type 16) or Picture (Type 13) or Placeholder (Type 14)
        # Note: Embedded videos might be type 16.
        if shape.Type in [13, 16] or (shape.Type == 14 and shape.PlaceholderFormat.ContainedType in [13, 16]):
             area = shape.Width * shape.Height
             if area > max_area:
                 max_area = area
                 target_shape = shape
    return target_shape

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
                
                # Get original dimensions
                orig_left = existing_shape.Left
                orig_top = existing_shape.Top
                orig_width = existing_shape.Width
                orig_height = existing_shape.Height
                
                # Calculate split layout
                (a_left, a_top, a_width, a_height), (b_left, b_top, b_width, b_height) = _split_rect(
                    orig_left, orig_top, orig_width, orig_height, gap
                )
                
                # Move existing shape to Left (A)
                existing_shape.Left = a_left
                existing_shape.Top = a_top
                existing_shape.Width = a_width
                existing_shape.Height = a_height
                
                # Insert New Video to Right (B)
                log(f"Inserting new video {filename} to right side...")
                try:
                    s = slide.Shapes.AddMediaObject2(video_path, False, True, b_left, b_top, b_width, b_height)
                except Exception:
                    s = slide.Shapes.AddMediaObject(video_path, b_left, b_top, b_width, b_height)
                
                s.Left = b_left
                s.Top = b_top
                s.Width = b_width
                s.Height = b_height
                
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
                log(f"Found existing image on Slide {slide_idx} for {key}")
                
                orig_left = existing_shape.Left
                orig_top = existing_shape.Top
                orig_width = existing_shape.Width
                orig_height = existing_shape.Height
                
                (a_left, a_top, a_width, a_height), (b_left, b_top, b_width, b_height) = _split_rect(
                    orig_left, orig_top, orig_width, orig_height, gap
                )
                
                existing_shape.Left = a_left
                existing_shape.Top = a_top
                existing_shape.Width = a_width
                existing_shape.Height = a_height
                
                log(f"Inserting new image for {key} to right side...")
                slide.Shapes.AddPicture(img_path, False, True, b_left, b_top, b_width, b_height)
                
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
