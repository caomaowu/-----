import os
import cv2
import win32com.client
import time
import numpy as np
import argparse
import sys

# --- Configuration ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Mapping from Key to Filename
# The key corresponds to VID_<key> and IMG_<key>
VIDEO_MAPPING = {
    "wendu": "wendu.mp4",
    "juanqi": "juanqi.mp4",
    "tijuanqi": "tijijuanqi.mp4", # Note: handling discrepancy between md (tijuanqi) and file (tijijuanqi)
    "qipao": "qipao.mp4",
    "yanghuazha": "yanghuazha.mp4",
    "tijiyanghua": "tijiyanghua.mp4",
    "liaoliuzhuizong": "liaoliuzhuizong.mp4",
    "suokong": "suokong.mp4",
    "suokonglv": "suokonglv.mp4",
    "suokongtiji": "suokongtiji.mp4",
    "rejie": "rejie.mp4",
    "qiya": "qiya.mp4",
}

def log(message):
    """Helper to print with flush for real-time GUI updates"""
    print(message, flush=True)

def extract_last_frame(video_path, output_image_path):
    """Extracts the last frame of a video."""
    if not os.path.exists(video_path):
        log(f"Error: Video not found at {video_path}")
        return False
    
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        log(f"Error: Could not open video {video_path}")
        return False
    
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if frame_count == 0:
        log(f"Error: Video {video_path} has 0 frames")
        cap.release()
        return False

    # Set to last frame
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_count - 1)
    ret, frame = cap.read()
    
    if ret:
        # cv2.imwrite fails with unicode paths on Windows sometimes
        # Use imencode and file write
        is_success, im_buf_arr = cv2.imencode(".png", frame)
        if is_success:
            im_buf_arr.tofile(output_image_path)
            log(f"Extracted last frame to {output_image_path}")
            cap.release()
            return True
        else:
             log(f"Error: Could not encode frame for {video_path}")
    else:
        log(f"Error: Could not read last frame of {video_path}")
    
    cap.release()
    return False

def generate_report(video_dir, template_path, output_path):
    log("Starting report generation...")
    log(f"Video Directory: {video_dir}")
    log(f"Template Path: {template_path}")
    log(f"Output Path: {output_path}")
    
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return
    if not os.path.exists(template_path):
        log(f"Error: Template file does not exist: {template_path}")
        return

    # 1. Prepare Images
    log("Step 1: Extracting frames...")
    temp_images = {}
    
    # Ensure temp dir exists or use base dir
    # We'll use the directory where the script is, or system temp. 
    # Using script dir is safer for now to avoid permission issues.
    temp_dir = BASE_DIR
    
    for key, filename in VIDEO_MAPPING.items():
        video_path = os.path.join(video_dir, filename)
        image_path = os.path.join(temp_dir, f"temp_{key}.png")
        
        if os.path.exists(video_path):
            if extract_last_frame(video_path, image_path):
                temp_images[key] = image_path
        else:
            log(f"Warning: Video file {filename} missing in {video_dir}.")

    # 2. Open PPT Application
    log("Step 2: Opening PPT...")
    ppt_app = None
    pres = None
    try:
        # Try WPS first as requested in the plan ("WPS PPT")
        try:
            ppt_app = win32com.client.Dispatch("Kwpp.Application")
            log("Using WPS Presentation.")
        except Exception:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            log("Using Microsoft PowerPoint.")
    except Exception:
        log("Error: Could not open WPS or PowerPoint.")
        return

    try:
        ppt_app.Visible = True
    except:
        pass
    
    try:
        pres = ppt_app.Presentations.Open(template_path)
    except Exception as e:
        log(f"Error opening template: {e}")
        return

    # 3. Iterate slides and shapes
    log("Step 3: Processing slides...")
    
    # We collect actions to perform to avoid modifying collection while iterating
    # List of (slide_index, shape_name, action_type, key, left, top, width, height)
    actions = []

    try:
        for i, slide in enumerate(pres.Slides):
            # slide index starts at 1 in COM, but enumerate is 0-based. 
            # We use slide object directly.
            for shape in slide.Shapes:
                name = shape.Name
                if name.startswith("VID_"):
                    key = name.replace("VID_", "")
                    if key in VIDEO_MAPPING:
                        actions.append({
                            "slide_index": i + 1,
                            "shape_name": name,
                            "type": "video",
                            "key": key,
                            "left": shape.Left,
                            "top": shape.Top,
                            "width": shape.Width,
                            "height": shape.Height
                        })
                elif name.startswith("IMG_"):
                    key = name.replace("IMG_", "")
                    if key in VIDEO_MAPPING and key in temp_images:
                         actions.append({
                            "slide_index": i + 1,
                            "shape_name": name,
                            "type": "image",
                            "key": key,
                            "left": shape.Left,
                            "top": shape.Top,
                            "width": shape.Width,
                            "height": shape.Height
                        })
    except Exception as e:
        log(f"Error iterating slides: {e}")

    # 4. Execute Actions
    log(f"Found {len(actions)} anchors to process.")
    
    for action in actions:
        try:
            slide = pres.Slides(action["slide_index"])
            
            # Remove the placeholder shape
            try:
                shape = slide.Shapes(action["shape_name"])
                shape.Delete()
            except Exception as e:
                log(f"Warning: Could not delete shape {action['shape_name']}: {e}")
            
            # Insert new content
            if action["type"] == "video":
                video_file = os.path.join(video_dir, VIDEO_MAPPING[action["key"]])
                if os.path.exists(video_file):
                    log(f"Inserting video {video_file} at Slide {action['slide_index']}")
                    try:
                        # Try AddMediaObject2 first (Explicitly Embed: LinkToFile=False, SaveWithDocument=True)
                        new_shape = slide.Shapes.AddMediaObject2(video_file, False, True, action["left"], action["top"], action["width"], action["height"])
                    except:
                         # Fallback
                         new_shape = slide.Shapes.AddMediaObject(video_file, action["left"], action["top"], action["width"], action["height"])
                         
                    # Ensure properties
                    new_shape.Left = action["left"]
                    new_shape.Top = action["top"]
                    new_shape.Width = action["width"]
                    new_shape.Height = action["height"]
                    
            elif action["type"] == "image":
                image_file = temp_images[action["key"]]
                log(f"Inserting image {image_file} at Slide {action['slide_index']}")
                new_shape = slide.Shapes.AddPicture(image_file, False, True, action["left"], action["top"], action["width"], action["height"])
        except Exception as e:
            log(f"Error inserting content for {action['key']}: {e}")

    # 5. Save and Close
    log("Step 4: Saving report...")
    
    # Ensure output directory exists and file is removed
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception as e:
            log(f"Warning: Could not remove existing output file: {e}")

    try:
        # 11 = ppSaveAsOpenXMLPresentation (pptx)
        pres.SaveAs(output_path)
        log(f"Saved to {output_path}")
    except Exception as e:
        log(f"Error saving: {e}")
        
    try:
        pres.Close()
        ppt_app.Quit()
    except:
        pass

    # 6. Cleanup
    log("Step 5: Cleaning up...")
    for img_path in temp_images.values():
        if os.path.exists(img_path):
            try:
                os.remove(img_path)
            except:
                pass
    
    log("Done!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Automation Report")
    parser.add_argument("--video_dir", default=os.path.join(BASE_DIR, "视频"), help="Path to video directory")
    parser.add_argument("--template_path", default=os.path.join(BASE_DIR, "自动化模板.pptx"), help="Path to PPT template")
    parser.add_argument("--output_path", help="Path to output PPT")
    
    args = parser.parse_args()

    output_path = args.output_path
    if not output_path:
        video_folder_name = os.path.basename(os.path.normpath(args.video_dir))
        current_date = time.strftime("%Y.%m.%d")
        output_path = os.path.join(BASE_DIR, f"{video_folder_name}-模流分析报告-{current_date}.pptx")
    
    generate_report(args.video_dir, args.template_path, output_path)
