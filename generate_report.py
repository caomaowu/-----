import os
import cv2
import win32com.client
import time
import numpy as np

# --- Configuration ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VIDEO_DIR = os.path.join(BASE_DIR, "视频")
TEMPLATE_PATH = os.path.join(BASE_DIR, "自动化模板.pptx")
OUTPUT_PATH = os.path.join(BASE_DIR, "自动化报告_generated.pptx")

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
}

def extract_last_frame(video_path, output_image_path):
    """Extracts the last frame of a video."""
    if not os.path.exists(video_path):
        print(f"Error: Video not found at {video_path}")
        return False
    
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"Error: Could not open video {video_path}")
        return False
    
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if frame_count == 0:
        print(f"Error: Video {video_path} has 0 frames")
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
            print(f"Extracted last frame to {output_image_path}")
            cap.release()
            return True
        else:
             print(f"Error: Could not encode frame for {video_path}")
    else:
        print(f"Error: Could not read last frame of {video_path}")
    
    cap.release()
    return False

def generate_report():
    print("Starting report generation...")
    
    # 1. Prepare Images
    print("Step 1: Extracting frames...")
    temp_images = {}
    for key, filename in VIDEO_MAPPING.items():
        video_path = os.path.join(VIDEO_DIR, filename)
        image_path = os.path.join(BASE_DIR, f"temp_{key}.png")
        
        if os.path.exists(video_path):
            if extract_last_frame(video_path, image_path):
                temp_images[key] = image_path
        else:
            print(f"Warning: Video file {filename} missing.")

    # 2. Open PPT Application
    print("Step 2: Opening PPT...")
    try:
        # Try WPS first as requested in the plan ("WPS PPT")
        ppt_app = win32com.client.Dispatch("Kwpp.Application")
        print("Using WPS Presentation.")
    except Exception:
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            print("Using Microsoft PowerPoint.")
        except Exception:
            print("Error: Could not open WPS or PowerPoint.")
            return

    ppt_app.Visible = True
    
    try:
        pres = ppt_app.Presentations.Open(TEMPLATE_PATH)
    except Exception as e:
        print(f"Error opening template: {e}")
        return

    # 3. Iterate slides and shapes
    print("Step 3: Processing slides...")
    
    # We collect actions to perform to avoid modifying collection while iterating
    # List of (slide_index, shape_name, action_type, key, left, top, width, height)
    actions = []

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

    # 4. Execute Actions
    print(f"Found {len(actions)} anchors to process.")
    
    for action in actions:
        slide = pres.Slides(action["slide_index"])
        
        # Remove the placeholder shape
        try:
            shape = slide.Shapes(action["shape_name"])
            shape.Delete()
        except Exception as e:
            print(f"Warning: Could not delete shape {action['shape_name']}: {e}")
        
        # Insert new content
        try:
            if action["type"] == "video":
                video_file = os.path.join(VIDEO_DIR, VIDEO_MAPPING[action["key"]])
                if os.path.exists(video_file):
                    print(f"Inserting video {video_file} at Slide {action['slide_index']}")
                    # AddMediaObject(FileName, Left, Top, Width, Height)
                    # Note: Arguments might vary slightly between versions, but usually this works.
                    # For strict positioning, we might need to insert then move.
                    
                    # Microsoft PPT: AddMediaObject is deprecated/not standard in some versions, AddVideo is newer.
                    # WPS might behave differently.
                    # Let's try standard AddMediaObject first, usually works for older compat.
                    # Or AddVideo2 for newer MS PPT.
                    
                    # Safe approach: Insert then resize.
                    try:
                        # Try AddVideo (standard in newer Office)
                        new_shape = slide.Shapes.AddMediaObject(video_file, action["left"], action["top"], action["width"], action["height"])
                    except:
                         # Fallback or specific WPS method?
                         # Some docs suggest AddMovie for older versions.
                         # Let's try AddMediaObject again or check methods.
                         # If it fails, we might just print error.
                         new_shape = slide.Shapes.AddMediaObject2(video_file, False, True, action["left"], action["top"], action["width"], action["height"])
                         
                    # Ensure properties
                    new_shape.Left = action["left"]
                    new_shape.Top = action["top"]
                    new_shape.Width = action["width"]
                    new_shape.Height = action["height"]
                    
            elif action["type"] == "image":
                image_file = temp_images[action["key"]]
                print(f"Inserting image {image_file} at Slide {action['slide_index']}")
                new_shape = slide.Shapes.AddPicture(image_file, False, True, action["left"], action["top"], action["width"], action["height"])
        except Exception as e:
            print(f"Error inserting content for {action['key']}: {e}")

    # 5. Save and Close
    print("Step 4: Saving report...")
    
    # Ensure output directory exists and file is removed
    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
        except Exception as e:
            print(f"Warning: Could not remove existing output file: {e}")

    try:
        # 11 = ppSaveAsOpenXMLPresentation (pptx)
        pres.SaveAs(OUTPUT_PATH)
        print(f"Saved to {OUTPUT_PATH}")
    except Exception as e:
        print(f"Error saving: {e}")
        
    try:
        pres.Close()
        ppt_app.Quit()
    except:
        pass

    # 6. Cleanup
    print("Step 5: Cleaning up...")
    for img_path in temp_images.values():
        if os.path.exists(img_path):
            os.remove(img_path)
    
    print("Done!")

if __name__ == "__main__":
    generate_report()
