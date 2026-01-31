import os
import argparse
import time
from utils import log, extract_last_frame, load_config, find_media_file
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_report(video_dir, template_path, output_path):
    log("Starting report generation...")
    log(f"Video Directory: {video_dir}")
    
    # config = load_config(CONFIG_PATH) # No longer needed for video_mapping
    
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return

    # 1. PPT Automation
    ppt = PPTAutomation(template_path, output_path)
    if not ppt.open():
        return

    # 2. Scan and Process Media (First Pass)
    log("Scanning slides for media placeholders...")
    actions = ppt.scan_placeholders()
    log(f"Found {len(actions)} media anchors to process.")
    
    temp_images = []
    
    for action in actions:
        key = action["key"]
        action_type = action["type"]
        is_video_placeholder = (action_type == "video")
        
        # Find media file
        media_path = find_media_file(video_dir, key)
        
        if not media_path:
            log(f"Skipping {key}: Media not found in '{video_dir}'. Please check file name.")
            continue
            
        is_video_file = media_path.lower().endswith(('.mp4', '.avi', '.mov', '.wmv', '.mkv'))
        final_path = media_path

        # --- Handle Media Placeholder ---
        # Handle Image Placeholder needing Video Frame
        if not is_video_placeholder and is_video_file:
            temp_img = os.path.join(BASE_DIR, f"temp_{key}_{int(time.time())}.png")
            if extract_last_frame(media_path, temp_img):
                final_path = temp_img
                temp_images.append(final_path)
            else:
                log(f"Warning: Failed to extract frame for {key} from {media_path}")
                continue
                
        ppt.process_media_placeholder(action, final_path)

    # 3. Process Curves (Second Pass - After Media Insertion)
    log("Scanning for curve data placeholders...")
    curve_actions = ppt.scan_curve_tags()
    log(f"Found {len(curve_actions)} curve tags to process.")
    
    for action in curve_actions:
         key = action["key"]
         slide = action["slide"]
         
         # Export current slide to image
         temp_slide_img = os.path.join(BASE_DIR, f"temp_slide_{slide.SlideIndex}_{int(time.time())}.png")
         temp_images.append(temp_slide_img)
         
         if ppt.export_slide_as_image(slide, temp_slide_img):
             # Parse Curve from Slide Screenshot
             try:
                 log(f"Analyzing curve for {key} from Slide {slide.SlideIndex}...")
                 parser = CurveParser(temp_slide_img)
                 val = parser.parse()
                 
                 if val is not None:
                     # Format value
                     options = action.get("options", "")
                     formatted_val = str(val)
                     
                     # Simple format handling (e.g. ".2f")
                     fmt = ".2f" # Default
                     if "." in options:
                         match = re.search(r'\.(\d+)f', options)
                         if match:
                             fmt = f".{match.group(1)}f"
                     
                     try:
                         formatted_val = format(val, fmt)
                     except:
                         formatted_val = f"{val:.2f}"
                         
                     # Replace in Text
                     original_text = action["original_text"]
                     
                     def replace_match(m):
                         matched_key = m.group(1).strip()
                         if matched_key == key:
                             return formatted_val
                         return m.group(0)
                         
                     new_text = re.sub(r'\{\{CURVE:([^}|]+)(?:\|([^}]+))?\}\}', replace_match, original_text)
                     
                     ppt.replace_text(action["shape"], new_text)
                     log(f"SUCCESS: Updated curve value for {key} -> {formatted_val}")
                 else:
                     log(f"FAILURE: Could not extract value for {key}. Check if chart is visible and axis is clear.")
             except Exception as e:
                 log(f"ERROR parsing curve for {key}: {e}")
                 import traceback
                 traceback.print_exc()
         else:
             log(f"Failed to export slide {slide.SlideIndex} for curve analysis.")

    # 4. Save
    log("Saving report...")
    ppt.save_and_close()
    
    # 4. Cleanup
    for img_path in temp_images:
        if os.path.exists(img_path):
            try: os.remove(img_path)
            except: pass
    
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
