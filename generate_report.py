import os
import argparse
import time
from utils import log, extract_last_frame, load_config, find_media_file, detect_curve_end_value
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

    # 2. Scan and Process
    log("Scanning slides for placeholders...")
    actions = ppt.scan_placeholders()
    log(f"Found {len(actions)} anchors to process.")
    
    temp_images = []
    processed_keys = set()
    
    for action in actions:
        key = action["key"]
        is_video_placeholder = (action["type"] == "video")
        
        # Find media file
        media_path = find_media_file(video_dir, key)
        
        if not media_path:
            log(f"Skipping {key}: Media not found")
            continue
            
        is_video_file = media_path.lower().endswith(('.mp4', '.avi', '.mov', '.wmv', '.mkv'))
        final_path = media_path
        
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
        processed_keys.add(key)

    # Curve Recognition Phase
    log("Starting curve value recognition...")
    temp_export_dir = os.path.join(BASE_DIR, "temp_exports")
    if not os.path.exists(temp_export_dir):
        os.makedirs(temp_export_dir)
        
    for key in processed_keys:
        # Optimization: Only export and process if there is a VAL_ placeholder for this key
        if not ppt.check_val_placeholder_exists(key):
            # log(f"Skipping OCR for {key}: No VAL_{key} placeholder found.")
            continue
            
        export_path = os.path.join(temp_export_dir, f"{key}.png")
        if ppt.export_smarttag_image(key, export_path):
            value = detect_curve_end_value(export_path)
            
            # Debug: Save debug image if recognition failed or just for trace
            # (Optional: can be enabled by a flag, but for now let's keep clean)
            
            if value is not None:
                val_str = f"{value:.3e}"
                log(f"Detected value for {key}: {val_str}")
                ppt.replace_text_placeholder(key, val_str)
            else:
                log(f"Could not detect value for {key}")
                # Save failed image for inspection
                try:
                    failed_debug_path = os.path.join(BASE_DIR, f"debug_failed_{key}.png")
                    import shutil
                    shutil.copy(export_path, failed_debug_path)
                    log(f"Saved failed image to {failed_debug_path}")
                except: pass
            
            if os.path.exists(export_path):
                try: os.remove(export_path)
                except: pass
                
    try: os.rmdir(temp_export_dir)
    except: pass

    # 3. Save
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
