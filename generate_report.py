import os
import argparse
import time
import json
from utils import log, extract_last_frame, load_config, find_media_file, detect_curve_end_value
from ppt_automation import PPTAutomation
from ai_analysis import AIAnalyzer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
UNIT_CONFIG_PATH = os.path.join(BASE_DIR, "unit_config.json")

def load_unit_config():
    """Load unit conversion configuration from JSON file"""
    if os.path.exists(UNIT_CONFIG_PATH):
        try:
            with open(UNIT_CONFIG_PATH, 'r', encoding='utf-8') as f:
                log(f"Loaded unit config from {UNIT_CONFIG_PATH}")
                return json.load(f)
        except Exception as e:
            log(f"Error loading unit config: {e}")
    else:
        log("No unit_config.json found, using defaults.")
    return {}

def generate_report(video_dir, template_path, output_path):
    log("Starting report generation...")
    
    # Load unit config
    unit_config_map = load_unit_config()

    log(f"Video Directory: {video_dir}")
    
    # config = load_config(CONFIG_PATH) # No longer needed for video_mapping
    main_config = load_config(CONFIG_PATH)
    
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return

    # 1. PPT Automation
    ppt = PPTAutomation(template_path, output_path)
    if not ppt.open():
        return

    # 1.5 Replace Metadata (name, date) on First Slide
    try:
        folder_name = os.path.basename(os.path.normpath(video_dir))
        current_date_str = time.strftime("%Y.%m.%d")
        
        replacements = {
            "name": folder_name,
            "date": current_date_str
        }
        log(f"Replacing metadata on first slide: {replacements}")
        ppt.replace_texts(replacements, only_first_slide=True)
    except Exception as e:
        log(f"Warning: Failed to replace metadata: {e}")

    # 2. Scan and Process
    log("Scanning slides for placeholders...")
    actions = ppt.scan_placeholders()
    log(f"Found {len(actions)} anchors to process.")
    
    temp_images = []
    processed_keys = set()
    key_to_image_path = {} # Map key to local image path for AI analysis
    
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
        
        # Logic update: We always want a static image for AI analysis if it's a video file.
        # So we extract frame for ALL video files found, not just when placeholder is Image.
        if is_video_file:
            temp_img = os.path.join(BASE_DIR, f"temp_{key}_{int(time.time())}.png")
            # Only extract if we haven't already (though keys should be unique usually)
            if extract_last_frame(media_path, temp_img):
                temp_images.append(temp_img)
                key_to_image_path[key] = temp_img
                
                # If placeholder is image type, we MUST use this frame as final_path
                if not is_video_placeholder:
                    final_path = temp_img
            else:
                log(f"Warning: Failed to extract frame for {key} from {media_path}")
                # If we failed to extract and it's an image placeholder, we might have issues.
                # But let's continue with original file if it's video placeholder.
                if not is_video_placeholder:
                    continue
        else:
            # It's an image file
            key_to_image_path[key] = media_path
                
        ppt.process_media_placeholder(action, final_path)
        processed_keys.add(key)

    # Curve Recognition Phase
    log("Starting curve value recognition...")
    temp_export_dir = os.path.join(BASE_DIR, "temp_exports")
    if not os.path.exists(temp_export_dir):
        os.makedirs(temp_export_dir)
        
    positive_only = main_config.get("ocr", {}).get("positive_only", False)
        
    for key in processed_keys:
        # Optimization: Only export and process if there is a VAL_ placeholder for this key
        if not ppt.check_val_placeholder_exists(key):
            # log(f"Skipping OCR for {key}: No VAL_{key} placeholder found.")
            continue
            
        export_path = os.path.join(temp_export_dir, f"{key}.png")
        if ppt.export_smarttag_image(key, export_path):
            value = detect_curve_end_value(export_path, positive_only=positive_only)
            
            # Debug: Save debug image if recognition failed or just for trace
            # (Optional: can be enabled by a flag, but for now let's keep clean)
            
            if value is not None:
                # Check for unit conversion config
                if key in unit_config_map:
                    config = unit_config_map[key]
                    scale = config.get("scale", 1.0)
                    unit = config.get("unit", "")
                    fmt = config.get("format", "{:.3e}")
                    
                    converted_value = value * scale
                    try:
                        val_str = fmt.format(converted_value)
                        if unit:
                            val_str += f" {unit}"
                    except Exception as e:
                        log(f"Format error for {key}: {e}")
                        val_str = f"{converted_value:.3e} {unit}"
                else:
                    # Default behavior
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

    # AI Analysis Phase
    ai_analyzer = AIAnalyzer(main_config)
    ai_analyzer.process_presentation(ppt, temp_export_dir, local_images_map=key_to_image_path)

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
