import os
import argparse
import time
from utils import log, extract_last_frame, load_config
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_report(video_dir, template_path, output_path):
    log("Starting report generation...")
    log(f"Video Directory: {video_dir}")
    
    config = load_config(CONFIG_PATH)
    video_mapping = config.get("video_mapping", {})
    
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return

    # 1. Prepare Images
    log("Step 1: Extracting frames...")
    temp_dir = BASE_DIR
    temp_images = {}
    
    for key, info in video_mapping.items():
        # Handle both simple filename (legacy config style if any) and dict style
        filename = info["file"] if isinstance(info, dict) else info
        video_path = os.path.join(video_dir, filename)
        image_path = os.path.join(temp_dir, f"temp_{key}.png")
        
        if os.path.exists(video_path):
            if extract_last_frame(video_path, image_path):
                temp_images[key] = image_path
        else:
            log(f"Warning: Video file {filename} missing in {video_dir}.")

    # 2. PPT Automation
    ppt = PPTAutomation(template_path, output_path)
    if not ppt.open():
        return

    # 3. Scan and Process
    log("Step 2: Processing slides...")
    actions = ppt.scan_placeholders(video_mapping.keys())
    log(f"Found {len(actions)} anchors to process.")
    
    for action in actions:
        key = action["key"]
        slide = action["slide"]
        shape_name = action["shape_name"]
        is_video = (action["type"] == "video")
        
        info = video_mapping[key]
        filename = info["file"] if isinstance(info, dict) else info
        
        if is_video:
            media_path = os.path.join(video_dir, filename)
        else:
            media_path = temp_images.get(key)
            
        if media_path:
            ppt.insert_video_by_name(slide, shape_name, media_path, is_video=is_video)
        else:
            log(f"Skipping {key}: Media not found")

    # 4. Save
    log("Step 3: Saving report...")
    ppt.save_and_close()
    
    # 5. Cleanup
    for img_path in temp_images.values():
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
