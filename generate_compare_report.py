import os
import argparse
import time
from utils import log, extract_last_frame, load_config, find_media_file
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_compare_report(video_dir_a, video_dir_b, template_path, output_path, gap):
    log("Starting compare report generation...")
    
    config = load_config(CONFIG_PATH)
    labels = config.get("labels", {})

    if not os.path.exists(video_dir_a) or not os.path.exists(video_dir_b):
        log("Error: One of the video directories does not exist.")
        return 2

    # 1. PPT Automation
    ppt = PPTAutomation(template_path, output_path, config_labels=labels)
    if not ppt.open():
        return 3

    # 2. Scan and Process
    log("Scanning slides for placeholders...")
    actions = ppt.scan_placeholders()
    log(f"Found {len(actions)} anchors to process.")
    
    temp_images = []
    
    for action in actions:
        key = action["key"]
        slide = action["slide"]
        shape = action["shape"]
        is_video_placeholder = (action["type"] == "video")
        
        # Find media files
        path_a = find_media_file(video_dir_a, key)
        path_b = find_media_file(video_dir_b, key)
        
        final_a = path_a
        final_b = path_b
        
        # Handle Image Placeholders (Extract frames if video found)
        if not is_video_placeholder:
            # Process A
            if path_a and path_a.lower().endswith(('.mp4', '.avi', '.mov', '.wmv', '.mkv')):
                temp_a = os.path.join(BASE_DIR, f"temp_{key}_a_{int(time.time())}.png")
                if extract_last_frame(path_a, temp_a):
                    final_a = temp_a
                    temp_images.append(temp_a)
                else:
                    final_a = None # Failed to extract
            
            # Process B
            if path_b and path_b.lower().endswith(('.mp4', '.avi', '.mov', '.wmv', '.mkv')):
                temp_b = os.path.join(BASE_DIR, f"temp_{key}_b_{int(time.time())}.png")
                if extract_last_frame(path_b, temp_b):
                    final_b = temp_b
                    temp_images.append(temp_b)
                else:
                    final_b = None
            
        ppt.insert_comparison(slide, shape, final_a, final_b, gap, is_video=is_video_placeholder)

    # 3. Save
    log("Saving report...")
    ppt.save_and_close()
    
    # 4. Cleanup
    for p in temp_images:
        if p and os.path.exists(p):
            try: os.remove(p)
            except: pass
            
    log("Done!")
    return 0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Compare Automation Report (Side-by-side)")
    parser.add_argument("--video_dir_a", required=True, help="Path to first simulation video directory")
    parser.add_argument("--video_dir_b", required=True, help="Path to second simulation video directory")
    parser.add_argument("--template_path", default=os.path.join(BASE_DIR, "自动化模板.pptx"), help="Path to PPT template")
    parser.add_argument("--output_path", help="Path to output PPT")
    parser.add_argument("--gap", default=10, type=float, help="Gap between A and B in points")
    args = parser.parse_args()

    output_path = args.output_path
    if not output_path:
        video_folder_name = os.path.basename(os.path.normpath(args.video_dir_a))
        current_date = time.strftime("%Y.%m.%d")
        output_path = os.path.join(BASE_DIR, f"{video_folder_name}-模流分析报告-{current_date}.pptx")

    raise SystemExit(
        generate_compare_report(
            video_dir_a=args.video_dir_a,
            video_dir_b=args.video_dir_b,
            template_path=args.template_path,
            output_path=output_path,
            gap=args.gap,
        )
    )
