import os
import argparse
import time
from utils import log, extract_last_frame, load_config
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_append_report(existing_ppt_path, video_dir, output_path, gap):
    log("Starting append report generation...")
    
    config = load_config(CONFIG_PATH)
    video_mapping = config.get("video_mapping", {})
    labels = config.get("labels", {})

    if not os.path.exists(existing_ppt_path):
        log(f"Error: Existing PPT does not exist: {existing_ppt_path}")
        return 2
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return 2

    # 1. Prepare Images
    log("Step 1: Extracting frames for new videos...")
    temp_dir = BASE_DIR
    temp_images = {}
    
    for key, info in video_mapping.items():
        filename = info["file"]
        video_path = os.path.join(video_dir, filename)
        out_path = os.path.join(temp_dir, f"temp_{key}_new.png")
        if os.path.exists(video_path) and extract_last_frame(video_path, out_path):
            temp_images[key] = out_path

    # 2. PPT Automation
    ppt = PPTAutomation(existing_ppt_path, output_path, config_labels=labels)
    if not ppt.open():
        return 3

    # 3. Process Slides (Based on Index)
    log("Step 2: Processing slides...")
    
    # Process Video Slides
    for key, info in video_mapping.items():
        slide_idx = info.get("slide")
        filename = info.get("file")
        
        if not slide_idx or not filename: continue
        
        video_path = os.path.join(video_dir, filename)
        if not os.path.exists(video_path):
            log(f"Skipping {key}: New video not found.")
            continue
            
        ppt.insert_comparison_with_existing(slide_idx, video_path, gap, is_video=True)

    # Process Image Slides
    for key, info in video_mapping.items():
        slide_idx = info.get("img_slide")
        img_path = temp_images.get(key)
        
        if not slide_idx or not img_path: continue
        
        ppt.insert_comparison_with_existing(slide_idx, img_path, gap, is_video=False)

    # 4. Save
    log("Step 3: Saving report...")
    ppt.save_and_close()
    
    # 5. Cleanup
    for p in temp_images.values():
        if os.path.exists(p):
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
