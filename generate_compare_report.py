import os
import argparse
import time
from utils import log, extract_last_frame, load_config
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_compare_report(video_dir_a, video_dir_b, template_path, output_path, gap):
    log("Starting compare report generation...")
    
    config = load_config(CONFIG_PATH)
    video_mapping = config.get("video_mapping", {})
    labels = config.get("labels", {})

    if not os.path.exists(video_dir_a) or not os.path.exists(video_dir_b):
        log("Error: One of the video directories does not exist.")
        return 2

    # 1. Prepare Images
    log("Step 1: Extracting frames...")
    temp_dir = BASE_DIR
    temp_images_a = {}
    temp_images_b = {}
    
    for key, info in video_mapping.items():
        filename = info["file"] if isinstance(info, dict) else info
        
        # A
        video_a = os.path.join(video_dir_a, filename)
        out_a = os.path.join(temp_dir, f"temp_{key}_a.png")
        if os.path.exists(video_a) and extract_last_frame(video_a, out_a):
            temp_images_a[key] = out_a
            
        # B
        video_b = os.path.join(video_dir_b, filename)
        out_b = os.path.join(temp_dir, f"temp_{key}_b.png")
        if os.path.exists(video_b) and extract_last_frame(video_b, out_b):
            temp_images_b[key] = out_b

    # 2. PPT Automation
    ppt = PPTAutomation(template_path, output_path, config_labels=labels)
    if not ppt.open():
        return 3

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
            media_path_a = os.path.join(video_dir_a, filename)
            media_path_b = os.path.join(video_dir_b, filename)
        else:
            media_path_a = temp_images_a.get(key)
            media_path_b = temp_images_b.get(key)
            
        ppt.insert_comparison(slide, shape_name, media_path_a, media_path_b, gap, is_video=is_video)

    # 4. Save
    log("Step 3: Saving report...")
    ppt.save_and_close()
    
    # 5. Cleanup
    for p in list(temp_images_a.values()) + list(temp_images_b.values()):
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
