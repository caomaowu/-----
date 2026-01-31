import os
import argparse
import time
from utils import log, extract_last_frame, load_config, find_media_file
from ppt_automation import PPTAutomation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def generate_append_report(existing_ppt_path, video_dir, output_path, gap):
    log("Starting append report generation...")
    
    # config = load_config(CONFIG_PATH)
    # labels = config.get("labels", {}) # We can still use labels if needed, but passed to PPTAutomation init
    
    # To keep consistent with other scripts, let's load config for labels
    config = load_config(CONFIG_PATH)
    labels = config.get("labels", {})

    if not os.path.exists(existing_ppt_path):
        log(f"Error: Existing PPT does not exist: {existing_ppt_path}")
        return 2
    if not os.path.exists(video_dir):
        log(f"Error: Video directory does not exist: {video_dir}")
        return 2

    # 1. PPT Automation
    ppt = PPTAutomation(existing_ppt_path, output_path, config_labels=labels)
    if not ppt.open():
        return 3

    # 2. Scan and Process
    log("Scanning slides for existing SmartTags...")
    # New logic: Scan for SmartTags instead of placeholders
    actions = ppt.scan_smart_tags()
    log(f"Found {len(actions)} existing media objects (SmartTags).")
    
    temp_images = []
    
    for action in actions:
        key = action["key"]
        slide = action["slide"]
        shape = action["shape"]
        
        # Find new media file to compare against
        media_path = find_media_file(video_dir, key)
        
        if not media_path:
            log(f"Skipping {key}: New media not found in video dir")
            continue
            
        is_video_file = media_path.lower().endswith(('.mp4', '.avi', '.mov', '.wmv', '.mkv'))
        final_path = media_path
        
        # Check if we need to extract frame (if existing is image, maybe we want image vs image?)
        # But append mode usually compares video vs new video.
        # Let's assume if new file is video, we insert video.
        # Unless we want to match types?
        # Let's stick to: if new file is video, insert as video.
        
        # However, we need to know if we should extract frame.
        # Usually append report is "Existing vs New".
        # If new is video, use video.
        # If new is video but we want image?
        # Let's assume default behavior: use the file as is.
        # EXCEPT if the existing shape suggests it was an image?
        # Hard to tell from just shape.
        # Let's check if the filename suggests it should be an image (legacy logic checked 'img_slide').
        # But now we don't have config.
        # Let's rely on the file extension found.
        
        is_video_insert = True
        
        if is_video_file:
            # But wait, maybe we want to extract frame if it's an image comparison?
            # In v2.0, we don't know if the user wants image or video just by key.
            # We can check if the key starts with IMG_? No, key is just 'wendu'.
            # We can check the existing shape type?
            # If existing shape is a picture, maybe we want a picture?
            if shape.Type == 13: # msoPicture
                 # It's a picture.
                 # Should we convert new video to picture?
                 # Let's try to be smart.
                 # If we found a video file but existing is picture, let's extract frame.
                 temp_img = os.path.join(BASE_DIR, f"temp_{key}_append_{int(time.time())}.png")
                 if extract_last_frame(media_path, temp_img):
                     final_path = temp_img
                     temp_images.append(final_path)
                     is_video_insert = False
        else:
            is_video_insert = False
        
        # Call new append logic
        ppt.insert_comparison_with_smart_tag(slide, shape, final_path, gap, is_video=is_video_insert)

    # 3. Save
    log("Saving report...")
    ppt.save_and_close()
    
    # 4. Cleanup
    for p in temp_images:
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
