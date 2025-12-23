import os
import cv2
import win32com.client
import argparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

VIDEO_MAPPING = {
    "wendu": "wendu.mp4",
    "juanqi": "juanqi.mp4",
    "tijuanqi": "tijijuanqi.mp4",
    "qipao": "qipao.mp4",
    "yanghuazha": "yanghuazha.mp4",
    "tijiyanghua": "tijiyanghua.mp4",
    "liaoliuzhuizong": "liaoliuzhuizong.mp4",
    "suokong": "suokong.mp4",
    "suokonglv": "suokonglv.mp4",
    "suokongtiji": "suokongtiji.mp4",
    "rejie": "rejie.mp4",
}


def log(message):
    print(message, flush=True)


def extract_last_frame(video_path, output_image_path):
    if not os.path.exists(video_path):
        log(f"Warning: Video not found at {video_path}")
        return False

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        log(f"Warning: Could not open video {video_path}")
        return False

    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if frame_count == 0:
        log(f"Warning: Video {video_path} has 0 frames")
        cap.release()
        return False

    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_count - 1)
    ret, frame = cap.read()

    if ret:
        is_success, im_buf_arr = cv2.imencode(".png", frame)
        if is_success:
            im_buf_arr.tofile(output_image_path)
            log(f"Extracted last frame to {output_image_path}")
            cap.release()
            return True
        log(f"Warning: Could not encode frame for {video_path}")
    else:
        log(f"Warning: Could not read last frame of {video_path}")

    cap.release()
    return False


def _split_rect(left, top, width, height, gap):
    try:
        gap = float(gap)
    except Exception:
        gap = 0.0

    if gap < 0:
        gap = 0.0

    if width <= gap:
        gap = 0.0

    half_width = (width - gap) / 2.0
    a = (left, top, half_width, height)
    b = (left + half_width + gap, top, half_width, height)
    return a, b


def generate_compare_report(video_dir_a, video_dir_b, template_path, output_path, gap):
    log("Starting compare report generation...")
    log(f"Video Directory A: {video_dir_a}")
    log(f"Video Directory B: {video_dir_b}")
    log(f"Template Path: {template_path}")
    log(f"Output Path: {output_path}")

    if not os.path.exists(video_dir_a):
        log(f"Error: Video directory A does not exist: {video_dir_a}")
        return 2
    if not os.path.exists(video_dir_b):
        log(f"Error: Video directory B does not exist: {video_dir_b}")
        return 2
    if not os.path.exists(template_path):
        log(f"Error: Template file does not exist: {template_path}")
        return 2

    log("Step 1: Extracting frames...")
    temp_dir = BASE_DIR
    temp_images_a = {}
    temp_images_b = {}

    for key, filename in VIDEO_MAPPING.items():
        video_a = os.path.join(video_dir_a, filename)
        out_a = os.path.join(temp_dir, f"temp_{key}_a.png")
        if os.path.exists(video_a) and extract_last_frame(video_a, out_a):
            temp_images_a[key] = out_a
        else:
            if os.path.exists(out_a):
                try:
                    os.remove(out_a)
                except Exception:
                    pass

        video_b = os.path.join(video_dir_b, filename)
        out_b = os.path.join(temp_dir, f"temp_{key}_b.png")
        if os.path.exists(video_b) and extract_last_frame(video_b, out_b):
            temp_images_b[key] = out_b
        else:
            if os.path.exists(out_b):
                try:
                    os.remove(out_b)
                except Exception:
                    pass

    log("Step 2: Opening PPT...")
    try:
        try:
            ppt_app = win32com.client.Dispatch("Kwpp.Application")
            log("Using WPS Presentation.")
        except Exception:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            log("Using Microsoft PowerPoint.")
    except Exception:
        log("Error: Could not open WPS or PowerPoint.")
        return 3

    try:
        ppt_app.Visible = True
    except Exception:
        pass

    try:
        pres = ppt_app.Presentations.Open(template_path)
    except Exception as e:
        log(f"Error opening template: {e}")
        return 4

    log("Step 3: Processing slides...")
    actions = []

    try:
        for i, slide in enumerate(pres.Slides):
            for shape in slide.Shapes:
                name = shape.Name
                if name.startswith("VID_"):
                    key = name.replace("VID_", "")
                    if key in VIDEO_MAPPING:
                        actions.append(
                            {
                                "slide_index": i + 1,
                                "shape_name": name,
                                "type": "video",
                                "key": key,
                                "left": shape.Left,
                                "top": shape.Top,
                                "width": shape.Width,
                                "height": shape.Height,
                            }
                        )
                elif name.startswith("IMG_"):
                    key = name.replace("IMG_", "")
                    if key in VIDEO_MAPPING and (key in temp_images_a or key in temp_images_b):
                        actions.append(
                            {
                                "slide_index": i + 1,
                                "shape_name": name,
                                "type": "image",
                                "key": key,
                                "left": shape.Left,
                                "top": shape.Top,
                                "width": shape.Width,
                                "height": shape.Height,
                            }
                        )
    except Exception as e:
        log(f"Error iterating slides: {e}")

    log(f"Found {len(actions)} anchors to process.")

    for action in actions:
        key = action["key"]
        try:
            slide = pres.Slides(action["slide_index"])

            try:
                slide.Shapes(action["shape_name"]).Delete()
            except Exception as e:
                log(f"Warning: Could not delete shape {action['shape_name']}: {e}")

            (a_left, a_top, a_width, a_height), (b_left, b_top, b_width, b_height) = _split_rect(
                action["left"], action["top"], action["width"], action["height"], gap
            )

            if action["type"] == "video":
                video_a = os.path.join(video_dir_a, VIDEO_MAPPING[key])
                if os.path.exists(video_a):
                    log(f"Inserting A video {video_a} at Slide {action['slide_index']}")
                    try:
                        s = slide.Shapes.AddMediaObject(video_a, a_left, a_top, a_width, a_height)
                    except Exception:
                        s = slide.Shapes.AddMediaObject2(video_a, False, True, a_left, a_top, a_width, a_height)
                    s.Left, s.Top, s.Width, s.Height = a_left, a_top, a_width, a_height
                else:
                    log(f"Warning: Missing A video for {key}: {video_a}")

                video_b = os.path.join(video_dir_b, VIDEO_MAPPING[key])
                if os.path.exists(video_b):
                    log(f"Inserting B video {video_b} at Slide {action['slide_index']}")
                    try:
                        s = slide.Shapes.AddMediaObject2(video_b, False, True, b_left, b_top, b_width, b_height)
                    except Exception:
                        s = slide.Shapes.AddMediaObject(video_b, b_left, b_top, b_width, b_height)
                    s.Left, s.Top, s.Width, s.Height = b_left, b_top, b_width, b_height
                else:
                    log(f"Warning: Missing B video for {key}: {video_b}")

            elif action["type"] == "image":
                img_a = temp_images_a.get(key)
                if img_a and os.path.exists(img_a):
                    log(f"Inserting A image {img_a} at Slide {action['slide_index']}")
                    slide.Shapes.AddPicture(img_a, False, True, a_left, a_top, a_width, a_height)
                else:
                    log(f"Warning: Missing A image for {key}")

                img_b = temp_images_b.get(key)
                if img_b and os.path.exists(img_b):
                    log(f"Inserting B image {img_b} at Slide {action['slide_index']}")
                    slide.Shapes.AddPicture(img_b, False, True, b_left, b_top, b_width, b_height)
                else:
                    log(f"Warning: Missing B image for {key}")
        except Exception as e:
            log(f"Error inserting compare content for {key}: {e}")

    log("Step 4: Saving report...")
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception as e:
            log(f"Warning: Could not remove existing output file: {e}")

    try:
        pres.SaveAs(output_path)
        log(f"Saved to {output_path}")
    except Exception as e:
        log(f"Error saving: {e}")
        try:
            pres.Close()
        except Exception:
            pass
        try:
            ppt_app.Quit()
        except Exception:
            pass
        return 5

    try:
        pres.Close()
        ppt_app.Quit()
    except Exception:
        pass

    log("Step 5: Cleaning up...")
    for p in list(temp_images_a.values()) + list(temp_images_b.values()):
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass

    log("Done!")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Compare Automation Report (Side-by-side)")
    parser.add_argument("--video_dir_a", required=True, help="Path to first simulation video directory")
    parser.add_argument("--video_dir_b", required=True, help="Path to second simulation video directory")
    parser.add_argument("--template_path", default=os.path.join(BASE_DIR, "自动化模板.pptx"), help="Path to PPT template")
    parser.add_argument("--output_path", default=os.path.join(BASE_DIR, "自动化报告_compare_generated.pptx"), help="Path to output PPT")
    parser.add_argument("--gap", default=10, type=float, help="Gap between A and B in points")
    args = parser.parse_args()

    raise SystemExit(
        generate_compare_report(
            video_dir_a=args.video_dir_a,
            video_dir_b=args.video_dir_b,
            template_path=args.template_path,
            output_path=args.output_path,
            gap=args.gap,
        )
    )

