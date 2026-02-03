import os
import cv2
import numpy as np
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
from utils import run_tesseract_custom, extract_float, load_config

# Increase max pixels to avoid DOS errors on large images
Image.MAX_IMAGE_PIXELS = None

class DiagnoseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Curve Recognition Debugger")
        self.root.geometry("1200x800")
        
        self.image_path = None
        self.original_img = None
        self.display_img = None
        self.scale_factor = 1.0
        
        # Parameters
        self.crop_ratio_var = tk.DoubleVar(value=0.5)
        self.axis_width_ratio_var = tk.DoubleVar(value=0.2)
        self.positive_only = False
        try:
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
            config = load_config(config_path)
            self.positive_only = config.get("ocr", {}).get("positive_only", False)
        except Exception:
            self.positive_only = False
        
        self._setup_ui()
        
        # Check if argument provided
        if len(sys.argv) > 1:
            self.load_image(sys.argv[1])
        else:
            # Try default test image
            default_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test.png")
            if os.path.exists(default_img):
                self.load_image(default_img)

    def _setup_ui(self):
        # Top Control Panel
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X)
        
        ttk.Button(control_frame, text="Load Image", command=self.browse_image).pack(side=tk.LEFT, padx=5)
        
        # Sliders
        param_frame = ttk.LabelFrame(control_frame, text="Parameters", padding=5)
        param_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # Crop Start X
        ttk.Label(param_frame, text="Crop Start X (Right Half) %:").grid(row=0, column=0, padx=5)
        self.crop_scale = ttk.Scale(param_frame, from_=0.0, to=0.9, variable=self.crop_ratio_var, orient=tk.HORIZONTAL, length=200, command=self.on_scale_change)
        self.crop_scale.grid(row=0, column=1, padx=5)
        
        # Entry for Crop
        self.crop_entry = ttk.Entry(param_frame, width=6)
        self.crop_entry.grid(row=0, column=2, padx=5)
        self.crop_entry.insert(0, "50.0")
        self.crop_entry.bind('<Return>', self.on_entry_change)
        self.crop_entry.bind('<FocusOut>', self.on_entry_change)
        
        # Axis Width
        ttk.Label(param_frame, text="Axis Zone Width %:").grid(row=0, column=3, padx=5)
        self.axis_scale = ttk.Scale(param_frame, from_=0.05, to=0.5, variable=self.axis_width_ratio_var, orient=tk.HORIZONTAL, length=200, command=self.on_scale_change)
        self.axis_scale.grid(row=0, column=4, padx=5)
        
        # Entry for Axis
        self.axis_entry = ttk.Entry(param_frame, width=6)
        self.axis_entry.grid(row=0, column=5, padx=5)
        self.axis_entry.insert(0, "20.0")
        self.axis_entry.bind('<Return>', self.on_entry_change)
        self.axis_entry.bind('<FocusOut>', self.on_entry_change)
        
        ttk.Button(control_frame, text="Run Diagnosis", command=self.run_diagnosis).pack(side=tk.LEFT, padx=10)
        
        # Main Content
        content_frame = ttk.Frame(self.root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Image Canvas (Scrollable)
        self.canvas_frame = ttk.Frame(content_frame, borderwidth=2, relief="sunken")
        self.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Log Area
        self.log_text = tk.Text(content_frame, width=40, state="disabled")
        self.log_text.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        print(msg)

    def browse_image(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")])
        if path:
            self.load_image(path)

    def load_image(self, path):
        self.image_path = path
        try:
            # Use imdecode for non-ascii paths
            self.original_img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if self.original_img is None:
                raise Exception("Failed to decode image")
            
            self.log(f"Loaded: {path}")
            self.display_preview(self.original_img)
            self.run_diagnosis()
        except Exception as e:
            messagebox.showerror("Error", f"Could not load image: {e}")

    def on_scale_change(self, event=None):
        # Update entry boxes from slider values
        c = self.crop_ratio_var.get()
        a = self.axis_width_ratio_var.get()
        
        # Only update if focus is not on entry to avoid fighting
        if self.root.focus_get() != self.crop_entry:
            self.crop_entry.delete(0, tk.END)
            self.crop_entry.insert(0, f"{c*100:.1f}")
            
        if self.root.focus_get() != self.axis_entry:
            self.axis_entry.delete(0, tk.END)
            self.axis_entry.insert(0, f"{a*100:.1f}")
        
        # Redraw
        if self.original_img is not None:
            self.draw_overlays_only()

    def on_entry_change(self, event=None):
        try:
            # Get values from entries
            c_str = self.crop_entry.get()
            a_str = self.axis_entry.get()
            
            c = float(c_str) / 100.0
            a = float(a_str) / 100.0
            
            # Clamp values
            c = max(0.0, min(0.9, c))
            a = max(0.05, min(0.5, a))
            
            # Update variables (will update sliders)
            self.crop_ratio_var.set(c)
            self.axis_width_ratio_var.set(a)
            
            # Redraw
            if self.original_img is not None:
                self.draw_overlays_only()
                
        except ValueError:
            pass # Ignore invalid input while typing

    def draw_overlays_only(self):
        if self.original_img is None: return
        
        debug_img = self.original_img.copy()
        h, w = debug_img.shape[:2]
        
        crop_ratio = self.crop_ratio_var.get()
        axis_ratio = self.axis_width_ratio_var.get()
        
        crop_x = int(w * crop_ratio)
        roi_w = w - crop_x
        axis_w = int(roi_w * axis_ratio)
        axis_end_x = crop_x + axis_w
        
        # Draw Crop Line (Yellow)
        cv2.line(debug_img, (crop_x, 0), (crop_x, h), (0, 255, 255), 2)
        cv2.putText(debug_img, "Crop Start", (crop_x + 5, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 255), 2)
        
        # Draw Axis Zone (Cyan)
        cv2.rectangle(debug_img, (crop_x, 0), (axis_end_x, h), (255, 255, 0), 2)
        cv2.putText(debug_img, "Axis Zone", (crop_x + 5, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 0), 2)
        
        self.display_preview(debug_img)

    def display_preview(self, img):
        # Resize to fit canvas
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        if canvas_w < 10 or canvas_h < 10:
            canvas_w = 800
            canvas_h = 600
            
        h, w = img.shape[:2]
        ratio = min(canvas_w/w, canvas_h/h)
        self.scale_factor = ratio
        
        new_w = int(w * ratio)
        new_h = int(h * ratio)
        
        resized = cv2.resize(img, (new_w, new_h))
        # Convert BGR to RGB
        rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        self.tk_img = ImageTk.PhotoImage(pil_img)
        
        self.canvas.delete("all")
        self.canvas.create_image(canvas_w//2, canvas_h//2, anchor=tk.CENTER, image=self.tk_img)

    def run_diagnosis(self):
        if self.original_img is None: return
        
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")
        
        self.log("Running diagnosis...")
        
        img = self.original_img.copy()
        h, w = img.shape[:2]
        debug_img = img.copy()
        
        crop_ratio = self.crop_ratio_var.get()
        axis_ratio = self.axis_width_ratio_var.get()
        
        # 1. Crop
        crop_x_start = int(w * crop_ratio)
        roi = img[:, crop_x_start:]
        roi_h, roi_w = roi.shape[:2]
        
        self.log(f"Crop: Start X={crop_x_start} ({crop_ratio:.1%}), Width={roi_w}")
        
        # 2. Axis OCR
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        axis_width = int(roi_w * axis_ratio)
        axis_roi = gray[:, :axis_width]
        
        self.log(f"Axis Search Width: {axis_width} ({axis_ratio:.1%})")
        
        # Advanced Preprocessing for OCR (Fix 0.x -> 6.0 issue)
        scale = 4
        roi_large = cv2.resize(axis_roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
        
        # 1. Gaussian Blur to reduce noise
        roi_blur = cv2.GaussianBlur(roi_large, (3, 3), 0)
        
        # 2. Adaptive Threshold (better for small details like decimals)
        roi_thresh = cv2.adaptiveThreshold(roi_blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                         cv2.THRESH_BINARY, 31, 10)
        
        # 3. Erosion (Thicken text/dots on white background)
        kernel = np.ones((2,2), np.uint8)
        roi_processed = cv2.erode(roi_thresh, kernel, iterations=1)
        
        # Save debug image
        debug_ocr_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_ocr_last_run.png")
        cv2.imwrite(debug_ocr_path, roi_processed)
        self.log(f"Saved OCR debug image to: {debug_ocr_path}")
        
        data = run_tesseract_custom(roi_processed)
        
        y_values = []
        if data:
            num_boxes = len(data['text'])
            for i in range(num_boxes):
                text = data['text'][i].strip()
                if not text: continue
                
                nums = extract_float(text)
                if nums:
                    val = nums[0]
                    # Relative to ROI
                    x_rel = int(data['left'][i] / scale)
                    y_rel = int(data['top'][i] / scale)
                    w_box = int(data['width'][i] / scale)
                    h_box = int(data['height'][i] / scale)
                    
                    # Global
                    x = crop_x_start + x_rel
                    y = y_rel
                    
                    y_center = y + h_box / 2
                    x_center = x + w_box / 2
                    
                    y_values.append({
                        'val': val,
                        'y_center': y_center,
                        'x_center': x_center,
                        'rect': (x, y, w_box, h_box)
                    })
                    
                    # Draw rect (Initially yellow for candidates)
                    cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), (0, 255, 255), 1)
                    cv2.putText(debug_img, str(val), (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1)

        candidate_bins = []
        if len(y_values) >= 2:
            bin_width = 20
            bins = {}
            for v in y_values:
                bin_idx = int(v['x_center'] / bin_width)
                if bin_idx not in bins:
                    bins[bin_idx] = []
                bins[bin_idx].append(v)
            candidate_bins = sorted(bins.items(), key=lambda x: (-len(x[1]), x[0]))
        else:
            self.log("FAILURE: Not enough Y-axis labels.")
            self.display_preview(debug_img)
            return
        
        # 3. Curve
        chart_start_x = crop_x_start + axis_width
        
        # Draw separation line
        cv2.line(debug_img, (chart_start_x, 0), (chart_start_x, h), (0, 0, 255), 2)
        
        chart_roi = gray[:, axis_width:]
        _, binary = cv2.threshold(chart_roi, 200, 255, cv2.THRESH_BINARY_INV)
        
        # Remove grid
        kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
        detected_lines_h = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_h)
        # Increased kernel height for better vertical line removal
        kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))
        detected_lines_v = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_v)
        clean_binary = cv2.subtract(binary, detected_lines_h)
        clean_binary = cv2.subtract(clean_binary, detected_lines_v)
        
        # --- Mask Top-Right Corner (Legend/Label Area) ---
        h_chart, w_chart = clean_binary.shape[:2]
        mask_h = int(h_chart * 0.08)
        mask_w = int(w_chart * 0.45)
        # Draw mask on debug image for visualization (Gray box)
        mask_x_global = chart_start_x + (w_chart - mask_w)
        cv2.rectangle(debug_img, (mask_x_global, 0), (chart_start_x + w_chart, mask_h), (100, 100, 100), -1)
        cv2.putText(debug_img, "Masked", (mask_x_global, mask_h//2), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Apply mask to binary
        clean_binary[0:mask_h, (w_chart - mask_w):] = 0
        
        # --- Mask Bottom (X-Axis/Border Area) ---
        mask_bottom_h = int(h_chart * 0.12)
        # Draw mask on debug image (Gray box)
        cv2.rectangle(debug_img, (chart_start_x, h_chart - mask_bottom_h), (chart_start_x + w_chart, h_chart), (100, 100, 100), -1)
        cv2.putText(debug_img, "Masked Bottom", (chart_start_x + 10, h_chart - 5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Apply mask
        clean_binary[(h_chart - mask_bottom_h):, :] = 0
        # -------------------------------------------------

        # --- Filter Vertical Lines by Aspect Ratio (Connected Components) ---
        num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(clean_binary, connectivity=8)
        
        for i in range(1, num_labels): # Skip background
            x, y, w, h_comp, area = stats[i]
            aspect_ratio = h_comp / float(w) if w > 0 else 0
            
            # Draw all components in faint gray for debug
            # x_global = chart_start_x + x
            # cv2.rectangle(debug_img, (x_global, y), (x_global + w, y + h_comp), (50, 50, 50), 1)

            if aspect_ratio > 10 and h_comp > (h_chart * 0.2):
                 clean_binary[labels == i] = 0
                 # Visualize removed line in BLUE
                 x_global = chart_start_x + x
                 cv2.rectangle(debug_img, (x_global, y), (x_global + w, y + h_comp), (255, 0, 0), 2)
                 cv2.putText(debug_img, f"RM AR={aspect_ratio:.1f}", (x_global, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 0), 1)
                 self.log(f"Removed vertical artifact: AR={aspect_ratio:.1f}, H={h_comp}")
        # ------------------------------------------------------------------

        points = cv2.findNonZero(clean_binary)
        if points is None:
            self.log("FAILURE: No curve detected.")
            self.display_preview(debug_img)
            return
            
        points = points[:, 0, :]
        points[:, 0] += chart_start_x
        
        # Rightmost
        sorted_points = points[points[:, 0].argsort()[::-1]]
        
        # User Request: Use the single last point (rightmost)
        rightmost_point = sorted_points[0]
        
        # Draw the chosen point
        cv2.circle(debug_img, tuple(rightmost_point), 5, (0, 0, 255), -1) # Red dot for final point
            
        y_end_px = rightmost_point[1]
        cv2.line(debug_img, (0, int(y_end_px)), (w, int(y_end_px)), (255, 0, 255), 1)
        
        if not candidate_bins:
            self.log("FAILURE: Not enough Y-axis labels.")
            self.display_preview(debug_img)
            return
        
        chosen = None
        for bin_idx, items in candidate_bins:
            selected_x = [v['x_center'] for v in items]
            median_x = np.median(selected_x)
            aligned = [v for v in y_values if abs(v['x_center'] - median_x) < 15]
            if self.positive_only:
                aligned = [v for v in aligned if v['val'] >= 0]
            if len(aligned) < 2:
                continue
            aligned.sort(key=lambda v: v['y_center'])
            y_top_px = aligned[0]['y_center']
            val_top = aligned[0]['val']
            y_bottom_px = aligned[-1]['y_center']
            val_bottom = aligned[-1]['val']
            if (y_bottom_px - y_top_px) == 0:
                continue
            slope = (val_bottom - val_top) / (y_bottom_px - y_top_px)
            result = val_top + slope * (y_end_px - y_top_px)
            if self.positive_only and val_top >= 0 and val_bottom >= 0 and result < 0:
                self.log(f"Positive-only enabled; negative result {result} from bin {bin_idx}, retrying")
                continue
            chosen = (aligned, median_x, result, y_top_px, val_top, y_bottom_px, val_bottom)
            break
        
        if not chosen:
            self.log("FAILURE: No valid positive mapping.")
            self.display_preview(debug_img)
            return
        
        aligned, median_x, result, y_top_px, val_top, y_bottom_px, val_bottom = chosen
        
        for v in y_values:
            if abs(v['x_center'] - median_x) < 15 and (not self.positive_only or v['val'] >= 0):
                x, y, w_box, h_box = v['rect']
                cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), (0, 255, 0), 2)
                cv2.putText(debug_img, str(v['val']), (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
            else:
                x, y, w_box, h_box = v['rect']
                cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), (0, 0, 255), 1)
        
        self.log(f"Found {len(aligned)} numbers (after filtering).")
        self.log(f"Top: {val_top} @ {y_top_px:.1f}px")
        self.log(f"Bottom: {val_bottom} @ {y_bottom_px:.1f}px")
        
        self.log(f"Curve End Y: {y_end_px:.1f}")
        self.log(f"RESULT: {result:.4e}")
        
        cv2.putText(debug_img, f"Result: {result:.4e}", (int(w/2), 50), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2)
        
        self.display_preview(debug_img)

if __name__ == "__main__":
    root = tk.Tk()
    app = DiagnoseApp(root)
    root.mainloop()
