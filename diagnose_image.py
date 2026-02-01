import os
import cv2
import numpy as np
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
from utils import run_tesseract_custom, extract_float

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
        
        scale = 2
        roi_large = cv2.resize(axis_roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
        data = run_tesseract_custom(roi_large)
        
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

        # Filter: Separation of Y-axis and X-axis
        if len(y_values) >= 2:
            # Cluster by X-coordinate (20px bins)
            bin_width = 20
            bins = {}
            for v in y_values:
                bin_idx = int(v['x_center'] / bin_width)
                if bin_idx not in bins: bins[bin_idx] = []
                bins[bin_idx].append(v)
            
            # Find best bin (most items, tie-break left)
            best_bin = None
            max_count = 0
            for bin_idx, items in bins.items():
                if len(items) > max_count:
                    max_count = len(items)
                    best_bin = items
                elif len(items) == max_count:
                    if best_bin and bin_idx < int(best_bin[0]['x_center'] / bin_width):
                        best_bin = items
            
            aligned_y_values = []
            if best_bin:
                selected_x = [v['x_center'] for v in best_bin]
                median_x = np.median(selected_x)
                
                # Strict filter +/- 15px
                for v in y_values:
                    if abs(v['x_center'] - median_x) < 15:
                        aligned_y_values.append(v)
                        # Draw accepted in GREEN
                        x, y, w_box, h_box = v['rect']
                        cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), (0, 255, 0), 2)
                        cv2.putText(debug_img, str(v['val']), (x, y-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
                    else:
                        # Draw rejected in RED
                        x, y, w_box, h_box = v['rect']
                        cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), (0, 0, 255), 1)
            
            if len(aligned_y_values) >= 2:
                y_values = aligned_y_values
            else:
                y_values = [] # Fail if no valid column found

        self.log(f"Found {len(y_values)} numbers (after filtering).")
        
        if len(y_values) < 2:
            self.log("FAILURE: Not enough Y-axis labels.")
            self.display_preview(debug_img)
            return
            
        y_values.sort(key=lambda v: v['y_center'])
        y_top_px = y_values[0]['y_center']
        val_top = y_values[0]['val']
        y_bottom_px = y_values[-1]['y_center']
        val_bottom = y_values[-1]['val']
        
        self.log(f"Top: {val_top} @ {y_top_px:.1f}px")
        self.log(f"Bottom: {val_bottom} @ {y_bottom_px:.1f}px")
        
        # 3. Curve
        chart_start_x = crop_x_start + axis_width
        
        # Draw separation line
        cv2.line(debug_img, (chart_start_x, 0), (chart_start_x, h), (0, 0, 255), 2)
        
        chart_roi = gray[:, axis_width:]
        _, binary = cv2.threshold(chart_roi, 200, 255, cv2.THRESH_BINARY_INV)
        
        # Remove grid
        kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 1))
        detected_lines_h = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_h)
        kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 20))
        detected_lines_v = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_v)
        clean_binary = cv2.subtract(binary, detected_lines_h)
        clean_binary = cv2.subtract(clean_binary, detected_lines_v)
        
        points = cv2.findNonZero(clean_binary)
        if points is None:
            self.log("FAILURE: No curve detected.")
            self.display_preview(debug_img)
            return
            
        points = points[:, 0, :]
        points[:, 0] += chart_start_x
        
        # Rightmost
        sorted_points = points[points[:, 0].argsort()[::-1]]
        top_n = min(20, len(sorted_points))
        rightmost = sorted_points[:top_n]
        
        for p in rightmost:
            cv2.circle(debug_img, tuple(p), 3, (255, 0, 255), -1)
            
        y_end_px = np.mean(rightmost[:, 1])
        cv2.line(debug_img, (0, int(y_end_px)), (w, int(y_end_px)), (255, 0, 255), 1)
        
        # Correct Linear Interpolation Logic
        # slope = (val_bottom - val_top) / (y_bottom_px - y_top_px)
        # result = val_top + slope * (y_current - y_top_px)
        
        if (y_bottom_px - y_top_px) != 0:
            slope = (val_bottom - val_top) / (y_bottom_px - y_top_px)
            result = val_top + slope * (y_end_px - y_top_px)
        else:
            result = 0.0
        
        self.log(f"Curve End Y: {y_end_px:.1f}")
        self.log(f"RESULT: {result:.4e}")
        
        cv2.putText(debug_img, f"Result: {result:.4e}", (int(w/2), 50), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2)
        
        self.display_preview(debug_img)

if __name__ == "__main__":
    root = tk.Tk()
    app = DiagnoseApp(root)
    root.mainloop()
