import os
import cv2
import numpy as np
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
from utils import analyze_curve_image, load_config

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
        self.axis_width_ratio_var = tk.DoubleVar(value=0.18)
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
        self.axis_entry.insert(0, "18.0")
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
        
        crop_ratio = self.crop_ratio_var.get()
        axis_ratio = self.axis_width_ratio_var.get()
        result = analyze_curve_image(
            self.image_path,
            positive_only=self.positive_only,
            crop_ratio=crop_ratio,
            axis_ratio=axis_ratio,
        )

        img = self.original_img.copy()
        h, w = img.shape[:2]
        debug_img = img.copy()

        crop_x_start = result.get('crop_x_start', int(w * crop_ratio))
        axis_width = result.get('axis_width', int((w - crop_x_start) * axis_ratio))
        chart_start_x = result.get('chart_start_x', crop_x_start + axis_width)

        cv2.line(debug_img, (crop_x_start, 0), (crop_x_start, h), (0, 255, 255), 2)
        cv2.putText(debug_img, "Crop Start", (crop_x_start + 5, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 255), 2)
        cv2.rectangle(debug_img, (crop_x_start, 0), (crop_x_start + axis_width, h), (255, 255, 0), 2)
        cv2.putText(debug_img, "Axis Zone", (crop_x_start + 5, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 0), 2)
        cv2.line(debug_img, (chart_start_x, 0), (chart_start_x, h), (0, 0, 255), 2)

        self.log(f"Crop: Start X={crop_x_start} ({crop_ratio:.1%})")
        actual_axis_ratio = result.get('axis_ratio', axis_ratio)
        self.log(f"Axis Search Width: {axis_width} ({actual_axis_ratio:.1%})")

        axis_candidates = result.get('axis_candidates', [])
        axis_model = result.get('axis_model')
        inlier_keys = set()
        if axis_model:
            for item in axis_model['inliers']:
                rect = item['rect']
                inlier_keys.add((rect, round(item['val'], 8)))

        for item in axis_candidates:
            x, y, w_box, h_box = item['rect']
            key = (item['rect'], round(item['val'], 8))
            color = (0, 255, 0) if key in inlier_keys else (0, 0, 255)
            thickness = 2 if key in inlier_keys else 1
            label = f"{item['val']} ({item['conf']:.0f})"
            cv2.rectangle(debug_img, (x, y), (x + w_box, y + h_box), color, thickness)
            cv2.putText(debug_img, label, (x, max(15, y - 5)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, color, 1)

        curve = result.get('curve')
        if curve:
            x, y, w_box, h_box = curve['bbox']
            x_global = chart_start_x + x
            cv2.rectangle(debug_img, (x_global, y), (x_global + w_box, y + h_box), (255, 0, 0), 2)
            legend_rect = curve.get('legend_rect')
            if legend_rect:
                lx, ly, lw_box, lh_box = legend_rect
                legend_x_global = chart_start_x + lx
                cv2.rectangle(
                    debug_img,
                    (legend_x_global, ly),
                    (legend_x_global + lw_box, ly + lh_box),
                    (0, 165, 255),
                    2,
                )
                cv2.putText(
                    debug_img,
                    "Legend",
                    (legend_x_global, max(20, ly - 8)),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.5,
                    (0, 165, 255),
                    1,
                )
            endpoint = (chart_start_x + int(round(curve['x_end'])), int(round(curve['y_end'])))
            cv2.circle(debug_img, endpoint, 6, (0, 0, 255), -1)
            cv2.line(debug_img, (0, endpoint[1]), (w, endpoint[1]), (255, 0, 255), 1)
            cv2.putText(
                debug_img,
                f"{curve['mask_name']} mask",
                (x_global, max(20, y - 8)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.5,
                (255, 0, 0),
                1,
            )

        if not result.get('ok'):
            self.log(f"FAILURE: {result.get('error', 'Unknown error')}")
            self.display_preview(debug_img)
            return

        self.log(f"Axis Inliers: {len(axis_model['inliers'])}")
        self.log(f"Top: {axis_model['val_top']} @ {axis_model['y_top_px']:.1f}px")
        self.log(f"Bottom: {axis_model['val_bottom']} @ {axis_model['y_bottom_px']:.1f}px")
        if curve.get('legend_rect'):
            self.log(f"Legend Rect: {curve['legend_rect']}")
        self.log(f"Curve End Y: {curve['y_end']:.1f} ({curve['mask_name']})")
        self.log(f"RESULT: {result['value']:.4e}")

        cv2.putText(debug_img, f"Result: {result['value']:.4e}", (int(w / 2), 50), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2)
        self.display_preview(debug_img)

if __name__ == "__main__":
    root = tk.Tk()
    app = DiagnoseApp(root)
    root.mainloop()
