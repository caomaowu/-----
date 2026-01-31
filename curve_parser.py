import cv2
import numpy as np
import pytesseract
import os
import re

# Set Tesseract path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_TESSERACT = os.path.join(BASE_DIR, "Tesseract-OCR", "tesseract.exe")

# Common system paths for Tesseract
SYSTEM_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.getenv("LOCALAPPDATA", ""), r"Tesseract-OCR\tesseract.exe")
]

# Determine Tesseract command
TESSERACT_CMD = None
if os.path.exists(LOCAL_TESSERACT):
    TESSERACT_CMD = LOCAL_TESSERACT
else:
    # Fallback to system paths
    for p in SYSTEM_PATHS:
        if os.path.exists(p):
            TESSERACT_CMD = p
            break

# Configure pytesseract
if TESSERACT_CMD:
    print(f"DEBUG: Using Tesseract at {TESSERACT_CMD}")
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    print("Warning: Tesseract-OCR not found in local or system paths. OCR features will fail.")

class CurveParser:
    def __init__(self, image_path):
        self.image_path = image_path
        # Use imdecode to handle Chinese paths on Windows
        try:
            self.img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception:
            self.img = None
            
        if self.img is None:
            # Fallback to standard imread just in case
            self.img = cv2.imread(image_path)
            
        if self.img is None:
            raise ValueError(f"Could not load image: {image_path}")
        self.gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        self.height, self.width = self.img.shape[:2]
        
    def detect_chart_area(self):
        """
        Detects potential chart areas.
        Returns a list of candidate rects (x, y, w, h), sorted by area (largest first).
        """
        # Method 1: Thresholding (good for white backgrounds)
        _, thresh = cv2.threshold(self.gray, 240, 255, cv2.THRESH_BINARY)
        # Invert so white background becomes black, black lines become white
        thresh_inv = cv2.bitwise_not(thresh)
        
        # Method 2: Canny Edges (good for finding frames)
        edges = cv2.Canny(self.gray, 50, 150)
        
        # Combine
        combined = cv2.bitwise_or(thresh_inv, edges)
        
        # Dilate to connect broken lines (grid lines, axis lines)
        kernel = np.ones((3,3), np.uint8)
        dilated = cv2.dilate(combined, kernel, iterations=2)
        
        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        candidates = []
        img_area = self.width * self.height
        
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            area = w * h
            # Filter out small areas (noise) - assume chart is at least 5% of image
            if area > (img_area * 0.05):
                candidates.append((x, y, w, h))
        
        # Sort by area descending
        candidates.sort(key=lambda r: r[2] * r[3], reverse=True)
        
        return candidates

    def calibrate_y_axis(self, chart_rect):
        """
        OCR the area to the left of the chart to find Y-axis labels.
        Returns a function (or m, c) to map pixel Y to physical Y.
        """
        x, y, w, h = chart_rect
        
        # Strategy Update:
        # The 'chart_rect' is likely the entire white background including labels.
        # We should look at the left side of this rect, not outside it, especially if x=0.
        
        # Define ROI: Left 15% of the chart rect, or fixed 150px, whichever is reasonable.
        # But we also need to consider if we detected the 'inner' plot area.
        # Let's cover a strip that straddles the left edge 'x'.
        # [x - 50, x + 150]
        
        roi_x_start = max(0, x - 50)
        roi_x_end = min(self.width, x + 150)
        roi_w = roi_x_end - roi_x_start
        
        roi_y = y - 20 
        roi_h = h + 40
        roi_y = max(0, roi_y)
        roi_h = min(self.height - roi_y, roi_h)
        
        if roi_w <= 10 or roi_h <= 10:
             return None
 
        y_axis_roi = self.img[roi_y:roi_y+roi_h, roi_x_start:roi_x_start+roi_w]
        
        # --- Preprocess for OCR (Enhanced) ---
        # 1. Resize: Scale up by 3x to help with small fonts
        scale_factor = 3
        roi_large = cv2.resize(y_axis_roi, None, fx=scale_factor, fy=scale_factor, interpolation=cv2.INTER_CUBIC)
        
        roi_gray = cv2.cvtColor(roi_large, cv2.COLOR_BGR2GRAY)
        
        # 2. Thresholding: Use adaptive or Otsu
        _, roi_thresh = cv2.threshold(roi_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Debug: save OCR ROI
        # cv2.imwrite(f"debug_roi_{x}.png", roi_thresh)
        
        # Use Tesseract to get data with bounding boxes
        # psm 6: Assume a single uniform block of text.
        try:
            data = pytesseract.image_to_data(roi_thresh, output_type=pytesseract.Output.DICT, config='--psm 6')
        except Exception as e:
            print(f"OCR Error: {e}")
            return None
        
        points = []
        for i, text in enumerate(data['text']):
            text = text.strip()
            if not text:
                continue
            
            # print(f"DEBUG OCR RAW: '{text}'") # ADDED DEBUG
                
            # Clean text: replace common OCR errors for numbers
            # e.g. "O" -> "0", "l" -> "1", spaces -> dots if likely float
            clean_text = text.replace('O', '0').replace('o', '0').replace('l', '1').replace(' ', '.')
            
            # Handle cases where "0.042" might be read as "0 042" or "0042" if dot is missed
            # But regex below handles most non-numeric removal.
            
            # Remove non-numeric chars except . and - and e (scientific notation)
            clean_text = re.sub(r'[^\d.-eE]', '', clean_text)
            
            try:
                if not clean_text: continue
                val = float(clean_text)
                
                # Get center Y of the text box relative to ROI
                # Remember to scale back the coordinates!
                text_mid_y_scaled = data['top'][i] + data['height'][i] / 2
                text_mid_y = text_mid_y_scaled / scale_factor
                
                # Convert to global Y
                global_y = roi_y + text_mid_y
                
                points.append((global_y, val))
                # Debug print for analysis
                print(f"DEBUG OCR: Found axis label: {val} at y={global_y}")
            except ValueError:
                continue
        
        if len(points) < 2:
            # print("Warning: Not enough Y-axis points found for calibration.")
            return None
            
        # Linear Regression: Phys = m * Pixel + c
        pixel_ys = np.array([p[0] for p in points])
        phys_vals = np.array([p[1] for p in points])
        
        # RANSAC-like robust fitting
        best_m, best_c = 0, 0
        best_inliers_count = 0
        best_error = float('inf')
        
        # Iterative fitting
        points_arr = np.array(points)
        pixel_ys = points_arr[:, 0]
        phys_vals = points_arr[:, 1]
        
        # Try to fit multiple times or just standard RANSAC
        # Simplified RANSAC:
        n_points = len(points)
        if n_points < 3:
            # Fallback to simple least squares
             A = np.vstack([pixel_ys, np.ones(len(pixel_ys))]).T
             m, c = np.linalg.lstsq(A, phys_vals, rcond=None)[0]
             return m, c

        for _ in range(20): # 20 iterations
            # Pick 2 random points
            indices = np.random.choice(n_points, 2, replace=False)
            p1 = points[indices[0]]
            p2 = points[indices[1]]
            
            if abs(p1[0] - p2[0]) < 5: continue # Too close in Y
            
            # Fit line
            m_curr = (p1[1] - p2[1]) / (p1[0] - p2[0])
            c_curr = p1[1] - m_curr * p1[0]
            
            # Count inliers
            expected_vals = m_curr * pixel_ys + c_curr
            errors = np.abs(expected_vals - phys_vals)
            
            # Threshold: 10% of value range or fixed small value?
            # Y-axis values are small (0.04), so error threshold must be small.
            # But X-axis labels (235) will have HUGE error.
            # Let's use a relative threshold or fixed small number if values are small.
            # If values are ~0.04, threshold 0.005 is good.
            
            inliers = errors < 0.01 # Strict threshold for small numbers
            inliers_count = np.sum(inliers)
            
            if inliers_count > best_inliers_count:
                best_inliers_count = inliers_count
                best_m = m_curr
                best_c = c_curr
                
        # Final refit with all inliers
        if best_inliers_count > 2:
            expected_vals = best_m * pixel_ys + best_c
            errors = np.abs(expected_vals - phys_vals)
            inliers_mask = errors < 0.01
            
            clean_pixels = pixel_ys[inliers_mask]
            clean_phys = phys_vals[inliers_mask]
            
            A = np.vstack([clean_pixels, np.ones(len(clean_pixels))]).T
            m, c = np.linalg.lstsq(A, clean_phys, rcond=None)[0]
            return m, c
        else:
            return None # Failed to find a good line

    def get_curve_end_value(self, chart_rect, m, c, target_color_mask=None):
        """
        Finds the right-most point of the curve in the chart area and maps it to physical value.
        """
        x, y, w, h = chart_rect
        chart_img = self.img[y:y+h, x:x+w]
        
        # 1. Segment the curve
        # Convert to HSV
        hsv = cv2.cvtColor(chart_img, cv2.COLOR_BGR2HSV)
        
        # If no specific color requested, assume non-white/non-black/non-gray is the curve
        # Mask out white background
        lower_white = np.array([0, 0, 200])
        upper_white = np.array([180, 50, 255])
        mask_bg = cv2.inRange(hsv, lower_white, upper_white)
        
        # Mask out black/gray grid lines
        lower_black = np.array([0, 0, 0])
        upper_black = np.array([180, 255, 150]) # Adjust brightness threshold for grid
        mask_grid = cv2.inRange(hsv, lower_black, upper_black)
        
        # Combine masks to find "content"
        mask_noise = cv2.bitwise_or(mask_bg, mask_grid)
        mask_curve = cv2.bitwise_not(mask_noise)
        
        # TODO: Apply specific color mask if provided
        
        # 2. Find right-most pixel
        # Scan from right to left
        h_chart, w_chart = chart_img.shape[:2]
        
        end_point = None
        # Ignore right-most margin (sometimes there's a border)
        margin_right = 5 
        
        for col in range(w_chart - 1 - margin_right, 0, -1):
            column_pixels = mask_curve[:, col]
            non_zeros = np.where(column_pixels > 0)[0]
            
            if len(non_zeros) > 0:
                # Found curve pixels
                # Take the median Y to represent the curve center
                avg_y = np.median(non_zeros)
                end_point = (col, avg_y)
                break
        
        if end_point:
            # Convert chart-relative Y to global Y
            global_end_y = y + end_point[1]
            
            # Map to physical value
            final_value = m * global_end_y + c
            return final_value
            
        print("Warning: No curve detected.")
        return None

    def parse(self):
        # 1. Detect Chart Candidates
        candidates = self.detect_chart_area()
        if not candidates:
            print("Error: Could not detect any potential chart area.")
            return None
            
        # 2. Try to Calibrate on candidates
        # We iterate through the top 3 candidates (largest areas)
        for i, rect in enumerate(candidates[:3]):
            # print(f"Checking candidate {i}: {rect}")
            calibration = self.calibrate_y_axis(rect)
            
            if calibration:
                m, c = calibration
                # print(f"Calibration successful on candidate {i}: m={m:.4f}, c={c:.2f}")
                
                # 3. Get End Value
                result = self.get_curve_end_value(rect, m, c)
                if result is not None:
                    return result
                else:
                    print(f"Warning: Calibration OK but no curve found in candidate {i}.")
            else:
                pass
                # print(f"Calibration failed on candidate {i}.")
        
        print("Error: Failed to extract data from any candidate area.")
        return None

if __name__ == "__main__":
    # Simple test if run directly
    import sys
    if len(sys.argv) > 1:
        parser = CurveParser(sys.argv[1])
        val = parser.parse()
        print(f"Extracted Value: {val}")
