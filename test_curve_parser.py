import cv2
import numpy as np
import os
from curve_parser import CurveParser

def create_dummy_chart(filename):
    # Create a 800x600 image
    img = np.zeros((600, 800, 3), dtype=np.uint8)
    img.fill(200) # Gray background
    
    # Draw White Chart Area
    # (x, y, w, h) = (100, 50, 600, 500)
    cv2.rectangle(img, (100, 50), (700, 550), (255, 255, 255), -1)
    
    # Draw Axis Labels (Black text on left)
    font = cv2.FONT_HERSHEY_SIMPLEX
    # Top label "100" at y=60 (inside chart area y=50..550)
    # Actually labels are usually centered on the tick.
    # Let's put "100.0" at y=50
    cv2.putText(img, "100.0", (10, 60), font, 1, (0, 0, 0), 2)
    
    # Bottom label "0.0" at y=550
    cv2.putText(img, "0.0", (30, 550), font, 1, (0, 0, 0), 2)
    
    # Draw Curve (Red)
    # Start at bottom-left (100, 550) -> Value 0
    # End at top-right (700, 50) -> Value 100
    # Color: Red (0, 0, 255)
    cv2.line(img, (100, 550), (700, 50), (0, 0, 255), 3)
    
    cv2.imwrite(filename, img)
    print(f"Created dummy chart: {filename}")
    return filename

def test_parser():
    dummy_file = "test_chart.png"
    create_dummy_chart(dummy_file)
    
    parser = CurveParser(dummy_file)
    
    print("--- Testing Chart Detection ---")
    candidates = parser.detect_chart_area()
    print(f"Detected Candidates: {candidates}")
    
    if not candidates:
        print("No chart detected!")
        return

    # Take the first candidate
    rect = candidates[0]
    
    print("\n--- Testing Calibration ---")
    # This might fail if Tesseract is not installed/configured or fails to read generated text
    try:
        m, c = parser.calibrate_y_axis(rect)
        print(f"Slope (m): {m}, Intercept (c): {c}")
        
        # Check logic:
        # Pixel Y=50 should be 100.0
        # Pixel Y=550 should be 0.0
        # Diff Y = 500 pixels. Diff Val = -100.
        # m should be approx -100/500 = -0.2
        print(f"Expected m ~ -0.2. Actual: {m}")
    except TypeError:
        print("Calibration failed (likely OCR issue). Skipping value check.")
        return

    print("\n--- Testing Curve Extraction ---")
    val = parser.get_curve_end_value(rect, m, c)
    print(f"Extracted End Value: {val}")
    print("Expected Value: ~100.0")
    
    # Clean up
    # os.remove(dummy_file)

if __name__ == "__main__":
    test_parser()
