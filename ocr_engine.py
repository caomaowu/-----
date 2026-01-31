import pytesseract
import cv2
import os
import re
import numpy as np

# Try to find tesseract executable in common Windows paths
COMMON_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.getenv("LOCALAPPDATA", ""), r"Tesseract-OCR\tesseract.exe")
]

tesseract_cmd = None
for p in COMMON_PATHS:
    if os.path.exists(p):
        tesseract_cmd = p
        break

if tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

# Predefined OCR Rules for Moldflow/Casting Simulation Results
OCR_RULES = {
    # 规则格式:
    # "rule_name": {
    #     "keywords": ["关键词1", "关键词2"], # 只要包含其中一个关键词，就认为可能包含该数据
    #     "pattern": r"正则",               # 提取数值的正则
    #     "unit_convert": func              # 单位转换/格式化函数
    # }
    
    # 1. 卷气质量 (Entrained Air Mass)
    "air_mass": {
        "keywords": ["Entrained Air Mass", "Air Mass"],
        "pattern": r"Entrained Air Mass:?\s*([\d\.eE\+\-]+)\s*kg",
        "unit_convert": lambda x: f"{float(x)*1000:.4f}g" # kg -> g
    },
    
    # 2. 表面缺陷/氧化物 (Surface Defect Mass)
    # 用户示例: "Surface Defect Mass: 7.360e-02 kg" -> 目标 "27g" (这里假设直接换算，如果是特定逻辑需调整)
    # 注意：用户示例中的27g与7.36e-02kg(73.6g)不符，可能需要用户确认逻辑。
    # 这里暂时按标准物理量换算实现。
    "defect_mass": {
        "keywords": ["Surface Defect Mass", "Defect Mass"],
        "pattern": r"Surface Defect Mass:?\s*([\d\.eE\+\-]+)\s*kg",
        "unit_convert": lambda x: f"{float(x)*1000:.1f}g"
    },
    
    # 3. 卷气体积 (Entrained Air Volume)
    "air_volume": {
        "keywords": ["Entrained Air Volume"],
        "pattern": r"Entrained Air Volume:?\s*([\d\.eE\+\-]+)\s*m\^3",
        "unit_convert": lambda x: f"{float(x)*1e6:.2f}cm³" # m^3 -> cm^3
    },

    # 4. 缩孔体积 (Porosity Volume)
    "porosity": {
        "keywords": ["Porosity Volume", "Shrinkage Volume"],
        "pattern": r"(?:Porosity|Shrinkage) Volume:?\s*([\d\.eE\+\-]+)\s*m\^3",
        "unit_convert": lambda x: f"{float(x)*1e6:.2f}cm³"
    },
    
    # 5. 通用数值提取 (Value) - 仅提取数字
    "value": {
        "keywords": [],
        "pattern": r"([\d\.eE\+\-]+)",
        "unit_convert": lambda x: str(x)
    }
}

class OCREngine:
    def __init__(self):
        self.available = tesseract_cmd is not None
        if not self.available:
            print("Warning: Tesseract-OCR not found. OCR features will be disabled.")

    def preprocess_image(self, image_path):
        """
        Reads image, converts to grayscale and applies thresholding for better OCR.
        """
        try:
            # Handle non-ascii paths
            img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if img is None: return None
            
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Simple thresholding
            # _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
            
            # Otsu's thresholding might be better for varying lighting
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # 2x scale up can help with small text
            h, w = binary.shape[:2]
            scaled = cv2.resize(binary, (w*2, h*2), interpolation=cv2.INTER_CUBIC)
            
            return scaled
        except Exception as e:
            print(f"Image preprocessing failed: {e}")
            return None

    def extract_data(self, image_path, rule_names):
        """
        Run OCR on image and extract data based on rule_names.
        rule_names: list of strings, e.g. ["mass", "volume"]
        Returns: dict { "rule_name": "extracted_value" }
        """
        if not self.available:
            return {}

        processed_img = self.preprocess_image(image_path)
        if processed_img is None:
            return {}

        try:
            # Run Tesseract
            # --psm 6: Assume a single uniform block of text. Good for screenshots.
            # --psm 3: Fully automatic page segmentation, but no OSD. (Default)
            # --psm 11: Sparse text.
            text = pytesseract.image_to_string(processed_img, lang='eng', config='--psm 3')
            
            # Also try without preprocessing if result is empty? 
            # Sometimes binary fails if contrast is already good.
            if not text.strip():
                img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_GRAYSCALE)
                text = pytesseract.image_to_string(img, lang='eng', config='--psm 3')

        except Exception as e:
            print(f"OCR execution failed: {e}")
            return {}

        results = {}
        for rule_key in rule_names:
            rule = OCR_RULES.get(rule_key)
            if not rule:
                continue
                
            # Regex Match
            match = re.search(rule["pattern"], text, re.IGNORECASE | re.MULTILINE)
            if match:
                raw_val = match.group(1)
                try:
                    val = rule["unit_convert"](raw_val)
                    results[rule_key] = val
                except:
                    results[rule_key] = raw_val # Fallback to raw
            else:
                # Debug info
                # print(f"DEBUG: Rule {rule_key} matched nothing in:\n{text[:100]}...")
                pass
                
        return results
