
根据您提供的 `pytesseract曲线识别方案.md`，我制定了以下实施计划。

## 1. 环境准备
- 安装 Python 依赖库 `pytesseract` (检测到尚未安装)。
- 确认 Tesseract 可执行文件路径配置正确 (已确认在 `third_party/tesseract/tesseract.exe`)。

## 2. 核心功能实现 (`utils.py`)
将在 `utils.py` 中新增以下功能：
- **配置 Tesseract**：设置 `tesseract_cmd` 和 `tessdata_dir_config` 指向项目内的 `third_party` 目录。
- **图像识别算法 (`detect_curve_end_value`)**：
    1.  **图像预处理**：读取图片，裁剪右半部分 (默认 50%)。
    2.  **坐标轴定位**：使用 Canny 边缘检测和 Hough 变换定位纵轴位置。
    3.  **刻度识别 (OCR)**：裁剪纵轴区域，使用 `pytesseract` 识别 Y 轴的最大值和最小值。
    4.  **曲线末端定位**：在绘图区通过颜色阈值和形态学操作提取曲线，找到最右侧的曲线像素点。
    5.  **数值计算**：根据末端点在 Y 轴的相对位置，结合 Y 轴刻度范围，线性插值计算出实际数值。

## 3. PPT 操作扩展 (`ppt_automation.py`)
在 `PPTAutomation` 类中新增方法：
- **`export_smarttag_image(self, key, output_path)`**：根据 `SmartTag_<key>` 名称定位形状，并将其导出为临时图片文件，作为识别算法的输入。
- **`replace_text_placeholder(self, key, value)`**：扫描全文档，查找 `VAL_<key>` 文本占位符，并将其替换为计算出的数值字符串。

## 4. 流程集成 (`generate_report.py`)
修改主程序逻辑：
- 在完成所有媒体文件的插入操作**之后**，新增一个处理阶段。
- 遍历已插入的 Key，对于存在对应 `VAL_<key>` 占位符的项目：
    1.  调用 `export_smarttag_image` 导出图片。
    2.  调用 `detect_curve_end_value` 识别数值。
    3.  调用 `replace_text_placeholder` 更新 PPT 内容。
- 最后清理临时导出的图片文件。

## 5. 验证计划
1.  **单元测试**：使用项目目录下的 `test.png` 运行 `detect_curve_end_value`，验证是否能输出接近 `0.027` 的数值。
2.  **集成测试**：运行 `generate_report.py`，检查生成的 PPT 中 `VAL_` 占位符是否被正确替换。
