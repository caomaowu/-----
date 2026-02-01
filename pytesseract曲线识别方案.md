# pytesseract 曲线末端纵坐标识别方案

## 1. 目标
- 自动识别结果图右侧曲线末端的纵坐标数值
- 将识别结果替换到 PPT 指定文本占位符
- 适配不同结果图片（尺寸、网格、范围变化）

## 2. 依赖与安装位置
### 2.1 Python 依赖
- `pytesseract`
- `opencv-python`
- `numpy`

### 2.2 Tesseract 本体安装位置（程序文件夹内，已经完成安装）
自动化方案/
├── third_party/
│   └── tesseract/
│       ├── tesseract.exe
│       └── tessdata/
│           ├── eng.traineddata
│           └── chi_sim.traineddata
```

运行时在代码中设置：
- `pytesseract.pytesseract.tesseract_cmd = <程序目录>/third_party/tesseract/tesseract.exe`
- `TESSDATA_PREFIX = <程序目录>/third_party/tesseract/`

## 3. 识别流程（算法）
### 3.1 图像来源（方案A：直接使用PPT内图片）
1. 插入完成后，按 `SmartTag_<key>` 定位图片形状
2. 使用 `shape.Export(临时路径)` 导出为 PNG
3. 将导出图作为识别输入
4. 仅对图片形状导出（视频形状跳过，或先抽帧再作为图片）
5. 只导出存在 `VAL_<key>` 占位符的 key，避免无效导出

### 3.1.1 导出路径与清理
- 临时导出目录：`<程序目录>/temp_exports`
- 文件命名：`<key>.png`
- 失败不清理，成功后清理临时导出目录

### 3.2 图像裁剪与绘图区定位
1. 输入图像（PPT导出的图片）
2. 右半区裁剪（默认右 50%，可配置 0.45~0.6）
3. 识别绘图区 ROI
   - 主策略：Canny + Hough 直线检测，定位纵轴/横轴边界
   - 兜底策略：阈值分离非白像素，取最大连通区域外接矩形
4. OCR 前可放大 2x 以提升小字识别率

### 3.3 曲线提取与末端点定位
1. 灰度化 + 自适应阈值，提取深色曲线
2. 形态学开操作移除网格线（水平/垂直细线）
3. 在 ROI 内寻找最靠右曲线像素簇
4. 取该簇 y 坐标中位数作为 `y_end`

### 3.4 纵轴刻度 OCR 与数值映射
1. 在绘图区左侧裁剪出纵轴刻度区
2. 使用 pytesseract 识别刻度文本
3. 提取 `y_min` 和 `y_max`
4. 线性映射：
```
value = y_max - (y_end - y_top) / (y_bottom - y_top) * (y_max - y_min)
```
5. 结果按统一格式输出（如 `%.3e` 或 `%.4f`）
6. 默认格式建议：`%.3e`

## 4. PPT 占位符替换规则
- 文本占位符统一使用：`VAL_<key>`
- 例如：`VAL_wendu`
- 识别完成后进行全局文本扫描替换

## 5. 失败兜底策略
- OCR 失败：返回空值并保留原占位符
- 刻度识别不完整：尝试多阈值组合重试
- 曲线提取失败：记录日志并跳过该项
- 导出失败：记录日志并跳过该 key

## 6. 与当前项目集成方式（不改变现有结构）
建议新增三个功能点：
1. `ppt_automation.py`
   - `export_smarttag_image(key, output_path)`
   - 用于导出 `SmartTag_<key>` 的图片形状
2. `utils.py`
   - `detect_curve_end_value(image_path) -> float`
   - 封装裁剪、ROI、曲线、OCR
3. `generate_report.py`
   - 在插入图片后导出、识别，再替换文本

## 7. 验证方式
- 选取 5~10 张不同结果图进行对比
- 输出识别数值与人工读取值对比
- 误差低于设定阈值（如 < 2%）


## 8. 测试样例
- **测试图片路径**：`C:\Users\gcb\Desktop\自动化方案\test.png`
- **读数**：曲线末端纵坐标 ≈ 0.027（区间 0.026–0.028）
- **期望识别误差**：< 2%
