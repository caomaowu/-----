# 自动化报告生成工具

这是一个用于自动化生成PPT报告的工具，可以将视频文件和截图自动插入到PPT模板中的指定位置。

## 功能特性

- **图形化界面**: 提供友好的GUI操作界面
- **批量处理**: 支持批量生成报告
- **视频处理**: 自动提取视频最后一帧作为截图
- **智能匹配**: 根据预定义规则自动匹配视频和PPT页面
- **锚点定位**: 使用锚点形状精确定位视频和图片位置
- **对比报告**: 支持生成对比报告

## 项目结构

```
自动化方案/
├── gui_main.py              # 主程序 - GUI界面
├── generate_report.py       # 普通报告生成脚本
├── generate_compare_report.py # 对比报告生成脚本
├── 自动化模板.pptx          # PPT模板文件
├── 视频命名与位置、坐标.md   # 视频映射配置文件
├── beifen/                  # 备份目录
└── README.md               # 项目说明文档
```

## 环境要求

- Python 3.10+
- Windows操作系统
- wps 

## 依赖库

- `opencv-python` (cv2)
- `pywin32` (win32com)
- `numpy`
- `tkinter` (Python内置)

## 安装步骤

1. 克隆或下载项目到本地
2. 安装依赖库：
```bash
pip install opencv-python pywin32 numpy
```

## 使用方法

### 1. 图形界面方式（推荐）

运行主程序：
```bash
python gui_main.py
```

界面操作步骤：
1. 选择执行脚本（普通生成或对比生成）
2. 选择PPT模板文件
3. 设置视频文件目录
4. 点击"开始生成"按钮

### 2. 命令行方式

普通报告生成：
```bash
python generate_report.py --template 自动化模板.pptx --video_dir 视频
```

对比报告生成：
```bash
python generate_compare_report.py --template 自动化模板.pptx --video_dir 视频
```

## 视频文件映射规则

视频文件与PPT页面的对应关系如下：

| 视频文件名 | PPT页面位置 | 最后一帧位置 |
|------------|-------------|--------------|
| wendu.mp4 | PPT第4页 | 最后一帧放在PPT第5页 |
| juanqi.mp4 | PPT第6页 | 最后一帧放在PPT第8页 |
| tijuanqi.mp4 | PPT第7页 | 最后一帧放在PPT第9页 |
| qipao.mp4 | PPT第10页 | 最后一帧放在PPT第11页 |
| yanghuazha.mp4 | PPT第12页 | 最后一帧放在PPT第14页 |
| tijiyanghua.mp4 | PPT第13页 | 最后一帧放在PPT第15页 |
| liaoliuzhuizong.mp4 | PPT第16页 | 最后一帧放在PPT第17页 |
| suokong.mp4 | PPT第18页 | 最后一帧放在PPT第21页 |
| suokonglv.mp4 | PPT第19页 | 最后一帧放在PPT第22页 |
| suokongtiji.mp4 | PPT第20页 | 最后一帧放在PPT第22页 |
| rejie.mp4 | PPT第23页 | 最后一帧放在PPT第24页 |
| qiya.mp4 | PPT第25页 | 最后一帧放在PPT第26页 |

## 锚点命名规则

在PPT模板中使用以下命名规则的锚点形状：

- **视频锚点**: `VID_<key>` （如：VID_wendu、VID_juanqi）
- **图片锚点**: `IMG_<key>` （如：IMG_wendu、IMG_juanqi）

其中`<key>`对应视频映射中的键名。

## 注意事项

1. 确保视频文件存在于指定的视频目录中
2. PPT模板需要预先设置好锚点形状
3. 如果某个视频文件不存在，对应的PPT页面将保持空白
4. 生成的报告将保存在与脚本相同的目录下
5. 建议使用英文路径，避免中文路径可能导致的兼容性问题

## 故障排除

### 常见问题

1. **PowerPoint无法启动**
   - 确保已安装Microsoft PowerPoint
   - 检查是否有其他PowerPoint实例正在运行

2. **视频文件无法读取**
   - 检查视频文件路径是否正确
   - 确保视频文件格式受支持（推荐mp4格式）
   - 检查视频文件是否损坏

3. **锚点形状找不到**
   - 确认PPT模板中的形状命名是否正确
   - 检查形状是否位于正确的幻灯片上

## 更新日志

- v1.0.0: 初始版本，支持基本报告生成功能
- v1.1.0: 添加GUI界面，支持对比报告生成


**提示**: 使用前请仔细阅读视频命名与位置、坐标.md文件，了解详细的配置规则和使用说明。