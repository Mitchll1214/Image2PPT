```markdown
# 🖼️ Image2PPT - 图片转PPT工具

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

一款**图片转PPT工具**的桌面工具，可将文件夹中的图片自动生成 **16:9 宽屏 PowerPoint 演示文稿**。支持智能布局、文件名显示、自定义前缀等功能。

## ✨ 特性

- 🖼️ **批量转换** - 支持 JPG、PNG、BMP、GIF、TIFF 等常见图片格式
- 📐 **智能布局** - 中心对称布局算法，自适应不同比例图片（允许±15%调整）
- 🎨 **鎏金UI** - 深黑背景 + 鎏金配色，专业视觉体验
- 📝 **文件名显示** - 可选在图片下方自动显示文件名（16号加粗居中）
- 🔢 **灵活分页** - 可自定义每页图片数量（默认6张）
- 🏷️ **前缀命名** - 自动为PPT文件添加时间戳和自定义前缀
- 📊 **进度提示** - 实时显示转换进度条

## 📸 界面预览

```
<img width="900" height="797" alt="image" src="https://github.com/user-attachments/assets/417df411-0cb5-411f-a7c4-842dc161fa0c" />
```

## 🚀 快速开始

### 环境要求(压缩包中已打包为exe)

- Python 3.7+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install python-pptx pillow pywin32
```

> **注意**：`pythoncom`（pywin32）仅在 Windows 上必需；macOS/Linux 用户可移除相关代码或安装 `pyobjc`

### 运行程序

```bash
python image2ppt.py
```

## 📖 使用说明

1. 点击 **浏览** 选择包含图片的源文件夹
2. 选择PPT输出文件夹（默认与源文件夹相同）
3. （可选）勾选并输入文件名前缀
4. 设置每页显示的图片数量（推荐 4~9 张）
5. 选择是否在图片下方显示文件名
6. 点击 **开始转换**
7. 等待进度条完成，弹出成功提示

## 🧮 布局算法说明

采用 **中心对称自适应布局**：

- 幻灯片尺寸：10" × 5.625"（16:9）
- 自动计算最优行列数（接近正方形排列）
- 保留10%页面边距，图片间距0.15英寸
- 图片比例差异超过15%时进行智能拉伸（不超过±15%）
- 所有图片在单元格内居中显示

## 🗂️ 项目结构

```
Image2PPT/
├── main.py          # 主程序
├── README.md             # 说明文档
└── requirements.txt      # 依赖列表
```

## 📝 依赖库

| 库 | 用途 |
|---|---|
| `python-pptx` | 创建/操作PPT文件 |
| `Pillow` | 获取图片尺寸和DPI信息 |
| `pywin32` | Windows COM初始化（可选） |
| `tkinter` | GUI界面（Python内置） |

## 🛠️ 常见问题

**Q: 为什么有些图片被拉伸了？**
A: 当图片原始比例与单元格比例差异超过15%时，程序会进行不超过15%的适度调整，以充分利用空间。如需完全保留原始比例，可修改代码中的 `0.15` 阈值。

**Q: 支持哪些图片格式？**
A: `.jpg` `.jpeg` `.png` `.bmp` `.gif` `.tif` `.tiff`

**Q: 图片按什么顺序排列？**
A: 按文件名升序排序。

**Q: 生成的PPT太大怎么办？**
A: 程序保留原始图片分辨率。如需压缩，可先批量调整图片尺寸。

## 📄 许可证

MIT License © 2025 FJY

## 👤 作者

**FJY** - 鎏金专业版

---

⭐ 如果这个工具帮到了你，欢迎给个Star！
```
