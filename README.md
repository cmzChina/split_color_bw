# split_color_bw

> 同页黑白 / 彩色内容分离工具 v2.0
> Split colored PDF/Word documents into two PDFs: one black-and-white, one color.

将一份包含彩色和黑白内容的文档，按**页内对象**自动拆分为两个 PDF：
- **黑白版** — 仅保留黑色 / 灰色文字与图形，适合黑白打印、降低打印成本
- **彩色版** — 仅保留彩色元素，适合核对彩色内容、单独彩印

---

## 功能特性 / Features

- **矢量无损模式（Vector）**：不把页面渲染成图片，直接按对象颜色删除。输出 PDF 仍是矢量，文字可复制、无限放大不模糊。
- **像素模式（Raster）**：将页面渲染为图片后逐像素判定颜色，适用于扫描件、图片型 PDF。
- **自动模式（Auto，默认）**：逐页判断——图片覆盖率超过 80% 的页面视为扫描页，走像素模式；其余页面走矢量模式。
- **支持 PDF / Word（.doc / .docx）**：Word 自动转 PDF 后再分离，支持 WPS、Microsoft Word（COM 自动化）和 LibreOffice（headless）三种后端。
- **批量 & 递归处理**：自动扫描目标文件夹及其子目录中的所有文档。
- **拖拽即用**：把文件或文件夹拖到程序图标上即可处理；也可直接双击运行（处理程序所在目录）。
- **统计报告**：输出每页彩色占比的 CSV 报告，便于核对。
- **自动安装依赖**：源码运行时若缺少库，会自动从清华 PyPI 镜像安装。

---

## 快速开始 / Quick Start

### 方式一：便携版 EXE（推荐普通用户）

1. 下载 `split_color_bw_portable.exe`（约 45 MB，已内置所有依赖）。
2. 将 exe 放到存放 PDF / Word 的文件夹中，**双击运行**。
3. 或把**单个文件 / 整个文件夹**拖到 exe 图标上。
4. 处理完成后，结果输出到源目录下的 `export/` 文件夹。

### 方式二：源码运行（推荐开发者）

```bash
# 安装依赖（程序也会自动安装）
pip install pymupdf numpy pillow pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple

# 处理当前目录
python split_color_bw.py

# 处理指定文件或文件夹
python split_color_bw.py "C:\path\to\your\file.pdf"
python split_color_bw.py "C:\path\to\your\folder"
```

> **Python 版本**：建议 3.10 及以上。
> **操作系统**：Windows（Word 转换依赖 Windows COM 或 LibreOffice；纯 PDF 处理可跨平台）。

---

## 输出说明 / Output

所有结果输出到**源文件所在目录**的 `export/` 子文件夹：

| 输出文件 | 说明 |
|---|---|
| `xxx_黑白.pdf` | 仅保留黑白内容的 PDF |
| `xxx_彩色.pdf` | 仅保留彩色内容的 PDF |
| `YYYY年MM月DD日HH时MM分_统计报告.csv` | 每页彩色 / 黑白占比统计 |

> 若同名输出已存在，会自动追加 `(1)`、`(2)` 等序号，不会覆盖已有文件。

---

## 可调参数 / Configuration

打开 `split_color_bw.py`，修改文件顶部的参数区：

```python
MODE     = "auto"   # "vector" 矢量无损 | "raster" 像素模式 | "auto" 自动判断
DPI      = 300      # 渲染精度，仅 raster 模式 / 扫描页生效
SAT_THR  = 0.12     # 饱和度阈值 (0~1)，漏判彩色可调小到 0.08
DIFF_THR = 18       # 通道差阈值 (0~255)，发黄扫描件全页判彩色可调大到 30
CLEAN    = 5        # 边缘清理 (奇数 3/5/7/9)，仅像素模式
DENOISE  = 0        # 去噪点 (奇数 3/5，0=关闭)，仅像素模式
EXPORT   = "export" # 输出文件夹名
RECURSIVE = True    # 是否递归扫描子目录

SCAN_RATIO      = 0.80  # auto 模式：单页图片覆盖率超过 80% 视为扫描页
IMG_COLOR_RATIO = 0.02  # 矢量模式：图片中彩色像素超过 2% 判为彩色图像
```

### 参数调优建议

| 现象 | 调整 |
|---|---|
| 部分彩色内容被误删（漏判彩色） | 把 `SAT_THR` 调小，如 `0.08` |
| 发黄 / 偏色的扫描件整页被判为彩色 | 把 `DIFF_THR` 调大，如 `30` |
| 扫描件彩色边缘有杂色 | 增大 `CLEAN`（如 `7`），或开启 `DENOISE`（如 `3`） |
| 扫描件处理慢 / 文件大 | 降低 `DPI`（如 `200`） |
| 不希望处理子文件夹 | 设 `RECURSIVE = False` |

---

## 颜色判定原理 / How Color Detection Works

对每个页面对象（文字 span、矢量图形 path、图片），提取其颜色并转换为 RGB：

```
diff = max(R, G, B) - min(R, G, B)
判定为彩色 ⟺ diff > DIFF_THR 且 (diff / max) > SAT_THR
```

- **矢量模式**：用 PyMuPDF 的 redaction（涂黑/删除）功能，物理删除对应颜色的对象矩形区域。
- **像素模式**：将页面渲染为位图，生成彩色掩码后分别合成黑白页和彩色页。
- **图片对象**：采样图片像素，若彩色像素比例超过 `IMG_COLOR_RATIO` 则整张图归为彩色。

---

## 常见问题 / FAQ

**Q：处理后 PDF 中的文字还能复制吗？**
A：矢量模式下可以，文字仍为可选中的文本层；像素模式（扫描页）下页面变为图片，文字不可复制。

**Q：Word 文件转换失败怎么办？**
A：请确认已安装 WPS 或 Microsoft Word；若都没有，可安装 [LibreOffice](https://www.libreoffice.org/)，程序会自动尝试 `soffice --headless` 转换。

**Q：程序会自动安装依赖，安全吗？**
A：程序仅从清华 PyPI 镜像（`pypi.tuna.tsinghua.edu.cn`）安装 `pymupdf`、`numpy`、`pillow`、`pywin32` 这四个公开库。如不希望自动安装，可提前手动 `pip install` 后再运行。

**Q：输出的黑白版和彩色版加起来不等于原文件大小？**
A：正常。分离后两个 PDF 各自经过 `garbage=4, deflate=True` 压缩优化，且删除了部分对象，总大小通常与原文件不同。

**Q：加密 / 受密码保护的 PDF 能处理吗？**
A：当前版本不支持加密 PDF。请先解除密码保护后再处理。

**Q：Mac / Linux 能用吗？**
A：纯 PDF 处理可以（需安装 Python 依赖）；Word 转换功能仅支持 Windows（依赖 COM 或 Windows 版 LibreOffice）。

---

## 项目结构 / Project Structure

```
split_color_bw/
├── split_color_bw.py              # 主程序源码
├── split_color_bw_portable.exe    # 便携版（单文件，已内置依赖）
├── split_color_bw_unportable/     # 非便携版（exe + _internal 依赖目录）
├── export/                         # 默认输出目录（运行时生成）
├── LICENSE                         # MIT 许可证
└── README.md                       # 本文件
```

> `split_color_bw_portable.exe` 和 `split_color_bw_unportable/` 为打包产物，
> 不纳入版本控制（见 `.gitignore`），可在 Releases 页面下载。

---

## 依赖 / Dependencies

| 库 | 用途 | 必需 |
|---|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF 读取、对象提取、redaction 删除 | 是 |
| [NumPy](https://numpy.org/) | 像素数组运算、颜色掩码 | 是 |
| [Pillow](https://python-pillow.org/) | 图片读取、格式转换、滤镜 | 是 |
| [pywin32](https://github.com/mhammond/pywin32) | Windows COM 调用 WPS/Word 转 PDF | 否（Word 转换需要） |
| LibreOffice | Word 转 PDF 备选后端 | 否（无 WPS/Word 时需要） |

---

## 许可证 / License

[MIT License](./LICENSE) — Copyright (c) 2026 chen
