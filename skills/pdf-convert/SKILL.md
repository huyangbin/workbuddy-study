---
name: pdf-convert
description: PDF 与 Word/PPT/Excel/HTML/图片 之间的常用格式互转。当用户说"PDF 转 Word/PPT/Excel/图片/HTML""Excel/HTML/图片 转 PDF""把 xxx 转成 pdf"等时使用。优先用纯 Python 工具，避开沙箱里 LibreOffice 写不出 OOXML 的坑。
agent_created: true
---

# PDF 常用格式互转工具（纯 Python，沙箱友好）

一个脚本覆盖 8 个方向，全部已实测可用。

## 关键环境约束（必读）
- **沙箱里的 LibreOffice 能写 PDF，但写不出合法的 docx/pptx/xlsx**（OOXML 导出过滤器报 `no export filter` / `impl_store failed`，产物是没有 `word/document.xml` / slide 的空壳）。这与 PPTX→PDF(Impress) 在沙箱失败同源，是环境限制。
- 因此：**凡是"写出 Office 格式"(docx/pptx/xlsx) 一律用纯 Python 库**，不要依赖 LibreOffice 的 OOXML 导出。
- LibreOffice **仅**用于「输入是 Office/HTML、输出是 PDF」时提升保真度（已验证 Writer→PDF、Calc→PDF 可用；Impress→PDF 在沙箱失败）。脚本里 `excel2pdf`/`html2pdf` 已做成"LibreOffice 优先、reportlab 兜底"。
- 网络：`pip` 走官方源 `https://pypi.org/simple` 可用，但**很慢**；`pdf2docx` 依赖 numpy+opencv(50MB+) 在沙箱十多分钟下不完，**勿用 pdf2docx**，改用手写 OOXML（见脚本 `cmd_pdf2docx`）。

## 依赖安装（一次性）
建一个 venv，把需要的库一次装好（约几分钟，比反复装快）：
```
PY=<managed python 绝对路径，如 C:/Users/27598/.workbuddy/binaries/python/versions/3.13.12/python.exe>
$PY -m venv .venv-pdf
.venv-pdf/Scripts/python.exe -m pip install --index-url https://pypi.org/simple --disable-pip-version-check \
    pypdfium2 pdfplumber reportlab pypdf Pillow openpyxl python-pptx
```
脚本位置：技能目录下的 `pdf_convert.py`（或项目内 `pdf_convert.py`）。运行时用 `.venv-pdf/Scripts/python.exe` 调它。

## 用法
```
python pdf_convert.py <子命令> <输入> --out <输出>
```
子命令与示例：
| 方向 | 子命令 | 示例 |
|------|--------|------|
| PDF → 图片 | `pdf2img` | `pdf2img in.pdf --out dir --dpi 150 [--format jpeg]` |
| 图片 → PDF | `img2pdf` | `img2pdf a.png b.png --out out.pdf`（也接受目录） |
| PDF → Excel | `pdf2excel` | `pdf2excel in.pdf --out out.xlsx`（每页一个 sheet，抽表格，无表则放文本） |
| Excel → PDF | `excel2pdf` | `excel2pdf data.xlsx --out data.pdf`（LO 优先，reportlab 兜底） |
| HTML → PDF | `html2pdf` | `html2pdf page.html --out page.pdf`（LO 优先，reportlab 兜底） |
| PDF → HTML | `pdf2html` | `pdf2html in.pdf --out in.html`（每页=截图base64+文本，单文件自包含） |
| PDF → Word | `pdf2docx` | `pdf2docx in.pdf --out in.docx`（手写 OOXML：每页=标题+整页图+文本） |
| PDF → PPT | `pdf2pptx` | `pdf2pptx in.pdf --out in.pptx`（python-pptx：每页=整页图） |

## 实现要点（排错用）
- **pdf2docx 手写 OOXML**：必须声明 `xmlns:wp=...` 命名空间，否则 Word 打不开；PDF 抽取文本含 XML 1.0 非法控制字符，必须用 `sanitize_xml_chars` 清洗；图片尺寸用 PNG IHDR 算 EMU。验证方式：用 LibreOffice 把生成的 docx 当 Writer 文档读并导出 PDF，能出 PDF 即合法。
- **pdf2excel**：openpyxl 拒绝控制字符（`IllegalCharacterError`），写单元格前用 `sanitize_cell`（正则 `[\x00-\x08\x0b\x0c\x0e-\x1f]` 清空）。
- **pdf2img / img2pdf**：统一用 `pypdfium2` 渲染、`Pillow` 拼多页 PDF（`img.save(..., "PDF", save_all=True, append_images=...)`）。
- **pdf2pptx**：`python-pptx` 逐页 `add_picture`，幻灯片尺寸按首页图片宽高比自适应。
- **excel2pdf / html2pdf**：`run_lo_convert()` 调 soffice，`cwd` 必须切到 `program` 目录、用独立 `UserInstallation` profile；产物大小 >1KB 才算成功，否则走 reportlab 兜底。

## 交付前自检
- PDF：读前 5 字节应为 `%PDF-`。
- docx/pptx/xlsx：用 `zipfile` 打开，确认关键部件存在（docx 有 `word/document.xml`+`word/media/`；pptx 有 `ppt/slides/slideN.xml`；xlsx 能被 openpyxl 打开）。
- html：用 `html.parser` 解析不报错，且含预期页数 `<section>`。
