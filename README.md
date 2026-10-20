# 📄 文档批量加密工具（Word / Excel / PPT / PDF）

一款用于 **批量原地加密** 办公文档与 PDF 的桌面小工具，支持 Microsoft Office、WPS Office、永中 Office 与 LibreOffice 多引擎自动切换。开箱即用的 Tkinter GUI，无需命令行经验。

---

## ✨ 功能特性

- **多格式支持**：`.docx / .docm / .doc / .xlsx / .xlsm / .xls / .pptx / .pptm / .ppt / .pdf`
- **多引擎自动探测**：依次尝试 Microsoft Office → WPS Office → 永中 Office → LibreOffice，取第一个可用引擎
- **智能加密检测**：通过文件魔数、`msoffcrypto`、COM / UNO 三级策略判断文件是否已加密，**已加密文件自动跳过**
- **零数据损坏**：加密成功后才会替换原文件；PDF 加密走临时文件 + 原子替换
- **保留文件时间**：加密后自动恢复原始访问/修改时间（Windows 下同时恢复创建时间）
- **子目录递归**：可选是否包含子文件夹
- **随机密码生成**：内置强随机密码生成器，一键填充
- **实时日志**：逐条显示每个文件的处理结果与失败原因
- **跨平台降级**：非 Windows 环境自动降级为 PDF 加密模式（或 LibreOffice 模式）

---

## 🧩 支持矩阵

| 格式 | Office COM | WPS | 永中 | LibreOffice | pikepdf |
|------|:----------:|:---:|:----:|:-----------:|:-------:|
| `.docx` / `.docm` | ✅ | ✅ | ✅ | ✅ | — |
| `.doc` | ✅ | ✅ | ✅ | ✅ | — |
| `.xlsx` / `.xlsm` | ✅ | ✅ | ✅ | ✅ | — |
| `.xls` | ✅ | ✅ | ✅ | ✅ | — |
| `.pptx` / `.pptm` | ✅ | ✅ | ✅ | ✅ | — |
| `.ppt` | ✅ | ✅ | ✅ | ✅ | — |
| `.pdf` | — | — | — | — | ✅ |

> 说明：加密能力由本机已安装的 Office 套件或 LibreOffice 提供；PDF 加密独立于 Office 套件。

---

## 📦 环境要求

- **Python**：3.8+
- **操作系统**：
  - Windows 10 / 11（推荐，完整支持所有 Office 引擎）
  - Linux / macOS（需安装 LibreOffice 才能处理 Office 文件）
- **Office 套件**（可选，处理 Office 文件时必需其一）：
  - Microsoft Office 2010+
  - WPS Office
  - 永中 Office
  - LibreOffice / OpenOffice

---

## 🚀 安装

```bash
git clone https://github.com/<your-name>/office-encryptor.git
cd office-encryptor

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt

```

---

## 许可证

本项目基于 [MIT License](LICENSE) 发布。
