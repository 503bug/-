# 文档批量加密工具

批量加密 Word / Excel / PPT / PDF 文件的 Windows 桌面工具。

## 特性

- 支持 Microsoft Office / WPS Office / 永中 Office / LibreOffice 多引擎自动切换
- 支持 `.doc` `.docx` `.docm` `.xls` `.xlsx` `.xlsm` `.ppt` `.pptx` `.pptm` `.pdf`
- 通过文件魔数 + msoffcrypto 智能识别已加密文件并跳过
- 加密后自动恢复文件的时间戳（访问 / 修改 / 创建时间）
- 内置安全随机密码生成器
- tkinter 图形界面，实时处理日志

## 环境要求

- **操作系统**：Windows（依赖 win32com）
- **Python**：3.8+
- **Office 引擎**（至少其一）：
  - Microsoft Office 2010+
  - WPS Office
  - 永中 Office
  - LibreOffice（需使用其自带的 Python 运行，以提供 `uno` 模块）

## 依赖说明
包	                    用途	                    是否必需
pywin32	              COM 接口调用 Office	        Windows 必需
pikepdf	              PDF 加密	                  加密 PDF 必需
msoffcrypto-tool	    旧格式加密检测	              推荐
uno	                  LibreOffice 接口	          使用 LibreOffice 时必需（随 LibreOffice 自带 Python 提供）

## 使用

bash
python src/doc_encryptor.py
选择目标文件夹

