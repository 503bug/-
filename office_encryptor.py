#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档批量加密工具（Word / Excel / PPT / PDF）
============================================

对指定文件夹中的 Office（Word/Excel/PPT）与 PDF 文件进行原地批量加密，
已加密文件自动跳过，加密后保留原始访问时间/修改时间。

引擎优先级：
    1. Microsoft Office / WPS Office / 永中 Office（COM，仅 Windows）
    2. LibreOffice / OpenOffice（UNO 桥接，跨平台）
    3. pikepdf（仅 PDF，跨平台）

License: MIT
"""

import os
import sys
import time
import socket
import string
import secrets
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


# ============================================================
# 可选依赖：Windows COM（Office 自动化）
# ============================================================
try:
    import win32com.client
    import win32file
    import win32con
    import pythoncom
    from pywintypes import Time as PyTime
    WINDOWS_COM_AVAILABLE = True
except ImportError:
    win32com = win32file = win32con = pythoncom = PyTime = None
    WINDOWS_COM_AVAILABLE = False

# PDF 加密依赖
try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except ImportError:
    PIKEPDF_AVAILABLE = False

# Office 加密检测依赖（可选，用于旧格式 .doc/.xls/.ppt）
try:
    import msoffcrypto
    MSOFFCRYPTO_AVAILABLE = True
except ImportError:
    MSOFFCRYPTO_AVAILABLE = False


# ============================================================
# 常量
# ============================================================
# OOXML 扩展名（未加密为 ZIP，加密为 CFB）
OOXML_EXTS = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"}
LEGACY_EXTS = {".doc", ".xls", ".ppt"}

# 文件魔数
ZIP_MAGIC = b"PK\x03\x04"
CFB_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

# 隐藏子进程窗口（仅 Windows）
if sys.platform == "win32":
    CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
else:
    CREATE_NO_WINDOW = 0


# ============================================================
# COM 引擎定义（Office / WPS / 永中）
# ============================================================
COM_ENGINES = [
    {
        "name": "Microsoft Office",
        "word": ["Word.Application"],
        "excel": ["Excel.Application"],
        "ppt": ["PowerPoint.Application"],
        "max_pwd": 255,
    },
    {
        "name": "WPS Office",
        "word": ["KWPS.Application", "Wps.Application"],
        "excel": ["KET.Application", "Et.Application"],
        "ppt": ["KWPP.Application", "Kwpp.Application"],
        "max_pwd": 15,
    },
    {
        "name": "永中 Office",
        "word": ["Yozo.Application", "Yozosoft.Application"],
        "excel": ["Yozo.Application", "Yozosoft.Application"],
        "ppt": ["Yozo.Application", "Yozosoft.Application"],
        "max_pwd": 100,
    },
]


def detect_com_engine():
    """探测可用的 COM 引擎，返回 (引擎名, word, excel, ppt, max_pwd)。"""
    if not WINDOWS_COM_AVAILABLE:
        return None, None, None, None, None

    pythoncom.CoInitialize()
    for engine in COM_ENGINES:
        word_app = excel_app = ppt_app = None

        for progid in engine["word"]:
            try:
                word_app = win32com.client.Dispatch(progid)
                word_app.Visible = False
                try:
                    word_app.DisplayAlerts = 0
                except Exception:
                    pass
                break
            except Exception:
                continue

        for progid in engine["excel"]:
            try:
                excel_app = win32com.client.Dispatch(progid)
                excel_app.Visible = False
                try:
                    excel_app.DisplayAlerts = False
                except Exception:
                    pass
                break
            except Exception:
                continue

        for progid in engine.get("ppt", []):
            try:
                ppt_app = win32com.client.Dispatch(progid)
                ppt_app.Visible = False
                try:
                    ppt_app.DisplayAlerts = 0
                except Exception:
                    pass
                break
            except Exception:
                continue

        if any(a is not None for a in (word_app, excel_app, ppt_app)):
            return engine["name"], word_app, excel_app, ppt_app, engine["max_pwd"]

    return None, None, None, None, None


# ============================================================
# LibreOffice UNO 适配
# ============================================================
LIBREOFFICE_PATHS = [
    # Windows
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    # Linux
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/usr/lib64/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    # macOS
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

UNO_PORT = 2002
UNO_HOST = "127.0.0.1"


class LibreOfficeEngine:
    def __init__(self):
        self.proc = None
        self.ctx = None
        self.desktop = None
        self.uno_available = self._check_uno()

    @staticmethod
    def _check_uno():
        try:
            import uno  # noqa: F401
            return True
        except ImportError:
            return False

    def start(self):
        if not self.uno_available:
            raise RuntimeError("缺少 uno 模块。请使用 LibreOffice 自带的 Python 运行。")
        soffice = self._find_soffice()
        if soffice is None:
            raise RuntimeError("未找到 LibreOffice 安装路径。")

        cmd = [
            soffice, "--headless",
            f"--accept=socket,host={UNO_HOST},port={UNO_PORT};urp;",
            "--norestore", "--nofirststartwizard",
        ]
        self.proc = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
        )

        deadline = time.time() + 15
        while time.time() < deadline:
            try:
                with socket.create_connection((UNO_HOST, UNO_PORT), timeout=1):
                    break
            except OSError:
                time.sleep(0.3)
        else:
            self.stop()
            raise RuntimeError("LibreOffice UNO 端口未就绪。")

        self._connect()

    @staticmethod
    def _find_soffice():
        for p in LIBREOFFICE_PATHS:
            if os.path.isfile(p):
                return p
        import shutil
        return shutil.which("soffice")

    def _connect(self):
        import uno
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        self.ctx = resolver.resolve(
            f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext"
        )
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

    def _load(self, filepath, password=None, hidden=True):
        from com.sun.star.beans import PropertyValue
        import uno

        file_url = uno.systemPathToFileUrl(os.path.abspath(filepath))
        props = []
        p_hidden = PropertyValue(); p_hidden.Name = "Hidden"; p_hidden.Value = hidden
        props.append(p_hidden)
        if password is not None:
            p_pwd = PropertyValue(); p_pwd.Name = "Password"; p_pwd.Value = password
            props.append(p_pwd)
        return self.desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))

    def is_encrypted(self, filepath):
        doc = None
        try:
            doc = self._load(filepath, password="", hidden=True)
            return False
        except Exception as e:
            msg = str(e).lower()
            return any(k in msg for k in ("password", "密码", "encrypt", "加密"))
        finally:
            if doc is not None:
                try:
                    doc.close(True)
                except Exception:
                    pass

    def encrypt(self, filepath, password):
        import uno
        from com.sun.star.beans import PropertyValue

        doc = self._load(filepath, password=None, hidden=True)
        if doc is None:
            raise RuntimeError(f"LibreOffice 无法打开文件：{filepath}")

        try:
            ext = os.path.splitext(filepath)[1].lower()
            filter_map = {
                ".docx": "MS Word 2007 XML", ".docm": "MS Word 2007 XML",
                ".doc": "MS Word 97",
                ".xlsx": "Calc MS Excel 2007 XML", ".xlsm": "Calc MS Excel 2007 XML",
                ".xls": "MS Excel 97",
                ".pptx": "Impress MS PowerPoint 2007 XML", ".pptm": "Impress MS PowerPoint 2007 XML",
                ".ppt": "MS PowerPoint 97",
            }
            filter_name = filter_map.get(ext)
            if not filter_name:
                raise RuntimeError(f"LibreOffice 不支持的扩展名：{ext}")

            file_url = uno.systemPathToFileUrl(os.path.abspath(filepath))
            props = []
            p_filter = PropertyValue(); p_filter.Name = "FilterName"; p_filter.Value = filter_name
            props.append(p_filter)
            p_pwd = PropertyValue(); p_pwd.Name = "Password"; p_pwd.Value = password
            props.append(p_pwd)
            doc.storeToURL(file_url, tuple(props))
        finally:
            doc.close(True)

    def stop(self):
        try:
            if self.desktop is not None:
                self.desktop.terminate()
        except Exception:
            pass
        if self.proc is not None:
            try:
                self.proc.terminate()
                self.proc.wait(timeout=5)
            except Exception:
                try:
                    self.proc.kill()
                except Exception:
                    pass
        self.proc = self.desktop = self.ctx = None


# ============================================================
# 主程序
# ============================================================
class MultiEngineEncryptorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文档批量加密工具（Word / Excel / PPT / PDF）")
        self.root.geometry("820x660")
        self.root.resizable(True, True)

        tk.Label(root, text="目标文件夹：", anchor="w").pack(fill="x", padx=10, pady=(10, 0))
        frame_dir = tk.Frame(root)
        frame_dir.pack(fill="x", padx=10, pady=5)
        self.dir_var = tk.StringVar()
        tk.Entry(frame_dir, textvariable=self.dir_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame_dir, text="浏览...", command=self.select_folder).pack(side="right", padx=(5, 0))

        tk.Label(root, text="密码（上限约 10 字符）：", anchor="w").pack(
            fill="x", padx=10, pady=(10, 0)
        )
        frame_pwd = tk.Frame(root)
        frame_pwd.pack(fill="x", padx=10, pady=5)
        self.pwd_var = tk.StringVar()
        tk.Entry(frame_pwd, textvariable=self.pwd_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame_pwd, text="生成随机密码", command=self.generate_password).pack(
            side="right", padx=(5, 0)
        )

        self.include_subdirs = tk.BooleanVar(value=False)
        tk.Checkbutton(root, text="包含子文件夹", variable=self.include_subdirs).pack(
            anchor="w", padx=10, pady=5
        )

        tk.Button(
            root, text="开始加密",
            command=self.start_encrypt, bg="#e1f0d1", height=2,
        ).pack(fill="x", padx=10, pady=10)

        tk.Label(root, text="处理日志：", anchor="w").pack(fill="x", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=17, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.is_running = False

    # ---------------- 日志 ----------------
    def log(self, message):
        self.root.after(0, self._log, message)

    def _log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ---------------- UI 事件 ----------------
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            self.dir_var.set(folder)

    def generate_password(self, length=10):
        alphabet = string.ascii_letters + string.digits + "!@#$%^&*-_"
        password = "".join(secrets.choice(alphabet) for _ in range(length))
        self.pwd_var.set(password)
        self.log(f"[提示] 已生成随机密码：{password}")

    def start_encrypt(self):
        if self.is_running:
            return
        folder = self.dir_var.get().strip()
        password = self.pwd_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("警告", "请先选择有效的文件夹。")
            return
        if not password:
            messagebox.showwarning("警告", "请输入密码或生成随机密码。")
            return
        if any(c in password for c in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']):
            messagebox.showwarning("警告", "密码中包含不推荐的字符，请更换。")
            return

        ok = messagebox.askyesno(
            "确认加密",
            "即将对文件夹中的 Word、Excel、PPT、PDF 文件进行原地加密。\n"
            "已加密的文件将自动跳过。\n\n是否继续？",
        )
        if not ok:
            return

        self.is_running = True
        self.log("=" * 60)
        self.log(f"目标文件夹：{folder}")
        self.log(f"密码：{password}")
        self.log("正在探测引擎...")
        include_subdirs = self.include_subdirs.get()
        threading.Thread(
            target=self._worker,
            args=(folder, password, include_subdirs),
            daemon=True,
        ).start()

    # ---------------- 工作线程 ----------------
    def _worker(self, folder, password, include_subdirs):
        com_word = com_excel = com_ppt = None
        lo_engine = None
        engine_name = None
        max_pwd = 255
        mode = None

        try:
            engine_name, com_word, com_excel, com_ppt, max_pwd = detect_com_engine()
            if engine_name is not None:
                mode = "com"
            else:
                lo_engine = LibreOfficeEngine()
                if not lo_engine.uno_available:
                    self.log("[错误] 无 COM 引擎，且当前 Python 缺少 uno 模块。")
                    self.log("        PDF 以外的文件将无法加密。")
                    mode = "pdf_only"
                else:
                    lo_engine.start()
                    engine_name = "LibreOffice / OpenOffice"
                    max_pwd = 255
                    mode = "lo"

            if engine_name:
                self.log(f"[引擎] {engine_name}（密码上限 {max_pwd}）")
            self.log(
                f"[检测] 魔数检测就绪；msoffcrypto："
                f"{'可用' if MSOFFCRYPTO_AVAILABLE else '不可用（旧格式回退 COM）'}"
            )
            self.log(
                f"[检测] PDF 加密库 pikepdf："
                f"{'可用' if PIKEPDF_AVAILABLE else '不可用'}"
            )

            office_exts = OOXML_EXTS | LEGACY_EXTS | {".pdf"}
            files = [
                f for f in self._collect_files(folder, include_subdirs)
                if os.path.splitext(f)[1].lower() in office_exts
            ]
            if not files:
                self.log("[结束] 未找到可加密的文件。")
                return

            total = len(files)
            encrypted = skipped = errors = 0

            for idx, filepath in enumerate(files, 1):
                ext = os.path.splitext(filepath)[1].lower()
                times = None
                try:
                    times = self._get_times(filepath)

                    # ---------- 第一步：检测是否已加密 ----------
                    if ext == ".pdf":
                        is_enc = self._is_encrypted_pdf(filepath)
                    else:
                        is_enc = self._is_encrypted_office(filepath, ext)
                        if is_enc is None:
                            if mode == "com":
                                is_enc = self._is_encrypted_com(
                                    com_word, com_excel, com_ppt, filepath, ext
                                )
                            elif mode == "lo":
                                is_enc = lo_engine.is_encrypted(filepath)
                            else:
                                is_enc = False

                    if is_enc:
                        self.log(f"[{idx}/{total}] ⏭️ 已加密，跳过：{filepath}")
                        skipped += 1
                        continue

                    if ext != ".pdf" and mode == "pdf_only":
                        self.log(f"[{idx}/{total}] ⏭️ 无可用 Office 引擎，跳过：{filepath}")
                        skipped += 1
                        continue

                    # ---------- 第二步：加密 ----------
                    if ext == ".pdf":
                        status = self._encrypt_pdf(filepath, password)
                    elif mode == "com":
                        status = self._encrypt_com(
                            com_word, com_excel, com_ppt, filepath, ext, password
                        )
                    elif mode == "lo":
                        lo_engine.encrypt(filepath, password)
                        status = "encrypted"
                    else:
                        skipped += 1
                        continue

                    if times:
                        self._restore_times(filepath, times)

                    if status == "encrypted":
                        self.log(f"[{idx}/{total}] ✅ 已加密：{filepath}")
                        encrypted += 1
                    elif status == "already_encrypted":
                        self.log(f"[{idx}/{total}] ⏭️ 已加密，跳过：{filepath}")
                        skipped += 1

                except Exception as e:
                    if times and os.path.exists(filepath):
                        try:
                            self._restore_times(filepath, times)
                        except Exception:
                            pass
                    self.log(f"[{idx}/{total}] ❌ 失败：{filepath}\n    {e}")
                    errors += 1

            self.log("-" * 60)
            self.log(f"完成。成功 {encrypted}，跳过 {skipped}，失败 {errors}。")
            if engine_name:
                self.log(f"Office 引擎：{engine_name}")
            self.log("=" * 60)

        except Exception as e:
            self.log(f"[严重错误] {e}")
        finally:
            for app in (com_word, com_excel, com_ppt):
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            if lo_engine is not None:
                lo_engine.stop()
            if WINDOWS_COM_AVAILABLE:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            self.is_running = False

    # ---------------- 已加密检测 ----------------
    @staticmethod
    def _is_encrypted_pdf(filepath):
        if not PIKEPDF_AVAILABLE:
            return False
        try:
            with pikepdf.open(filepath):
                return False
        except pikepdf.PasswordError:
            return True
        except Exception:
            return False

    @staticmethod
    def _is_encrypted_office(filepath, ext):
        """
        返回 True / False / None。
        - OOXML：用魔数 100% 判断（未加密=ZIP，加密=CFB）。
        - 旧格式：CFB 无法用魔数区分，交给 msoffcrypto；不可用返回 None 让上层回退 COM。
        """
        try:
            with open(filepath, "rb") as f:
                magic = f.read(8)
        except Exception:
            return None

        if magic.startswith(ZIP_MAGIC):
            # 普通 ZIP => 未加密的 OOXML
            return False
        if magic == CFB_MAGIC:
            if ext in OOXML_EXTS:
                # OOXML 被包成 CFB => 已加密
                return True
            # 旧格式（.doc/.xls/.ppt）正常就是 CFB，继续往下用 msoffcrypto

        if MSOFFCRYPTO_AVAILABLE:
            try:
                with open(filepath, "rb") as f:
                    of = msoffcrypto.OfficeFile(f)
                    return bool(of.is_encrypted())
            except Exception:
                return None

        return None

    def _is_encrypted_com(self, word, excel, ppt, filepath, ext):
        """COM 回退检测（旧格式且 msoffcrypto 不可用时）。"""
        if ext == ".doc":
            if word is None:
                return False
            doc = None
            try:
                doc = word.Documents.Open(
                    filepath, ReadOnly=True, PasswordDocument="",
                    AddToRecentFiles=False, Visible=False,
                )
                return bool(getattr(doc, "HasPassword", False))
            except Exception as e:
                return self._is_pwd_err(e)
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

        if ext == ".xls":
            if excel is None:
                return False
            wb = None
            try:
                wb = excel.Workbooks.Open(
                    filepath, ReadOnly=True, Password="", UpdateLinks=0,
                )
                return bool(getattr(wb, "HasPassword", False))
            except Exception as e:
                return self._is_pwd_err(e)
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass

        if ext == ".ppt":
            if ppt is None:
                return False
            pres = None
            try:
                pres = ppt.Presentations.Open(
                    filepath, WithWindow=False, ReadOnly=True,
                )
                return bool(pres.Password)
            except Exception as e:
                return self._is_pwd_err(e)
            finally:
                if pres is not None:
                    try:
                        pres.Close()
                    except Exception:
                        pass

        return False

    # ---------------- 加密实现 ----------------
    def _encrypt_com(self, word, excel, ppt, filepath, ext, password):
        if ext in (".doc", ".docx", ".docm"):
            if word is None:
                raise RuntimeError("无文字组件")
            return self._encrypt_word(word, filepath, password)
        elif ext in (".xls", ".xlsx", ".xlsm"):
            if excel is None:
                raise RuntimeError("无表格组件")
            return self._encrypt_excel(excel, filepath, password)
        elif ext in (".ppt", ".pptx", ".pptm"):
            if ppt is None:
                raise RuntimeError("无演示组件")
            return self._encrypt_ppt(ppt, filepath, password)
        else:
            raise RuntimeError(f"不支持的扩展名：{ext}")

    def _encrypt_word(self, app, filepath, password):
        doc = None
        try:
            doc = app.Documents.Open(
                filepath, ReadOnly=False, AddToRecentFiles=False, Visible=False,
            )
            doc.Password = password
            doc.Save()
            doc.Close(SaveChanges=False)
            doc = None
            return "encrypted"
        except Exception as e:
            if self._is_pwd_err(e):
                return "already_encrypted"
            raise
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    def _encrypt_excel(self, app, filepath, password):
        wb = None
        try:
            wb = app.Workbooks.Open(filepath, ReadOnly=False, UpdateLinks=0)
            wb.SaveAs(filepath, None, password)
            wb.Close(SaveChanges=False)
            wb = None
            return "encrypted"
        except Exception as e:
            if self._is_pwd_err(e):
                return "already_encrypted"
            raise
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def _encrypt_ppt(self, app, filepath, password):
        pres = None
        try:
            pres = app.Presentations.Open(
                filepath, WithWindow=False, ReadOnly=False,
            )
            pres.Password = password
            pres.Save()
            pres.Close()
            pres = None
            return "encrypted"
        except Exception as e:
            if self._is_pwd_err(e):
                return "already_encrypted"
            raise
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:
                    pass

    def _encrypt_pdf(self, filepath, password):
        if not PIKEPDF_AVAILABLE:
            raise RuntimeError("pikepdf 未安装，无法加密 PDF。请执行：pip install pikepdf")
        temp = filepath + ".enc_tmp.pdf"
        try:
            with pikepdf.open(filepath) as pdf:
                pdf.save(
                    temp,
                    encryption=pikepdf.Encryption(user=password, owner=password, R=6),
                )
            os.replace(temp, filepath)
        finally:
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
        return "encrypted"

    # ---------------- 工具 ----------------
    @staticmethod
    def _is_pwd_err(exc):
        """
        COM 回退路径下的密码错误判断。优先 hresult，文本兜底仅保留精确关键词，
        避免把"工作表保护"等误判为文件加密。
        """
        if WINDOWS_COM_AVAILABLE:
            try:
                import pywintypes
                if isinstance(exc, pywintypes.com_error):
                    hresult = exc.hresult & 0xFFFFFFFF
                    if hresult in (
                        0x800A1521,  # Word: 密码错误 / 文档受保护
                        0x800A1520,  # Word: 受保护
                        0x800A1522,
                        0x800A1533,
                        0x800A03EC,  # Excel: 通用错误（加密文件打开常见）
                    ):
                        return True
            except Exception:
                pass

        m = str(exc).lower()
        return any(k in m for k in (
            "password", "密码", "文档受保护", "document is protected",
            "is password protected",
        ))

    def _collect_files(self, folder, include_subdirs):
        result = []
        if include_subdirs:
            for root, _, names in os.walk(folder):
                for n in names:
                    result.append(os.path.join(root, n))
        else:
            for n in os.listdir(folder):
                p = os.path.join(folder, n)
                if os.path.isfile(p):
                    result.append(p)
        return result

    @staticmethod
    def _get_times(filepath):
        st = os.stat(filepath)
        return st.st_atime, st.st_mtime, st.st_ctime

    def _restore_times(self, filepath, times):
        atime, mtime, ctime = times
        try:
            os.utime(filepath, (atime, mtime))
        except Exception:
            pass

        if not WINDOWS_COM_AVAILABLE:
            return

        try:
            h = win32file.CreateFile(
                filepath, win32con.GENERIC_WRITE,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None,
            )
            try:
                win32file.SetFileTime(h, PyTime(ctime), PyTime(atime), PyTime(mtime))
            finally:
                win32file.CloseHandle(h)
        except Exception as e:
            self.log(f"    [警告] 恢复创建时间失败：{e}")


def main():
    root = tk.Tk()
    MultiEngineEncryptorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()