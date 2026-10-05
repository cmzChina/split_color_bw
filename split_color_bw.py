# -*- coding: utf-8 -*-
"""
=========================================================
   同页黑白 / 彩色内容分离工具  v2.0   (无界面，双击即用)
=========================================================
v2.0 新增「矢量无损模式」：不把页面渲染成图片，直接按对象颜色删除，
     输出的 PDF 仍是矢量，文字可复制、无限放大不模糊。

用法：放进 存放 PDF / Word 的文件夹双击，或把文件夹/文件拖到程序图标上。
输出：export/xxx_黑白.pdf、export/xxx_彩色.pdf、export/时间_统计报告.csv
"""

import os
import io
import sys
import csv
import time
import shutil
import datetime
import tempfile
import subprocess

# ==================== 可调参数 ====================
MODE     = "auto"   # "vector" 矢量无损 | "raster" 像素模式(旧) | "auto" 自动判断
DPI      = 300      # 仅 raster 模式 / 扫描页 生效
SAT_THR  = 0.12     # 饱和度阈值(0~1)。漏判彩色 -> 调小到 0.08
DIFF_THR = 18       # 通道差阈值(0~255)。发黄扫描件全页判彩色 -> 调到 30
CLEAN    = 5        # 边缘清理(奇数 3/5/7/9)，仅像素模式
DENOISE  = 0        # 去噪点(奇数 3/5，0=关闭)，仅像素模式
EXPORT   = "export"
RECURSIVE = True

SCAN_RATIO      = 0.80  # auto：单页被图片覆盖超 80% -> 视为扫描页，该页走像素模式
IMG_COLOR_RATIO = 0.02  # 矢量模式：图中彩色像素超 2% -> 判为彩色图像
# =================================================

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
EXTS = (".pdf", ".doc", ".docx")
IS_FROZEN = getattr(sys, "frozen", False)


def get_base_dir():
    if IS_FROZEN:
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ---------------- 依赖 ----------------
def pip_install(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg,
                           "-i", MIRROR, "--disable-pip-version-check"],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def require(pkg, imp):
    try:
        __import__(imp)
    except ImportError:
        pip_install(pkg)
        __import__(imp)


def setup_deps():
    if IS_FROZEN:
        print("[1/3] 运行模式：EXE（依赖已内置）")
        return
    print("[1/3] 检查运行环境 ...")
    for pkg, imp in (("pymupdf", "pymupdf"), ("numpy", "numpy"), ("pillow", "PIL")):
        require(pkg, imp)
    try:
        require("pywin32", "win32com.client")
    except Exception:
        print("    [提示] pywin32 不可用，Word 将尝试用 LibreOffice")
    print("    环境就绪")


def get_fitz():
    try:
        import fitz
    except ImportError:
        import pymupdf as fitz
    return fitz


def rc(name, default):
    """兼容不同 PyMuPDF 版本的 redaction 常量。"""
    return getattr(get_fitz(), name, default)


# ---------------- 颜色判定 ----------------
def to_rgb255(c):
    """sRGB 整数 或 (r,g,b) 浮点元组 -> (R,G,B) 0~255"""
    fitz = get_fitz()
    if isinstance(c, (tuple, list)):
        v = list(c)[:3]
    else:
        try:
            v = list(fitz.sRGB_to_rgb(int(c)))
        except Exception:
            v = [0, 0, 0]
    return tuple(int(round(max(0.0, min(1.0, x)) * 255)) for x in v)


def is_color(rgb):
    mx, mn = max(rgb), min(rgb)
    if mx <= 0:
        return False
    diff = mx - mn
    return diff > DIFF_THR and (diff / mx) > SAT_THR


def image_is_color(doc, xref):
    """采样图片像素，判断是否彩色图。"""
    import numpy as np
    from PIL import Image
    try:
        raw = doc.extract_image(xref)["image"]
        img = Image.open(io.BytesIO(raw)).convert("RGB")
    except Exception:
        return False
    img.thumbnail((400, 400))
    arr = np.asarray(img).astype(np.int16)
    mx, mn = arr.max(axis=2), arr.min(axis=2)
    diff = mx - mn
    sat = np.where(mx > 0, diff / np.maximum(mx, 1), 0.0)
    return float(((diff > DIFF_THR) & (sat > SAT_THR)).mean()) > IMG_COLOR_RATIO


# ---------------- 收集页面对象（矢量模式核心） ----------------
def collect_objects(doc, page):
    """返回 (彩色对象矩形, 黑白对象矩形, 图片覆盖率)"""
    fitz = get_fitz()
    color_rects, bw_rects = [], []
    page_area = abs(page.rect.get_area()) or 1.0
    img_area = 0.0

    # 1) 文字：span["color"] 是 sRGB 整数
    try:
        d = page.get_text("dict")
        for b in d.get("blocks", []):
            for ln in b.get("lines", []):
                for sp in ln.get("spans", []):
                    if not (sp.get("text") or "").strip():
                        continue
                    r = fitz.Rect(sp["bbox"])
                    if abs(r.get_area()) <= 0:
                        continue
                    (color_rects if is_color(to_rgb255(sp.get("color", 0)))
                     else bw_rects).append(r)
    except Exception:
        pass

    # 2) 矢量图形：path 的 fill / color 是 (r,g,b) 浮点
    try:
        for dr in page.get_drawings():
            cols = [c for c in (dr.get("fill"), dr.get("color")) if c]
            r = fitz.Rect(dr["rect"])
            if not cols or abs(r.get_area()) <= 0:
                continue
            (color_rects if any(is_color(to_rgb255(c)) for c in cols)
             else bw_rects).append(r)
    except Exception:
        pass

    # 3) 图片
    try:
        for im in page.get_image_info():
            r = fitz.Rect(im["bbox"])
            if abs(r.get_area()) <= 0:
                continue
            img_area += abs(r.get_area())
            (color_rects if image_is_color(doc, im["xref"])
             else bw_rects).append(r)
    except Exception:
        pass

    return color_rects, bw_rects, img_area / page_area


# ---------------- 物理删除 ----------------
def apply_erase(page, rects):
    """用 redaction 真正删除指定区域内的文字/图形/图片。"""
    fitz = get_fitz()
    for r in rects:
        try:
            page.add_redact_annot(fitz.Rect(r), fill=False)
        except Exception:
            pass
    page.apply_redactions(
        images=rc("PDF_REDACT_IMAGE_REMOVE", 1),
        graphics=rc("PDF_REDACT_LINE_ART_REMOVE", 1),
        text=rc("PDF_REDACT_TEXT_REMOVE", 1))


# ---------------- 模式 A：矢量无损 ----------------
def page_vector(pb, pc, c_rects, b_rects):
    ca = sum(abs(get_fitz().Rect(r).get_area()) for r in c_rects)
    ba = sum(abs(get_fitz().Rect(r).get_area()) for r in b_rects)
    apply_erase(pb, c_rects)   # 黑白页：删彩色
    apply_erase(pc, b_rects)   # 彩色页：删黑白
    return (ca / (ca + ba) * 100.0) if (ca + ba) > 0 else 0.0


# ---------------- 模式 B：像素（保留旧逻辑，扫描件用） ----------------
def build_color_mask(arr):
    import numpy as np
    from PIL import Image, ImageFilter
    mx, mn = arr.max(axis=2), arr.min(axis=2)
    diff = mx - mn
    sat = np.where(mx > 0, diff / np.maximum(mx, 1), 0.0)
    mask = (diff > DIFF_THR) & (sat > SAT_THR)
    m = Image.fromarray(mask.astype(np.uint8) * 255)
    if CLEAN and CLEAN > 1:
        m = m.filter(ImageFilter.MaxFilter(int(CLEAN)))
    if DENOISE and DENOISE > 1:
        m = m.filter(ImageFilter.MinFilter(int(DENOISE)))
        m = m.filter(ImageFilter.MaxFilter(int(DENOISE)))
    return np.array(m) > 127


def page_raster(pb, pc):
    fitz = get_fitz()
    import numpy as np
    from PIL import Image
    mat = fitz.Matrix(DPI / 72.0, DPI / 72.0)
    pix = pb.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    arr = np.asarray(img).astype(np.int16)
    mask = build_color_mask(arr)[..., None]
    arr8 = arr.astype(np.uint8)
    pairs = ((pb, np.where(mask, 255, arr8).astype(np.uint8)),
             (pc, np.where(mask, arr8, 255).astype(np.uint8)))
    rect = pb.rect
    for page, out in pairs:
        buf = io.BytesIO()
        Image.fromarray(out).save(buf, format="PNG")
        apply_erase(page, [page.rect])
        page.insert_image(fitz.Rect(0, 0, rect.width, rect.height),
                          stream=buf.getvalue())
    return float(mask.mean()) * 100.0


# ---------------- 处理一个 PDF ----------------
def process_document(src, out_bw, out_color):
    fitz = get_fitz()
    doc_bw = fitz.open(src)
    doc_color = fitz.open(src)
    try:
        total = doc_bw.page_count
        stats, modes = [], {}

        for i in range(total):
            pb, pc = doc_bw[i], doc_color[i]
            if MODE == "raster":
                st, tag = page_raster(pb, pc), "像素"
            else:
                c_rects, b_rects, ratio = collect_objects(doc_bw, pb)
                if MODE == "auto" and ratio > SCAN_RATIO:
                    st, tag = page_raster(pb, pc), "像素(扫描页)"
                else:
                    st, tag = page_vector(pb, pc, c_rects, b_rects), "矢量"
            stats.append(st)
            modes[tag] = modes.get(tag, 0) + 1
            sys.stdout.write("\r    进度: %d/%d 页 [%s]" % (i + 1, total, tag))
            sys.stdout.flush()
        sys.stdout.write("\n")

        doc_bw.save(out_bw, garbage=4, deflate=True)
        doc_color.save(out_color, garbage=4, deflate=True)
        return stats, modes
    finally:
        doc_bw.close()
        doc_color.close()


# ---------------- Word 转 PDF ----------------
def word_to_pdf(src, out_dir):
    src_abs = os.path.abspath(src)
    out_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    co_initialized = False
    try:
        import win32com.client as win32
        try:
            import pythoncom
            pythoncom.CoInitialize()
            co_initialized = True
        except Exception:
            pass
        for app_name in ("KWps.Application", "Word.Application"):
            app = None
            doc = None
            try:
                # DispatchEx 始终创建新实例，避免连接到用户已运行的窗口并将其隐藏
                app = win32.DispatchEx(app_name)
            except Exception:
                continue
            try:
                app.Visible = False
                doc = app.Documents.Open(src_abs, ReadOnly=True)
                doc.SaveAs2(out_pdf, FileFormat=17)
                doc.Close(False)
                doc = None
                app.Quit()
                app = None
                if os.path.exists(out_pdf):
                    return out_pdf
            except Exception:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
    except Exception:
        pass
    finally:
        if co_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
    for cand in (shutil.which("soffice"),
                 r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if cand and os.path.exists(cand):
            subprocess.run([cand, "--headless", "--convert-to", "pdf",
                            "--outdir", out_dir, src_abs],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.exists(out_pdf):
                return out_pdf
    return None


def unique_path(p):
    if not os.path.exists(p):
        return p
    stem, ext = os.path.splitext(p)
    n = 1
    while os.path.exists("%s(%d)%s" % (stem, n, ext)):
        n += 1
    return "%s(%d)%s" % (stem, n, ext)


def collect_files(base):
    found = []
    if RECURSIVE:
        for root, dirs, names in os.walk(base):
            dirs[:] = [d for d in dirs
                       if d.lower() != EXPORT.lower() and not d.startswith(".")]
            for name in names:
                low = name.lower()
                if name.startswith("~$"):
                    continue
                if low.endswith(("_黑白.pdf", "_彩色.pdf")):
                    continue
                if low.endswith(EXTS):
                    p = os.path.join(root, name)
                    found.append((p, os.path.relpath(p, base)))
    else:
        for name in sorted(os.listdir(base)):
            low = name.lower()
            p = os.path.join(base, name)
            if os.path.isdir(p) or name.startswith("~$"):
                continue
            if low.endswith(("_黑白.pdf", "_彩色.pdf")):
                continue
            if low.endswith(EXTS):
                found.append((p, name))
    return sorted(found, key=lambda x: x[1])


def report_empty(base):
    print("    ! 在下面这个目录里没有找到任何 PDF / Word 文件：")
    print("      %s" % base)
    try:
        items = sorted(os.listdir(base))
    except Exception as e:
        print("      （无法读取：%s）" % e)
        return
    print("    ! 该目录实际内容（前 60 项，[D] 表示文件夹）：")
    if not items:
        print("      （空文件夹）")
    for n in items[:60]:
        try:
            print("      %s%s" % ("[D] " if os.path.isdir(os.path.join(base, n)) else "    ", n))
        except Exception:
            pass
    print("    ! 也可把 PDF/Word 所在文件夹直接拖到本程序图标上")


# ---------------- 主流程 ----------------
def main():
    if len(sys.argv) > 1:
        target = os.path.abspath(sys.argv[1])
        base = os.path.dirname(target) if os.path.isfile(target) else target
    else:
        base = get_base_dir()

    out_dir = os.path.join(base, EXPORT)
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="bwcolor_")
    try:
        print("[2/3] 扫描目录：%s （模式：%s）" % (base, MODE))
        items = collect_files(base)
        if not items:
            report_empty(base)
            return

        print("      共发现 %d 个文件" % len(items))
        print("[3/3] 开始处理 ...")

        all_stats = []
        for path, rel in items:
            stem = os.path.splitext(rel)[0].replace(os.sep, "_").replace("/", "_")
            print("  -> %s" % rel)
            work = path
            if path.lower().endswith((".doc", ".docx")):
                work = word_to_pdf(path, tmp)
                if not work:
                    print("     [跳过] Word 转 PDF 失败")
                    continue
            t0 = time.time()
            try:
                out_bw = unique_path(os.path.join(out_dir, stem + "_黑白.pdf"))
                out_color = unique_path(os.path.join(out_dir, stem + "_彩色.pdf"))
                stats, modes = process_document(work, out_bw, out_color)
                avg = sum(stats) / len(stats) if stats else 0
                all_stats.append((rel, len(stats), avg, stats))
                print("     完成 %.1f 秒，%d 页 [%s]，平均彩色占比 %.1f%%"
                      % (time.time() - t0, len(stats),
                         " ".join("%s×%d" % (k, v) for k, v in modes.items()), avg))
            except Exception as e:
                print("     [失败] %s" % e)

        if all_stats:
            now = datetime.datetime.now()
            rep = unique_path(os.path.join(
                out_dir, now.strftime("%Y年%m月%d日%H时%M分") + "_统计报告.csv"))
            with open(rep, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["生成时间", now.strftime("%Y-%m-%d %H:%M:%S")])
                w.writerow(["扫描目录", base])
                w.writerow(["处理模式", MODE])
                w.writerow([])
                w.writerow(["文件名", "页码", "彩色占比%", "黑白占比%"])
                for rel, _, _, stats in all_stats:
                    for i, r in enumerate(stats, 1):
                        w.writerow([rel, i, "%.2f" % r, "%.2f" % (100 - r)])
            print("\n统计报告：%s" % rep)

        print("全部完成，结果见 %s 文件夹。" % EXPORT)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    try:
        setup_deps()
        main()
    except Exception as e:
        print("\n出错了：%r" % e)
        if not IS_FROZEN:
            print("缺库可执行：pip install pymupdf numpy pillow pywin32 -i %s" % MIRROR)
    input("\n按回车键退出...")