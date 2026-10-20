#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Probe / render PDFs for the 租金收缴专项核查 workflow.

Usage:
    python render_pdf_pages.py <pdf> [--out DIR] [--scale 2.78] [--pages 1-8,12] [--prefix P]
    python render_pdf_pages.py <pdf> --text            # dump text layer
    python render_pdf_pages.py <pdf|dir> --probe [--json]
                                                       # 逐份体检：ok / no_text_layer / encrypted / broken

Probe 输出供底稿"资料清单"与"疑点清单"直接引用：
    - encrypted  → 是否提供填 √，备注写"文件加密损坏，内容线下人工核验原件"，疑点清单加一条
    - broken     → 同上
    - no_text_layer → 扫描件，渲染 PNG 后读图

Notes:
    - scale ~2.78 => ~200 DPI (72 * 2.78).
    - Chinese paths: pass via env var or quote properly on Windows PowerShell.
    - For a 25MB+ report, render only the pages you need (use --pages).
"""
import argparse
import json
import os
import sys


def parse_pages(spec, n):
    if not spec:
        return list(range(n))
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    return [i for i in out if 0 <= i < n]


def _is_encrypted(err):
    s = str(err).lower()
    return any(k in s for k in ("password", "encrypt", "security", "decrypt"))


def probe_one(path):
    """返回 dict(file, status, pages, encrypted, note)。status ∈ ok/no_text_layer/encrypted/broken。"""
    rec = {"file": os.path.basename(path), "path": path,
           "status": "ok", "pages": None, "encrypted": False, "note": ""}
    try:
        import pypdfium2 as pdfium
    except Exception as e:  # pragma: no cover
        rec.update(status="broken", note="pypdfium2 不可用：%s" % e)
        return rec
    pdf = None
    try:
        pdf = pdfium.PdfDocument(path)
        n = len(pdf)
        rec["pages"] = n
        txt = ""
        if n:
            try:
                txt = pdf[0].get_textpage().get_text_range() or ""
            except Exception as e:
                if _is_encrypted(e):
                    rec.update(status="encrypted", encrypted=True, note=str(e))
                    return rec
                txt = ""
        if not txt.strip():
            rec["status"] = "no_text_layer"
            rec["note"] = "无文字层，需渲染 PNG 读图"
    except Exception as e:
        if _is_encrypted(e):
            rec.update(status="encrypted", encrypted=True, note=str(e))
        else:
            rec.update(status="broken", note="%s: %s" % (type(e).__name__, e))
    finally:
        try:
            if pdf is not None:
                pdf.close()
        except Exception:
            pass
    return rec


def collect(path):
    """传入单个文件或目录（递归），返回 pdf 路径列表。"""
    if os.path.isfile(path):
        return [path]
    if not os.path.isdir(path):
        print("ERROR: not found:", path, file=sys.stderr)
        sys.exit(2)
    out = []
    for root, _dirs, files in os.walk(path):
        for f in files:
            if f.lower().endswith(".pdf"):
                out.append(os.path.join(root, f))
    return sorted(out)


def do_probe(path, as_json):
    files = collect(path)
    if not files:
        print("ERROR: no PDF found under:", path, file=sys.stderr)
        sys.exit(2)
    recs = [probe_one(p) for p in files]
    if as_json:
        print(json.dumps(recs, ensure_ascii=False, indent=2))
    else:
        print("%-4s %-14s %-6s %-4s %s" % ("#", "status", "pages", "enc", "file"))
        for i, r in enumerate(recs, 1):
            print("%-4d %-14s %-6s %-4s %s"
                  % (i, r["status"], r["pages"] if r["pages"] is not None else "-",
                     "Y" if r["encrypted"] else "", r["file"]))
            if r["note"]:
                print("     note: %s" % r["note"])
    bad = [r for r in recs if r["status"] in ("encrypted", "broken")]
    print("\n合计 %d 份：ok=%d / 无文字层=%d / 加密=%d / 损坏=%d"
          % (len(recs),
             sum(1 for r in recs if r["status"] == "ok"),
             sum(1 for r in recs if r["status"] == "no_text_layer"),
             sum(1 for r in recs if r["status"] == "encrypted"),
             sum(1 for r in recs if r["status"] == "broken")), file=sys.stderr)
    if bad:
        print("以下文件须在底稿备注写「文件加密损坏，内容线下人工核验原件」，并在疑点清单留痕：",
              file=sys.stderr)
        for r in bad:
            print("  - %s (%s)" % (r["file"], r["status"]), file=sys.stderr)


def main():
    ap = argparse.ArgumentParser(description="Probe or render PDF pages.")
    ap.add_argument("pdf")
    ap.add_argument("--out", default=None, help="output dir (default ./_scan)")
    ap.add_argument("--scale", type=float, default=2.78, help="render scale, 2.78 ~= 200DPI")
    ap.add_argument("--pages", default=None, help="e.g. 1-8,12 (1-based)")
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--text", action="store_true", help="dump text layer instead of rendering")
    ap.add_argument("--probe", action="store_true",
                    help="只体检：逐份报告 ok/no_text_layer/encrypted/broken")
    ap.add_argument("--json", action="store_true", help="与 --probe 搭配，输出 JSON")
    args = ap.parse_args()

    pdf_path = os.environ.get("PDF", args.pdf)
    if not os.path.exists(pdf_path):
        d, base = os.path.split(pdf_path)
        if os.path.isdir(d):
            for name in os.listdir(d):
                if name.strip() == base.strip():
                    pdf_path = os.path.join(d, name)
                    break

    if args.probe:
        do_probe(pdf_path, args.json)
        return

    if not os.path.exists(pdf_path):
        print("ERROR: file not found:", pdf_path, file=sys.stderr)
        sys.exit(2)

    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(pdf_path)
    n = len(pdf)
    idx = parse_pages(args.pages, n)
    print("pages total=%d, selected=%d" % (n, len(idx)), file=sys.stderr)

    if args.text:
        for i in idx:
            txt = pdf[i].get_textpage().get_text_range()
            print("===== page %d (chars=%d) =====" % (i + 1, len(txt)))
            print(txt[:3000])
        return

    out = args.out or os.path.join(os.getcwd(), "_scan")
    os.makedirs(out, exist_ok=True)
    prefix = args.prefix or os.path.splitext(os.path.basename(pdf_path))[0]
    for i in idx:
        img = pdf[i].render(scale=args.scale).to_pil()
        p = os.path.join(out, "%s_%02d.png" % (prefix, i + 1))
        img.save(p)
        print(p)


if __name__ == "__main__":
    main()
