# -*- coding: utf-8 -*-
r"""Generic Origin plotting driver for this machine (verified COM flow).

Writes a worksheet into the already-open Origin, plots it, styles the axes,
exports a PNG via LabTalk expGraph, and verifies the file.

Usage:
    python origin_plot.py --data FILE --x-col 1 --y-cols 2,3 --plot-type bar \
        --out C:\OriginData\fig.png --xlabel "Category" --ylabel "Value"

Data file: whitespace/comma/tab separated text, no header required, columns 1-based.
Does NOT save .opju - press Ctrl+S in Origin to keep the figure.
"""
import argparse
import os
import sys
import time
import traceback

import pythoncom
import win32com.client

PLOT_IDS = {"line": 201, "line+symbol": 202, "scatter": 200, "bar": 203}


def ex(app, cmd):
    try:
        return app.Execute(cmd)
    except Exception as e:
        raise RuntimeError("[FAIL] LabTalk %r\n%s" % (cmd, e))


def num(app, expr):
    try:
        return float(app.Evaluate(expr))
    except Exception:
        return None


def color_num(app, rgb):
    v = num(app, "color(%d,%d,%d)" % rgb)
    if v is None:
        raise RuntimeError("color(%s) failed" % (rgb,))
    return int(v)


def connect():
    app = None
    for pid in ("Origin.ApplicationSI", "Origin.Application"):
        try:
            app = win32com.client.GetActiveObject(pid)
            break
        except Exception:
            pass
    if app is None:
        for pid in ("Origin.ApplicationSI", "Origin.Application"):
            try:
                app = win32com.client.Dispatch(pid)
                break
            except Exception:
                pass
    if app is None:
        raise RuntimeError("[FAIL] cannot connect Origin")
    try:
        app.Visible = True
    except Exception:
        pass
    return app


def read_table(path):
    rows = []
    with open(path, "r", encoding="utf-8-sig") as f:
        for line in f:
            s = line.strip()
            if not s or s[0] in "*#":
                continue
            parts = s.replace(",", " ").replace("\t", " ").split()
            vals = []
            for p in parts:
                try:
                    vals.append(float(p))
                except ValueError:
                    vals = None
                    break
            if vals:
                rows.append(vals)
    if not rows:
        raise RuntimeError("[FAIL] no numeric rows in %s" % path)
    ncol = min(len(r) for r in rows)
    return [r[:ncol] for r in rows]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--x-col", type=int, default=1, help="1-based x column")
    ap.add_argument("--y-cols", required=True, help="comma list, e.g. 2,3,4")
    ap.add_argument("--plot-type", default="line", choices=sorted(PLOT_IDS))
    ap.add_argument("--out", required=True, help="PNG output path")
    ap.add_argument("--xlabel", default="X")
    ap.add_argument("--ylabel", default="Y")
    ap.add_argument("--legend", default="auto", choices=["auto", "on", "off"])
    ap.add_argument("--width", type=float, default=1.6)
    args = ap.parse_args()

    ycols = [int(c) for c in args.y_cols.split(",") if c.strip()]
    if not os.path.isfile(args.data):
        raise RuntimeError("[FAIL] missing data: %s" % args.data)
    rows = read_table(args.data)
    xcol = args.x_col
    keep = [xcol] + ycols
    for c in keep:
        if c < 1 or c > len(rows[0]):
            raise RuntimeError("[FAIL] column %d out of range (have %d)" % (c, len(rows[0])))
    table = [[r[c - 1] for c in keep] for r in rows]

    out = os.path.abspath(args.out)
    outdir = os.path.dirname(out)
    outname = os.path.splitext(os.path.basename(out))[0]

    pythoncom.CoInitialize()
    try:
        app = connect()
        ex(app, "newbook;")
        time.sleep(1.0)
        ex(app, "wks.ncols=%d;" % len(keep))
        page = str(app.WorksheetPages.Item(int(app.WorksheetPages.Count) - 1).Name)
        app.PutWorksheet(page, table, 0, 0)
        time.sleep(1.5)
        ex(app, "wks.col1.type=1;")
        for i in range(2, len(keep) + 1):
            ex(app, "wks.col%d.type=2;" % i)
        nrows = num(app, "wks.nrows")
        if nrows is None or int(nrows) != len(table):
            raise RuntimeError("[FAIL] nrows mismatch: %s vs %d" % (nrows, len(table)))

        before = int(app.GraphPages.Count)
        yspec = "%d:%d" % (2, len(keep)) if len(keep) > 2 else "2"
        ex(app, "plotxy iy:=(1,%s) plot:=%d;" % (yspec, PLOT_IDS[args.plot_type]))
        time.sleep(2.5)
        after = int(app.GraphPages.Count)
        if after <= before:
            raise RuntimeError("[FAIL] no new graph")
        gp = app.GraphPages.Item(after - 1)
        gname = str(gp.Name)
        nplots = int(gp.Layers.Item(0).DataPlots.Count)
        ex(app, 'win -a "%s";' % gname)

        for i in range(1, nplots + 1):
            ex(app, "layer -p %d;" % i)
            ex(app, "layer.plot%d.line.width=%.2f;" % (i, args.width))

        if args.legend == "off" or (args.legend == "auto" and nplots < 2):
            ex(app, "legend.show=0;")
        else:
            ex(app, "legend.show=1;")

        ex(app, 'xb.text$="%s";' % args.xlabel)
        ex(app, 'yl.text$="%s";' % args.ylabel)
        for t in ("xb", "yl"):
            ex(app, "%s.font=font(Times New Roman);" % t)
            ex(app, "%s.fsize=20;" % t)

        os.makedirs(outdir, exist_ok=True)
        if os.path.isfile(out):
            os.remove(out)
        ex(app, 'win -a "%s";' % gname)
        time.sleep(0.5)
        ex(app, 'expGraph type:=png path:="%s" filename:="%s" overwrite:=1;'
           % (outdir, outname))
        time.sleep(3.0)
        if not os.path.isfile(out) or os.path.getsize(out) < 1000:
            raise RuntimeError("[FAIL] PNG not created or too small: %s" % out)

        print("[OK] graph=%s plots=%d rows=%d" % (gname, nplots, len(table)))
        print("[OK] PNG=%s (%d bytes)" % (out, os.path.getsize(out)))
        print("[OK] .opju NOT saved - press Ctrl+S in Origin to keep it.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
