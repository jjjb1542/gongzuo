#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fill the 《房产土地租金收缴专项核查底稿》 from a JSON spec.

Usage:
    python fill_workingpaper.py spec.json
    python fill_workingpaper.py --example        # print a template spec

The script NEVER edits the template in place: it copies `template` to `output`
and fills the copy. `{SRC}` inside any string is replaced by `src_ref`.

Spec schema (all keys optional except template/output):
{
  "template": "<path>/workingpaper-template.xlsx",
  "output":   "<path>/…底稿_<标的>.xlsx",
  "unit":     "被审计单位（取台账『权属公司』列）",
  "baseline": "2026-09-30",          // 清查基准日（入参，写入 5 张表的抬头）
  "exec":     "审计组",
  "date":     "2026-10-16",
  "src_ref":  "成都香城产业发展集团有限公司8.12_副本.xlsx → sheet「房产」",
  "material_list": { "8": ["√", "备注 {SRC} 第56行（标的）", "原件扫描件"] },
                     // 三元组：是否提供 / 备注 / 资料载体类型（原件扫描件|复印件拍照件|截图|电子文档）
  "procedures":    { "8": ["《审计资料清单》", "审计组", "2026-10-16", null],
                     "10": ["《合同与财务对账差异明细表》", "审计组", "2026-10-16",
                            "【本项目执行情况】…"] },
  "ledger": {
     "common": {"B": "成都香城产业发展集团有限公司", "E": "项目名", "H": "园区外", ...},
     "rows":   [ {"A": 1, "D": "标的", "F": 132.79, "Q": 65333, ...}, ... ],
     "first_row": 9,
     "fmt_cols": [17, 18, 19, 20, 24, 27]
  },
  "reconciliation": {
     "rows": [ {"A": 1, "C": "位置", "D": "承租方", "E": 65333, "I": "描述 {SRC} 第56行。",
                "J": 1}],                       // J=对应台账序号，用于 Q≠X 差异登记核销
     "first_row": 9,
     "note_target": {"anchor": "A14", "merge": "A14:I21"}   // required if data reaches row 11+
  },
  "issues": { "title": "…疑点清单（客观列示，不作定性）",
              "items": [ ["1", "疑点", "事实依据", "核查建议", "资料充分性"], ... ] }
}
"""
import argparse
import json
import os
import re
import shutil
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# 欠租统计台账 列映射（A..AH 共 34 列）
LEDGER_COL = {
    "A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8, "I": 9, "J": 10,
    "K": 11, "L": 12, "M": 13, "N": 14, "O": 15, "P": 16, "Q": 17, "R": 18, "S": 19,
    "T": 20, "U": 21, "V": 22, "W": 23, "X": 24, "Y": 25, "Z": 26,
    "AA": 27, "AB": 28, "AC": 29, "AD": 30, "AE": 31, "AF": 32, "AG": 33, "AH": 34,
}

SHEET_MATERIAL = "审计资料清单"
SHEET_PROCEDURE = "审计程序表"
SHEET_LEDGER = "欠租统计台账"
SHEET_RECON = "合同与财务对账差异明细表"
SHEET_ISSUES = "疑点清单"

# 通用列字母→索引：A~AH 全覆盖（台账用 A~AH，对账明细表用 A~J，均为其子集）
COL = dict(LEDGER_COL)

BASE_RX = re.compile(r"（\s*核查基准日：[^）]*）")
DEFAULT_BASELINE = "2026-09-30"


def fmt_baseline(b):
    """把 2026-09-30 / 2026/9/30 / 2026年9月30日 统一为 2026年9月30日。"""
    s = str(b or "").strip()
    m = re.match(r"^(\d{4})\s*[-/年.]\s*(\d{1,2})\s*[-/月.]\s*(\d{1,2})", s)
    if not m:
        return s
    return "%d年%d月%d日" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))


def col(k):
    """Accept int or column letter."""
    return k if isinstance(k, int) else COL[k]


class Spec(object):
    def __init__(self, d):
        self.d = d
        self.src_ref = d.get("src_ref", "")
        self.baseline = d.get("baseline") or DEFAULT_BASELINE
        self.baseline_disp = fmt_baseline(self.baseline)

    def s(self, text):
        if text is None:
            return None
        return str(text).replace("{SRC}", self.src_ref)


def copy_template(spec):
    tpl, out = spec.d["template"], spec.d["output"]
    if not os.path.exists(tpl):
        raise SystemExit("template not found: %s" % tpl)
    if os.path.abspath(tpl) == os.path.abspath(out):
        raise SystemExit("refuse to overwrite the template in place")
    shutil.copyfile(tpl, out)
    return openpyxl.load_workbook(out)


def set_units(wb, spec):
    """写被审计单位 + 基准日（基准日为入参，同步 4 张模板表的 A6）。"""
    unit = spec.d.get("unit")
    if unit:
        wb[SHEET_MATERIAL]["B2"] = unit
        wb[SHEET_PROCEDURE]["B2"] = unit
        wb[SHEET_LEDGER]["H2"] = unit
        wb[SHEET_RECON]["C2"] = unit

    disp = "（核查基准日：%s）" % spec.baseline_disp
    for sn in (SHEET_MATERIAL, SHEET_PROCEDURE, SHEET_LEDGER, SHEET_RECON):
        cell = wb[sn]["A6"]
        old = str(cell.value or "")
        new = BASE_RX.sub(disp, old, count=1)
        if new == old and disp not in old:
            new = (old.rstrip() + disp) if old else disp
        cell.value = new


def fill_material_list(wb, spec):
    """写资料清单：D 是否提供 / E 备注 / F 资料载体类型。"""
    rows = spec.d.get("material_list") or {}
    ws = wb[SHEET_MATERIAL]
    for r, pair in rows.items():
        r = int(r)
        d, e, f = (list(pair) + [None, None, None])[:3]
        ws.cell(r, 4).value = d
        ws.cell(r, 5).value = spec.s(e)
        ws.cell(r, 6).value = spec.s(f)


def fill_procedures(wb, spec):
    rows = spec.d.get("procedures") or {}
    ws = wb[SHEET_PROCEDURE]
    exec_ = spec.d.get("exec", "审计组")
    date = spec.d.get("date", "")
    for r, tup in rows.items():
        r = int(r)
        d, e, f, note = (list(tup) + [None, None, None, None])[:4]
        ws.cell(r, 4).value = spec.s(d)
        ws.cell(r, 5).value = e if e is not None else exec_
        ws.cell(r, 6).value = f if f is not None else date
        if note:
            base = ws.cell(r, 3).value or ""
            if "【本项目执行情况】" not in base:
                ws.cell(r, 3).value = base + "\n" + spec.s(note)


def fill_ledger(wb, spec):
    lg = spec.d.get("ledger") or {}
    if not lg:
        return
    ws = wb[SHEET_LEDGER]
    common = lg.get("common") or {}
    first = int(lg.get("first_row", 9))
    for i, row in enumerate(lg.get("rows") or []):
        vals = dict(common)
        vals.update(row)
        r = first + i
        for k, v in vals.items():
            ws.cell(r, col(k)).value = spec.s(v) if isinstance(v, str) else v

    fmt_cols = lg.get("fmt_cols") or [17, 18, 19, 20, 24, 27]
    n = len(lg.get("rows") or [])
    for c in fmt_cols:
        for i in range(n):
            cell = ws.cell(first + i, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"

    # 清除模板残留占位行（数据行之后到第 10 行）；模板 A9/A10、AD9..AG10 有 □ 勾选串
    for r in range(first + n, 11):
        for c in range(1, 35):
            ws.cell(r, c).value = None


def fill_reconciliation(wb, spec):
    rc = spec.d.get("reconciliation") or {}
    if not rc:
        return
    ws = wb[SHEET_RECON]
    first = int(rc.get("first_row", 9))
    rows = rc.get("rows") or []
    note_target = rc.get("note_target")
    note_text = ws["A11"].value

    last = first + len(rows) - 1
    if last >= 11 and not note_target:
        raise SystemExit("data reaches row %d but note_target not provided "
                         "(would collide with the template's merged note A11:I18)" % last)

    if note_target:
        for rng in list(ws.merged_cells.ranges):
            if str(rng) == "A11:I18":
                ws.unmerge_cells("A11:I18")
        ws["A11"].value = None

    for i, d in enumerate(rows):
        r = first + i
        for k, v in d.items():
            ws.cell(r, col(k)).value = spec.s(v) if isinstance(v, str) else v
            if k in ("E", "G") and isinstance(v, (int, float)):
                ws.cell(r, col(k)).number_format = "#,##0.00"
        # 未给差异类型时清掉模板的 □ 勾选串，避免占位残留
        if "H" not in d:
            ws.cell(r, col("H")).value = None

    # clear unused template placeholder rows (between last data row and row 10)
    for r in range(last + 1, 11):
        for c in range(1, 11):
            ws.cell(r, c).value = None

    if note_target and note_text:
        ws.merge_cells(note_target["merge"])
        a = ws[note_target["anchor"]]
        a.value = note_text
        a.alignment = Alignment(wrap_text=True, vertical="top")
        a.font = Font(size=10)


def make_issue_sheet(wb, spec):
    """《疑点清单》为第 5 张表，无论有无疑点都必须建（空表也要有标题与表头）。"""
    isx = spec.d.get("issues") or {}
    if SHEET_ISSUES in wb.sheetnames:
        del wb[SHEET_ISSUES]
    ws = wb.create_sheet(SHEET_ISSUES)
    thin = Side(style="thin", color="BDBDBD")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = isx.get("title", "疑点清单（客观列示，不作定性）")
    ws["A2"] = "被审计单位：%s    核查基准日：%s" % (spec.d.get("unit", ""), spec.baseline_disp)
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"].font = Font(size=10)

    head = ["序号", "疑点（客观事实）", "事实依据", "核查建议", "资料充分性"]
    for j, h in enumerate(head, 1):
        c = ws.cell(4, j)
        c.value = h
        c.font = Font(bold=True)
        c.border = bd
        c.fill = PatternFill("solid", fgColor="F5F7FA")
        c.alignment = Alignment(horizontal="center")

    for i, item in enumerate(isx.get("items") or []):
        r = 5 + i
        for j, v in enumerate(list(item)[:5], 1):
            cell = ws.cell(r, j)
            cell.value = spec.s(v)
            cell.border = bd
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for c, w in zip("ABCDE", [6, 56, 34, 40, 18]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A5"


def do_fill(spec_dict):
    spec = Spec(spec_dict)
    wb = copy_template(spec)
    set_units(wb, spec)
    fill_material_list(wb, spec)
    fill_procedures(wb, spec)
    fill_ledger(wb, spec)
    fill_reconciliation(wb, spec)
    make_issue_sheet(wb, spec)
    out = spec.d["output"]
    wb.save(out)
    print("SAVED:", out, os.path.exists(out), os.path.getsize(out))
    print("sheets:", wb.sheetnames)
    return out


EXAMPLE = {
    "template": "assets/workingpaper-template.xlsx",
    "output": "C:/…/…底稿_标的.xlsx",
    "unit": "成都市新都兴工建设投资有限公司",
    "baseline": "2026-09-30",
    "exec": "审计组",
    "date": "2026-10-16",
    "src_ref": "成都香城产业发展集团有限公司8.12_副本.xlsx → sheet「房产」",
    "material_list": {
        "8": ["√", "已提供（{SRC} 第56行：标的）", "电子文档"],
        "9": ["√", "《房屋租赁合同》（2024-05-01，24 页）；{SRC} 第56行", "原件扫描件"],
    },
    "procedures": {
        "8": ["《审计资料清单》", "审计组", "2026-10-16", None],
        "10": ["《合同与财务对账差异明细表》", "审计组", "2026-10-16",
               "【本项目执行情况】仅完成合同测算的单侧比对，财务数据未取得；未做审计定性。"],
    },
    "ledger": {
        "common": {"B": "成都香城产业发展集团有限公司", "E": "项目名", "H": "园区外",
                   "I": "房产", "L": "已出租", "U": "—（无欠缴）"},
        "rows": [
            {"A": 1, "C": "权属公司", "D": "新都区XX路41号商铺", "F": 132.79,
             "O": "2026-05-01", "P": "2027-04-30",
             "Q": 65333, "R": 32666.5, "S": 32666.5, "T": 0,
             "X": 65333, "AA": 0, "AB": "—", "AC": "—",
             "V": "41元/㎡·月（不递增）", "W": "按半年结算，先付后用",
             "AD": "□是 ☑否（合同免租期栏填“/”）", "AE": "□有 ☑无（不适用）",
             "AF": "□是 □否（未获取催收记录，待核实）",
             "AG": "□完整 □不完整（未获取催收记录，待核实）",
             "AH": "①事实：按付款周期逐期测算，截至基准日已到期应收65333元，台账应收65333元，两者一致。"
                   "②差异与原因：无。③其他：无产权证。"
                   "④来源：《房屋租赁合同》（2024-05-01，24 页）→ sheet「房产」第56行（新都区XX路41号商铺）。"},
            {"A": 2, "C": "权属公司", "D": "新都区YY路12号商铺", "F": 380.00,
             "O": "2025-03-01", "P": "2028-02-29",
             "Q": 579038.04, "R": 369495.71, "S": 0, "T": 209542.33,
             "X": 580000.00, "AA": 209542.33, "AB": "2025-09", "AC": "13个月",
             "V": "58元/㎡·月（第2年起递增3%）", "W": "按季结算",
             "AD": "☑是 □否（免租1个月）", "AE": "□有 ☑无（无减免审批资料）",
             "AF": "☑是 □否", "AG": "□完整 ☑不完整（缺送达回执）",
             "AH": "①事实：按付款周期逐期测算，截至基准日已到期应收580000.00元，台账应收579038.04元。"
                   "②差异与原因：台账少计961.96元，已在《合同与财务对账差异明细表》登记，待凭证核实，不作定性。"
                   "③其他：免租1个月已扣减。"
                   "④来源：《房屋租赁合同》（2025-03-01，22 页）→ sheet「房产」第61行（新都区YY路12号商铺）。"},
        ],
        "first_row": 9,
    },
    "reconciliation": {
        "rows": [
            {"A": 1, "B": "—（台账未列资产编号）", "C": "新都区XX路41号商铺", "D": "承租方甲",
             "E": 65333, "F": "资料未提供，待补", "G": "待补（财务数据未取得）",
             "H": "□时间性差异 □金额差异 □缺合同 □其他（待补后判定）",
             "I": "合同测算全期应收65333元，与台账应收一致；财务数据未取得，无法完成双面比对。"
                  "来源：《房屋租赁合同》（2024-05-01，24 页）→ 第56行。",
             "J": 1},
            {"A": 2, "B": "—（台账未列资产编号）", "C": "新都区YY路12号商铺", "D": "承租方乙",
             "E": 580000, "F": "资料未提供，待补", "G": 961.96,
             "H": "□时间性差异 ☑金额差异 □缺合同 □其他（台账与合同测算差异）",
             "I": "合同测算累计应收580000.00元，台账应收租金579038.04元，差额961.96元属台账与合同测算差异，"
                  "待凭证核实，不作定性。来源：《房屋租赁合同》（2025-03-01，22 页）→ 第61行。",
             "J": 2},
        ],
        "first_row": 9,
    },
    "issues": {
        "title": "…疑点清单（客观列示，不作定性）",
        "items": [
            ["1", "台账应收租金579038.04元与合同测算累计应收580000.00元相差961.96元，"
                  "台账未按月登记。",
             "台账 sheet「房产」第61行（新都区YY路12号商铺）；《房屋租赁合同》（2025-03-01，22 页）第3条。",
             "调取该合同项下全部收款凭证与应收明细账，逐期核对应收余额，差异确认后调整台账或补登凭证。",
             "部分充分（缺财务明细账）"],
            ["2", "资料清单第13行「租金收款记账凭证」未提供，账面累计实收租金无法核对。",
             "资料清单第13行（是否提供＝×）；台账 sheet「房产」第56行、第61行。",
             "向被审计单位索取租金收款记账凭证及发票，与银行流水逐笔核对实收金额。",
             "不充分（仅台账单面数据）"],
        ],
    },
}


def main():
    ap = argparse.ArgumentParser(description="Fill 租金收缴专项核查底稿 from a JSON spec.")
    ap.add_argument("spec", nargs="?", help="path to spec.json")
    ap.add_argument("--example", action="store_true", help="print an example spec and exit")
    args = ap.parse_args()
    if args.example or not args.spec:
        print(json.dumps(EXAMPLE, ensure_ascii=False, indent=2))
        return
    with open(args.spec, "r", encoding="utf-8") as f:
        do_fill(json.load(f))


if __name__ == "__main__":
    main()
