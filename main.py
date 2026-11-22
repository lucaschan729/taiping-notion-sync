import io
import math
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from playwright.sync_api import sync_playwright

START = "https://www.tplhk.cntaiping.com/info/fulfillment_ratio"
COMPANY = "中國太平"
OUT_DIR = "output"
NOTION_VERSION = "2025-09-03"
AUTO = "自動生成"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

JS_TABLES = """() => Array.from(document.querySelectorAll('table')).map(
  t => Array.from(t.rows).map(
    r => Array.from(r.cells).map(
      c => (c.innerText || c.textContent || '').replace(/\\s+/g, ' ').trim())))"""
JS_TEXT = "() => document.body ? document.body.innerText : ''"

PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*[%％]$")
NA = re.compile(r"^不適用\s*[\(（]\s*([A-Za-z])\s*[\)）]$")
TITLE = re.compile(r"(20\d{2})\s*報告年度的\s*(\S{2,24}?)實現率")
REASON = re.compile(r"不適用\s*[\(（]\s*([A-Za-z])\s*[\)）]\s*[:：]\s*([^；;。]+)")

N_FIXED = 2
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid("FFE9ECEF")
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color="FF1F2933")
TYPE_FONT = Font(color="FF2F5597")
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_data_table(t):
    return any(row and "生效年份" in row[0] for row in t[:3])


# ---------- 1. 讀取網頁 ----------

def fetch_once(p, label, **launch):
    print("\n--- 嘗試：%s ---" % label)
    browser = p.chromium.launch(
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        **launch)
    tables, text = [], ""
    try:
        ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                  viewport={"width": 1366, "height": 900})
        page = ctx.new_page()
        try:
            r = page.goto(START, wait_until="domcontentloaded", timeout=60000)
            print("   首次回應狀態：%s" % (r.status if r else "沒有"))
        except Exception as e:
            print("   載入出錯：%s" % one_line(str(e))[:300])
        last, stable = -1, 0
        for i in range(40):
            page.wait_for_timeout(2000)
            try:
                tables = page.evaluate(JS_TABLES)
                text = page.evaluate(JS_TEXT)
            except Exception:
                continue
            n = sum(len(t) for t in tables if is_data_table(t))
            if n > 0 and n == last:
                stable += 1
                if stable >= 2:
                    break
            else:
                stable = 0
            last = n
        try:
            print("   標題：%s" % page.title())
        except Exception:
            pass
        print("   實現率表格 %d 個｜頁面文字 %d 字" % (
            sum(1 for t in tables if is_data_table(t)), len(one_line(text))))
    finally:
        browser.close()
    return tables, text


def fetch():
    attempts = [("內置 Chromium", {}), ("內置 Chromium（重試）", {}),
                ("系統 Google Chrome", {"channel": "chrome"})]
    with sync_playwright() as p:
        for label, launch in attempts:
            try:
                tables, text = fetch_once(p, label, **launch)
            except Exception as e:
                print("   出錯：%s" % one_line(str(e))[:300])
                continue
            if any(is_data_table(t) for t in tables):
                return tables, text
    print("三次嘗試都讀不到實現率表格，停止。")
    sys.exit(1)


# ---------- 2. 解析 ----------

def parse_cell(s):
    s = (s or "").strip()
    m = PCT.match(s)
    if m:
        return round(float(m.group(1)) / 100.0, 4), None
    m = NA.match(s)
    if m:
        return "/", m.group(1).lower()
    if s in ("", "-", "–", "—", "/", "不適用", "N/A"):
        return "/", "無註腳"
    return s, "!"


def parse(tables, text):
    problems = []
    flat = one_line(text)

    titles = []
    for m in TITLE.finditer(flat):
        if "生效年份" in flat[m.end():m.end() + 30]:
            kind = m.group(2)
            if kind.endswith("分紅") and len(kind) > 2:
                kind = kind[:-2]
            titles.append((m.group(1), kind))
    data = [t for t in tables if is_data_table(t)]
    print("\n=== 解析 ===")
    print("   表格 %d 個｜標題 %d 個：%s" % (
        len(data), len(titles), "、".join("%s %s" % x for x in titles)))
    named = len(titles) == len(data)
    if not named:
        problems.append("表格數目（%d）與標題數目（%d）不一致" % (len(data), len(titles)))

    years = set(y for y, _ in titles)
    sections = []
    for i, t in enumerate(data):
        kind = titles[i][1] if named else "表格%d" % (i + 1)
        yrs, pol, rows = None, None, []
        for row in t:
            if not row or not any(row):
                continue
            head = row[0]
            if "生效年份" in head:
                yrs = row[1:]
                continue
            if "保單年度" in head:
                pol = row[1:]
                m = re.search(r"截至\s*(20\d{2})", head)
                if m:
                    years.add(m.group(1))
                continue
            rows.append(row)
        if not yrs or not pol or len(yrs) != len(pol):
            problems.append("%s：找不到完整的表頭" % kind)
            continue
        if not all(re.fullmatch(r"(19|20)\d{2}", y) for y in yrs):
            problems.append("%s：生效年份不是四位數年份：%s" % (kind, "、".join(yrs)))
        parsed = []
        for row in rows:
            name = row[0]
            if not name:
                problems.append("%s：有一行沒有產品名稱：%s" % (kind, "｜".join(row)[:80]))
                continue
            if len(row) - 1 != len(yrs):
                problems.append("%s／%s：欄數 %d，應為 %d" % (kind, name, len(row) - 1, len(yrs)))
                continue
            cells = [parse_cell(c) for c in row[1:]]
            for j, (v, code) in enumerate(cells):
                if code == "!":
                    problems.append("%s／%s／%s生效：無法識別「%s」" % (kind, name, yrs[j], v))
            parsed.append((name, cells))
        names = [n for n, _ in parsed]
        dup = sorted(set(n for n in names if names.count(n) > 1))
        if dup:
            print("   注意：%s 有重複產品名稱（全部保留）：%s" % (kind, "、".join(dup)))
        if not parsed:
            problems.append("%s：沒有產品" % kind)
        sections.append({"kind": kind, "cols": yrs, "rows": parsed})

    if not sections:
        print("沒有可用的表格，停止。")
        for x in problems:
            print("   ✗ %s" % x)
        sys.exit(1)
    if len(years) != 1:
        problems.append("報告年度不一致：%s" % "、".join(sorted(years)))
    year = sorted(years)[-1] if years else "未知年度"
    cols = sections[0]["cols"]
    for sec in sections[1:]:
        if sec["cols"] != cols:
            problems.append("%s 的年度欄位與其他表格不一致" % sec["kind"])

    reasons = {}
    for m in REASON.finditer(flat):
        reasons.setdefault(m.group(1).lower(), m.group(2).strip())
    print("   官方註釋：%s" % "｜".join("(%s) %s" % (k, reasons[k]) for k in sorted(reasons)))
    return year, cols, sections, problems, reasons


def numbers(rows):
    return [v for r in rows for v in r[N_FIXED:] if is_num(v)]


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def fmt_pct(v):
    return "%.2f%%" % (v * 100) if is_num(v) else "/"


def to_rows(sections):
    """同一產品的不同紅利類型排在一起；回傳 [產品名稱, 紅利類型, 各年度數值...]"""
    order, by = [], {}
    for sec in sections:
        for name, cells in sec["rows"]:
            if name not in by:
                by[name] = []
                order.append(name)
            by[name].append([name, sec["kind"]] + [v for v, _ in cells])
    return [row for name in order for row in by[name]]


def report(sections, rows, cols):
    print("\n=== 核對 ===")
    for sec in sections:
        sub = [r for r in rows if r[1] == sec["kind"]]
        nums = numbers(sub)
        with_val = sum(1 for r in sub if any(is_num(v) for v in r[N_FIXED:]))
        print("   %s：%d 產品（%d 個有數值）｜%d 個數值｜%d 格不適用｜平均 %s｜最高 %s｜最低 %s" % (
            sec["kind"], len(sub), with_val, len(nums), len(sub) * len(cols) - len(nums),
            fmt_pct(mean_or_slash(nums)),
            fmt_pct(max(nums) if nums else "/"), fmt_pct(min(nums) if nums else "/")))
        items = [(r[N_FIXED + j], r[0], cols[j]) for r in sub
                 for j in range(len(cols)) if is_num(r[N_FIXED + j])]
        if items:
            hi, lo = max(items), min(items)
            print("      最高：%s／%s生效／%.0f%%" % (hi[1], hi[2], hi[0] * 100))
            print("      最低：%s／%s生效／%.0f%%" % (lo[1], lo[2], lo[0] * 100))
        codes = {}
        for _, cells in sec["rows"]:
            for _, code in cells:
                if code:
                    codes[code] = codes.get(code, 0) + 1
        print("      不適用註腳：%s" % "、".join("(%s) %d" % (k, codes[k]) for k in sorted(codes)))


def build_summary(cols, groups):
    n = len(cols)
    width = N_FIXED + n
    head = ["所有年份平均值"] + ["持滿%d年平均值" % (i + 1) for i in range(1, n)]
    cum = []
    for label, sub in groups:
        line = [label]
        for start in range(N_FIXED, width):
            line.append(mean_or_slash(
                [r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(N_FIXED, width)]
    return {"head": head, "cum": cum,
            "year_head": ["%s生效" % y for y in cols], "year": [year_line]}


# ---------- 3. Excel（參考保誠格式） ----------

def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    """100% 或以上綠底；低於 80% 紅底；沒有數字用淺灰字。"""
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def pct_format(v):
    p = v * 100
    return "0%" if abs(p - round(p)) < 1e-6 else "0.0%"


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 16]
    caps = [52, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(N_FIXED)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = TYPE_FONT
            else:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = "A%d:%s%d" % (HEAD_ROW, get_column_letter(last_col), data_end)

    lc = N_FIXED
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= N_FIXED else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "%d:%d" % (HEAD_ROW, HEAD_ROW)


def add_overview(wb, year, today, title, rows, kinds, reasons):
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def body(r, values, pct_cols):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c in pct_cols:
                cell.number_format = pct_cols[c]
                heat(cell, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    def merged_line(r, first, desc):
        cell = ws.cell(row=r, column=1, value=first)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H
        return cell

    banner(1, "%s分紅表現概覽（報告年度 %s）" % (COMPANY, year), TITLE_FONT, DARK, 40)
    banner(2, "資料來源：中國太平人壽保險（香港）網站｜更新日期：%s" % today, NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    nums = numbers(rows)
    low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
    r += 1
    body(r, [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low],
         {5: "0.00%"})
    ws.cell(row=r, column=6).number_format = "0.0%"
    link = ws.cell(row=r, column=1)
    link.font = LINK_FONT
    link.hyperlink = Hyperlink(ref=link.coordinate, location="'%s'!A1" % title)

    r += 2
    section(r, "按紅利類型的統計")
    r += 1
    table_head(r, ["紅利類型", "產品數", "有數字的格數", "平均值", "最高", "最低"])
    for kind in kinds:
        sub = [row for row in rows if row[1] == kind]
        ns = numbers(sub)
        r += 1
        body(r, [kind, len(sub), len(ns), mean_or_slash(ns),
                 max(ns) if ns else "/", min(ns) if ns else "/"],
             {4: "0.00%", 5: "0%", 6: "0%"})
        ws.cell(row=r, column=1).font = BLACK_BOLD

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方網站顯示「不適用」，原因見下方")]
    for v, desc in samples:
        r += 1
        cell = merged_line(r, v, desc)
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)

    if reasons:
        r += 2
        section(r, "「/」的原因（官方網站註釋）")
        for k in sorted(reasons):
            r += 1
            merged_line(r, "不適用 (%s)" % k, reasons[k])

    r += 2
    section(r, "使用提示")
    tips = ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
            "工作表第 3 行有篩選箭嘴，可按產品名稱或紅利類型篩選。",
            "向右或向下捲動時，標題列和產品名稱會固定不動。",
            "平均值摘要在工作表最底部。",
            "官方網站沒有公布總現金價值比率；萬用壽險派息率不納入。",
            "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"]
    for t in tips:
        r += 1
        cell = ws.cell(row=r, column=1, value="・" + t)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = NOTE_FONT
        cell.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H


def build_excel(year, cols, sections, rows, reasons, now):
    today = now.strftime("%Y-%m-%d")
    kinds = [sec["kind"] for sec in sections]
    groups = [("所有分紅", rows)]
    for kind in kinds:
        sub = [r for r in rows if r[1] == kind]
        if numbers(sub):
            groups.append((kind, sub))
    summary = build_summary(cols, groups)
    print("\n=== 平均值摘要 ===")
    for line in summary["cum"]:
        print("   累計 %s：%s" % (line[0], " ".join(fmt_pct(v) for v in line[1:])))
    for line in summary["year"]:
        print("   當年 %s：%s" % (line[0], " ".join(fmt_pct(v) for v in line[1:])))

    header = ["产品名称", "红利类型"] + ["保单年度%d\n%s" % (i + 1, y) for i, y in enumerate(cols)]
    notes = ("分紅實現率＝實際派發的累積非保證利益 ÷ 銷售時所示的金額。"
             "100% 即與銷售時一致，高於 100% 即比銷售時多。年度欄下方的年份為保單生效年份。\n"
             "綠色＝100% 或以上；紅色＝低於 80%；「/」＝官方網站顯示「不適用」（原因見「概覽」）。\n"
             "資料來源：中國太平人壽保險（香港）網站（報告年度 %s）｜更新日期：%s" % (year, today))
    title = "%s分紅實現率" % year
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, title, "%s分紅實現率（報告年度 %s）" % (COMPANY, year), notes,
              header, rows, summary, "分紅類型", "2E75B6")
    add_overview(wb, year, today, title, rows, kinds, reasons)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()

    os.makedirs(OUT_DIR, exist_ok=True)
    name = "%s分紅實現率（%s年）%s.xlsx" % (COMPANY, year, AUTO)
    path = os.path.join(OUT_DIR, name)
    with open(path, "wb") as fh:
        fh.write(content)
    print("\n=== Excel ===")
    print("   %s｜%d KB｜%d 個產品｜%d 行｜工作表：%s" % (
        path, len(content) // 1024, len({r[0] for r in rows}), len(rows), "、".join(wb.sheetnames)))
    return name, content


# ---------- 4. Notion ----------

def notion_upload(name, content, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = (os.environ.get("NOTION_DATABASE_URL") or os.environ.get("NOTION_DATABASE_ID") or "").strip()
    print("\n=== Notion ===")
    if not token or not raw:
        print("   未設定 Notion Secrets，只生成 Excel。")
        return
    m = re.search(r"([0-9a-fA-F]{32})$", raw.split("?")[0].rstrip("/").replace("-", ""))
    if not m:
        print("   資料庫網址內找不到 ID，停止。")
        sys.exit(1)
    db_id = m.group(1)
    auth = {"Authorization": "Bearer " + token, "Notion-Version": NOTION_VERSION}

    def api(method, url, body=None):
        r = requests.request(method, "https://api.notion.com/v1" + url, headers=auth,
                             json=body, timeout=60)
        if r.status_code >= 300:
            raise RuntimeError("Notion %s %s：%s %s" % (
                method, url.split("?")[0], r.status_code, r.text[:300]))
        return r.json()

    sources = api("GET", "/databases/" + db_id).get("data_sources") or []
    if not sources:
        print("   資料庫沒有 data source，停止。")
        sys.exit(1)
    ds = sources[0]["id"]
    props = api("GET", "/data_sources/" + ds).get("properties", {})
    title_prop = next((k for k, v in props.items() if v.get("type") == "title"), None)
    if not title_prop:
        print("   找不到標題欄，停止。")
        sys.exit(1)

    page_id, cursor = None, None
    while True:
        body = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        res = api("POST", "/data_sources/%s/query" % ds, body)
        for pg in res.get("results", []):
            if pg.get("in_trash") or pg.get("archived"):
                continue
            parts = pg.get("properties", {}).get(title_prop, {}).get("title", [])
            title = "".join(x.get("plain_text", "") for x in parts)
            if str(year) in title:
                page_id = pg["id"]
                print("   使用現有頁面：%s" % title)
                break
        if page_id or not res.get("has_more"):
            break
        cursor = res.get("next_cursor")
    if not page_id:
        pg = api("POST", "/pages", {
            "parent": {"type": "data_source_id", "data_source_id": ds},
            "properties": {title_prop: {"title": [{"text": {"content": str(year)}}]}}})
        page_id = pg["id"]
        print("   已建立頁面：%s" % year)

    old, cursor = [], None
    while True:
        url = "/blocks/%s/children?page_size=100" % page_id
        if cursor:
            url += "&start_cursor=" + cursor
        res = api("GET", url)
        for blk in res.get("results", []):
            if blk.get("type") != "file":
                continue
            f = blk.get("file", {})
            cap = "".join(x.get("plain_text", "") for x in f.get("caption", []))
            if AUTO in (f.get("name") or "") or AUTO in cap:
                old.append(blk["id"])
        if not res.get("has_more"):
            break
        cursor = res.get("next_cursor")

    up = api("POST", "/file_uploads", {"mode": "single_part", "filename": name,
                                       "content_type": XLSX})
    r = requests.post("https://api.notion.com/v1/file_uploads/%s/send" % up["id"],
                      headers=auth, timeout=120,
                      files={"file": (name, content, XLSX)})
    if r.status_code >= 300:
        raise RuntimeError("傳送檔案失敗：%s %s" % (r.status_code, r.text[:300]))
    api("PATCH", "/blocks/%s/children" % page_id, {"children": [{
        "object": "block", "type": "file",
        "file": {"type": "file_upload", "file_upload": {"id": up["id"]}}}]})
    for bid in old:
        api("DELETE", "/blocks/" + bid)
    print("   已上傳：%s（刪除舊的自動生成檔案 %d 個）" % (name, len(old)))


def main():
    now = datetime.now(timezone(timedelta(hours=8)))
    tables, text = fetch()
    year, cols, sections, problems, reasons = parse(tables, text)
    rows = to_rows(sections)
    report(sections, rows, cols)
    name, content = build_excel(year, cols, sections, rows, reasons, now)
    if problems:
        print("\n=== 發現 %d 個問題，不上傳 Notion ===" % len(problems))
        for x in problems[:60]:
            print("   ✗ %s" % x)
        sys.exit(1)
    notion_upload(name, content, year)
    print("\n=== 完成 ===")


if __name__ == "__main__":
    main()
