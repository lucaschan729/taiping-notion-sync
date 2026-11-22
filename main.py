import os
import re
import sys
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from playwright.sync_api import sync_playwright

START = "https://www.tplhk.cntaiping.com/info/fulfillment_ratio"
COMPANY = "中國太平人壽"
OUT_DIR = "output"
NOTION_VERSION = "2025-09-03"
AUTO_MARK = "自動生成"
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

JS_TABLES = """() => Array.from(document.querySelectorAll('table')).map(
  t => Array.from(t.rows).map(
    r => Array.from(r.cells).map(
      c => (c.innerText || c.textContent || '').replace(/\\s+/g, ' ').trim())))"""
JS_TEXT = "() => document.body ? document.body.innerText : ''"

PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*[%％]$")
NA = re.compile(r"^不適用\s*[\(（]\s*([A-Za-z])\s*[\)）]$")
TITLE = re.compile(r"(20\d{2})\s*報告年度的\s*(\S{2,24}?)實現率")

BOLD = Font(bold=True)
GREY = Font(color="808080")
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
SEP = PatternFill("solid", fgColor="D9D9D9")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


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
        return float(m.group(1)) / 100.0, None
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
        cols = ["%s\n%s生效" % (a, b) for a, b in zip(pol, yrs)]
        parsed = []
        for row in rows:
            name = row[0]
            if not name:
                problems.append("%s：有一行沒有產品名稱：%s" % (kind, "｜".join(row)[:80]))
                continue
            if len(row) - 1 != len(cols):
                problems.append("%s／%s：欄數 %d，應為 %d" % (kind, name, len(row) - 1, len(cols)))
                continue
            cells = [parse_cell(c) for c in row[1:]]
            for j, (v, code) in enumerate(cells):
                if code == "!":
                    problems.append("%s／%s／%s：無法識別「%s」" % (
                        kind, name, cols[j].replace("\n", " "), v))
            parsed.append((name, cells))
        names = [n for n, _ in parsed]
        dup = sorted(set(n for n in names if names.count(n) > 1))
        if dup:
            print("   注意：%s 有重複產品名稱（全部保留）：%s" % (kind, "、".join(dup)))
        if not parsed:
            problems.append("%s：沒有產品" % kind)
        sections.append({"kind": kind, "cols": cols, "rows": parsed})

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

    i = flat.find("產品已停售")
    if i >= 0:
        print("   官方註釋（節錄）：%s" % flat[max(0, i - 400):i + 10])
    return year, cols, sections, problems


def stat(rows):
    vals = [v for _, cells in rows for v, _ in cells if isinstance(v, float)]
    return {
        "products": len(rows),
        "with_val": sum(1 for _, cells in rows if any(isinstance(v, float) for v, _ in cells)),
        "n": len(vals),
        "na": sum(1 for _, cells in rows for v, _ in cells if not isinstance(v, float)),
        "avg": sum(vals) / len(vals) if vals else None,
        "max": max(vals) if vals else None,
        "min": min(vals) if vals else None,
    }


def pct(v):
    return "沒有" if v is None else "%.2f%%" % (v * 100)


def report(sections):
    print("\n=== 核對 ===")
    for sec in sections:
        s = stat(sec["rows"])
        print("   %s：%d 產品（%d 個有數值）｜%d 個數值｜%d 格不適用｜平均 %s｜最高 %s｜最低 %s" % (
            sec["kind"], s["products"], s["with_val"], s["n"], s["na"],
            pct(s["avg"]), pct(s["max"]), pct(s["min"])))
        items = [(v, name, sec["cols"][j].replace("\n", " "))
                 for name, cells in sec["rows"]
                 for j, (v, _) in enumerate(cells) if isinstance(v, float)]
        if items:
            hi, lo = max(items), min(items)
            print("      最高：%s／%s／%.0f%%" % (hi[1], hi[2], hi[0] * 100))
            print("      最低：%s／%s／%.0f%%" % (lo[1], lo[2], lo[0] * 100))
        codes = {}
        for _, cells in sec["rows"]:
            for _, code in cells:
                if code:
                    codes[code] = codes.get(code, 0) + 1
        print("      不適用註腳：%s" % "、".join("(%s) %d" % (k, codes[k]) for k in sorted(codes)))


# ---------- 3. Excel ----------

def put(cell, v, fmt=None):
    cell.alignment = CENTER
    cell.border = BORDER
    cell.value = v
    if isinstance(v, float):
        p = v * 100
        cell.number_format = fmt or ("0%" if abs(p - round(p)) < 1e-6 else "0.0%")
        if round(v, 6) >= 1:
            cell.fill = GREEN
        elif round(v, 6) < 0.8:
            cell.fill = RED
    elif v == "/":
        cell.font = GREY


def head(ws, row, labels):
    for c, label in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.alignment = CENTER
        cell.border = BORDER


def build_excel(year, cols, sections, now):
    order, by = [], {}
    for sec in sections:
        for name, cells in sec["rows"]:
            if name not in by:
                by[name] = []
                order.append(name)
            by[name].append((sec["kind"], cells))

    wb = Workbook()
    ov = wb.active
    ov.title = "概覽"
    ws = wb.create_sheet("%s分紅實現率" % year)

    # 概覽
    ov["A1"] = "%s（香港）分紅實現率概覽" % COMPANY
    ov["A1"].font = Font(bold=True, size=14)
    info = [("報告年度", year), ("資料來源", START),
            ("更新時間", now.strftime("%Y-%m-%d %H:%M") + "（香港時間）")]
    for i, (k, v) in enumerate(info, 2):
        ov.cell(row=i, column=1, value=k).font = BOLD
        ov.cell(row=i, column=2, value=v)
    head(ov, 6, ["紅利類別", "產品數", "有數值產品", "數值格數", "「不適用」格數", "平均", "最高", "最低"])
    r = 7
    lines = [(sec["kind"], stat(sec["rows"])) for sec in sections]
    total = stat([row for sec in sections for row in sec["rows"]])
    total["products"] = len(order)
    total["with_val"] = sum(1 for n in order
                            if any(isinstance(v, float) for _, cells in by[n] for v, _ in cells))
    lines.append(("合計", total))
    for kind, s in lines:
        c = ov.cell(row=r, column=1, value=kind)
        c.font = BOLD
        c.border = BORDER
        c.alignment = LEFT
        for j, key in enumerate(["products", "with_val", "n", "na"], 2):
            put(ov.cell(row=r, column=j), s[key])
        for j, key in enumerate(["avg", "max", "min"], 6):
            put(ov.cell(row=r, column=j), s[key] if s[key] is not None else "/",
                "0.00%" if key == "avg" else None)
        r += 1
    notes = ["「不適用」在工作表中以 / 顯示，原因見官方網頁註釋。",
             "100% 或以上為綠底，低於 80% 為紅底。",
             "官方網頁沒有公布總現金價值比率；萬用壽險派息率不納入。",
             "平均為各數值的簡單平均。"]
    for i, t in enumerate(notes, r + 1):
        ov.cell(row=i, column=1, value=t)
    ov.column_dimensions["A"].width = 26
    for col in "BCDEFGH":
        ov.column_dimensions[col].width = 16

    # 明細
    ncol = 2 + len(cols) + 1
    head(ws, 1, ["產品名稱", "紅利類別"] + cols + ["平均"])
    ws.row_dimensions[1].height = 34
    r = 2
    for k, name in enumerate(order):
        if k:
            for c in range(1, ncol + 1):
                ws.cell(row=r, column=c).fill = SEP
            ws.row_dimensions[r].height = 6
            r += 1
        for kind, cells in by[name]:
            a = ws.cell(row=r, column=1, value=name)
            a.font = BOLD
            a.alignment = LEFT
            a.border = BORDER
            b = ws.cell(row=r, column=2, value=kind)
            b.alignment = CENTER
            b.border = BORDER
            vals = []
            for j, (v, _) in enumerate(cells):
                put(ws.cell(row=r, column=3 + j), v)
                if isinstance(v, float):
                    vals.append(v)
            put(ws.cell(row=r, column=ncol), sum(vals) / len(vals) if vals else "/", "0.0%")
            r += 1
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 20
    for c in range(3, ncol + 1):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 11
    ws.freeze_panes = "C2"

    os.makedirs(OUT_DIR, exist_ok=True)
    name = "%s_%s分紅實現率.xlsx" % (COMPANY, year)
    path = os.path.join(OUT_DIR, name)
    wb.save(path)
    print("\n=== Excel ===")
    print("   %s｜%d 個產品｜%d 行" % (path, len(order), sum(len(by[n]) for n in order)))
    return path, name


# ---------- 4. Notion ----------

def notion_upload(path, name, year, now):
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
                print("   找到年份頁面：%s" % title)
                break
        if page_id or not res.get("has_more"):
            break
        cursor = res.get("next_cursor")
    if not page_id:
        pg = api("POST", "/pages", {
            "parent": {"type": "data_source_id", "data_source_id": ds},
            "properties": {title_prop: {"title": [{"text": {"content": str(year)}}]}}})
        page_id = pg["id"]
        print("   已建立年份頁面：%s" % year)

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
            if (f.get("name") or "").startswith(COMPANY + "_") or AUTO_MARK in cap:
                old.append(blk["id"])
        if not res.get("has_more"):
            break
        cursor = res.get("next_cursor")

    up = api("POST", "/file_uploads", {"mode": "single_part", "filename": name,
                                       "content_type": XLSX_TYPE})
    with open(path, "rb") as fh:
        r = requests.post("https://api.notion.com/v1/file_uploads/%s/send" % up["id"],
                          headers=auth, timeout=120,
                          files={"file": ("taiping_%s.xlsx" % year, fh, XLSX_TYPE)})
    if r.status_code >= 300 or r.json().get("status") != "uploaded":
        raise RuntimeError("上傳檔案失敗：%s %s" % (r.status_code, r.text[:300]))
    caption = "%s｜%s 更新" % (AUTO_MARK, now.strftime("%Y-%m-%d"))
    api("PATCH", "/blocks/%s/children" % page_id, {"children": [{
        "object": "block", "type": "file",
        "file": {"type": "file_upload", "file_upload": {"id": up["id"]},
                 "caption": [{"type": "text", "text": {"content": caption}}]}}]})
    print("   已上傳：%s" % name)
    for bid in old:
        api("DELETE", "/blocks/" + bid)
    print("   已移除舊的自動生成檔案 %d 個" % len(old))


def main():
    now = datetime.now(timezone(timedelta(hours=8)))
    tables, text = fetch()
    year, cols, sections, problems = parse(tables, text)
    report(sections)
    path, name = build_excel(year, cols, sections, now)
    if problems:
        print("\n=== 發現 %d 個問題，不上傳 Notion ===" % len(problems))
        for x in problems[:60]:
            print("   ✗ %s" % x)
        sys.exit(1)
    notion_upload(path, name, year, now)
    print("\n=== 完成 ===")


if __name__ == "__main__":
    main()
