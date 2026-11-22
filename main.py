import re

from playwright.sync_api import sync_playwright

START = "https://www.tplhk.cntaiping.com/info/fulfillment_ratio"
KEYWORDS = ["報告年度", "分紅實現率", "總現金價值", "現金價值比率", "過往派息率",
            "萬用壽險", "週年紅利", "終期紅利", "復歸紅利", "特別紅利", "不適用", "N/A"]
LINK_PAT = re.compile(r"實現率|現金價值|派息|紅利|fulfil|ratio|dividend|\.pdf|\.xlsx?", re.I)

JS_TABLES = """() => Array.from(document.querySelectorAll('table')).map(
  t => Array.from(t.rows).map(
    r => Array.from(r.cells).map(c => c.innerText.trim().replace(/\\s+/g, ' '))))"""
JS_SELECTS = """() => Array.from(document.querySelectorAll('select')).map(
  s => ({id: s.id, name: s.name, opts: Array.from(s.options).map(o => o.text.trim())}))"""
JS_LINKS = """() => Array.from(document.querySelectorAll('a')).map(
  a => [a.innerText.trim().replace(/\\s+/g, ' ').slice(0, 40), a.href])"""
JS_TEXT = "() => document.body ? document.body.innerText : ''"


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


def probe(p, label, **launch):
    print("\n########## 嘗試：%s ##########" % label)
    try:
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            **launch)
    except Exception as e:
        print("   無法啟動瀏覽器：%s" % one_line(str(e))[:300])
        return False

    ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                              viewport={"width": 1366, "height": 900})
    page = ctx.new_page()
    responses = []

    def on_response(resp):
        try:
            if resp.request.resource_type in ("document", "xhr", "fetch"):
                responses.append(resp)
        except Exception:
            pass

    page.on("response", on_response)

    print("=== 1. 載入 ===")
    try:
        r = page.goto(START, wait_until="domcontentloaded", timeout=60000)
        print("   首次回應狀態：%s" % (r.status if r else "沒有"))
    except Exception as e:
        print("   載入出錯：%s" % one_line(str(e))[:300])

    ok = False
    for i in range(25):
        page.wait_for_timeout(2000)
        try:
            html = page.content()
        except Exception:
            continue
        if "實現率" in html:
            ok = True
            print("   約 %d 秒後出現「實現率」字樣" % ((i + 1) * 2))
            break
    if not ok:
        print("   等了 50 秒仍未出現「實現率」字樣")
    page.wait_for_timeout(6000)

    try:
        print("   最終網址：%s" % page.url)
        print("   標題：%s" % page.title())
    except Exception as e:
        print("   讀取標題出錯：%s" % one_line(str(e))[:200])

    print("\n=== 2. 網絡請求（頁面及資料介面）：%d 個 ===" % len(responses))
    for resp in responses[:80]:
        try:
            rt = resp.request.resource_type
            ctype = resp.headers.get("content-type", "")
            print("   [%s] %s %s｜%s｜%s" % (rt, resp.request.method, resp.status, ctype[:40], resp.url[:200]))
            if rt in ("xhr", "fetch"):
                post = resp.request.post_data
                if post:
                    print("        送出：%s" % one_line(post)[:300])
                try:
                    body = resp.text()
                    print("        回傳 %d 字：%s" % (len(body), one_line(body)[:400]))
                except Exception as e:
                    print("        讀不到內容：%s" % one_line(str(e))[:120])
        except Exception:
            pass

    print("\n=== 3. 關鍵字 ===")
    try:
        text = page.evaluate(JS_TEXT)
    except Exception:
        text = ""
    flat_text = one_line(text)
    print("   頁面文字總長 %d 字" % len(flat_text))
    for kw in KEYWORDS:
        n = flat_text.count(kw)
        i = flat_text.find(kw)
        ctx_txt = flat_text[max(0, i - 50):i + 120] if i >= 0 else ""
        print("   %s：%d 次｜%s" % (kw, n, ctx_txt))

    frames = page.frames
    print("\n=== 4. 框架：%d 個 ===" % len(frames))
    for k, f in enumerate(frames, 1):
        print("\n--- 框架 %d：%s ---" % (k, f.url[:200]))
        try:
            tables = f.evaluate(JS_TABLES)
        except Exception as e:
            print("   讀取出錯：%s" % one_line(str(e))[:200])
            continue
        print("   表格 %d 個" % len(tables))
        for n, t in enumerate(tables[:8], 1):
            width = max([len(row) for row in t] or [0])
            print("   表格 %d｜%d 行｜最多 %d 欄" % (n, len(t), width))
            for row in t[:7]:
                print("      %s" % "｜".join(row)[:260])
        try:
            selects = f.evaluate(JS_SELECTS)
        except Exception:
            selects = []
        print("   下拉選單 %d 個" % len(selects))
        for s in selects[:8]:
            print("      id=%s name=%s｜選項 %d 個｜%s" % (
                s.get("id"), s.get("name"), len(s.get("opts", [])), "、".join(s.get("opts", [])[:20])[:400]))
        try:
            links = f.evaluate(JS_LINKS)
        except Exception:
            links = []
        seen = []
        for label_txt, href in links:
            if href and (LINK_PAT.search(href) or LINK_PAT.search(label_txt)):
                item = "%s｜%s" % (label_txt, href)
                if item not in seen:
                    seen.append(item)
        print("   相關連結 %d 條" % len(seen))
        for item in seen[:60]:
            print("      %s" % item[:260])

    print("\n=== 5. 頁面文字（節錄）===")
    i = flat_text.find("報告年度")
    if i < 0:
        i = flat_text.find("實現率")
    start = max(0, i - 200) if i >= 0 else 0
    print("   %s" % flat_text[start:start + 3500])

    browser.close()
    return ok


def main():
    with sync_playwright() as p:
        ok = probe(p, "Playwright 內置 Chromium")
        if not ok:
            ok = probe(p, "系統 Google Chrome", channel="chrome")
    print("\n=== 結論：%s ===" % ("成功讀到頁面內容" if ok else "兩種瀏覽器都未能通過驗證"))
    print("=== 完 ===")


if __name__ == "__main__":
    main()
