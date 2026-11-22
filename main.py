import html as htmllib
import re
from urllib.parse import urljoin

from curl_cffi import requests as creq

START = "https://www.tplhk.cntaiping.com/info/fulfillment_ratio"
KEYWORDS = ["報告年度", "分紅實現率", "總現金價值", "現金價值比率", "過往派息率",
            "萬用壽險", "週年紅利", "終期紅利", "復歸紅利", "特別紅利", "不適用", "N/A"]
LINK_PAT = re.compile(r"實現率|現金價值|派息|紅利|fulfil|ratio|dividend|\.pdf|\.xlsx?", re.I)
API_PAT = re.compile(r"[\"'](/?[A-Za-z0-9_\-/\.]*(?:api|ajax|json|query|list|get)[A-Za-z0-9_\-/\.\?=&]*)[\"']", re.I)

s = creq.Session(impersonate="chrome")


def flat(t):
    t = re.sub(r"<(script|style)\b.*?</\1>", " ", t or "", flags=re.S | re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", htmllib.unescape(t)).strip()


def attr(tag, name):
    m = re.search(r"\b" + name + r"\s*=\s*[\"']([^\"']*)[\"']", tag, flags=re.I)
    return m.group(1) if m else ""


def main():
    r = s.get(START, timeout=90, headers={"Accept-Language": "zh-HK,zh;q=0.9"})
    raw = r.content.decode("utf-8", "replace")
    print("=== 1. 網頁 ===")
    print("   狀態 %s｜%d KB｜最終網址 %s" % (r.status_code, len(raw) // 1024, r.url))
    m = re.search(r"<title[^>]*>(.*?)</title>", raw, flags=re.S | re.I)
    print("   標題：%s" % (flat(m.group(1)) if m else "沒有"))

    print("\n=== 2. 關鍵字 ===")
    for kw in KEYWORDS:
        hits = [x.start() for x in re.finditer(re.escape(kw), raw)]
        ctx = flat(raw[max(0, hits[0] - 60):hits[0] + 120])[:150] if hits else ""
        print("   %s：%d 次｜%s" % (kw, len(hits), ctx))

    tables = re.findall(r"<table\b.*?</table>", raw, flags=re.S | re.I)
    print("\n=== 3. 表格：%d 個 ===" % len(tables))
    for k, t in enumerate(tables[:4], 1):
        trs = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
        print("   表格 %d｜%d 行" % (k, len(trs)))
        for tr in trs[:5]:
            cells = [flat(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
            print("      %s" % "｜".join(cells)[:220])

    selects = re.findall(r"(<select\b[^>]*>)(.*?)</select>", raw, flags=re.S | re.I)
    print("\n=== 4. 下拉選單：%d 個 ===" % len(selects))
    for tag, body in selects[:6]:
        opts = [flat(o) for o in re.findall(r"<option\b[^>]*>(.*?)</option>", body, flags=re.S | re.I)]
        print("   id=%s name=%s｜選項 %d 個｜%s" % (attr(tag, "id"), attr(tag, "name"), len(opts), "、".join(opts[:15])[:300]))

    print("\n=== 5. iframe ===")
    for tag in re.findall(r"<iframe\b[^>]*>", raw, flags=re.I)[:10]:
        print("   %s" % urljoin(r.url, attr(tag, "src")))

    print("\n=== 6. 相關連結 ===")
    seen = []
    for tag, body in re.findall(r"(<a\b[^>]*>)(.*?)</a>", raw, flags=re.S | re.I):
        href = attr(tag, "href")
        label = flat(body)[:40]
        if href and (LINK_PAT.search(href) or LINK_PAT.search(label)):
            item = "%s｜%s" % (label, urljoin(r.url, href))
            if item not in seen:
                seen.append(item)
    print("   共 %d 條" % len(seen))
    for item in seen[:60]:
        print("   %s" % item)

    print("\n=== 7. Script ===")
    srcs = [attr(t, "src") for t in re.findall(r"<script\b[^>]*>", raw, flags=re.I)]
    srcs = [x for x in srcs if x]
    print("   外部 script %d 個" % len(srcs))
    for x in srcs[:30]:
        print("   %s" % urljoin(r.url, x))
    inline = [b for b in re.findall(r"<script\b[^>]*>(.*?)</script>", raw, flags=re.S | re.I) if b.strip()]
    print("   內嵌 script %d 個｜大小（字元）：%s" % (len(inline), [len(b) for b in inline][:30]))
    apis = []
    for b in inline:
        for x in API_PAT.findall(b):
            if x not in apis and len(x) > 4:
                apis.append(x)
    print("   疑似資料介面 %d 個" % len(apis))
    for x in apis[:40]:
        print("   %s" % x)
    for k, b in enumerate(inline, 1):
        if "%" in b and re.search(r"\d{2,3}%", b):
            print("   內嵌 script %d 含百分比數值，節錄：%s" % (k, re.sub(r"\s+", " ", b)[:600]))

    print("\n=== 8. 頁面文字（節錄）===")
    body = re.search(r"<body\b.*</body>", raw, flags=re.S | re.I)
    txt = flat(body.group(0) if body else raw)
    print("   文字總長 %d 字" % len(txt))
    i = txt.find("分紅實現率")
    print("   %s" % txt[max(0, i):max(0, i) + 2500])

    print("\n=== 完 ===")


if __name__ == "__main__":
    main()
