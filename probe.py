import io
import re
import sys
import time
from urllib.parse import urljoin

import pdfplumber
import requests
from bs4 import BeautifulSoup

PAGE = "https://www.hklife.com.hk/tc/customer-info/fulfillment-ratios/index.html"
KEYS = ("fulfil", "實現率", "現金價值", "cash-value", "cashvalue", "tcv")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
PCT = re.compile(r"\d+(?:\.\d+)?\s*[%％]")
PDF_URL = re.compile(r"""["']([^"'<>\s]+?\.pdf)["']""", re.I)
MAX_DOCS = 8
MAX_PAGES = 30


def get(url):
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60)
            print(f"HTTP {r.status_code}｜{len(r.content)} bytes｜{r.headers.get('content-type')}"
                  f"｜last-modified {r.headers.get('last-modified')}｜{url}")
            if r.status_code == 200:
                return r
        except requests.RequestException as e:
            print(f"{type(e).__name__}｜{url}")
        time.sleep(5)
    return None


def short(s, n=18):
    return re.sub(r"\s+", " ", s or "").strip()[:n]


def dump_pdf(content):
    with pdfplumber.open(io.BytesIO(content)) as doc:
        print(f"PDF 共 {len(doc.pages)} 頁")
        total = 0
        for no, page in enumerate(doc.pages[:MAX_PAGES], 1):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            total += len(PCT.findall(text))
            print(f"\n== 第 {no} 頁｜文字 {len(text)} 字｜百分比 {len(PCT.findall(text))} 個"
                  f"｜不適用／N/A {text.count('不適用') + len(re.findall(r'N/?A', text))} 個｜表格 {len(tables)} 個")
            print(text[:1200 if no <= 2 else 400])
            for t in tables:
                print(f"-- 表格：{len(t)} 行 x {max(len(x) for x in t)} 欄")
                for row in t[:8]:
                    print("   ", [short(c).replace("\n", "⏎") for c in row])
        print(f"\n合計百分比 {total} 個")


def dump_html(url, content):
    html = content.decode("utf-8", errors="replace")
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.extract()
    text = soup.get_text("\n", strip=True)
    tables = soup.find_all("table")
    print(f"網頁文字 {len(text)} 字｜百分比 {len(PCT.findall(text))} 個｜不適用 {text.count('不適用')} 個"
          f"｜表格 {len(tables)} 個｜iframe {len(soup.find_all('iframe'))} 個")
    at = text.find("實現率")
    print("--- 文字節錄 ---")
    print(text[max(at, 0):max(at, 0) + 1500])
    for i, t in enumerate(tables[:12], 1):
        rows = t.find_all("tr")
        print(f"-- 表格 {i}：{len(rows)} 行")
        for tr in rows[:8]:
            print("   ", [short(c.get_text()) + (f"[c{c.get('colspan')}]" if c.get("colspan") else "")
                          + (f"[r{c.get('rowspan')}]" if c.get("rowspan") else "")
                          for c in tr.find_all(["th", "td"])])
    for f in soup.find_all("iframe"):
        print("   iframe：", f.get("src"))
    for s in soup.find_all("select"):
        print("   下拉選單：", [short(o.get_text()) for o in s.find_all("option")][:15])
    found = []
    for a in soup.find_all("a", href=True):
        label = short(a.get_text(), 40)
        href = a["href"]
        if any(k in href.lower() or k in label.lower() for k in KEYS) or href.lower().endswith(".pdf"):
            full = urljoin(url, href).split("#")[0]
            if full not in [u for u, _ in found] and full != url:
                found.append((full, label))
    for m in PDF_URL.finditer(html):
        full = urljoin(url, m.group(1))
        if full not in [u for u, _ in found]:
            found.append((full, "（原始碼內）"))
    print(f"--- 相關連結 {len(found)} 個 ---")
    for u, label in found[:40]:
        print(f"   {label}｜{u}")
    return found


def main():
    r = get(PAGE)
    if r is None:
        print("讀不到官方頁")
        sys.exit(1)
    found = dump_html(PAGE, r.content)
    picked = [u for u, label in found if any(k in u.lower() or k in label.lower() for k in KEYS)]
    for u in picked[:MAX_DOCS]:
        print(f"\n########## {u}")
        d = get(u)
        if d is None:
            print("讀不到")
        elif d.content.startswith(b"%PDF"):
            dump_pdf(d.content)
        else:
            dump_html(u, d.content)


if __name__ == "__main__":
    main()
