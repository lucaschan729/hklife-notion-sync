import datetime as dt
import hashlib
import json
import os
import re
import sys
import time
from collections import Counter

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PAGE = "https://www.hklife.com.hk/tc/customer-info/fulfillment-ratios/index.html"
COMPANY = "香港人壽"
PREFIX = f"{COMPANY}分紅實現率"
MARK = "資料指紋"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
YEAR_COL = re.compile(r"第\s*(\d+\+?)\s*個保單年度")
ISSUED = re.compile(r"[（(]\s*([\d\s\-–—至]+?)\s*年繕發")
NUM = re.compile(r"^(\d+(?:\.\d+)?)\s*[%％]")
HKT = dt.timezone(dt.timedelta(hours=8))

NAVY, LIGHT, GRAY = "2B4C7E", "DCE6F1", "E9EBF0"
THIN = Side(style="thin", color="B8C4D6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
WHITE_BOLD = Font(size=10, bold=True, color="FFFFFF")
BOLD = Font(size=10, bold=True)


# ---------- 讀取與解析 ----------

def fetch():
    for _ in range(3):
        try:
            r = requests.get(PAGE, headers=UA, timeout=60)
            print(f"HTTP {r.status_code}｜{len(r.content)} bytes｜last-modified {r.headers.get('last-modified')}")
            if r.status_code == 200:
                return r.content.decode("utf-8", errors="replace")
        except requests.RequestException as e:
            print(f"連線失敗：{type(e).__name__}")
        time.sleep(5)
    sys.exit("讀不到官方頁，停止")


def clean(s):
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


def span(v):
    return max(int(re.sub(r"\D", "", str(v or "")) or 1), 1)


def expand(table):
    out, carry = [], {}

    def fill(row, col):
        while col in carry:
            row.append(carry[col][0])
            carry[col][1] -= 1
            if carry[col][1] == 0:
                del carry[col]
            col += 1
        return col

    for tr in table.find_all("tr"):
        row, col = [], 0
        for cell in tr.find_all(["th", "td"]):
            col = fill(row, col)
            text = clean(cell.get_text())
            rs, cs = span(cell.get("rowspan")), span(cell.get("colspan"))
            for _ in range(cs):
                row.append(text)
                if rs > 1:
                    carry[col] = [text, rs - 1]
                col += 1
        fill(row, col)
        out.append(row)
    return out


def parse(html):
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.extract()
    text = soup.get_text(" ", strip=True)
    m = re.search(r"(20\d{2})\s*報告年度", text)
    if not m:
        sys.exit("找不到報告年度，停止")
    year = m.group(1)
    m = re.search(r"20\d{2}\s*年\s*\d{1,2}\s*月", text)
    updated = clean(m.group(0)) if m else "未能辨識"

    cols, rows = None, []
    blanks, notes, skipped = Counter(), Counter(), []
    for tno, table in enumerate(soup.find_all("table"), 1):
        sups = table.find_all("sup")
        for s in sups:
            s.decompose()
        for br in table.find_all("br"):
            br.replace_with("\n")
        g = expand(table)
        h = next((i for i, r in enumerate(g) if sum(bool(YEAR_COL.search(c)) for c in r) >= 3), None)
        if h is None:
            print(f"表格 {tno}：找不到保單年度標題，略過")
            continue
        ycols = [i for i, c in enumerate(g[h]) if YEAR_COL.search(c)]

        def find(key):
            for r in g[:h + 1]:
                for i, c in enumerate(r[:ycols[0]]):
                    if key in c:
                        return i
            return None

        ni, ti, ci = find("名稱"), find("種類"), find("貨幣")
        if ni is None or ci is None:
            print(f"表格 {tno}：找不到產品名稱或貨幣欄，略過")
            continue
        top = " ".join(r[ycols[0]] for r in g[:h] if len(r) > ycols[0])
        if "終期" in top:
            cat = "終期紅利"
        elif "週年" in top:
            cat = "週年紅利、特別紅利及非保證每月入息"
        else:
            cat = top or f"表格{tno}"

        labels = []
        for i in ycols:
            m = ISSUED.search(g[h][i])
            labels.append((YEAR_COL.search(g[h][i]).group(1), clean(m.group(1)) if m else ""))
        if cols is None:
            cols = labels
        elif [a for a, _ in cols] != [a for a, _ in labels]:
            sys.exit("兩個表格的保單年度欄不一致，停止")

        n = 0
        for r in g[h + 1:]:
            if len(set(r)) <= 1:
                if r and r[0]:
                    print(f"表格 {tno} 分節行（未納入）：{r[0][:40]}")
                continue
            if len(r) <= ycols[-1] or not r[ni] or not r[ci]:
                skipped.append((tno, r))
                continue
            vals = []
            for i in ycols:
                m = NUM.match(r[i])
                if m:
                    vals.append(round(float(m.group(1)) / 100, 6))
                    if r[i][m.end():].strip():
                        notes[r[i]] += 1
                else:
                    vals.append(None)
                    blanks[r[i] or "（空白）"] += 1
            rows.append({"name": r[ni], "type": r[ti] if ti is not None else "",
                         "cur": r[ci], "cat": cat, "vals": vals})
            n += 1
        print(f"表格 {tno}｜{cat}｜標題 {h + 1} 行｜數據 {n} 行｜年度欄 {len(ycols)} 個｜移除上標 {len(sups)} 個")

    if not rows:
        sys.exit("沒有讀到任何數據，停止")
    order = {}
    for r in rows:
        order.setdefault(r["name"], len(order))
    cats = list(dict.fromkeys(r["cat"] for r in rows))
    curs = list(dict.fromkeys(r["cur"] for r in rows))
    rows.sort(key=lambda r: (order[r["name"]], curs.index(r["cur"]), cats.index(r["cat"])))

    print("無數字格內容：", dict(blanks))
    if notes:
        print("數字後帶附註的格：", dict(notes))
    if skipped:
        print(f"未能辨識的行 {len(skipped)} 行，頭 3 行：")
        for tno, r in skipped[:3]:
            print(f"   表格 {tno}：{r}")
    return year, updated, cols, rows, cats


def avg(v):
    return sum(v) / len(v) if v else None


def summary(rows):
    v = [x for r in rows for x in r["vals"] if x is not None]
    return {"products": len({r["name"] for r in rows}), "rows": len(rows), "n": len(v),
            "empty": sum(len(r["vals"]) for r in rows) - len(v), "avg": avg(v),
            "min": min(v) if v else None, "max": max(v) if v else None,
            "low": sum(x < 0.8 - 1e-9 for x in v), "high": sum(x >= 1 - 1e-9 for x in v)}


def show(x):
    return "/" if x is None else f"{x:.2%}"


# ---------- Excel ----------

def put(ws, r, c, v, font=None, fill=None, align=CENTER, fmt=None, border=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or Font(size=10)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = align
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = BOX
    return cell


def band(ws, r, c1, c2, v, **kw):
    for c in range(c1, c2 + 1):
        put(ws, r, c, v if c == c1 else None, **kw)
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)


def put_pct(ws, r, c, v, fmt, colour=True):
    if v is None:
        return put(ws, r, c, "/", font=Font(size=10, color="A6A6A6"))
    font, fill = Font(size=10), None
    if colour and v >= 1 - 1e-9:
        font, fill = Font(size=10, color="006100"), "C6EFCE"
    elif colour and v < 0.8 - 1e-9:
        font, fill = Font(size=10, color="9C0006"), "FFC7CE"
    return put(ws, r, c, v, font=font, fill=fill, fmt=fmt)


def build(year, updated, cols, rows, cats, path, fp):
    wb = Workbook()
    groups = [("所有產品", rows)]
    if len(cats) > 1:
        groups += [(c, [r for r in rows if r["cat"] == c]) for c in cats]

    # 概覽
    ov = wb.active
    ov.title = "概覽"
    for i, w in enumerate([36, 10, 10, 10, 12, 12, 10, 10, 14, 16], 1):
        ov.column_dimensions[get_column_letter(i)].width = w
    band(ov, 1, 1, 10, f"{COMPANY}分紅實現率（{year} 報告年度）概覽",
         font=Font(size=14, bold=True, color="FFFFFF"), fill=NAVY, align=LEFT, border=False)
    ov.row_dimensions[1].height = 30
    info = [("公司", f"{COMPANY}保險有限公司"), ("報告年度", year), ("官方頁面標示日期", updated),
            ("官方來源", PAGE),
            ("生成時間（香港時間）", dt.datetime.now(HKT).strftime("%Y-%m-%d %H:%M")),
            (MARK, fp)]
    for i, (k, v) in enumerate(info, 3):
        put(ov, i, 1, k, font=BOLD, fill=LIGHT)
        band(ov, i, 2, 10, v, align=LEFT)
    r = 3 + len(info) + 1
    heads = ["紅利類別", "產品數", "行數", "數值格", "無數字格", "平均值", "最低", "最高", "低於80%格數", "100%或以上格數"]
    for c, t in enumerate(heads, 1):
        put(ov, r, c, t, font=WHITE_BOLD, fill=NAVY)
    ov.row_dimensions[r].height = 30
    for name, part in groups:
        r += 1
        s = summary(part)
        put(ov, r, 1, name, font=BOLD, fill=LIGHT)
        for c, k in enumerate(["products", "rows", "n", "empty"], 2):
            put(ov, r, c, s[k])
        for c, k in enumerate(["avg", "min", "max"], 6):
            put_pct(ov, r, c, s[k], "0.00%", colour=False)
        put(ov, r, 9, s["low"])
        put(ov, r, 10, s["high"])
        ov.row_dimensions[r].height = 30
    r += 2
    for t in ["數值照錄官方網頁；「/」＝官方標示不適用或沒有數字。",
              "綠底＝100% 或以上；紅底＝低於 80%。",
              "官方只披露 2010 年起繕發、報告年度內仍有生效保單的分紅產品。"]:
        band(ov, r, 1, 10, t, font=Font(size=9, color="595959"), align=LEFT, border=False)
        r += 1

    # 數據
    ws = wb.create_sheet(f"{year}分紅實現率")
    L = 4
    last = L + len(cols)
    for i, w in enumerate([36, 16, 8, 36] + [14] * len(cols), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    band(ws, 1, 1, last, f"{COMPANY}分紅實現率（{year} 報告年度）",
         font=Font(size=14, bold=True, color="FFFFFF"), fill=NAVY, align=LEFT, border=False)
    ws.row_dimensions[1].height = 30
    band(ws, 2, 1, last, "分紅實現率＝實際派發的非保證金額 ÷ 銷售時所示的金額。100% 即與銷售時一致，高於 100% 即比銷售時多。",
         font=Font(size=9, color="595959"), align=LEFT, border=False)
    band(ws, 3, 1, last, f"「/」＝官方標示不適用或沒有數字。來源：{PAGE}（{updated}）",
         font=Font(size=9, color="595959"), align=LEFT, border=False)
    hr = 5
    heads = ["產品名稱", "產品種類", "貨幣", "紅利類別"] + \
            [f"保單年度{y}" + (f"\n（{iss}年繕發）" if iss else "") for y, iss in cols]
    for c, t in enumerate(heads, 1):
        put(ws, hr, c, t, font=WHITE_BOLD, fill=NAVY)
    ws.row_dimensions[hr].height = 44
    allv = [x for row in rows for x in row["vals"] if x is not None]
    fmt = "0%" if all(abs(x * 100 - round(x * 100)) < 1e-6 for x in allv) else "0.00%"
    r, prev = hr + 1, None
    for row in rows:
        if prev is not None and row["name"] != prev:
            for c in range(1, last + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=GRAY)
            ws.row_dimensions[r].height = 6
            r += 1
        prev = row["name"]
        put(ws, r, 1, row["name"], font=BOLD, align=LEFT)
        put(ws, r, 2, row["type"])
        put(ws, r, 3, row["cur"], font=Font(size=10, bold=True, color="7030A0"))
        put(ws, r, 4, row["cat"])
        for i, v in enumerate(row["vals"]):
            put_pct(ws, r, L + 1 + i, v, fmt)
        ws.row_dimensions[r].height = 30
        r += 1
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(last)}{r - 1}"
    ws.freeze_panes = f"A{hr + 1}"

    # 平均值摘要
    r += 1
    put(ws, r, 1, "平均值摘要", font=Font(size=11, bold=True, color=NAVY), align=LEFT, border=False)
    r += 1
    band(ws, r, 1, L, "紅利類別", font=WHITE_BOLD, fill=NAVY)
    for i, (y, _) in enumerate(cols):
        put(ws, r, L + 1 + i, "所有年份平均值" if i == 0 else f"持滿{y}年平均值", font=WHITE_BOLD, fill=NAVY)
    ws.row_dimensions[r].height = 30
    for name, part in groups:
        r += 1
        band(ws, r, 1, L, name, font=BOLD, fill=LIGHT)
        for i in range(len(cols)):
            put_pct(ws, r, L + 1 + i, avg([x for row in part for x in row["vals"][i:] if x is not None]),
                    "0.00%", colour=False)
    r += 1
    band(ws, r, 1, L, "當年平均值", font=WHITE_BOLD, fill=NAVY)
    for i, (y, _) in enumerate(cols):
        put(ws, r, L + 1 + i, f"保單年度{y}", font=BOLD, fill=LIGHT)
    for name, part in groups:
        r += 1
        band(ws, r, 1, L, name, font=BOLD, fill=LIGHT)
        for i in range(len(cols)):
            put_pct(ws, r, L + 1 + i, avg([row["vals"][i] for row in part if row["vals"][i] is not None]),
                    "0.00%", colour=False)
    r += 1
    band(ws, r, 1, last, "「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。"
                         "平均值計入所有貨幣，為固定數字，不會隨篩選變動。",
         font=Font(size=9, color="595959"), align=LEFT, border=False)
    wb.save(path)


# ---------- Notion ----------

def notion(method, path, token, soft=False, **kw):
    h = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}
    r = requests.request(method, "https://api.notion.com/v1" + path, headers=h, timeout=120, **kw)
    if r.status_code >= 400:
        if soft:
            return None
        try:
            msg = r.json().get("message")
        except ValueError:
            msg = r.text[:300]
        sys.exit(f"Notion {method} {path.split('?')[0]} 失敗：HTTP {r.status_code}｜{msg}")
    return r.json()


def upload(path, year, fp):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        sys.exit("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    m = re.search(r"[0-9a-fA-F]{32}|[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}", raw)
    if not m:
        sys.exit("NOTION_DATABASE_ID 內找不到 32 位 ID")
    dbid = m.group(0)
    db = notion("GET", f"/databases/{dbid}", token, soft=True)
    dsid = db["data_sources"][0]["id"] if db and db.get("data_sources") else dbid
    ds = notion("GET", f"/data_sources/{dsid}", token)
    title = next(k for k, v in ds["properties"].items() if v["type"] == "title")
    q = notion("POST", f"/data_sources/{dsid}/query", token,
               json={"filter": {"property": title, "title": {"equals": year}}, "page_size": 5})
    if q["results"]:
        page = q["results"][0]["id"]
        print(f"找到 Notion「{year}」頁")
    else:
        page = notion("POST", "/pages", token, json={
            "parent": {"type": "data_source_id", "data_source_id": dsid},
            "properties": {title: {"title": [{"text": {"content": year}}]}}})["id"]
        print(f"已新增 Notion「{year}」頁")

    blocks, cursor = [], None
    while True:
        d = notion("GET", f"/blocks/{page}/children?page_size=100" +
                   (f"&start_cursor={cursor}" if cursor else ""), token)
        blocks += d["results"]
        if not d.get("has_more"):
            break
        cursor = d["next_cursor"]
    old, has_file, same = [], False, False
    for b in blocks:
        if b["type"] == "file":
            name = b["file"].get("name") or ""
            if name.startswith(PREFIX) and "自動生成" in name:
                old.append(b["id"])
                has_file = True
        elif b["type"] == "paragraph":
            txt = "".join(x.get("plain_text", "") for x in b["paragraph"]["rich_text"])
            if txt.startswith(MARK):
                old.append(b["id"])
                same = same or fp in txt
    if same and has_file and os.environ.get("FORCE", "").lower() != "true":
        print("數據與 Notion 現有檔案相同，略過上傳")
        return

    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    fu = notion("POST", "/file_uploads", token,
                json={"mode": "single_part", "filename": os.path.basename(path), "content_type": mime})
    with open(path, "rb") as f:
        notion("POST", f"/file_uploads/{fu['id']}/send", token, files={"file": ("upload.xlsx", f, mime)})
    stamp = dt.datetime.now(HKT).strftime("%Y-%m-%d %H:%M")
    notion("PATCH", f"/blocks/{page}/children", token, json={"children": [
        {"object": "block", "type": "file",
         "file": {"type": "file_upload", "file_upload": {"id": fu["id"]}}},
        {"object": "block", "type": "paragraph", "paragraph": {"rich_text": [{"type": "text", "text": {
            "content": f"{MARK}：{fp}｜{COMPANY} {year} 報告年度｜自動上傳 {stamp}（香港時間）"}}]}}]})
    for bid in old:
        notion("DELETE", f"/blocks/{bid}", token)
    print(f"已上傳到 Notion「{year}」頁；移除舊區塊 {len(old)} 個")


# ---------- 主程式 ----------

def main():
    year, updated, cols, rows, cats = parse(fetch())
    print(f"\n報告年度 {year}｜頁面標示 {updated}｜年度欄 {[y for y, _ in cols]}")
    for name, part in [("合計", rows)] + [(c, [r for r in rows if r["cat"] == c]) for c in cats]:
        s = summary(part)
        print(f"{name}：{s['products']} 產品／{s['rows']} 行／{s['n']} 數值／{s['empty']} 無數字格｜"
              f"平均 {show(s['avg'])}｜範圍 {show(s['min'])}–{show(s['max'])}｜"
              f"低於80% {s['low']} 格｜100%或以上 {s['high']} 格")
    print("\n頭 6 行（供抽查）：")
    for r in rows[:6]:
        print(f"   {r['name']}｜{r['cur']}｜{r['cat'][:4]}｜" + " ".join(show(x) for x in r["vals"]))

    total = summary(rows)
    if total["rows"] < 20 or total["n"] < 100:
        sys.exit("資料量異常偏少，可能官方頁改版，停止（不會上傳）")
    data = [year, cols, [[r[k] for k in ("name", "type", "cur", "cat", "vals")] for r in rows]]
    fp = hashlib.sha256(json.dumps(data, ensure_ascii=False).encode()).hexdigest()[:12]
    path = f"{PREFIX}（{year}年）自動生成.xlsx"
    build(year, updated, cols, rows, cats, path, fp)
    print(f"\n已生成 {path}｜{MARK} {fp}")

    if os.environ.get("UPLOAD", "").lower() != "true":
        print("只讀試跑：沒有上傳 Notion。Excel 可在本次執行的 Artifacts 下載。")
        return
    upload(path, year, fp)


if __name__ == "__main__":
    main()
