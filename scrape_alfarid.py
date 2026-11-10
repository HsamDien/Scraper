"""
استخراج منتجات وأسعار أسواق الفريد (طلبات - الياسمين) إلى ملف إكسل.

التثبيت (مرة واحدة):
    pip install playwright openpyxl
    playwright install chromium

التشغيل:
    python scrape_alfarid.py alfarid_products_ar.xlsx

يقرأ الروابط من ورقة "الأقسام" في الملف، ويكتب النتيجة في alfarid_all_products.xlsx
ويحفظ تقدمه كل قسم، فإذا انقطع التشغيل أعد تشغيله وسيكمل من حيث توقف.
"""
import re
import sys
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

SRC = sys.argv[1] if len(sys.argv) > 1 else "alfarid_products_ar.xlsx"
OUT = "alfarid_all_products.xlsx"
MAX_PAGES = 40
PRICE_RE = re.compile(r"JOD\s*([\d.,]+)")
STOCK_RE = re.compile(r"^(?:loading\s*)?(?:فقط\s*\d+\s*صنف\s*في\s*المخزن\s*\d+)?\s*", re.I)


def parse_card(text, href):
    text = " ".join(text.split())
    prices = PRICE_RE.findall(text)
    if not prices:
        return None
    name = text[: text.find("JOD")].strip()
    name = STOCK_RE.sub("", name).strip()
    name = re.sub(r"^loading\s*", "", name, flags=re.I).strip()
    if not name:
        return None
    current = float(prices[-1].replace(",", ""))
    original = float(prices[0].replace(",", "")) if len(prices) > 1 else None
    return name, current, original, href


def scrape_category(page, url):
    items, seen_first = [], None
    for n in range(1, MAX_PAGES + 1):
        u = url if n == 1 else f"{url}&page={n}"
        page.goto(u, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_selector('a[href*="/product/"]', timeout=12000)
        except PWTimeout:
            break  # لا منتجات في هذه الصفحة
        page.wait_for_timeout(800)
        cards = page.eval_on_selector_all(
            'a[href*="/product/"]',
            "els => els.map(e => [e.innerText, e.href])",
        )
        rows = [r for r in (parse_card(t, h) for t, h in cards) if r]
        if not rows:
            break
        first = rows[0][3]
        if first == seen_first:  # الموقع أعاد نفس الصفحة
            break
        seen_first = first
        items.extend(rows)
        # آخر صفحة؟ إذا لا يوجد رابط للصفحة التالية
        if not page.query_selector(f'a[href*="page={n + 1}"]'):
            break
    return items


def main():
    wb_in = load_workbook(SRC)
    ws_in = wb_in["الأقسام"]
    cats = [(r[0], r[1], r[3]) for r in ws_in.iter_rows(min_row=2, values_only=True) if r[3]]

    done = set()
    if Path(OUT).exists():
        wb = load_workbook(OUT)
        ws = wb.active
        done = {(r[0], r[1]) for r in ws.iter_rows(min_row=2, values_only=True)}
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "المنتجات"
        ws.append(["القسم الرئيسي", "القسم الفرعي", "المنتج", "السعر (دينار)", "السعر قبل الخصم", "رابط المنتج"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_context(locale="ar-JO").new_page()
        for i, (main_ar, sub_ar, url) in enumerate(cats, 1):
            if (main_ar, sub_ar) in done:
                continue
            print(f"[{i}/{len(cats)}] {main_ar} > {sub_ar}", flush=True)
            try:
                items = scrape_category(page, url)
            except Exception as e:  # noqa: BLE001
                print("   خطأ:", e)
                continue
            if not items:
                ws.append([main_ar, sub_ar, "(لا توجد منتجات)", None, None, url])
            for name, cur, orig, href in items:
                ws.append([main_ar, sub_ar, name, cur, orig, href])
            print(f"   {len(items)} منتج")
            wb.save(OUT)
            time.sleep(1)
        browser.close()

    # تنسيق نهائي
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="F55905")
        c.alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDEF", [30, 26, 44, 14, 16, 60]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True
    wb.save(OUT)
    print("تم الحفظ في", OUT)


if __name__ == "__main__":
    main()
