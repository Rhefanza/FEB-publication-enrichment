
import re
import time
import random
import argparse
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


BASE_URL = "https://sinta.kemdiktisaintek.go.id"

INPUT_FILE = r"D:\Proyek FEB\source data\unique_author_2_sumber_sinta.xlsx"
INPUT_SHEET = "Unique Authors"
AUTHOR_COLUMN = "Normalized Key"

OUTPUT_FILE = "books_sinta_2020_2026.xlsx"

START_YEAR = 2020
END_YEAR = 2026

USER_DATA_DIR = "sinta_browser_profile"

AFFILIATION_KEYWORDS = [
    "universitas airlangga",
    "univ airlangga",
    "airlangga",
    "airlngga",
    "unair",
]

DEPARTMENT_KEYWORDS = [
    "ilmu manajemen",
    "sains manajemen",
    "manajemen",

    "ilmu ekonomi islam",
    "sains ekonomi islam",
    "ekonomi islam",
    "ekonomi syariah",
    "ekonomi syari ah",
    "ekonomi syari'ah",

    "ilmu akuntansi",
    "sains akuntansi",
    "akuntansi",
    "akuntan",
    "pendidikan profesi akuntan",
    "pendidikan profesi akuntan profesi",

    "ilmu ekonomi",
    "sains ekonomi",
    "ekonomi pembangunan",

    "pengembangan sumber daya manusia",
]

UNKNOWN_DEPARTMENT_VALUES = [
    "",
    "-",
    "unknown",
    "none",
    "null",
    "n/a",
    "na",
]

FINAL_COLUMNS = [
    "Title",
    "Category",
    "Authors",
    "Publisher",
    "Tahun",
    "City",
    "ISBN",
]


def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


def normalize_text(text):
    text = clean_text(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_unknown_department(text):
    text_lower = normalize_text(text)
    return (
        text_lower == ""
        or text_lower in UNKNOWN_DEPARTMENT_VALUES
        or "unknown" in text_lower
    )


def is_affiliation_match(text):
    text_lower = normalize_text(text)
    normalized_keywords = [normalize_text(k) for k in AFFILIATION_KEYWORDS]
    return any(k in text_lower for k in normalized_keywords)


def is_department_match(text):
    text_lower = normalize_text(text)
    normalized_keywords = [normalize_text(k) for k in DEPARTMENT_KEYWORDS]
    return any(k in text_lower for k in normalized_keywords)


def extract_year(text):
    text = clean_text(text)
    match = re.search(r"\b(2020|2021|2022|2023|2024|2025|2026)\b", text)
    return int(match.group(1)) if match else None


def extract_isbn(text):
    text = clean_text(text)
    match = re.search(r"ISBN\s*:\s*([0-9Xx\-]+)", text, re.I)
    if match:
        return clean_text(match.group(1))
    return ""


def read_authors(limit=None):
    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)

    if AUTHOR_COLUMN not in df.columns:
        raise ValueError(
            f"Kolom '{AUTHOR_COLUMN}' tidak ditemukan. "
            f"Kolom tersedia: {list(df.columns)}"
        )

    authors = (
        df[AUTHOR_COLUMN]
        .dropna()
        .astype(str)
        .map(clean_text)
        .drop_duplicates()
        .tolist()
    )

    authors = [a for a in authors if a]

    if limit:
        authors = authors[:limit]

    return authors


def launch_context(playwright, headless=False):
    context = playwright.chromium.launch_persistent_context(
        USER_DATA_DIR,
        headless=headless,
        slow_mo=120,
        viewport={"width": 1366, "height": 768},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        args=[
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--start-maximized",
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ],
    )

    if context.pages:
        page = context.pages[0]
    else:
        page = context.new_page()

    return context, page


def manual_login(minutes=5):
    with sync_playwright() as p:
        context, page = launch_context(p, headless=False)

        page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)

        print("=" * 80)
        print("LOGIN MANUAL SINTA")
        print("1. Browser akan terbuka.")
        print("2. Silakan login ke akun SINTA.")
        print("3. Setelah login berhasil, biarkan browser tetap terbuka.")
        print(f"4. Script akan menunggu {minutes} menit lalu menyimpan session.")
        print("=" * 80)

        time.sleep(minutes * 60)

        context.close()

    print("Session login tersimpan di folder:", USER_DATA_DIR)


def extract_candidates_from_search_page(html):
    soup = BeautifulSoup(html, "html.parser")

    candidate_blocks = soup.select(".list-item")

    if not candidate_blocks:
        candidate_blocks = soup.select(".profile-side, .row, .card")

    candidates = []

    for block in candidate_blocks:
        block_text = clean_text(block.get_text(" ", strip=True))

        sinta_match = re.search(r"SINTA\s*ID\s*:\s*(\d+)", block_text, re.I)
        if not sinta_match:
            continue

        sinta_id = sinta_match.group(1)

        name_tag = block.select_one(".profile-name")
        affil_tag = block.select_one(".profile-affil")
        dept_tag = block.select_one(".profile-dept")

        candidate_name = clean_text(name_tag.get_text(" ", strip=True)) if name_tag else ""
        candidate_affil = clean_text(affil_tag.get_text(" ", strip=True)) if affil_tag else ""
        candidate_dept = clean_text(dept_tag.get_text(" ", strip=True)) if dept_tag else ""

        affil_check_text = candidate_affil if candidate_affil else block_text
        dept_check_text = candidate_dept if candidate_dept else block_text

        affil_ok = is_affiliation_match(affil_check_text)
        dept_ok = is_department_match(dept_check_text)
        dept_unknown = is_unknown_department(candidate_dept)

        candidates.append({
            "sinta_id": sinta_id,
            "name": candidate_name,
            "affil": candidate_affil,
            "dept": candidate_dept,
            "affil_ok": affil_ok,
            "dept_ok": dept_ok,
            "dept_unknown": dept_unknown,
            "block_text": block_text,
        })

    return candidates


def go_to_next_search_page(page):
    selectors = [
        "nav[aria-label='pagination-sample'] li.page-item:not(.disabled) a:has-text('Next')",
        "ul.pagination li.page-item:not(.disabled) a:has-text('Next')",
        "li.page-item:not(.disabled) a:has-text('Next')",
    ]

    for selector in selectors:
        try:
            next_btn = page.locator(selector)

            if next_btn.count() > 0 and next_btn.first.is_visible():
                old_text = clean_text(page.inner_text("body"))

                next_btn.first.click(timeout=5000)
                page.wait_for_timeout(random.randint(2000, 4000))

                new_text = clean_text(page.inner_text("body"))

                if old_text != new_text:
                    return True
        except Exception:
            pass

    return False


def choose_best_sinta_candidate(candidates, author_name):
    if not candidates:
        print(f"SKIP: SINTA ID tidak terbaca -> {author_name}")
        return ""

    unair_candidates = [c for c in candidates if c["affil_ok"]]

    if not unair_candidates:
        print(f"SKIP: Tidak ada afiliasi Universitas Airlangga -> {author_name}")
        return ""

    dept_matched = [c for c in unair_candidates if c["dept_ok"]]

    if dept_matched:
        c = dept_matched[0]

        print("  Kandidat dipilih: UNAIR + Prodi FEB")
        print(f"  Nama    : {c['name'] if c['name'] else '-'}")
        print(f"  Afiliasi: {c['affil'] if c['affil'] else '-'}")
        print(f"  Prodi   : {c['dept'] if c['dept'] else '-'}")
        print(f"  SINTA ID: {c['sinta_id']}")

        return c["sinta_id"]

    if len(unair_candidates) == 1:
        c = unair_candidates[0]

        if c["dept_unknown"]:
            print("  Kandidat dipilih: UNAIR tunggal prodi unknown/kosong")
            print(f"  Nama    : {c['name'] if c['name'] else '-'}")
            print(f"  Afiliasi: {c['affil'] if c['affil'] else '-'}")
            print(f"  Prodi   : {c['dept'] if c['dept'] else 'UNKNOWN / kosong'}")
            print(f"  SINTA ID: {c['sinta_id']}")

            return c["sinta_id"]

    print(f"SKIP: Kandidat UNAIR ada, tapi prodi tidak cocok FEB -> {author_name}")
    return ""


def search_sinta_id(page, author_name):
    search_url = f"{BASE_URL}/authors?q={quote_plus(author_name)}"

    page.goto(search_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(random.randint(1500, 3000))

    try:
        body_text = clean_text(page.inner_text("body"))
    except Exception:
        return ""

    if "Data Not Found" in body_text:
        print(f"SKIP: Data Not Found -> {author_name}")
        return ""

    all_candidates = []
    seen_page_signatures = set()

    for search_page_no in range(1, 30):
        print(f"  Cek hasil search author page {search_page_no}")

        page.wait_for_timeout(random.randint(1000, 2000))

        html = page.content()
        body_text = clean_text(page.inner_text("body"))
        signature = body_text[:2000]

        if signature in seen_page_signatures:
            print("  Stop search pagination: halaman berulang.")
            break

        seen_page_signatures.add(signature)

        page_candidates = extract_candidates_from_search_page(html)

        if page_candidates:
            print(f"  Kandidat di page ini: {len(page_candidates)}")
            all_candidates.extend(page_candidates)
        else:
            print("  Tidak ada kandidat terbaca di page ini.")

        has_next = go_to_next_search_page(page)

        if not has_next:
            print("  Stop search pagination: tidak ada tombol Next aktif.")
            break

    return choose_best_sinta_candidate(all_candidates, author_name)


def parse_books_page(html):
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    book_items = soup.select(".ar-list-item")

    for item in book_items:
        title_tag = item.select_one(".ar-title a")
        if not title_tag:
            continue

        title = clean_text(title_tag.get_text(" ", strip=True))
        item_text = clean_text(item.get_text(" ", strip=True))

        year = extract_year(item_text)
        if year is None:
            continue

        if year < START_YEAR or year > END_YEAR:
            continue

        category = ""
        authors = ""
        publisher = ""
        city = ""
        isbn = ""

        meta_blocks = item.select(".ar-meta")

        for meta in meta_blocks:
            meta_text = clean_text(meta.get_text(" ", strip=True))

            if "Category" in meta_text:
                category_match = re.search(r"Category\s*:\s*(.*)", meta_text, re.I)
                if category_match:
                    category = clean_text(category_match.group(1))

            pub_tag = meta.select_one(".ar-pub")
            if pub_tag:
                publisher = clean_text(pub_tag.get_text(" ", strip=True))

            year_tag = meta.select_one(".ar-year")
            if year_tag:
                y = extract_year(year_tag.get_text(" ", strip=True))
                if y:
                    year = y

            city_tag = meta.select_one(".ar-cited")
            if city_tag:
                city = clean_text(city_tag.get_text(" ", strip=True))

            quartile_tag = meta.select_one(".ar-quartile")
            if quartile_tag:
                q_text = clean_text(quartile_tag.get_text(" ", strip=True))
                if "ISBN" in q_text.upper():
                    isbn = extract_isbn(q_text)

        # Authors biasanya meta tanpa class khusus setelah category
        if not authors:
            for meta in meta_blocks:
                meta_text = clean_text(meta.get_text(" ", strip=True))

                if (
                    "Category" not in meta_text
                    and "ISBN" not in meta_text.upper()
                    and not meta.select_one(".ar-year")
                    and not meta.select_one(".ar-cited")
                    and not meta.select_one(".ar-quartile")
                ):
                    # Hapus publisher dari teks jika ikut terbaca
                    if publisher and publisher in meta_text:
                        meta_text = meta_text.replace(publisher, "")
                    authors = clean_text(meta_text)
                    break

        # Fallback
        if not category:
            category_match = re.search(r"Category\s*:\s*(.*?)(?=[A-Z][a-z]+|$)", item_text, re.I)
            if category_match:
                category = clean_text(category_match.group(1))

        if not isbn:
            isbn = extract_isbn(item_text)

        rows.append({
            "Title": title,
            "Category": category,
            "Authors": authors,
            "Publisher": publisher,
            "Tahun": year,
            "City": city,
            "ISBN": isbn,
        })

    return rows


def go_to_next_page(page):
    selectors = [
        "nav[aria-label='pagination-sample'] li.page-item:not(.disabled) a:has-text('Next')",
        "ul.pagination li.page-item:not(.disabled) a:has-text('Next')",
        "li.page-item:not(.disabled) a:has-text('Next')",
    ]

    for selector in selectors:
        try:
            next_btn = page.locator(selector)
            if next_btn.count() > 0 and next_btn.first.is_visible():
                old_text = clean_text(page.inner_text("body"))

                next_btn.first.click(timeout=5000)
                page.wait_for_timeout(random.randint(2500, 4500))

                new_text = clean_text(page.inner_text("body"))

                if old_text != new_text:
                    return True
        except Exception:
            pass

    return False


def scrape_books_by_sinta_id(page, sinta_id):
    books_url = f"{BASE_URL}/authors/profile/{sinta_id}/?view=books"

    page.goto(books_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(random.randint(2500, 4500))

    all_rows = []
    seen_page_signatures = set()

    for page_no in range(1, 50):
        print(f"  Ambil Books page {page_no}")

        page.wait_for_timeout(random.randint(1000, 2000))

        html = page.content()
        signature = clean_text(page.inner_text("body"))[:2000]

        if signature in seen_page_signatures:
            print("  Stop pagination: halaman berulang.")
            break

        seen_page_signatures.add(signature)

        rows = parse_books_page(html)
        all_rows.extend(rows)

        has_next = go_to_next_page(page)

        if not has_next:
            print("  Stop pagination: tidak ada tombol Next aktif.")
            break

    return all_rows


def save_output(all_rows):
    df_out = pd.DataFrame(all_rows, columns=FINAL_COLUMNS)

    if not df_out.empty:
        df_out = df_out.drop_duplicates(
            subset=[
                "Title",
                "Authors",
                "Publisher",
                "Tahun",
                "ISBN",
            ],
            keep="first"
        )

        df_out = df_out.sort_values(
            by=["Tahun", "Title"],
            ascending=[False, True]
        )

    df_out.to_excel(OUTPUT_FILE, index=False)

    print("=" * 80)
    print("Selesai.")
    print("Total baris output:", len(df_out))
    print("File tersimpan:", OUTPUT_FILE)

    return df_out


def scrape(limit=None, headless=False):
    authors = read_authors(limit=limit)

    print("Total author diproses:", len(authors))

    all_rows = []

    with sync_playwright() as p:
        context, page = launch_context(p, headless=headless)

        for idx, author_name in enumerate(authors, start=1):
            print("=" * 80)
            print(f"[{idx}/{len(authors)}] Search author: {author_name}")

            try:
                sinta_id = search_sinta_id(page, author_name)

                if not sinta_id:
                    continue

                rows = scrape_books_by_sinta_id(
                    page=page,
                    sinta_id=sinta_id
                )

                if not rows:
                    print(f"Tidak ada Books 2020-2026: {author_name}")
                    continue

                all_rows.extend(rows)
                print(f"Books diambil: {len(rows)}")

                page.wait_for_timeout(random.randint(2000, 5000))

            except PlaywrightTimeoutError:
                print(f"TIMEOUT: {author_name}")
                continue

            except Exception as e:
                print(f"ERROR pada author {author_name}: {e}")
                continue

        context.close()

    save_output(all_rows)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--login",
        action="store_true",
        help="Buka browser untuk login manual SINTA dan simpan session."
    )

    parser.add_argument(
        "--login-minutes",
        type=int,
        default=5,
        help="Durasi tunggu login manual dalam menit."
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Batasi jumlah author untuk testing."
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help="Jalankan browser tanpa tampilan."
    )

    args = parser.parse_args()

    if args.login:
        manual_login(minutes=args.login_minutes)
    else:
        scrape(limit=args.limit, headless=args.headless)
