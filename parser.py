"""
Продвинутый парсер книг с сайта https://books.toscrape.com

Заходит на страницу КАЖДОЙ книги и собирает: название, категорию, цену,
точное количество на складе, рейтинг (1–5), UPC, описание, ссылку на обложку
и ссылку на страницу книги.

Результат сохраняется в три файла: books.xlsx (с листом «Сводка»),
books.csv (UTF-8 с BOM для Excel) и books.json.

Примеры запуска:
    python parser.py                      # весь каталог (1000 книг)
    python parser.py --category Poetry    # только одна категория
    python parser.py --limit 20           # тест на первых 20 книгах
"""

import argparse
import csv
import json
import random
import re
import sys
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

# Главная страница сайта: с неё берём список категорий
BASE_URL = "https://books.toscrape.com/"

# Первая страница общего каталога (все 1000 книг)
CATALOG_URL = "https://books.toscrape.com/catalogue/page-1.html"

# Имена итоговых файлов
XLSX_FILE = "books.xlsx"
CSV_FILE = "books.csv"
JSON_FILE = "books.json"

# Сколько раз пробовать загрузить страницу, прежде чем сдаться
MAX_RETRIES = 3

# Пауза между запросами (в секундах) — случайное число из этого диапазона,
# чтобы не нагружать сайт
DELAY_RANGE = (0.3, 0.7)

# Сколько секунд ждать ответа сервера, прежде чем считать запрос неудачным
TIMEOUT = 10

# На сайте рейтинг записан словом в CSS-классе: <p class="star-rating Three">.
# Этот словарь переводит слово в число.
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# Заголовок User-Agent: представляемся браузером, а не «python-requests»
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}

# Колонки итоговой таблицы: пары «ключ в словаре книги» -> «заголовок».
# Один и тот же порядок используется в Excel и CSV.
COLUMNS = [
    ("title", "Название"),
    ("category", "Категория"),
    ("price", "Цена (£)"),
    ("stock", "На складе (шт.)"),
    ("rating", "Рейтинг"),
    ("upc", "UPC"),
    ("description", "Описание"),
    ("image_url", "Обложка"),
    ("url", "Ссылка"),
]


def pause():
    """Случайная пауза между запросами."""
    time.sleep(random.uniform(*DELAY_RANGE))


def fetch_page(session, url):
    """
    Загружает страницу и возвращает её HTML-текст.
    Если загрузка не удалась — повторяет до MAX_RETRIES раз.
    Если все попытки неудачны — возвращает None.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, headers=HEADERS, timeout=TIMEOUT)
            # raise_for_status() выбросит исключение, если код ответа 4xx или 5xx
            response.raise_for_status()
            # Сайт отдаёт UTF-8, но не всегда сообщает об этом в заголовках.
            # Без этой строки символ «£» превращается в «Â£».
            response.encoding = "utf-8"
            return response.text
        except requests.RequestException as error:
            print(f"  ! Ошибка при загрузке (попытка {attempt}/{MAX_RETRIES}): {error}")
            if attempt < MAX_RETRIES:
                # Перед повтором ждём чуть дольше с каждой попыткой
                time.sleep(attempt * 2)
    return None


# ---------------------------------------------------------------------------
# Шаг 1. Список ссылок на книги
# ---------------------------------------------------------------------------

def get_categories(session):
    """
    Возвращает словарь {название категории: ссылка на её первую страницу}.
    Категории берутся из бокового меню на главной странице.
    """
    html = fetch_page(session, BASE_URL)
    if html is None:
        return None
    soup = BeautifulSoup(html, "html.parser")
    categories = {}
    # Боковое меню: ul.nav-list > li > ul > li > a — это и есть категории
    for link in soup.select("ul.nav-list ul a"):
        name = link.get_text(strip=True)
        categories[name] = urljoin(BASE_URL, link["href"])
    return categories


def collect_book_links(session, start_url, limit=None):
    """
    Обходит страницы списка (каталога или категории) по кнопке «next»
    и возвращает список ссылок на страницы книг.
    Если задан limit — останавливается, как только набрано нужное число ссылок.
    """
    links = []
    url = start_url
    page_number = 1

    while url:
        print(f"Сбор ссылок, страница {page_number}: {url}")
        html = fetch_page(session, url)
        if html is None:
            # Без страницы не узнать ссылку «next», поэтому дальше не идём
            print(f"  ! Страница {page_number} не загрузилась — сбор ссылок остановлен")
            break

        soup = BeautifulSoup(html, "html.parser")
        # Каждая книга в списке — это <article class="product_pod">,
        # ссылка на её страницу лежит в <h3><a href="...">
        for card in soup.select("article.product_pod"):
            links.append(urljoin(url, card.h3.a["href"]))

        if limit and len(links) >= limit:
            return links[:limit]

        # Ищем кнопку «next». На последней странице её нет — цикл закончится.
        next_link = soup.select_one("li.next a")
        if next_link:
            url = urljoin(url, next_link["href"])
            page_number += 1
            pause()
        else:
            url = None

    return links


# ---------------------------------------------------------------------------
# Шаг 2. Разбор страницы одной книги
# ---------------------------------------------------------------------------

def parse_book_page(html, url):
    """
    Достаёт все данные книги с её собственной страницы.
    """
    soup = BeautifulSoup(html, "html.parser")
    main = soup.select_one("div.product_main")

    title = main.h1.get_text(strip=True)

    # Хлебные крошки: Home / Books / <Категория> / <Название книги>.
    # Категория — предпоследний пункт.
    crumbs = soup.select("ul.breadcrumb li")
    category = crumbs[-2].get_text(strip=True) if len(crumbs) >= 2 else ""

    # Цена в виде "£51.77": убираем знак фунта и превращаем в число
    price_text = main.select_one("p.price_color").get_text(strip=True)
    price = float(price_text.replace("£", ""))

    # Наличие: "In stock (22 available)" — вытаскиваем число из скобок.
    # Если числа нет (книги нет в наличии) — считаем, что 0.
    availability_text = main.select_one("p.availability").get_text(" ", strip=True)
    match = re.search(r"\((\d+) available\)", availability_text)
    stock = int(match.group(1)) if match else 0

    # Рейтинг: второй CSS-класс у <p class="star-rating Three">
    rating_tag = main.select_one("p.star-rating")
    rating_word = next((c for c in rating_tag["class"] if c in RATING_WORDS), None)
    rating = RATING_WORDS.get(rating_word)

    # Таблица «Product Information»: строки <th>Название</th><td>Значение</td>.
    # Собираем её в словарь и берём из него UPC.
    info = {
        row.th.get_text(strip=True): row.td.get_text(strip=True)
        for row in soup.select("table.table-striped tr")
    }
    upc = info.get("UPC", "")

    # Описание — это <p> сразу после блока <div id="product_description">.
    # У некоторых книг описания нет — тогда оставляем пустую строку.
    description = ""
    description_header = soup.select_one("#product_description")
    if description_header:
        description_tag = description_header.find_next_sibling("p")
        if description_tag:
            description = description_tag.get_text(strip=True)
            # Сайт дописывает в конец каждого описания « ...more» — убираем
            description = re.sub(r"\s*\.\.\.more$", "", description)

    # Обложка: ссылка вида "../../media/cache/....jpg" — делаем из неё полный адрес
    image_tag = soup.select_one("#product_gallery img")
    image_url = urljoin(url, image_tag["src"]) if image_tag else ""

    return {
        "title": title,
        "category": category,
        "price": price,
        "stock": stock,
        "rating": rating,
        "upc": upc,
        "description": description,
        "image_url": image_url,
        "url": url,
    }


def scrape_books(session, links, books):
    """
    Заходит на страницу каждой книги из списка links и добавляет результат в books.
    Список books передаётся снаружи: если парсинг прервут (Ctrl+C),
    в нём останется всё, что уже успели собрать.
    Возвращает список ссылок, которые так и не удалось обработать.
    """
    failed = []
    total = len(links)

    for number, url in enumerate(links, start=1):
        print(f"Книга {number} из {total}: {url}")
        html = fetch_page(session, url)

        if html is None:
            print("  ! Страница не загрузилась — книга пропущена")
            failed.append(url)
        else:
            try:
                books.append(parse_book_page(html, url))
            except (AttributeError, KeyError, ValueError, TypeError) as error:
                # Если вёрстка страницы неожиданная — не падаем, а пропускаем книгу
                print(f"  ! Не удалось разобрать страницу: {error}")
                failed.append(url)

        # Пауза перед следующим запросом (после последней книги не нужна)
        if number < total:
            pause()

    return failed


# ---------------------------------------------------------------------------
# Шаг 3. Сохранение результатов
# ---------------------------------------------------------------------------

def style_header(ws):
    """Оформляет первую строку листа как заголовок таблицы."""
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    # Закрепляем строку заголовков (она не уезжает при прокрутке)
    ws.freeze_panes = "A2"
    # Включаем фильтр-стрелочки в заголовках
    ws.auto_filter.ref = ws.dimensions


def auto_width(ws, rows, max_width=60):
    """
    Подбирает ширину колонок по самому длинному значению.
    rows — список строк таблицы вместе с заголовком.
    """
    for col_index in range(1, len(rows[0]) + 1):
        max_length = max(len(str(row[col_index - 1])) for row in rows)
        ws.column_dimensions[get_column_letter(col_index)].width = min(max_length + 2, max_width)


def build_summary(books):
    """
    Считает по каждой категории количество книг и среднюю цену.
    Возвращает список строк [категория, количество, средняя цена],
    отсортированный по количеству книг (больше — выше).
    """
    groups = {}
    for book in books:
        groups.setdefault(book["category"], []).append(book["price"])

    summary = [
        [category, len(prices), round(sum(prices) / len(prices), 2)]
        for category, prices in groups.items()
    ]
    summary.sort(key=lambda row: (-row[1], row[0]))
    return summary


def save_to_excel(books, filename):
    """
    Сохраняет книги в Excel: лист «Книги» со всеми данными
    и лист «Сводка» с количеством книг и средней ценой по категориям.
    """
    wb = Workbook()

    # --- Лист 1: все книги ---
    ws = wb.active
    ws.title = "Книги"
    header = [title for _, title in COLUMNS]
    rows = [[book[key] for key, _ in COLUMNS] for book in books]

    ws.append(header)
    for row in rows:
        ws.append(row)
    style_header(ws)

    # Формат цены: два знака после запятой
    price_col = get_column_letter([key for key, _ in COLUMNS].index("price") + 1)
    for cell in ws[price_col][1:]:
        cell.number_format = "0.00"

    auto_width(ws, [header] + rows)
    # Описание бывает очень длинным — фиксируем ширину, чтобы не растягивать лист
    desc_col = get_column_letter([key for key, _ in COLUMNS].index("description") + 1)
    ws.column_dimensions[desc_col].width = 60

    # --- Лист 2: сводка по категориям ---
    ws_summary = wb.create_sheet("Сводка")
    summary_header = ["Категория", "Книг", "Средняя цена (£)"]
    summary_rows = build_summary(books)

    ws_summary.append(summary_header)
    for row in summary_rows:
        ws_summary.append(row)
    style_header(ws_summary)

    for cell in ws_summary["C"][1:]:
        cell.number_format = "0.00"

    # Итоговая строка по всему каталогу (без фильтра — она ниже таблицы)
    if books:
        total_avg = round(sum(b["price"] for b in books) / len(books), 2)
        ws_summary.append([])
        ws_summary.append(["Итого", len(books), total_avg])
        last = ws_summary.max_row
        for cell in ws_summary[last]:
            cell.font = Font(bold=True)
        ws_summary.cell(row=last, column=3).number_format = "0.00"

    auto_width(ws_summary, [summary_header] + summary_rows)

    wb.save(filename)


def save_to_csv(books, filename):
    """
    Сохраняет книги в CSV. Кодировка utf-8-sig добавляет в начало файла
    метку BOM — по ней Excel понимает, что файл в UTF-8, и не ломает символы.
    Разделитель «;» — его Excel с русскими настройками понимает сразу.
    """
    with open(filename, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow([title for _, title in COLUMNS])
        for book in books:
            writer.writerow([book[key] for key, _ in COLUMNS])


def save_to_json(books, filename):
    """Сохраняет книги в JSON. ensure_ascii=False оставляет текст читаемым."""
    with open(filename, "w", encoding="utf-8") as file:
        json.dump(books, file, ensure_ascii=False, indent=2)


def save_all(books):
    """Сохраняет результат во все три формата."""
    if not books:
        print("Нечего сохранять — не собрано ни одной книги")
        return
    save_to_excel(books, XLSX_FILE)
    save_to_csv(books, CSV_FILE)
    save_to_json(books, JSON_FILE)
    print(f"Результат сохранён: {XLSX_FILE}, {CSV_FILE}, {JSON_FILE}")


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

def parse_args():
    """Разбирает аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Продвинутый парсер books.toscrape.com (данные с каждой страницы книги)"
    )
    parser.add_argument(
        "--category",
        help='парсить только эту категорию, например --category Poetry '
             '(для названий с пробелами: --category "Science Fiction")',
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="собрать только первые N книг (удобно для теста)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    if args.limit is not None and args.limit < 1:
        print("--limit должен быть положительным числом")
        sys.exit(1)

    books = []  # сюда складываются книги; при Ctrl+C сохраним то, что есть
    failed = []
    interrupted = False
    started = time.time()

    # Session переиспользует одно соединение для всех запросов — так быстрее
    with requests.Session() as session:
        try:
            # Определяем, с какой страницы начинать: каталог целиком или категория
            if args.category:
                categories = get_categories(session)
                if categories is None:
                    print("Не удалось загрузить список категорий")
                    sys.exit(1)
                # Ищем категорию без учёта регистра: poetry == Poetry
                found = {name.lower(): name for name in categories}
                name = found.get(args.category.strip().lower())
                if name is None:
                    print(f"Категория «{args.category}» не найдена. Доступные категории:")
                    print(", ".join(sorted(categories)))
                    sys.exit(1)
                start_url = categories[name]
                print(f"Старт парсинга категории «{name}»\n")
                pause()
            else:
                start_url = CATALOG_URL
                print("Старт парсинга всего каталога books.toscrape.com\n")

            links = collect_book_links(session, start_url, args.limit)
            print(f"\nНайдено книг: {len(links)}. Заходим на страницу каждой книги...\n")
            pause()

            failed = scrape_books(session, links, books)

        except KeyboardInterrupt:
            interrupted = True
            print("\n\nПарсинг прерван (Ctrl+C) — сохраняем то, что уже собрано...")

    elapsed = time.time() - started
    print(f"\nСобрано книг: {len(books)} за {elapsed:.0f} сек.")
    if failed:
        print(f"Не удалось обработать: {len(failed)}")
        for url in failed:
            print(f"  - {url}")

    save_all(books)
    if interrupted:
        sys.exit(130)  # стандартный код выхода для прерывания по Ctrl+C


# Этот блок выполняется только при запуске файла напрямую (python parser.py),
# но не при импорте его в другой скрипт
if __name__ == "__main__":
    main()
