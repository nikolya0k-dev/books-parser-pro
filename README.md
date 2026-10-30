# Продвинутый парсер книг с books.toscrape.com

**Полный сбор данных о каждом товаре — в Excel, CSV и JSON.** Парсер заходит на
страницу каждой из 1000 книг и собирает подробную информацию: категорию, точный
остаток на складе, артикул (UPC), описание, ссылку на обложку и другое. Можно
парсить весь каталог или только одну категорию. На выходе — три файла:
`books.xlsx` (с отдельным листом-сводкой по категориям), `books.csv` (сразу
открывается в Excel без «кракозябр») и `books.json` (для программ и API).
Парсер бережно относится к сайту (паузы между запросами), устойчив к сбоям сети
(повторные попытки) и не теряет данные при остановке по Ctrl+C.

Стек: Python, requests, BeautifulSoup, openpyxl.

Это развитие [первой версии парсера](https://github.com/nikolya0k-dev/books-parser),
которая собирала только данные с карточек в списке каталога.

## Пример результата

Первые 5 строк из [books.xlsx](books.xlsx) (всего 1000 строк, описание сокращено):

| Название | Категория | Цена (£) | На складе | Рейтинг | UPC | Описание |
|---|---|---:|---:|:---:|---|---|
| [A Light in the Attic](https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html) | Poetry | 51.77 | 22 | 3 | a897fe39b1053632 | It's hard to imagine a world without A Light in the Attic.… |
| [Tipping the Velvet](https://books.toscrape.com/catalogue/tipping-the-velvet_999/index.html) | Historical Fiction | 53.74 | 20 | 1 | 90fa61229261140a | "Erotic and absorbing...Written with starling power."--"The… |
| [Soumission](https://books.toscrape.com/catalogue/soumission_998/index.html) | Fiction | 50.10 | 20 | 1 | 6957f44c3847a760 | Dans une France assez proche de la nôtre, un homme s’engage… |
| [Sharp Objects](https://books.toscrape.com/catalogue/sharp-objects_997/index.html) | Mystery | 47.82 | 20 | 4 | e00eb4fd7b871a48 | WICKED above her hipbone, GIRL across her heart Words are… |
| [Sapiens: A Brief History of Humankind](https://books.toscrape.com/catalogue/sapiens-a-brief-history-of-humankind_996/index.html) | History | 54.23 | 20 | 5 | 4165285e1663650f | From a renowned historian comes a groundbreaking narrative… |

Лист «Сводка» — количество книг и средняя цена по каждой категории
(первые 5 строк из 50 категорий):

| Категория | Книг | Средняя цена (£) |
|---|---:|---:|
| Default | 152 | 34.39 |
| Nonfiction | 110 | 34.26 |
| Sequential Art | 75 | 34.57 |
| Add a comment | 67 | 35.80 |
| Fiction | 65 | 36.07 |

> «Default» и «Add a comment» — это настоящие названия категорий на сайте
> (учебный магазин так устроен), парсер их не придумывает.

## Что делает

1. Проходит по страницам каталога (или выбранной категории) по кнопке «next»
   и собирает ссылки на все книги.
2. Заходит на страницу каждой книги и собирает:

| Колонка          | Пример                                                        |
|------------------|---------------------------------------------------------------|
| Название         | A Light in the Attic                                          |
| Категория        | Poetry (из «хлебных крошек»)                                  |
| Цена (£)         | 51.77                                                         |
| На складе (шт.)  | 22 (число из текста «In stock (22 available)»)                |
| Рейтинг          | 3 (число от 1 до 5)                                           |
| UPC              | a897fe39b1053632 (артикул)                                    |
| Описание         | текст описания; если его нет — пусто                          |
| Обложка          | https://books.toscrape.com/media/cache/fe/72/fe72f0532301ec28892ae79a629a293c.jpg |
| Ссылка           | https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html |

3. Сохраняет результат в три файла:

- `books.xlsx` — лист «Книги» со всеми данными (заголовки выделены, ширина
  колонок подобрана, первая строка закреплена, включён фильтр) и лист «Сводка»:
  сколько книг в каждой категории и их средняя цена, плюс итоговая строка;
- `books.csv` — кодировка UTF-8 с BOM и разделитель `;`, поэтому файл
  открывается двойным щелчком в Excel с правильными символами и колонками;
- `books.json` — список книг в формате JSON.

Особенности:

- можно парсить только одну категорию (`--category`) или первые N книг (`--limit`);
- пауза 0.3–0.7 секунды между запросами, чтобы не нагружать сайт;
- если страница не загрузилась — до 3 попыток; если так и не вышло, книга
  пропускается, а в конце выводится список пропущенных ссылок;
- в консоль выводится прогресс: «Книга 37 из 1000»;
- если остановить парсинг (Ctrl+C), всё уже собранное сохраняется в файлы.

## Быстрый старт на Windows

Нужен Python 3.9 или новее ([python.org](https://www.python.org/downloads/)).
Команды выполняются в PowerShell в папке проекта. На Windows используем `py` —
это стандартный запускатель Python, он работает, даже если команда `python`
открывает Microsoft Store.

1. Создать виртуальное окружение (один раз):

   ```bash
   py -m venv .venv
   ```

2. Установить зависимости в него (один раз):

   ```bash
   .venv\Scripts\python -m pip install -r requirements.txt
   ```

3. Запустить парсер:

   ```bash
   .venv\Scripts\python parser.py
   ```

Активировать окружение не обязательно: команды выше сразу вызывают Python из `.venv`.
Готовые файлы `books.xlsx`, `books.csv` и `books.json` появятся в папке проекта.

## Установка (macOS / Linux и общий способ)

Нужен Python 3.9 или новее.

1. Откройте терминал в папке проекта.
2. (Рекомендуется) создайте и активируйте виртуальное окружение:

   ```bash
   python -m venv .venv
   ```

   Windows (PowerShell):

   ```bash
   .venv\Scripts\Activate.ps1
   ```

   macOS / Linux:

   ```bash
   source .venv/bin/activate
   ```

3. Установите зависимости:

   ```bash
   pip install -r requirements.txt
   ```

> На Windows, если команда `python` открывает Microsoft Store, используйте `py` вместо `python`.

## Запуск

Весь каталог (1000 книг):

```bash
python parser.py
```

Только одна категория (регистр букв не важен):

```bash
python parser.py --category Poetry
```

Если в названии категории есть пробел — возьмите его в кавычки:

```bash
python parser.py --category "Science Fiction"
```

Тест на первых 20 книгах:

```bash
python parser.py --limit 20
```

Аргументы можно сочетать: `--category Poetry --limit 5`. Если указать
несуществующую категорию, парсер выведет список всех доступных.
Справка по аргументам: `python parser.py --help`.

Пример вывода:

```
Старт парсинга всего каталога books.toscrape.com

Сбор ссылок, страница 1: https://books.toscrape.com/catalogue/page-1.html
Сбор ссылок, страница 2: https://books.toscrape.com/catalogue/page-2.html
...

Найдено книг: 1000. Заходим на страницу каждой книги...

Книга 1 из 1000: https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html
Книга 2 из 1000: https://books.toscrape.com/catalogue/tipping-the-velvet_999/index.html
...

Собрано книг: 1000 за 823 сек.
Результат сохранён: books.xlsx, books.csv, books.json
```

Полный обход занимает около 14 минут: парсер делает больше тысячи
запросов и делает паузу перед каждым.

## Настройки

В начале `parser.py` есть константы, которые можно менять:

- `DELAY_RANGE` — диапазон паузы между запросами (сек);
- `MAX_RETRIES` — сколько раз повторять загрузку страницы;
- `TIMEOUT` — сколько ждать ответа сервера;
- `XLSX_FILE`, `CSV_FILE`, `JSON_FILE` — имена итоговых файлов.

## Используемые библиотеки

- `requests` — загрузка страниц;
- `beautifulsoup4` — разбор HTML;
- `openpyxl` — создание Excel-файла;
- `csv`, `json`, `argparse` — стандартные модули Python для CSV, JSON и
  аргументов командной строки.
