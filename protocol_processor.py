# -*- coding: utf-8 -*-
"""
Обработка сканированных протоколов измерений базовых станций.

Цепочка: SEC_Document.xlsx (ссылки) -> скачивание RAR из внутреннего
сервиса -> 7z -> OCR (tesseract) -> номер/дата протокола -> запись
в Pool.xlsx (столбец R) -> удаление временных файлов.

Статус: этап 3 (авторизация и скачивание архивов).
"""

import logging
import shutil
import argparse
import getpass
import tempfile
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
import re
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import subprocess


# ===================== КОНСТАНТЫ =====================

DATA_DIR = Path(r"C:\Users\OZhigalin\Documents\PRTO")
SEC_FILE = DATA_DIR / "SEC_Document.xlsx"   # файл со ссылками (только чтение)
POOL_FILE = DATA_DIR / "Pool.xlsx"          # файл, куда пишем результат

SEC_LINK_COL = "AG"        # столбец со ссылкой на архив
SEC_BS_COL = "C"           # столбец с номером базовой станции
SEC_HEADER_ROWS = 2        # строки 1-2 - шапка, данные с 3-й

POOL_BS_COL = "A"          # столбец с номером базовой станции
POOL_RESULT_COL = "R"      # столбец, куда пишем "№NNNN ДД.ММ.ГГГГ"
POOL_HEADER_ROWS = 3       # первые 3 строки - шапка

# Сервис: страница входа (формы). j_security_check - по стандарту Java,
# POST уходит в контекст рядом с login (см. функцию login()).
LOGIN_PAGE_URL = "http://nri-application.vimpelcom.ru/nrijs-web/rest/login"
DOWNLOAD_PAUSE_SEC = 1
BATCH_SIZE = 10

TESSERACT_EXE = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
SEVENZIP_EXE = Path(r"C:\Program Files\7-Zip\7z.exe")
OCR_LANG = "rus"
OCR_DPI = 300
pytesseract.pytesseract.tesseract_cmd = str(TESSERACT_EXE)

# Родительные падежи месяцев -> номер месяца (для дат вида "03" июня 2026 г.)
MONTHS = {
    "января": "01", "февраля": "02", "марта": "03", "апреля": "04",
    "мая": "05", "июня": "06", "июля": "07", "августа": "08",
    "сентября": "09", "октября": "10", "ноября": "11", "декабря": "12",
}

PROTOCOL_KEYWORD = "протокол"

RAR_MAGIC = b"Rar!"       # сигнатура rar-архива
HTTP_TIMEOUT = 60         # сек на запрос

# ===================== ЛОГГЕР =====================

def setup_logging() -> logging.Logger:
    log_file = Path(__file__).parent / f"protocol_log_{datetime.now():%Y%m%d}.txt"
    logger = logging.getLogger("protocol_processor")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s",
                            datefmt="%H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    logger.info("Лог-файл: %s", log_file)
    return logger

# ===================== БЭКАП =====================

def backup_pool(log: logging.Logger) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    dst = POOL_FILE.parent / f"Pool_backup_{stamp}.xlsx"
    shutil.copy2(POOL_FILE, dst)
    log.info("Бэкап создан: %s", dst)
    return dst

# ===================== ЧТЕНИЕ EXCEL =====================

def col_index(letter: str) -> int:
    """'AG' -> 33 (номер столбца openpyxl, с 1)."""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def norm_bs(value) -> str:
    """Номер BS к строке без лишних пробелов."""
    if value is None:
        return ""
    return str(value).strip()


def read_sec_rows(log: logging.Logger) -> list:
    """Читает SEC_Document.xlsx с строки 3 до первой пустой.

    Ссылка берётся из текста ячейки AG, а если текст пуст -
    из гиперссылки ячейки (в Excel ссылка бывает именно объектом).
    """
    # Без read_only: иначе openpyxl не загружает hyperlinks.
    wb = openpyxl.load_workbook(SEC_FILE, data_only=True)
    ws = wb.active
    tasks = []
    link_idx = col_index(SEC_LINK_COL)
    bs_idx = col_index(SEC_BS_COL)
    row_num = 0
    for row in ws.iter_rows(values_only=True):
        row_num += 1
        if row_num <= SEC_HEADER_ROWS:
            continue
        bs = norm_bs(ws.cell(row=row_num, column=bs_idx).value)
        if ws.cell(row=row_num, column=link_idx).value is None and bs == "":
            break
        link = ws.cell(row=row_num, column=link_idx).value
        if link is None or str(link).strip() == "":
            cell = ws.cell(row=row_num, column=link_idx)
            if cell.hyperlink is not None and cell.hyperlink.target:
                link = cell.hyperlink.target
        link = str(link).strip() if link else ""
        if link == "" or bs == "":
            log.warning("Строка %s в SEC: пропущена (ссылка=%r, BS=%r)",
                        row_num, link, bs)
            continue
        tasks.append({"sec_row": row_num, "bs": bs, "url": link})
    wb.close()
    log.info("SEC_Document.xlsx: прочитано строк с данными: %d", len(tasks))
    return tasks


def read_pool_state(log: logging.Logger) -> dict:
    """Читает Pool.xlsx: BS -> номер строки и заполнен ли R."""
    wb = openpyxl.load_workbook(POOL_FILE, read_only=True, data_only=True)
    ws = wb.active
    state = {}
    bs_idx = col_index(POOL_BS_COL) - 1
    r_idx = col_index(POOL_RESULT_COL) - 1
    row_num = 0
    for row in ws.iter_rows(values_only=True):
        row_num += 1
        if row_num <= POOL_HEADER_ROWS:
            continue
        row = tuple(row) + (None,) * (max(bs_idx, r_idx) + 1 - len(row))
        bs = norm_bs(row[bs_idx])
        if bs == "":
            break
        r_value = row[r_idx]
        state[bs] = {
            "pool_row": row_num,
            "r_filled": r_value is not None and str(r_value).strip() != "",
        }
    wb.close()
    log.info("Pool.xlsx: строк с номерами BS: %d", len(state))
    return state

# ===================== АВТОРИЗАЦИЯ И СКАЧИВАНИЕ (этап 3) =====================

def login(log: logging.Logger) -> requests.Session:
    """Открывает страницу входа, POST j_security_check, возвращает сессию."""
    session = requests.Session()
    session.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                     "protocol_processor/1.0")

    log.info("Открываю страницу входа: %s", LOGIN_PAGE_URL)
    resp = session.get(LOGIN_PAGE_URL, timeout=HTTP_TIMEOUT)
    log.info("GET login: HTTP %s, длина ответа %s", resp.status_code, len(resp.text))
    if resp.status_code != 200:
        log.warning("Страница входа вернула HTTP %s - продолжаю попытку логина",
                    resp.status_code)

    username = input("Логин: ").strip()
    password = getpass.getpass("Пароль: ")

    # j_security_check лежит рядом с защищаемым контекстом.
    base = LOGIN_PAGE_URL.rsplit("/", 1)[0]          # .../nrijs-web/rest
    check_url = base.rsplit("/", 1)[0] + "/j_security_check"  # .../nrijs-web/j_security_check
    payload = {"j_username": username, "j_password": password}
    log.info("POST %s", check_url)
    resp = session.post(check_url, data=payload, timeout=HTTP_TIMEOUT,
                        allow_redirects=True)
    log.info("POST j_security_check: HTTP %s, итоговый URL: %s",
             resp.status_code, resp.url)

    # Диагностика: если вернулась страница с формой входа - пароль не принят.
    body = resp.text
    if "j_security_check" in body or "j_password" in body:
        log.error("Похоже, авторизация НЕ прошла: в ответе снова форма входа.")
        log.error("Первые 300 символов ответа: %.300s", body)
        raise SystemExit(1)
    log.info("Авторизация: похоже, прошла успешно (cookies: %d)",
             len(session.cookies))
    return session


def looks_like_rar(content: bytes) -> bool:
    return content[:4] == RAR_MAGIC


def download_archives(session: requests.Session, tasks: list,
                      log: logging.Logger) -> list:
    """Скачивает архивы по ссылкам. Возвращает задачи с полем 'archive'."""
    work_dir = Path(tempfile.mkdtemp(prefix="prto_"))
    log.info("Временная папка: %s", work_dir)

    ok_tasks = []
    for n, t in enumerate(tasks, 1):
        log.info("[%d/%d] Скачиваю BS=%s (SEC стр.%s): %.100s",
                 n, len(tasks), t["bs"], t["sec_row"], t["url"])
        try:
            resp = session.get(t["url"], timeout=HTTP_TIMEOUT)
        except requests.RequestException as e:
            log.error("[%d] Ошибка сети: %s", n, e)
            continue
        if resp.status_code != 200:
            log.error("[%d] HTTP %s - пропуск", n, resp.status_code)
            continue
        if not looks_like_rar(resp.content):
            log.error("[%d] Ответ НЕ rar-архив (начало: %.20r). "
                      "Первые 200 символов как текст: %.200s",
                      n, resp.content[:20],
                      resp.content.decode("utf-8", errors="replace"))
            continue
        archive_path = work_dir / f"bs_{t['bs']}_{t['sec_row']}.rar"
        archive_path.write_bytes(resp.content)
        log.info("[%d] Скачано: %s (%d байт)", n, archive_path.name,
                 len(resp.content))
        t["archive"] = archive_path
        ok_tasks.append(t)
        if n % BATCH_SIZE == 0:
            log.info("--- Пауза после пакета из %d ---", BATCH_SIZE)
        time.sleep(DOWNLOAD_PAUSE_SEC)

    log.info("Скачано архивов: %d из %d", len(ok_tasks), len(tasks))
    return ok_tasks


# ===================== РАСПАКОВКА И ВЫБОР PDF (этап 4) =====================

def extract_pdf(archive: Path, log: logging.Logger = None) -> Path:
    """Распаковывает RAR через 7z и возвращает PDF с 'протокол' в имени."""
    if log is None:
        log = logging.getLogger("protocol_processor")

    out_dir = archive.parent / archive.stem          # ...\bs_14004_3\

    # Удаляем прошлую распаковку, если была (повторные запуски).
    if out_dir.exists():
        shutil.rmtree(out_dir, ignore_errors=True)

    cmd = [str(SEVENZIP_EXE), "x", str(archive), f"-o{out_dir}", "-y"]
    log.info("Распаковка: %s", archive.name)
    proc = subprocess.run(cmd, capture_output=True, timeout=120)
    if proc.returncode != 0:
        tail = proc.stderr.decode("utf-8", errors="replace")[-300:]
        raise RuntimeError(f"7z вернул код {proc.returncode}: {tail}")

    # Рекурсивный поиск PDF (в архиве - папка с двумя PDF).
    pdfs = list(out_dir.rglob("*.pdf")) + list(out_dir.rglob("*.PDF"))
    if not pdfs:
        raise RuntimeError("В архиве нет PDF-файлов")

    # Ищем "протокол" без учёта регистра, только кириллица.
    kw = PROTOCOL_KEYWORD.lower()
    matches = [p for p in pdfs if kw in p.name.lower()]

    if len(matches) == 0:
        log.error("PDF с '%s' в имени не найден. Все файлы в архиве:",
                  PROTOCOL_KEYWORD)
        for p in pdfs:
            log.error("  - %s", p.name)
        raise RuntimeError(f"PDF с '{PROTOCOL_KEYWORD}' в имени не найден")
    if len(matches) > 1:
        log.warning("Найдено несколько PDF с '%s', беру первый: %s",
                    PROTOCOL_KEYWORD, matches[0].name)
    log.info("Выбран PDF: %s", matches[0].name)
    return matches[0]

# ===================== OCR И ИЗВЛЕЧЕНИЕ (этап 5) =====================

DATE_RE = re.compile(r"[«\"]?\s*(\d{1,2})\s*[»\"]?\s+([а-яё]+)\s+(\d{4})\s*г", re.IGNORECASE)
NUM_RE = re.compile(r"№\s*(\d{1,4})")


def _ocr_pdf(pdf_path: Path, log: logging.Logger, debug: bool) -> list:
    """Рендерит страницы PDF и распознаёт текст. Возвращает список по страницам."""
    page_texts = []
    doc = fitz.open(pdf_path)
    log.info("OCR: %s (%d стр.)", pdf_path.name, len(doc))
    zoom = OCR_DPI / 72
    mat = fitz.Matrix(zoom, zoom)
    for page_no, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        text = pytesseract.image_to_string(img, lang=OCR_LANG)
        page_texts.append(text)
        if debug:
            log.info("--- Фрагмент стр. %d (первые 500 симв.) ---", page_no)
            log.info("%.500s", " ".join(text.split()))
    doc.close()
    return page_texts


def _find_protocol_number(pages: list, log: logging.Logger,
                          pdf_path: Path = None) -> str:
    """Ищет номер протокола на стр.1: все кандидаты '№ + цифры',
    сверка с номером из имени PDF-файла."""
    flat1 = " ".join(pages[0].split())
    num_re = re.compile(r"[№Nn][oO°o0]?\s*(\d{1,4})(?!\d)")
    candidates = []
    for m in num_re.finditer(flat1):
        num = m.group(1)
        ctx = flat1[max(0, m.start() - 40):m.end() + 30]
        candidates.append((num, ctx))
        log.info("Кандидат номера: №%s | контекст: %s", num, ctx)

    # Эталон для сверки - номер из имени файла ('Протокол №519_БС 14004.pdf')
    fname_num = None
    if pdf_path is not None:
        mf = re.search(r"№\s*(\d{1,4})", pdf_path.name)
        if mf:
            fname_num = mf.group(1)
            log.info("Номер в имени файла: №%s", fname_num)

    if fname_num:
        for num, ctx in candidates:
            if num == fname_num:
                log.info("Номер подтверждён именем файла: №%s", num)
                return num
        log.warning("Номер из имени файла (%s) не найден среди кандидатов стр.1",
                    fname_num)

    if len(candidates) == 1:
        log.warning("Единственный кандидат - беру его (сверь на пилоте): №%s",
                    candidates[0][0])
        return candidates[0][0]

    # Несколько кандидатов: предпочитаем тот, где рядом 'протокол'/'акта'
    for num, ctx in candidates:
        if re.search(r"ротокол|акта", ctx, re.IGNORECASE):
            log.info("Выбран кандидат рядом с 'протокол/акта': №%s", num)
            return num

    if candidates:
        log.error("Несколько кандидатов, выбрать не удалось: %s",
                  [n for n, _ in candidates])
        raise RuntimeError("Номер протокола: несколько кандидатов")

    if fname_num:
        log.warning("Кандидатов на стр.1 нет - беру номер из имени файла: №%s "
                    "(ОБЯЗАТЕЛЬНО сверь с PDF на пилоте)", fname_num)
        return fname_num

    for label, rx in (("Номер акта измерений", r".{0,60}омер\s*акта.{0,60}"),
                      ("протокол", r".{0,60}ротокол.{0,60}")):
        m = re.search(rx, flat1, re.IGNORECASE)
        log.error("Контекст '%s': %s", label,
                  m.group(0) if m else "не найдено на стр.1")
    raise RuntimeError("Номер протокола не распознан на стр.1")


def ocr_extract(pdf_path: Path, log: logging.Logger = None,
                debug: bool = False) -> tuple:
    """OCR PDF. Возвращает (номер_протокола, дата ДД.ММ.ГГГГ)."""
    if log is None:
        log = logging.getLogger("protocol_processor")

    pages = _ocr_pdf(pdf_path, log, debug)

    # Дата: сначала стр.1, при неудаче - весь документ
    date = None
    for scope in (pages[:1], pages):
        flat = " ".join(" ".join(p.split()) for p in scope)
        m = DATE_RE.search(flat)
        if m:
            month_num = MONTHS.get(m.group(2).lower())
            if month_num:
                date = f"{int(m.group(1)):02d}.{month_num}.{m.group(3)}"
                break
    if date is None:
        flat = " ".join(" ".join(p.split()) for p in pages)
        raise RuntimeError(f"Дата не распознана. Фрагмент: %.300s" % flat)

    number = _find_protocol_number(pages, log, pdf_path=pdf_path)
    log.info("Распознано: №%s от %s", number, date)
    return number, date

# ===================== ЗАПИСЬ И ОЧИСТКА (этап 6) =====================

def write_results(results: list, log: logging.Logger) -> int:
    """Пакетно пишет результаты в Pool.xlsx (столбец R).

    results: список кортежей (pool_row, текст_для_R).
    Возвращает число записанных строк.
    """
    if not results:
        return 0
    wb = openpyxl.load_workbook(POOL_FILE)
    ws = wb.active
    r_col = col_index(POOL_RESULT_COL)
    written = 0
    for pool_row, text in results:
        ws.cell(row=pool_row, column=r_col).value = text
        log.info("Pool стр.%s (R): %s", pool_row, text)
        written += 1
    wb.save(POOL_FILE)
    wb.close()
    log.info("Записано в Pool.xlsx: %d строк", written)
    return written


def cleanup_temp(log: logging.Logger) -> None:
    """Удаляет временные папки prto_* из %TEMP%."""
    temp = Path(tempfile.gettempdir())
    removed = 0
    for d in temp.glob("prto_*"):
        shutil.rmtree(d, ignore_errors=True)
        removed += 1
    log.info("Удалено временных папок: %d", removed)

# ===================== MAIN =====================

def main() -> None:
    parser = argparse.ArgumentParser(description="Обработка протоколов измерений")
    parser.add_argument("--pilot", type=int, default=0,
                        help="обработать только первые N необработанных строк")
    parser.add_argument("--debug", action="store_true",
                        help="выводить распознанный текст в консоль")
    args = parser.parse_args()

    log = setup_logging()
    log.info("=== Запуск (pilot=%s, debug=%s) ===", args.pilot, args.debug)

    for path, name in ((TESSERACT_EXE, "tesseract.exe"),
                       (SEVENZIP_EXE, "7z.exe"),
                       (SEC_FILE, "SEC_Document.xlsx"),
                       (POOL_FILE, "Pool.xlsx")):
        if not path.exists():
            log.error("НЕ НАЙДЕН: %s (%s)", name, path)

    tasks = read_sec_rows(log)
    pool_state = read_pool_state(log)

    pending = []
    for t in tasks:
        info = pool_state.get(t["bs"])
        if info is None:
            log.warning("SEC стр.%s BS=%s: не найдена в Pool.xlsx - пропуск",
                        t["sec_row"], t["bs"])
            continue
        if info["r_filled"]:
            continue
        pending.append(t)

    if args.pilot and args.pilot > 0:
        pending = pending[:args.pilot]

    log.info("К обработке: %d строк", len(pending))
    if not pending:
        log.info("Обрабатывать нечего - выход.")
        return

    session = login(log)

    total = len(pending)
    written_total = 0
    failed = 0

    for start in range(0, total, BATCH_SIZE):
        batch = pending[start:start + BATCH_SIZE]
        log.info("=== ПАКЕТ %d–%d из %d ===",
                 start + 1, start + len(batch), total)

        # 1. Скачивание архива пакета (с паузами внутри)
        downloaded = download_archives(session, batch, log)
        failed += len(batch) - len(downloaded)

        # 2. Распаковка и выбор PDF
        pdf_tasks = []
        for t in downloaded:
            try:
                t["pdf"] = extract_pdf(t["archive"], log)
                pdf_tasks.append(t)
            except Exception as e:
                log.error("BS=%s (SEC стр.%s): архив не обработан - %s",
                          t["bs"], t["sec_row"], e)
                failed += 1

        # 3. OCR и извлечение
        results = []
        for t in pdf_tasks:
            try:
                number, date = ocr_extract(t["pdf"], log, args.debug)
                row = pool_state[t["bs"]]["pool_row"]
                results.append((row, f"№{number} {date}"))
            except Exception as e:
                log.error("BS=%s (SEC стр.%s): OCR не удался - %s",
                          t["bs"], t["sec_row"], e)
                failed += 1

        # 4. Запись результатов пакета (<= 10 строк, одно открытие файла)
        if results:
            try:
                written_total += write_results(results, log)
            except PermissionError:
                log.error("Pool.xlsx ЗАКРЫТ В EXCEL? Запись пакета не выполнена. "
                          "Закрой файл и запусти заново - обработка продолжится "
                          "с незаполненных строк.")
                raise SystemExit(1)

        # 5. Удаление временных файлов пакета
        cleanup_temp(log)

        log.info("=== Пакет завершён: записано %d, всего записано %d, "
                 "ошибок/пропусков %d ===", len(results), written_total, failed)

    log.info("=== ИТОГИ: обработано %d, ошибок/пропусков %d, всего задач %d ===",
             written_total, failed, total)
    log.info("=== Готово. ===")

    # Проверка сессии: пробуем первую ссылку, если не rar - сообщаем и выходим.
    downloaded = download_archives(session, pending, log)

    downloaded = download_archives(session, pending, log)

    # Этап 4: распаковка и выбор PDF с "протокол"
    pdf_tasks = []
    for t in downloaded:
        try:
            pdf = extract_pdf(t["archive"], log)
            t["pdf"] = pdf
            pdf_tasks.append(t)
        except Exception as e:
            log.error("BS=%s (SEC стр.%s): архив не обработан - %s",
                      t["bs"], t["sec_row"], e)

    log.info("Архивов распаковано, PDF выбран: %d из %d",
             len(pdf_tasks), len(downloaded))

    log.info("Архивов распаковано, PDF выбран: %d из %d",
             len(pdf_tasks), len(downloaded))

    # Этап 5: OCR и извлечение номера/даты
    for t in pdf_tasks:
        try:
            number, date = ocr_extract(t["pdf"], log, args.debug)
            t["number"], t["date"] = number, date
        except Exception as e:
            log.error("BS=%s (SEC стр.%s): OCR не удался - %s",
                      t["bs"], t["sec_row"], e)

    ok = [t for t in pdf_tasks if "number" in t]
    log.info("OCR успешен: %d из %d", len(ok), len(pdf_tasks))

    # Этап 6 (вариант Б): запись пакетами по 10 с сохранением после каждого -
    # при сбое обработанное уже в файле, повторный запуск продолжит дальше.
    buffer = []
    written_total = 0
    for t in ok:
        buffer.append((pool_state[t["bs"]]["pool_row"],
                       f"№{t['number']} {t['date']}"))
        if len(buffer) >= BATCH_SIZE:
            try:
                written_total += write_results(buffer, log)
            except PermissionError:
                log.error("Pool.xlsx ЗАКРЫТ В EXCEL? Блок из %d строк не записан.",
                          len(buffer))
                raise SystemExit(1)
            buffer = []
    if buffer:  # хвост пакета
        try:
            written_total += write_results(buffer, log)
        except PermissionError:
            log.error("Pool.xlsx ЗАКРЫТ В EXCEL? Блок из %d строк не записан.",
                      len(buffer))
            raise SystemExit(1)

    cleanup_temp(log)

    # Итоговая сводка
    errors = len(pending) - written_total
    log.info("=== ИТОГИ: обработано %d, ошибок/пропусков %d, всего задач %d ===",
             written_total, errors, len(pending))

    log.info("=== Этап 6 завершён. ===")

if __name__ == "__main__":
    main()