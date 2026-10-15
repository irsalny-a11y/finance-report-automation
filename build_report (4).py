# -*- coding: utf-8 -*-
"""
Пересборка месячного отчёта автосервиса.

Запуск:
    python build_report.py "Отчет_август.xlsx" "Отчет_август_исправленный.xlsx"

Что делает:
  * читает исходный файл (лист «Итого» + листы заказ-нарядов);
  * на каждый проект создаёт лист с корректным расчётом:
    выручка и себестоимость отдельно с НДС и без НДС, валовая прибыль;
    налог на прибыль по проектам НЕ считается — это налог компании в целом;
  * «Сводка по проектам» — все проекты в одной таблице;
  * «Общие расходы» — накладные расходы и прочие доходы/расходы (а/м);
  * «ОПиУ» — отчёт о прибылях и убытках (всё без НДС);
  * «Налоги» — НДС и налог на прибыль за месяц, считаются формулами;
  * «Параметры» — ставки налогов (меняются в одном месте).
Все итоги — формулы Excel, при изменении исходных цифр пересчитываются.
"""
import os
import re
import sys
import traceback

# Папка, где лежит сам скрипт: там ищем исходный отчёт и туда сохраняем результат
BASE = os.path.dirname(os.path.abspath(sys.argv[0]))


class _Tee:
    """Печатает в окно и одновременно пишет в журнал рядом со скриптом."""
    def __init__(self, stream, log):
        self.stream, self.log = stream, log

    def write(self, text):
        for out in (self.stream, self.log):
            try:
                if out:
                    out.write(text)
                    out.flush()
            except Exception:
                pass

    def flush(self):
        pass


LOG_PATH = os.path.join(BASE, "журнал_запуска.txt")
try:
    _log = open(LOG_PATH, "w", encoding="utf-8")
except Exception:
    _log = None
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")   # не падать на символах, которых нет в кодировке консоли
    except Exception:
        pass
sys.stdout = _Tee(sys.stdout, _log)
sys.stderr = _Tee(sys.stderr, _log)


def _pause():
    # чтобы окно не закрывалось сразу при запуске двойным щелчком
    try:
        input("\nНажмите Enter, чтобы закрыть окно...")
    except Exception:
        pass


def _on_error(exc_type, exc, tb):
    print("\n!!! ОШИБКА — файл отчёта НЕ создан.\n")
    if exc_type is PermissionError:
        print("Нет доступа к файлу. Скорее всего, итоговый файл открыт в Excel — "
              "закройте его и запустите скрипт снова.")
    elif exc_type is FileNotFoundError:
        print(f"Файл не найден: {exc}")
    else:
        traceback.print_exception(exc_type, exc, tb)
    print(f"\nПодробности записаны в файл: {LOG_PATH}")
    _pause()
    sys.exit(1)


sys.excepthook = _on_error

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
except ImportError:
    print("Не установлена библиотека openpyxl.\n"
          "Откройте командную строку и выполните:  pip install openpyxl")
    _pause()
    sys.exit(1)

print(f"Скрипт запущен. Папка: {BASE}")
print("Файлы Excel в папке:", ", ".join(f for f in os.listdir(BASE) if f.lower().endswith(".xlsx")) or "нет")
OUT_PREFIX = "Отчет по автосервису за"   # начало имени итогового файла
SRC_NAME = "отчет"   # ищем файлы, в имени которых есть это слово (регистр и «ё» не важны)

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
# корни названий месяцев: находят «август», «августа», «Август», «авг» и т.п.
MONTH_PATTERNS = [r"янв", r"фев", r"мар(?!к)", r"апр", r"ма[йя]", r"июн", r"июл",
                  r"авг", r"сен", r"окт", r"ноя", r"дек"]


def _norm(s):
    import unicodedata
    return unicodedata.normalize("NFC", s).lower().replace("ё", "е")


def find_source():
    if len(sys.argv) > 1:
        path = sys.argv[1]
        return path if os.path.isabs(path) else os.path.join(BASE, path)
    files = sorted(f for f in os.listdir(BASE)
                   if _norm(f).endswith(".xlsx") and not f.startswith("~$")
                   and SRC_NAME in _norm(f)
                   and not _norm(f).startswith(_norm(OUT_PREFIX))   # не берём уже готовые отчёты
                   and "исправлен" not in _norm(f))
    if not files:
        raise FileNotFoundError(
            f"в папке {BASE} нет файла Excel (.xlsx) со словом «Отчет» в названии — "
            f"переименуйте отчёт или положите его рядом со скриптом")
    if len(files) == 1:
        return os.path.join(BASE, files[0])
    print("Найдено несколько файлов со словом «Отчет»:")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    n = int(input("Введите номер файла с отчётом: "))
    return os.path.join(BASE, files[n - 1])


def detect_period(src_path):
    """Месяц и год отчёта: из имени исходного файла, иначе спрашиваем."""
    import datetime
    name = _norm(os.path.basename(src_path))
    month = None
    for i, pat in enumerate(MONTH_PATTERNS):
        if re.search(r"(?<![а-я])" + pat, name):
            month = i + 1
            break
    if month is None:
        print("Не удалось определить месяц по имени файла.")
        while month is None:
            ans = _norm(input("Введите месяц отчёта (например: август или 8): ").strip())
            if ans.isdigit() and 1 <= int(ans) <= 12:
                month = int(ans)
            else:
                for i, pat in enumerate(MONTH_PATTERNS):
                    if re.match(pat, ans):
                        month = i + 1
                        break
            if month is None:
                print("Не понял месяц, попробуйте ещё раз.")
    y = re.search(r"(20\d{2})", name)
    if y:
        year = int(y.group(1))
    else:
        today = datetime.date.today()
        # отчёт составляют за прошедший месяц: декабрьский отчёт в январе — прошлого года
        year = today.year if month <= today.month else today.year - 1
    return MONTHS[month - 1], year


SRC = find_source()
MONTH, YEAR = detect_period(SRC)
PERIOD = f"{MONTH} {YEAR}"
if len(sys.argv) > 2:
    DST = sys.argv[2] if os.path.isabs(sys.argv[2]) else os.path.join(BASE, sys.argv[2])
else:
    DST = os.path.join(BASE, f"{OUT_PREFIX} {PERIOD}.xlsx")
print(f"Исходный файл: {SRC}")
print(f"Период отчёта: {PERIOD}")

# ---------------------------------------------------------------- стили
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")          # исходные данные
F_LINK = Font(name=FONT, size=10, color="008000")           # ссылки на другие листы
F_BOLD = Font(name=FONT, size=10, bold=True)
F_BOLD_LINK = Font(name=FONT, size=10, bold=True, color="008000")
F_TITLE = Font(name=FONT, size=13, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_NOTE = Font(name=FONT, size=9, italic=True, color="666666")
FILL_HDR = PatternFill("solid", fgColor="1F4E78")
FILL_SEC = PatternFill("solid", fgColor="DDEBF7")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_ASSUMP = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TOP2 = Border(top=Side(style="medium", color="000000"))
MONEY = '#,##0.00;(#,##0.00);"-"'
QTY = '#,##0.##;(#,##0.##);"-"'
PCT = '0.0%;(0.0%);"-"'
FLAG = '0'

VAT = "'Параметры'!$B$4"
PTAX = "'Параметры'!$B$5"
PTAX_FED = "'Параметры'!$B$6"
PTAX_REG = "'Параметры'!$B$7"
INS = "'Параметры'!$B$8"


def q(name):
    return "'" + name.replace("'", "''") + "'"


def put(ws, ref, value, font=F_BASE, fmt=None, fill=None, border=None, align=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = align
    return c


def header(ws, row, titles, col0=1):
    for i, t in enumerate(titles):
        put(ws, f"{L(col0 + i)}{row}", t, F_HDR, fill=FILL_HDR, border=BOX,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = 42


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[L(i)].width = x


def num(v):
    return v if isinstance(v, (int, float)) else 0


# ---------------------------------------------------------------- чтение исходника
src = load_workbook(SRC)
src_v = load_workbook(SRC, data_only=True)   # значения, посчитанные Excel в исходном файле
tot = src["Итого"]
tot_v = src_v["Итого"]
# если файл сохранён не в Excel, в нём может не быть посчитанных значений формул
if any(isinstance(c.value, str) and c.value.startswith("=") for row in tot.iter_rows() for c in row) and \
        all(c.value is None for row in tot_v.iter_rows() for c in row
            if isinstance(tot.cell(c.row, c.column).value, str)
            and str(tot.cell(c.row, c.column).value).startswith("=")):
    print("! ВНИМАНИЕ: в исходном файле нет сохранённых результатов формул.\n"
          "  Откройте его в Excel, нажмите «Сохранить» и запустите скрипт снова —\n"
          "  иначе лист «Сверка» покажет неверные данные исходника.")


def sheet_of_row(r):
    for c in (2, 3, 4, 6):
        v = tot.cell(r, c).value
        if isinstance(v, str) and v.startswith("="):
            m = re.match(r"^=(?:'((?:[^']|'')+)'|([^!]+))!", v)
            if m:
                return (m.group(1) or m.group(2)).replace("''", "'")
    name = str(tot.cell(r, 1).value)
    for s in src.sheetnames:
        if s.strip().lower() == name.strip().lower():
            return s
    return None


def _cached(wv, r, c, fallback):
    v = wv.cell(r, c).value
    return v if isinstance(v, (int, float)) else fallback


def parse_project(ws):
    wv = src_v[ws.title]
    chk = dict(works_gross=0.0, works_n=0, parts_sale=0.0, parts_buy=0.0, parts_vat_in=0.0, parts_n=0)
    works, parts = [], []
    w_hdr = p_hdr = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            if v.strip().startswith("Работы") and w_hdr is None:
                w_hdr = row[0].row
            if v.strip().startswith("Наимен") and p_hdr is None:
                p_hdr = row[0].row
    if w_hdr:
        r = w_hdr + 1
        while ws.cell(r, 1).value not in (None, ""):
            works.append(dict(name=str(ws.cell(r, 1).value).strip(),
                              qty=num(ws.cell(r, 2).value),
                              price=num(ws.cell(r, 3).value)))
            # колонка E «стоимость работ с НДС» исходника
            chk["works_gross"] += _cached(wv, r, 5, works[-1]["qty"] * works[-1]["price"])
            chk["works_n"] += 1
            r += 1
    if p_hdr:
        r = p_hdr + 1
        while ws.cell(r, 1).value not in (None, ""):
            k = ws.cell(r, 11).value
            m = ws.cell(r, 13).value
            parts.append(dict(
                name=str(ws.cell(r, 1).value).strip().replace("\\\\", "\\"),
                art=(str(ws.cell(r, 2).value).strip() if ws.cell(r, 2).value else None),
                qty=num(ws.cell(r, 3).value),
                price=num(ws.cell(r, 4).value),
                markup=num(ws.cell(r, 7).value),
                fixed_sale=k if isinstance(k, (int, float)) else None,
                vat_in=0 if (isinstance(m, (int, float)) and m == 0) else 1,
            ))
            it = parts[-1]
            buy = it["qty"] * it["price"]
            sale = it["fixed_sale"] if it["fixed_sale"] is not None else buy * (1 + it["markup"])
            chk["parts_buy"] += _cached(wv, r, 6, buy)                     # F «Закупка»
            chk["parts_sale"] += _cached(wv, r, 11, sale)                  # K «Цена продажи с НДС»
            chk["parts_vat_in"] += _cached(wv, r, 13, buy * it["vat_in"] * 0.22 / 1.22)  # M «входящий НДС»
            chk["parts_n"] += 1
            r += 1
    return works, parts, chk


projects = []
r = 2
while True:
    a = tot.cell(r, 1).value
    if isinstance(a, str) and a.strip().upper() == "ИТОГО":
        break
    if r > 200:
        break
    if a:
        sname = sheet_of_row(r)
        if sname is None:
            print(f"! Не найден лист для строки {r}: {a}")
        else:
            w, p, chk = parse_project(src[sname])
            in_itogo = tot.cell(r, 2).value not in (None, "") or tot.cell(r, 3).value not in (None, "")
            projects.append(dict(title=str(a).strip(), src_sheet=sname,
                                 works=w, parts=p, chk=chk, in_itogo=in_itogo,
                                 cash_work=num(tot.cell(r, 7).value),
                                 cash_parts=num(tot.cell(r, 8).value)))
    r += 1

# листы заказ-нарядов, которые есть в файле, но не записаны в «Итого»
_taken = {p_["src_sheet"] for p_ in projects}
for sname in src.sheetnames:
    if sname == "Итого" or sname in _taken:
        continue
    w, p, chk = parse_project(src[sname])
    if w or p:
        print(f"! Лист «{sname}» не указан в «Итого» исходника — добавлен в отчёт автоматически")
        projects.append(dict(title=sname.strip(), src_sheet=sname, works=w, parts=p, chk=chk,
                             in_itogo=False, cash_work=0, cash_parts=0))
SRC_PROJECT_SHEETS = sum(1 for sn in src.sheetnames if sn != "Итого"
                         and any(parse_project(src[sn])[:2]))

# общие расходы: колонки O (сумма), P (статья), Q (форма оплаты)
CATEGORY = {
    "зп": "Заработная плата", "аренда": "Аренда",
    "оборудование": "Инструмент и оборудование", "инструмент": "Инструмент и оборудование",
    "ст.машина": "Инструмент и оборудование",
    "программа": "Услуги (бухгалтерия, ПО)", "бухгалтер": "Услуги (бухгалтерия, ПО)",
    "комиссия": "Банковские комиссии", "доставка масла": "Транспорт и доставка",
}
overheads = []
for r in range(3, 34):
    s = tot.cell(r, 15).value
    if not isinstance(s, (int, float)):
        continue
    item = tot.cell(r, 16).value
    item = str(item).strip() if item else f"Без описания (стр. {r} исходного листа «Итого»)"
    pay = tot.cell(r, 17).value
    overheads.append(dict(item=item, amount=s, pay=str(pay).strip() if pay else "",
                          cat=CATEGORY.get(item.lower(), "Прочие расходы"),
                          deductible=0 if item.lower() == "кешбек" else 1))

# прочие доходы/расходы (а/м): J/K — расход, L/M — приход
other_inc, other_exp = [], []
for r in range(3, 34):
    if isinstance(tot.cell(r, 13).value, (int, float)):
        other_inc.append(dict(item=str(tot.cell(r, 12).value).strip(), amount=tot.cell(r, 13).value))
    if isinstance(tot.cell(r, 11).value, (int, float)):
        other_exp.append(dict(item=str(tot.cell(r, 10).value).strip(), amount=tot.cell(r, 11).value))

def _sum_col(ws, col, r1=2, r2=34):
    return sum(ws.cell(rr, col).value for rr in range(r1, r2)
               if isinstance(ws.cell(rr, col).value, (int, float)))


def _itogo_total(col, fallback_col):
    """Строка «ИТОГО» исходника, если она посчитана; иначе сумма колонки."""
    for rr in range(2, 60):
        a = tot_v.cell(rr, 1).value
        if isinstance(a, str) and a.strip().upper() == "ИТОГО":
            v = tot_v.cell(rr, col).value
            if isinstance(v, (int, float)):
                return v
    return _sum_col(tot_v, fallback_col)


SRC_TOTALS = dict(
    works=_itogo_total(2, 2), parts=_itogo_total(3, 3),
    opex=_itogo_total(15, 15), cash_work=_sum_col(tot_v, 7, 2, 31),
    inc=_sum_col(tot_v, 13, 3, 34), exp=_sum_col(tot_v, 11, 3, 34),
    net_cash=_itogo_total(14, 14),
    n_opex=len([1 for rr in range(3, 34) if isinstance(tot.cell(rr, 15).value, (int, float))]),
)

orig_result = None
try:
    orig_result = load_workbook(SRC, data_only=True)["Итого"]["F39"].value
except Exception:
    pass

# ---------------------------------------------------------------- новая книга
wb = Workbook()
ws_pl = wb.active
ws_pl.title = "ОПиУ"
ws_tax = wb.create_sheet("Налоги")
ws_chk = wb.create_sheet("Сверка")
ws_sum = wb.create_sheet("Сводка по проектам")
ws_oh = wb.create_sheet("Общие расходы")
ws_par = wb.create_sheet("Параметры")
ws_fix = wb.create_sheet("Исправления")

# ---- Параметры
p = ws_par
put(p, "A1", "Параметры расчёта", F_TITLE)
put(p, "A2", "Синий шрифт — вводимые значения, чёрный — формулы, зелёный — ссылки на другие листы. "
             "Жёлтая заливка — допущения, которые стоит проверить.", F_NOTE)
header(p, 3, ["Параметр", "Значение", "Комментарий"])
rows = [
    ("Ставка НДС", 0.22, PCT, "Общая ставка НДС с 01.01.2026 (ст. 164 НК РФ). Цены в исходном файле — с НДС."),
    ("Ставка налога на прибыль", 0.25, PCT, "С 01.01.2025, ст. 284 НК РФ."),
    ("  в т.ч. федеральный бюджет", 0.08, PCT, "Доля федерального бюджета (2025–2030)."),
    ("  в т.ч. бюджет субъекта", "=B5-B6", PCT, "Остаток ставки."),
    ("Страховые взносы, % от ФОТ", 0.0, PCT,
     "Заполните, если суммы зарплаты в учёте НЕ включают взносы (обычно 30%, для МСП 15% сверх 1,5 МРОТ). "
     "По умолчанию 0 — взносы не начисляются."),
    ("Период", PERIOD, None, ""),
    ("Порог выручки объекта для рейтинга маржинальности, руб.", 0, MONEY,
     "Объекты с выручкой без НДС ниже порога не участвуют в выборе самого низкомаржинального "
     "(чтобы мелкие заказы не заслоняли крупные). 0 — участвуют все."),
]
for i, (lbl, val, fmt, com) in enumerate(rows, start=4):
    put(p, f"A{i}", lbl, border=BOX)
    is_formula = isinstance(val, str) and val.startswith("=")
    put(p, f"B{i}", val, F_BASE if is_formula else F_INPUT, fmt, border=BOX,
        fill=FILL_ASSUMP if lbl.startswith("Страховые") else None)
    put(p, f"C{i}", com, F_NOTE, border=BOX, align=Alignment(wrap_text=True))
widths(p, [34, 14, 90])

# ---- Листы проектов
used_names = set(wb.sheetnames)
proj_refs = []  # (title, sheetname, dict label->cell)
SUM_ROWS = [
    ("works_gross", "Выручка за работы, с НДС"),
    ("works_vat", "НДС с работ"),
    ("works_net", "Выручка за работы, без НДС"),
    ("parts_gross", "Выручка за запчасти, с НДС"),
    ("parts_vat", "НДС с запчастей"),
    ("parts_net", "Выручка за запчасти, без НДС"),
    ("rev_net", "Выручка всего, без НДС"),
    ("parts_cost", "Себестоимость запчастей, без НДС"),
    ("cash_work", "Прямые расходы по работам (нал., без НДС)"),
    ("cost", "Себестоимость всего, без НДС"),
    ("gp", "Валовая прибыль"),
    ("gpm", "Рентабельность по валовой прибыли"),
    ("vat_in", "НДС входящий по запчастям (к вычету)"),
    ("cash_parts", "Справочно: расход нал. на з/ч (из исходного файла)"),
]

for pr in projects:
    name = pr["src_sheet"].strip()[:31]
    base = name
    k = 2
    while name in used_names:
        name = f"{base[:28]}_{k}"
        k += 1
    used_names.add(name)
    ws = wb.create_sheet(name)
    put(ws, "A1", f"Заказ-наряд: {pr['title']}", F_TITLE)
    put(ws, "A2", "Цены — с НДС. Налог на прибыль здесь не считается: он рассчитывается по компании "
                  "в целом (лист «Налоги»), т.к. база — прибыль после всех расходов.", F_NOTE)

    # размещение таблиц
    first = 4
    put(ws, f"A{first}", "ИТОГИ ПО ПРОЕКТУ", F_BOLD, fill=FILL_SEC)
    ws.merge_cells(start_row=first, start_column=1, end_row=first, end_column=3)
    srow = {key: first + 1 + i for i, (key, _) in enumerate(SUM_ROWS)}
    w_hdr = first + len(SUM_ROWS) + 3
    works = pr["works"]
    w_first = w_hdr + 1
    w_last = w_first + max(len(works), 1) - 1
    w_tot = w_last + 1
    p_title = w_tot + 3
    p_hdr = p_title + 1
    parts = pr["parts"]
    p_first = p_hdr + 1
    p_last = p_first + max(len(parts), 1) - 1
    p_tot = p_last + 1

    # --- таблица работ
    put(ws, f"A{w_hdr - 1}", "РАБОТЫ", F_BOLD, fill=FILL_SEC)
    header(ws, w_hdr, ["№", "Наименование работ", "Кол-во, н/ч", "Цена н/ч, с НДС",
                       "Сумма, с НДС", "НДС", "Сумма, без НДС"])
    if works:
        for i, w in enumerate(works):
            rr = w_first + i
            put(ws, f"A{rr}", i + 1, border=BOX)
            put(ws, f"B{rr}", w["name"], border=BOX, align=Alignment(wrap_text=True))
            put(ws, f"C{rr}", w["qty"], F_INPUT, QTY, border=BOX)
            put(ws, f"D{rr}", w["price"], F_INPUT, MONEY, border=BOX)
            put(ws, f"E{rr}", f"=C{rr}*D{rr}", fmt=MONEY, border=BOX)
            put(ws, f"F{rr}", f"=E{rr}*{VAT}/(1+{VAT})", fmt=MONEY, border=BOX)
            put(ws, f"G{rr}", f"=E{rr}-F{rr}", fmt=MONEY, border=BOX)
    else:
        put(ws, f"B{w_first}", "— работ нет —", F_NOTE, border=BOX)
        for c in "ACDEFG":
            ws[f"{c}{w_first}"].border = BOX
    put(ws, f"B{w_tot}", "Итого работы", F_BOLD, fill=FILL_TOT, border=BOX)
    for c in "ACD":
        put(ws, f"{c}{w_tot}", None, fill=FILL_TOT, border=BOX)
    for c in "EFG":
        put(ws, f"{c}{w_tot}", f"=SUM({c}{w_first}:{c}{w_last})", F_BOLD, MONEY, FILL_TOT, BOX)

    # --- таблица запчастей
    put(ws, f"A{p_title}", "ЗАПЧАСТИ И МАТЕРИАЛЫ", F_BOLD, fill=FILL_SEC)
    header(ws, p_hdr, ["№", "Наименование", "Артикул", "Кол-во", "Цена закупки за ед.",
                       "Сумма закупки", "Закупка с НДС? (1/0)", "НДС входящий",
                       "Закупка без НДС", "Наценка", "Цена продажи за ед., с НДС",
                       "Сумма продажи, с НДС", "НДС исходящий", "Продажа без НДС",
                       "Маржа без НДС", "Маржа, % к себест."])
    if parts:
        for i, it in enumerate(parts):
            rr = p_first + i
            put(ws, f"A{rr}", i + 1, border=BOX)
            put(ws, f"B{rr}", it["name"], border=BOX, align=Alignment(wrap_text=True))
            put(ws, f"C{rr}", it["art"], F_INPUT, border=BOX)
            put(ws, f"D{rr}", it["qty"], F_INPUT, QTY, border=BOX)
            put(ws, f"E{rr}", it["price"], F_INPUT, MONEY, border=BOX)
            put(ws, f"F{rr}", f"=D{rr}*E{rr}", fmt=MONEY, border=BOX)
            put(ws, f"G{rr}", it["vat_in"], F_INPUT, FLAG, border=BOX,
                align=Alignment(horizontal="center"))
            put(ws, f"H{rr}", f"=F{rr}*G{rr}*{VAT}/(1+{VAT})", fmt=MONEY, border=BOX)
            put(ws, f"I{rr}", f"=F{rr}-H{rr}", fmt=MONEY, border=BOX)
            if it["fixed_sale"] is not None:
                put(ws, f"L{rr}", it["fixed_sale"], F_INPUT, MONEY, border=BOX)
                put(ws, f"J{rr}", f"=IF(F{rr}=0,0,L{rr}/F{rr}-1)", fmt=PCT, border=BOX)
                ws[f"L{rr}"].comment = Comment("Цена продажи задана вручную в исходном файле.", "script")
            else:
                put(ws, f"J{rr}", it["markup"], F_INPUT, PCT, border=BOX)
                put(ws, f"L{rr}", f"=F{rr}*(1+J{rr})", fmt=MONEY, border=BOX)
            put(ws, f"K{rr}", f"=IF(D{rr}=0,0,L{rr}/D{rr})", fmt=MONEY, border=BOX)
            put(ws, f"M{rr}", f"=L{rr}*{VAT}/(1+{VAT})", fmt=MONEY, border=BOX)
            put(ws, f"N{rr}", f"=L{rr}-M{rr}", fmt=MONEY, border=BOX)
            put(ws, f"O{rr}", f"=N{rr}-I{rr}", fmt=MONEY, border=BOX)
            put(ws, f"P{rr}", f"=IF(I{rr}=0,0,O{rr}/I{rr})", fmt=PCT, border=BOX)
    else:
        put(ws, f"B{p_first}", "— запчастей нет —", F_NOTE, border=BOX)
        for c in "ACDEFGHIJKLMNOP":
            ws[f"{c}{p_first}"].border = BOX
    put(ws, f"B{p_tot}", "Итого запчасти", F_BOLD, fill=FILL_TOT, border=BOX)
    for c in "ACDEGJK":
        put(ws, f"{c}{p_tot}", None, fill=FILL_TOT, border=BOX)
    for c in "FHILMNO":
        put(ws, f"{c}{p_tot}", f"=SUM({c}{p_first}:{c}{p_last})", F_BOLD, MONEY, FILL_TOT, BOX)
    put(ws, f"P{p_tot}", f"=IF(I{p_tot}=0,0,O{p_tot}/I{p_tot})", F_BOLD, PCT, FILL_TOT, BOX)

    # --- блок итогов
    f = {
        "works_gross": f"=E{w_tot}",
        "works_vat": f"=F{w_tot}",
        "works_net": f"=G{w_tot}",
        "parts_gross": f"=L{p_tot}",
        "parts_vat": f"=M{p_tot}",
        "parts_net": f"=N{p_tot}",
        "rev_net": f"=B{srow['works_net']}+B{srow['parts_net']}",
        "parts_cost": f"=I{p_tot}",
        "cash_work": pr["cash_work"],
        "cost": f"=B{srow['parts_cost']}+B{srow['cash_work']}",
        "gp": f"=B{srow['rev_net']}-B{srow['cost']}",
        "gpm": f"=IF(B{srow['rev_net']}=0,0,B{srow['gp']}/B{srow['rev_net']})",
        "vat_in": f"=H{p_tot}",
        "cash_parts": pr["cash_parts"],
    }
    for key, label in SUM_ROWS:
        rr = srow[key]
        bold = key in ("rev_net", "cost", "gp")
        put(ws, f"A{rr}", label, F_BOLD if bold else F_BASE, border=BOX,
            fill=FILL_TOT if bold else None)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
        val = f[key]
        is_input = not (isinstance(val, str) and val.startswith("="))
        font = F_INPUT if is_input else (F_BOLD if bold else F_BASE)
        put(ws, f"C{rr}", val, font, PCT if key == "gpm" else MONEY,
            FILL_TOT if bold else None, BOX)
    ws[f"C{srow['cash_work']}"].comment = Comment(
        "Колонка «расход нал работы» листа «Итого» исходного файла. НДС не выделяется.", "script")
    ws[f"C{srow['cash_parts']}"].comment = Comment(
        "Колонка «расход нал з/ч» листа «Итого». В расчёт НЕ входит: стоимость запчастей уже "
        "учтена в закупке (таблица ниже). Если это дополнительные расходы — перенесите сумму в "
        "строку «Прямые расходы по работам».", "script")
    # в формулах выше использовалась колонка B — переписываем на C
    for key, _ in SUM_ROWS:
        c = ws[f"C{srow[key]}"]
        if isinstance(c.value, str):
            c.value = re.sub(r"\bB(\d+)\b", r"C\1", c.value)

    widths(ws, [5, 52, 16, 11, 14, 15, 12, 13, 14, 10, 15, 16, 14, 15, 14, 12])
    ws.freeze_panes = "A4"
    _refs = {k2: f"C{v}" for k2, v in srow.items()}
    _refs.update(_w=(w_first, w_last), _p=(p_first, p_last), parts_buy=f"F{p_tot}")
    proj_refs.append((pr["title"], name, _refs))

# ---- Сводка по проектам
s = ws_sum
put(s, "A1", "Сводка по проектам (заказ-нарядам)", F_TITLE)
put(s, "A2", "Все значения — ссылки на листы проектов. Название проекта — гиперссылка на его лист.", F_NOTE)
cols = [("works_gross", "Работы, с НДС"), ("works_vat", "НДС с работ"), ("works_net", "Работы, без НДС"),
        ("parts_gross", "Запчасти, с НДС"), ("parts_vat", "НДС с запчастей"),
        ("parts_net", "Запчасти, без НДС"), ("rev_net", "Выручка, без НДС"),
        ("parts_cost", "Себест. з/ч, без НДС"), ("cash_work", "Прямые нал. расходы по работам"),
        ("cost", "Себестоимость всего"), ("gp", "Валовая прибыль"), ("gpm", "Рентаб-ть"),
        ("vat_in", "НДС входящий по з/ч"), ("cash_parts", "Справочно: нал. на з/ч")]
header(s, 4, ["№", "Проект"] + [c[1] for c in cols])
S_FIRST = 5
for i, (title, sh, refs) in enumerate(proj_refs):
    rr = S_FIRST + i
    put(s, f"A{rr}", i + 1, border=BOX)
    c = put(s, f"B{rr}", title, Font(name=FONT, size=10, color="0563C1", underline="single"), border=BOX)
    c.hyperlink = f"#{q(sh)}!A1"
    for j, (key, _) in enumerate(cols):
        put(s, f"{L(3 + j)}{rr}", f"={q(sh)}!{refs[key]}", F_LINK, PCT if key == "gpm" else MONEY, border=BOX)
S_LAST = S_FIRST + len(proj_refs) - 1
S_TOT = S_LAST + 1
SC = {key: L(3 + j) for j, (key, _) in enumerate(cols)}
put(s, f"B{S_TOT}", "ИТОГО", F_BOLD, fill=FILL_TOT, border=BOX)
put(s, f"A{S_TOT}", None, fill=FILL_TOT, border=BOX)
for key, col in SC.items():
    if key == "gpm":
        put(s, f"{col}{S_TOT}", f"=IF({SC['rev_net']}{S_TOT}=0,0,{SC['gp']}{S_TOT}/{SC['rev_net']}{S_TOT})",
            F_BOLD, PCT, FILL_TOT, BOX)
    else:
        put(s, f"{col}{S_TOT}", f"=SUM({col}{S_FIRST}:{col}{S_LAST})", F_BOLD, MONEY, FILL_TOT, BOX)
widths(s, [5, 30] + [14] * len(cols))
s.freeze_panes = "C5"


def SUMREF(key):
    return f"{q('Сводка по проектам')}!${SC[key]}${S_TOT}"


# ---- Общие расходы
o = ws_oh
put(o, "A1", "Общие расходы, прочие доходы и расходы", F_TITLE)
put(o, "A2", "Жёлтые колонки — допущения: ставка входящего НДС (0 — нет счёта-фактуры) и признак "
             "«учитывается в налоге на прибыль» (0 — нет подтверждающих документов). Проверьте их.", F_NOTE)
put(o, "A4", "1. ОБЩИЕ (ОПЕРАЦИОННЫЕ) РАСХОДЫ", F_BOLD, fill=FILL_SEC)
header(o, 5, ["№", "Статья", "Категория ОПиУ", "Форма оплаты", "Сумма (как в учёте)",
              "Ставка вх. НДС", "НДС входящий", "Сумма без НДС", "Учитывается в НУ (1/0)",
              "Расход для налога на прибыль"])
O_FIRST = 6
for i, e in enumerate(overheads):
    rr = O_FIRST + i
    put(o, f"A{rr}", i + 1, border=BOX)
    put(o, f"B{rr}", e["item"], F_INPUT, border=BOX)
    put(o, f"C{rr}", e["cat"], F_INPUT, border=BOX)
    put(o, f"D{rr}", e["pay"], F_INPUT, border=BOX)
    put(o, f"E{rr}", e["amount"], F_INPUT, MONEY, border=BOX)
    put(o, f"F{rr}", 0, F_INPUT, PCT, FILL_ASSUMP, BOX)
    put(o, f"G{rr}", f"=E{rr}*F{rr}/(1+F{rr})", fmt=MONEY, border=BOX)
    put(o, f"H{rr}", f"=E{rr}-G{rr}", fmt=MONEY, border=BOX)
    put(o, f"I{rr}", e["deductible"], F_INPUT, FLAG, FILL_ASSUMP, BOX, Alignment(horizontal="center"))
    put(o, f"J{rr}", f"=H{rr}*I{rr}", fmt=MONEY, border=BOX)
    if e["item"].lower() == "кешбек":
        o[f"I{rr}"].comment = Comment("Выплаты «кешбека» обычно не подтверждены документами и не "
                                      "уменьшают налог на прибыль. Поставьте 1, если документы есть.", "script")
O_LAST = O_FIRST + len(overheads) - 1
O_TOT = O_LAST + 1
put(o, f"B{O_TOT}", "Итого общие расходы", F_BOLD, fill=FILL_TOT, border=BOX)
for c in "ACDFI":
    put(o, f"{c}{O_TOT}", None, fill=FILL_TOT, border=BOX)
for c in "EGHJ":
    put(o, f"{c}{O_TOT}", f"=SUM({c}{O_FIRST}:{c}{O_LAST})", F_BOLD, MONEY, FILL_TOT, BOX)

# прочие доходы
I_SEC = O_TOT + 3
put(o, f"A{I_SEC}", "2. ПРОЧИЕ ДОХОДЫ (а/м, наличные — колонки «Приход а/м» исходного файла)", F_BOLD, fill=FILL_SEC)
header(o, I_SEC + 1, ["№", "Наименование", "", "Форма оплаты", "Сумма", "Ставка НДС",
                      "НДС исходящий", "Сумма без НДС"])
I_FIRST = I_SEC + 2
for i, e in enumerate(other_inc):
    rr = I_FIRST + i
    put(o, f"A{rr}", i + 1, border=BOX)
    put(o, f"B{rr}", e["item"], F_INPUT, border=BOX)
    put(o, f"C{rr}", None, border=BOX)
    put(o, f"D{rr}", "нал", F_INPUT, border=BOX)
    put(o, f"E{rr}", e["amount"], F_INPUT, MONEY, border=BOX)
    put(o, f"F{rr}", f"={VAT}", F_LINK, PCT, FILL_ASSUMP, BOX)
    put(o, f"G{rr}", f"=E{rr}*F{rr}/(1+F{rr})", fmt=MONEY, border=BOX)
    put(o, f"H{rr}", f"=E{rr}-G{rr}", fmt=MONEY, border=BOX)
o[f"F{I_FIRST}"].comment = Comment(
    "Плательщик НДС начисляет НДС на любую реализацию услуг, в т.ч. за наличные. Поэтому по "
    "умолчанию выделяется 22%. Поставьте 0, если эти поступления не являются выручкой компании.", "script")
I_LAST = I_FIRST + len(other_inc) - 1
I_TOT = I_LAST + 1
put(o, f"B{I_TOT}", "Итого прочие доходы", F_BOLD, fill=FILL_TOT, border=BOX)
for c in "ACDF":
    put(o, f"{c}{I_TOT}", None, fill=FILL_TOT, border=BOX)
for c in "EGH":
    put(o, f"{c}{I_TOT}", f"=SUM({c}{I_FIRST}:{c}{I_LAST})", F_BOLD, MONEY, FILL_TOT, BOX)

# расходы по а/м
X_SEC = I_TOT + 3
put(o, f"A{X_SEC}", "3. РАСХОДЫ ПО ПРОЧИМ ДОХОДАМ (колонки «Расход а/м» исходного файла)", F_BOLD, fill=FILL_SEC)
header(o, X_SEC + 1, ["№", "Наименование", "", "Форма оплаты", "Сумма", "Ставка вх. НДС",
                      "НДС входящий", "Сумма без НДС", "Учитывается в НУ (1/0)", "Расход для налога на прибыль"])
X_FIRST = X_SEC + 2
for i, e in enumerate(other_exp):
    rr = X_FIRST + i
    put(o, f"A{rr}", i + 1, border=BOX)
    put(o, f"B{rr}", e["item"], F_INPUT, border=BOX)
    put(o, f"C{rr}", None, border=BOX)
    put(o, f"D{rr}", "нал", F_INPUT, border=BOX)
    put(o, f"E{rr}", e["amount"], F_INPUT, MONEY, border=BOX)
    put(o, f"F{rr}", 0, F_INPUT, PCT, FILL_ASSUMP, BOX)
    put(o, f"G{rr}", f"=E{rr}*F{rr}/(1+F{rr})", fmt=MONEY, border=BOX)
    put(o, f"H{rr}", f"=E{rr}-G{rr}", fmt=MONEY, border=BOX)
    put(o, f"I{rr}", 1, F_INPUT, FLAG, FILL_ASSUMP, BOX, Alignment(horizontal="center"))
    put(o, f"J{rr}", f"=H{rr}*I{rr}", fmt=MONEY, border=BOX)
X_LAST = X_FIRST + len(other_exp) - 1
X_TOT = X_LAST + 1
put(o, f"B{X_TOT}", "Итого расходы по прочим доходам", F_BOLD, fill=FILL_TOT, border=BOX)
for c in "ACDFI":
    put(o, f"{c}{X_TOT}", None, fill=FILL_TOT, border=BOX)
for c in "EGHJ":
    put(o, f"{c}{X_TOT}", f"=SUM({c}{X_FIRST}:{c}{X_LAST})", F_BOLD, MONEY, FILL_TOT, BOX)
widths(o, [5, 42, 28, 12, 15, 11, 14, 15, 13, 17])
o.freeze_panes = "A6"

OH = q("Общие расходы")

# ---- ОПиУ
pl = ws_pl
put(pl, "A1", f'=\"Отчёт о прибылях и убытках за \"&{q("Параметры")}!B9', F_TITLE)
put(pl, "A2", "Все суммы — без НДС, в рублях. НДС не является ни доходом, ни расходом компании.", F_NOTE)
header(pl, 4, ["Показатель", "Сумма, руб.", "% от выручки"])
PL = {}
r = 5


def pl_row(key, label, formula, kind="normal", fmt=MONEY):
    global r
    font = {"normal": F_BASE, "sub": F_BASE, "total": F_BOLD, "grand": F_BOLD}[kind]
    link = isinstance(formula, str) and "!" in formula and kind in ("normal", "sub")
    fill = FILL_TOT if kind in ("total", "grand") else None
    put(pl, f"A{r}", label, font, fill=fill, border=BOX,
        align=Alignment(indent=2 if kind == "sub" else 0))
    put(pl, f"B{r}", formula, F_LINK if link else font, fmt, fill, BOX)
    if fmt == MONEY:
        put(pl, f"C{r}", f"=IF($B${{REV}}=0,0,B{r}/$B${{REV}})", font, PCT, fill, BOX)
    else:
        put(pl, f"C{r}", None, fill=fill, border=BOX)
    if kind == "grand":
        for c in "ABC":
            pl[f"{c}{r}"].border = Border(top=Side(style="medium"), bottom=Side(style="medium"),
                                          left=THIN, right=THIN)
    PL[key] = r
    r += 1


def blank():
    global r
    r += 1


pl_row("rev", "ВЫРУЧКА (без НДС)", "=B{0}+B{1}+B{2}", "total")
pl_row("rev_w", "Работы", f"={SUMREF('works_net')}", "sub")
pl_row("rev_p", "Запчасти и материалы", f"={SUMREF('parts_net')}", "sub")
pl_row("rev_o", "Прочие доходы (а/м)", f"={OH}!H{I_TOT}", "sub")
pl[f"B{PL['rev']}"].value = f"=B{PL['rev_w']}+B{PL['rev_p']}+B{PL['rev_o']}"
blank()
pl_row("cogs", "СЕБЕСТОИМОСТЬ (без НДС)", "", "total")
pl_row("cogs_p", "Запчасти и материалы", f"={SUMREF('parts_cost')}", "sub")
pl_row("cogs_w", "Прямые расходы по работам (нал.)", f"={SUMREF('cash_work')}", "sub")
pl_row("cogs_o", "Расходы по прочим доходам (а/м)", f"={OH}!H{X_TOT}", "sub")
pl[f"B{PL['cogs']}"].value = f"=B{PL['cogs_p']}+B{PL['cogs_w']}+B{PL['cogs_o']}"
blank()
pl_row("gp", "ВАЛОВАЯ ПРИБЫЛЬ", f"=B{PL['rev']}-B{PL['cogs']}", "total")
blank()
pl_row("opex", "ОПЕРАЦИОННЫЕ РАСХОДЫ (без НДС)", "", "total")
cats = ["Заработная плата", "Аренда", "Инструмент и оборудование", "Услуги (бухгалтерия, ПО)",
        "Банковские комиссии", "Транспорт и доставка"]
cat_rows = []
for cat in cats:
    key = "cat_" + cat
    pl_row(key, cat, f"=SUMIF({OH}!$C${O_FIRST}:$C${O_LAST},A{r},{OH}!$H${O_FIRST}:$H${O_LAST})", "sub")
    cat_rows.append(PL[key])
pl_row("cat_other", "Прочие расходы",
       f"={OH}!H{O_TOT}-SUM(B{cat_rows[0]}:B{cat_rows[-1]})", "sub")
pl[f"A{PL['cat_other']}"].comment = Comment("Все статьи с категорией, не вошедшей в строки выше "
                                            "(в т.ч. «Прочие расходы»).", "script")
pl_row("ins", "Страховые взносы с ФОТ", f"=B{PL['cat_Заработная плата']}*{INS}", "sub")
pl[f"B{PL['ins']}"].font = F_LINK
pl[f"B{PL['opex']}"].value = f"=SUM(B{cat_rows[0]}:B{PL['ins']})"
blank()
pl_row("ebt", "ПРИБЫЛЬ ДО НАЛОГООБЛОЖЕНИЯ", f"=B{PL['gp']}-B{PL['opex']}", "total")
pl_row("tax", "Налог на прибыль", "=0", "normal")  # ссылка заполнится ниже
blank()
pl_row("np", "ЧИСТАЯ ПРИБЫЛЬ", f"=B{PL['ebt']}-B{PL['tax']}", "grand")
blank()
put(pl, f"A{r}", "Справочно", F_BOLD, fill=FILL_SEC)
r += 1
ref_start = r
put(pl, f"A{r}", "Итог по исходному файлу (Итого!F39)", border=BOX)
put(pl, f"B{r}", orig_result if orig_result is not None else 0, F_INPUT, MONEY, border=BOX)
pl[f"B{r}"].comment = Comment("Значение из исходного файла, для сравнения.", "script")
r += 1
put(pl, f"A{r}", "Разница (исправленный − исходный)", border=BOX)
put(pl, f"B{r}", f"=B{PL['np']}-B{ref_start}", fmt=MONEY, border=BOX)
r += 1
put(pl, f"A{r}", "Выручка с НДС (для справки)", border=BOX)
put(pl, f"B{r}", f"={SUMREF('works_gross')}+{SUMREF('parts_gross')}+{OH}!E{I_TOT}", F_LINK, MONEY, border=BOX)
# подставить номер строки выручки в формулы % от выручки
for row in pl.iter_rows(min_row=5, max_row=r, min_col=3, max_col=3):
    c = row[0]
    if isinstance(c.value, str):
        c.value = c.value.replace("{REV}", str(PL["rev"]))
widths(pl, [46, 18, 14])
pl.freeze_panes = "A5"

# ---- Налоги
t = ws_tax
put(t, "A1", f'=\"Расчёт налогов за \"&{q("Параметры")}!B9', F_TITLE)
put(t, "A2", "Расчёт за месяц — оценочный: НДС декларируется за квартал и платится равными долями "
             "в течение 3 месяцев следующего квартала; налог на прибыль — нарастающим итогом.", F_NOTE)
TX = {}
r = 4


def tx_sec(title):
    global r
    put(t, f"A{r}", title, F_BOLD, fill=FILL_SEC)
    put(t, f"B{r}", None, fill=FILL_SEC)
    r += 1


def tx_row(key, label, formula, kind="normal", fmt=MONEY, note=None):
    global r
    font = F_BOLD if kind in ("total", "grand") else F_BASE
    fill = FILL_TOT if kind in ("total", "grand") else None
    link = isinstance(formula, str) and "!" in formula and kind == "sub"
    put(t, f"A{r}", label, font, fill=fill, border=BOX,
        align=Alignment(indent=2 if kind == "sub" else 0))
    put(t, f"B{r}", formula, F_LINK if link else font, fmt, fill, BOX)
    if note:
        put(t, f"C{r}", note, F_NOTE)
    TX[key] = r
    r += 1


tx_sec("НДС")
tx_row("vat_out", "НДС начисленный (с реализации)", "", "total")
tx_row("vo_w", "с работ", f"={SUMREF('works_vat')}", "sub")
tx_row("vo_p", "с запчастей", f"={SUMREF('parts_vat')}", "sub")
tx_row("vo_o", "с прочих доходов (а/м)", f"={OH}!G{I_TOT}", "sub")
t[f"B{TX['vat_out']}"].value = f"=SUM(B{TX['vo_w']}:B{TX['vo_o']})"
tx_row("vat_in", "НДС к вычету (входящий)", "", "total")
tx_row("vi_p", "по закупке запчастей", f"={SUMREF('vat_in')}", "sub",
       note="Только позиции, купленные с НДС (признак 1 на листах проектов)")
tx_row("vi_o", "по общим расходам", f"={OH}!G{O_TOT}", "sub",
       note="По умолчанию 0 — укажите ставку на листе «Общие расходы», где есть счета-фактуры")
tx_row("vi_x", "по расходам на прочие доходы", f"={OH}!G{X_TOT}", "sub")
t[f"B{TX['vat_in']}"].value = f"=SUM(B{TX['vi_p']}:B{TX['vi_x']})"
tx_row("vat_pay", "НДС К УПЛАТЕ ЗА МЕСЯЦ (минус — к возмещению)",
       f"=B{TX['vat_out']}-B{TX['vat_in']}", "grand")
r += 1

tx_sec("НАЛОГ НА ПРИБЫЛЬ")
PLs = q("ОПиУ")
tx_row("inc", "Доходы (выручка без НДС)", f"={PLs}!B{PL['rev']}", "normal")
t[f"B{TX['inc']}"].font = F_LINK
tx_row("exp", "Расходы, уменьшающие налоговую базу", "", "total")
tx_row("e_cogs", "себестоимость запчастей и прямые расходы",
       f"={SUMREF('parts_cost')}+{SUMREF('cash_work')}", "sub")
tx_row("e_x", "расходы по прочим доходам (учитываемые)", f"={OH}!J{X_TOT}", "sub")
tx_row("e_oh", "общие расходы (учитываемые)", f"={OH}!J{O_TOT}", "sub",
       note="Без статей с признаком 0 (напр., «кешбек»)")
tx_row("e_ins", "страховые взносы", f"={PLs}!B{PL['ins']}", "sub")
t[f"B{TX['exp']}"].value = f"=SUM(B{TX['e_cogs']}:B{TX['e_ins']})"
tx_row("nd", "Справочно: расходы, не учитываемые в НУ", f"={OH}!H{O_TOT}-{OH}!J{O_TOT}+{OH}!H{X_TOT}-{OH}!J{X_TOT}",
       "normal")
tx_row("base", "Налоговая база (минус — убыток)", f"=B{TX['inc']}-B{TX['exp']}", "total")
tx_row("ptax", "НАЛОГ НА ПРИБЫЛЬ ЗА МЕСЯЦ", f"=MAX(0,B{TX['base']})*{PTAX}", "grand",
       note="При убытке налог 0; убыток переносится на будущее")
tx_row("pt_fed", "в федеральный бюджет", f"=MAX(0,B{TX['base']})*{PTAX_FED}", "sub")
tx_row("pt_reg", "в бюджет субъекта", f"=MAX(0,B{TX['base']})*{PTAX_REG}", "sub")
r += 1
tx_sec("ИТОГО НАЛОГОВ")
tx_row("all", "НДС + налог на прибыль за месяц", f"=MAX(0,B{TX['vat_pay']})+B{TX['ptax']}", "grand")
widths(t, [58, 18, 80])
# ссылки ОПиУ -> Налоги
pl[f"B{PL['tax']}"].value = f"={q('Налоги')}!B{TX['ptax']}"
pl[f"B{PL['tax']}"].font = F_LINK

# ---- Исправления
x = ws_fix
put(x, "A1", "Что исправлено по сравнению с исходным файлом", F_TITLE)
fixes = [
    ("Налог на прибыль с работ считался как 25% от выручки без НДС (H = G*25%).",
     "Налог на прибыль берётся с прибыли, а не с выручки. Считается один раз по компании: "
     "(доходы − расходы) × 25% — лист «Налоги»."),
    ("Общие расходы (зарплата, аренда и т.д.) вычитались уже после «налога», не уменьшая налоговую базу.",
     "Учитываемые расходы уменьшают базу налога на прибыль; налог ~ на 25% от их суммы меньше."),
    ("Проекты «гарант бетон» и «маз» были на своих листах, но не попали в «Итого» (117 260 руб. с НДС).",
     "Все листы проектов включены в сводку автоматически."),
    ("Не было общей суммы НДС за месяц; входящий НДС по общим расходам не учитывался.",
     "Лист «Налоги»: исходящий НДС − вычеты = НДС к уплате. Ставка вх. НДС по расходам задаётся на листе «Общие расходы»."),
    ("Итог смешивал суммы с НДС и без НДС (выручка без НДС, а общие расходы и нал. расходы — с НДС, "
     "наличные поступления по а/м — без выделения НДС).",
     "В ОПиУ все показатели без НДС: выручка, себестоимость, операционные расходы."),
    ("«Чистая прибыль» считалась по каждой строке работ/запчастей отдельно.",
     "На листах проектов — валовая прибыль (выручка − себестоимость). Чистая прибыль — только по компании в ОПиУ."),
    ("Ставки 22% и 25% были вписаны в сотни формул.",
     "Ставки вынесены на лист «Параметры» — меняются в одной ячейке."),
]
header(x, 3, ["Было", "Стало"])
for i, (a, b) in enumerate(fixes, start=4):
    put(x, f"A{i}", a, border=BOX, align=Alignment(wrap_text=True, vertical="top"))
    put(x, f"B{i}", b, border=BOX, align=Alignment(wrap_text=True, vertical="top"))
widths(x, [70, 80])

# ---- Сверка с исходными данными
from openpyxl.formatting.rule import CellIsRule, FormulaRule

k = ws_chk
TOL = 0.01
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
SV = q("Сводка по проектам")
TXs = q("Налоги")
put(k, "A1", "Сверка итогового отчёта с исходными данными", F_TITLE)
put(k, "A2", f"Исходный файл: {os.path.basename(SRC)}. Колонка «Исходный файл» — значения, посчитанные в "
             f"исходнике (синие, зафиксированы при создании отчёта); «Итоговый отчёт» — формулы этого файла. "
             f"Допуск расхождения — {TOL} руб.", F_NOTE)
put(k, "A3", "Общий результат сверки:", F_BOLD)
status_col_cells = []   # диапазоны колонок статусов


def chk_status(rr):
    put(k, f"F{rr}", f'=IF(ABS(D{rr}-E{rr})<{TOL},"ОК","РАСХОЖДЕНИЕ")', F_BOLD, border=BOX,
        align=Alignment(horizontal="center"))


def chk_row(label, src_val, rep_formula, explained=None, note="", fmt=MONEY):
    global r
    put(k, f"A{r}", label, border=BOX, align=Alignment(wrap_text=True, vertical="center"))
    is_formula = isinstance(src_val, str) and src_val.startswith("=")
    put(k, f"B{r}", src_val, (F_LINK if "!" in src_val else F_BASE) if is_formula else F_INPUT, fmt, border=BOX)
    put(k, f"C{r}", rep_formula, F_LINK if "!" in str(rep_formula) else F_BASE, fmt, border=BOX)
    put(k, f"D{r}", f"=C{r}-B{r}", fmt=fmt, border=BOX)
    put(k, f"E{r}", explained if explained is not None else 0,
        F_LINK if explained is not None else F_BASE, fmt, border=BOX)
    chk_status(r)
    put(k, f"G{r}", note, F_NOTE, border=BOX, align=Alignment(wrap_text=True, vertical="center"))
    r += 1


def chk_section(title, cols):
    global r
    r += 1
    put(k, f"A{r}", title, F_BOLD, fill=FILL_SEC)
    for c in "BCDEFG":
        put(k, f"{c}{r}", None, fill=FILL_SEC)
    r += 1
    header(k, r, cols)
    r += 1


COLS_SRC = ["Проверка", "Исходный файл", "Итоговый отчёт", "Разница", "Объяснённая разница",
            "Статус", "Пояснение"]
COLS_INT = ["Проверка", "Расчёт по составляющим", "Значение в отчёте", "Разница", "Объяснённая разница",
            "Статус", "Пояснение"]
CNT = '#,##0;(#,##0);"0"'
r = 4
first_status_row = None

# 1. Полнота переноса
chk_section("1. ПОЛНОТА ПЕРЕНОСА ДАННЫХ (количество строк)", COLS_SRC)
first_status_row = r
cnt_w = "=" + "+".join(f"COUNT({q(sh)}!C{rf['_w'][0]}:C{rf['_w'][1]})" for _, sh, rf in proj_refs)
cnt_p = "=" + "+".join(f"COUNT({q(sh)}!D{rf['_p'][0]}:D{rf['_p'][1]})" for _, sh, rf in proj_refs)
chk_row("Листов заказ-нарядов", SRC_PROJECT_SHEETS, f"=COUNTA({SV}!B{S_FIRST}:B{S_LAST})", fmt=CNT,
        note="Листы исходника с таблицей работ или запчастей")
chk_row("Строк работ (все проекты)", sum(p_["chk"]["works_n"] for p_ in projects), cnt_w, fmt=CNT)
chk_row("Строк запчастей (все проекты)", sum(p_["chk"]["parts_n"] for p_ in projects), cnt_p, fmt=CNT)
chk_row("Статей общих расходов", SRC_TOTALS["n_opex"], f"=COUNT({OH}!E{O_FIRST}:E{O_LAST})", fmt=CNT)
chk_row("Строк прочих доходов (а/м)", len(other_inc), f"=COUNT({OH}!E{I_FIRST}:E{I_LAST})", fmt=CNT)
chk_row("Строк расходов по прочим доходам", len(other_exp), f"=COUNT({OH}!E{X_FIRST}:E{X_LAST})", fmt=CNT)

# 2. Сверка с листом «Итого» исходника
chk_section("2. СВЕРКА С ЛИСТОМ «ИТОГО» ИСХОДНОГО ФАЙЛА (суммы в учёте, с НДС)", COLS_SRC)
missing = [i for i, p_ in enumerate(projects) if not p_["in_itogo"]]
miss_names = ", ".join(f"«{projects[i]['title']}»" for i in missing)


def explained(key):
    if not missing:
        return None
    return "=" + "+".join(f"{SV}!{SC[key]}{S_FIRST + i}" for i in missing)


note_miss = (f"Разница = проекты, не вошедшие в «Итого» исходника: {miss_names}" if missing else "")
chk_row("Выручка за работы, с НДС", SRC_TOTALS["works"], f"={SUMREF('works_gross')}",
        explained("works_gross"), note_miss)
chk_row("Выручка за запчасти, с НДС", SRC_TOTALS["parts"], f"={SUMREF('parts_gross')}",
        explained("parts_gross"), note_miss)
chk_row("Прямые нал. расходы по работам", SRC_TOTALS["cash_work"], f"={SUMREF('cash_work')}")
chk_row("Общие расходы (в сумме учёта)", SRC_TOTALS["opex"], f"={OH}!E{O_TOT}")
chk_row("Прочие доходы (а/м)", SRC_TOTALS["inc"], f"={OH}!E{I_TOT}")
chk_row("Расходы по прочим доходам", SRC_TOTALS["exp"], f"={OH}!E{X_TOT}")
chk_row("Прочие доходы за вычетом расходов («итого нал чистыми»)", SRC_TOTALS["net_cash"],
        f"={OH}!E{I_TOT}-{OH}!E{X_TOT}")

# 3. Внутренняя согласованность отчёта
chk_section("3. ВНУТРЕННЯЯ СОГЛАСОВАННОСТЬ ОТЧЁТА (цифры одинаковы на всех листах)", COLS_INT)


def all_proj(key):
    return "+".join(f"{q(sh)}!{rf[key]}" for _, sh, rf in proj_refs)


P = lambda key: f"{PLs}!B{PL[key]}"
T = lambda key: f"{TXs}!B{TX[key]}"
chk_row("Выручка без НДС: листы проектов → «Сводка»", "=" + all_proj("rev_net"), f"={SUMREF('rev_net')}")
chk_row("Себестоимость: листы проектов → «Сводка»", "=" + all_proj("cost"), f"={SUMREF('cost')}")
chk_row("Валовая прибыль: листы проектов → «Сводка»", "=" + all_proj("gp"), f"={SUMREF('gp')}")
chk_row("Выручка ОПиУ = проекты + прочие доходы (без НДС)",
        f"={SUMREF('rev_net')}+{OH}!H{I_TOT}", f"={P('rev')}")
chk_row("Выручка с НДС = выручка без НДС + НДС начисленный",
        f"={SUMREF('works_gross')}+{SUMREF('parts_gross')}+{OH}!E{I_TOT}", f"={P('rev')}+{T('vat_out')}")
chk_row("Себестоимость ОПиУ = проекты + расходы по прочим доходам",
        f"={SUMREF('cost')}+{OH}!H{X_TOT}", f"={P('cogs')}")
chk_row("Операционные расходы: статьи по категориям = итог листа «Общие расходы»",
        f"=SUM({PLs}!B{PL['cat_Заработная плата']}:B{PL['cat_other']})", f"={OH}!H{O_TOT}")
chk_row("Операционные расходы ОПиУ = общие расходы без НДС + страховые взносы",
        f"={OH}!H{O_TOT}+{P('ins')}", f"={P('opex')}")
chk_row("Общие расходы: сумма без НДС + входящий НДС = сумма в учёте",
        f"={OH}!H{O_TOT}+{OH}!G{O_TOT}", f"={OH}!E{O_TOT}")
chk_row("Валовая прибыль = выручка − себестоимость", f"={P('rev')}-{P('cogs')}", f"={P('gp')}")
chk_row("Прибыль до налога = валовая прибыль − операционные расходы",
        f"={P('gp')}-{P('opex')}", f"={P('ebt')}")
chk_row("Налог на прибыль: лист «Налоги» → ОПиУ", f"={T('ptax')}", f"={P('tax')}")
chk_row("Чистая прибыль = прибыль до налога − налог", f"={P('ebt')}-{P('tax')}", f"={P('np')}")
chk_row("Доходы для налога = выручка ОПиУ", f"={P('rev')}", f"={T('inc')}")
chk_row("Налоговая база = доходы − учитываемые расходы", f"={T('inc')}-{T('exp')}", f"={T('base')}")
chk_row("Расходы ОПиУ − расходы для налога = не учитываемые в НУ",
        f"={P('cogs')}+{P('opex')}-{T('exp')}", f"={T('nd')}")
chk_row("Налог на прибыль = база × ставка", f"=MAX(0,{T('base')})*{PTAX}", f"={T('ptax')}")
chk_row("Налог на прибыль = федеральная часть + региональная", f"={T('pt_fed')}+{T('pt_reg')}", f"={T('ptax')}")
chk_row("НДС начисленный: листы проектов + прочие доходы → «Налоги»",
        "=" + all_proj("works_vat") + "+" + all_proj("parts_vat") + f"+{OH}!G{I_TOT}", f"={T('vat_out')}")
chk_row("НДС к вычету по запчастям: листы проектов → «Налоги»", "=" + all_proj("vat_in"), f"={T('vi_p')}")
chk_row("НДС к уплате = начисленный − к вычету", f"={T('vat_out')}-{T('vat_in')}", f"={T('vat_pay')}")
last_status_row = r - 1

# 4. По каждому проекту
r += 1
put(k, f"A{r}", "4. СВЕРКА ПО КАЖДОМУ ПРОЕКТУ (исходный лист заказ-наряда → лист проекта в отчёте, с НДС)",
    F_BOLD, fill=FILL_SEC)
r += 1
groups = [("Работы, с НДС", "works_gross", "works_gross"),
          ("Продажа з/ч, с НДС", "parts_sale", "parts_gross"),
          ("Закупка з/ч, с НДС", "parts_buy", "parts_buy"),
          ("НДС входящий по з/ч", "parts_vat_in", "vat_in")]
# двухуровневая шапка
put(k, f"A{r}", None, fill=FILL_HDR)
col = 3
top = r
for g, _, _ in groups:
    k.merge_cells(start_row=top, start_column=col, end_row=top, end_column=col + 2)
    put(k, f"{L(col)}{top}", g, F_HDR, fill=FILL_HDR, align=Alignment(horizontal="center"))
    col += 3
k.merge_cells(start_row=top, start_column=col, end_row=top, end_column=col + 1)
put(k, f"{L(col)}{top}", "Строк работ", F_HDR, fill=FILL_HDR, align=Alignment(horizontal="center"))
k.merge_cells(start_row=top, start_column=col + 2, end_row=top, end_column=col + 3)
put(k, f"{L(col + 2)}{top}", "Строк з/ч", F_HDR, fill=FILL_HDR, align=Alignment(horizontal="center"))
r += 1
sub = ["Проект", "Учтён в «Итого» исходника"]
for _ in groups:
    sub += ["Исходный файл", "Отчёт", "Разница"]
sub += ["Исх.", "Отчёт", "Исх.", "Отчёт", "Статус"]
header(k, r, sub)
for c in range(1, 3):
    put(k, f"{L(c)}{top}", None, fill=FILL_HDR)
r += 1
PT_FIRST = r
STATUS_COL = L(3 + 3 * len(groups) + 4)
for i, ((title, sh, rf), pr) in enumerate(zip(proj_refs, projects)):
    c_ = put(k, f"A{r}", title, Font(name=FONT, size=10, color="0563C1", underline="single"), border=BOX)
    c_.hyperlink = f"#{q(sh)}!A1"
    put(k, f"B{r}", "да" if pr["in_itogo"] else "НЕТ", F_BOLD if not pr["in_itogo"] else F_BASE,
        border=BOX, fill=None if pr["in_itogo"] else FILL_ASSUMP, align=Alignment(horizontal="center"))
    col = 3
    diffs = []
    for _, src_key, ref_key in groups:
        put(k, f"{L(col)}{r}", pr["chk"][src_key], F_INPUT, MONEY, border=BOX)
        put(k, f"{L(col + 1)}{r}", f"={q(sh)}!{rf[ref_key]}", F_LINK, MONEY, border=BOX)
        put(k, f"{L(col + 2)}{r}", f"={L(col + 1)}{r}-{L(col)}{r}", fmt=MONEY, border=BOX)
        diffs.append(f"ABS({L(col + 2)}{r})<{TOL}")
        col += 3
    put(k, f"{L(col)}{r}", pr["chk"]["works_n"], F_INPUT, CNT, border=BOX)
    put(k, f"{L(col + 1)}{r}", f"=COUNT({q(sh)}!C{rf['_w'][0]}:C{rf['_w'][1]})", F_LINK, CNT, border=BOX)
    put(k, f"{L(col + 2)}{r}", pr["chk"]["parts_n"], F_INPUT, CNT, border=BOX)
    put(k, f"{L(col + 3)}{r}", f"=COUNT({q(sh)}!D{rf['_p'][0]}:D{rf['_p'][1]})", F_LINK, CNT, border=BOX)
    diffs += [f"{L(col)}{r}={L(col + 1)}{r}", f"{L(col + 2)}{r}={L(col + 3)}{r}"]
    put(k, f"{STATUS_COL}{r}", f'=IF(AND({",".join(diffs)}),"ОК","РАСХОЖДЕНИЕ")', F_BOLD, border=BOX,
        align=Alignment(horizontal="center"))
    r += 1
PT_LAST = r - 1
put(k, f"A{r}", "ИТОГО", F_BOLD, fill=FILL_TOT, border=BOX)
put(k, f"B{r}", None, fill=FILL_TOT, border=BOX)
for c in range(3, 3 + 3 * len(groups) + 4):
    cl = L(c)
    put(k, f"{cl}{r}", f"=SUM({cl}{PT_FIRST}:{cl}{PT_LAST})", F_BOLD,
        CNT if c >= 3 + 3 * len(groups) else MONEY, FILL_TOT, BOX)
put(k, f"{STATUS_COL}{r}", f'=IF(COUNTIF({STATUS_COL}{PT_FIRST}:{STATUS_COL}{PT_LAST},"ОК")={len(proj_refs)},"ОК","РАСХОЖДЕНИЕ")',
    F_BOLD, fill=FILL_TOT, border=BOX, align=Alignment(horizontal="center"))
PT_TOT = r

# общий результат
bad = (f'COUNTIF(F{first_status_row}:F{last_status_row},"РАСХОЖДЕНИЕ")'
       f'+COUNTIF({STATUS_COL}{PT_FIRST}:{STATUS_COL}{PT_LAST},"РАСХОЖДЕНИЕ")')
put(k, "C3", f'=IF({bad}=0,"ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ","ЕСТЬ РАСХОЖДЕНИЯ: "&({bad}))', F_TITLE)
k.merge_cells("C3:G3")

# подсветка статусов
for rng in (f"F{first_status_row}:F{last_status_row}", f"{STATUS_COL}{PT_FIRST}:{STATUS_COL}{PT_TOT}"):
    k.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"ОК"'], fill=FILL_OK))
    k.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"РАСХОЖДЕНИЕ"'], fill=FILL_BAD))
k.conditional_formatting.add("C3", FormulaRule(formula=['LEFT($C$3,3)="ВСЕ"'], fill=FILL_OK))
k.conditional_formatting.add("C3", FormulaRule(formula=['LEFT($C$3,3)<>"ВСЕ"'], fill=FILL_BAD))
widths(k, [52, 16, 16, 13, 14, 15, 60] + [13] * 14)
k.column_dimensions["B"].width = 28
k.freeze_panes = "B4"

# статус сверки — на первом листе ОПиУ
put(pl, "A3", f'="Сверка с исходными данными: "&{q("Сверка")}!C3', F_BOLD)

# ======================================================================
# ---- Пояснительная записка с мини-дашбордом
# ======================================================================
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule

d = wb.create_sheet("Пояснительная записка", 0)
wb.active = 0
d.sheet_view.showGridLines = False
PLs_, TXs_ = q("ОПиУ"), q("Налоги")
PLc = lambda key: f"{PLs_}!$B${PL[key]}"
TXc = lambda key: f"{TXs_}!$B${TX[key]}"
THR = "'Параметры'!$B$10"
RUB = '#,##0" ₽";[Red]-#,##0" ₽";"–"'
RUB_S = '#,##0;[Red]-#,##0;"–"'
PCT0 = '0.0%;[Red]-0.0%;"–"'
C_DARK, C_BLUE, C_GREEN, C_RED, C_GREY = "1F4E78", "2E75B6", "548235", "C00000", "7F7F7F"
FILL_TILE = PatternFill("solid", fgColor="F3F7FB")
FILL_TOP = PatternFill("solid", fgColor="EAF4E4")
FILL_LOW = PatternFill("solid", fgColor="FBE9E9")
EDGE = lambda color: Border(left=Side(style="thick", color=color))

# ---------- расчёты для текста (Python повторяет формулы отчёта)
V_ = 0.22
py = []
all_lines = []
for pr in projects:
    wg = sum(w["qty"] * w["price"] for w in pr["works"])
    hours = sum(w["qty"] for w in pr["works"])
    ps = pb = pvi = 0.0
    lines = []
    for it in pr["parts"]:
        buy = it["qty"] * it["price"]
        sale = it["fixed_sale"] if it["fixed_sale"] is not None else buy * (1 + it["markup"])
        vi = buy * it["vat_in"] * V_ / (1 + V_)
        ps, pb, pvi = ps + sale, pb + buy, pvi + vi
        cost_net, sale_net = buy - vi, sale / (1 + V_)
        mk = sale_net / cost_net - 1 if cost_net else None
        lines.append(dict(name=it["name"], cost=cost_net, sale=sale_net, markup=mk, proj=pr["title"]))
    all_lines += lines
    wn, pn = wg / (1 + V_), ps / (1 + V_)
    cost = (pb - pvi) + pr["cash_work"]
    rev = wn + pn
    py.append(dict(title=pr["title"], rev=rev, works=wn, parts=pn, parts_cost=pb - pvi, cash=pr["cash_work"],
                   cost=cost, gp=rev - cost, m=(rev - cost) / rev if rev else 0, hours=hours, lines=lines))
inc_net = sum(e["amount"] for e in other_inc) / (1 + V_)
exp_net = sum(e["amount"] for e in other_exp)
opex_py = sum(e["amount"] for e in overheads)
rev_py = sum(p_["rev"] for p_ in py) + inc_net
gp_py = rev_py - sum(p_["cost"] for p_ in py) - exp_net
be = opex_py / rev_py if rev_py else 0            # маржа безубыточности
for p_ in py:
    p_["alloc"] = opex_py * p_["rev"] / rev_py if rev_py else 0
    p_["after"] = p_["gp"] - p_["alloc"]
proj_gp = sum(p_["gp"] for p_ in py)
top = max(py, key=lambda x: x["gp"])
low = min((p_ for p_ in py if p_["rev"] > 0), key=lambda x: x["m"])
big = [p_ for p_ in py if p_["rev"] >= 0.05 * rev_py]
low_big = min(big, key=lambda x: x["m"]) if big else low
parts_cost_all = sum(l_["cost"] for l_ in all_lines)
avg_mk = sum(l_["sale"] for l_ in all_lines) / parts_cost_all - 1 if parts_cost_all else 0
def plural(n, one, few, many):
    n_ = abs(n) % 100
    if 11 <= n_ <= 14:
        return many
    return {1: one, 2: few, 3: few, 4: few}.get(n_ % 10, many)


fm = lambda x: f"{x:,.0f}".replace(",", " ") + " руб."
fp = lambda x: f"{x * 100:.1f}".replace(".", ",") + "%"


def proj_markup(p_):
    c_ = sum(l_["cost"] for l_ in p_["lines"])
    return sum(l_["sale"] for l_ in p_["lines"]) / c_ - 1 if c_ else None


def why_low(p_):
    out = []
    share = p_["parts"] / p_["rev"] if p_["rev"] else 0
    mk = proj_markup(p_)
    if share > 0.5:
        out.append(f"выручка на {fp(share)} состоит из запчастей, а на запчастях маржа ниже, чем на работах")
    if mk is not None and mk < avg_mk:
        out.append(f"средняя наценка на запчасти {fp(mk)} против {fp(avg_mk)} в среднем по сервису")
    if p_["cash"] > 0:
        out.append(f"наличные расходы по работам {fm(p_['cash'])} ({fp(p_['cash'] / p_['rev'])} выручки)")
    cheap = sorted([l_ for l_ in p_["lines"] if l_["markup"] is not None], key=lambda l_: l_["markup"])[:3]
    if cheap:
        out.append("минимальная наценка: " + "; ".join(f"«{l_['name'][:40]}» — {fp(l_['markup'])}" for l_ in cheap))
    return out


top3 = sorted(py, key=lambda x: -x["gp"])[:3]
below = sorted([p_ for p_ in py if p_["rev"] > 0 and p_["m"] < be], key=lambda x: x["after"])
low_mk_lines = [l_ for l_ in all_lines if l_["markup"] is not None and l_["markup"] < 0.3]
works_only = [p_ for p_ in py if p_["parts"] == 0 and p_["cash"] == 0 and p_["works"] > 0]

# ---------- сетка листа
widths(d, [2] + [11.5] * 12 + [2])
r = 1
d.merge_cells("B1:M1")
put(d, "B1", f'="Пояснительная записка к отчёту за "&\'Параметры\'!B9', Font(name=FONT, size=18, bold=True, color="FFFFFF"),
    fill=PatternFill("solid", fgColor=C_DARK), align=Alignment(vertical="center", indent=1))
d.row_dimensions[1].height = 36
d.merge_cells("B2:M2")
put(d, "B2", f'="Все суммы без НДС, в рублях · "&{q("Сверка")}!C3', Font(name=FONT, size=9, color="FFFFFF"),
    fill=PatternFill("solid", fgColor=C_DARK), align=Alignment(indent=1))


def section(row, title):
    d.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
    put(d, f"B{row}", title, Font(name=FONT, size=12, bold=True, color=C_DARK),
        border=Border(bottom=Side(style="medium", color=C_DARK)))
    for c in range(3, 14):
        d.cell(row, c).border = Border(bottom=Side(style="medium", color=C_DARK))
    d.row_dimensions[row].height = 22


def tile(row, col, title, value, fmt, sub, color):
    c1, c2 = L(col), L(col + 2)
    for rr, v, f_, fnt in ((row, title, None, Font(name=FONT, size=9, color=C_GREY)),
                           (row + 1, value, fmt, Font(name=FONT, size=17, bold=True, color=color)),
                           (row + 2, sub, None, Font(name=FONT, size=9, color=C_GREY))):
        d.merge_cells(f"{c1}{rr}:{c2}{rr}")
        put(d, f"{c1}{rr}", v, fnt, f_, FILL_TILE, align=Alignment(horizontal="left", indent=1, vertical="center"))
        for c in range(col, col + 3):
            d.cell(rr, c).fill = FILL_TILE
        d[f"{c1}{rr}"].border = EDGE(color)
    d.row_dimensions[row + 1].height = 30


# ---------- служебная таблица (скрытые колонки P:AD)
H0 = 16  # колонка P
hcols = ["name", "rev", "works", "parts", "cost", "gp", "m", "pshare", "gpshare", "alloc", "after", "hours", "kgp", "km"]
HC = {k_: L(H0 + i) for i, k_ in enumerate(hcols)}
HF = 5
HL = HF + len(proj_refs) - 1
put(d, f"{HC['name']}{HF - 1}", "Служебные расчёты для дашборда", F_BOLD)
for i, k_ in enumerate(hcols):
    put(d, f"{L(H0 + i)}{HF - 1}" if i else f"{HC['name']}{HF - 1}", None if i else "Служебные расчёты", F_BOLD)
for i, (title, sh, rf) in enumerate(proj_refs):
    rr, sr = HF + i, S_FIRST + i
    f_ = {
        "name": f"={SV}!B{sr}", "rev": f"={SV}!{SC['rev_net']}{sr}", "works": f"={SV}!{SC['works_net']}{sr}",
        "parts": f"={SV}!{SC['parts_net']}{sr}", "cost": f"={SV}!{SC['cost']}{sr}", "gp": f"={SV}!{SC['gp']}{sr}",
        "m": f"=IF({HC['rev']}{rr}=0,0,{HC['gp']}{rr}/{HC['rev']}{rr})",
        "pshare": f"=IF({HC['rev']}{rr}=0,0,{HC['parts']}{rr}/{HC['rev']}{rr})",
        "gpshare": f"=IF({SUMREF('gp')}=0,0,{HC['gp']}{rr}/{SUMREF('gp')})",
        "alloc": f"=IF({PLc('rev')}=0,0,{PLc('opex')}*{HC['rev']}{rr}/{PLc('rev')})",
        "after": f"={HC['gp']}{rr}-{HC['alloc']}{rr}",
        "hours": f"=SUM({q(sh)}!C{rf['_w'][0]}:C{rf['_w'][1]})",
        "kgp": f"={HC['gp']}{rr}+ROW()/1000000",
        "km": f"=IF(AND({HC['rev']}{rr}>0,{HC['rev']}{rr}>={THR}),{HC['m']}{rr},1000000000)+ROW()/1000000000",
    }
    for k_ in hcols:
        put(d, f"{HC[k_]}{rr}", f_[k_], F_LINK if "!" in f_[k_] else F_BASE)
HR = lambda k_: f"${HC[k_]}${HF}:${HC[k_]}${HL}"
# данные для графиков структуры
GD = HL + 3
chart_data = [("Работы", f"={PLc('rev_w')}"), ("Запчасти", f"={PLc('rev_p')}"), ("Прочие доходы", f"={PLc('rev_o')}")]
for i, (lb, fx) in enumerate(chart_data):
    put(d, f"{HC['name']}{GD + i}", lb)
    put(d, f"{HC['rev']}{GD + i}", fx, F_LINK, "#,##0")
GW = GD + 5
wf = [("Выручка", f"={PLc('rev')}"), ("Себестоимость", f"={PLc('cogs')}"),
      ("Опер. расходы", f"={PLc('opex')}"), ("Налог", f"={PLc('tax')}"),
      ("Чистая прибыль", f"={PLc('np')}")]
for i, (lb, fx) in enumerate(wf):
    put(d, f"{HC['name']}{GW + i}", lb)
    put(d, f"{HC['rev']}{GW + i}", fx, F_LINK, "#,##0")
for i in range(len(hcols)):
    d.column_dimensions[L(H0 + i)].hidden = True
d.column_dimensions["O"].width = 4

# ---------- 1. KPI-плитки
section(4, "Ключевые показатели месяца")
tile(5, 2, "Выручка без НДС", f"={PLc('rev')}", RUB,
     f'="с НДС: "&FIXED({PLc("rev")}+{TXc("vat_out")},0)', C_BLUE)
tile(5, 5, "Валовая прибыль", f"={PLc('gp')}", RUB,
     f'="маржа "&FIXED(IF({PLc("rev")}=0,0,{PLc("gp")}/{PLc("rev")})*100,1)&"%"', C_BLUE)
tile(5, 8, "Чистая прибыль", f"={PLc('np')}", RUB,
     f'="рентабельность "&FIXED(IF({PLc("rev")}=0,0,{PLc("np")}/{PLc("rev")})*100,1)&"%"', C_GREEN)
tile(5, 11, "Налоги к уплате", f"=MAX(0,{TXc('vat_pay')})+{TXc('ptax')}", RUB,
     f'="НДС "&FIXED(MAX(0,{TXc("vat_pay")}),0)&" · прибыль "&FIXED({TXc("ptax")},0)', C_RED)
tile(9, 2, "Заказ-нарядов", f"=COUNTA({SV}!B{S_FIRST}:B{S_LAST})", "0",
     f'="с работами: "&COUNTIF({HR("works")},">0")', C_DARK)
tile(9, 5, "Средний чек (с НДС)",
     f"=IF(COUNTA({SV}!B{S_FIRST}:B{S_LAST})=0,0,({SUMREF('works_gross')}+{SUMREF('parts_gross')})/COUNTA({SV}!B{S_FIRST}:B{S_LAST}))",
     RUB, f'="максимум: "&FIXED(MAX({HR("rev")})*(1+{VAT}),0)', C_DARK)
tile(9, 8, "Нормо-часов", f"=SUM({HR('hours')})", '#,##0.0',
     f'="ставка н/ч без НДС: "&FIXED(IF(SUM({HR("hours")})=0,0,{PLc("rev_w")}/SUM({HR("hours")})),0)', C_DARK)
tile(9, 11, "Маржа безубыточности", f"=IF({PLc('rev')}=0,0,{PLc('opex')}/{PLc('rev')})", '0.0%',
     '="объект с маржой ниже — убыточен после накладных"', C_RED)
d["K11"].alignment = Alignment(indent=1, wrap_text=True, vertical="center")
d.row_dimensions[11].height = 26

# ---------- 2. Карточки лидера и аутсайдера
section(13, "Лидер и аутсайдер месяца")
GP_HDR = 0  # будет задан ниже (рейтинг) — ссылки на idx
card_rows = [("Выручка без НДС", "rev", RUB), ("  в т.ч. работы", "works", RUB), ("  в т.ч. запчасти", "parts", RUB),
             ("Себестоимость", "cost", RUB), ("Валовая прибыль", "gp", RUB), ("Маржа", "m", PCT0),
             ("Доля в валовой прибыли сервиса", "gpshare", PCT0),
             ("Накладные (пропорц. выручке)", "alloc", RUB), ("Прибыль после накладных", "after", RUB),
             ("Нормо-часов", "hours", '#,##0.0')]


def card(col, title, idx_ref, color, fill):
    c1, c4, c5, c6 = L(col), L(col + 3), L(col + 4), L(col + 5)
    rr = 14
    d.merge_cells(f"{c1}{rr}:{c6}{rr}")
    put(d, f"{c1}{rr}", title, Font(name=FONT, size=9, bold=True, color=color), fill=fill,
        align=Alignment(indent=1))
    rr += 1
    d.merge_cells(f"{c1}{rr}:{c6}{rr}")
    put(d, f"{c1}{rr}", f'=IFERROR(INDEX({HR("name")},{idx_ref}),"—")',
        Font(name=FONT, size=15, bold=True, color=color), fill=fill, align=Alignment(indent=1, vertical="center"))
    d.row_dimensions[rr].height = 26
    for lbl, k_, f_ in card_rows:
        rr += 1
        d.merge_cells(f"{c1}{rr}:{c4}{rr}")
        d.merge_cells(f"{c5}{rr}:{c6}{rr}")
        bold = k_ in ("gp", "m", "after")
        put(d, f"{c1}{rr}", lbl, F_BOLD if bold else F_BASE, fill=fill, align=Alignment(indent=1))
        put(d, f"{c5}{rr}", f'=IFERROR(INDEX({HR(k_)},{idx_ref}),"")', F_BOLD if bold else F_BASE, f_, fill,
            align=Alignment(horizontal="right", indent=1))
    for row_ in range(14, rr + 1):
        for c in range(col, col + 6):
            d.cell(row_, c).fill = fill
        d[f"{c1}{row_}"].border = EDGE(color)
    return rr


# рейтинговые таблицы располагаются ниже; номера строк известны заранее
N = len(proj_refs)
RT = 64 + 0  # заголовок раздела рейтинга (уточняется ниже)
# ---------- 3. Текстовые выводы (формируются при создании отчёта)
paras = []
paras.append(("Общие итоги",
    f"За период {plural(len(py), 'выполнен', 'выполнено', 'выполнено')} {len(py)} {plural(len(py), 'заказ-наряд', 'заказ-наряда', 'заказ-нарядов')}. Выручка без НДС — {fm(rev_py)}, валовая прибыль — {fm(gp_py)} "
    f"(маржа {fp(gp_py / rev_py if rev_py else 0)}). Общие расходы — {fm(opex_py)}, то есть {fp(be)} выручки: "
    f"это минимальная валовая маржа, при которой объект окупает свою долю накладных. "
    f"Три крупнейших объекта ({', '.join('«' + p_['title'] + '»' for p_ in top3)}) дают "
    f"{fp(sum(p_['gp'] for p_ in top3) / proj_gp if proj_gp else 0)} валовой прибыли по заказ-нарядам."))
tmk = proj_markup(top)
paras.append((f"Самый прибыльный объект — «{top['title']}»",
    f"Валовая прибыль {fm(top['gp'])} — {fp(top['gp'] / proj_gp if proj_gp else 0)} всей валовой прибыли по заказ-нарядам, "
    f"маржа {fp(top['m'])}. Выручка {fm(top['rev'])}: работы {fp(top['works'] / top['rev'])} "
    f"({top['hours']:.1f} н/ч), запчасти {fp(top['parts'] / top['rev'])}"
    + (f" со средней наценкой {fp(tmk)}" if tmk is not None else "") +
    f". После распределения накладных объект принёс {fm(top['after'])}. " +
    ("Лидерство обеспечено объёмом: это крупнейший заказ месяца по выручке."
     if top["rev"] == max(p_["rev"] for p_ in py) else
     f"По выручке объект на {sorted(py, key=lambda x: -x['rev']).index(top) + 1}-м месте, "
     f"лидерство обеспечено маржой {fp(top['m'])} против {fp(proj_gp / sum(p_['rev'] for p_ in py))} в среднем.")))
lw = why_low(low)
txt = (f"Маржа {fp(low['m'])} при выручке {fm(low['rev'])}. Причины: " + "; ".join(lw) + ". "
       f"После распределения накладных — {fm(low['after'])}.")
if low is not low_big:
    txt += (f" Объект небольшой ({fp(low['rev'] / rev_py)} выручки), поэтому на итог почти не влияет. "
            f"Среди крупных объектов (от 5% выручки) самая низкая маржа у «{low_big['title']}» — {fp(low_big['m'])} "
            f"при выручке {fm(low_big['rev'])}; причины: " + "; ".join(why_low(low_big)) +
            f". С учётом накладных этот объект дал {fm(low_big['after'])}.")
paras.append((f"Самый низкомаржинальный объект — «{low['title']}»", txt))
if below:
    paras.append(("Объекты ниже маржи безубыточности",
        f"{len(below)} {plural(len(below), 'объект имеет', 'объекта имеют', 'объектов имеют')} маржу ниже {fp(be)} "
        f"и после распределения накладных {plural(len(below), 'даёт', 'дают', 'дают')} минус "
        f"{fm(-sum(p_['after'] for p_ in below))} в сумме: " +
        ", ".join(f"«{p_['title']}» ({fp(p_['m'])}, {fm(p_['after'])})" for p_ in below[:6]) +
        (" и др." if len(below) > 6 else "") + "."))
recs = []
if low_mk_lines:
    recs.append(f"Пересмотреть цены на запчасти с наценкой ниже 30%: таких позиций {len(low_mk_lines)} из "
                f"{len(all_lines)}. Каждые +5 п.п. наценки на весь объём запчастей дают около "
                f"{fm(parts_cost_all * 0.05)} дополнительной валовой прибыли.")
if be < 1:
    k_be = be / (1 - be)
    recs.append(f"Заказ, состоящий только из запчастей, окупает свою долю накладных при наценке от {fp(k_be)} "
                f"(средняя наценка сейчас {fp(avg_mk)}). Отсюда два пути: поднимать цены на запчасти у объектов "
                f"ниже порога либо распределять накладные не по выручке, а по нормо-часам — продажа запчастей почти "
                f"не загружает мастеров, и тогда картина по объектам станет справедливее.")
recs.append("На работах нет прямой себестоимости (зарплата мастеров — в общих расходах), поэтому маржа работ близка "
            f"к 100%{' (' + str(len(works_only)) + ' ' + plural(len(works_only), 'объект', 'объекта', 'объектов') + ' только с работами)' if works_only else ''}. Для честного "
            "сравнения объектов стоит распределять зарплату по нормо-часам — это можно добавить в отчёт.")
recs.append("Проверить допущения, влияющие на прибыль: страховые взносы с ФОТ (по умолчанию не начисляются), "
            "входящий НДС по аренде и оборудованию, документы по наличным расходам.")
paras.append(("Рекомендации", " ".join(f"{i + 1}) {t_}" for i, t_ in enumerate(recs))))

# ---------- размещение: карточки (14..25), графики, текст, рейтинг
end_cards = max(card(2, "▲ САМЫЙ ПРИБЫЛЬНЫЙ ОБЪЕКТ (по валовой прибыли)", "$M${GPF}", C_GREEN, FILL_TOP),
                card(8, "▼ САМЫЙ НИЗКОМАРЖИНАЛЬНЫЙ ОБЪЕКТ", "$M${LMF}", C_RED, FILL_LOW))
r = end_cards + 1
put(d, f"B{r}", f'="Порог выручки для рейтинга маржинальности: "&FIXED({THR},0)&" руб. (меняется на листе «Параметры»)"',
    F_NOTE)
r += 2
section(r, "Графики")
CH_ROW = r + 1
r = CH_ROW + 34
section(r, "Выводы и рекомендации")
put(d, f"B{r + 1}", "Текст сформирован по данным на момент создания отчёта; цифры на плитках, в карточках и графиках — формулы.",
    F_NOTE)
r += 2
for head, body in paras:
    body = body.replace("руб..", "руб.")
    d.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    put(d, f"B{r}", head, Font(name=FONT, size=10, bold=True, color=C_DARK))
    r += 1
    d.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    put(d, f"B{r}", body, F_BASE, align=Alignment(wrap_text=True, vertical="top"))
    d.row_dimensions[r].height = 14 * (len(body) // 125 + 1) + 6
    r += 2

# ---------- 4. Рейтинги
cols_rt = ["№", "Объект", "", "", "Выручка без НДС", "Валовая прибыль", "Маржа", "Доля з/ч",
           "Доля в ВП", "Накладные", "После накладных", "idx"]
keys_rt = [None, "name", None, None, "rev", "gp", "m", "pshare", "gpshare", "alloc", "after"]
fmts_rt = [None, None, None, None, RUB_S, RUB_S, PCT0, PCT0, PCT0, RUB_S, RUB_S]


def rating(start, title, idx_formula):
    section(start, title)
    hr = start + 1
    header(d, hr, cols_rt, col0=2)
    d[f"M{hr}"].value = "служ."
    d.merge_cells(start_row=hr, start_column=3, end_row=hr, end_column=5)
    for i in range(N):
        rr = hr + 1 + i
        put(d, f"M{rr}", idx_formula(i + 1), Font(name=FONT, size=8, color="BFBFBF"))
        put(d, f"B{rr}", f'=IF($M{rr}="","",{i + 1})', border=BOX, align=Alignment(horizontal="center"))
        d.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=5)
        for j, k_ in enumerate(keys_rt):
            if k_ is None:
                continue
            put(d, f"{L(2 + j)}{rr}", f'=IF($M{rr}="","",INDEX({HR(k_)},$M{rr}))', fmt=fmts_rt[j], border=BOX)
        for c in range(3, 13):
            d.cell(rr, c).border = BOX
    return hr + 1, hr + N


RT = r + 1
GPF, GPL = rating(RT, "Рейтинг объектов по валовой прибыли",
                  lambda k_: f"=MATCH(LARGE({HR('kgp')},{k_}),{HR('kgp')},0)")
LT = GPL + 3
LMF, LML = rating(LT, "Рейтинг маржинальности (от низкой к высокой)",
                  lambda k_: f'=IF(SMALL({HR("km")},{k_})>=100000000,"",MATCH(SMALL({HR("km")},{k_}),{HR("km")},0))')
# карточки ссылались на заглушки — подставляем реальные строки
for row in d.iter_rows(min_row=14, max_row=end_cards):
    for c in row:
        if isinstance(c.value, str) and "{GPF}" in c.value or isinstance(c.value, str) and "{LMF}" in c.value:
            c.value = c.value.replace("{GPF}", str(GPF)).replace("{LMF}", str(LMF))
for rng in (f"G{GPF}:G{GPL}", f"G{LMF}:G{LML}"):
    d.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="max", color="70AD47"))
for rng in (f"H{GPF}:H{GPL}", f"H{LMF}:H{LML}"):
    d.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile",
                                                     mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
for rng in (f"L{GPF}:L{GPL}", f"L{LMF}:L{LML}"):
    d.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"],
                                                 fill=PatternFill("solid", fgColor="FBE9E9")))
d.column_dimensions["M"].width = 11.5

# ---------- графики
def _labels(series, num_fmt=None, percent=False):
    lb = DataLabelList()
    lb.showVal = not percent
    lb.showPercent = percent
    lb.showSerName = False
    lb.showCatName = False
    lb.showLegendKey = False
    lb.showLeaderLines = False
    if num_fmt:
        lb.numFmt = num_fmt
    series.dLbls = lb


def _axes(ch):
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.visible_cells_only = False
    ch.legend = None


ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Топ-10 по валовой прибыли"
ch.add_data(Reference(d, min_col=7, min_row=GPF, max_row=GPF + 9), titles_from_data=False)
ch.set_categories(Reference(d, min_col=3, min_row=GPF, max_row=GPF + 9))
ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.numFmt = "#,##0"; ch.y_axis.majorGridlines = None
ch.series[0].graphicalProperties.solidFill = C_GREEN
_labels(ch.series[0], "#,##0"); ch.gapWidth = 40
_axes(ch); ch.width, ch.height = 12.4, 8.5
d.add_chart(ch, f"B{CH_ROW}")

ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Самая низкая маржа"
ch.add_data(Reference(d, min_col=8, min_row=LMF, max_row=LMF + 9), titles_from_data=False)
ch.set_categories(Reference(d, min_col=3, min_row=LMF, max_row=LMF + 9))
ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.numFmt = "0%"; ch.y_axis.majorGridlines = None
ch.y_axis.scaling.min = 0
ch.series[0].graphicalProperties.solidFill = C_RED
_labels(ch.series[0], "0%"); ch.gapWidth = 40
_axes(ch); ch.width, ch.height = 12.4, 8.5
d.add_chart(ch, f"H{CH_ROW}")

hc_name, hc_rev = H0, H0 + 1
ch = DoughnutChart(); ch.title = "Структура выручки"; ch.holeSize = 55
ch.add_data(Reference(d, min_col=hc_rev, min_row=GD, max_row=GD + 2), titles_from_data=False)
ch.set_categories(Reference(d, min_col=hc_name, min_row=GD, max_row=GD + 2))
for i, col_ in enumerate([C_BLUE, "ED7D31", "A5A5A5"]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col_; ch.series[0].dPt.append(pt)
_labels(ch.series[0], "0%", percent=True)
ch.visible_cells_only = False; ch.legend.position = "b"; ch.width, ch.height = 12.4, 8.5
d.add_chart(ch, f"B{CH_ROW + 17}")

ch = BarChart(); ch.type = "col"; ch.style = 10
ch.title = "От выручки к прибыли"
ch.add_data(Reference(d, min_col=hc_rev, min_row=GW, max_row=GW + 4), titles_from_data=False)
ch.set_categories(Reference(d, min_col=hc_name, min_row=GW, max_row=GW + 4))
for i, col_ in enumerate([C_BLUE, "ED7D31", "FFC000", C_RED, C_GREEN]):
    pt = DataPoint(idx=i); pt.graphicalProperties.solidFill = col_; ch.series[0].dPt.append(pt)
_labels(ch.series[0], "#,##0")
ch.y_axis.numFmt = "#,##0"; ch.y_axis.majorGridlines = None; ch.gapWidth = 60
_axes(ch); ch.width, ch.height = 12.4, 8.5
d.add_chart(ch, f"H{CH_ROW + 17}")

d.freeze_panes = "A3"
d.sheet_view.zoomScale = 90

_existed = os.path.exists(DST)
print("Сохраняю итоговый файл...")
wb.save(DST)
print(f"\nГотово! Файл отчёта {'ЗАМЕНЁН (файл с таким именем уже был)' if _existed else 'создан'}:\n{DST}")
# открыть готовый файл, чтобы сразу было видно результат
try:
    if sys.platform.startswith("win"):
        os.startfile(DST)
    elif sys.platform == "darwin":
        import subprocess
        subprocess.Popen(["open", DST])
except Exception:
    pass
print(f"Проектов: {len(projects)}, общих расходов: {len(overheads)}, "
      f"прочих доходов: {len(other_inc)}, прочих расходов: {len(other_exp)}")
_pause()
