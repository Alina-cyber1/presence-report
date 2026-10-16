from pathlib import Path
import math, re, os, pickle
from datetime import datetime
from flask import Flask, request, send_file, redirect, url_for
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / 'uploads'
EXPORT_DIR = BASE_DIR / 'exports'
CONFIG_DIR = BASE_DIR / 'config'
HISTORY_DIR = CONFIG_DIR / 'history'
STATE_FILE = CONFIG_DIR / 'state.pkl'
LAST_FILE = UPLOAD_DIR / 'last.xlsx'
for folder in (UPLOAD_DIR, EXPORT_DIR, CONFIG_DIR, HISTORY_DIR):
    folder.mkdir(exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

NO_DEPT = 'Без департамента'
NO_UNIT = 'Без подразделения'

MONTHS_RU = {
    'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4, 'май': 5, 'июн': 6,
    'июл': 7, 'август': 8, 'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12
}

def txt(v):
    return '' if v is None else str(v).strip()

def num(v):
    try:
        return float(txt(v).replace(' ', '').replace(',', '.'))
    except Exception:
        return 0.0

def hm(v):
    m = re.search(r'(?<!\d)(\d{1,2})[:\-.](\d{2})(?!\d)', txt(v))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None

def fm(v):
    if v is None: return '—'
    sign = '-' if v < 0 else ''
    v = abs(round(v))
    return f'{sign}{v // 60:02d}:{v % 60:02d}'

def fh(v): return fm(v * 60)
def fsigned(v): return ('+' if round(v) > 0 else '') + fm(v)
def esc(v):
    return (txt(v).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace(chr(34), '&quot;'))

def parse_interval(v):
    m = re.search(r'(\d{1,2}[:\-.]\d{2})\s*[-–—]\s*(\d{1,2}[:\-.]\d{2})', txt(v))
    if not m: return None, None
    start, end = hm(m.group(1)), hm(m.group(2))
    if start is not None and end is not None and end < start:
        end += 1440
    return start, end

def parse_schedule(v):
    source = txt(v)
    low = source.lower().replace('ё', 'е')
    if not source:
        return {'type': 'missing', 'start': None, 'hours': None}
    if 'свободн' in low:
        return {'type': 'free', 'start': None, 'hours': None}
    m = re.search(r'с\s*(\d{1,2})\s*[-:]\s*(\d{2})', low)
    start = int(m.group(1)) * 60 + int(m.group(2)) if m else None
    m = re.search(r'\((\d+(?:[\.,]\d+)?)\s*час', low)
    hours = float(m.group(1).replace(',', '.')) if m else None
    typ = 'five_two' if 'пятиднев' in low else ('shift' if 'сменн' in low else 'unknown')
    return {'type': typ, 'start': start, 'hours': hours}

def format_kind(v):
    low = txt(v).lower().replace('ё', 'е')
    if 'гибрид' in low: return 'hybrid'
    if 'удален' in low: return 'remote'
    return 'office'

def parse_period(head_text):
    low = txt(head_text).lower().replace('ё', 'е')
    month = None
    for key, val in MONTHS_RU.items():
        if key in low:
            month = val
            break
    m = re.search(r'20\d{2}', low)
    year = int(m.group()) if m else datetime.now().year
    if month is None:
        month = datetime.now().month
    return year, month

def is_employee_row(row):
    """
    Структура после prepare_upload:
    A (0) — Департамент
    B (1) — №
    C (2) — Подразделение
    D (3) — Сотрудник (ФИО)
    E (4) — Табельный номер (0000-XXXXX)
    Не считаем сотрудником строки, где в D не ФИО, а «Руководство …» или «Курьер …».
    """
    if row is None: return False
    number = txt(row[1]) if len(row) > 1 else ''
    name = txt(row[3]) if len(row) > 3 else ''
    personnel = txt(row[4]) if len(row) > 4 else ''
    if not re.fullmatch(r'\d+', number):
        return False
    if not re.fullmatch(r'\d{4}-\d{5}', personnel):
        return False
    if not name or name.lower() == 'сотрудник':
        return False
    low = name.lower().replace('ё', 'е')
    for kw in ('руководство', 'курьер'):
        if kw in low:
            return False
    words = [w for w in re.split(r'\s+', name) if w]
    if len(words) < 2:
        return False
    for w in words:
        if len(w) < 2 or not w[0].isupper():
            return False
    return True

def strict_day(rule, missing, start, end, rv, skud):
    z = {'plan': 0, 'fact': 0, 'late': 0, 'early': 0, 'late_leave': 0,
         'off': 0, 'early_arr': 0, 'remote_rv': 0, 'free': 0}
    has_int = start is not None and end is not None
    presence = max(end - start, 0) if has_int else round(skud * 60)
    if rule['type'] == 'free':
        if has_int or skud > 0: z['free'] = presence
        return z
    if rv <= 0:
        if has_int or skud > 0: z['off'] = presence
        return z
    planm = round(rv * 60)
    if not has_int and skud <= 0:
        z['remote_rv'] = planm
        return z
    if missing:
        z['plan'] = planm
        z['fact'] = round(skud * 60) if skud > 0 else presence
        return z
    z['plan'] = planm
    if rule['start'] is None:
        z['fact'] = presence
        return z
    ps = rule['start']
    pe = ps + planm
    if has_int:
        worked = max(0, min(end, pe) - max(start, ps))
        extra = max(0, end - max(start, pe))
        z['late'] = min(max(0, start - ps), planm)
        z['early'] = min(max(0, pe - end), planm)
        z['early_arr'] = max(0, min(end, ps) - start)
        z['late_leave'] = extra
        z['fact'] = worked + extra
    else:
        z['fact'] = round(skud * 60)
    return z

def load_departments():
    path = BASE_DIR / 'Справочник.xlsx'
    if not path.exists():
        return {}
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb['Справочник'] if 'Справочник' in wb.sheetnames else wb.active
        m = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            name = txt(row[1]) if len(row) > 1 else ''
            dept = txt(row[5]) if len(row) > 5 else ''
            if not dept or dept.startswith('='):
                dept = txt(row[4]) if len(row) > 4 else ''
            if name and dept and not dept.startswith('='):
                m[name] = dept
        wb.close()
        return m
    except Exception:
        return {}

def prepare_upload(src_path, dst_path):
    """
    Заполняет пустую колонку A «Департаментом» через справочник.
    НЕ пересобирает файл — сохраняет его как есть.
    """
    DEPT_MAP = load_departments()

    wb = load_workbook(src_path, data_only=False, keep_links=False)
    ws = wb.active

    header_row = None
    for r in range(1, min(20, ws.max_row) + 1):
        v = ws.cell(row=r, column=2).value
        if txt(v) == '№':
            header_row = r
            break
    if header_row is None:
        header_row = 4

    ws.cell(row=header_row, column=1).value = 'Департамент'

    for r in range(header_row + 1, ws.max_row + 1):
        unit = txt(ws.cell(row=r, column=3).value)
        if not unit:
            continue
        dept = DEPT_MAP.get(unit, '') or unit
        ws.cell(row=r, column=1).value = dept

    wb.save(dst_path)
    wb.close()

def analyze(path):
    """
    A (0)  — Департамент
    B (1)  — №
    C (2)  — Подразделение
    D (3)  — Сотрудник (ФИО)
    E (4)  — Табельный
    F (5)  — Должность
    G (6)  — Дата приема
    H (7)  — Дата увольнения
    I (8)  — Формат работы
    J (9)  — График
    K (10) — Болезнь (РВ)
    L (11) — Болезнь (СКУД)
    ...
    AE (30) — Разница СБ и плановых часов
    AF (31) — 01 Сб — РВ
    AG (32) — 01 Сб — СКУД
    """
    COL_DEPARTMENT = 0
    COL_NUMBER     = 1
    COL_UNIT       = 2
    COL_NAME       = 3
    COL_PERSONNEL  = 4
    COL_POSITION   = 5
    COL_FORMAT     = 8
    COL_SCHEDULE   = 9
    COL_DAY_START  = 31
    COL_DAY_STEP   = 2

    DEPT_MAP = load_departments()
    wb = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    year, month = datetime.now().year, datetime.now().month
    for idx, row in enumerate(rows_iter):
        joined = ' '.join(txt(c) for c in row[:15])
        if re.search(r'\bза\s+[А-Яа-я]+', joined) or re.search(r'\b[А-Яа-я]+\s+20\d{2}', joined):
            year, month = parse_period(joined)
            break
        if idx > 10:
            break

    employees = []
    days = []
    window = []
    for row in rows_iter:
        window.append(row)
        if len(window) > 3:
            window.pop(0)
        if len(window) == 3 and is_employee_row(window[0]):
            emp_row = window[0]
            data_row = window[2]

            department_direct = txt(emp_row[COL_DEPARTMENT]) if len(emp_row) > COL_DEPARTMENT else ''
            unit = txt(emp_row[COL_UNIT]) if len(emp_row) > COL_UNIT else ''
            department = department_direct or DEPT_MAP.get(unit, '') or unit or NO_DEPT

            number    = txt(emp_row[COL_NUMBER])    if len(emp_row) > COL_NUMBER    else ''
            name      = txt(emp_row[COL_NAME])      if len(emp_row) > COL_NAME      else ''
            personnel = txt(emp_row[COL_PERSONNEL]) if len(emp_row) > COL_PERSONNEL else ''
            position  = txt(emp_row[COL_POSITION])  if len(emp_row) > COL_POSITION  else ''
            work_format = txt(emp_row[COL_FORMAT])  if len(emp_row) > COL_FORMAT    else ''
            if not work_format:
                work_format = 'Не указан'
            fmt = format_kind(work_format)
            raw_schedule = txt(emp_row[COL_SCHEDULE]) if len(emp_row) > COL_SCHEDULE else ''
            rule = parse_schedule(raw_schedule)
            is_missing = rule['type'] == 'missing'

            e = {
                'number': number, 'name': name, 'position': position,
                'department': department, 'unit': unit, 'personnel': personnel,
                'work_format': work_format,
                'schedule': raw_schedule or 'График не указан',
                'rule': rule,
                'summary': {
                    'work': 0, 'remote': 0, 'late': 0, 'early': 0, 'anomaly': 0,
                    'late_leave': 0, 'above': 0.0, 'below': 0.0, 'off': 0,
                    'problems': 0, 'excluded': False, 'termination': is_missing,
                    'plan': 0, 'fact': 0
                }
            }

            for day in range(1, 32):
                c = COL_DAY_START + (day - 1) * COL_DAY_STEP
                if c + 1 >= len(emp_row) or c + 1 >= len(data_row):
                    break
                try:
                    date = datetime(year, month, day).strftime('%d.%m.%Y')
                except Exception:
                    continue
                raw = emp_row[c] if c < len(emp_row) else None
                rv = num(data_row[c]) if c < len(data_row) else 0.0
                skud = num(data_row[c + 1]) if c + 1 < len(data_row) else 0.0
                start, end = parse_interval(raw)
                statuses = []
                kind = ''
                plan = '—'
                m = strict_day(rule, is_missing, start, end, rv, skud)
                e['summary']['plan'] += m['plan']
                e['summary']['fact'] += m['fact']

                if is_missing:
                    if rv > 0:
                        e['summary']['work'] += 1
                        if skud > 0:
                            d = (skud - rv) * 60
                            if d > 5:
                                statuses.append('Отклонение длительности СКУД от РВ: +' + fm(d))
                                e['summary']['above'] += skud - rv
                            elif d < -5:
                                statuses.append('Отклонение длительности СКУД от РВ: ' + fm(d))
                                e['summary']['below'] += rv - skud
                            else:
                                statuses.append('СКУД соответствует РВ')
                        else:
                            statuses = ['Удалённая работа (потенциально)']
                            kind = 'remote'
                            e['summary']['remote'] += 1
                elif rule['type'] == 'free':
                    if raw or skud > 0:
                        statuses = ['Свободное посещение/Подработка']
                        kind = 'free'
                elif rv <= 0:
                    if raw or skud > 0:
                        statuses = ['Посещение вне графика / переработка']
                        kind = 'off_schedule'
                        e['summary']['off'] += 1
                elif fmt == 'remote' and not raw and skud <= 0:
                    statuses = ['Удалённая работа']
                    kind = 'remote'
                    e['summary']['remote'] += 1
                    e['summary']['work'] += 1
                elif fmt == 'hybrid' and not raw and skud <= 0:
                    statuses = ['Гибрид/удалённая работа']
                    kind = 'remote'
                    e['summary']['remote'] += 1
                    e['summary']['work'] += 1
                else:
                    e['summary']['work'] += 1
                    if rule['start'] is None:
                        statuses = ['Нераспознанный график']
                        kind = 'unknown'
                        e['summary']['problems'] += 1
                    else:
                        pe = rule['start'] + rv * 60
                        plan = f"{fm(rule['start'])}–{fm(pe % 1440)}" + (' (+1 день)' if pe >= 1440 else '')
                        if start is None or end is None:
                            statuses = ['Проблема: нет полного интервала']
                            kind = 'problem'
                            e['summary']['problems'] += 1
                        else:
                            if start < rule['start'] - 10:
                                statuses.append('Аномально ранний приход ' + fm(rule['start'] - start))
                                e['summary']['anomaly'] += 1
                            elif start > rule['start']:
                                statuses.append('Опоздание ' + fm(start - rule['start']))
                                e['summary']['late'] += 1
                            elif start < rule['start']:
                                statuses.append('Ранний приход ' + fm(rule['start'] - start))
                            delta = end - pe
                            if delta < 0:
                                statuses.append('Ранний уход ' + fm(-delta))
                                e['summary']['early'] += 1
                            elif delta > 5:
                                statuses.append('Поздний уход ' + fm(delta))
                                e['summary']['late_leave'] += 1
                            d = (skud - rv) * 60
                            if d > 5:
                                statuses.append('Отклонение длительности СКУД от РВ: +' + fm(d))
                                e['summary']['above'] += skud - rv
                            elif d < -5:
                                statuses.append('Отклонение длительности СКУД от РВ: ' + fm(d))
                                e['summary']['below'] += rv - skud
                            else:
                                statuses.append('СКУД соответствует РВ')

                days.append({
                    'number': number, 'name': name, 'position': position,
                    'department': department, 'unit': unit,
                    'work_format': work_format, 'date': date,
                    'rv': rv, 'skud': skud, 'interval': txt(raw) or '—',
                    'plan': plan, 'statuses': statuses, 'kind': kind, 'm': m
                })

            employees.append(e)
            window = []

    wb.close()
    meta = {'year': year, 'month': month, 'period_key': f'{year:04d}-{month:02d}'}
    return employees, days, meta

COLUMN_DESCRIPTIONS = [
    ('Рабочие дни', 'Количество дней, когда по графику было запланировано присутствие.'),
    ('Удалённые дни', 'Дни без визита в офис и без СКУД.'),
    ('Опоздания', 'Дни прихода позже начала смены. С первой минуты, без допуска.'),
    ('Ранние уходы', 'Дни ухода раньше планового окончания. С первой минуты.'),
    ('Аномально ранние приходы', 'Приход раньше плана более чем на 10 минут.'),
    ('Поздние уходы', 'Задержка после планового окончания более чем на 5 минут.'),
    ('СКУД выше РВ', 'Суммарные часы превышения присутствия над нормой.'),
    ('СКУД ниже РВ', 'Суммарные часы недостачи присутствия.'),
    ('Статус формирования отчёта', '"Сформирован" или "Увольнение сотрудника" (график не указан).'),
]

DASH_DESCRIPTIONS = [
    ('Сотрудники', 'Количество сотрудников.'),
    ('План (РВ), ч', 'Сумма нормы РВ по рабочим дням.'),
    ('Факт, ч', 'Фактически отработано от прихода до ухода.'),
    ('Разница план/факт, ч', 'Факт минус план.'),
    ('Опоздания, ч', 'Суммарное время опозданий.'),
    ('Ранние уходы, ч', 'Суммарное время ранних уходов.'),
    ('Поздние уходы, ч', 'Суммарное время после планового окончания смены.'),
    ('Вне графика, ч', 'Работа в дни, когда смена по графику не запланирована (РВ=0).'),
    ('Ранний приход, ч (справочно)', 'Справочно, в факт и разницу не входит.'),
    ('Потенциально удалённые, ч', 'Норма РВ дней без интервала и СКУД.'),
    ('Свободный график, ч', 'Время по свободному графику, без нормы.'),
]

DKEYS = ['plan', 'fact', 'late', 'early', 'late_leave', 'off', 'early_arr', 'remote_rv', 'free']

COL_TOGGLES = [
    ('plan',      'План (РВ), ч'),
    ('fact',      'Факт, ч'),
    ('diff',      'Количество часов недоработки по графику'),
    ('norm',      'Статус нормы'),
    ('pct',       'Отклонение, %'),
    ('late',      'Опоздания, ч'),
    ('early',     'Ранние уходы, ч'),
    ('off',       'Вне графика, ч'),
    ('early_arr', 'Ранний приход, ч (справочно)'),
    ('remote_rv', 'Потенциально удалённые, ч'),
    ('free',      'Свободный график, ч'),
]
ALL_COLS = {k for k, _ in COL_TOGGLES}

def safe_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|]', '_', name).strip()
    return cleaned or 'export'

def filter_bar(tab, emps, dept_sel='', unit_sel='', emp_sel='', show_employee=False):
    dept_sel = str(dept_sel or '')
    unit_sel = str(unit_sel or '')
    emp_sel = str(emp_sel or '')

    depts = sorted({e.get('department', NO_DEPT) for e in emps})
    dept_opts = ''.join(
        '<option value="' + esc(d) + '"' + (' selected' if d == dept_sel else '') + '>' +
        esc(d) + '</option>' for d in depts
    )
    units_pool = [e for e in emps if not dept_sel or e.get('department') == dept_sel]
    units = sorted({e.get('unit', NO_UNIT) for e in units_pool})
    unit_opts = ''.join(
        '<option value="' + esc(u) + '"' + (' selected' if u == unit_sel else '') + '>' +
        esc(u) + '</option>' for u in units
    )

    emp_html = ''
    if show_employee:
        names = sorted({e['name'] for e in emps
                        if (not dept_sel or e.get('department') == dept_sel)
                        and (not unit_sel or e.get('unit') == unit_sel)})
        name_opts = ''.join('<option value="' + esc(n) + '">' for n in names)
        emp_html = (
            '<label>Сотрудник</label>'
            '<input list="names" name="employee" value="' + esc(emp_sel) + '" '
            'placeholder="Начните вводить ФИО" autocomplete="off">'
            '<datalist id="names">' + name_opts + '</datalist>'
        )

    return (
        '<div class="card"><form method="get" action="/" class="select-row">'
        '<input type="hidden" name="tab" value="' + esc(tab) + '">'
        '<label>Департамент</label>'
        '<select name="department"><option value="">Все</option>' + dept_opts + '</select>'
        '<label>Подразделение</label>'
        '<select name="unit"><option value="">Все</option>' + unit_opts + '</select>'
        + emp_html +
        '<button class="btn">Показать</button>'
        '<a class="btn btn-ghost" href="/?tab=' + esc(tab) + '">Сбросить</a>'
        '</form></div>'
    )

def render_cols_form(tab, active, dept='', unit=''):
    boxes = []
    for key, label in COL_TOGGLES:
        checked = ' checked' if key in active else ''
        boxes.append(
            '<label><input type="checkbox" name="col_' + key + '" value="1"' + checked + '>'
            + esc(label) + '</label>'
        )
    return (
        '<div class="card">'
        '<h2>Показывать показатели</h2>'
        '<form method="get" action="/" class="cols-form">'
        '<input type="hidden" name="tab" value="' + esc(tab) + '">'
        '<input type="hidden" name="department" value="' + esc(dept) + '">'
        '<input type="hidden" name="unit" value="' + esc(unit) + '">'
        + ''.join(boxes) +
        '<button class="btn" type="submit">Применить и сохранить</button>'
        '</form></div>'
    )

def collect_cols(args):
    any_marker = any(k.startswith('col_') for k in args.keys())
    if not any_marker:
        return set(ALL_COLS)
    active = set()
    for key, _ in COL_TOGGLES:
        if args.get('col_' + key):
            active.add(key)
    return active

def cols_to_url(active):
    if active == ALL_COLS:
        return ''
    return ','.join(sorted(active))

def render_pie(title, parts):
    total = sum(v for _, v, _ in parts) or 1
    R = 60
    C = 2 * 3.14159265 * R
    svg_parts = []
    offset = 0.0
    for _, v, color in parts:
        frac = (v / total) if total else 0
        dash = C * frac
        svg_parts.append(
            '<circle r="' + str(R) + '" cx="80" cy="80" fill="transparent" '
            'stroke="' + color + '" stroke-width="28" '
            'stroke-dasharray="' + f'{dash:.2f} {C - dash:.2f}' + '" '
            'stroke-dashoffset="' + f'{-offset:.2f}' + '" '
            'transform="rotate(-90 80 80)"></circle>'
        )
        offset += dash
    svg = ('<svg width="160" height="160" viewBox="0 0 160 160">'
           + ''.join(svg_parts) +
           '<text x="80" y="78" text-anchor="middle" font-size="13" fill="#6b7684">Всего</text>'
           '<text x="80" y="98" text-anchor="middle" font-size="16" font-weight="700" fill="#1b2430">'
           + fm(total) + '</text>'
           '</svg>')
    legend = []
    for label, v, color in parts:
        pct = (v / total * 100) if total else 0
        legend.append(
            '<div><span class="dot" style="background:' + color + '"></span>' +
            esc(label) + '<span class="val">' + fm(v) + ' · ' +
            f'{pct:.1f}%</span></div>'
        )
    return ('<div class="card"><h2>' + esc(title) + '</h2>'
            '<div class="pie">' + svg + '<div class="pie-legend">' +
            ''.join(legend) + '</div></div>'
            '<p class="note">График отражает только выбранные категории.</p></div>')

def export_xlsx(emps, days, path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Анализ данных'
    headers = ['№', 'Департамент', 'Подразделение', 'Сотрудник', 'Должность',
               'Табельный номер', 'Формат работы', 'График',
               'Рабочие дни', 'Удалённые дни', 'Опоздания', 'Ранние уходы',
               'Аномально ранние приходы', 'Поздние уходы >5 мин',
               'СКУД выше РВ, ч', 'СКУД ниже РВ, ч', 'Вне графика, дней',
               'Проблемные дни', 'Статус формирования отчёта']
    ws.append(headers)
    header_fill = PatternFill(start_color='2F6FED', end_color='2F6FED', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    thin = Side(style='thin', color='D9DEE3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 32
    green = PatternFill(start_color='E3F6E9', end_color='E3F6E9', fill_type='solid')
    gray = PatternFill(start_color='FDECEC', end_color='FDECEC', fill_type='solid')
    for e in emps:
        s = e['summary']
        status = 'Увольнение сотрудника' if s['termination'] else 'Сформирован'
        ws.append([e['number'], e.get('department', ''), e.get('unit', ''),
                   e['name'], e.get('position', ''), e['personnel'],
                   e['work_format'], e['schedule'],
                   s['work'], s['remote'], s['late'], s['early'], s['anomaly'],
                   s['late_leave'], round(s['above'], 2), round(s['below'], 2),
                   s['off'], s['problems'], status])
        r = ws.max_row
        ws.cell(r, 19).fill = gray if s['termination'] else green
        for c in range(1, 20):
            ws.cell(r, c).border = border
            if c not in (4, 5, 7, 8):
                ws.cell(r, c).alignment = center
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    widths = {1: 8, 2: 28, 3: 28, 4: 32, 5: 28, 6: 16, 7: 16, 8: 40,
              9: 11, 10: 12, 11: 11, 12: 11, 13: 14, 14: 14,
              15: 12, 16: 12, 17: 12, 18: 12, 19: 22}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    wd = wb.create_sheet('Дни')
    dheaders = ['№', 'Департамент', 'Подразделение', 'Сотрудник', 'Должность',
                'Формат работы', 'Дата', 'РВ', 'СКУД', 'Интервал', 'План', 'Статусы']
    wd.append(dheaders)
    for cell in wd[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    wd.row_dimensions[1].height = 28
    bad = PatternFill(start_color='FDECEC', end_color='FDECEC', fill_type='solid')
    warn = PatternFill(start_color='FFF6DF', end_color='FFF6DF', fill_type='solid')
    blue = PatternFill(start_color='EAF1FF', end_color='EAF1FF', fill_type='solid')
    for d in days:
        wd.append([d['number'], d.get('department', ''), d.get('unit', ''),
                   d['name'], d.get('position', ''), d['work_format'],
                   d['date'], d['rv'], d['skud'], d['interval'],
                   d['plan'], '; '.join(d['statuses'])])
        st = d['statuses']
        b = any(x.startswith(('Опоздание', 'Ранний уход', 'Аномально', 'Проблема', 'Нераспознанный')) for x in st)
        w = any(x.startswith(('Поздний уход', 'Отклонение длительности')) for x in st)
        fill = bad if b else (warn if w else (blue if d['kind'] in ('remote', 'free', 'off_schedule') else None))
        if fill:
            r = wd.max_row
            for c in range(1, 13):
                wd.cell(r, c).fill = fill
    wd.freeze_panes = 'A2'
    wd.auto_filter.ref = wd.dimensions
    dwidths = {1: 8, 2: 28, 3: 28, 4: 32, 5: 28, 6: 16, 7: 12, 8: 9, 9: 9, 10: 18, 11: 20, 12: 60}
    for c, w in dwidths.items():
        wd.column_dimensions[get_column_letter(c)].width = w

    de = wb.create_sheet('Описание значений')
    de.append(['Колонка', 'Описание'])
    for cell in de[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    for name, desc in COLUMN_DESCRIPTIONS:
        de.append([name, desc])
        r = de.max_row
        de.cell(r, 1).font = Font(bold=True)
        de.cell(r, 1).alignment = Alignment(vertical='top', wrap_text=True)
        de.cell(r, 2).alignment = Alignment(vertical='top', wrap_text=True)
        de.cell(r, 1).border = border
        de.cell(r, 2).border = border
    de.column_dimensions['A'].width = 28
    de.column_dimensions['B'].width = 100

    wb.save(path)

def dashboard_data(emps, days, group_by='department', filter_dept='', filter_unit=''):
    agg = {}
    for e in emps:
        if filter_dept and e.get('department') != filter_dept:
            continue
        if filter_unit and e.get('unit') != filter_unit:
            continue
        key = (e.get('department') or NO_DEPT) if group_by == 'department' else (e.get('unit') or NO_UNIT)
        a = agg.setdefault(key, {'emp': 0, **{k: 0 for k in DKEYS}})
        a['emp'] += 1
    for x in days:
        if filter_dept and x.get('department') != filter_dept:
            continue
        if filter_unit and x.get('unit') != filter_unit:
            continue
        key = (x.get('department') or NO_DEPT) if group_by == 'department' else (x.get('unit') or NO_UNIT)
        a = agg.get(key)
        m = x.get('m')
        if a is None or not m:
            continue
        for k in DKEYS:
            a[k] += m[k]
    rows = []
    for name, a in agg.items():
        r = dict(a)
        r['name'] = name
        r['diff'] = r['fact'] - r['plan']
        rows.append(r)
    rows.sort(key=lambda r: (-r['diff'], r['name']))
    total = {'emp': sum(r['emp'] for r in rows), **{k: sum(r[k] for r in rows) for k in DKEYS}}
    total['diff'] = total['fact'] - total['plan']
    return rows, total

def export_dashboard_xlsx(emps, days, group_by, filter_dept, filter_unit, path, sheet_name):
    rows, total = dashboard_data(emps, days, group_by, filter_dept, filter_unit)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    header_fill = PatternFill(start_color='2F6FED', end_color='2F6FED', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    thin = Side(style='thin', color='D9DEE3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title = 'Департамент' if group_by == 'department' else 'Подразделение'

    ws.append([title] + [n for n, _ in DASH_DESCRIPTIONS] + ['Статус нормы', 'Отклонение, %'])
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 48

    def hrs(v): return round(v / 60, 2)

    def line(label, r):
        diff_min = r['diff']
        if diff_min < -5:
            norm_status = 'Недоработка'
        elif diff_min > 5:
            norm_status = 'Переработка'
        else:
            norm_status = 'Норма'
        pct = round((r['diff'] / r['plan'] * 100), 1) if r['plan'] else 0.0
        return [label, r['emp'], hrs(r['plan']), hrs(r['fact']), hrs(r['diff']),
                hrs(r['late']), hrs(r['early']), hrs(r['late_leave']), hrs(r['off']),
                hrs(r['early_arr']), hrs(r['remote_rv']), hrs(r['free']),
                norm_status, pct]

    red = PatternFill(start_color='FDECEC', end_color='FDECEC', fill_type='solid')
    grn = PatternFill(start_color='E3F6E9', end_color='E3F6E9', fill_type='solid')

    for r in rows:
        ws.append(line(r['name'], r))
        n = ws.max_row
        if r['diff'] < 0:
            ws.cell(n, 6).fill = red
        elif r['diff'] > 0:
            ws.cell(n, 6).fill = grn

    ws.append(line('ИТОГО', total))
    n = ws.max_row
    for c in range(1, 15):
        ws.cell(n, c).font = Font(bold=True)

    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in r:
            cell.border = border
            if cell.column > 1:
                cell.alignment = Alignment(horizontal='center')

    ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = 34
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 16

    wb.save(path)

def export_dynamics_xlsx(history_data, path, filter_dept='', filter_unit=''):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Динамика'
    header_fill = PatternFill(start_color='2F6FED', end_color='2F6FED', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    thin = Side(style='thin', color='D9DEE3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(['Период', 'Сотрудники', 'План (РВ), ч', 'Факт, ч', 'Разница, ч',
               'Опоздания, ч', 'Ранние уходы, ч', 'Поздние уходы, ч', 'Вне графика, ч',
               'Δ разницы к пред., ч', 'Δ опозданий к пред., ч'])
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 40

    def period_label(meta):
        m = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь',
             7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'}
        return f"{m.get(meta['month'], meta['month'])} {meta['year']}"

    def hrs(v): return round(v / 60, 2)
    prev = None
    for snap in history_data:
        emps = snap['employees']
        days = snap['days']
        if filter_dept:
            emps = [e for e in emps if e.get('department') == filter_dept]
            ids = {e['number'] for e in emps}
            days = [d for d in days if d['number'] in ids]
        if filter_unit:
            emps = [e for e in emps if e.get('unit') == filter_unit]
            ids = {e['number'] for e in emps}
            days = [d for d in days if d['number'] in ids]
        agg = {'emp': len(emps), **{k: 0 for k in DKEYS}}
        for x in days:
            m = x.get('m')
            if not m: continue
            for k in DKEYS:
                agg[k] += m[k]
        agg['diff'] = agg['fact'] - agg['plan']
        line = [period_label(snap['meta']), agg['emp'],
                hrs(agg['plan']), hrs(agg['fact']), hrs(agg['diff']),
                hrs(agg['late']), hrs(agg['early']), hrs(agg['late_leave']), hrs(agg['off'])]
        if prev is not None:
            line.append(round(hrs(agg['diff']) - hrs(prev['diff']), 2))
            line.append(round(hrs(agg['late']) - hrs(prev['late']), 2))
        else:
            line.append('—')
            line.append('—')
        ws.append(line)
        prev = agg
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in r:
            cell.border = border
            if cell.column > 1:
                cell.alignment = Alignment(horizontal='center')
    ws.column_dimensions['A'].width = 18
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 16
    wb.save(path)

def snapshot_path(period_key):
    return HISTORY_DIR / f'{period_key}.pkl'

def save_snapshot(employees, days, meta):
    try:
        with open(snapshot_path(meta['period_key']), 'wb') as f:
            pickle.dump({'meta': meta, 'employees': employees, 'days': days}, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception as ex:
        print(f'[save_snapshot] error: {ex}')

def load_snapshot(period_key):
    p = snapshot_path(period_key)
    if not p.exists():
        return None
    try:
        with open(p, 'rb') as f:
            return pickle.load(f)
    except Exception:
        return None

def list_history():
    keys = []
    for p in HISTORY_DIR.glob('*.pkl'):
        name = p.stem
        if re.fullmatch(r'20\d{2}-\d{2}', name):
            keys.append(name)
    keys.sort()
    return keys

def load_all_history():
    result = []
    for key in list_history():
        snap = load_snapshot(key)
        if snap:
            result.append(snap)
    return result

def save_state(employees, days, meta, filename):
    try:
        with open(STATE_FILE, 'wb') as f:
            pickle.dump({
                'employees': employees, 'days': days,
                'meta': meta, 'filename': filename
            }, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception as ex:
        print(f'[save_state] error: {ex}')

def load_state():
    if not STATE_FILE.exists():
        return None
    try:
        with open(STATE_FILE, 'rb') as f:
            return pickle.load(f)
    except Exception as ex:
        print(f'[load_state] error: {ex}')
        try:
            STATE_FILE.unlink()
        except Exception:
            pass
        return None

CSS = (
    ':root{--bg:#f2f4f7;--card:#ffffff;--ink:#1b2430;--muted:#6b7684;'
    '--accent:#2f6fed;--accent-dark:#1f56c4;--green:#1f9d55;--red:#b42318;--radius:14px}'
    '*{box-sizing:border-box}body{font-family:"Segoe UI",Arial,sans-serif;margin:0;'
    'padding:28px 32px;background:linear-gradient(180deg,#eef2f8,#f6f8fb);color:var(--ink)}'
    'h1{font-size:26px;margin:0 0 4px}h2{font-size:17px;margin:0 0 14px;color:var(--ink)}'
    '.subtitle{color:var(--muted);margin:0 0 20px;font-size:14px}'
    '.card{background:var(--card);border-radius:var(--radius);padding:20px 22px;margin:16px 0;'
    'box-shadow:0 4px 16px rgba(20,30,60,0.06);border:1px solid #e7ebf1}'
    '.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px}'
    '.stat{text-align:center;padding:6px 0}'
    '.metric{font-size:30px;font-weight:700;color:var(--accent-dark)}'
    '.stat-label{color:var(--muted);font-size:13px;margin-top:2px}'
    'table{border-collapse:collapse;width:100%;background:var(--card);font-size:13px}'
    'th,td{padding:10px 12px;border-bottom:1px solid #eef1f5;text-align:left;vertical-align:top}'
    'th{background:#f7f9fc;color:#48505c;font-weight:600;position:sticky;top:0;'
    'border-bottom:2px solid #e3e8ef}'
    'tr:hover td{background:#f8fafc}.name{font-weight:600}'
    '.wrap{overflow:auto;max-height:600px;border-radius:10px;border:1px solid #eef1f5}'
    '.bad td{background:#fdecec}.warn td{background:#fff6df}.blue td{background:#eaf1ff}'
    '.btn{background:var(--accent);color:#fff;padding:10px 18px;border:0;border-radius:9px;'
    'text-decoration:none;cursor:pointer;font-weight:600;font-size:14px;display:inline-block;transition:background .15s}'
    '.btn:hover{background:var(--accent-dark)}.btn-green{background:var(--green)}'
    '.btn-green:hover{background:#177a41}.btn-ghost{background:#eef1f6;color:var(--muted)}'
    '.btn-ghost:hover{background:#e3e7ee}'
    'input[type=file]{padding:8px;border:1px dashed #c7cede;border-radius:8px;background:#fbfcfe}'
    '.select-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}'
    '.select-row label{font-weight:600;color:var(--muted);font-size:13px}'
    'select,input[list]{padding:9px 12px;min-width:260px;border-radius:9px;border:1px solid #d7dce4;background:#fff;font-size:13px}'
    '.muted{color:var(--muted);margin:10px 0 0}'
    '.tag{display:inline-block;padding:3px 9px;border-radius:999px;font-size:12px;font-weight:600}'
    '.tag-green{background:#e3f6e9;color:#177a41}.tag-gray{background:#eceff3;color:#5b6472}'
    '.tag-red{background:#fdecec;color:#b42318}'
    '.tabs{display:flex;gap:8px;margin:18px 0 4px;border-bottom:2px solid #e3e8ef;flex-wrap:wrap}'
    '.tab{padding:11px 22px;border-radius:10px 10px 0 0;text-decoration:none;font-weight:600;'
    'font-size:15px;color:var(--muted);background:#eef1f6;border:1px solid #e3e8ef;'
    'border-bottom:none;margin-bottom:-2px}'
    '.tab:hover{color:var(--accent-dark)}.tab.active{background:#fff;color:var(--accent-dark);border-bottom:2px solid #fff}'
    '.pos{color:#177a41}.neg{color:#b42318}.strong{font-weight:700}.grey{color:#8a93a0}'
    'tr.total td{background:#f1f4f9;font-weight:700;border-top:2px solid #d7dce4}'
    '.info-th{cursor:help;position:relative}.info-icon{font-size:11px;color:#9aa4b2}'
    '.note{color:var(--muted);font-size:12px;margin:12px 0 0;line-height:1.5}'
    '.notice{background:#fff6df;border:1px solid #f0c34a;border-radius:10px;padding:12px 16px;margin:12px 0;color:#8a5a00;font-size:13px}'
    '.err{background:#fdecec;border:1px solid #f0b0b0;border-radius:10px;padding:12px 16px;margin:12px 0;color:#b42318;font-size:13px}'
    '.cols-form{display:flex;gap:14px;flex-wrap:wrap;align-items:center;padding:4px 0 0;'
    'font-size:13px;color:var(--ink)}'
    '.cols-form label{display:flex;gap:6px;align-items:center;font-weight:500}'
    '.cols-form input[type=checkbox]{margin:0;width:15px;height:15px}'
    '.pie{display:flex;gap:24px;align-items:center;flex-wrap:wrap}'
    '.pie svg{flex:0 0 auto;display:block}'
    '.pie-legend{font-size:13px;line-height:1.9}'
    '.pie-legend .dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:8px;vertical-align:middle}'
    '.pie-legend .val{margin-left:8px;color:var(--muted)}'
)

TOP = ('<!doctype html><html lang="ru"><head><meta charset="utf-8">'
       '<meta name="viewport" content="width=device-width, initial-scale=1">'
       '<title>Анализ присутствия сотрудников</title>'
       '<style>' + CSS + '</style></head><body>')
BOTTOM = '</body></html>'

STATE = {
    'employees': None, 'days': None, 'meta': None,
    'filename': None, 'status': 'idle', 'error': None
}

def load_state_to_memory():
    saved = load_state()
    if saved:
        STATE['employees'] = saved.get('employees')
        STATE['days'] = saved.get('days')
        STATE['meta'] = saved.get('meta')
        STATE['filename'] = saved.get('filename')

load_state_to_memory()

def admin_form():
    if STATE['status'] == 'processing':
        note = '<p class="muted">⏳ Файл обрабатывается, подождите несколько секунд.</p>'
    elif STATE['error']:
        note = '<div class="err">Ошибка обработки: ' + esc(STATE['error']) + '</div>'
    elif STATE['filename']:
        meta = STATE.get('meta') or {}
        per = f"{meta.get('month', '?')}.{meta.get('year', '?')}" if meta else '—'
        note = ('<p class="muted">Опубликован файл: <b>' + esc(STATE['filename']) +
                '</b> · период: <b>' + esc(per) + '</b>. '
                'Загрузите новый файл, чтобы заменить активный отчёт.</p>')
    else:
        note = ('<p class="muted">Загрузите Excel-файл отчёта о присутствии. '
                'Он станет активным для руководителей сразу после обработки.</p>')

    hist = list_history()
    hist_html = ''
    if hist:
        hist_html = ('<div class="card"><h2>Загруженные периоды (в истории)</h2>'
                     '<p class="muted">' + ', '.join(hist) + '</p></div>')

    return ('<h1>Анализ присутствия сотрудников — Админ</h1>'
            '<p class="subtitle">Загрузка файлов и публикация дашборда для руководителей</p>'
            '<div class="card"><form method="post" enctype="multipart/form-data" action="/upload">'
            '<input type="file" name="file" accept=".xlsx,.xls" required> '
            '<button class="btn">Обработать и опубликовать</button></form>'
            + note + '</div>' + hist_html)

def public_empty():
    return ('<h1>Анализ присутствия сотрудников</h1>'
            '<p class="subtitle">Дашборд для руководителей</p>'
            '<div class="card"><div class="notice">'
            'Отчёт пока не опубликован. Обратитесь к ответственному за обновление.'
            '</div></div>')

def tabs_bar(tab, dept='', unit='', cols=''):
    items = [
        ('dashboard_dept', 'Дашборд по департаментам'),
        ('dashboard_unit', 'Дашборд по подразделениям'),
        ('employees', 'Анализ по сотрудникам'),
        ('dynamics', 'Динамика по периодам'),
    ]
    q = []
    if dept: q.append('department=' + esc(dept))
    if unit: q.append('unit=' + esc(unit))
    if cols: q.append('cols=' + esc(cols))
    qs = ('&' + '&'.join(q)) if q else ''
    out = '<div class="tabs">'
    for key, label in items:
        cls = 'tab active' if key == tab else 'tab'
        out += f'<a class="{cls}" href="/?tab={key}{qs}">{label}</a>'
    out += '</div>'
    return out

def render_employees(emps, days, selected, dept_selected, unit_selected):
    selected = str(selected or '')
    dept_selected = str(dept_selected or '')
    unit_selected = str(unit_selected or '')
    es = emps
    if selected:
        es = [e for e in es if e['name'] == selected]
    if dept_selected:
        es = [e for e in es if e.get('department') == dept_selected]
    if unit_selected:
        es = [e for e in es if e.get('unit') == unit_selected]
    ids = {e['number'] for e in es}
    ds = [d for d in days if d['number'] in ids]

    title = selected if selected else 'Все сотрудники'
    if dept_selected: title += ' — ' + dept_selected
    if unit_selected: title += ' / ' + unit_selected

    filter_query = ''
    if selected: filter_query += 'employee=' + esc(selected) + '&'
    if dept_selected: filter_query += 'department=' + esc(dept_selected) + '&'
    if unit_selected: filter_query += 'unit=' + esc(unit_selected) + '&'
    filter_query = filter_query.rstrip('&')

    filters = filter_bar('employees', emps,
                         dept_sel=dept_selected,
                         unit_sel=unit_selected,
                         emp_sel=selected,
                         show_employee=True)

    export_btns = '<div class="stat"><a class="btn btn-green" href="/export/full">Export Excel (полная выгрузка)</a></div>'
    if selected or dept_selected or unit_selected:
        export_btns += ('<div class="stat"><a class="btn" href="/export/filtered?' +
                        filter_query + '">Export Excel по фильтру</a></div>')

    sr = []
    for e in es:
        s = e['summary']
        status = ('<span class="tag tag-gray">Увольнение сотрудника</span>'
                  if s['termination'] else '<span class="tag tag-green">Сформирован</span>')

        diff_min = s['fact'] - s['plan']
        if diff_min < -5:
            norm_status = '<span class="tag tag-red">Недоработка</span>'
            norm_class = 'neg strong'
        elif diff_min > 5:
            norm_status = '<span class="tag" style="background:#fff6df;color:#8a5a00">Переработка</span>'
            norm_class = 'pos strong'
        else:
            norm_status = '<span class="tag tag-green">Норма</span>'
            norm_class = 'grey'

        pct = round((diff_min / s['plan'] * 100), 1) if s['plan'] else 0.0
        if abs(pct) <= 1:
            pct_class = 'grey'
        elif abs(pct) <= 5:
            pct_class = 'pos' if pct > 0 else 'neg'
        else:
            pct_class = 'neg strong' if pct < 0 else 'pos strong'

        sr.append(
            '<tr>'
            '<td>' + esc(e['number']) + '</td>'
            '<td class="name">' + esc(e['name']) + '</td>'
            '<td>' + esc(e.get('position', '')) + '</td>'
            '<td>' + esc(e.get('department', '')) + '</td>'
            '<td>' + esc(e.get('unit', '')) + '</td>'
            '<td>' + esc(e['work_format']) + '</td>'
            '<td>' + esc(e['schedule']) + '</td>'
            '<td>' + str(s['work']) + '</td>'
            '<td>' + str(s['remote']) + '</td>'
            '<td>' + str(s['late']) + '</td>'
            '<td>' + str(s['early']) + '</td>'
            '<td>' + str(s['anomaly']) + '</td>'
            '<td>' + str(s['late_leave']) + '</td>'
            '<td>' + fh(s['above']) + '</td>'
            '<td>' + fh(s['below']) + '</td>'
            '<td class="' + norm_class + '">' + fsigned(diff_min) + '</td>'
            '<td class="' + pct_class + '">' + f'{pct:+.1f}%' + '</td>'
            '<td>' + norm_status + '</td>'
            '<td>' + status + '</td>'
            '</tr>')

    dr = []
    for d in ds:
        st = '; '.join(d['statuses']) or '—'
        bad = any(x.startswith(('Опоздание', 'Ранний уход', 'Аномально', 'Проблема', 'Нераспознанный'))
                  for x in d['statuses'])
        warn = any(x.startswith(('Поздний уход', 'Отклонение длительности')) for x in d['statuses'])
        cl = 'bad' if bad else ('warn' if warn else
                                ('blue' if d['kind'] in ('remote', 'free', 'off_schedule') else ''))
        dr.append(
            '<tr class="' + cl + '">'
            '<td>' + esc(d['number']) + '</td>'
            '<td class="name">' + esc(d['name']) + '</td>'
            '<td>' + esc(d.get('position', '')) + '</td>'
            '<td>' + esc(d.get('department', '')) + '</td>'
            '<td>' + esc(d.get('unit', '')) + '</td>'
            '<td>' + esc(d['date']) + '</td>'
            '<td>' + fh(d['rv']) + '</td>'
            '<td>' + fh(d['skud']) + '</td>'
            '<td>' + esc(d['interval']) + '</td>'
            '<td>' + esc(d['plan']) + '</td>'
            '<td>' + esc(st) + '</td>'
            '</tr>')

    thead = ('<tr><th>№</th><th>Сотрудник</th><th>Должность</th><th>Департамент</th>'
             '<th>Подразделение</th><th>Формат работы</th><th>График</th>'
             '<th>Рабочие дни</th><th>Удалённые дни</th><th>Опоздания</th><th>Ранние уходы</th>'
             '<th>Аномально ранние</th><th>Поздние уходы</th><th>СКУД выше РВ</th>'
             '<th>СКУД ниже РВ</th>'
             '<th>Отклонение, ч</th><th>Отклонение, %</th><th>Статус нормы</th>'
             '<th>Статус</th></tr>')

    return (filters +
            '<div class="card"><div class="grid">'
            '<div class="stat"><div class="metric">' + str(len(es)) + '</div>'
            '<div class="stat-label">сотрудников</div></div>'
            '<div class="stat"><div class="metric">' + str(len(ds)) + '</div>'
            '<div class="stat-label">дневных записей</div></div>'
            + export_btns + '</div></div>'
            '<div class="card"><h2>Анализ данных: ' + esc(title) + '</h2>'
            '<div class="wrap"><table>' + thead + ''.join(sr) + '</table></div></div>'
            '<div class="card"><h2>Детализация по дням</h2>'
            '<div class="wrap"><table>'
            '<tr><th>№</th><th>Сотрудник</th><th>Должность</th><th>Департамент</th>'
            '<th>Подразделение</th><th>Дата</th><th>РВ</th><th>СКУД</th>'
            '<th>Интервал</th><th>План</th><th>Статусы</th></tr>'
            + ''.join(dr) + '</table></div></div>')

def render_dashboard(emps, days, group_by, filter_dept='', filter_unit='', active_cols=None):
    rows, t = dashboard_data(emps, days, group_by, filter_dept, filter_unit)
    if active_cols is None:
        active_cols = set(ALL_COLS)

    def sgn(v):
        return 'neg' if round(v) < 0 else ('pos' if round(v) > 0 else '')

    tab_key = 'dashboard_dept' if group_by == 'department' else 'dashboard_unit'
    title = 'Дашборд по департаментам' if group_by == 'department' else 'Дашборд по подразделениям'
    subtitle = ('Департаменты' if group_by == 'department' else 'Подразделения') + \
               (f' департамента «{filter_dept}»' if filter_dept else '') + \
               (f' · {filter_unit}' if filter_unit else '') + \
               ' отсортированы по разнице план/факт: лучшие сверху'

    form = filter_bar(tab_key, emps, dept_sel=filter_dept, unit_sel=filter_unit)
    cols_form = render_cols_form(tab_key, active_cols, dept=filter_dept, unit=filter_unit)

    q = []
    if filter_dept: q.append('department=' + esc(filter_dept))
    if filter_unit: q.append('unit=' + esc(filter_unit))
    cols_str = cols_to_url(active_cols)
    if cols_str: q.append('cols=' + esc(cols_str))
    qs = ('?' + '&'.join(q)) if q else ''
    export_url = ('/export/dashboard_dept' if group_by == 'department'
                  else '/export/dashboard_unit') + qs

    cards = ('<div class="card"><div class="grid">'
             '<div class="stat"><div class="metric">' + str(len(rows)) + '</div>'
             '<div class="stat-label">строк</div></div>'
             '<div class="stat"><div class="metric">' + fm(t['plan']) + '</div>'
             '<div class="stat-label">план (РВ), ч</div></div>'
             '<div class="stat"><div class="metric">' + fm(t['fact']) + '</div>'
             '<div class="stat-label">факт, ч</div></div>'
             '<div class="stat"><div class="metric ' + sgn(t['diff']) + '">' +
             fsigned(t['diff']) + '</div><div class="stat-label">разница, ч</div></div>'
             '<div class="stat"><a class="btn btn-green" href="' + export_url +
             '">Export Excel</a></div></div></div>')

    head_cells = ['<th>' + ('Департамент' if group_by == 'department' else 'Подразделение') + '</th>']
    head_cells.append('<th>Сотрудники</th>')
    col_titles = {
        'plan': 'План (РВ), ч', 'fact': 'Факт, ч', 'diff': 'Разница, ч',
        'norm': 'Статус нормы', 'pct': 'Отклонение, %',
        'late': 'Опоздания, ч', 'early': 'Ранние уходы, ч',
        'late_leave': 'Поздние уходы, ч', 'off': 'Вне графика, ч',
        'early_arr': 'Ранний приход, ч', 'remote_rv': 'Потенц. удалённые, ч',
        'free': 'Свободный график, ч',
    }
    for k, _ in COL_TOGGLES:
        if k in active_cols:
            head_cells.append('<th>' + col_titles.get(k, k) + '</th>')
    head = '<tr>' + ''.join(head_cells) + '</tr>'

    def line(label, r, cls=''):
        cells = ['<td class="name">' + esc(label) + '</td>',
                 '<td>' + str(r['emp']) + '</td>']
        for k, _ in COL_TOGGLES:
            if k not in active_cols:
                continue
            if k == 'diff':
                cells.append('<td class="' + sgn(r['diff']) + ' strong">' + fsigned(r['diff']) + '</td>')
            elif k == 'norm':
                dm = r['diff']
                if dm < -5:
                    cells.append('<td><span class="tag tag-red">Недоработка</span></td>')
                elif dm > 5:
                    cells.append('<td><span class="tag" style="background:#fff6df;color:#8a5a00">Переработка</span></td>')
                else:
                    cells.append('<td><span class="tag tag-green">Норма</span></td>')
            elif k == 'pct':
                pct = round((r['diff'] / r['plan'] * 100), 1) if r['plan'] else 0.0
                if abs(pct) <= 1:
                    cl = 'grey'
                elif abs(pct) <= 5:
                    cl = 'pos' if pct > 0 else 'neg'
                else:
                    cl = 'neg strong' if pct < 0 else 'pos strong'
                cells.append('<td class="' + cl + '">' + f'{pct:+.1f}%' + '</td>')
            elif k in ('early_arr', 'remote_rv', 'free'):
                cells.append('<td class="grey">' + fm(r[k]) + '</td>')
            else:
                cells.append('<td>' + fm(r[k]) + '</td>')
        return '<tr class="' + cls + '">' + ''.join(cells) + '</tr>'

    body = ''.join(line(r['name'], r) for r in rows) + line('ИТОГО', t, 'total')
    note = ('<p class="note">Расчёт без допусков: любое отклонение считается с первой минуты. '
            'Время раннего прихода в факт не входит. Дни, где РВ>0, но нет ни интервала, ни СКУД, '
            'считаются потенциально удалёнными и в план/факт не входят. '
            'Часы показаны в формате ЧЧ:ММ.</p>')

    pie_parts = [
        ('Опоздания',    t.get('late', 0),  '#e57373'),
        ('Ранние уходы', t.get('early', 0), '#f0b24a'),
        ('Вне графика',  t.get('off', 0),   '#7c8cf0'),
    ]
    pie_html = render_pie('Структура отклонений', pie_parts)

    table = ('<div class="card"><h2>' + esc(title) + '</h2>'
             '<p class="muted">' + esc(subtitle) + '</p>'
             '<div class="wrap"><table>' + head + body + '</table></div>' +
             note + '</div>')
    return form + cols_form + cards + pie_html + table

def render_dynamics(filter_dept='', filter_unit=''):
    history = load_all_history()
    if not history:
        return ('<div class="card"><h2>Динамика по периодам</h2>'
                '<p class="muted">Пока нет загруженных периодов. Загрузите файлы за несколько месяцев '
                'на странице <a href="/admin">/admin</a> — они появятся здесь.</p></div>')

    all_emps = [e for s in history for e in s['employees']]
    form = filter_bar('dynamics', all_emps, dept_sel=filter_dept, unit_sel=filter_unit)

    def period_label(meta):
        m = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь',
             7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'}
        return f"{m.get(meta['month'], meta['month'])} {meta['year']}"

    rows_html = ''
    prev = None
    for snap in history:
        emps = snap['employees']
        days = snap['days']
        if filter_dept:
            emps = [e for e in emps if e.get('department') == filter_dept]
            ids = {e['number'] for e in emps}
            days = [d for d in days if d['number'] in ids]
        if filter_unit:
            emps = [e for e in emps if e.get('unit') == filter_unit]
            ids = {e['number'] for e in emps}
            days = [d for d in days if d['number'] in ids]
        agg = {'emp': len(emps), **{k: 0 for k in DKEYS}}
        for x in days:
            m = x.get('m')
            if not m: continue
            for k in DKEYS:
                agg[k] += m[k]
        agg['diff'] = agg['fact'] - agg['plan']

        def sgn(v):
            return 'neg' if round(v) < 0 else ('pos' if round(v) > 0 else '')

        diff_html = ('<span class="' + sgn(agg['diff']) + ' strong">' + fsigned(agg['diff']) + '</span>')
        delta_diff = ''
        delta_late = ''
        if prev is not None:
            dd = round((agg['diff'] - prev['diff']) / 60, 2)
            dl = round((agg['late'] - prev['late']) / 60, 2)
            cls_dd = 'neg' if dd < 0 else ('pos' if dd > 0 else '')
            cls_dl = 'neg' if dl < 0 else ('pos' if dl > 0 else '')
            delta_diff = f'<span class="{cls_dd}">{dd:+.2f} ч</span>'
            delta_late = f'<span class="{cls_dl}">{dl:+.2f} ч</span>'
        else:
            delta_diff = '—'
            delta_late = '—'

        rows_html += (
            '<tr>'
            '<td class="name">' + esc(period_label(snap['meta'])) + '</td>'
            '<td>' + str(agg['emp']) + '</td>'
            '<td>' + fm(agg['plan']) + '</td>'
            '<td>' + fm(agg['fact']) + '</td>'
            '<td>' + diff_html + '</td>'
            '<td>' + fm(agg['late']) + '</td>'
            '<td>' + fm(agg['early']) + '</td>'
            '<td>' + fm(agg['late_leave']) + '</td>'
            '<td>' + fm(agg['off']) + '</td>'
            '<td>' + delta_diff + '</td>'
            '<td>' + delta_late + '</td>'
            '</tr>')
        prev = agg

    head = ('<tr><th>Период</th><th>Сотрудники</th><th>План (РВ), ч</th><th>Факт, ч</th>'
            '<th>Разница, ч</th><th>Опоздания, ч</th><th>Ранние уходы, ч</th>'
            '<th>Поздние уходы, ч</th><th>Вне графика, ч</th>'
            '<th>Δ разницы к пред.</th><th>Δ опозданий к пред.</th></tr>')

    export_url = '/export/dynamics'
    q = []
    if filter_dept: q.append('department=' + esc(filter_dept))
    if filter_unit: q.append('unit=' + esc(filter_unit))
    if q: export_url += '?' + '&'.join(q)

    table = ('<div class="card"><h2>Динамика по периодам</h2>'
             '<p class="muted">Сравнение показателей месяц к месяцу. '
             'Δ — изменение к предыдущему периоду (зелёный — улучшение, красный — ухудшение).</p>'
             '<div class="wrap"><table>' + head + rows_html + '</table></div>'
             '<p class="note"><a class="btn btn-green" href="' + export_url +
             '">Export Excel (Динамика)</a></p></div>')
    return form + table

def build_page(tab, args):
    tab = tab or 'dashboard_dept'
    if tab not in ('dashboard_dept', 'dashboard_unit', 'employees', 'dynamics'):
        tab = 'dashboard_dept'

    if STATE['employees'] is None:
        if request.path.startswith('/admin') or request.path.startswith('/upload'):
            return TOP + admin_form() + BOTTOM
        return TOP + public_empty() + BOTTOM

    dept = args.get('department', '')
    unit = args.get('unit', '')
    active_cols = collect_cols(args)
    cols_str = cols_to_url(active_cols)

    if tab == 'dashboard_dept':
        body = render_dashboard(STATE['employees'], STATE['days'], 'department',
                                dept, unit, active_cols=active_cols)
    elif tab == 'dashboard_unit':
        body = render_dashboard(STATE['employees'], STATE['days'], 'unit',
                                dept, unit, active_cols=active_cols)
    elif tab == 'employees':
        body = render_employees(STATE['employees'], STATE['days'],
                                args.get('employee', ''), dept, unit)
    else:
        body = render_dynamics(dept, unit)

    header = admin_form() if request.path.startswith('/admin') else ''
    return TOP + header + tabs_bar(tab, dept, unit, cols_str) + body + BOTTOM

@app.route('/', methods=['GET'])
def index():
    return build_page(request.args.get('tab', 'dashboard_dept'), request.args)

@app.route('/admin', methods=['GET'])
def admin_index():
    return build_page(request.args.get('tab', 'dashboard_dept'), request.args)

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename:
        return redirect('/admin')
    if Path(f.filename).suffix.lower() not in ('.xlsx', '.xls'):
        STATE['error'] = 'Файл должен иметь расширение .xlsx или .xls'
        return redirect('/admin')

    safe_name = Path(f.filename).name
    target = UPLOAD_DIR / safe_name
    f.save(target)
    try:
        target.replace(LAST_FILE)
    except Exception:
        pass

    STATE['status'] = 'processing'
    STATE['error'] = None
    try:
        prepared = UPLOAD_DIR / 'prepared.xlsx'
        prepare_upload(LAST_FILE, prepared)
        emps, days, meta = analyze(prepared)

        STATE['employees'] = emps
        STATE['days'] = days
        STATE['meta'] = meta
        STATE['filename'] = safe_name
        save_state(emps, days, meta, safe_name)
        save_snapshot(emps, days, meta)
        for old in EXPORT_DIR.glob('*.xlsx'):
            try: old.unlink()
            except Exception: pass
    except Exception as ex:
        STATE['error'] = str(ex)
    finally:
        STATE['status'] = 'idle'
    return redirect('/admin')

@app.route('/export/full')
def export_full():
    if STATE['employees'] is None:
        return redirect('/')
    target = EXPORT_DIR / 'result_full.xlsx'
    if not target.exists():
        export_xlsx(STATE['employees'], STATE['days'], target)
    return send_file(target, as_attachment=True, download_name='Полная выгрузка.xlsx')

@app.route('/export/filtered')
def export_filtered():
    if STATE['employees'] is None:
        return redirect('/')
    selected = request.args.get('employee', '')
    dept = request.args.get('department', '')
    unit = request.args.get('unit', '')
    es = STATE['employees']
    if selected:
        es = [e for e in es if e['name'] == selected]
    if dept:
        es = [e for e in es if e.get('department') == dept]
    if unit:
        es = [e for e in es if e.get('unit') == unit]
    ids = {e['number'] for e in es}
    ds = [d for d in STATE['days'] if d['number'] in ids]
    fname = 'Моя выгрузка'
    if unit and not selected and not dept:
        fname = safe_filename(unit)
    target = EXPORT_DIR / (fname + '.xlsx')
    export_xlsx(es, ds, target)
    return send_file(target, as_attachment=True, download_name=fname + '.xlsx')

@app.route('/export/dashboard_dept')
def export_dashboard_dept():
    if STATE['employees'] is None:
        return redirect('/')
    dept = request.args.get('department', '')
    unit = request.args.get('unit', '')
    target = EXPORT_DIR / 'dashboard_dept.xlsx'
    export_dashboard_xlsx(STATE['employees'], STATE['days'], 'department',
                          dept, unit, target, 'Дашборд по департаментам')
    return send_file(target, as_attachment=True, download_name='Дашборд по департаментам.xlsx')

@app.route('/export/dashboard_unit')
def export_dashboard_unit():
    if STATE['employees'] is None:
        return redirect('/')
    dept = request.args.get('department', '')
    unit = request.args.get('unit', '')
    target = EXPORT_DIR / 'dashboard_unit.xlsx'
    export_dashboard_xlsx(STATE['employees'], STATE['days'], 'unit',
                          dept, unit, target, 'Дашборд по подразделениям')
    fname = 'Дашборд по подразделениям'
    if dept: fname += ' - ' + safe_filename(dept)
    if unit: fname += ' - ' + safe_filename(unit)
    return send_file(target, as_attachment=True, download_name=fname + '.xlsx')

@app.route('/export/dynamics')
def export_dynamics():
    history = load_all_history()
    if not history:
        return redirect('/?tab=dynamics')
    dept = request.args.get('department', '')
    unit = request.args.get('unit', '')
    target = EXPORT_DIR / 'dynamics.xlsx'
    export_dynamics_xlsx(history, target, dept, unit)
    return send_file(target, as_attachment=True, download_name='Динамика по периодам.xlsx')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True, use_reloader=False)
