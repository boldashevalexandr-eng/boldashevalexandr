"""
Сетевой график (Гантт-стиль) для объекта «Магистральный трубопровод ПНД
DN 50 – DN 500», Лот №2, шифр 18053-ТУ (оферта ООО «МАГУС-СТРОЙ»).

Файл пишется в отдельный лист «Сетевой график» того же рабочего файла
`output/Основной_график_18053-ТУ_лот2.xlsx`, где уже есть лист
«Основной график» (если файла нет — создаётся заново).

Формат
------
Левая часть  : № | Шифр | Работа | Предш. | Бригада | Начало | Окончание |
              Дней | Резерв, дн.
Правая часть : календарная шкала по неделям (35 недель) с месячной
              группировкой.
Полосы       : критический путь  — красно-оранжевая «★»;
               некритические    — синяя.
"""

from datetime import date, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

# ---------------------------------------------------------------------------
OBJECT_NAME = ('Магистральный трубопровод из труб ПНД (DN 50 – DN 500). '
               'Лот №2, шифр 18053-ТУ')
CONTRACTOR  = 'ООО «МАГУС-СТРОЙ»'
START_DATE  = date(2026, 6, 27)
TOTAL_DAYS  = 240
END_DATE    = START_DATE + timedelta(days=TOTAL_DAYS - 1)     # 21.02.2027

# ---------------------------------------------------------------------------
# Работы (офсет — день от старта, 0‑based).
#   (code, name, off, dur, pred, crew, float_days, is_critical)
# Предшественники заданы в нотации СНиП:
#   "1"         — ОН (окончание‑начало) лаг 0
#   "1(НН+9)"  — НН (начало‑начало)   лаг +9
#   "5(ОО-5)"  — ОО (окончание‑оконч.) лаг -5
# ---------------------------------------------------------------------------
STAGES = [
    ('1',  'Подготовительные и орг.-техн. работы (мобилизация, бытовка, '
           'ограждение)',                                       0,  15,
     '—',                   'Бр.№1 подгот.',                    0,  True),
    ('2',  'Геодезическая разбивка трассы, вынос осей и реперов', 9, 10,
     '1(НН+9)',             'Геодез. гр.',                      5,  True),
    ('3',  'Земляные работы: разработка траншеи DN 50 – DN 500, '
           'песчаная подготовка',                               15, 95,
     '1; 2(НН+6)',          'Бр.№2 земл.',                      0,  True),
    ('4',  'Монтаж магистрали ПНД SDR17 DN 500 — сварка встык, '
           'укладка',                                           30, 135,
     '3(НН+15)',            'Бр.№3 монт.',                      0,  True),
    ('5',  'Монтаж отводов ПНД DN 50 / DN 100 (электромуфты, '
           'врезки)',                                           85, 95,
     '4(НН+55)',            'Бр.№3а монт.',                     0,  True),
    ('6',  'Устройство смотровых / поворотных колодцев (ж/б '
           'сборные)',                                          60, 90,
     '3(НН+45)',            'Бр.№4 кол.',                       90, False),
    ('7',  'Установка запорной арматуры, фасонных изделий, '
           'ковров',                                            120, 70,
     '4(НН+90)',            'Бр.№3б арм.',                      50, False),
    ('8',  'Переходы через коммуникации и автодороги '
           '(ГНБ, футляры)',                                    70, 60,
     '3(НН+55)',            'Бр.№5 ГНБ',                        110, False),
    ('9',  'Контроль сварных стыков (ВИК, УЗК/рентген), ЭХЗ',   130, 60,
     '4(НН+100)',           'Лаб. НК',                          50, False),
    ('10', 'Обратная засыпка и уплотнение траншеи '
           '(Ку ≥ 0,95)',                                       110, 110,
     '4(НН+80); 5(НН+25)',  'Бр.№2 земл.',                      0,  True),
    ('11', 'Гидравлические испытания на прочность и '
           'герметичность',                                     195, 20,
     '10(НН+85)',           'Бр.№6 исп.',                       0,  True),
    ('12', 'Пусконаладочные работы, промывка, дезинфекция',     215, 15,
     '11',                  'Бр.№6 ПНР',                        0,  True),
    ('13', 'Благоустройство и рекультивация, восстановление '
           'покрытий',                                          140, 95,
     '10(НН+30)',           'Бр.№7 благ.',                      5,  False),
    ('14', 'Оформление испол. документации, сдача комиссии, '
           'КС-11',                                             220, 20,
     '12(НН+5)',            'ПТО + ИТР',                        0,  True),
]


def d(off):       return START_DATE + timedelta(days=off)
def fmt(dt):      return dt.strftime('%d.%m.%Y')
def month_name(m):
    return ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
            'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
            'декабрь'][m - 1]


# ---------------------------------------------------------------------------
# Open / create workbook
# ---------------------------------------------------------------------------
fname = 'output/Основной_график_18053-ТУ_лот2.xlsx'
if os.path.exists(fname):
    wb = load_workbook(fname)
else:
    wb = Workbook()
    wb.active.title = 'Placeholder'

# drop existing network sheet so we re-create it cleanly
if 'Сетевой график' in wb.sheetnames:
    del wb['Сетевой график']

ws = wb.create_sheet('Сетевой график', 0)   # make it the first sheet
wb.active = 0

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
thin    = Side(style='thin',    color='9AA0A6')
hair    = Side(style='hair',    color='BFBFBF')
medium  = Side(style='medium',  color='000000')
border  = Border(left=thin, right=thin, top=thin, bottom=thin)
hborder = Border(left=hair, right=hair, top=hair, bottom=hair)

fill_title   = PatternFill('solid', fgColor='1F4E78')
fill_month   = PatternFill('solid', fgColor='2E75B6')
fill_week    = PatternFill('solid', fgColor='D9E1F2')
fill_colhdr  = PatternFill('solid', fgColor='1F4E78')
fill_alt     = PatternFill('solid', fgColor='F5F7FA')
fill_crit    = PatternFill('solid', fgColor='FDE9D9')   # row highlight
fill_bar_c   = PatternFill('solid', fgColor='C00000')   # critical bar
fill_bar_n   = PatternFill('solid', fgColor='5B9BD5')   # non‑critical bar
fill_bar_e   = PatternFill('solid', fgColor='FFF2CC')   # edge / partial week
fill_weekend = PatternFill('solid', fgColor='FAFAFA')

white_bold   = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
small_bold   = Font(name='Calibri', size=9,  bold=True, color='FFFFFF')
small        = Font(name='Calibri', size=9)
small_b      = Font(name='Calibri', size=9, bold=True)
bar_font     = Font(name='Calibri', size=9, bold=True, color='FFFFFF')

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)

# ---------------------------------------------------------------------------
# Layout dimensions
# ---------------------------------------------------------------------------
DATA_COLS = ['№', 'Шифр', 'Работа', 'Предш.', 'Бригада',
             'Начало', 'Окончание', 'Дней', 'Резерв, дн.']
DATA_N = len(DATA_COLS)                          # 9
FIRST_WEEK_COL = DATA_N + 1                      # column 10 (J)

WEEK_DAYS = 7
N_WEEKS   = -(-TOTAL_DAYS // WEEK_DAYS)          # ceil = 35

TITLE_ROW  = 1
MONTH_ROW  = 2
WEEK_ROW   = 3
HEADER_ROW = 3          # same as WEEK_ROW for data columns (merged 2‑3)
FIRST_DATA = 4

# ---------------------------------------------------------------------------
# Title
# ---------------------------------------------------------------------------
last_col = FIRST_WEEK_COL + N_WEEKS - 1
ws.merge_cells(start_row=TITLE_ROW, start_column=1,
               end_row=TITLE_ROW,   end_column=last_col)
t = ws.cell(row=TITLE_ROW, column=1)
t.value = (f'СЕТЕВОЙ (КАЛЕНДАРНО-СЕТЕВОЙ) ГРАФИК ПРОИЗВОДСТВА РАБОТ\n'
           f'Объект: {OBJECT_NAME}.   Подрядчик: {CONTRACTOR}\n'
           f'Начало: {fmt(START_DATE)}    Окончание: {fmt(END_DATE)}    '
           f'Срок: {TOTAL_DAYS} календ. дней    '
           f'★ — критический путь')
t.fill = fill_title
t.font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
t.alignment = center
ws.row_dimensions[TITLE_ROW].height = 60

# ---------------------------------------------------------------------------
# Data‑column headers (merged rows 2‑3)
# ---------------------------------------------------------------------------
for idx, name in enumerate(DATA_COLS, start=1):
    ws.merge_cells(start_row=MONTH_ROW, start_column=idx,
                   end_row=WEEK_ROW,    end_column=idx)
    c = ws.cell(row=MONTH_ROW, column=idx, value=name)
    c.font = white_bold
    c.alignment = center
    c.fill = fill_colhdr
    c.border = border
ws.row_dimensions[MONTH_ROW].height = 28
ws.row_dimensions[WEEK_ROW].height  = 24

# ---------------------------------------------------------------------------
# Week + month headers
# ---------------------------------------------------------------------------
week_starts = [START_DATE + timedelta(days=i * WEEK_DAYS)
               for i in range(N_WEEKS)]

# Week row (row 3)
for i, ws_start in enumerate(week_starts):
    col = FIRST_WEEK_COL + i
    cell = ws.cell(row=WEEK_ROW, column=col,
                   value=ws_start.strftime('%d.%m'))
    cell.font = small_b
    cell.alignment = center
    cell.fill = fill_week
    cell.border = hborder

# Month row (row 2) – merge consecutive weeks of same month
cursor = 0
while cursor < N_WEEKS:
    m = week_starts[cursor].month
    y = week_starts[cursor].year
    j = cursor
    while j + 1 < N_WEEKS and week_starts[j + 1].month == m and \
            week_starts[j + 1].year == y:
        j += 1
    c0 = FIRST_WEEK_COL + cursor
    c1 = FIRST_WEEK_COL + j
    ws.merge_cells(start_row=MONTH_ROW, start_column=c0,
                   end_row=MONTH_ROW,   end_column=c1)
    cell = ws.cell(row=MONTH_ROW, column=c0,
                   value=f'{month_name(m).upper()} {y}')
    cell.font = white_bold
    cell.fill = fill_month
    cell.alignment = center
    cell.border = border
    cursor = j + 1

# ---------------------------------------------------------------------------
# Data + bars
# ---------------------------------------------------------------------------
def week_index(day_offset):
    return day_offset // WEEK_DAYS          # 0‑based week

for idx, (code, name, off, dur, pred, crew, tf, crit) in enumerate(STAGES):
    r = FIRST_DATA + idx
    ws.row_dimensions[r].height = 36

    start   = d(off)
    finish  = d(off + dur - 1)
    data = [idx + 1, code, name, pred, crew,
            fmt(start), fmt(finish), dur,
            '0 ★' if crit else tf]

    for col, v in enumerate(data, 1):
        cell = ws.cell(row=r, column=col, value=v)
        cell.border = border
        cell.font = small_b if (crit and col == 3) else small
        cell.alignment = left if col in (3, 4, 5) else center
        if crit:
            cell.fill = fill_crit
        elif idx % 2 == 0:
            cell.fill = fill_alt

    # Draw bar across week columns
    w_start = week_index(off)
    w_end   = week_index(off + dur - 1)
    bar_fill = fill_bar_c if crit else fill_bar_n

    for w in range(N_WEEKS):
        col = FIRST_WEEK_COL + w
        cell = ws.cell(row=r, column=col)
        cell.border = hborder
        if w_start <= w <= w_end:
            cell.fill = bar_fill
            # Put marker in first bar week
            if w == w_start:
                cell.value = ('★' if crit else '') + code
                cell.font = bar_font
                cell.alignment = Alignment(horizontal='left',
                                           vertical='center')
        else:
            if idx % 2 == 0 and not crit:
                cell.fill = fill_alt
            if crit:
                # light tint for critical row outside bar
                cell.fill = PatternFill('solid', fgColor='FFF5EC')

# ---------------------------------------------------------------------------
# Column widths
# ---------------------------------------------------------------------------
widths_left = {1: 4, 2: 6, 3: 40, 4: 17, 5: 14,
               6: 11, 7: 11, 8: 6, 9: 9}
for col, w in widths_left.items():
    ws.column_dimensions[get_column_letter(col)].width = w
for w in range(N_WEEKS):
    ws.column_dimensions[get_column_letter(FIRST_WEEK_COL + w)].width = 4

# ---------------------------------------------------------------------------
# Summary / legend row
# ---------------------------------------------------------------------------
legend_row = FIRST_DATA + len(STAGES) + 1
ws.merge_cells(start_row=legend_row, start_column=1,
               end_row=legend_row,   end_column=last_col)
l = ws.cell(row=legend_row, column=1)
l.value = ('УСЛОВНЫЕ ОБОЗНАЧЕНИЯ:   '
           '█ критический путь (★) — красно-оранжевый;   '
           '█ работы с резервом — синий;   '
           'Предшественники: ОН — окончание‑начало, НН — начало‑начало, '
           'ОО — окончание‑окончание,  цифра — № работы,  (+X) — лаг, дн.')
l.font  = small
l.fill  = fill_week
l.alignment = left
ws.row_dimensions[legend_row].height = 34

# Итого
total_row = legend_row + 1
ws.merge_cells(start_row=total_row, start_column=1,
               end_row=total_row,   end_column=7)
tc = ws.cell(row=total_row, column=1,
             value=f'ИТОГО по объекту: с {fmt(START_DATE)} '
                   f'по {fmt(END_DATE)}')
tc.font = Font(bold=True)
tc.alignment = Alignment(horizontal='right', vertical='center')
tc.border = border
ws.cell(row=total_row, column=8, value=TOTAL_DAYS).font = Font(bold=True)
ws.cell(row=total_row, column=8).alignment = center
ws.cell(row=total_row, column=8).border = border
ws.cell(row=total_row, column=9, value='★').font = Font(bold=True,
                                                        color='C00000')
ws.cell(row=total_row, column=9).alignment = center
ws.cell(row=total_row, column=9).border = border

# Freeze panes so the task table stays visible when scrolling the timeline
ws.freeze_panes = ws.cell(row=FIRST_DATA, column=FIRST_WEEK_COL).coordinate

# Print setup
ws.print_options.gridLines = False
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.page_setup.paperSize   = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth  = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f'{TITLE_ROW}:{WEEK_ROW}'

os.makedirs('output', exist_ok=True)

# remove stray default sheet if it was created
if 'Placeholder' in wb.sheetnames:
    del wb['Placeholder']
if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > 1:
    del wb['Sheet']

wb.save(fname)
print(f'Saved: {fname}')
print(f'Sheets: {wb.sheetnames}')
print(f'Weeks: {N_WEEKS}   Columns: {last_col}   Rows: {total_row}')
