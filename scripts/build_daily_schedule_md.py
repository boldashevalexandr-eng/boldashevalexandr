"""
Generates a day-by-day Markdown schedule for the same project
(Лот №2, шифр 18053-ТУ, ООО «МАГУС-СТРОЙ»).

Output : docs/График_по_дням.md

Two tables are produced:
  1. Compact day-by-day calendar grid – rows = calendar day,
     columns = stage shifts (by stage code 1..14). Active days are
     marked with the stage code; if the stage lies on the critical
     path the marker is wrapped in **bold**. Idle days show "·".
  2. A summary footer with daily head-count, active brigades and
     totals.

The data set is identical to the one used in the Excel generators
(`scripts/build_main_schedule.py` / `build_network_schedule.py`).
"""

from datetime import date, timedelta
import os

START_DATE = date(2026, 6, 27)
TOTAL_DAYS = 240
END_DATE   = START_DATE + timedelta(days=TOTAL_DAYS - 1)

# (code, name, off, dur, crew_short, headcount, is_critical)
STAGES = [
    ('1',  'Подготовительные работы',                      0,  15,
     'Бр.№1 подгот.',         10, True),
    ('2',  'Геодезическая разбивка',                       9,  10,
     'Геодез. гр.',            4, True),
    ('3',  'Земляные работы (траншея DN 50 – DN 500)',     15, 95,
     'Бр.№2 земл.',           22, True),
    ('4',  'Монтаж магистрали ПНД DN 500',                 30, 135,
     'Бр.№3 монт.',           20, True),
    ('5',  'Монтаж отводов ПНД DN 50 / DN 100',            85, 95,
     'Бр.№3а монт.',          12, True),
    ('6',  'Колодцы (ж/б сборные)',                        60, 90,
     'Бр.№4 кол.',            10, False),
    ('7',  'Запорная арматура, фасонные изделия',          120, 70,
     'Бр.№3б арм.',            6, False),
    ('8',  'Переходы ГНБ / футляры',                       70, 60,
     'Бр.№5 ГНБ',              8, False),
    ('9',  'Контроль качества (ВИК / УЗК / ЭХЗ)',          130, 60,
     'Лаб. НК',                6, False),
    ('10', 'Обратная засыпка и уплотнение',                110, 110,
     'Бр.№2 земл.',           14, True),
    ('11', 'Гидравлические испытания',                     195, 20,
     'Бр.№6 исп.',             8, True),
    ('12', 'Пусконаладочные работы',                       215, 15,
     'Бр.№6 ПНР',              6, True),
    ('13', 'Благоустройство и рекультивация',              140, 95,
     'Бр.№7 благ.',            6, False),
    ('14', 'Сдача объекта (исполнительная, КС-11)',        220, 20,
     'ПТО + ИТР',              4, True),
]

WD = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
RU_MONTHS = ['янв', 'фев', 'мар', 'апр', 'май', 'июн',
             'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']


def is_active(stage, day_off):
    _, _, off, dur, *_ = stage
    return off <= day_off <= off + dur - 1


def day_marker(stage, day_off):
    code, name, off, dur, crew, hc, crit = stage
    if not is_active(stage, day_off):
        return '·'
    return f'**{code}**' if crit else code


def fmt(dt: date) -> str:
    return f'{dt.day:02d}.{RU_MONTHS[dt.month - 1]}'


# ---------------------------------------------------------------------------
# Build markdown
# ---------------------------------------------------------------------------
lines = []
lines.append('# График производства работ по дням')
lines.append('')
lines.append('**Объект:** Магистральный трубопровод ПНД (DN 50 – DN 500). '
             'Лот №2, шифр 18053-ТУ  ')
lines.append('**Подрядчик:** ООО «МАГУС-СТРОЙ»  ')
lines.append(f'**Начало:** {START_DATE.strftime("%d.%m.%Y")}  '
             f'**Окончание:** {END_DATE.strftime("%d.%m.%Y")}  '
             f'**Срок:** {TOTAL_DAYS} календ. дней  ')
lines.append('**Обозначение:** жирный код (**N**) — этап на критическом '
             'пути; «·» — этап в этот день не ведётся.')
lines.append('')

# ---------------------------------------------------------------------------
# Stage legend
# ---------------------------------------------------------------------------
lines.append('## Условные обозначения этапов')
lines.append('')
lines.append('| Шифр | Этап | Бригада | Раб., чел. | Крит. путь |')
lines.append('|:----:|------|---------|:----------:|:----------:|')
for code, name, off, dur, crew, hc, crit in STAGES:
    star = '★' if crit else ''
    lines.append(f'| **{code}** | {name} | {crew} | {hc} | {star} |')
lines.append('')

# ---------------------------------------------------------------------------
# Day-by-day table
# ---------------------------------------------------------------------------
lines.append('## График по календарным дням')
lines.append('')

# header
header = ['День', 'Дата', 'Дн нед'] + [c for c, *_ in STAGES] + \
         ['Σ раб.', 'Активные бригады']
sep    = [':---:', ':----:', ':---:'] + [':-:'] * len(STAGES) + \
         [':---:', '------']
lines.append('| ' + ' | '.join(header) + ' |')
lines.append('|' + '|'.join(sep) + '|')

cur_month = None
for n in range(TOTAL_DAYS):
    dt = START_DATE + timedelta(days=n)

    # month separator (new month banner row)
    if dt.month != cur_month:
        cur_month = dt.month
        banner = f'**{RU_MONTHS[dt.month-1].upper()} {dt.year}**'
        # render banner spanning first three columns + dashes elsewhere
        cells = [banner, '', ''] + [''] * len(STAGES) + ['', '']
        lines.append('| ' + ' | '.join(cells) + ' |')

    cells = [str(n + 1), fmt(dt), WD[dt.weekday()]]
    active_codes = []
    head = 0
    crews = []
    for s in STAGES:
        cells.append(day_marker(s, n))
        if is_active(s, n):
            code, name, off, dur, crew, hc, crit = s
            active_codes.append(code)
            head += hc
            crews.append(crew)
    cells.append(str(head) if head else '—')
    cells.append('; '.join(crews) if crews else '—')
    lines.append('| ' + ' | '.join(cells) + ' |')

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
lines.append('')
lines.append('## Сводка')
lines.append('')

# Peak head-count
peak = 0
peak_day = 0
total_man_days = 0
for n in range(TOTAL_DAYS):
    head = sum(hc for s in STAGES if is_active(s, n)
               for code, name, off, dur, crew, hc, crit in [s])
    total_man_days += head
    if head > peak:
        peak, peak_day = head, n

peak_dt = START_DATE + timedelta(days=peak_day)
lines.append(f'- **Пиковая численность рабочих:** {peak} чел. '
             f'({peak_dt.strftime("%d.%m.%Y")}, день {peak_day + 1}).')
lines.append(f'- **Среднесуточная численность:** '
             f'{total_man_days / TOTAL_DAYS:.1f} чел.')
lines.append(f'- **Общее количество чел.‑дней:** {total_man_days}.')
lines.append(f'- **Критический путь:** 1 → 2 → 3 → 4 → 5 → 10 → 11 → '
             f'12 → 14 (резерв = 0).')
lines.append('')

out = 'docs/График_по_дням.md'
os.makedirs('docs', exist_ok=True)
with open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print(f'Saved: {out}')
print(f'Rows : {TOTAL_DAYS} day rows + month banners')
print(f'Size : {os.path.getsize(out):,} bytes')
