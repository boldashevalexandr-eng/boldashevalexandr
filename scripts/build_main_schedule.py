"""
Generates the main schedule (Основной график) in Excel for the object
"Лот №2 (оферта ООО «МАГУС-СТРОЙ», шифр 18053-ТУ)".

Input parameters
----------------
Object      : Магистральный трубопровод из труб ПНД (DN 50 ... DN 500),
              Лот №2, шифр 18053-ТУ (оферта ООО «МАГУС-СТРОЙ» от 06.04.2026).
Start date  : 27.06.2026
Duration    : 240 calendar days
End date    : 21.02.2027

Columns
-------
№ | Этап | Начало | Окончание | Дней | Бригада | Техника и материалы

Stages that lie on the critical path are marked with "★" in column «Этап».
"""

from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Project parameters
# ---------------------------------------------------------------------------
OBJECT_NAME   = ('Магистральный трубопровод из труб ПНД (DN 50 – DN 500). '
                 'Лот №2, шифр 18053-ТУ')
CONTRACTOR    = 'ООО «МАГУС-СТРОЙ»'
START_DATE    = date(2026, 6, 27)
TOTAL_DAYS    = 240
END_DATE      = START_DATE + timedelta(days=TOTAL_DAYS - 1)   # 21.02.2027


def d(offset_days: int) -> date:
    """Return project_start + offset_days (0-based)."""
    return START_DATE + timedelta(days=offset_days)


def fmt(dt: date) -> str:
    return dt.strftime('%d.%m.%Y')


# ---------------------------------------------------------------------------
# Schedule rows
# Each row : (stage, start_offset, duration_days, crew, resources, is_critical)
# Offsets measured from project start (day 0 = 27.06.2026).
# ---------------------------------------------------------------------------
STAGES = [
    # 1. Подготовительный период
    ('Подготовительные и организационно-технические работы '
     '(мобилизация, бытовой городок, ограждение, схема движения, '
     'вынос знаков, входной контроль)',
     0, 15,
     'Бригада №1 — подготовительная (ИТР 2, раб. 8)',
     'Бульдозер Б-10, экскаватор ЭО-3323, бортовой МАЗ, вагончики-бытовки, '
     'сетка СРО, биотуалеты, ГСМ, щебень ф.20–40 (подъездные пути)',
     True),

    # 2. Геодезия
    ('Геодезическая разбивка трассы, вынос осей и реперов, '
     'закрепление знаков',
     9, 10,
     'Геодезическая группа (ИТР 2, раб. 2)',
     'Тахеометр Sokkia CX-105, нивелир, GPS-приёмник, вешки, краска, арматура',
     True),

    # 3. Земляные работы
    ('Земляные работы: разработка траншеи DN 50 – DN 500, '
     'планировка дна, устройство песчаной подготовки h=100 мм',
     15, 95,
     'Бригада №2 — земляная (раб. 18), машинисты 4',
     'Экскаватор Hitachi ZX-200, Caterpillar 320, бульдозер Komatsu D65, '
     'самосвалы КамАЗ-6520 (4 ед.), виброплита Wacker, песок средней крупности, '
     'листовая стальная крепь (инв.), водоотливные насосы',
     True),

    # 4. Монтаж магистрали DN 500
    ('Монтаж магистрального трубопровода из труб ПНД SDR17 '
     'DN 500 — сварка встык, опуск в траншею, центровка',
     30, 135,
     'Бригада №3 — монтажная (раб. 16), сварщики-пластики 4',
     'Трубы ПНД SDR17 PE100 DN 500, сварочные аппараты Widos 7000 / Ritmo Basic, '
     'трубоукладчик ТГ-124, автокран КС-55713 (25 т), центраторы, '
     'строповочные захваты',
     True),

    # 5. Отводящие линии DN 50–100
    ('Монтаж отводящих трубопроводов ПНД DN 50 / DN 100 '
     '(PE-RT 1,04 МПа), врезки, компенсаторы',
     85, 95,
     'Бригада №3а — монтажная (раб. 10), сварщики-пластики 2',
     'Трубы ПНД DN 50 / DN 100 PE100, электромуфты, сёдла-врезки, '
     'аппараты электромуфтовой сварки Hürner HST 300, мини-экскаватор JCB 8025',
     True),

    # 6. Смотровые колодцы
    ('Устройство смотровых / поворотных колодцев (ж/б сборные, '
     'установка лотков, люков, лестниц)',
     60, 90,
     'Бригада №4 — колодезная (раб. 10)',
     'Кольца КС 15.9 / 20.9, плиты перекрытия ПП, люки чугунные Т/Л, '
     'автокран КС-45717 (25 т), раствор М100, гидроизоляция "Техноэласт"',
     False),

    # 7. Запорная арматура и фасонные изделия
    ('Установка запорной арматуры, фасонных изделий, '
     'задвижек, обратных клапанов, ковров',
     120, 70,
     'Бригада №3б — арматурщики (раб. 6)',
     'Задвижки стальные DN 500 / DN 100, клапаны обратные, фланцевые '
     'переходы ПНД-сталь, крепёж 5.8, ключи динамометрические, '
     'гидравлические опрессовочные насосы',
     False),

    # 8. Переходы
    ('Переходы через действующие коммуникации и автодороги '
     '(ГНБ, футляры стальные, защитные кожухи)',
     70, 60,
     'Бригада №5 — переходы / ГНБ (раб. 8)',
     'Установка ГНБ Vermeer D24x40, буровой раствор (бентонит), '
     'стальные футляры Ø 720/820, манжеты "Манжет-К", электрохим. защита',
     False),

    # 9. Контроль качества
    ('Контроль сварных стыков (100 % визуально, 10 % УЗК/рентген), '
     'электрохимзащита, антикоррозионное покрытие стыков',
     130, 60,
     'Лаборатория НК + электромонтажники (ИТР 2, раб. 4)',
     'Дефектоскоп УД2-70, рентген-аппарат РПД-200, термоусаживаемые '
     'манжеты Canusa, станция катодной защиты, анодные заземлители',
     False),

    # 10. Обратная засыпка
    ('Обратная засыпка пазух и траншеи с послойным '
     'уплотнением (Ку ≥ 0,95), восстановление профиля',
     110, 110,
     'Бригада №2 — земляная (раб. 14)',
     'Экскаватор-погрузчик JCB 3CX, виброкатки Bomag BW 120, '
     'виброплиты Wacker DPU, песок, ПГС, грунт обратной засыпки',
     True),

    # 11. Гидроиспытания
    ('Гидравлические испытания трубопровода на прочность '
     'и герметичность, составление актов',
     195, 20,
     'Бригада №6 — испытаний (ИТР 2, раб. 6)',
     'Опрессовочный агрегат АО-161, манометры кл. 0,4, заглушки '
     'DN 500 / DN 100, емкости для воды V=25 м³, насос Grundfos CR, '
     'компрессор ПКС-5,25',
     True),

    # 12. Пусконаладка
    ('Пусконаладочные работы, промывка, дезинфекция, '
     'продувка, комплексные испытания',
     215, 15,
     'Бригада №6 — ПНР (ИТР 2, раб. 4)',
     'Гипохлорит Na, хлораторная установка, комплекс КИПиА, '
     'передвижная лаборатория, компрессор, ГСМ',
     True),

    # 13. Благоустройство
    ('Благоустройство и рекультивация нарушенных земель, '
     'восстановление дорожного покрытия и газонов',
     140, 95,
     'Бригада №7 — благоустройство (раб. 6)',
     'Асфальтоукладчик Vögele Super 1300, каток ДУ-47, а/б смесь '
     'тип Б, бордюр БР 100.30.15, плодородный грунт, семена газонных трав',
     False),

    # 14. Сдача объекта
    ('Оформление исполнительной документации, приёмо-сдаточные '
     'мероприятия, сдача комиссии, подписание КС-11',
     220, 20,
     'ПТО + ИТР (инж. 3, геодез. 1)',
     'Исп. схемы, акты скрытых работ, паспорта материалов, '
     'протоколы ВИК / УЗК / гидроиспытаний, ПО AutoCAD/Credo',
     True),
]


# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = 'Основной график'

thin     = Side(style='thin', color='808080')
medium   = Side(style='medium', color='000000')
border   = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill('solid', fgColor='1F4E78')
crit_fill = PatternFill('solid', fgColor='FDE9D9')
alt_fill  = PatternFill('solid', fgColor='F2F2F2')

# Title row
ws.merge_cells('A1:G1')
c = ws['A1']
c.value = (f'ОСНОВНОЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ\n'
           f'Объект: {OBJECT_NAME}\n'
           f'Подрядчик: {CONTRACTOR}. '
           f'Начало: {fmt(START_DATE)}   Окончание: {fmt(END_DATE)}   '
           f'Срок: {TOTAL_DAYS} календ. дней   ★ — критический путь')
c.font = Font(name='Calibri', size=12, bold=True, color='1F4E78')
c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws.row_dimensions[1].height = 70

# Header row
headers = ['№', 'Этап', 'Начало', 'Окончание', 'Дней',
           'Бригада', 'Техника и материалы']
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=2, column=col, value=h)
    cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal='center', vertical='center',
                               wrap_text=True)
    cell.border = border
ws.row_dimensions[2].height = 32

# Data rows
for idx, (stage, off, dur, crew, res, crit) in enumerate(STAGES, start=1):
    r = idx + 2
    start = d(off)
    finish = d(off + dur - 1)
    name = ('★ ' if crit else '') + stage
    row = [idx, name, fmt(start), fmt(finish), dur, crew, res]
    for col, v in enumerate(row, 1):
        cell = ws.cell(row=r, column=col, value=v)
        cell.border = border
        cell.font = Font(name='Calibri', size=10,
                         bold=crit and col == 2)
        cell.alignment = Alignment(
            horizontal='center' if col in (1, 3, 4, 5) else 'left',
            vertical='center',
            wrap_text=True)
        if crit:
            cell.fill = crit_fill
        elif idx % 2 == 0:
            cell.fill = alt_fill

# Totals
total_row = len(STAGES) + 3
ws.cell(row=total_row, column=1, value='Итого').font = Font(bold=True)
ws.merge_cells(start_row=total_row, start_column=1,
               end_row=total_row, end_column=4)
c = ws.cell(row=total_row, column=1)
c.value = f'ИТОГО по объекту: с {fmt(START_DATE)} по {fmt(END_DATE)}'
c.alignment = Alignment(horizontal='right', vertical='center')
c.font = Font(bold=True)
c.border = border
ws.cell(row=total_row, column=5, value=TOTAL_DAYS).font = Font(bold=True)
ws.cell(row=total_row, column=5).alignment = Alignment(horizontal='center')
ws.cell(row=total_row, column=5).border = border
for col in (6, 7):
    ws.cell(row=total_row, column=col, value='').border = border

# Column widths
widths = {1: 5, 2: 55, 3: 13, 4: 13, 5: 7, 6: 35, 7: 60}
for col, w in widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# Freeze header
ws.freeze_panes = 'A3'
ws.print_options.gridLines = False
ws.page_setup.orientation  = ws.ORIENTATION_LANDSCAPE
ws.page_setup.paperSize    = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth   = 1
ws.page_setup.fitToHeight  = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = '1:2'

out = 'output/Основной_график_18053-ТУ_лот2.xlsx'
import os
os.makedirs('output', exist_ok=True)
wb.save(out)
print(f'Saved: {out}')
print(f'Stages: {len(STAGES)}   Start: {fmt(START_DATE)}   '
      f'End: {fmt(END_DATE)}   Duration: {TOTAL_DAYS} days')
