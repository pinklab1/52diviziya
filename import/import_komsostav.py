"""
import_komsostav.py — сверка базы с новым списком «Командный состав 52 сд»
                      и кавалерами ордена Славы

Работа в два шага, между ними — автор списка.

1. Сверка. Скрипт сопоставляет базу со списком и пишет один xlsx: все
   люди из базы и из списка, по строке на каждого, со статусом и
   решением по умолчанию. Автор проверяет оранжевые строки «ПРОВЕРИТЬ»
   (похожие написания, дубли в самой базе) и при несогласии меняет
   решение в остальных. База не меняется.

   Кроме листа «Командный состав» читаются листы «Кавалеры_II» и
   «Кавалеры_III» (с версии в-03): кавалеры — такие же люди в общем списке,
   плюс награда «Орден Славы II/III степени». Человек, который есть в
   нескольких листах (обе степени, или кавалер и командир), — одна строка
   сверки с несколькими наградами. Награды пишутся в таблицу person_awards
   (миграция 004) и только добавляются: повторный импорт их не дублирует.

   С версии 3 списка читается и лист «Рабочие_войны» — такие же люди, без
   награды. Ответы автора берутся из «Решения», а если он ответил словами
   в «Комментарии» — оттуда (author_answers.py). Коды строк листов кавалеров
   после вставок в лист находятся заново по ФИО и году.

2. Запись. Тот же скрипт с файлом, заполненным автором, и --apply:
   уточняет тех, кто есть, добавляет новых, объединяет подтверждённые
   дубли, удаляет то, что автор отметил «Удалить». Одной транзакцией.

В отличие от import_to_db.py --reset, без решения автора ничего не
стирает. Ответы заказчика с листов «Вопросы» и «Написание ФИО» самого
списка применяются до сопоставления. Сопоставление — по ФИО и году; кроме
написания из списка пробуется написание из poteri.xlsx, откуда собиралась
база. Похожие написания («Выхрыстюк / Выхристюк») сами не склеиваются.

Запуск:
  Сверка по дампу (на своём компьютере, без сервера):
    python import_komsostav.py --file data/Командный_состав_52сд-в-03.xlsx         --dump ../backups/52sd_2026-10-06.sql.gz --sverka сверка.xlsx

  Запись (на сервере; перед ней — свежий дамп базы):
    python3 import_komsostav.py --file ... --decisions сверка_от_автора.xlsx         --sverka итог.xlsx --apply

  Зависимости на сервере:
    pip install openpyxl psycopg2-binary python-dotenv --break-system-packages
"""

import argparse
import difflib
import gzip
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import author_answers

SOURCE_TAG = 'КОМСОСТАВ в-02'

# ── Подразделения: код из списка → название строки в units ─────────────────
# Названия совпадают с теми, что завёл import_from_excel.py, — так новые
# записи о службе ссылаются на уже существующие подразделения.
UNIT_NAMES = {
    '429сп':    '429-й стрелковый полк',
    '431сп':    '431-й стрелковый полк',
    '439сп':    '439-й стрелковый полк',
    '1028ап':   '1028-й артиллерийский полк',
    '52сд':     'Управление 52-й стрелковой дивизии',
    '106осанб': '106-й отдельный санитарный батальон',
    '164осапб': '164-й отдельный сапёрный батальон',
    '405оиптд': '405-й истребительно-противотанковый дивизион',
    '587обс':   '587-й отдельный батальон связи',
    '127орр':   '127-я отдельная разведрота',
    '42орхз':   '42-я рота химзащиты',
    '842двл':   '842-й дивизионный ветеринарный лазарет',
    '371пхп':   '371-й полевой хлебозавод',
    '563орс':   '563-я отдельная рота связи',
}

# Написания, которые в списке встречаются по одному-два раза. Сведены
# к основному коду по смыслу; в отчёте они помечены, чтобы проверить.
UNIT_ALIASES = {
    '405оиптад': '405оиптд',
    '405оиптдн': '405оиптд',
    '106мсб':    '106осанб',
    '106омедб':  '106осанб',
    '106омснб':  '106осанб',
    '106санб':   '106осанб',
    'управлениядивизии52сд': '52сд',
}

# Часть не установлена — записи о службе остаются без подразделения.
UNIT_UNKNOWN = {'?сп', '?полка', '???сп/ап'}

# «убит», «погиб» и т.п. уже выражены в outcome — в подробности не пишем.
PLAIN_REASONS = {'убит', 'убита', 'погиб', 'погибла', 'умер от ран', 'умерла от ран',
                 'умер от болезни', 'умерла от болезни',
                 'пропал без вести', 'пропала без вести'}

# Листы с кавалерами ордена Славы: лист → степень. Колонки — как в
# «Командном составе» (COLS). Колонка «Вопрос №» на этих листах осталась от
# шаблона и к их строкам не относится — не читаем.
GLORY_SHEETS = {'Кавалеры_II': 2, 'Кавалеры_III': 3}
# Листы с людьми без награды (с версии 3 списка): лист → префикс кода в сверке.
PEOPLE_SHEETS = {'Рабочие_войны': 'W'}
ROMAN = {1: 'I', 2: 'II', 3: 'III'}
SRC_MAIN = 'Командный состав'

COLS = {
    'A': 'num', 'B': 'last', 'C': 'first', 'D': 'patr', 'F': 'fate',
    'H': 'birth', 'I': 'year', 'K': 'rank', 'L': 'role', 'M': 'unit',
    'N': 'out_date', 'O': 'out_reason', 'P': 'extra', 'S': 'poteri_fio',
}


# ── Чтение xlsx ─────────────────────────────────────────────────────────────

def clean(v):
    if v is None:
        return None
    s = ' '.join(str(v).split())
    return s or None


def read_sheet(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return rows[0], rows[1:]


def load_list(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Командный состав']
    rows = []
    main_block = True
    for xl_row, values in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        r = {key: clean(values[ord(col) - 65]) if ord(col) - 65 < len(values) else None
             for col, key in COLS.items()}
        if not r['last']:
            continue
        # После №1934 в файле идёт дописанный блок со своей нумерацией с 1.
        # Номера в «Вопросах» относятся к основному блоку.
        if rows and r['num'] == '1':
            main_block = False
        r['xl_row'] = xl_row
        r['main'] = main_block
        r['orig_fio'] = fio(r)
        r['notes'] = []
        r['src'] = [SRC_MAIN]
        r['awards'] = []
        rows.append(r)

    for sheet, degree in GLORY_SHEETS.items():
        if sheet in wb.sheetnames:
            rows += load_extra_sheet(wb[sheet], sheet, f'G{degree}', [glory(degree, sheet)])
    for sheet, prefix in PEOPLE_SHEETS.items():
        if sheet in wb.sheetnames:
            rows += load_extra_sheet(wb[sheet], sheet, prefix, [])

    _, questions = read_sheet(wb, 'Вопросы')
    _, spellings = read_sheet(wb, 'Написание ФИО')
    return rows, questions, spellings


def load_extra_sheet(ws, sheet, prefix, awards):
    """Кавалеры ордена Славы, «Рабочие войны» — те же колонки, что у
    командного состава. Лист отформатирован на миллион строк, поэтому читаем
    до 300 пустых подряд. Строки-разделители «А», «Б» (только буква)
    пропускаем."""
    rows, empty = [], 0
    for xl_row, values in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not any(v not in (None, '') and str(v).strip() for v in values):
            empty += 1
            if empty > 300:
                break
            continue
        empty = 0
        r = {key: clean(values[ord(col) - 65]) if ord(col) - 65 < len(values) else None
             for col, key in COLS.items()}
        if not r['last'] or not (r['first'] or r['patr']):
            continue
        r['xl_row'] = f'{prefix}_{xl_row}'    # код в сверке: RG3_45 — лист и строка
        r['main'] = False                     # «Вопросы» относятся к основному листу
        r['orig_fio'] = fio(r)
        r['notes'] = []
        r['src'] = [sheet.replace('_', ' ')]
        r['awards'] = list(awards)
        rows.append(r)
    return rows


def glory(degree, sheet):
    return {'kind': 'order_of_glory', 'degree': degree, 'status': 'awarded',
            'title': f'Орден Славы {ROMAN[degree]} степени', 'source': sheet}


def list_twins(rows):
    """Один человек под двумя написаниями фамилии внутри самого списка:
    «Цветков / Цветов Иван Никифорович», оба 439 сп. Имя и отчество
    совпадают, фамилии похожи, части не противоречат друг другу. Сами не
    сливаем — спрашиваем автора, какую строку оставить."""
    pairs = []
    for i, a in enumerate(rows):
        for b in rows[i + 1:]:
            if (norm(a['first']), norm(a['patr'])) != (norm(b['first']), norm(b['patr'])):
                continue
            if norm(a['last']) == norm(b['last']):
                continue
            if ratio(letters(a['last']), letters(b['last'])) < 0.75:
                continue
            ua, ub = unit_code(a['unit'])[0], unit_code(b['unit'])[0]
            ya, yb = birth_of(a)[1], birth_of(b)[1]
            if (ua and ub and ua != ub) or (ya and yb and ya != yb):
                continue
            pairs.append((a, b))
    return pairs


def merge_sheets(rows, log):
    """Один человек в нескольких листах — одна строка: к первой (командный
    состав, потом II, потом III) дописываются награды и откуда он.

    Тот же человек — совпали ФИО и год. В листах кавалеров год часто не
    указан: тогда, кроме ФИО, должна совпасть часть (Барашкин, 1028 ап, в
    листе II с годом, в III без). Однофамильцы разных лет — разные люди."""
    seen, out = defaultdict(list), []
    for r in rows:
        key = (norm(r['last']), norm(r['first']), norm(r['patr']))
        y, unit = birth_of(r)[1], unit_code(r['unit'])[0]
        keep = None
        if not r['main']:
            for k in seen[key]:
                ky = birth_of(k)[1]
                if (y and ky and y == ky) or ((not y or not ky) and unit and unit == unit_code(k['unit'])[0]):
                    keep = k
                    break
        if keep:
            keep['awards'] += [a for a in r['awards'] if a not in keep['awards']]
            keep['src'] += [x for x in r['src'] if x not in keep['src']]
            if not birth_of(keep)[1] and y:
                keep['year'], keep['birth'] = r['year'], r['birth']
            log.append(f"{fio(r)}: строка {r['xl_row']} слита с {keep['xl_row']}")
            continue
        seen[key].append(r)
        out.append(r)
    return out


def awards_text(awards):
    return '; '.join(a['title'] + (' (представлен)' if a['status'] == 'nominated' else '')
                     for a in awards)


def fio(r):
    return ' '.join(x for x in (r['last'], r['first'], r['patr']) if x)


# ── Ответы заказчика ────────────────────────────────────────────────────────

def apply_answers(rows, questions, spellings, log):
    """Вносит ответы с листов «Вопросы» и «Написание ФИО» в rows.
    Ответ, которого скрипт не понимает, останавливает импорт: молча
    пропустить решение заказчика хуже, чем упасть."""
    by_num = {r['num']: r for r in rows if r['main']}
    drop = set()

    for q in questions:
        qnum, topic, where, who, seen, _, answer = (list(q) + [None] * 8)[:7]
        answer = clean(answer)
        if not topic:
            continue
        nums = re.findall(r'№(\d+)', str(where or ''))

        m = re.match(r'Удалить строку №(\d+)', answer or '')
        if m:
            drop.add(m.group(1))
            log.append((qnum, topic, who, answer, f'строка №{m.group(1)} не импортируется'))
            continue

        if topic == 'Номер части':
            m = re.match(r'Исправить на «(.+)»', answer or '')
            if not m:
                raise SystemExit(f'Вопрос {qnum}: непонятный ответ «{answer}»')
            r = by_num[nums[0]]
            r['unit'] = m.group(1)
            log.append((qnum, topic, who, answer, f'часть → {m.group(1)}'))
            continue

        if topic == 'Не установлено «?»':
            r = by_num[nums[0]]
            field = str(seen).split(':')[0]
            if answer == 'Записать «не установлено»':
                if field == 'Должность':
                    r['role'] = 'не установлено'
                elif field == 'Звание':
                    r['rank'] = 'не установлено'
                elif field == 'Часть':
                    r['unit'] = None
                else:
                    raise SystemExit(f'Вопрос {qnum}: неизвестное поле «{field}»')
                log.append((qnum, topic, who, answer, f'{field.lower()} → не установлено'))
            elif field == 'Должность' and answer:
                role = answer
                if r['unit'] and role.endswith(' ' + r['unit']):
                    role = role[:-len(r['unit'])].strip()
                r['role'] = role
                log.append((qnum, topic, who, answer, f'должность → {role}'))
            else:
                raise SystemExit(f'Вопрос {qnum}: непонятный ответ «{answer}»')
            continue

        # Остальное — подтверждения как есть: «разные люди», «оставить»,
        # «значок означает погиб», «считать по таблице».
        if answer and (answer.startswith(('Оставить', 'Строки', 'Значок', 'Число', 'Считать'))):
            log.append((qnum, topic, who, answer, 'без изменений'))
            continue
        raise SystemExit(f'Вопрос {qnum}: непонятный ответ «{answer}»')

    for s in spellings:
        num, new_fio, _, _, poteri, _, answer, comment = (list(s) + [None] * 8)[:8]
        answer = clean(answer)
        r = by_num.get(str(num))
        if r is None:
            raise SystemExit(f'«Написание ФИО»: нет строки №{num}')
        if not answer:
            r['notes'].append('написание ФИО без ответа — оставлено из нового списка')
            log.append((None, 'Написание ФИО', new_fio, '—', 'ответа нет, оставлено как в новом списке'))
        elif answer == 'Приписку не надо':
            pass
        elif str(comment or '').startswith('отчество'):
            if r['patr'] != answer:
                log.append((None, 'Написание ФИО', new_fio, answer, f'отчество → {answer}'))
                r['patr'] = answer
        else:
            raise SystemExit(f'«Написание ФИО», №{num}: непонятный ответ «{answer}»')

    kept = [r for r in rows if not (r['main'] and r['num'] in drop)]
    return kept


# ── Разбор значений ─────────────────────────────────────────────────────────

def norm(s):
    return ' '.join((s or '').lower().replace('ё', 'е').split())


def letters(s):
    return re.sub(r'[^а-я]', '', norm(s))


def parse_date(s):
    """«27.04.1944», «8.04.44», «13.08.42. г.», число-дата Excel, ячейка-дата
    Excel («1943-02-03 00:00:00» — так в в-03 стоят даты выбытия) → 'YYYY-MM-DD'.
    «00.08.42», «__.08.1942» — день не известен, датой не считаем."""
    if not s:
        return None
    s = s.strip()
    if re.fullmatch(r'\d{5}', s):
        return (date(1899, 12, 30) + timedelta(days=int(s))).isoformat()
    m = re.fullmatch(r'(\d{4})-(\d\d)-(\d\d)( 00:00:00)?', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        s = re.sub(r'[\s.]*г\.?$', '', s).rstrip('.')
        m = re.fullmatch(r'(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})', s)
        if not m:
            return None
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 1900
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return None


def birth_of(r):
    full = parse_date(r['birth'])
    if full:
        return full, full[:4]
    year = r['year'] if r['year'] and re.fullmatch(r'\d{4}', r['year']) else None
    return (f'{year}-01-01' if year else None), year


def fate_parts(fate):
    """«погиб (автокатастрофа)» → ('погиб', ['автокатастрофа']). Мусор в
    скобках вроде «(;bd)(» отбрасываем: пояснение — это слова."""
    fate = fate or ''
    notes = [x.strip() for x in re.findall(r'\(([^()]*)\)?', fate)
             if len(re.findall(r'[а-яё]', x.lower())) >= 3]
    head = norm(re.split(r'\(', fate)[0])
    return head, notes


def outcome_of(r):
    """Судьба из столбца «Судьба». С в-03 автор пишет точнее: «умер от ран»,
    «умер от болезни», «погибла», «погиб (автокатастрофа)», а «… (жив)» —
    значит, человек на самом деле выжил."""
    head, notes = fate_parts(r['fate'])
    text = (r['out_reason'] or '').lower()
    if 'жив' in notes or ('жив' in text and 'умер' not in text):
        return 'survived'
    if head in ('погиб', 'погибла', 'осужден и погиб'):   # norm(): ё → е
        if 'умер' in text and 'ран' in text:
            return 'died_of_wounds'
        return 'killed'
    if head.startswith(('умер', 'умерла')):
        # «умер от ран / болезни / отравления» — в базе один вариант
        # died_of_wounds («умер от ран/болезни»); что именно — в подробностях.
        if head in ('умер', 'умерла') and 'ран' not in text:
            return None
        return 'died_of_wounds'
    if head in ('пропал без вести', 'пропала без вести'):
        if 'попал в плен' in notes or 'попала в плен' in notes:
            return 'pow'
        return 'survived' if 'выжил' in text else 'missing'
    if head == 'плен':
        return 'pow_survived' if 'освобожден' in text else 'pow'
    if head == 'расстрелян':
        return 'executed'
    return None


def details_of(r):
    """Подробности судьбы: «умер от болезни», «автокатастрофа», причина
    выбытия, если она говорит больше, чем сама судьба."""
    head, notes = fate_parts(r['fate'])
    parts = [n for n in notes if n not in ('жив', 'попал в плен', 'попала в плен')]
    if head.startswith(('умер', 'умерла')) and 'ран' not in head and head not in ('умер', 'умерла'):
        parts.insert(0, head)                       # «умер от болезни»
    reason = (r['out_reason'] or '').strip().rstrip('.')
    if reason and norm(reason) not in PLAIN_REASONS and norm(reason) != head:
        # «автокатастрофа» + «погиб в автокатастрофе» — второе полнее, первое лишнее.
        parts = [x for x in parts if norm(x)[:7] not in norm(reason)]
        parts.append(reason)
    return '; '.join(parts) or None


def unit_code(raw):
    if not raw:
        return None, None
    code = re.sub(r'\s+', '', raw.lower())
    if code in UNIT_UNKNOWN:
        return None, None
    if code in UNIT_NAMES:
        return code, None
    if code in UNIT_ALIASES:
        return UNIT_ALIASES[code], f'«{raw}» считается как {UNIT_ALIASES[code]}'
    return None, f'«{raw}» не распознано — запись без подразделения'


def extra_note(r):
    """«Прочие сведения»: убираем остатки разметки исходного документа."""
    parts = []
    for p in (r['extra'] or '').split(';'):
        p = re.sub(r'^\s*_extra:\s*', '', p).strip()
        p = re.sub(r'=+.*$', '', p).strip()
        if p:
            parts.append(p)
    return '; '.join(parts) or None


# ── Загрузка базы ───────────────────────────────────────────────────────────

def load_from_dump(path):
    tables = {}
    current = None
    opener = gzip.open if str(path).endswith('.gz') else open
    with opener(path, 'rt', encoding='utf-8') as f:
        for line in f:
            if current is None:
                m = re.match(r'COPY public\.(\w+) \((.*)\) FROM stdin;', line)
                if m and m.group(1) in ('persons', 'service_records', 'units', 'divisions',
                                        'person_awards'):
                    current = m.group(1)
                    cols = m.group(2).split(', ')
                    tables[current] = []
                continue
            if line.startswith('\\.'):
                current = None
                continue
            vals = [None if v == '\\N' else v for v in line.rstrip('\n').split('\t')]
            tables[current].append(dict(zip(cols, vals)))
    for p in tables['persons']:
        p['id'] = int(p['id'])
        p['sources'] = json.loads(p['sources']) if p['sources'] else []
    for s in tables['service_records']:
        s['id'] = int(s['id'])
        s['person_id'] = int(s['person_id'])
        s['unit_id'] = int(s['unit_id']) if s['unit_id'] else None
    units = {u['name']: int(u['id']) for u in tables['units']}
    division_id = int(tables['divisions'][0]['id'])
    # В дампах до миграции 004 таблицы наград нет — значит, наград нет.
    attach_awards(tables['persons'], tables.get('person_awards', []))
    return tables['persons'], tables['service_records'], units, division_id


def award_key(a):
    return (a['kind'], int(a['degree']) if a.get('degree') else None, a['status'])


def attach_awards(persons, awards):
    by_person = defaultdict(set)
    for a in awards:
        by_person[int(a['person_id'])].add(award_key(a))
    for p in persons:
        p['award_keys'] = by_person[p['id']]


def get_db():
    import psycopg2
    from dotenv import load_dotenv
    load_dotenv()
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


def load_from_db(cur):
    def fetch(sql):
        cur.execute(sql)
        names = [d[0] for d in cur.description]
        return [dict(zip(names, row)) for row in cur.fetchall()]

    persons = fetch("""SELECT id, last_name, first_name, patronymic, birth_date,
                              death_date, outcome, outcome_details, sources
                         FROM persons ORDER BY id""")
    for p in persons:
        for k in ('birth_date', 'death_date'):
            p[k] = p[k].isoformat() if p[k] else None
        p['sources'] = p['sources'] or []
    records = fetch("""SELECT id, person_id, unit_id, rank, role, date_to,
                              end_reason, note FROM service_records ORDER BY id""")
    for s in records:
        s['date_to'] = s['date_to'].isoformat() if s['date_to'] else None
    units = {u['name']: u['id'] for u in fetch("SELECT id, name FROM units")}
    division_id = fetch("SELECT id FROM divisions ORDER BY id LIMIT 1")[0]['id']
    if not fetch("SELECT 1 FROM information_schema.tables WHERE table_name = 'person_awards'"):
        sys.exit('В базе нет таблицы person_awards — сначала миграция 004_person_awards.sql')
    attach_awards(persons, fetch("SELECT person_id, kind, degree, status FROM person_awards"))
    return persons, records, units, division_id


# ── Сопоставление ───────────────────────────────────────────────────────────

def name_variants(r):
    """Написания, под которыми человек может лежать в базе."""
    out = [(r['last'], r['first'], r['patr'])]
    for s in [r['orig_fio']] + (r['poteri_fio'] or '').split(';'):
        w = (s or '').split()
        if len(w) >= 2:
            out.append((w[0], w[1], w[2] if len(w) > 2 else None))
    seen, uniq = set(), []
    for v in out:
        key = tuple(norm(x) for x in v)
        if key not in seen:
            seen.add(key)
            uniq.append(key)
    return uniq


def joined_name(last, first, patr):
    return letters((last or '') + (first or '') + (patr or ''))


def ratio(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def looks_same(a, b):
    """Похожие написания одного человека: «Рассоха / Россоха», «Ратин /
    Гатин», «АрхипенкоТихон Алексеевич1902 / Архипенко Тихон Алексеевич».
    Но не «Крылов / Калобухов Александр Григорьевич»: похожи должны быть
    именно фамилии. a и b — (фамилия, имя, отчество)."""
    ja, jb = joined_name(*a), joined_name(*b)
    if min(len(ja), len(jb)) >= 10 and (ja.startswith(jb) or jb.startswith(ja)):
        return True
    return (ratio(letters(a[0]), letters(b[0])) >= 0.75
            and norm(a[1])[:1] == norm(b[1])[:1]
            and ratio(ja, jb) >= 0.8)


def p_name(p):
    return (p['last_name'], p['first_name'], p['patronymic'])


def r_name(r):
    return (r['last'], r['first'], r['patr'])


def match(rows, persons):
    by_fio = defaultdict(list)
    for p in persons:
        by_fio[tuple(norm(x) for x in p_name(p))].append(p)
    year = lambda p: (p['birth_date'] or '')[:4] or None

    claimed = {}           # person id → строка списка
    result = {}            # xl_row → ('match'|'new'|'disputed', [кандидаты], почему)

    # Строки из ответов автора уже знают своего человека (force_pid) или
    # строку списка, к чьему человеку относятся (same_as).
    by_id = {p['id']: p for p in persons}
    for r in rows:
        if r.get('force_pid'):
            claimed.setdefault(r['force_pid'], r)
            result[r['xl_row']] = ('match', [by_id[r['force_pid']]], None)
    forced = [r for r in rows if r.get('force_pid') or r.get('same_as')]
    rows = [r for r in rows if r not in forced]

    def candidates(r):
        found = {}
        for v in name_variants(r):
            for p in by_fio.get(v, []):
                found[p['id']] = p
        if not found:
            # Слитые при старом разборе имена: «АрхипенкоТихон» «Алексеевич1902».
            for v in name_variants(r):
                key = joined_name(*v)
                if len(key) < 10:
                    continue
                for p in persons:
                    if joined_name(*p_name(p)).startswith(key):
                        found[p['id']] = p
        return list(found.values())

    # Проход 1: совпали ФИО и год. Проход 2: ФИО совпали, год другой или
    # пустой — только среди тех, кого не забрал проход 1 (однофамильцы
    # Сергеевы 1919 и 1910 не должны сесть на одну запись).
    pending = []
    for r in rows:
        _, y = birth_of(r)
        cands = candidates(r)
        exact = [p for p in cands if year(p) == y and y]
        if len(exact) == 1 and exact[0]['id'] not in claimed:
            claimed[exact[0]['id']] = r
            result[r['xl_row']] = ('match', exact, None)
        elif len(exact) > 1:
            result[r['xl_row']] = ('disputed', exact, 'в базе несколько человек с таким ФИО и годом')
        else:
            pending.append((r, y, cands))

    unmatched = []
    for r, y, cands in pending:
        free = [p for p in cands if p['id'] not in claimed]
        # Тёзки разных лет — разные люди: так их помечает автор списка.
        # Тем же человеком считаем, только если год не указан с одной стороны.
        namesakes = [p for p in free if year(p) and y and year(p) != y]
        free = [p for p in free if p not in namesakes]
        if len(free) == 1:
            claimed[free[0]['id']] = r
            result[r['xl_row']] = ('match', free, None)
            continue
        if len(free) > 1:
            result[r['xl_row']] = ('disputed', free, 'в базе несколько человек с таким ФИО')
            continue
        if namesakes:
            # По правилу автора тёзки разных лет — разные люди, но год часто
            # описка (та же часть и должность, разница в год). Спрашиваем,
            # а по умолчанию ставим «Новый человек».
            result[r['xl_row']] = ('disputed', namesakes,
                                   f'ФИО совпало, а год разный: на сайте '
                                   f'{"/".join(year(p) for p in namesakes)}, в списке {y}. '
                                   f'По умолчанию — разные люди; если это тот же человек, '
                                   f'выберите «Тот же человек»')
            continue
        if cands:
            result[r['xl_row']] = ('disputed', cands, 'в базе такой человек уже сопоставлен с другой строкой списка')
            continue
        unmatched.append((r, y))

    # Проход 3: похожие написания — только как вопрос, без автоматического
    # слияния. Идёт последним: иначе «Сенюгин» успевает занять запись
    # «Синюгина» раньше, чем до неё дойдёт сам Синюгин из списка.
    pool = [(p, joined_name(*p_name(p))) for p in persons]
    for r, y in unmatched:
        key = joined_name(*r_name(r))
        best, best_ratio = None, 0
        for p, joined in pool:
            if p['id'] in claimed or not joined:
                continue
            py = year(p)
            close = (not (py and y and py != y) and joined[0] == key[:1]
                     and ratio(key, joined) >= 0.85)
            if close or looks_same(r_name(r), p_name(p)):
                rt = ratio(key, joined)
                if rt > best_ratio:
                    best, best_ratio = p, rt
        if best:
            result[r['xl_row']] = ('disputed', [best], 'похожее написание')
        else:
            result[r['xl_row']] = ('new', [], None)

    for r in forced:
        if r.get('same_as'):
            kind, cands, _ = result[r['same_as']]
            if kind != 'match':
                raise SystemExit(f"Ответ автора к R{r['same_as']}: строка не сопоставлена с сайтом")
            result[r['xl_row']] = ('match', cands, None)

    return result, claimed


# ── План изменений ──────────────────────────────────────────────────────────

def fix_same_person(r, p, fix):
    """Поправки автора к «тому же человеку»: год, написание, «как на сайте»."""
    if fix.get('base_name'):
        r['last'], r['first'], r['patr'] = p['last_name'], p['first_name'], p['patronymic']
    if fix.get('base_year') and year_of(p):
        fix = dict(fix, year=year_of(p))
    if fix.get('year'):
        if (birth_of(r)[0] or '')[:4] != fix['year']:
            r['birth'] = None
        r['year'] = fix['year']
    spelling = fix.get('spelling', '').split()
    if len(spelling) >= 2:
        r['last'], r['first'] = spelling[0], spelling[1]
        r['patr'] = spelling[2] if len(spelling) > 2 else r['patr']
    elif spelling:
        # одно слово — заменяет самую похожую часть ФИО
        part = max(('last', 'first', 'patr'),
                   key=lambda f: ratio(letters(r[f]), letters(spelling[0])))
        r[part] = spelling[0]


def build_plan(rows, persons, records, units, decisions, fixes=None, comments=None):
    """План — по пункту на каждого человека из базы и из списка. Решение
    автора (decisions) уже учтено в сопоставлении; что из плана писать
    в базу, решает apply_plan по тому же решению. fixes — поправки из
    комментариев автора, comments — сами комментарии (для файла сверки)."""
    fixes, comments = fixes or {}, comments or {}
    result, claimed = match(rows, persons)
    by_id = {p['id']: p for p in persons}

    # «Тот же человек» / «Новый человек» на спорных.
    used = set()
    for r in rows:
        kind, cands, why = result[r['xl_row']]
        if kind != 'disputed':
            continue
        for p in cands:
            key = f"R{r['xl_row']}-P{p['id']}"
            verdict = decisions.get(key)
            if verdict == 'Тот же человек' and p['id'] not in claimed:
                claimed[p['id']] = r
                fix_same_person(r, p, fixes.get(key, {}))
                result[r['xl_row']] = ('match', [p], why)
                # Код остаётся «R…-P…»: при следующей сверке решение
                # «Тот же человек» найдёт ту же пару.
                r['plan_key'] = key
                used.add(key)
                break
            if verdict == 'Новый человек':
                result[r['xl_row']] = ('new', [], why)
                r['plan_key'] = key
                used.add(key)
                break

    recs_by_person = defaultdict(list)
    for s in records:
        recs_by_person[s['person_id']].append(s)

    plan = {'items': [], 'unit_notes': Counter(), 'missing_units': set(),
            'recs_by_person': recs_by_person}
    disputed_ids = set()

    # «В списке дважды»: если на сайте есть только та строка пары, которую
    # автор не оставил, оставленная строка уточняет этого человека, а не
    # заводит нового (Беликов на сайте, автор оставил «Беляков»).
    twins_in_list = list_twins([r for r in rows if not (r.get('force_pid') or r.get('same_as'))])
    for a, b in twins_in_list:
        verdict = decisions.get(f"R{a['xl_row']}-R{b['xl_row']}")
        drop, keep = {'Оставить первую': (b, a), 'Оставить вторую': (a, b)}.get(verdict, (None, None))
        if drop and result[drop['xl_row']][0] == 'match' and result[keep['xl_row']][0] == 'new':
            p = result[drop['xl_row']][1][0]
            result[keep['xl_row']] = ('match', [p], f'в списке дважды: оставлено «{fio(keep)}»')

    for r in rows:
        kind, cands, why = result[r['xl_row']]
        code, unit_note = unit_code(r['unit'])
        if unit_note:
            plan['unit_notes'][unit_note] += 1
        unit_name = UNIT_NAMES.get(code)
        if unit_name and unit_name not in units:
            plan['missing_units'].add(unit_name)
        birth, _ = birth_of(r)
        outcome = outcome_of(r)
        died = outcome in ('killed', 'died_of_wounds', 'executed')
        out_date = parse_date(r['out_date'])
        notes = [x for x in [extra_note(r)] if x]
        if r['out_date'] and not out_date:
            notes.append(f'дата выбытия: {r["out_date"]}')
        record = {
            'unit': unit_name, 'rank': r['rank'], 'role': r['role'],
            'date_to': out_date, 'end_reason': r['out_reason'],
            'note': '; '.join(notes) or None,
        }
        person = {
            'last_name': r['last'], 'first_name': r['first'], 'patronymic': r['patr'],
            'birth_date': birth, 'outcome': outcome, 'outcome_details': details_of(r),
            'death_date': out_date if died else None,
        }
        base = {'row': r, 'record': record, 'person_new': person, 'why': why,
                'awards': r['awards']}

        if kind == 'disputed':
            for p in cands:
                disputed_ids.add(p['id'])
                key = f"R{r['xl_row']}-P{p['id']}"
                default = 'Новый человек' if (why or '').startswith('ФИО совпало, а год разный') else ''
                plan['items'].append(dict(base, kind='disputed', key=key, person=p,
                                          decision=decisions.get(key, default)))
            continue
        key = r.get('plan_key') or f"R{r['xl_row']}"
        if kind == 'new':
            plan['items'].append(dict(base, kind='new', key=key, person=None,
                                      decision=decisions.get(key, 'Добавить')))
            continue

        p = cands[0]
        changes = {}
        for field in ('last_name', 'first_name', 'patronymic', 'outcome',
                      'outcome_details', 'death_date'):
            new = person[field]
            if new and new != p[field]:
                changes[field] = (p[field], new)
        # «погиб» в списке — общее слово; «умер от ран» в базе точнее.
        if changes.get('outcome') == ('died_of_wounds', 'killed'):
            del changes['outcome']
        # Год из списка не затирает точную дату рождения того же года.
        old_b = p['birth_date']
        if birth and old_b != birth and not (birth.endswith('-01-01') and old_b
                                              and old_b[:4] == birth[:4]):
            changes['birth_date'] = (old_b, birth)

        op = plan_record(p['id'], recs_by_person[p['id']], record, units)
        new_awards = [a for a in r['awards'] if award_key(a) not in p['award_keys']]
        decision = decisions.get(key, 'Обновить')
        if decision == 'Добавить':      # строка была новой, теперь нашёлся человек на сайте
            decision = 'Обновить'
        plan['items'].append(dict(base, kind='match', key=key, person=p, changes=changes,
                                  op=op, new_awards=new_awards, decision=decision))

    # Кто есть только в базе. Среди них — старые дубли: тот же человек под
    # другим написанием уже сопоставлен со строкой списка.
    matched = [by_id[pid] for pid in claimed]
    for p in persons:
        if p['id'] in claimed or p['id'] in disputed_ids:
            continue
        twin = next((m for m in matched if looks_same(p_name(p), p_name(m))
                     and not (year_of(p) and year_of(m) and year_of(p) != year_of(m))), None)
        if twin:
            key = f"P{p['id']}-P{twin['id']}"
            # Двойник нашёлся только теперь (его пару дали ответы автора), а
            # саму запись автор уже велел удалить — решение то же.
            old = 'Удалить' if decisions.get(f"P{p['id']}") == 'Удалить' else ''
            if old:
                used.add(f"P{p['id']}")
            plan['items'].append({'kind': 'twin', 'key': key, 'person': p, 'twin': twin,
                                  'row': claimed[twin['id']],
                                  'decision': decisions.get(key, old)})
        else:
            key = f"P{p['id']}"
            plan['items'].append({'kind': 'only_db', 'key': key, 'person': p, 'row': None,
                                  'decision': decisions.get(key, 'Оставить')})

    for a, b in twins_in_list:
        key = f"R{a['xl_row']}-R{b['xl_row']}"
        plan['items'].append({'kind': 'list_twin', 'key': key, 'person': None, 'row': a,
                              'other': b, 'decision': decisions.get(key, '')})

    for it in plan['items']:
        it['comment'] = comments.get(it['key'], '')

    unused = set(decisions) - {i['key'] for i in plan['items']} - used
    if unused:
        print(f'⚠ Решений, которым нет места в плане: {len(unused)} — база или список '
              f'изменились после сверки. Примеры: {sorted(unused)[:5]}')
    return plan


def stems(s):
    return {w[:4] for w in re.findall(r'[а-я0-9]+', norm(s)) if len(w) > 2}


def similar_roles(a, b):
    """«командир взвода» ≈ «командир стрелкового взвода», но не
    «командир минометного батальона» ≈ «ПНШ-2»."""
    sa, sb = stems(a), stems(b)
    if not sa or not sb:
        return False
    return sa <= sb or sb <= sa or len(sa & sb) / len(sa | sb) >= 0.4


def plan_record(person_id, existing, record, units):
    """Как лечь записи о службе из списка на записи человека в базе.

    Уточняем запись той же части с похожей должностью. Запись без части
    (обрывки старого разбора вроде «комвзв10282») заменяем, если в этой
    части у человека ничего нет. Иначе добавляем новую: другая должность
    или другая часть — это другой этап службы или расхождение источников,
    и прежнюю запись терять нельзя."""
    unit_id = units.get(record['unit'])
    # Часть в списке не указана — сверяем со всеми записями человека.
    same_unit = [s for s in existing if s['unit_id'] == unit_id or not unit_id]
    no_unit = [s for s in existing if s['unit_id'] is None]
    target = next((s for s in same_unit if norm(s['role']) == norm(record['role'])), None)
    if target is None and record['role']:
        similar = [s for s in same_unit if s['role'] and similar_roles(s['role'], record['role'])]
        target = similar[0] if similar else None
    if target is None:
        # Пустая должность или «?» — уточнять нечего, записываем поверх.
        target = next((s for s in same_unit if not stems(s['role'])), None)
    if target is None and record['role'] == 'не установлено' and same_unit:
        # Неизвестная должность не повод заводить ещё одну запись.
        target = same_unit[0]
    if target is None and not same_unit and no_unit:
        target = no_unit[0]
    if target is None:
        return ('insert', existing, record)
    diff = {}
    if record['unit'] and target['unit_id'] is None:
        diff['unit'] = (None, record['unit'])
    for field in ('rank', 'role', 'date_to', 'end_reason', 'note'):
        new = record[field]
        # «не установлено» — ответ на пустую клетку списка; то, что
        # про человека уже известно в базе, им не затираем.
        if new == 'не установлено' and target[field]:
            continue
        if new and norm(new) != norm(target[field]):
            diff[field] = (target[field], new)
    return ('update', target, diff) if diff else ('same', target, {})



# ── Файл сверки для автора ──────────────────────────────────────────────────

FIELD_RU = {
    'last_name': 'фамилия', 'first_name': 'имя', 'patronymic': 'отчество',
    'birth_date': 'дата рождения', 'outcome': 'судьба', 'outcome_details': 'подробности',
    'death_date': 'дата гибели', 'rank': 'звание', 'role': 'должность',
    'date_to': 'дата выбытия', 'end_reason': 'причина выбытия', 'note': 'примечание',
    'unit': 'часть',
}
OUTCOME_RU = {
    'killed': 'погиб', 'died_of_wounds': 'умер от ран', 'missing': 'пропал без вести',
    'pow': 'плен', 'pow_survived': 'плен, выжил', 'executed': 'расстрелян',
    'survived': 'выжил', 'unknown': 'неизвестно',
}

# Статус, цвет строки и варианты решения; первый вариант — по умолчанию.
KINDS = {
    'same':     ('Есть в базе, без изменений',          'FFFFFF', ['Обновить', 'Не обновлять']),
    'match':    ('Есть в базе, данные уточнятся',       'DDEBF7', ['Обновить', 'Не обновлять']),
    'new':      ('Новый — добавить в базу',             'E2EFDA', ['Добавить', 'Не добавлять']),
    'disputed': ('ПРОВЕРИТЬ: похож на человека в базе', 'FCE4D6',
                 ['Тот же человек', 'Новый человек', 'Не добавлять']),
    'twin':     ('ПРОВЕРИТЬ: в базе дважды',            'FCE4D6', ['Объединить', 'Разные люди', 'Удалить']),
    'only_db':  ('Только в базе, в новом списке нет',   'EDEDED', ['Оставить', 'Удалить']),
    'list_twin': ('ПРОВЕРИТЬ: в списке дважды?',        'FCE4D6',
                  ['Оставить первую', 'Оставить вторую', 'Разные люди']),
}


def ru(field, v):
    if v is None:
        return '—'
    if field == 'outcome':
        return OUTCOME_RU.get(v, v)
    m = re.fullmatch(r'(\d{4})-(\d\d)-(\d\d)', str(v))
    if m:
        y, mo, d = m.groups()
        return y if field == 'birth_date' and (mo, d) == ('01', '01') else f'{d}.{mo}.{y}'
    return str(v)


def p_fio(p):
    return ' '.join(x for x in p_name(p) if x)


def year_of(p):
    return (p['birth_date'] or '')[:4] or ''


def kind_of(item):
    if (item['kind'] == 'match' and not item['changes'] and item['op'][0] == 'same'
            and not item['new_awards']):
        return 'same'
    return item['kind']


def what_changes(item):
    parts = [f"{FIELD_RU[f]}: {ru(f, old)} → {ru(f, new)}"
             for f, (old, new) in item['changes'].items()]
    op, target, diff = item['op']
    if op == 'insert':
        rec = diff
        parts.append(f"добавится служба: {rec['role'] or '—'}, {rec['unit'] or 'без части'}"
                     + (' (прежняя остаётся)' if target else ''))
    elif op == 'update':
        parts += [f"{FIELD_RU[f]}: {ru(f, old)} → {ru(f, new)}" for f, (old, new) in diff.items()]
    if item['new_awards']:
        parts.append('добавится награда: ' + awards_text(item['new_awards']))
    return '; '.join(parts)


def write_sverka(plan, units, path):
    unit_names = {uid: name for name, uid in units.items()}
    recs = plan['recs_by_person']

    def service(p):
        return '; '.join(f"{s['role'] or '—'} ({unit_names.get(s['unit_id'], 'без части')})"
                         for s in recs[p['id']]) if p else ''

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Сверка'
    header = ['Код', 'Статус', 'Откуда', 'ФИО в новом списке', 'Год', 'Звание', 'Должность',
              'Часть', 'Судьба', 'Награды', 'ФИО в базе', 'Год в базе', 'Служба в базе',
              'Что изменится / пояснение', 'Решение', 'Комментарий']
    decision_col = chr(65 + header.index('Решение'))
    ws.append(header)

    lines = []
    for it in plan['items']:
        k = kind_of(it)
        r, p = it['row'], it['person']
        if k == 'twin':
            r = None   # строка списка относится к «двойнику», а не к этой записи
        src = ', '.join(r['src']) if r else 'только на сайте'
        list_part = ([fio(r), birth_of(r)[1] or '', r['rank'], r['role'], r['unit'],
                      r['fate'] or '', awards_text(r['awards'])] if r else [''] * 7)
        if k in ('same', 'match'):
            note = what_changes(it)
        elif k == 'disputed':
            note = it['why']
        elif k == 'new':
            note = '; '.join(x for x in [it['why'],
                                         it['awards'] and 'награда: ' + awards_text(it['awards'])] if x)
        elif k == 'twin':
            t = it['twin']
            note = (f"похоже, это тот же человек, что «{p_fio(t)}» {year_of(t)} (код P{t['id']}), "
                    f"который есть в новом списке. «Объединить» — служба переносится к нему, "
                    f"эта запись удаляется")
        elif k == 'only_db':
            note = 'источник: ' + (', '.join(p['sources']) or 'не указан')
        elif k == 'list_twin':
            b = it['other']
            note = (f"в списке есть ещё «{fio(b)}» {birth_of(b)[1] or 'год не указан'}, {b['unit'] or 'без части'} "
                    f"(код R{b['xl_row']}) — то же имя и отчество, похожая фамилия. Если это один человек, "
                    f"«Оставить первую» — останется эта строка («{fio(r)}»), «Оставить вторую» — «{fio(b)}»")
        else:
            note = ''
        sort = norm(fio(it['row']) if it['row'] else p_fio(p))
        lines.append((sort, k, [it['key'], KINDS[k][0], src] + list_part
                      + [p_fio(p) if p else '', year_of(p) if p else '', service(p),
                         note, it['decision'], it.get('comment') or '']))

    lines.sort(key=lambda x: (x[0], x[2][0]))
    validations = {}
    for n, (_, k, values) in enumerate(lines, 2):
        ws.append(values)
        fill = PatternFill('solid', fgColor=KINDS[k][1])
        for c in ws[n]:
            c.fill = fill
        ws[f'{decision_col}{n}'].fill = PatternFill('solid', fgColor='FFF2CC')
        if k not in validations:
            dv = DataValidation(type='list', formula1='"' + ','.join(KINDS[k][2]) + '"',
                                allow_blank=True)
            ws.add_data_validation(dv)
            validations[k] = dv
        validations[k].add(f'{decision_col}{n}')

    for c in ws[1]:
        c.font = Font(bold=True)
    widths = [11, 34, 18, 32, 6, 20, 40, 11, 14, 26, 32, 8, 40, 60, 16, 30]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = ws.dimensions

    counts = Counter(k for _, k, _ in lines)
    info = wb.create_sheet('Как заполнять', 0)
    text = [
        ('Сверка базы сайта 52sd.ru с новым списком «Командный состав 52 сд» и кавалерами ордена Славы', ''),
        ('', ''),
        ('На листе «Сверка» — все люди: и те, кто уже есть на сайте, и те, кто есть в новом списке.', ''),
        ('Строки идут по алфавиту, поэтому похожие записи стоят рядом. Столбец «Откуда» — из какого листа человек; по нему удобно фильтровать.', ''),
        ('', ''),
        ('Что сделать', ''),
        ('1. Оранжевые строки «ПРОВЕРИТЬ» — выбрать решение в жёлтом столбце «Решение». Без ответа они не попадут на сайт.', ''),
        ('2. Остальные строки уже заполнены решением по умолчанию. Менять только там, где не согласны.', ''),
        ('3. Если что-то не так с самими данными — написать в столбце «Комментарий».', ''),
        ('4. Столбец «Код» не трогать: по нему решения переносятся в базу.', ''),
        ('', ''),
        ('Статус', 'Строк'),
    ] + [(f'{KINDS[k][0]}  —  {", ".join(KINDS[k][2])}', counts.get(k, 0)) for k in KINDS]
    for row in text:
        info.append(row)
    info['A1'].font = Font(bold=True, size=13)
    info['A6'].font = Font(bold=True)
    info['A12'].font = Font(bold=True)
    for n, k in enumerate(KINDS, 13):
        info[f'A{n}'].fill = PatternFill('solid', fgColor=KINDS[k][1])
    info.column_dimensions['A'].width = 110
    info.column_dimensions['B'].width = 8

    wb.save(path)
    return counts


def kind_of_key(key, status):
    """Тип строки сверки по коду: R…-P… спорный, P…-P… дубль в базе,
    R…-R… дубль в списке, P… только в базе. Код надёжнее статуса: пара,
    которую автор признал «тем же человеком», в новой сверке стоит как
    «данные уточнятся», а код у неё прежний."""
    if re.fullmatch(r'R[^-]+-P\d+', key):
        return 'disputed'
    if re.fullmatch(r'P\d+-P\d+', key):
        return 'twin'
    if re.fullmatch(r'R[^-]+-R[^-]+', key):
        return 'list_twin'
    if re.fullmatch(r'P\d+', key):
        return 'only_db'
    return 'new' if str(status or '').startswith('Новый') else 'match'


def read_decisions(path, ranks=()):
    """Решения автора: жёлтый столбец «Решение», а если автор ответил
    словами в «Комментарии» — то, что сказано там (author_answers.py).
    → решения, поправки, комментарии и ФИО/год строки списка по коду."""
    empty = ({}, {}, {}, {})
    if not path:
        return empty
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    decisions, fixes, comments, info, errors = {}, {}, {}, {}, []
    rows = list(wb['Сверка'].iter_rows(values_only=True))
    head = rows[0]
    col, com, st = head.index('Решение'), head.index('Комментарий'), head.index('Статус')
    name, year = head.index('ФИО в новом списке'), head.index('Год')
    for row in rows[1:]:
        key = row[0]
        if not key:
            continue
        info[key] = (clean(row[name]), clean(row[year]))
        decision = (clean(row[col]) or '').rstrip(':') or None
        comment = clean(row[com])
        if comment:
            comments[key] = comment
        if key in author_answers.ANSWERS:
            decision, fix = author_answers.ANSWERS[key]
            if fix:
                fixes[key] = fix
        elif comment:
            try:
                answer = author_answers.interpret(kind_of_key(key, row[st]), comment, ranks)
            except ValueError as e:
                errors.append(f'{key}: {e}')
                continue
            if answer:
                decision, fix = answer
                if fix:
                    fixes[key] = fix
        if decision:
            decisions[key] = decision
    if errors:
        raise SystemExit('Ответы автора, которые скрипт не понял:\n  ' + '\n  '.join(errors))
    return decisions, fixes, comments, info


def rekey(tables, info, rows):
    """Коды строк листов кавалеров — номер строки в листе. Автор вставляет
    людей в середину листа, строки сдвигаются, и код из прошлой сверки
    указывает на другого человека. Такие коды находим заново по ФИО и году."""
    by_code = {f"R{r['xl_row']}": r for r in rows}
    by_name = defaultdict(list)
    for r in rows:
        by_name[(str(r['xl_row']).split('_')[0], norm(fio(r)), birth_of(r)[1])].append(r)
    moved, lost = {}, []
    for key, (name, year) in info.items():
        code = re.match(r'R[^-]+', key)
        if not code or '_' not in code.group(0):
            continue
        code = code.group(0)
        r = by_code.get(code)
        if r and norm(fio(r)) == norm(name):
            continue
        found = by_name.get((code[1:].split('_')[0], norm(name), year), [])
        if len(found) == 1:
            moved[key] = f"R{found[0]['xl_row']}" + key[len(code):]
        else:
            lost.append(key)
    for table in tables:
        for old, new in moved.items():
            if old in table:
                table[new] = table.pop(old)
    if moved:
        print(f'Коды строк кавалеров сдвинулись в новом списке — найдены по ФИО: {len(moved)}')
    if lost:
        print(f'⚠ Не нашлись в новом списке (решение не перенесено): {lost}')
    return moved


def blank_row(**fields):
    r = {key: None for key in COLS.values()}
    r.update(main=False, notes=[], awards=[], src=['ответ автора'])
    r.update(fields)
    r['orig_fio'] = fio(r)
    return r


def author_rows(rows, persons, fixes):
    """Поправки автора к строкам списка (ROW_FIXES) и новые строки из его
    ответов: данные человека, который есть только на сайте («Оставить: ФИО,
    год, звание…»), и вторая служба (EXTRA_ROWS)."""
    by_code = {f"R{r['xl_row']}": r for r in rows}
    for code, patch in author_answers.ROW_FIXES.items():
        if code not in by_code:
            raise SystemExit(f'Поправка автора к {code}: такой строки в списке нет')
        by_code[code].update(patch)

    extra = []
    for code, patches in author_answers.EXTRA_ROWS.items():
        base = by_code[code]
        for i, patch in enumerate(patches, 1):
            r = blank_row(last=base['last'], first=base['first'], patr=base['patr'],
                          year=base['year'], birth=base['birth'], **patch)
            r.update(xl_row=f"{base['xl_row']}+{i}", same_as=base['xl_row'],
                     plan_key=f'{code}+{i}')
            extra.append(r)

    by_id = {p['id']: p for p in persons}
    for key, fix in fixes.items():
        data = fix.get('data')
        if not data:
            continue
        data = dict(data, **author_answers.DATA_FIXES.get(key, {}))
        p = by_id.get(int(key[1:]))
        if p is None:
            print(f'⚠ {key}: на сайте такого человека уже нет — ответ автора пропущен')
            continue
        words = data['name']
        if len(words) >= 2 and not any(w.endswith('.') for w in words):
            last, first, patr = words[0], words[1], ' '.join(words[2:]) or None
        else:                       # «Варфоломеев У. И.» — инициалы, имя берём с сайта
            last, first, patr = p['last_name'], p['first_name'], p['patronymic']
        r = blank_row(last=last, first=first, patr=patr, year=data['year'],
                      rank=data['rank'], role=data['role'], unit=data['unit'],
                      fate=data['fate'], out_date=data['out_date'],
                      out_reason=data['out_reason'])
        r.update(xl_row=f'A{key}', force_pid=p['id'], plan_key=key)
        extra.append(r)
    return extra


# ── Запись в базу ───────────────────────────────────────────────────────────

def sheet_tags(r):
    """Листы таблицы автора, где есть человек: «Командный состав»,
    «Кавалеры III», «Рабочие войны». Строки из ответов автора листа не имеют."""
    return [s for s in r['src'] if s != 'ответ автора']


def apply_plan(db, plan, units, division_id):
    from psycopg2.extras import Json
    cur = db.cursor()
    done = Counter()

    for name in sorted(plan['missing_units']):
        cur.execute("INSERT INTO units (division_id, name, unit_type) "
                    "VALUES (%s, %s, 'unit') RETURNING id", (division_id, name))
        units[name] = cur.fetchone()[0]

    def write_awards(person_id, awards):
        for a in awards:
            cur.execute("""INSERT INTO person_awards (person_id, kind, degree, title, status, source)
                           VALUES (%s, %s, %s, %s, %s, %s)
                           ON CONFLICT (person_id, kind, coalesce(degree, 0), status) DO NOTHING""",
                        (person_id, a['kind'], a['degree'], a['title'], a['status'], a['source']))
            done['наград'] += cur.rowcount

    def write_record(person_id, op, target, rec):
        if op == 'insert':
            cur.execute("""INSERT INTO service_records
                             (person_id, unit_id, division_id, rank, role, date_to, end_reason, note)
                           VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
                        (person_id, units.get(rec['unit']), division_id, rec['rank'],
                         rec['role'], rec['date_to'], rec['end_reason'], rec['note']))
        elif op == 'update':
            values = {f: new for f, (_, new) in rec.items()}
            if 'unit' in values:
                values['unit_id'] = units[values.pop('unit')]
            sets = ', '.join(f'{f} = %s' for f in values)
            cur.execute(f'UPDATE service_records SET {sets} WHERE id = %s',
                        list(values.values()) + [target['id']])

    gone = set()

    def merge(drop, keep):
        """Запись drop — дубль keep: всё ценное переносится к keep, drop удаляется."""
        # Пустая запись дубля в той же части ничего не добавляет.
        cur.execute(r"""DELETE FROM service_records d
                        WHERE d.person_id = %s
                          AND coalesce(d.role, '') ~ '^[\s?—-]*$'
                          AND EXISTS (SELECT 1 FROM service_records k
                                       WHERE k.person_id = %s
                                         AND k.unit_id IS NOT DISTINCT FROM d.unit_id)""",
                    (drop, keep))
        cur.execute('UPDATE service_records SET person_id = %s WHERE person_id = %s', (keep, drop))
        cur.execute('UPDATE documents SET person_id = %s WHERE person_id = %s', (keep, drop))
        cur.execute("""INSERT INTO person_awards
                         (person_id, kind, degree, title, status, document_url, note, source)
                       SELECT %s, kind, degree, title, status, document_url, note, source
                         FROM person_awards WHERE person_id = %s
                       ON CONFLICT (person_id, kind, coalesce(degree, 0), status) DO NOTHING""",
                    (keep, drop))
        cur.execute("""INSERT INTO person_events (person_id, event_id)
                       SELECT %s, event_id FROM person_events WHERE person_id = %s
                       ON CONFLICT DO NOTHING""", (keep, drop))
        # Фото у дубля бывает, а у основной записи нет.
        cur.execute("""UPDATE persons k SET
                         photo_url = COALESCE(k.photo_url, d.photo_url),
                         sources = (SELECT jsonb_agg(DISTINCT x)
                                      FROM jsonb_array_elements(k.sources || d.sources) x),
                         last_updated = NOW()
                       FROM persons d WHERE k.id = %s AND d.id = %s""", (keep, drop))
        cur.execute('DELETE FROM persons WHERE id = %s', (drop,))
        gone.add(drop)

    # «В списке дважды»: строка, которую автор не оставил, в базу не идёт.
    dropped, list_pairs = set(), []
    for it in plan['items']:
        if it['kind'] == 'list_twin' and it['decision'] == 'Оставить первую':
            dropped.add(it['other']['xl_row'])
            list_pairs.append((it['other']['xl_row'], it['row']['xl_row']))
        elif it['kind'] == 'list_twin' and it['decision'] == 'Оставить вторую':
            dropped.add(it['row']['xl_row'])
            list_pairs.append((it['row']['xl_row'], it['other']['xl_row']))
    matched_pid = {it['row']['xl_row']: it['person']['id'] for it in plan['items']
                   if it['kind'] == 'match'}

    for it in plan['items']:
        kind, d, p = it['kind'], it['decision'], it['person']
        if kind in ('match', 'new') and it['row']['xl_row'] in dropped:
            done['строк-дублей пропущено'] += 1
            continue

        if kind == 'match':
            # Лист таблицы автора — даже при «Не обновлять»: по нему на сайте
            # работает фильтр «Командный состав». Дописываем к тому, что уже
            # есть: у человека может быть несколько строк в плане.
            cur.execute("""UPDATE persons SET sources = coalesce(
                             (SELECT jsonb_agg(DISTINCT x)
                                FROM jsonb_array_elements(coalesce(sources, '[]') || %s) x),
                             '[]')
                           WHERE id = %s""",
                        (Json(sheet_tags(it['row']) + ([SOURCE_TAG] if d != 'Не обновлять' else [])),
                         p['id']))
        if kind == 'match' and d != 'Не обновлять':
            changes = it['changes']
            if changes:
                sets = [f'{f} = %s' for f in changes] + ['last_updated = NOW()']
                params = [new for _, new in changes.values()]
                cur.execute(f"UPDATE persons SET {', '.join(sets)} WHERE id = %s", params + [p['id']])
            op, target, rec = it['op']
            write_record(p['id'], op, target, rec)
            write_awards(p['id'], it['new_awards'])
            done['обновлено'] += 1

        elif kind == 'new' and d != 'Не добавлять':
            pr = it['person_new']
            cur.execute("""INSERT INTO persons
                             (last_name, first_name, patronymic, birth_date, death_date,
                              outcome, outcome_details, sources, last_updated)
                           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW()) RETURNING id""",
                        (pr['last_name'], pr['first_name'], pr['patronymic'], pr['birth_date'],
                         pr['death_date'], pr['outcome'], pr['outcome_details'],
                         Json([SOURCE_TAG] + sheet_tags(it['row']))))
            new_id = cur.fetchone()[0]
            write_record(new_id, 'insert', None, it['record'])
            write_awards(new_id, it['awards'])
            done['добавлено'] += 1

        elif kind in ('only_db', 'twin') and d == 'Удалить':
            # Записи о службе и связи с событиями уходят каскадом.
            cur.execute('DELETE FROM persons WHERE id = %s', (p['id'],))
            done['удалено'] += 1

        elif kind == 'twin' and d == 'Объединить':
            merge(p['id'], it['twin']['id'])
            done['объединено'] += 1

    # Обе строки пары «в списке дважды» были на сайте двумя записями — раз
    # автор сказал, что это один человек, сливаем и записи на сайте.
    for drop_row, keep_row in list_pairs:
        drop, keep = matched_pid.get(drop_row), matched_pid.get(keep_row)
        if drop and keep and drop != keep and drop not in gone and keep not in gone:
            merge(drop, keep)
            done['объединено (дважды в списке)'] += 1

    db.commit()
    cur.close()
    return done


# ── Главная функция ─────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[1])
    ap.add_argument('--file', required=True, help='xlsx со списком командного состава')
    ap.add_argument('--dump', help='читать базу из дампа pg_dump (.sql или .sql.gz)')
    ap.add_argument('--sverka', default='sverka_komsostav.xlsx', help='куда записать файл сверки')
    ap.add_argument('--decisions', help='файл сверки с решениями автора')
    ap.add_argument('--apply', action='store_true', help='записать изменения в базу')
    args = ap.parse_args()

    if args.apply and args.dump:
        sys.exit('--apply пишет в живую базу, вместе с --dump его не запустить')
    if args.apply and not args.decisions:
        sys.exit('--apply только по файлу сверки с решениями автора: --decisions')
    if not Path(args.file).exists():
        sys.exit(f'Не найден: {args.file}')

    rows, questions, spellings = load_list(args.file)
    rows = apply_answers(rows, questions, spellings, [])
    merged = []
    rows = merge_sheets(rows, merged)
    print(f'Людей в списке: {len(rows)} (из них в нескольких листах сразу: {len(merged)})')
    for line in merged:
        print('  ', line)

    db = None
    if args.dump:
        persons, records, units, division_id = load_from_dump(args.dump)
    else:
        db = get_db()
        persons, records, units, division_id = load_from_db(db.cursor())
    print(f'Персон в базе: {len(persons)}')

    ranks = {r['rank'] for r in rows if r['rank']}
    decisions, fixes, comments, info = read_decisions(args.decisions, ranks)
    rekey((decisions, fixes, comments), info, rows)
    extra = author_rows(rows, persons, fixes)
    if extra:
        print(f'Строк из ответов автора: {len(extra)}')
    plan = build_plan(rows + extra, persons, records, units, decisions, fixes, comments)
    counts = write_sverka(plan, units, args.sverka)
    for k, (label, _, _) in KINDS.items():
        print(f'  {label:<40} {counts.get(k, 0)}')
    print(f'Сверка: {args.sverka}')

    if args.apply:
        done = apply_plan(db, plan, units, division_id)
        print('✓ Записано в базу:', ', '.join(f'{k} {v}' for k, v in done.items()))
    if db:
        db.close()


if __name__ == '__main__':
    main()
