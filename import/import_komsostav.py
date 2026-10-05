"""
import_komsostav.py — импорт списка «Командный состав 52 сд» без потери данных

В отличие от import_to_db.py --reset, ничего не стирает: людей, которые
уже есть в базе, уточняет, новых добавляет, остальных не трогает.
Ответы заказчика с листов «Вопросы» и «Написание ФИО» применяются
к списку до сопоставления.

Сопоставление с базой — по ФИО и году рождения. Кроме написания из
нового списка пробуется написание из poteri.xlsx (столбец «Написание в
poteri.xlsx»): база собиралась оттуда. Похожие, но не совпавшие записи
(«Выхрыстюк / Выхристюк») сами не склеиваются — попадают на лист
«Спорные», ответ на них даётся там же и подаётся обратно через --answers.

Режимы:
  1. Отчёт без записи в базу, по дампу (можно на своём компьютере):
       python import_komsostav.py --file data/Командный_состав_52сд-в-02.xlsx \
           --dump ../backups/52sd_2026-10-06.sql.gz --report report.xlsx

  2. Отчёт по живой базе (на сервере, база не меняется):
       python3 import_komsostav.py --file ... --report report.xlsx

  3. Запись в базу (на сервере, одной транзакцией):
       python3 import_komsostav.py --file ... --report report.xlsx \
           [--answers report_с_ответами.xlsx] --apply

  Перед --apply — свежий дамп базы. Зависимости на сервере:
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
    'управлениядивизии52сд': '52сд',
}

# Часть не установлена — записи о службе остаются без подразделения.
UNIT_UNKNOWN = {'?сп', '?полка', '???сп/ап'}

# «убит», «погиб» и т.п. уже выражены в outcome — в подробности не пишем.
PLAIN_REASONS = {'убит', 'убита', 'погиб', 'умер от ран', 'умерла от ран',
                 'пропал без вести', 'пропала без вести'}

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
        rows.append(r)

    _, questions = read_sheet(wb, 'Вопросы')
    _, spellings = read_sheet(wb, 'Написание ФИО')
    return rows, questions, spellings


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
    """«27.04.1944», «8.04.44», число-дата Excel → 'YYYY-MM-DD'."""
    if not s:
        return None
    s = s.strip()
    if re.fullmatch(r'\d{5}', s):
        return (date(1899, 12, 30) + timedelta(days=int(s))).isoformat()
    m = re.fullmatch(r'(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(\s*г\.?)?', s)
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


def outcome_of(r):
    fate = r['fate']
    text = (r['out_reason'] or '').lower()
    alive = 'жив' in text and 'умер' not in text
    if fate == 'погиб' or fate == 'осуждён и погиб':
        if alive:
            return 'survived'
        if 'умер' in text and 'ран' in text:
            return 'died_of_wounds'
        return 'killed'
    if fate == 'пропал без вести':
        return 'survived' if alive or 'выжил' in text else 'missing'
    if fate == 'плен':
        return 'pow_survived' if 'освобожден' in text else 'pow'
    if fate == 'расстрелян':
        return 'executed'
    if fate == 'умер':
        return 'died_of_wounds' if 'ран' in text else None
    return None


def details_of(r):
    reason = r['out_reason']
    if reason and norm(reason) not in PLAIN_REASONS:
        return reason
    return None


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
                if m and m.group(1) in ('persons', 'service_records', 'units', 'divisions'):
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
    return tables['persons'], tables['service_records'], units, division_id


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


def match(rows, persons, answers):
    by_fio = defaultdict(list)
    for p in persons:
        by_fio[(norm(p['last_name']), norm(p['first_name']), norm(p['patronymic']))].append(p)
    year = lambda p: (p['birth_date'] or '')[:4] or None

    claimed = {}           # person id → строка списка
    result = {}            # xl_row → ('match'|'new'|'disputed', person, why)

    def candidates(r):
        found = {}
        for v in name_variants(r):
            for p in by_fio.get(v, []):
                found[p['id']] = p
        if not found:
            # Слитые при старом разборе имена: «АрхипенкоТихон» «Алексеевич1902».
            for v in name_variants(r):
                key = letters(''.join(x or '' for x in v))
                if len(key) < 10:
                    continue
                for p in persons:
                    joined = letters((p['last_name'] or '') + (p['first_name'] or '')
                                     + (p['patronymic'] or ''))
                    if joined.startswith(key):
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
            result[r['xl_row']] = ('match', exact[0], None)
        elif len(exact) > 1:
            result[r['xl_row']] = ('disputed', exact, 'в базе несколько человек с таким ФИО и годом')
        else:
            pending.append((r, y, cands))

    unmatched = []
    for r, y, cands in pending:
        free = [p for p in cands if p['id'] not in claimed]
        if len(free) == 1:
            p = free[0]
            py = year(p)
            if py and y and abs(int(py) - int(y)) > 2:
                result[r['xl_row']] = ('disputed', [p], f'ФИО совпало, год {py} в базе и {y} в списке')
            else:
                claimed[p['id']] = r
                result[r['xl_row']] = ('match', p, None)
            continue
        if len(free) > 1:
            result[r['xl_row']] = ('disputed', free, 'в базе несколько человек с таким ФИО')
            continue
        if cands:
            result[r['xl_row']] = ('disputed', cands, 'запись в базе уже занята другой строкой списка')
            continue
        unmatched.append((r, y))

    # Проход 3: похожие написания — только как вопрос, без автоматического
    # слияния. Идёт последним: иначе «Сенюгин» успевает занять запись
    # «Синюгина» раньше, чем до неё дойдёт сам Синюгин из списка.
    fuzzy_pool = [(p, letters((p['last_name'] or '') + (p['first_name'] or '')
                              + (p['patronymic'] or ''))) for p in persons]
    for r, y in unmatched:
        key = letters(fio(r))
        best, best_ratio = None, 0
        for p, joined in fuzzy_pool:
            if p['id'] in claimed:
                continue
            py = year(p)
            if py and y and py != y:
                continue
            # Опечатки бывают в середине фамилии, не в первой букве:
            # без этого «Кузин» и «Кулик» тянутся к «Чудину».
            if not joined or joined[0] != key[:1]:
                continue
            ratio = difflib.SequenceMatcher(None, key, joined).ratio()
            if ratio > best_ratio:
                best, best_ratio = p, ratio
        if best and best_ratio >= 0.85:
            result[r['xl_row']] = ('disputed', [best], f'похожее написание ({best_ratio:.0%})')
        else:
            result[r['xl_row']] = ('new', None, None)

    # Ответы на «Спорные» из прошлого отчёта.
    by_id = {p['id']: p for p in persons}
    for r in rows:
        kind, cand, why = result[r['xl_row']]
        if kind != 'disputed' or r['xl_row'] not in answers:
            continue
        verdict, pid = answers[r['xl_row']]
        if verdict == 'Тот же человек' and pid in by_id and pid not in claimed:
            claimed[pid] = r
            result[r['xl_row']] = ('match', by_id[pid], f'по ответу: {why}')
        elif verdict == 'Новый человек':
            result[r['xl_row']] = ('new', None, f'по ответу: {why}')

    return result, claimed


# ── План изменений ──────────────────────────────────────────────────────────

def build_plan(rows, persons, records, units, answers):
    result, claimed = match(rows, persons, answers)
    recs_by_person = defaultdict(list)
    for s in records:
        recs_by_person[s['person_id']].append(s)

    plan = {'updates': [], 'inserts': [], 'disputed': [], 'unit_notes': Counter(),
            'missing_units': set(), 'record_ops': []}

    for r in rows:
        kind, p, why = result[r['xl_row']]
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

        if kind == 'disputed':
            plan['disputed'].append((r, p, why))
            continue
        if kind == 'new':
            plan['inserts'].append((r, person, record, why))
            continue

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
        plan['updates'].append((r, p, changes, op, why))

    plan['matched_ids'] = set(claimed)
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


# ── Отчёт ───────────────────────────────────────────────────────────────────

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


def ru(field, v):
    if v is None:
        return '—'
    return OUTCOME_RU.get(v, v) if field == 'outcome' else v


def where(r):
    tail = '' if r['main'] else ' (доп. блок)'
    return f"№{r['num']}{tail}, стр. {r['xl_row']}"


def p_fio(p):
    return ' '.join(x for x in (p['last_name'], p['first_name'], p['patronymic']) if x)


def write_report(plan, rows, persons, units, log, path):
    unit_names = {uid: name for name, uid in units.items()}
    wb = openpyxl.Workbook()
    head = Font(bold=True)
    yellow = PatternFill('solid', fgColor='FFF2CC')

    def sheet(title, header, data, widths):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font = head
        for row in data:
            ws.append(list(row))
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        ws.freeze_panes = 'A2'
        if data:
            ws.auto_filter.ref = ws.dimensions
        return ws

    upd = plan['updates']
    changed = [u for u in upd if u[2] or u[3][0] != 'same']
    rec_new = sum(1 for u in upd if u[3][0] == 'insert')
    rec_upd = sum(1 for u in upd if u[3][0] == 'update')
    untouched = [p for p in persons if p['id'] not in plan['matched_ids']]

    ws = wb.active
    ws.title = 'Сводка'
    summary = [
        ('Строк в новом списке после ответов на вопросы', len(rows)),
        ('', ''),
        ('Уже есть в базе', len(upd)),
        ('   из них что-то уточнится', len(changed)),
        ('   из них без изменений', len(upd) - len(changed)),
        ('   добавится запись о службе', rec_new),
        ('   уточнится запись о службе', rec_upd),
        ('Новые — будут добавлены', len(plan['inserts'])),
        ('Спорные — НЕ импортируются, пока нет ответа', len(plan['disputed'])),
        ('', ''),
        ('Есть в базе, но нет в списке — не трогаем', len(untouched)),
        ('', ''),
        ('Национальность из списка не переносится: в базе для неё нет поля.', ''),
        ('Фотографии не трогаем: они подбираются отдельно, match_photos.py.', ''),
    ]
    for row in summary:
        ws.append(row)
    for c in ws['A']:
        if c.value and not str(c.value).startswith(' '):
            c.font = head
    ws.column_dimensions['A'].width = 62
    ws.column_dimensions['B'].width = 10

    disp = []
    for r, cands, why in plan['disputed']:
        for p in (cands if isinstance(cands, list) else [cands]):
            disp.append((r['xl_row'], where(r), fio(r), birth_of(r)[1] or '—', r['unit'] or '—',
                         p['id'], p_fio(p), (p['birth_date'] or '')[:4] or '—', why, ''))
    ws = sheet('Спорные',
               ['Строка Excel', 'Где в списке', 'ФИО в списке', 'Год', 'Часть',
                'id в базе', 'ФИО в базе', 'Год в базе', 'Почему спорно', 'Ответ'],
               disp, [8, 18, 34, 7, 10, 8, 34, 9, 40, 18])
    dv = DataValidation(type='list', formula1='"Тот же человек,Новый человек,Пропустить"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    for i in range(2, len(disp) + 2):
        dv.add(f'J{i}')
        ws[f'J{i}'].fill = yellow

    sheet('Новые',
          ['Строка Excel', 'Где в списке', 'ФИО', 'Дата рождения', 'Звание', 'Должность',
           'Часть', 'Судьба', 'Примечание'],
          [(r['xl_row'], where(r), fio(r), pr['birth_date'] or '—', rec['rank'], rec['role'],
            rec['unit'] or '—', ru('outcome', pr['outcome']), why or '')
           for r, pr, rec, why in plan['inserts']],
          [8, 18, 34, 12, 20, 44, 34, 14, 30])

    lines = []
    for r, p, changes, (op, target, diff), why in upd:
        for field, (old, new) in changes.items():
            lines.append((p['id'], p_fio(p), where(r), FIELD_RU[field],
                          ru(field, old), ru(field, new)))
        if op == 'insert':
            rec = diff
            had = '; '.join(f"{s['role'] or '—'} · {unit_names.get(s['unit_id'], 'без части')}"
                            for s in target)
            lines.append((p['id'], p_fio(p), where(r), 'служба: новая запись',
                          f'в базе остаётся: {had}' if had else 'записей не было',
                          f"{rec['rank'] or ''} · {rec['role'] or ''} · {rec['unit'] or '—'}"))
        elif op == 'update':
            for field, (old, new) in diff.items():
                lines.append((p['id'], p_fio(p), where(r), f'служба: {FIELD_RU[field]}',
                              old or '—', new))
    sheet('Уточнения', ['id в базе', 'ФИО в базе', 'Где в списке', 'Что', 'Было', 'Станет'],
          lines, [8, 34, 18, 22, 40, 50])

    # Слитые старым разбором имена («АрхипенкоТихон» «Алексеевич1902»)
    # часто лежат рядом с нормальной записью того же человека — дубль.
    joined = lambda p: letters(p_fio(p).replace(' ', ''))
    matched = [(joined(p), p) for p in persons if p['id'] in plan['matched_ids']]

    def twin(p):
        key = joined(p)
        for k, m in matched:
            if len(k) >= 10 and key.startswith(k):
                return f"похоже на дубль id {m['id']} {p_fio(m)}"
        return ''

    sheet('Нет в списке',
          ['id в базе', 'ФИО в базе', 'Год', 'Откуда в базе', 'Заметка'],
          [(p['id'], p_fio(p), (p['birth_date'] or '')[:4] or '—', ', '.join(p['sources']),
            twin(p)) for p in sorted(untouched, key=p_fio)],
          [8, 40, 7, 40, 50])

    unit_rows = [(note, n) for note, n in plan['unit_notes'].most_common()]
    unit_rows += [(f'Подразделения «{u}» нет в базе — будет создано', '') for u in
                  sorted(plan['missing_units'])]
    sheet('Части', ['Что с частью', 'Строк'], unit_rows, [70, 8])

    sheet('Ответы применены', ['Вопрос', 'Тема', 'Кто', 'Ответ', 'Что сделано'],
          log, [8, 30, 34, 50, 40])

    wb.save(path)


def read_answers(path):
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for row in list(wb['Спорные'].iter_rows(values_only=True))[1:]:
        xl_row, pid, verdict = row[0], row[5], clean(row[9])
        if verdict in ('Тот же человек', 'Новый человек'):
            # Строка может стоять напротив нескольких кандидатов —
            # «Тот же человек» берётся у того, напротив кого он выбран.
            if verdict == 'Тот же человек' or xl_row not in out:
                out[int(xl_row)] = (verdict, int(pid))
    return out


# ── Запись в базу ───────────────────────────────────────────────────────────

def apply_plan(db, plan, units, division_id):
    from psycopg2.extras import Json
    cur = db.cursor()

    for name in sorted(plan['missing_units']):
        cur.execute("INSERT INTO units (division_id, name, unit_type) "
                    "VALUES (%s, %s, 'unit') RETURNING id", (division_id, name))
        units[name] = cur.fetchone()[0]

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

    for r, p, changes, (op, target, rec), _ in plan['updates']:
        sources = list(p['sources'])
        if SOURCE_TAG not in sources:
            sources.append(SOURCE_TAG)
        sets = [f'{f} = %s' for f in changes] + ['sources = %s']
        params = [new for _, new in changes.values()] + [Json(sources)]
        if changes:
            sets.append('last_updated = NOW()')
        cur.execute(f"UPDATE persons SET {', '.join(sets)} WHERE id = %s", params + [p['id']])
        write_record(p['id'], op, target, rec)

    for r, pr, rec, _ in plan['inserts']:
        cur.execute("""INSERT INTO persons
                         (last_name, first_name, patronymic, birth_date, death_date,
                          outcome, outcome_details, sources, last_updated)
                       VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW()) RETURNING id""",
                    (pr['last_name'], pr['first_name'], pr['patronymic'], pr['birth_date'],
                     pr['death_date'], pr['outcome'], pr['outcome_details'],
                     Json([SOURCE_TAG])))
        write_record(cur.fetchone()[0], 'insert', None, rec)

    db.commit()
    cur.close()


# ── Главная функция ─────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[1])
    ap.add_argument('--file', required=True, help='xlsx со списком командного состава')
    ap.add_argument('--dump', help='читать базу из дампа pg_dump (.sql или .sql.gz)')
    ap.add_argument('--report', default='report_komsostav.xlsx')
    ap.add_argument('--answers', help='отчёт с заполненным столбцом «Ответ» на листе «Спорные»')
    ap.add_argument('--apply', action='store_true', help='записать изменения в базу')
    args = ap.parse_args()

    if args.apply and args.dump:
        sys.exit('--apply пишет в живую базу, вместе с --dump его не запустить')
    if not Path(args.file).exists():
        sys.exit(f'Не найден: {args.file}')

    rows, questions, spellings = load_list(args.file)
    log = []
    rows = apply_answers(rows, questions, spellings, log)
    print(f'Строк в списке: {len(rows)}')

    db = None
    if args.dump:
        persons, records, units, division_id = load_from_dump(args.dump)
    else:
        db = get_db()
        persons, records, units, division_id = load_from_db(db.cursor())
    print(f'Персон в базе: {len(persons)}')

    plan = build_plan(rows, persons, records, units, read_answers(args.answers))
    write_report(plan, rows, persons, units, log, args.report)
    print(f"Уже в базе: {len(plan['updates'])}, новых: {len(plan['inserts'])}, "
          f"спорных: {len(plan['disputed'])}")
    print(f'Отчёт: {args.report}')

    if args.apply:
        apply_plan(db, plan, units, division_id)
        print('✓ Записано в базу')
    if db:
        db.close()


if __name__ == '__main__':
    main()
