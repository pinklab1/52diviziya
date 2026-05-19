"""
parse_alfavit.py — парсинг алфавитного списка офицеров 52-й дивизии
Работает без API — только регулярные выражения и python-docx.

Установка:
  pip install python-docx python-dotenv

Запуск:
  python parse_alfavit.py --file АЛФАВИТ.docx --dry-run
  python parse_alfavit.py --file АЛФАВИТ.docx --output result.json
"""

import argparse
import json
import re
import sys
from pathlib import Path

# ── Словари расшифровки ───────────────────────────────────────────────────

RANKS = {
    'млл-т':            'младший лейтенант',
    'мл.л-т':           'младший лейтенант',
    'мл л-т':           'младший лейтенант',
    'стл-т':            'старший лейтенант',
    'ст.л-т':           'старший лейтенант',
    'л-т':              'лейтенант',
    'гвстл-т':          'гвардии старший лейтенант',
    'стл-тмс':          'старший лейтенант медслужбы',
    'млл-тмс':          'младший лейтенант медслужбы',
    'мл.л-тмс':         'младший лейтенант медслужбы',
    'техник-лейтенант': 'техник-лейтенант',
    'техл-т':           'техник-лейтенант',
    'майоринтсл':       'майор интендантской службы',
    'майор':            'майор',
    'капитан':          'капитан',
    'подполковник':     'подполковник',
    'полковник':        'полковник',
    'генерал-майор':    'генерал-майор',
    'военфельд':        'военный фельдшер',
    'военфельдшер':     'военный фельдшер',
}

ROLES = {
    'комбат':           'командир батальона',
    'комвзв':           'командир взвода',
    'комвзвода':        'командир взвода',
    'кмвзв':            'командир взвода',
    'комср':            'командир стрелковой роты',
    'комсроты':         'командир стрелковой роты',
    'компульроты':      'командир пулемётной роты',
    'компульвзв':       'командир пулемётного взвода',
    'комминроты':       'командир миномётной роты',
    'комминбата':       'командир миномётного батальона',
    'комминвзв':        'командир миномётного взвода',
    'комогнвзв':        'командир огневого взвода',
    'комвзвсвязи':      'командир взвода связи',
    'комротысвязи':     'командир роты связи',
    'комсанвзв':        'командир санитарного взвода',
    'комбатр':          'командир батареи',
    'командир':         'командир',
    'ком':              'командир',
    'комроты':          'командир роты',
    'начштаб':          'начальник штаба',
    'нш':               'начальник штаба',
    'замком':           'заместитель командира',
    'замкомдив':        'заместитель командира дивизии',
    'замполит':         'заместитель по политчасти',
    'пнш':              'помощник начальника штаба',
    'пнш-1':            'помощник начальника штаба-1',
    'пнш-2':            'помощник начальника штаба-2',
    'пнш-3':            'помощник начальника штаба-3',
    'пнш-4':            'помощник начальника штаба-4',
    'военком':          'военный комиссар',
    'военкомсб':        'военком стрелкового батальона',
    'парторг':          'партийный организатор',
    'начразв':          'начальник разведки',
    'начразведки':      'начальник разведки',
    'адст':             'адъютант старший',
    'офицсвязи':        'офицер связи',
    'врач':             'врач',
    'млврач':           'младший врач',
    'мл.врач':          'младший врач',
    'фельд':            'фельдшер',
    'ветфельд':         'ветеринарный фельдшер',
    'ветлазар':         'ветеринарный лазарет',
    'вет':              'ветеринар',
    'х':                'хозяйственная часть',
    'агитатор':         'агитатор',
    'редакция':         'редакция',
    'снаб':             'снабжение',
    'поммедснаб':       'помощник по медснабжению',
    'начпфс':           'начальник продфуражной службы',
    'начотдтроф':       'начальник отдела трофеев',
    'овсполка':         'офицер по военной службе полка',
    'замкомсб':         'заместитель командира батальона',
    'помнач':           'помощник начальника',
    'полкинж':          'полковой инженер',
    'полкинжен':        'полковой инженер',
    'секретарьдпк':     'секретарь ДПК',
    'комвзвпешразв':    'командир взвода пешей разведки',
    'комогнвзвбатр':    'командир огневого взвода батареи',
    'комвзвбатр':       'командир взвода батареи',
    'начотдела':        'начальник отдела',
    'военфельд':        'военный фельдшер',
    'фельдсанроты':     'фельдшер санитарной роты',
    'фельдстсанроты':   'фельдшер старший санитарной роты',
}

UNITS = {
    '429сп':    '429-й стрелковый полк',
    '429':      '429-й стрелковый полк',
    '431сп':    '431-й стрелковый полк',
    '431':      '431-й стрелковый полк',
    '439сп':    '439-й стрелковый полк',
    '439':      '439-й стрелковый полк',
    '1028ап':   '1028-й артиллерийский полк',
    '1028':     '1028-й артиллерийский полк',
    '52сд':     'Штаб 52-й стрелковой дивизии',
    '52':       'Штаб 52-й стрелковой дивизии',
    '106':      '106-е подразделение',
    '164осапб': '164-й отдельный сапёрный батальон',
    '164':      '164-й отдельный сапёрный батальон',
    '405оиптд': '405-й истребительно-противотанковый дивизион',
    '405':      '405-й истребительно-противотанковый дивизион',
    '127':      '127-е подразделение',
    '201':      '201-е подразделение',
}

# ── Чтение файла ──────────────────────────────────────────────────────────

def read_docx(path: str) -> list[str]:
    from docx import Document
    doc = Document(path)
    lines = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(' '.join(cells))
    return lines


def clean_lines(lines: list[str]) -> list[str]:
    """Убирает заголовки букв и пустые строки."""
    result = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # Заголовки типа "** А (45) **" или просто "А"
        if re.match(r'^\*+\s*[А-ЯЁ]\s*[\(\*]', line):
            continue
        if re.match(r'^[А-ЯЁ]\s*\(\d+\)', line):
            continue
        result.append(line)
    return result

# ── Нормализация строки ───────────────────────────────────────────────────

def normalize(line: str) -> str:
    """Заменяет _ на пробел, убирает лишние пробелы."""
    line = line.replace('_', ' ')
    line = re.sub(r'\s+', ' ', line)
    return line.strip()

# ── Извлечение даты гибели ────────────────────────────────────────────────

def extract_death(line: str) -> tuple[str, str | None]:
    """Возвращает (строка без даты гибели, дата гибели YYYY-MM-DD или None)."""
    m = re.search(r'\(\+(\d{1,2})\.(\d{2})\.(\d{2,4})\)', line)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = '19' + y
        death = f"{y}-{mo.zfill(2)}-{d.zfill(2)}"
        line = line[:m.start()] + line[m.end():]
        return line.strip(), death
    return line, None

# ── Извлечение ФИО ────────────────────────────────────────────────────────

# Паттерн: три слова с заглавной буквы подряд (фамилия имя отчество)
FIO_RE = re.compile(
    r'^([А-ЯЁ][а-яё]+(?:[А-ЯЁ][а-яё]+)?)\s+'   # фамилия (может быть слитно двойная)
    r'([А-ЯЁ][а-яё]+)\s+'                         # имя
    r'([А-ЯЁ][а-яё]+(?:овна|евна|авна|ович|евич|ич)?)'  # отчество
)

def extract_fio(line: str) -> tuple[str | None, str | None, str | None, str]:
    """Возвращает (фамилия, имя, отчество, остаток строки)."""
    m = FIO_RE.match(line)
    if m:
        last, first, patr = m.groups()
        rest = line[m.end():].strip()
        return last, first, patr, rest
    # Попробуем хотя бы фамилию
    m2 = re.match(r'^([А-ЯЁ][а-яё]+)\s+(.*)', line)
    if m2:
        return m2.group(1), None, None, m2.group(2)
    return None, None, None, line

# ── Извлечение года рождения ──────────────────────────────────────────────

def extract_birth_year(rest: str) -> tuple[int | None, str]:
    m = re.search(r'\b(18\d{2}|19[0-2]\d)\b', rest)
    if m:
        year = int(m.group(1))
        rest = rest[:m.start()] + rest[m.end():]
        return year, rest.strip()
    return None, rest

# ── Извлечение звания ─────────────────────────────────────────────────────

def extract_rank(rest: str) -> tuple[str | None, str]:
    rest_lower = rest.lower()
    # Ищем по убыванию длины чтобы длинные совпадали раньше коротких
    for abbr in sorted(RANKS.keys(), key=len, reverse=True):
        pattern = r'(?<![а-яё])' + re.escape(abbr) + r'(?![а-яё])'
        if re.search(pattern, rest_lower):
            rank = RANKS[abbr]
            rest = re.sub(pattern, '', rest_lower, flags=re.IGNORECASE)
            rest = re.sub(r'\s+', ' ', rest).strip()
            return rank, rest
    return None, rest

# ── Извлечение подразделений ──────────────────────────────────────────────

UNIT_RE = re.compile(
    r'\b(1028ап|1028|429сп|429|431сп|431|439сп|439|'
    r'52сд|52|164осапб|164|405оиптд|405|106|127|201)\b',
    re.IGNORECASE
)

def extract_units(rest: str) -> tuple[list[str], str]:
    found = []
    for m in UNIT_RE.finditer(rest):
        code = m.group(1).lower()
        # Нормализуем код
        for key in sorted(UNITS.keys(), key=len, reverse=True):
            if code == key.lower():
                found.append(key)
                break
    rest_clean = UNIT_RE.sub('', rest)
    rest_clean = re.sub(r'\s+', ' ', rest_clean).strip()
    return found, rest_clean

# ── Извлечение должности ──────────────────────────────────────────────────

def extract_role(rest: str) -> tuple[str | None, str]:
    rest_lower = rest.lower().strip()
    # Убираем мусор: запятые, точки с запятой, цифры-номера рот/взводов
    rest_lower = re.sub(r'[,;]\s*', ' ', rest_lower)
    rest_lower = re.sub(r'\s+\d+\s+', ' ', rest_lower)
    rest_lower = rest_lower.strip()

    for abbr in sorted(ROLES.keys(), key=len, reverse=True):
        if rest_lower.startswith(abbr) or f' {abbr}' in rest_lower:
            return ROLES[abbr], ''
    # Если ничего не нашли — возвращаем как есть (для диагностики)
    return rest_lower if rest_lower else None, ''

# ── Разбор нескольких должностей (через ";") ─────────────────────────────

def parse_multi_role(rest: str) -> list[dict]:
    """Разбивает по ';' и парсит каждую часть."""
    parts = re.split(r';', rest)
    roles = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        units, part_clean = extract_units(part)
        role_text, _ = extract_role(part_clean)
        unit_codes = units if units else [None]
        for uc in unit_codes:
            roles.append({
                'role':      role_text,
                'unit_code': uc,
                'unit_name': UNITS.get(uc) if uc else None,
            })
    return roles

# ── Главная функция парсинга одной строки ────────────────────────────────

def parse_line(raw_line: str) -> dict | None:
    line = normalize(raw_line)

    # Дата гибели
    line, death_date = extract_death(line)

    # ФИО
    last, first, patr, rest = extract_fio(line)
    if not last:
        return None

    # Год рождения
    birth_year, rest = extract_birth_year(rest)

    # Звание
    rank, rest = extract_rank(rest)

    # Несколько должностей?
    if ';' in rest:
        roles = parse_multi_role(rest)
    else:
        units, rest2 = extract_units(rest)
        role_text, _ = extract_role(rest2)
        if units:
            roles = [{'role': role_text, 'unit_code': u, 'unit_name': UNITS.get(u)} for u in units]
        else:
            roles = [{'role': role_text, 'unit_code': None, 'unit_name': None}]

    return {
        'last_name':  last,
        'first_name': first,
        'patronymic': patr,
        'birth_year': birth_year,
        'death_date': death_date,
        'rank':       rank,
        'roles':      roles,
        'raw':        raw_line,  # оригинальная строка для проверки
    }

# ── CLI ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Парсинг алфавитного списка офицеров 52-й дивизии (без API)'
    )
    parser.add_argument('--file',       required=True, help='Путь к .docx файлу')
    parser.add_argument('--dry-run',    action='store_true', help='Показать первые 20 записей')
    parser.add_argument('--output',     metavar='FILE', help='Сохранить результат в JSON')
    parser.add_argument('--limit',      type=int, default=0, help='Обработать только первые N строк')
    parser.add_argument('--show-errors',action='store_true', help='Показать строки которые не распарсились')
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"Файл не найден: {args.file}")
        sys.exit(1)

    print(f"Читаю {args.file}...")
    lines = read_docx(args.file)
    lines = clean_lines(lines)
    print(f"Строк после очистки: {len(lines)}")

    if args.limit:
        lines = lines[:args.limit]
        print(f"Ограничение: первые {args.limit} строк")

    # Парсинг
    results = []
    errors  = []
    for line in lines:
        rec = parse_line(line)
        if rec:
            results.append(rec)
        else:
            errors.append(line)

    print(f"\nУспешно распарсено: {len(results)}")
    print(f"Не распарсилось:    {len(errors)}")

    # Статистика
    with_role  = sum(1 for r in results if any(x['role']      for x in r['roles']))
    with_unit  = sum(1 for r in results if any(x['unit_code'] for x in r['roles']))
    with_rank  = sum(1 for r in results if r['rank'])
    with_birth = sum(1 for r in results if r['birth_year'])
    with_death = sum(1 for r in results if r['death_date'])

    print(f"\nСтатистика:")
    print(f"  С должностью:      {with_role}")
    print(f"  С подразделением:  {with_unit}")
    print(f"  Со званием:        {with_rank}")
    print(f"  С годом рождения:  {with_birth}")
    print(f"  С датой гибели:    {with_death}")

    if args.dry_run:
        print("\n── Первые 20 записей ──")
        for r in results[:20]:
            print(json.dumps(r, ensure_ascii=False, indent=2))

    if args.show_errors:
        print(f"\n── Строки без ФИО ({len(errors)}) ──")
        for e in errors[:30]:
            print(f"  {e}")

    if args.output:
        # Убираем поле 'raw' из финального JSON
        clean = [{k: v for k, v in r.items() if k != 'raw'} for r in results]
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(clean, f, ensure_ascii=False, indent=2)
        print(f"\nJSON сохранён: {args.output}")

if __name__ == '__main__':
    main()
