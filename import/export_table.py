"""
export_table.py — общая таблица Excel: все люди сайта, строка на человека.

Решено 24.11.2026: вместо листов автора и сверки по ФИО — одна таблица,
выгруженная из базы. У каждой строки «Код» — номер человека на сайте
(persons.id): по нему правки автора ложатся на нужного человека без
угадывания по фамилии. Новых автор вносит строками без кода, с пометкой.

«Состав» — командный, политический, сержантский или рядовой
(composition.py, те же правила, что у фильтров сайта).

Национальности в базе нет — она берётся из таблицы автора (--list), где
человек находится тем же сопоставлением, что в import_komsostav.py.

Запуск:
  python export_table.py --dump ../../backups/52sd_2026-11-24_posle.sql.gz \
      --list ../data/Командный_состав_52сд-в-04.xlsx --out ../data/Общая_таблица.xlsx
"""

import argparse
import re
from collections import defaultdict
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import composition
import import_komsostav as ik

# Часть — так, как пишет автор: «429 сп», «1028 ап».
UNIT_SHORT = {name: re.sub(r'^(\d+)', r'\1 ', code) for code, name in ik.UNIT_NAMES.items()}

COLUMNS = [  # (заголовок, ширина)
    ('Код', 7), ('Состав', 14), ('Фамилия', 16), ('Имя', 13), ('Отчество', 16),
    ('Дата рождения', 12), ('Год рождения', 8), ('Национальность', 12),
    ('Звание', 18), ('Должность', 40), ('Часть', 9),
    ('Дата выбытия', 12), ('Причина выбытия', 22), ('Судьба', 16), ('Подробности судьбы', 28),
    ('Награды', 24), ('Прочие сведения', 28), ('Другие службы', 40), ('Фото', 6),
    ('Пометка', 30),
]


def ru_date(iso):
    if not iso:
        return None
    y, m, d = iso[:10].split('-')
    return f'{d}.{m}.{y}'


def fate(p):
    """На сайте «умер от ран» и «умер от болезни» — один вариант
    (died_of_wounds), что именно — в подробностях. В таблице — как у автора."""
    if p['outcome'] == 'died_of_wounds' and (p['outcome_details'] or '').startswith('умер от болезни'):
        return 'умер от болезни'
    return ik.OUTCOME_RU.get(p['outcome'], p['outcome'])


def filled(s):
    return sum(1 for f in ('rank', 'role', 'unit_id', 'date_to', 'end_reason', 'note') if s[f])


def service_text(s, unit_by_id):
    unit = unit_by_id.get(s['unit_id'])
    parts = [s['rank'], s['role'], UNIT_SHORT.get(unit, unit), ru_date(s['date_to'])]
    return ', '.join(x for x in parts if x)


def nationalities(list_path, persons):
    """Национальность из таблицы автора → id человека."""
    rows, questions, spellings = ik.load_list(list_path)
    rows = ik.apply_answers(rows, questions, spellings, [])
    rows = ik.merge_sheets(rows, [])
    wb = openpyxl.load_workbook(list_path, read_only=True, data_only=True)
    nat_by_code = {}
    sheets = [('Командный состав', '')] + [(s, f'G{d}_') for s, d in ik.GLORY_SHEETS.items()] \
        + [(s, f'{p}_') for s, p in ik.PEOPLE_SHEETS.items()]
    for sheet, prefix in sheets:
        if sheet not in wb.sheetnames:
            continue
        for xl_row, values in enumerate(wb[sheet].iter_rows(min_row=2, values_only=True), 2):
            nat = ik.clean(values[9]) if len(values) > 9 else None   # столбец J
            if nat:
                code = xl_row if not prefix else f'{prefix}{xl_row}'
                nat_by_code[code] = nat
    result, _ = ik.match(rows, persons)
    out = {}
    for r in rows:
        kind, cands, _ = result[r['xl_row']]
        nat = nat_by_code.get(r['xl_row'])
        # «Похожее написание» с одним кандидатом — это те, чьё написание
        # автор поправил в ответах (Джараян → Джароян): для национальности
        # этого достаточно.
        if nat and (kind == 'match' or (kind == 'disputed' and len(cands) == 1)):
            out.setdefault(cands[0]['id'], nat)
    return out


def main():
    ap = argparse.ArgumentParser(description='Общая таблица Excel из базы сайта')
    ap.add_argument('--dump', required=True, help='дамп базы (.sql.gz)')
    ap.add_argument('--list', help='таблица автора — для национальности')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    tables = {}
    import gzip
    with gzip.open(args.dump, 'rt', encoding='utf-8') as f:
        cur = None
        for line in f:
            if cur is None:
                m = re.match(r'COPY public\.(\w+) \((.*)\) FROM stdin;', line)
                if m:
                    cur, cols = m.group(1), m.group(2).split(', ')
                    tables[cur] = []
                continue
            if line.startswith('\\.'):
                cur = None
                continue
            vals = [None if v == '\\N' else v for v in line.rstrip('\n').split('\t')]
            tables[cur].append(dict(zip(cols, vals)))

    persons, records, units, _ = ik.load_from_dump(args.dump)
    unit_by_id = {uid: name for name, uid in units.items()}
    recs = defaultdict(list)
    for s in records:
        recs[s['person_id']].append(s)
    awards = defaultdict(list)
    for a in tables.get('person_awards', []):
        awards[int(a['person_id'])].append(a)
    photo = {int(p['id']): p['photo_url'] for p in tables['persons']}
    nat = nationalities(args.list, persons) if args.list else {}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Все люди'
    ws.append([c for c, _ in COLUMNS])
    for p in sorted(persons, key=lambda p: (ik.norm(p['last_name']), ik.norm(p['first_name']),
                                            ik.norm(p['patronymic']), p['id'])):
        rs = sorted(recs[p['id']], key=lambda s: (-filled(s), s['id']))
        main_s = rs[0] if rs else {}
        aw = awards[p['id']]
        birth = p['birth_date']
        unit = unit_by_id.get(main_s.get('unit_id'))
        ws.append([
            p['id'], composition.of_person(recs[p['id']]), p['last_name'], p['first_name'], p['patronymic'],
            ru_date(birth) if birth and not birth.endswith('-01-01') else None,
            int(birth[:4]) if birth else None,
            nat.get(p['id']),
            main_s.get('rank'), main_s.get('role'), UNIT_SHORT.get(unit, unit),
            ru_date(main_s.get('date_to')), main_s.get('end_reason'),
            fate(p), p['outcome_details'],
            '; '.join(a['title'] + (' (представлен)' if a['status'] == 'nominated' else '')
                      for a in aw) or None,
            main_s.get('note'),
            '; '.join(service_text(s, unit_by_id) for s in rs[1:]) or None,
            'есть' if photo.get(p['id']) else None,
            None,
        ])

    head_fill = PatternFill('solid', fgColor='DDDDDD')
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(wrap_text=True, vertical='center')
    yellow = PatternFill('solid', fgColor='FFF2CC')
    grey = PatternFill('solid', fgColor='EDEDED')
    n = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=n):
        row[0].fill = grey                      # код — не трогать
        row[-1].fill = yellow                   # пометка автора
    for i, (_, w) in enumerate(COLUMNS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = ws.dimensions
    dv = DataValidation(type='list', formula1='"' + ','.join(composition.COMPOSITIONS) + '"',
                        allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(f'B2:B{n + 2000}')                   # и для новых строк ниже

    info = wb.create_sheet('Как заполнять', 0)
    text = [
        ('Общая таблица 52-й стрелковой дивизии — все люди сайта 52sd.ru', True),
        (f'Выгружено из базы сайта {date.today().strftime("%d.%m.%Y")}. Людей: {n - 1}.', False),
        ('', False),
        ('Как пользоваться', True),
        ('1. Одна строка — один человек. Править прямо в клетках: ФИО, год, звание, должность, часть, судьба…', False),
        ('2. Серый столбец «Код» — номер человека на сайте. НЕ МЕНЯТЬ И НЕ УДАЛЯТЬ: по нему правка попадает к нужному человеку.', False),
        ('3. Новый человек — новая строка в конце листа, «Код» пустой; в «Пометке» — «новый».', False),
        ('4. Жёлтый столбец «Пометка» — для слов: «новый», «удалить», «дубль кода 123», «проверить», любые пояснения.', False),
        ('5. Строки не удалять — вместо этого написать «удалить» в «Пометке».', False),
        ('6. «Другие службы» — если у человека несколько должностей: основная — в столбцах «Звание/Должность/Часть», остальные здесь. Исправления к ним — словами в «Пометке».', False),
        ('7. «Награды» и «Фото» — для сведения; заводятся на сайте отдельно.', False),
        ('', False),
        ('Состав', True),
        ('Командный, политический, сержантский или рядовой — считается по званию и должности. '
         'Политический: политрук, комиссар, замполитрука, а также офицер на политической должности '
         '(замполит, парторг, комсорг, военком, агитатор, политотдел). Составы не пересекаются. '
         'Если состав определён неверно или пуст — исправьте прямо в клетке.', False),
        ('', False),
    ]
    counts = defaultdict(int)
    for row in ws.iter_rows(min_row=2, values_only=True):
        counts[row[1] or 'не определён'] += 1
    keys = composition.COMPOSITIONS + ['не определён']
    text += [(f'{s}: {counts.get(s, 0)}', False) for s in keys]
    for line, bold in text:
        info.append([line])
        if bold:
            info.cell(info.max_row, 1).font = Font(bold=True, size=12)
    info.column_dimensions['A'].width = 120

    wb.save(args.out)
    print(f'Людей: {n - 1}; с национальностью: {len(nat)}; файл: {args.out}')
    for s in keys:
        print(f'  {s}: {counts.get(s, 0)}')


if __name__ == '__main__':
    main()
