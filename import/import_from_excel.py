"""
import_from_excel.py — идемпотентный импорт людей из Excel-таблиц

Принципы:
- НЕ удаляет существующие данные (если только с --clear)
- НЕ затирает заполненные поля пустыми
- Обогащает существующих людей: новые данные дописываются к старым
- Отслеживает источники в колонке sources

Запуск:
  python3 import_from_excel.py --file poteri.xlsx
  python3 import_from_excel.py --file poteri.xlsx --clear   # для разработки
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
import psycopg2
from psycopg2.extras import Json
from dotenv import load_dotenv

load_dotenv()


UNIT_NAMES = {
    '429 сп':     '429-й стрелковый полк',
    '431 сп':     '431-й стрелковый полк',
    '439 сп':     '439-й стрелковый полк',
    '1028 ап':    '1028-й артиллерийский полк',
    '52 сд':      'Управление 52-й стрелковой дивизии',
    '106 осанб':  '106-й отдельный санитарный батальон',
    '164 осапб':  '164-й отдельный сапёрный батальон',
    '405 оиптд':  '405-й истребительно-противотанковый дивизион',
    '587 обс':    '587-й отдельный батальон связи',
    '127 орр':    '127-я отдельная разведрота',
    '42 орхз':    '42-я рота химзащиты',
    '527 оатрп':  '527-й отдельный артиллерийский транспортный полк',
    '842 ДВЛ':    '842-й дивизионный ветеринарный лазарет',
    '371 ПХП':    '371-й полевой хлебозавод',
    '563 орс':    '563-я отдельная рота связи',
    'политотдел': 'Политотдел дивизии',
}


def get_db():
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


# ── Парсеры ─────────────────────────────────────────────────────────

def parse_fio(raw):
    """ФИО → (last, first, patronymic, has_photo, photo_hint)"""
    if pd.isna(raw):
        return None, None, None, False, None
    s = str(raw).strip()
    has_photo = '_' in s
    photo_hint = s if has_photo else None

    if has_photo:
        s = s.replace('_', ' ')
        s = re.sub(r'\b(19|18)\d{2}\b', '', s)
        s = re.sub(r'\bх\b', '', s)
        s = re.sub(r'ПНШ-?\d+', '', s, flags=re.IGNORECASE)
        s = re.sub(r'\d+\s*(сп|ап|осапб|осанб|оиптд|обс|орр|орхз|оатрп)', '', s)
        s = re.sub(r'(ОВСполка|начарт|снаб|зам\w*|политотд\w*)', '',
                   s, flags=re.IGNORECASE)

    s = ' '.join(s.split())
    parts = s.split()
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2], has_photo, photo_hint
    if len(parts) == 2:
        return parts[0], parts[1], None, has_photo, photo_hint
    if len(parts) == 1:
        return parts[0], None, None, has_photo, photo_hint
    return None, None, None, has_photo, photo_hint


def parse_year(val):
    if pd.isna(val): return None
    try:
        year = int(float(val))
        if 1850 <= year <= 1930:
            return f"{year}-01-01"
    except (ValueError, TypeError):
        pass
    return None


def parse_date_ru(val):
    if pd.isna(val): return None
    s = str(val).strip()
    m = re.match(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*г?', s)
    if m:
        d, mo, yy = m.groups()
        yy = int(yy)
        year = yy if yy > 1900 else (1900 + yy if yy > 30 else 2000 + yy)
        try:
            datetime(year, int(mo), int(d))
            return f"{year}-{int(mo):02d}-{int(d):02d}"
        except ValueError:
            return None
    return None


def is_real_name(x):
    if pd.isna(x): return False
    s = str(x).strip()
    return len(s) >= 5 and any(c.isalpha() for c in s) and not s.isdigit()


def classify_note(note):
    """Примечание → (outcome, details)
    Возвращает None если примечание не содержит info об исходе."""
    if pd.isna(note): return None, None
    s = str(note).strip().lower()
    if not s: return None, None

    if 'жив' in s and 'плен' in s:
        return 'pow_survived', str(note).strip()
    if s == 'жив' or s.startswith('жив '):
        return 'survived', str(note).strip()
    if 'расстрелян' in s:
        return 'executed', str(note).strip()
    if 'плен' in s:
        return 'pow', str(note).strip()
    if 'штрафн' in s or 'ошб' in s:
        return None, str(note).strip()  # это контекст, не исход
    return None, None


# ── Работа с БД ────────────────────────────────────────────────────

def get_or_create_unit(cur, raw_unit, division_id, _cache={}):
    if pd.isna(raw_unit): return None
    s = str(raw_unit).strip()
    if not s: return None
    if s in _cache: return _cache[s]

    short = None
    for code in UNIT_NAMES:
        if code in s:
            short = code
            break
    name = UNIT_NAMES.get(short, s) if short else s

    cur.execute("SELECT id FROM units WHERE name = %s", (name,))
    row = cur.fetchone()
    if row:
        _cache[s] = row[0]
        return row[0]

    cur.execute(
        "INSERT INTO units (division_id, name, unit_type) VALUES (%s, %s, %s) RETURNING id",
        (division_id, name, 'unit')
    )
    uid = cur.fetchone()[0]
    _cache[s] = uid
    return uid


def find_person(cur, last, first, pat):
    cur.execute("""
        SELECT id, birth_date, death_date, photo_url, outcome, sources
        FROM persons
        WHERE last_name = %s
          AND COALESCE(first_name,'') = COALESCE(%s,'')
          AND COALESCE(patronymic,'')  = COALESCE(%s,'')
    """, (last, first, pat))
    return cur.fetchone()


def add_source(existing_sources, new_source):
    """Добавить источник в список (без дублей)"""
    if not existing_sources:
        sources = []
    elif isinstance(existing_sources, list):
        sources = list(existing_sources)
    else:
        sources = []
    if new_source not in sources:
        sources.append(new_source)
    return sources


def upsert_person(cur, data, source_tag, stats):
    """Создаёт нового или обогащает существующего.
    data — dict с полями last_name, first_name, patronymic, birth_date,
           death_date, photo_url, outcome, outcome_details
    """
    existing = find_person(cur, data['last_name'], data.get('first_name'),
                           data.get('patronymic'))

    if existing is None:
        # Новый человек
        sources = [source_tag]
        cur.execute("""
            INSERT INTO persons (last_name, first_name, patronymic, birth_date,
                                 death_date, photo_url, outcome, outcome_details,
                                 sources, last_updated)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, NOW())
            RETURNING id
        """, (
            data['last_name'], data.get('first_name'), data.get('patronymic'),
            data.get('birth_date'), data.get('death_date'), data.get('photo_url'),
            data.get('outcome'), data.get('outcome_details'),
            Json(sources),
        ))
        stats['inserted'] += 1
        return cur.fetchone()[0]

    # Обогащаем существующего
    pid, old_birth, old_death, old_photo, old_outcome, old_sources = existing
    updates = []
    params = []

    if data.get('birth_date') and not old_birth:
        updates.append("birth_date = %s")
        params.append(data['birth_date'])

    if data.get('death_date') and not old_death:
        updates.append("death_date = %s")
        params.append(data['death_date'])

    if data.get('photo_url') and not old_photo:
        updates.append("photo_url = %s")
        params.append(data['photo_url'])

    if data.get('outcome') and not old_outcome:
        updates.append("outcome = %s")
        params.append(data['outcome'])

    if data.get('outcome_details'):
        updates.append("outcome_details = COALESCE(outcome_details, '') || %s")
        params.append(data['outcome_details'])

    # Обновляем sources всегда
    new_sources = add_source(old_sources, source_tag)
    updates.append("sources = %s")
    params.append(Json(new_sources))

    updates.append("last_updated = NOW()")

    params.append(pid)
    cur.execute(f"UPDATE persons SET {', '.join(updates)} WHERE id = %s", params)
    stats['updated'] += 1
    return pid


# ── Главная функция ────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', default='poteri.xlsx')
    parser.add_argument('--clear', action='store_true',
                        help='DEV: очистить persons, service_records, units перед импортом')
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"Не найден: {args.file}")
        sys.exit(1)

    print(f"Читаю {args.file}...")
    all_sheets = pd.read_excel(args.file, sheet_name=None)

    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT id FROM divisions LIMIT 1")
    division_id = cur.fetchone()[0]

    if args.clear:
        print("⚠ Очищаю persons, service_records, units...")
        cur.execute("TRUNCATE persons, service_records, units RESTART IDENTITY CASCADE")
        db.commit()

    stats = {'inserted': 0, 'updated': 0}

    # ═══ ЭТАП 1: АЛФАВИТ ФОТО ═══
    print("\n=== ЭТАП 1: АЛФАВИТ ФОТО ===")
    df = all_sheets.get('АЛФАВИТ ФОТО')
    if df is not None:
        df = df[df['ФИО'].apply(is_real_name)]
        for _, row in df.iterrows():
            last, first, pat, has_photo, photo_hint = parse_fio(row['ФИО'])
            if not last: continue

            data = {
                'last_name': last,
                'first_name': first,
                'patronymic': pat,
                'birth_date': parse_year(row.get('Год рождения')),
                'photo_url': photo_hint if has_photo else None,
            }
            pid = upsert_person(cur, data, 'АЛФАВИТ ФОТО', stats)

            # Должность и подразделение (отдельной записью)
            unit_id = get_or_create_unit(cur, row.get('Подразделение'), division_id)
            role = row.get('Должность ')
            rank = row.get('Звание')
            if pd.isna(role): role = None
            if pd.isna(rank): rank = None

            if unit_id or role or rank:
                # Проверим что такой записи о службе ещё нет
                cur.execute("""
                    SELECT id FROM service_records
                    WHERE person_id = %s
                      AND COALESCE(unit_id, 0) = COALESCE(%s, 0)
                      AND COALESCE(role, '') = COALESCE(%s, '')
                """, (pid, unit_id, role))
                if not cur.fetchone():
                    cur.execute("""
                        INSERT INTO service_records (person_id, unit_id, division_id, rank, role)
                        VALUES (%s, %s, %s, %s, %s)
                    """, (pid, unit_id, division_id, rank, role))

        print(f"  обработано: добавлено {stats['inserted']}, обновлено {stats['updated']}")
    s1 = dict(stats)

    # ═══ ЭТАП 2: ЛИСТЫ ПОДРАЗДЕЛЕНИЙ ═══
    print("\n=== ЭТАП 2: Листы подразделений ===")
    loss_sheets = [n for n in all_sheets if n != 'АЛФАВИТ ФОТО']

    for sheet_name in loss_sheets:
        df = all_sheets[sheet_name]
        fio_col = 'Ф.И.О' if 'Ф.И.О' in df.columns else 'ФИО'
        if fio_col not in df.columns: continue

        df = df[df[fio_col].apply(is_real_name)]
        before_ins = stats['inserted']
        before_upd = stats['updated']

        for _, row in df.iterrows():
            last, first, pat, has_photo, photo_hint = parse_fio(row[fio_col])
            if not last: continue

            killed = parse_date_ru(row.get('Дата и причина выбытия'))
            died = parse_date_ru(row.get('Unnamed: 7'))
            missing = parse_date_ru(row.get('Unnamed: 8'))
            death_date = killed or died or missing

            # Базовый outcome по колонке смерти
            if killed: outcome = 'killed'
            elif died: outcome = 'died_of_wounds'
            elif missing: outcome = 'missing'
            else: outcome = None

            # Уточняющий outcome из примечания
            note_outcome, note_details = classify_note(row.get('Примечание'))
            if note_outcome:
                outcome = note_outcome  # примечание приоритетнее
            details = note_details

            data = {
                'last_name': last,
                'first_name': first,
                'patronymic': pat,
                'birth_date': parse_year(row.get('Год рождения')),
                'death_date': death_date,
                'photo_url': photo_hint if has_photo else None,
                'outcome': outcome,
                'outcome_details': details,
            }

            pid = upsert_person(cur, data, f'ПОТЕРИ:{sheet_name}', stats)

            # Служебная запись
            unit_raw = row.get('Подразделение')
            if pd.isna(unit_raw):
                unit_raw = sheet_name
            unit_id = get_or_create_unit(cur, unit_raw, division_id)

            role = row.get('Должность')
            rank = row.get('Звание')
            if pd.isna(role): role = None
            if pd.isna(rank): rank = None

            if unit_id or role or rank:
                cur.execute("""
                    SELECT id FROM service_records
                    WHERE person_id = %s
                      AND COALESCE(unit_id, 0) = COALESCE(%s, 0)
                      AND COALESCE(role, '') = COALESCE(%s, '')
                """, (pid, unit_id, role))
                if not cur.fetchone():
                    cur.execute("""
                        INSERT INTO service_records (person_id, unit_id, division_id, rank, role)
                        VALUES (%s, %s, %s, %s, %s)
                    """, (pid, unit_id, division_id, rank, role))

        delta_ins = stats['inserted'] - before_ins
        delta_upd = stats['updated'] - before_upd
        print(f"  {sheet_name:<12} → +{delta_ins} новых, ~{delta_upd} обогащено")

    db.commit()

    # ═══ ИТОГИ ═══
    print("\n=== ИТОГО В БД ===")
    cur.execute("SELECT COUNT(*) FROM persons")
    print(f"  persons:           {cur.fetchone()[0]}")
    cur.execute("SELECT COUNT(*) FROM persons WHERE birth_date IS NOT NULL")
    print(f"  с годом рождения:  {cur.fetchone()[0]}")
    cur.execute("SELECT COUNT(*) FROM persons WHERE photo_url IS NOT NULL")
    print(f"  с маркером фото:   {cur.fetchone()[0]}")
    cur.execute("SELECT COUNT(*) FROM persons WHERE death_date IS NOT NULL")
    print(f"  с датой смерти:    {cur.fetchone()[0]}")
    print("\nПо исходам:")
    cur.execute("""
        SELECT COALESCE(outcome, 'не указан'), COUNT(*)
        FROM persons GROUP BY outcome ORDER BY COUNT(*) DESC
    """)
    for outcome, cnt in cur.fetchall():
        print(f"  {outcome:<18} {cnt}")
    cur.execute("SELECT COUNT(*) FROM units")
    print(f"\n  units:             {cur.fetchone()[0]}")
    cur.execute("SELECT COUNT(*) FROM service_records")
    print(f"  service_records:   {cur.fetchone()[0]}")

    cur.close()
    db.close()


if __name__ == '__main__':
    main()
