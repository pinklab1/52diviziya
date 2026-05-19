"""
import_to_db.py — загрузка result_fixed.json в PostgreSQL

Запуск на сервере:
  cd /opt/52diviziya
  apt-get install -y python3-pip
  pip install psycopg2-binary python-dotenv --break-system-packages
  python3 import_to_db.py [result_fixed.json] [--reset]

  --reset  очищает persons и service_records (TRUNCATE ... RESTART
           IDENTITY CASCADE) перед загрузкой. Нужен при ПОВТОРНОМ
           импорте: без него существующие люди пропускаются как дубли
           и обновлённые данные (год рождения, исправленные роли) в
           базу не попадают.
"""

import json
import os
import sys
from pathlib import Path

import psycopg2
from dotenv import load_dotenv

load_dotenv()

# ── Подразделения: код → (номер, название) ────────────────────────────────
UNIT_MAP = {
    '429сп':    (429,  '429-й стрелковый полк'),
    '429':      (429,  '429-й стрелковый полк'),
    '431сп':    (431,  '431-й стрелковый полк'),
    '431':      (431,  '431-й стрелковый полк'),
    '439сп':    (439,  '439-й стрелковый полк'),
    '439':      (439,  '439-й стрелковый полк'),
    '1028ап':   (1028, '1028-й артиллерийский полк'),
    '1028':     (1028, '1028-й артиллерийский полк'),
    '52сд':     (52,   'Штаб 52-й стрелковой дивизии'),
    '52':       (52,   'Штаб 52-й стрелковой дивизии'),
    '106':      (106,  '106-е подразделение'),
    '164осапб': (164,  '164-й отдельный сапёрный батальон'),
    '164':      (164,  '164-й отдельный сапёрный батальон'),
    '405оиптд': (405,  '405-й истребительно-противотанковый дивизион'),
    '405':      (405,  '405-й истребительно-противотанковый дивизион'),
    '127':      (127,  '127-е подразделение'),
    '201':      (201,  '201-е подразделение'),
}


def get_db():
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


def get_or_create_unit(cur, unit_code, division_id):
    if not unit_code or unit_code not in UNIT_MAP:
        return None
    _, name = UNIT_MAP[unit_code]
    cur.execute("SELECT id FROM units WHERE name = %s AND division_id = %s",
                (name, division_id))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute(
        "INSERT INTO units (division_id, name, unit_type) "
        "VALUES (%s, %s, %s) RETURNING id",
        (division_id, name, 'regiment')
    )
    return cur.fetchone()[0]


def main():
    args = sys.argv[1:]
    reset = '--reset' in args
    positional = [a for a in args if not a.startswith('--')]
    json_file = positional[0] if positional else 'result_fixed.json'

    if not Path(json_file).exists():
        print(f"Файл не найден: {json_file}")
        sys.exit(1)

    print(f"Читаю {json_file}...")
    with open(json_file, encoding='utf-8') as f:
        records = json.load(f)
    print(f"Записей в файле: {len(records)}")

    print("Подключаюсь к БД...")
    db = get_db()
    cur = db.cursor()

    if reset:
        print("--reset: очищаю persons и service_records...")
        cur.execute(
            "TRUNCATE persons, service_records RESTART IDENTITY CASCADE"
        )
        db.commit()
        print("Таблицы очищены.")

    # division_id
    cur.execute("SELECT id FROM divisions LIMIT 1")
    division_id = cur.fetchone()[0]
    print(f"Дивизия: id={division_id}")

    inserted = 0
    skipped = 0
    roles_added = 0

    for i, r in enumerate(records, 1):
        if not r.get('last_name'):
            skipped += 1
            continue

        # Проверка дубля
        cur.execute("""
            SELECT id FROM persons
            WHERE last_name = %s
              AND COALESCE(first_name,'') = COALESCE(%s,'')
              AND COALESCE(patronymic,'')  = COALESCE(%s,'')
        """, (r['last_name'], r.get('first_name'), r.get('patronymic')))
        if cur.fetchone():
            skipped += 1
            continue

        # Дата гибели
        death_date = r.get('death_date')

        # Год рождения → дата (1 января N-го года); схема хранит birth_date
        birth_year = r.get('birth_year')
        birth_date = f'{birth_year}-01-01' if birth_year else None

        # Вставка человека
        cur.execute("""
            INSERT INTO persons (last_name, first_name, patronymic, birth_date, death_date)
            VALUES (%s, %s, %s, %s, %s)
            RETURNING id
        """, (
            r['last_name'],
            r.get('first_name'),
            r.get('patronymic'),
            birth_date,
            death_date,
        ))
        person_id = cur.fetchone()[0]

        # Записи о службе
        for role_rec in (r.get('roles') or []):
            unit_id = get_or_create_unit(cur, role_rec.get('unit_code'), division_id)
            cur.execute("""
                INSERT INTO service_records (person_id, unit_id, division_id, rank, role)
                VALUES (%s, %s, %s, %s, %s)
            """, (
                person_id,
                unit_id,
                division_id,
                r.get('rank'),
                role_rec.get('role'),
            ))
            roles_added += 1

        inserted += 1

        if i % 100 == 0:
            print(f"  обработано {i}/{len(records)}...")
            db.commit()

    db.commit()
    cur.close()
    db.close()

    print(f"\n✓ Готово")
    print(f"  людей добавлено:   {inserted}")
    print(f"  записей о службе:  {roles_added}")
    print(f"  пропущено дублей:  {skipped}")


if __name__ == '__main__':
    main()
