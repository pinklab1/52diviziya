"""
match_photos.py — привязка фотографий к людям в БД

Принимает корневую папку с фотографиями (внутри могут быть подпапки по подразделениям).
Для каждого файла:
  1. Парсит ФИО из имени
  2. Ищет человека в БД
  3. Если нашёл — копирует файл в /var/52diviziya/photos/officers/ и обновляет photo_url
  4. Если не нашёл — пишет в отчёт unmatched.txt

Использование:
  python3 match_photos.py --source /path/to/photos --dest /var/52diviziya/photos/officers --report unmatched.txt
  python3 match_photos.py --source ./photos --dry-run   # только посмотреть что найдётся
"""

import argparse
import os
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from collections import defaultdict

import psycopg2
from dotenv import load_dotenv

load_dotenv()


def get_db():
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


# Слова которые часто оказываются в именах файлов и НЕ являются частью ФИО
STOPWORDS_RE = re.compile(
    r'\b(?:Enhanced|Repaired|командир|начштаб\w*|комиссар|политрук|нач\w+|зам\w+|'
    r'пом\w+|инстр\w+|ПНШ\d*|интенд\w*|трибунал|редактор\w*|капельмейстер|'
    r'дивинженер|комарт|комвзв\w*|замредактора|переводчик|секретарь|радио\w*|'
    r'разв\w+|штаб\w*|тыла|тыл|кадров|связи|свзи|связь|капитан|майор|лейтенант|'
    r'подп-?к\w*|подполковник|полковник|техинтен\w*|сттех\w*|л-т|газета|ВЛКСМ|ДПК|'
    r'офсвязи|комдив\w*|замкомдив|компул\w*|пульрот\w*|сд|сп|ап|осапб|осанб|'
    r'оиптд|обс|орр|орхз|оатрп|управ\w*|штаба|штабарт\w*|резерва|резерв|'
    r'отдштаба|отделат\w*|отделат\w*|снаба|снаб|артснаб|ВТС|ОВС|ОЗПР|ВТП|х\b)',
    re.IGNORECASE
)


def split_camelcase(s):
    """ДемченкоАндрейСтепанович → Демченко Андрей Степанович"""
    return re.sub(r'(?<=[а-я])(?=[А-ЯЁ])', ' ', s)


def parse_filename(filename):
    """Имя файла → возможный (фамилия, имя, отчество)"""
    name = Path(filename).stem  # без расширения

    # Уберём (2), (3) - маркеры дублей
    name = re.sub(r'\s*\(\d+\)\s*$', '', name)
    # Уберём -Enhanced, -Repaired
    name = re.sub(r'-(?:Enhanced|Repaired|001)+', '', name, flags=re.IGNORECASE)
    # Уберём префикс 52_ или 52 в начале
    name = re.sub(r'^52[_\s]+', '', name)

    # Заменим подчёркивания на пробелы
    name = name.replace('_', ' ')

    # Удалим стопслова (должности, подразделения)
    name = STOPWORDS_RE.sub('', name)

    # Удалим годы (4 цифры подряд между 1850 и 1930)
    name = re.sub(r'\b(18|19)\d{2}\b', '', name)

    # Удалим прочие цифры и спецсимволы
    name = re.sub(r'[\d,;:\-./\\]', ' ', name)

    # Нормализуем пробелы
    name = ' '.join(name.split())

    # Возможно CamelCase — разделим
    parts = name.split()
    expanded = []
    for p in parts:
        expanded.extend(split_camelcase(p).split())
    parts = expanded

    # Фильтруем мусор (слишком короткие, не на кириллице)
    parts = [p for p in parts if len(p) >= 2 and re.search(r'[А-Яа-яЁё]', p)]

    # Капитализируем (Бубнов, не БУБНОВ или бубнов)
    parts = [p[0].upper() + p[1:].lower() if len(p) > 1 else p.upper() for p in parts]

    # Берём первые 3 значимых слова как ФИО
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    if len(parts) == 2:
        return parts[0], parts[1], None
    if len(parts) == 1:
        return parts[0], None, None
    return None, None, None


def find_person(cur, last, first, pat):
    """Ищет в БД сначала точно, потом мягко"""
    # Точное совпадение
    cur.execute("""
        SELECT id, last_name, first_name, patronymic FROM persons
        WHERE last_name = %s
          AND COALESCE(first_name,'') = COALESCE(%s,'')
          AND COALESCE(patronymic,'') = COALESCE(%s,'')
    """, (last, first, pat))
    row = cur.fetchone()
    if row:
        return row, 'exact'

    # Мягкое: только фамилия + имя (без отчества)
    if first:
        cur.execute("""
            SELECT id, last_name, first_name, patronymic FROM persons
            WHERE last_name = %s AND first_name = %s
        """, (last, first))
        rows = cur.fetchall()
        if len(rows) == 1:
            return rows[0], 'fuzzy_no_patronymic'

    # Мягкое: только фамилия (если уникальна)
    cur.execute("""
        SELECT id, last_name, first_name, patronymic FROM persons
        WHERE last_name = %s
    """, (last,))
    rows = cur.fetchall()
    if len(rows) == 1:
        return rows[0], 'fuzzy_last_only'

    return None, 'not_found'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--source', required=True, help='Папка с фото (рекурсивно)')
    parser.add_argument('--dest', default='/var/52diviziya/photos/officers',
                        help='Папка куда копировать (на сервере)')
    parser.add_argument('--report', default='unmatched.txt',
                        help='Файл для отчёта о ненайденных')
    parser.add_argument('--dry-run', action='store_true',
                        help='Только показать что найдётся, не копировать')
    args = parser.parse_args()

    source = Path(args.source)
    if not source.exists():
        print(f"Не существует: {source}")
        sys.exit(1)

    dest = Path(args.dest)
    if not args.dry_run:
        dest.mkdir(parents=True, exist_ok=True)

    db = get_db()
    cur = db.cursor()

    # Найдём все .jpg/.jpeg/.png файлы
    photos = []
    for ext in ('*.jpg', '*.JPG', '*.jpeg', '*.png'):
        photos.extend(source.rglob(ext))
    print(f"Найдено фотографий: {len(photos)}")

    stats = defaultdict(int)
    unmatched = []
    matched_log = []

    for photo in photos:
        last, first, pat = parse_filename(photo.name)
        if not last:
            stats['parse_failed'] += 1
            unmatched.append((photo.name, 'не удалось распарсить ФИО'))
            continue

        result, match_type = find_person(cur, last, first, pat)
        stats[match_type] += 1

        if not result:
            unmatched.append((photo.name, f'не найдено: {last} {first} {pat}'))
            continue

        person_id, db_last, db_first, db_pat = result
        new_filename = f"{person_id}_{db_last}.jpg"
        matched_log.append((photo.name, person_id, db_last, db_first, db_pat, match_type))

        if not args.dry_run:
            # Копируем файл
            shutil.copy2(photo, dest / new_filename)
            # Записываем относительный путь в photo_url
            photo_url = f"/photos/officers/{new_filename}"
            cur.execute(
                "UPDATE persons SET photo_url = %s WHERE id = %s",
                (photo_url, person_id)
            )

    if not args.dry_run:
        db.commit()

    # ── ОТЧЁТ ─────────────────────────────────────────────────────
    print("\n=== РЕЗУЛЬТАТЫ ===")
    for k, v in sorted(stats.items(), key=lambda x: -x[1]):
        print(f"  {k}: {v}")
    print(f"\nНе привязано: {len(unmatched)}")

    # Записываем подробный отчёт
    with open(args.report, 'w', encoding='utf-8') as f:
        f.write("=== ПРИВЯЗАНО ===\n")
        for fname, pid, l, fi, p, mt in matched_log:
            f.write(f"[{mt:25}] {fname} → #{pid} {l} {fi or ''} {p or ''}\n")
        f.write("\n=== НЕ ПРИВЯЗАНО ===\n")
        for fname, reason in unmatched:
            f.write(f"{fname} — {reason}\n")
    print(f"Отчёт сохранён: {args.report}")

    cur.close()
    db.close()


if __name__ == '__main__':
    main()
