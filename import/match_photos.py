"""
match_photos.py — привязка фотографий к людям в БД

Принимает корневую папку с фотографиями.
Для каждого файла:
  1. Парсит ФИО из имени
  2. Ищет человека в БД
  3. Если нашёл — копирует файл в /var/52diviziya/photos/officers/ и обновляет photo_url
  4. Если не нашёл — пишет в отчёт unmatched.txt

Использование:
  python3 match_photos.py --source ./photos --dry-run    # только посмотреть
  python3 match_photos.py --source ./photos              # реальная привязка
"""

import argparse
import os
import re
import shutil
import sys
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


# ════════════════════════════════════════════════════════════════════
# СТОП-СЛОВА: должности и подразделения которые НЕ являются частью ФИО.
#
# Принцип: разные категории слов вырезаются разными способами.
# - Длинные должности (от 5 букв): можно вырезать в любом месте — низкий риск задеть имя
# - Короткие аббревиатуры (сп, ап, сд, мм): только после цифры (429сп) или
#   между подчёркиваниями (_х_), чтобы не задеть "Спесивцев", "Уральский"
# - Сокращения званий с дефисом: л-т, к-н, мл-т, стл-т
# ════════════════════════════════════════════════════════════════════

# Длинные слова — должности и звания (≥5 букв). Безопасно убирать в любом месте.
LONG_STOPWORDS = [
    'Enhanced', 'Repaired',
    # Должности с "ком" в начале
    'командир', 'комиссар', 'комбат', 'комбатр', 'комбатареи', 'комбатарей',
    'комвзв', 'комминрот', 'комминбатр', 'комминвзв', 'комогнвзв', 'комогн',
    'компульрот', 'компульвзв', 'компул', 'компультвзв', 'комроты', 'комарт',
    'комотд', 'комсан', 'комсамобот', 'коморудия', 'кмвзв',
    'комсорг',
    # С "зам"
    'замкомдив', 'замкомандира', 'замкомполка', 'замредактора', 'замполит', 'замполитсб',
    # С "нач"
    'начштаба', 'начштаб', 'начштабарт', 'начпо', 'начхим', 'начарт', 'начкадров',
    'начтыла', 'начсвязи', 'начотделен', 'начмастбоепит', 'начразв', 'начразвштаба',
    'начполитотд',
    # С "пом"
    'помначПО', 'помнач', 'помначПОпо', 'помначартснаб',
    'помначотделен', 'помначотделат', 'помначВТС',
    # Прочие должности
    'политрук', 'политотдела', 'политотд', 'политрукроты', 'политруквзвпешразведки',
    'парторг', 'парторгсб', 'парткомиссия', 'парторгрот',
    'дивинженер', 'трибунал', 'капельмейстер', 'переводчик', 'секретарь',
    'инструктор', 'интендант', 'интенд', 'редактор',
    'адъютанткомдива', 'офсвязи', 'офсвязиштабартил',
    'техинтен', 'сттех', 'стветфлд', 'стветфлддивизиона',
    'фельдлазар', 'фельдшер', 'радиокинотехник', 'радиостанции',
    'военком', 'военкомсб', 'военфельдшер',
    'завделопроиз', 'делопроиз', 'мастбоепит', 'наводчикоруд',
    'госпитвзв', 'санинст', 'санитар',
    'пешразв', 'пешразведки', 'разведки',
    # Звания (длинные — 5+ букв)
    'подполковник', 'полковник', 'майор', 'лейтенант', 'капитан',
    'старшина', 'старшинамс', 'сержант',
    'подполковника',
    # Прочее (4+ букв)
    'инстрполитотдела', 'инстрполитотд', 'инструкторполитотдела',
    'связи', 'свзи', 'штаба', 'штабартил', 'штабарт',
    'отдштаба', 'отдельной', 'отделат',
    'тыла', 'кадров', 'резерва',
    'газета', 'ВЛКСМ', 'ДПК', 'НШСБ',
    'боепит', 'снаба', 'артснаб', 'управления',
    'минроты', 'политотдел',
    'врач', 'сестра', 'начштаб',
]

# Сортируем по убыванию длины (длинные слова матчатся раньше)
LONG_STOPWORDS.sort(key=len, reverse=True)
# Используем lookbehind+lookahead чтобы матчить только когда слово окружено разделителями
# То есть "комвзв" в "Комвзв_..." матчится, но не в "Овсейчук"
LONG_STOPWORDS_RE = re.compile(
    r'(?<![а-яёА-ЯЁa-zA-Z])(?:' + '|'.join(re.escape(w) for w in LONG_STOPWORDS) + r')(?:[а-яёА-ЯЁa-zA-Z0-9]*)',
    re.IGNORECASE
)

# Короткие аббревиатуры подразделений: ТОЛЬКО после цифр
DIGIT_UNIT_RE = re.compile(
    r'\d+\s*(?:сд|сп|ап|осапб|осанб|оиптд|обс|орр|орхз|оатрп|пхп|двзн|сб|ср|мсб|мм)\b',
    re.IGNORECASE
)

# Короткие токены между разделителями (нет букв ни до ни после)
BETWEEN_TOKENS_RE = re.compile(
    r'(?<![а-яёА-ЯЁa-zA-Z])(?:х|пхп|ПХП|двзн|сб|ср|сд|сп|ап|обс|орр|орхз|осапб|осанб|оиптд|оатрп|'
    r'мм|мс|нт|вет|ст|стс|001|002|003|ОВС|ВТС|ОЗПР|ВТП|ПТР)(?![а-яёА-ЯЁa-zA-Z])',
    re.IGNORECASE
)

# ПНШ-X
PNSH_RE = re.compile(r'\bПНШ[-]?\d*', re.IGNORECASE)

# Сокращения званий с дефисом
RANK_DASH_RE = re.compile(
    r'(?<![а-яёА-ЯЁ])(?:мл?л?-?т|с-?т|стл-?т|к-н|подп-?к|гв\.?[а-я]*|стс-?т|гвстаршина|млл-?т|л-т)(?![а-яёА-ЯЁ])',
    re.IGNORECASE
)


def split_camelcase(s):
    """ДемченкоАндрейСтепанович → Демченко Андрей Степанович"""
    return re.sub(r'(?<=[а-я])(?=[А-ЯЁ])', ' ', s)


def parse_filename(filename):
    """Имя файла → (фамилия, имя, отчество)"""
    name = Path(filename).stem

    # Уберём маркер дубля
    name = re.sub(r'\s*\(\d+\)\s*', ' ', name)
    # Уберём -Enhanced, -Repaired, -001
    name = re.sub(r'-(?:Enhanced|Repaired|\d{3})', ' ', name, flags=re.IGNORECASE)
    # Уберём префикс типа "52_", "429_", "1028_"
    name = re.sub(r'^\d+[_\s]+', '', name)

    # Удалим аббревиатуры подразделений с цифрой ПЕРЕД именем
    name = DIGIT_UNIT_RE.sub(' ', name)
    # Удалим короткие токены между разделителями
    name = BETWEEN_TOKENS_RE.sub(' ', name)
    # ПНШ-X
    name = PNSH_RE.sub(' ', name)
    # Удалим длинные стоп-слова
    name = LONG_STOPWORDS_RE.sub(' ', name)
    # Сокращения званий
    name = RANK_DASH_RE.sub(' ', name)

    # Заменим разделители на пробелы
    name = name.replace('_', ' ')

    # Удалим годы
    name = re.sub(r'\b(?:18|19)\d{2}\b', ' ', name)

    # Удалим остатки цифр и спецсимволов
    name = re.sub(r'[\d,;:./\\()]+', ' ', name)
    name = re.sub(r'\s*-\s*', ' ', name)  # дефисы между словами тоже считаем разделителями

    # Нормализуем пробелы
    name = ' '.join(name.split())

    # CamelCase
    parts = []
    for p in name.split():
        parts.extend(split_camelcase(p).split())

    # Фильтр: только кириллические токены длиной >= 2
    parts = [p for p in parts if len(p) >= 2 and re.search(r'[А-Яа-яЁё]', p)]
    # Капитализация
    parts = [p[0].upper() + p[1:].lower() if len(p) > 1 else p.upper() for p in parts]

    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    if len(parts) == 2:
        return parts[0], parts[1], None
    if len(parts) == 1:
        return parts[0], None, None
    return None, None, None


def find_person(cur, last, first, pat):
    """Ищет: точно → без отчества → только фамилия"""
    cur.execute("""
        SELECT id, last_name, first_name, patronymic FROM persons
        WHERE last_name = %s
          AND COALESCE(first_name,'') = COALESCE(%s,'')
          AND COALESCE(patronymic,'') = COALESCE(%s,'')
    """, (last, first, pat))
    row = cur.fetchone()
    if row:
        return row, 'exact'

    if first:
        cur.execute("""
            SELECT id, last_name, first_name, patronymic FROM persons
            WHERE last_name = %s AND first_name = %s
        """, (last, first))
        rows = cur.fetchall()
        if len(rows) == 1:
            return rows[0], 'fuzzy_no_patronymic'

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
    parser.add_argument('--source', required=True)
    parser.add_argument('--dest', default='/var/52diviziya/photos/officers')
    parser.add_argument('--report', default='unmatched.txt')
    parser.add_argument('--dry-run', action='store_true')
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

    photos = []
    for ext in ('*.jpg', '*.JPG', '*.jpeg', '*.png'):
        photos.extend(source.rglob(ext))
    print(f"Найдено фотографий: {len(photos)}")

    stats = defaultdict(int)
    unmatched = []
    matched = []

    for photo in photos:
        last, first, pat = parse_filename(photo.name)
        if not last:
            stats['parse_failed'] += 1
            unmatched.append((photo.name, 'не удалось распарсить ФИО'))
            continue

        result, match_type = find_person(cur, last, first, pat)
        stats[match_type] += 1

        if not result:
            unmatched.append((photo.name, f'не найдено: {last} {first or ""} {pat or ""}'))
            continue

        person_id, db_last, db_first, db_pat = result
        new_filename = f"{person_id}_{db_last}.jpg"
        matched.append((photo.name, person_id, db_last, db_first, db_pat, match_type))

        if not args.dry_run:
            shutil.copy2(photo, dest / new_filename)
            photo_url = f"/photos/officers/{new_filename}"
            cur.execute(
                "UPDATE persons SET photo_url = %s WHERE id = %s",
                (photo_url, person_id)
            )

    if not args.dry_run:
        db.commit()

    print("\n=== РЕЗУЛЬТАТЫ ===")
    for k, v in sorted(stats.items(), key=lambda x: -x[1]):
        print(f"  {k}: {v}")
    print(f"\nНе привязано: {len(unmatched)}")

    with open(args.report, 'w', encoding='utf-8') as f:
        f.write("=== ПРИВЯЗАНО ===\n")
        for fname, pid, l, fi, p, mt in matched:
            f.write(f"[{mt:25}] {fname} → #{pid} {l} {fi or ''} {p or ''}\n")
        f.write("\n=== НЕ ПРИВЯЗАНО ===\n")
        for fname, reason in unmatched:
            f.write(f"{fname} — {reason}\n")
    print(f"Отчёт сохранён: {args.report}")

    cur.close()
    db.close()


if __name__ == '__main__':
    main()
