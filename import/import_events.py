"""
import_events.py — загрузка events.json в таблицу events

Запуск:
  python3 import_events.py
  python3 import_events.py --file events.json --clear
"""

import argparse
import json
import os
import sys
from pathlib import Path

import psycopg2
from dotenv import load_dotenv

load_dotenv()

# Известные координаты основных мест боевого пути
LOCATION_COORDS = {
    'Ржев':       (56.2629, 34.3289),
    'Старица':    (56.5142, 34.9337),
    'Ельня':      (54.5763, 33.1737),
    'Калинин':    (56.8587, 35.9176),  # Тверь
    'Москва':     (55.7558, 37.6173),
    'Курск':      (51.7307, 36.1937),
    'Орёл':       (52.9685, 36.0693),
    'Орел':       (52.9685, 36.0693),
    'Воронеж':    (51.6608, 39.2003),
    'Полунино':   (56.3500, 34.4500),
    'Галахово':   (56.3700, 34.4800),
    'Кишинев':    (47.0105, 28.8638),
    'Кишинёв':    (47.0105, 28.8638),
    'Яссы':       (47.1585, 27.6014),
    'Бухарест':   (44.4268, 26.1025),
    'София':      (42.6977, 23.3219),
    'Шумен':      (43.2706, 26.9229),
    'Белград':    (44.7866, 20.4489),
    'Будапешт':   (47.4979, 19.0402),
    'Балатон':    (46.8473, 17.7330),
    'Вена':       (48.2082, 16.3738),
    'Прага':      (50.0755, 14.4378),
    'Хинган':     (47.5000, 121.0000),  # Большой Хинган
    'Маньчжурия': (49.5973, 117.4254),
    'Молдавия':   (47.4116, 28.3699),
    'Румыния':    (45.9432, 24.9668),
    'Болгария':   (42.7339, 25.4858),
    'Югославия':  (44.0165, 21.0059),
    'Венгрия':    (47.1625, 19.5033),
    'Австрия':    (47.5162, 14.5501),
    'Чехословакия': (49.8175, 15.4730),
}


def get_db():
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


def get_coords(location_str):
    """Возвращает (lat, lng) по первому совпадению места."""
    if not location_str:
        return None, None
    for place, coords in LOCATION_COORDS.items():
        if place in location_str:
            return coords
    return None, None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', default='events.json')
    parser.add_argument('--clear', action='store_true',
                        help='Очистить таблицу events перед импортом')
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"Файл не найден: {args.file}")
        sys.exit(1)

    print(f"Читаю {args.file}...")
    with open(args.file, encoding='utf-8') as f:
        events = json.load(f)
    print(f"Событий в файле: {len(events)}")

    db = get_db()
    cur = db.cursor()

    cur.execute("SELECT id FROM divisions LIMIT 1")
    division_id = cur.fetchone()[0]

    if args.clear:
        print("Очищаю таблицу events...")
        cur.execute("TRUNCATE events RESTART IDENTITY CASCADE")

    inserted = 0
    for e in events:
        lat, lng = get_coords(e.get('location'))
        cur.execute("""
            INSERT INTO events
              (division_id, date_from, date_to, title, description,
               location, lat, lng, event_type, is_key)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            division_id,
            e['date_from'],
            e.get('date_to'),
            e['title'],
            e.get('description'),
            e.get('location'),
            lat,
            lng,
            e.get('event_type', 'other'),
            bool(e.get('is_key', False)),
        ))
        inserted += 1

    db.commit()
    cur.close()
    db.close()

    print(f"\n✓ Импортировано: {inserted} событий")
    print(f"  с координатами на карте: примерно {sum(1 for e in events if get_coords(e.get('location'))[0])}")


if __name__ == '__main__':
    main()
