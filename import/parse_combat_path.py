"""
parse_combat_path.py — извлекает события из книги о боевом пути 52-й СД

Использование:
  python parse_combat_path.py --file Боевой_путь_52_СД.docx --output events.json
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from collections import Counter

from docx import Document

MONTHS = {
    'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4,
    'мая': 5, 'май': 5, 'июн': 6, 'июл': 7, 'август': 8,
    'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12,
}

DATE_RE = re.compile(
    r'\b(\d{1,2})\s+(январ\w*|феврал\w*|март\w*|апрел\w*|ма[йя]|июн\w*|июл\w*|август\w*|сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\s*(\d{4})?',
    re.IGNORECASE
)
PERIOD_RE = re.compile(
    r'\b[Сс]?\s*(\d{1,2})\s*(?:по|–|-)\s*(\d{1,2})\s+(январ\w*|феврал\w*|март\w*|апрел\w*|ма[йя]|июн\w*|июл\w*|август\w*|сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\s*(\d{4})?',
    re.IGNORECASE
)

# Контекстные подсказки по годам для боевого пути 52-й сд
CONTEXT_HINTS = [
    ('Ржевско-Сычевск', 1942), ('Ржев', 1942), ('Полунино', 1942),
    ('Галахово', 1942), ('Курск', 1943), ('Курская', 1943),
    ('Кутузов', 1943), ('Орёл', 1943), ('Орел', 1943),
    ('Багратион', 1944), ('Ясско-Кишинев', 1944), ('Кишинев', 1944),
    ('Кишинёв', 1944), ('Молдавия', 1944), ('Румыния', 1944),
    ('Бухарест', 1944), ('Болгария', 1944), ('София', 1944),
    ('Шумен', 1944), ('Югослав', 1944), ('Белград', 1944),
    ('Венгрия', 1944), ('Будапешт', 1944),
    ('Балатон', 1945), ('Венская', 1945), ('Вена', 1945),
    ('Чехослов', 1945), ('Прага', 1945),
    ('Манчжур', 1945), ('Маньчжур', 1945), ('Хинган', 1945),
    ('Япон', 1945),
]

TYPE_KEYWORDS = {
    'liberation': ['освобожд', 'взят', 'овладел', 'занял', 'вступил в'],
    'battle': ['бои', 'бой ', 'наступлени', 'атак', 'штурм', 'контратак',
               'оборон', 'сражени', 'удар', 'противник', 'враг', 'плацдарм',
               'форсировал', 'прорвал', 'разгром'],
    'move': ['прибыл', 'передислоц', 'переброс', 'марш', 'выгруз', 'погруз',
             'эшелон', 'выступил', 'выдвиж', 'сосредоточ', 'переправ',
             'отошл', 'отступа'],
    'reorg': ['пополнен', 'переформ', 'отдых', 'тыл', 'резерв', 'доукомплект'],
    'award': ['награжд', 'орден', 'медал', 'присвоено', 'благодарност', 'геро'],
    'command': ['приказ', 'назначен', 'принял командование', 'звание'],
}

KNOWN_LOCATIONS = [
    'Ржев', 'Ельня', 'Старица', 'Калинин', 'Москва', 'Курск', 'Орёл', 'Орел',
    'Воронеж', 'Полунино', 'Галахово', 'Полтинино', 'Зеленкино', 'Дешевки',
    'Космариха', 'Кишинев', 'Кишинёв', 'Яссы', 'Бухарест', 'София', 'Белград',
    'Будапешт', 'Вена', 'Прага', 'Дунай', 'Тиса', 'Молдавия', 'Румыния',
    'Болгария', 'Югославия', 'Венгрия', 'Австрия', 'Чехословакия',
    'Шумен', 'Хинган', 'Маньчжурия', 'Балатон',
]

OPERATIONS = [
    'Ржевско-Сычевская', 'Ясско-Кишинев', 'Будапештск', 'Венск',
    'Балатонск', 'Маньчжурск', 'Кутузов', 'Багратион', 'Белградск',
]


def parse_month(s):
    s = s.lower()
    for key, num in MONTHS.items():
        if s.startswith(key):
            return num
    return None


def guess_year_from_text(text):
    for kw, year in CONTEXT_HINTS:
        if kw in text:
            return year
    return None


def detect_type(text):
    text_l = text.lower()
    scores = {t: 0 for t in TYPE_KEYWORDS}
    for t, words in TYPE_KEYWORDS.items():
        for w in words:
            if w in text_l:
                scores[t] += 1
    best = max(scores.items(), key=lambda x: x[1])
    return best[0] if best[1] > 0 else 'other'


def extract_locations(text):
    return [loc for loc in KNOWN_LOCATIONS if loc in text][:3]


def extract_operations(text):
    return [op for op in OPERATIONS if op in text]


def make_title(text, etype, locations, operations):
    if operations:
        return f'Операция «{operations[0]}»'
    text_l = text.lower()
    if etype == 'liberation' and locations:
        return f'Освобождение: {locations[0]}'
    if etype == 'battle':
        if 'штурм' in text_l:
            return f'Штурм {locations[0]}' if locations else 'Штурм'
        if 'оборон' in text_l:
            return f'Оборона {locations[0]}' if locations else 'Оборонительный бой'
        if 'наступлени' in text_l:
            return f'Наступление: {locations[0]}' if locations else 'Наступление'
        return f'Бой за {locations[0]}' if locations else 'Бой'
    if etype == 'move':
        if 'прибыл' in text_l and locations:
            return f'Прибытие: {locations[0]}'
        if 'переброс' in text_l or 'передислоц' in text_l:
            return 'Передислокация' + (f' в {locations[0]}' if locations else '')
        if 'марш' in text_l:
            return 'Марш' + (f' к {locations[0]}' if locations else '')
        return 'Перемещение' + (f': {locations[0]}' if locations else '')
    if etype == 'reorg':
        return 'Пополнение и отдых' + (f' в {locations[0]}' if locations else '')
    if etype == 'award':
        return 'Награждения и приказы'
    first = re.split(r'[.!?]', text)[0]
    return first[:80].strip() + ('...' if len(first) > 80 else '')


def build_year_anchors(paragraphs):
    """Определяет якорные годы для каждого параграфа."""
    anchors = {}
    for i, para in enumerate(paragraphs):
        year_m = re.search(r'\b(194[1-5])\b', para)
        if year_m:
            anchors[i] = int(year_m.group(1))
            continue
        hint = guess_year_from_text(para)
        if hint:
            anchors[i] = hint

    # Распространяем якоря на соседей
    final = {}
    anchor_keys = sorted(anchors.keys())
    for i in range(len(paragraphs)):
        if i in anchors:
            final[i] = anchors[i]
        else:
            before = [k for k in anchor_keys if k < i]
            after = [k for k in anchor_keys if k > i]
            if before and after:
                b_dist = i - before[-1]
                a_dist = after[0] - i
                final[i] = anchors[before[-1]] if b_dist <= a_dist else anchors[after[0]]
            elif before:
                final[i] = anchors[before[-1]]
            elif after:
                final[i] = anchors[after[0]]
            else:
                final[i] = 1942
    return final


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--file', required=True)
    p.add_argument('--output', default='events.json')
    args = p.parse_args()

    if not Path(args.file).exists():
        print(f"Файл не найден: {args.file}")
        sys.exit(1)

    print(f"Читаю {args.file}...")
    doc = Document(args.file)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    print(f"Параграфов: {len(paragraphs)}")

    print("Определяю годы по контексту...")
    para_years = build_year_anchors(paragraphs)

    print("Извлекаю события...")
    events = []
    for i, para in enumerate(paragraphs):
        if i < 10: continue
        if 'Страница' in para and len(para) < 30: continue

        year = para_years.get(i, 1942)
        found_dates = []

        for m in PERIOD_RE.finditer(para):
            d1, d2, mw, ys = m.group(1), m.group(2), m.group(3), m.group(4)
            month = parse_month(mw)
            if not month: continue
            actual_year = int(ys) if ys else year
            if not (1941 <= actual_year <= 1945): continue
            try:
                d1i, d2i = int(d1), int(d2)
                datetime(actual_year, month, d1i)
                datetime(actual_year, month, d2i)
                found_dates.append((
                    f"{actual_year}-{month:02d}-{d1i:02d}",
                    f"{actual_year}-{month:02d}-{d2i:02d}",
                ))
            except ValueError:
                continue

        for m in DATE_RE.finditer(para):
            day, mw, ys = m.group(1), m.group(2), m.group(3)
            month = parse_month(mw)
            if not month: continue
            actual_year = int(ys) if ys else year
            if not (1941 <= actual_year <= 1945): continue
            try:
                d = int(day)
                datetime(actual_year, month, d)
            except ValueError:
                continue
            date_str = f"{actual_year}-{month:02d}-{d:02d}"
            if any(pp[0] == date_str or pp[1] == date_str for pp in found_dates):
                continue
            found_dates.append((date_str, None))

        if not found_dates: continue

        etype = detect_type(para)
        locations = extract_locations(para)
        operations = extract_operations(para)
        title = make_title(para, etype, locations, operations)
        sentences = re.split(r'(?<=[.!?])\s+', para)
        description = ' '.join(sentences[:3])[:500]
        is_key = bool(operations) or 'герой' in para.lower() or 'благодарност' in para.lower()

        for date_from, date_to in found_dates:
            events.append({
                'date_from': date_from,
                'date_to': date_to,
                'title': title,
                'description': description,
                'location': ', '.join(locations) if locations else None,
                'event_type': etype,
                'is_key': is_key,
            })

    # Дедупликация
    seen = set()
    unique = []
    for e in events:
        key = (e['date_from'], e['date_to'], e['title'][:60])
        if key in seen: continue
        seen.add(key)
        unique.append(e)
    unique.sort(key=lambda e: e['date_from'])

    print(f"\n✓ Извлечено: {len(unique)} событий")
    print(f"По годам: {dict(Counter(e['date_from'][:4] for e in unique))}")
    print(f"По типам: {dict(Counter(e['event_type'] for e in unique))}")
    print(f"Ключевых: {sum(1 for e in unique if e['is_key'])}")

    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(unique, f, ensure_ascii=False, indent=2)
    print(f"Сохранено: {args.output}")


if __name__ == '__main__':
    main()
