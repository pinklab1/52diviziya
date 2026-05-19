"""
import_alfavit.py — парсинг алфавитного списка офицеров 52-й дивизии

Формат строк в файле (непоследовательный, несколько вариантов):
  Абаньшин_Семен_Григорьевич_1917_начштаб 431; командир 429, 431
  Алавидзе_Шалва_Багратович_1920_комср431сп_млл-т(+15.03.44)
  Багров Петр Деевич 1916_капитан_1028ап
  Бахметьева Ираида Александровна 106
  Грачева Надежда Ивановна_1921_млврач431сп_стл-тмс_Вихорева

Установка:
  pip install python-docx anthropic psycopg2-binary python-dotenv

Запуск:
  python import_alfavit.py --file АЛФАВИТ.doc --dry-run   # проверка
  python import_alfavit.py --file АЛФАВИТ.doc             # реальный импорт
"""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import anthropic
import psycopg2
from dotenv import load_dotenv

load_dotenv()

# ── Расшифровка сокращений ────────────────────────────────────────────────
# Используется в промпте как подсказка для Claude

ABBREVIATIONS = """
ДОЛЖНОСТИ (сокращения):
  комдив              — командир дивизии
  ком, командир       — командир (подразделения)
  комбат              — командир батальона
  комбатр             — командир батареи
  комвзв              — командир взвода
  комср, комроты      — командир стрелковой роты
  компульроты         — командир пулемётной роты
  комминроты          — командир миномётной роты
  комогнвзв           — командир огневого взвода
  комминвзв           — командир миномётного взвода
  комвзвсвязи         — командир взвода связи
  комсанвзв           — командир санитарного взвода
  начштаб, НШ         — начальник штаба
  замкомдив           — заместитель командира дивизии
  замполит            — заместитель командира по политической части
  ПНШ, ПНШ-1..6      — помощник начальника штаба
  военком             — военный комиссар
  парторг             — партийный организатор
  начразв             — начальник разведки
  начПФС              — начальник продовольственно-фуражной службы
  ОВСполка            — обозно-вещевая служба полка
  адст                — адъютант старший батальона, дивизиона
  комендантштабаупр   — комендант штаба управления
  комендант           — комендант
  офицсвязи           — офицер связи
  ветфельд            — ветеринарный фельдшер
  врач, млврач        — врач, младший врач
  фельд               — фельдшер
  хоз                 — хозяйственная часть / завхоз
  редакция            — редакция газеты
  снаб                — снабжение

ДОЛЖНОСТИ (расширенный список — расшифровывай ВСЕГДА точно, не подгоняй под «командир»):
  комтоповзв          — командир топографического взвода
  комтелефвзв         — командир взвода телефонной связи
  комтелкабвзв        — командир телефонно-кабельного взвода
  комштабвзвсвязи     — командир штабного взвода связи
  комрадиовзв         — командир радиовзвода
  комкомендвзв        — командир комендантского взвода
  комштаброты         — командир штабной роты
  комсанроты          — командир санитарной роты
  комотдзаградбат     — командир отдельного заградительного батальона
  комотдучроты        — командир отдельной учебной роты
  комэвакоотд         — командир эвакоотделения
  комтрансвзв, комтранроты — командир транспортного взвода / роты
  комарт              — командир артиллерии
  коморудия           — командир орудия
  комдвз              — командир дивизиона
  ПНШ, ПНШ-1..6       — помощник начальника штаба (номер = отдел)
  помПНШ              — помощник помощника начальника штаба
  замком, замкомполка — заместитель командира (полка/дивизиона)
  замкомпостр         — заместитель командира по строевой части
  помнач              — помощник начальника
  нач1отдштаба        — начальник 1-го отдела штаба
  начштаба            — начальник штаба
  делопроизштаба      — делопроизводитель штаба
  начарт, начартил    — начальник артиллерии
  начарттехснаб       — начальник арттехснабжения
  начартмаст          — начальник артмастерской
  начхим, начхимслуж  — начальник химической службы
  начсанслужбы        — начальник санитарной службы
  начсндив            — начальник снабжения дивизии
  начфин              — начальник финансовой части
  начкадров           — начальник отдела кадров
  начавтосл           — начальник автослужбы
  начзагототд         — начальник заготовительного отдела
  начинжпарка         — начальник инженерного парка
  начТС, начВТС       — начальник технической / ветеринарно-технической службы
  начотдтроф          — начальник отдела трофеев
  начПО, начполитотд  — начальник политотдела
  замначполитотд      — заместитель начальника политотдела
  инстрполитотдела    — инструктор политотдела
  стврач              — старший врач
  военфельд           — военный фельдшер
  ветврач, стветврач, млветврач — (старший/младший) ветеринарный врач
  фельдветер          — ветеринарный фельдшер
  начветлазар         — начальник ветеринарного лазарета
  ординатор           — ординатор
  терапевт            — врач-терапевт
  поммедснаб          — помощник по медицинскому снабжению
  арттехник           — артиллерийский техник
  автотехник          — автотехник
  сттехл-т, сттехнл-т — старший техник-лейтенант
  техник-интендант    — техник-интендант
  полкинженер, полкинж — полковой инженер
  дивинженер          — дивизионный инженер
  оперуполСМЕРШ, уполнОКРСмерш — оперуполномоченный (ОКР) СМЕРШ
  трибунал            — военный трибунал
  парткомиссия        — партийная комиссия
  секретарьДПК        — секретарь дивизионной парткомиссии
  секрбюро ВЛКСМ      — секретарь бюро ВЛКСМ
  наводчикоруд        — наводчик орудия
  капельмейстер       — капельмейстер
  начКлуб             — начальник клуба
  казначей            — казначей
  редактор, замредактора — редактор / заместитель редактора
  МТО, замМТО         — материально-техническое обеспечение / заместитель по МТО

ЗВАНИЯ:
  млл-т               — младший лейтенант
  стл-т               — старший лейтенант
  л-т                 — лейтенант
  гвстл-т             — гвардии старший лейтенант
  стл-тмс             — старший лейтенант медслужбы
  млл-тмс             — младший лейтенант медслужбы
  техник-лейтенант    — техник-лейтенант
  майоринтсл          — майор интендантской службы

ПОДРАЗДЕЛЕНИЯ:
  52 или 52сд         — штаб, политотдел, тыл дивизии
  429, 429сп          — 429-й стрелковый полк
  431, 431сп          — 431-й стрелковый полк
  439, 439сп          — 439-й стрелковый полк
  1028, 1028ап        — 1028-й артиллерийский полк
  106, 106мсб         — 106-й медико-санитарный батальон
  164, 164осапб       — 164-й отдельный сапёрный батальон
  127 орр             — 127-я отдельная разведывательная рота
  405, 405оиптд       — 405-й отдельный истребительно-противотанковый дивизион
  587 обс             — 587-й отдельный батальон связи
  527 оавр            — 527-я отдельная автомобильная рота подвоза
  42 орхз             — 42-я отдельная рота химической защиты
  845 ветлазарет      — 845-й ветеринарный лазарет
  371-я пх            — 371-я полевая хлебопекарня
"""

# ── Промпт для Claude ─────────────────────────────────────────────────────

PROMPT = f"""Ты помогаешь оцифровать алфавитный список офицеров 52-й стрелковой дивизии СССР (Великая Отечественная война).

Вот справочник сокращений из этого документа:
{ABBREVIATIONS}

Из текста ниже извлеки каждого человека как отдельный объект.
Верни ТОЛЬКО валидный JSON — массив объектов. Никаких пояснений, никакого markdown.

Правила разбора строк:
- Разделители: подчёркивание (_) и пробел взаимозаменяемы
- Год рождения: 4-значное число (1917, 1923 и т.д.)
- Дата гибели: в скобках со знаком плюс, например (+15.03.44) → смерть 15 марта 1944
- Несколько должностей/подразделений разделены точкой с запятой (;)
- Звание может стоять до или после должности
- Двойная фамилия: Бубнова_Гущина — девичья и замужняя
- Если подразделение только число (429, 431, 439, 52, 1028) — расшифруй по справочнику

Структура каждого объекта:
{{
  "last_name":   "фамилия, обязательно",
  "first_name":  "имя или null",
  "patronymic":  "отчество или null",
  "birth_year":  1917 или null,
  "death_date":  "YYYY-MM-DD или null",
  "rank":        "полное звание (расшифрованное) или null",
  "roles": [
    {{
      "role":       "расшифрованная должность",
      "unit_code":  "код подразделения: 429сп / 431сп / 439сп / 1028ап / 52сд / 106 / 164осапб / 405оиптд / 127 или null",
      "unit_name":  "полное название подразделения"
    }}
  ],
  "note": "всё что не вписалось в структуру, например девичья фамилия, или null"
}}

Текст:
"""

# ── Чтение файла ──────────────────────────────────────────────────────────

def read_doc(path: str) -> str:
    """Конвертирует .doc в .docx и извлекает текст."""
    p = Path(path)
    suffix = p.suffix.lower()

    if suffix == '.doc':
        print("  Конвертирую .doc → .docx...")
        with tempfile.TemporaryDirectory() as tmp:
            result = subprocess.run(
                ['soffice', '--headless', '--convert-to', 'docx',
                 str(p.resolve()), '--outdir', tmp],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                raise RuntimeError(f"Ошибка конвертации: {result.stderr}")
            docx_path = Path(tmp) / (p.stem + '.docx')
            return _extract_text(str(docx_path))
    elif suffix == '.docx':
        return _extract_text(path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {suffix}")


def _extract_text(docx_path: str) -> str:
    from docx import Document
    doc = Document(docx_path)
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(' | '.join(cells))
    return '\n'.join(lines)


def get_person_lines(text: str) -> list[str]:
    """Оставляет только строки с людьми, убирает заголовки букв."""
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Пропускаем заголовки вида ** А (45) ** и итоговую строку
        if re.match(r'^\*+\s*[А-ЯЁA-Z]', line):
            continue
        if re.match(r'^\d{3,4}\s*\(', line):
            continue
        lines.append(line)
    return lines


# ── Claude API ────────────────────────────────────────────────────────────

def extract_with_claude(lines: list[str], batch_size: int = 40) -> list[dict]:
    """Отправляет строки батчами в Claude, получает структурированный JSON."""
    client = anthropic.Anthropic()
    all_results = []
    total = len(lines)
    batches = [lines[i:i+batch_size] for i in range(0, total, batch_size)]

    for i, batch in enumerate(batches):
        chunk = '\n'.join(batch)
        print(f"  Батч {i+1}/{len(batches)} ({len(batch)} строк)...", end=' ', flush=True)

        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=4000,
                messages=[{"role": "user", "content": PROMPT + chunk}]
            )
            raw = response.content[0].text.strip()
            raw = re.sub(r'^```json\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw)

            data = json.loads(raw)
            if isinstance(data, list):
                all_results.extend(data)
                print(f"✓ {len(data)} записей")
            else:
                print("⚠ неожиданный формат")

        except json.JSONDecodeError as e:
            print(f"✗ ошибка JSON: {e}")
            # Сохраняем сырой ответ для диагностики
            with open(f'error_batch_{i+1}.txt', 'w', encoding='utf-8') as f:
                f.write(raw)
        except Exception as e:
            print(f"✗ ошибка API: {e}")

    return all_results


# ── Импорт в PostgreSQL ───────────────────────────────────────────────────

UNIT_MAP = {
    '429сп':    (429,  '429-й стрелковый полк'),
    '431сп':    (431,  '431-й стрелковый полк'),
    '439сп':    (439,  '439-й стрелковый полк'),
    '1028ап':   (1028, '1028-й артиллерийский полк'),
    '52сд':     (52,   'Штаб 52-й стрелковой дивизии'),
    '106':      (106,  '106-е подразделение'),
    '164осапб': (164,  '164-й отдельный сапёрный батальон'),
    '405оиптд': (405,  '405-й отдельный истребительно-противотанковый дивизион'),
    '127':      (127,  '127-е подразделение'),
}


def get_or_create_unit(cur, unit_code: str, division_id: int) -> int | None:
    """Возвращает id подразделения, создаёт если нет."""
    if not unit_code or unit_code not in UNIT_MAP:
        return None

    num, name = UNIT_MAP[unit_code]
    cur.execute("SELECT id FROM units WHERE name = %s AND division_id = %s", (name, division_id))
    row = cur.fetchone()
    if row:
        return row[0]

    cur.execute("""
        INSERT INTO units (division_id, name, unit_type)
        VALUES (%s, %s, %s) RETURNING id
    """, (division_id, name, 'regiment'))
    return cur.fetchone()[0]


def import_persons(records: list[dict], db, division_id: int = 1) -> tuple[int, int]:
    cur = db.cursor()
    inserted = 0
    skipped = 0

    for r in records:
        if not r.get('last_name'):
            skipped += 1
            continue

        # Проверка дубликата
        cur.execute("""
            SELECT id FROM persons
            WHERE last_name = %s
              AND (first_name = %s OR (first_name IS NULL AND %s IS NULL))
              AND (patronymic = %s OR (patronymic IS NULL AND %s IS NULL))
        """, (
            r['last_name'], r.get('first_name'), r.get('first_name'),
            r.get('patronymic'), r.get('patronymic')
        ))
        if cur.fetchone():
            skipped += 1
            continue

        # Дата гибели
        death_date = None
        if r.get('death_date'):
            death_date = _parse_date(r['death_date'])

        # Вставка персоны
        cur.execute("""
            INSERT INTO persons (last_name, first_name, patronymic, death_date, note)
            VALUES (%s, %s, %s, %s, %s)
            RETURNING id
        """, (
            r['last_name'],
            r.get('first_name'),
            r.get('patronymic'),
            death_date,
            r.get('note'),
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

        inserted += 1

    db.commit()
    cur.close()
    return inserted, skipped


# ── Вспомогательные функции ───────────────────────────────────────────────

def _parse_date(s: str | None) -> str | None:
    if not s:
        return None
    s = str(s).strip()
    # DD.MM.YY или DD.MM.YYYY
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$', s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = '19' + y
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"
    # YYYY-MM-DD
    if re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return s
    return None


def get_db():
    return psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'diviziya52'),
        user=os.getenv('POSTGRES_USER', 'diviziya52_user'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432'),
    )


# ── CLI ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Импорт алфавитного списка офицеров 52-й дивизии'
    )
    parser.add_argument('--file', required=True, help='Путь к .doc или .docx файлу')
    parser.add_argument('--dry-run', action='store_true',
                        help='Показать результат без записи в БД')
    parser.add_argument('--batch-size', type=int, default=40,
                        help='Строк за один запрос к Claude (по умолчанию 40)')
    parser.add_argument('--limit', type=int, default=0,
                        help='Обработать только первые N строк (для тестирования)')
    parser.add_argument('--output-json', metavar='FILE',
                        help='Сохранить результат в JSON-файл')
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"Файл не найден: {args.file}")
        sys.exit(1)

    # Чтение
    print(f"Читаю {args.file}...")
    text = read_doc(args.file)
    lines = get_person_lines(text)
    print(f"Найдено {len(lines)} строк с людьми")

    if args.limit:
        lines = lines[:args.limit]
        print(f"Ограничение: обрабатываю первые {args.limit} строк")

    # Claude
    print(f"Отправляю в Claude (батчи по {args.batch_size} строк)...")
    records = extract_with_claude(lines, batch_size=args.batch_size)
    print(f"Итого извлечено: {len(records)} записей")

    # Сохранить JSON если нужно
    if args.output_json:
        with open(args.output_json, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"JSON сохранён: {args.output_json}")

    # Dry run — показать первые 5
    if args.dry_run:
        print("\n── DRY RUN (первые 5 записей) ──")
        for r in records[:5]:
            print(json.dumps(r, ensure_ascii=False, indent=2))
        print(f"\nВсего записей: {len(records)}")
        return

    # Импорт в БД
    print("Подключаюсь к БД...")
    db = get_db()
    ins, skp = import_persons(records, db)
    db.close()
    print(f"\n✓ Готово: вставлено {ins}, пропущено дублей {skp}")


if __name__ == '__main__':
    main()
