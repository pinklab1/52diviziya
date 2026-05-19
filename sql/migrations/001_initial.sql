-- ============================================================
-- 52-я стрелковая дивизия — схема базы данных
-- migration: 001_initial.sql
-- ============================================================

-- Расширения
CREATE EXTENSION IF NOT EXISTS unaccent;  -- нормализация букв (е/ё, и/й)

-- Текстовый поиск с unaccent для лучшей работы с русским
CREATE TEXT SEARCH CONFIGURATION russian_unaccent (COPY = russian);
ALTER TEXT SEARCH CONFIGURATION russian_unaccent
  ALTER MAPPING FOR hword, hword_part, word WITH unaccent, russian_stem;


-- ============================================================
-- ДИВИЗИИ
-- ============================================================
CREATE TABLE divisions (
  id            serial PRIMARY KEY,
  name          text NOT NULL,
  short_name    text,
  formed_at     date,
  disbanded_at  date,
  description   text,
  created_at    timestamptz DEFAULT now()
);

INSERT INTO divisions (name, short_name, formed_at)
VALUES ('52-я стрелковая дивизия', '52 сд', '1941-06-01');


-- ============================================================
-- ПОДРАЗДЕЛЕНИЯ (иерархия через parent_id)
-- ============================================================
CREATE TABLE units (
  id            serial PRIMARY KEY,
  division_id   int REFERENCES divisions ON DELETE CASCADE,
  parent_id     int REFERENCES units,           -- полк → батальон → рота
  name          text NOT NULL,
  unit_type     text CHECK (unit_type IN (
                  'division','regiment','battalion',
                  'company','platoon','other')),
  formed_at     date,
  disbanded_at  date,
  note          text
);


-- ============================================================
-- ЛЮДИ — базовая сущность
-- ============================================================
CREATE TABLE persons (
  id            serial PRIMARY KEY,
  last_name     text NOT NULL,
  first_name    text,
  patronymic    text,
  birth_date    date,
  birth_place   text,
  death_date    date,
  death_place   text,
  biography     text,
  photo_url     text,                           -- /photos/officers/ivanov-np.jpg
  awards        text[],                         -- массив наград
  sources       text[],                         -- архивные ссылки
  created_at    timestamptz DEFAULT now(),

  -- полнотекстовый поиск
  search_vector tsvector GENERATED ALWAYS AS (
    to_tsvector('russian_unaccent',
      coalesce(last_name,'')    || ' ' ||
      coalesce(first_name,'')   || ' ' ||
      coalesce(patronymic,'')   || ' ' ||
      coalesce(birth_place,'')  || ' ' ||
      coalesce(biography,'')
    )
  ) STORED
);

CREATE INDEX persons_search_idx ON persons USING GIN(search_vector);
CREATE INDEX persons_last_name_idx ON persons (last_name);


-- ============================================================
-- СЛУЖБА — человек в должности в период
-- (один человек мог служить в разных подразделениях)
-- ============================================================
CREATE TABLE service_records (
  id            serial PRIMARY KEY,
  person_id     int NOT NULL REFERENCES persons ON DELETE CASCADE,
  unit_id       int REFERENCES units,
  division_id   int REFERENCES divisions,
  rank          text,                           -- «генерал-майор»
  role          text,                           -- «командир полка», «рядовой»
  date_from     date,
  date_to       date,
  end_reason    text,                           -- «ранен», «убит», «переведён»
  note          text
);

CREATE INDEX service_records_person_idx ON service_records (person_id);
CREATE INDEX service_records_unit_idx   ON service_records (unit_id);
CREATE INDEX service_records_dates_idx  ON service_records (date_from, date_to);


-- ============================================================
-- СОБЫТИЯ боевого пути
-- ============================================================
CREATE TABLE events (
  id            serial PRIMARY KEY,
  division_id   int REFERENCES divisions ON DELETE CASCADE,
  date_from     date NOT NULL,
  date_to       date,                           -- NULL если одна дата
  title         text NOT NULL,
  description   text,
  location      text,                           -- «Ельня, Смоленская обл.»
  lat           numeric(9,6),
  lng           numeric(9,6),
  event_type    text CHECK (event_type IN (
                  'battle','move','defense',
                  'offense','reorg','command','other')),
  is_key        boolean DEFAULT false,          -- ключевое сражение
  sources       text[],
  created_at    timestamptz DEFAULT now(),

  search_vector tsvector GENERATED ALWAYS AS (
    to_tsvector('russian_unaccent',
      coalesce(title,'')       || ' ' ||
      coalesce(description,'') || ' ' ||
      coalesce(location,'')
    )
  ) STORED
);

CREATE INDEX events_search_idx   ON events USING GIN(search_vector);
CREATE INDEX events_date_idx     ON events (date_from);
CREATE INDEX events_type_idx     ON events (event_type);
CREATE INDEX events_key_idx      ON events (is_key) WHERE is_key = true;


-- ============================================================
-- ПОТЕРИ по периодам
-- ============================================================
CREATE TABLE casualties (
  id            serial PRIMARY KEY,
  division_id   int REFERENCES divisions ON DELETE CASCADE,
  period_from   date NOT NULL,
  period_to     date NOT NULL,
  killed        int DEFAULT 0,
  wounded       int DEFAULT 0,
  missing       int DEFAULT 0,
  captured      int DEFAULT 0,
  note          text,

  CONSTRAINT casualties_dates CHECK (period_to >= period_from)
);

CREATE INDEX casualties_period_idx ON casualties (period_from, period_to);


-- ============================================================
-- СВЯЗИ событий с людьми
-- (кто участвовал, кто был ранен, кто упомянут в приказе)
-- ============================================================
CREATE TABLE person_events (
  id            serial PRIMARY KEY,
  person_id     int NOT NULL REFERENCES persons ON DELETE CASCADE,
  event_id      int NOT NULL REFERENCES events  ON DELETE CASCADE,
  role          text,   -- «командовал», «ранен», «награждён», «погиб»
  note          text,
  UNIQUE (person_id, event_id, role)
);

CREATE INDEX person_events_person_idx ON person_events (person_id);
CREATE INDEX person_events_event_idx  ON person_events (event_id);


-- ============================================================
-- ДОКУМЕНТЫ И ФОТОГРАФИИ
-- ============================================================
CREATE TABLE documents (
  id            serial PRIMARY KEY,
  person_id     int REFERENCES persons ON DELETE SET NULL,
  event_id      int REFERENCES events  ON DELETE SET NULL,
  title         text,
  doc_type      text CHECK (doc_type IN (
                  'photo','order','map','report',
                  'award','letter','other')),
  file_url      text NOT NULL,
  taken_at      date,
  description   text,
  created_at    timestamptz DEFAULT now()
);

CREATE INDEX documents_person_idx ON documents (person_id);
CREATE INDEX documents_event_idx  ON documents (event_id);


-- ============================================================
-- РОЛЬ ДЛЯ POSTGREST (только чтение)
-- ============================================================
CREATE ROLE web_anon NOLOGIN;
GRANT USAGE  ON SCHEMA public TO web_anon;
GRANT SELECT ON ALL TABLES IN SCHEMA public TO web_anon;


-- ============================================================
-- УДОБНЫЕ ПРЕДСТАВЛЕНИЯ для API
-- ============================================================

-- Командиры с периодами службы (для боковой панели)
CREATE VIEW v_commanders AS
SELECT
  p.id,
  p.last_name,
  p.first_name,
  p.patronymic,
  p.photo_url,
  p.awards,
  sr.rank,
  sr.role,
  sr.date_from,
  sr.date_to,
  sr.end_reason,
  u.name AS unit_name,
  u.unit_type
FROM persons p
JOIN service_records sr ON sr.person_id = p.id
LEFT JOIN units u ON u.id = sr.unit_id
ORDER BY sr.date_from;

GRANT SELECT ON v_commanders TO web_anon;

-- События с количеством участников (для временной шкалы)
CREATE VIEW v_events_summary AS
SELECT
  e.*,
  COUNT(DISTINCT pe.person_id) AS persons_count
FROM events e
LEFT JOIN person_events pe ON pe.event_id = e.id
GROUP BY e.id
ORDER BY e.date_from;

GRANT SELECT ON v_events_summary TO web_anon;
