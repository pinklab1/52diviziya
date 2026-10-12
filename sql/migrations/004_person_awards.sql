-- Миграция 004: награды человека — отдельной таблицей
-- Запуск: docker compose exec -T db psql -U diviziya52_user -d diviziya52 < sql/migrations/004_person_awards.sql
--         затем docker compose restart postgrest (обновить схему API)
--
-- Почему не persons.awards (text[]): у награды есть статус. «Представлен к
-- награде, но не награждён» — тоже факт биографии, и на сайте он должен
-- отличаться от полученной награды (контур звезды вместо закрашенной).
-- Плюс степень ордена и картинка наградного листа. В массив строк это не
-- ложится, а отдельная таблица позволяет и фильтр «Герои Советского Союза».
-- persons.awards остаётся как был (пустой у всех) — не используется.

CREATE TABLE IF NOT EXISTS person_awards (
  id            serial PRIMARY KEY,
  person_id     integer NOT NULL REFERENCES persons(id) ON DELETE CASCADE,
  -- Вид награды определяет значок на сайте. Новые виды — дописать сюда
  -- и в frontend/index.html (AWARDS).
  kind          text NOT NULL CHECK (kind IN (
                  'hero_ussr',        -- Герой Советского Союза — звезда
                  'order_of_glory',   -- орден Славы, степени I–III
                  'other'             -- прочие, показываются по title
                )),
  degree        smallint CHECK (degree BETWEEN 1 AND 3),  -- степень ордена
  title         text NOT NULL,      -- как показывать: «Орден Славы III степени»
  status        text NOT NULL DEFAULT 'awarded' CHECK (status IN (
                  'awarded',          -- награждён — закрашенный значок
                  'nominated'         -- представлен, не награждён — контур
                )),
  document_url  text,               -- картинка наградного листа: /photos/awards/…
  note          text,               -- «представление отклонено командармом» и т.п.
  source        text,               -- откуда сведения: «Кавалеры_III», ЦАМО …
  created_at    timestamptz DEFAULT now()
);

-- Одна и та же награда у человека — один раз (повторный импорт не дублирует).
CREATE UNIQUE INDEX IF NOT EXISTS person_awards_once
  ON person_awards (person_id, kind, coalesce(degree, 0), status);
CREATE INDEX IF NOT EXISTS idx_person_awards_kind ON person_awards (kind, status);

GRANT SELECT ON person_awards TO web_anon;
