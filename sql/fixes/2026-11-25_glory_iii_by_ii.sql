-- Кавалерам ордена Славы II степени — орден Славы III степени (25.11.2026).
--
-- Орденом Славы награждали по порядку: III, потом II, потом I. На сайте у
-- 72 из 74 кавалеров II степени III не было: в таблице автора они только на
-- листе «Кавалеры_II». Правило автора: II степень — значит, и III.
--
-- Источник — «по старшей степени» (как IMPLIED_SOURCE в import_komsostav.py):
-- видно, что награда добавлена по правилу, а не взята из листа.
-- Повторный запуск ничего не задвоит (ON CONFLICT DO NOTHING).
--
-- Запуск на сервере из /opt/52diviziya:
--   docker compose exec -T db sh -c 'psql -U "$POSTGRES_USER" -d "$POSTGRES_DB"' \
--     < sql/fixes/2026-11-25_glory_iii_by_ii.sql

BEGIN;

INSERT INTO person_awards (person_id, kind, degree, title, status, source)
SELECT DISTINCT a.person_id, 'order_of_glory', 3, 'Орден Славы III степени', 'awarded',
       'по старшей степени'
  FROM person_awards a
 WHERE a.kind = 'order_of_glory'
   AND a.degree IN (1, 2)
   AND a.status = 'awarded'
ON CONFLICT (person_id, kind, coalesce(degree, 0), status) DO NOTHING;

-- I степень — значит, и II (в базе таких пока нет, правило на будущее).
INSERT INTO person_awards (person_id, kind, degree, title, status, source)
SELECT DISTINCT a.person_id, 'order_of_glory', 2, 'Орден Славы II степени', 'awarded',
       'по старшей степени'
  FROM person_awards a
 WHERE a.kind = 'order_of_glory'
   AND a.degree = 1
   AND a.status = 'awarded'
ON CONFLICT (person_id, kind, coalesce(degree, 0), status) DO NOTHING;

SELECT count(DISTINCT person_id) FILTER (WHERE degree = 2) AS "кавалеров II",
       count(*) FILTER (WHERE degree = 3 AND source = 'по старшей степени') AS "III добавлено по правилу",
       count(DISTINCT person_id) AS "кавалеров всего"
  FROM person_awards WHERE kind = 'order_of_glory';

COMMIT;
