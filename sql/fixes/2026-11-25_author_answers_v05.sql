-- Ответы автора на «Вопросы_автору_в-05» (25.11.2026).
--
-- 1. Кувшинов Н. А. (2147) — 1921: на сайте уже так.
-- 2. Филиппов Г. С. (2407) — 1919: на сайте уже так.
-- 3. Мезенцев А. А. — один человек: радист старший, 587 обс, 1909 г. р.
--    Остаётся запись 2319 (с орденом Славы III), запись 2632 (429 сп,
--    из «Рабочих войны») удаляется, её источники переходят к 2319.
-- 4. Полные кавалеры ордена Славы — все три степени:
--    Атаев Пулат (1648, 127 орр), Бугаев Семен Васильевич (1678, 439 сп),
--    Катечкин Иван Егорович (1788, 1028 ап). На сайте орденов у них не было.
--
-- Каждая правка проверяет ФИО по коду — на чужого человека не ляжет.
-- Повторный запуск ничего не задвоит.
--
-- Запуск на сервере из /opt/52diviziya:
--   docker compose exec -T db sh -c 'psql -U "$POSTGRES_USER" -d "$POSTGRES_DB"' \
--     < sql/fixes/2026-11-25_author_answers_v05.sql

BEGIN;

-- 3. Мезенцев
UPDATE persons k SET
       birth_date = COALESCE(k.birth_date, DATE '1909-01-01'),
       sources = (SELECT jsonb_agg(DISTINCT x)
                    FROM jsonb_array_elements(k.sources || COALESCE(
                         (SELECT d.sources FROM persons d WHERE d.id = 2632
                             AND d.last_name = 'Мезенцев' AND d.first_name = 'Александр'), '[]')) x),
       last_updated = NOW()
 WHERE k.id = 2319 AND k.last_name = 'Мезенцев' AND k.first_name = 'Александр';

UPDATE service_records SET role = 'радист старший'
 WHERE person_id = 2319 AND role = 'радист'
   AND EXISTS (SELECT 1 FROM persons WHERE id = 2319 AND last_name = 'Мезенцев');

DELETE FROM persons
 WHERE id = 2632 AND last_name = 'Мезенцев' AND first_name = 'Александр'
   AND EXISTS (SELECT 1 FROM persons WHERE id = 2319 AND last_name = 'Мезенцев');

-- 4. Полные кавалеры: I, II и III степени
INSERT INTO person_awards (person_id, kind, degree, title, status, source)
SELECT p.id, 'order_of_glory', d.degree,
       'Орден Славы ' || d.roman || ' степени', 'awarded', 'ответ автора 25.11.2026'
  FROM persons p
 CROSS JOIN (VALUES (1, 'I'), (2, 'II'), (3, 'III')) AS d(degree, roman)
 WHERE (p.id, p.last_name, p.first_name) IN
       ((1648, 'Атаев', 'Пулат'), (1678, 'Бугаев', 'Семен'), (1788, 'Катечкин', 'Иван'))
ON CONFLICT (person_id, kind, coalesce(degree, 0), status) DO NOTHING;

-- Проверка
SELECT p.id, p.last_name, p.first_name, p.birth_date,
       (SELECT string_agg(a.degree::text, ',' ORDER BY a.degree) FROM person_awards a
         WHERE a.person_id = p.id AND a.kind = 'order_of_glory') AS "степени Славы",
       (SELECT string_agg(s.role, '; ') FROM service_records s WHERE s.person_id = p.id) AS "должность"
  FROM persons p
 WHERE p.id IN (2319, 2632, 1648, 1678, 1788)
 ORDER BY p.id;

COMMIT;
