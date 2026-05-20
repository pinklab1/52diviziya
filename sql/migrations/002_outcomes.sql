-- Миграция 002: добавляем колонки для судьбы человека и источников данных
-- Запуск: docker compose exec -T db psql -U diviziya52_user -d diviziya52 < sql/migrations/002_outcomes.sql

ALTER TABLE persons
    ADD COLUMN IF NOT EXISTS outcome TEXT
        CHECK (outcome IN (
            'killed',           -- погиб в бою
            'died_of_wounds',   -- умер от ран/болезни
            'missing',          -- пропал без вести
            'pow',              -- попал в плен (судьба неизвестна)
            'pow_survived',     -- был в плену, выжил
            'executed',         -- расстрелян
            'survived',         -- жив (пережил войну)
            'unknown'           -- сведения не сохранились
        )),
    ADD COLUMN IF NOT EXISTS outcome_details TEXT,    -- "шталаг IX A", "в составе 9-го ОШБ"
    ADD COLUMN IF NOT EXISTS sources JSONB DEFAULT '[]'::jsonb,
    ADD COLUMN IF NOT EXISTS last_updated TIMESTAMP DEFAULT NOW();

CREATE INDEX IF NOT EXISTS idx_persons_outcome ON persons(outcome);
