CREATE SCHEMA IF NOT EXISTS quality_bot;

CREATE TABLE IF NOT EXISTS quality_bot.users (
    id BIGSERIAL PRIMARY KEY,
    telegram_id BIGINT NOT NULL UNIQUE,
    full_name VARCHAR(255) NOT NULL,
    city VARCHAR(255) NOT NULL DEFAULT '',
    is_whitelisted BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE quality_bot.users
    ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE quality_bot.users
    ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

CREATE TABLE IF NOT EXISTS quality_bot.tests (
    id BIGSERIAL PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    test_date DATE NOT NULL,
    active BOOLEAN NOT NULL DEFAULT TRUE,
    approved BOOLEAN NOT NULL DEFAULT FALSE,
    approved_at TIMESTAMPTZ,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_quality_bot_tests_test_date UNIQUE (test_date)
);

ALTER TABLE quality_bot.tests
    ADD COLUMN IF NOT EXISTS approved_at TIMESTAMPTZ;

ALTER TABLE quality_bot.tests
    ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE quality_bot.tests
    ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

CREATE TABLE IF NOT EXISTS quality_bot.questions (
    id BIGSERIAL PRIMARY KEY,
    test_id BIGINT NOT NULL REFERENCES quality_bot.tests(id) ON DELETE CASCADE,
    test_type VARCHAR(50) NOT NULL DEFAULT 'control',
    question_order INTEGER NOT NULL,
    image_file_id VARCHAR(500) NOT NULL,
    correct_answer INTEGER NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_quality_bot_questions_order UNIQUE (test_id, test_type, question_order)
);

ALTER TABLE quality_bot.questions
    DROP CONSTRAINT IF EXISTS ck_quality_bot_questions_correct_answer;

ALTER TABLE quality_bot.questions
    ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE quality_bot.questions
    ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

CREATE TABLE IF NOT EXISTS quality_bot.test_sessions (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT NOT NULL REFERENCES quality_bot.users(id) ON DELETE CASCADE,
    test_id BIGINT NOT NULL REFERENCES quality_bot.tests(id) ON DELETE CASCADE,
    test_type VARCHAR(50) NOT NULL DEFAULT 'control',
    current_question INTEGER NOT NULL DEFAULT 0,
    completed BOOLEAN NOT NULL DEFAULT FALSE,
    attempts INTEGER NOT NULL DEFAULT 1,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    completed_at TIMESTAMPTZ,
    CONSTRAINT uq_quality_bot_test_sessions_user_test_type UNIQUE (user_id, test_id, test_type)
);

ALTER TABLE quality_bot.test_sessions
    ADD COLUMN IF NOT EXISTS started_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

ALTER TABLE quality_bot.test_sessions
    ADD COLUMN IF NOT EXISTS completed_at TIMESTAMPTZ;

CREATE TABLE IF NOT EXISTS quality_bot.test_attempts (
    id BIGSERIAL PRIMARY KEY,
    session_id BIGINT NOT NULL REFERENCES quality_bot.test_sessions(id) ON DELETE CASCADE,
    attempt_number INTEGER NOT NULL DEFAULT 1,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    completed_at TIMESTAMPTZ,
    passed BOOLEAN NOT NULL DEFAULT FALSE,
    CONSTRAINT uq_quality_bot_test_attempts_session_number UNIQUE (session_id, attempt_number)
);

CREATE TABLE IF NOT EXISTS quality_bot.answers (
    id BIGSERIAL PRIMARY KEY,
    attempt_id BIGINT NOT NULL REFERENCES quality_bot.test_attempts(id) ON DELETE CASCADE,
    user_id BIGINT NOT NULL REFERENCES quality_bot.users(id) ON DELETE CASCADE,
    question_id BIGINT NOT NULL REFERENCES quality_bot.questions(id) ON DELETE CASCADE,
    answer INTEGER NOT NULL,
    is_correct BOOLEAN NOT NULL,
    correct_answer_at_time INTEGER NOT NULL,
    answered_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_quality_bot_answers_attempt_question UNIQUE (attempt_id, question_id)
);

ALTER TABLE quality_bot.answers
    DROP CONSTRAINT IF EXISTS ck_quality_bot_answers_answer;

ALTER TABLE quality_bot.answers
    DROP CONSTRAINT IF EXISTS ck_quality_bot_answers_correct_answer_at_time;

ALTER TABLE quality_bot.answers
    ADD COLUMN IF NOT EXISTS attempt_id BIGINT REFERENCES quality_bot.test_attempts(id) ON DELETE CASCADE;

ALTER TABLE quality_bot.answers
    ADD COLUMN IF NOT EXISTS is_correct BOOLEAN;

ALTER TABLE quality_bot.answers
    ADD COLUMN IF NOT EXISTS correct_answer_at_time INTEGER;

ALTER TABLE quality_bot.answers
    ADD COLUMN IF NOT EXISTS answered_at TIMESTAMPTZ NOT NULL DEFAULT NOW();

CREATE INDEX IF NOT EXISTS idx_quality_bot_users_telegram_id
    ON quality_bot.users(telegram_id);

CREATE INDEX IF NOT EXISTS idx_quality_bot_tests_test_date
    ON quality_bot.tests(test_date);

CREATE INDEX IF NOT EXISTS idx_quality_bot_questions_test_type
    ON quality_bot.questions(test_id, test_type);

CREATE INDEX IF NOT EXISTS idx_quality_bot_answers_user_question
    ON quality_bot.answers(user_id, question_id);

CREATE UNIQUE INDEX IF NOT EXISTS idx_quality_bot_answers_attempt_question
    ON quality_bot.answers(attempt_id, question_id)
    WHERE attempt_id IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_quality_bot_attempts_session
    ON quality_bot.test_attempts(session_id, attempt_number);

CREATE INDEX IF NOT EXISTS idx_quality_bot_sessions_user_test
    ON quality_bot.test_sessions(user_id, test_id, test_type);
