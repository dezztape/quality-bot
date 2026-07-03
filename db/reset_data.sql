TRUNCATE TABLE
    quality_bot.answers,
    quality_bot.test_attempts,
    quality_bot.test_sessions,
    quality_bot.questions,
    quality_bot.tests,
    quality_bot.users
RESTART IDENTITY CASCADE;
