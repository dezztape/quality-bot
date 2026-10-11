import asyncio
import os
import logging
import time
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from html import escape
from pathlib import Path
from typing import Any, Awaitable, Callable

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from dotenv import load_dotenv

from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramNetworkError, TelegramServerError
from aiogram.filters import Command
from aiogram.client.default import DefaultBotProperties

from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

from aiogram.types import (
    Message,
    CallbackQuery,
    BotCommand,
    FSInputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
)

from aiogram.utils.keyboard import InlineKeyboardBuilder

from sqlalchemy import (
    BigInteger,
    Boolean,
    CheckConstraint,
    Date,
    DateTime,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
    func,
    text,
    select,
)

from sqlalchemy.ext.asyncio import (
    AsyncSession,
    create_async_engine,
    async_sessionmaker,
)

from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


# =========================
# ENV
# =========================

load_dotenv()

# Логирование
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)
logger.info("🤖 Quality Bot starting...")

BOT_TOKEN = os.getenv("BOT_TOKEN")
DATABASE_URL = os.getenv("DATABASE_URL")
REDIS_URL = os.getenv("REDIS_URL")
SLOW_UPDATE_LOG_SECONDS = float(os.getenv("SLOW_UPDATE_LOG_SECONDS", "2"))
ADMIN_IDS = [
    int(x)
    for x in os.getenv("ADMIN_IDS", "").split(",")
    if x
]
MODERATOR_IDS = [
    int(x)
    for x in os.getenv("MODERATOR_IDS", "").split(",")
    if x
]
STATISTICIAN_IDS = [
    int(x)
    for x in os.getenv("STATISTICIAN_IDS", "").split(",")
    if x
]

logger.info(
    f"✅ Loaded {len(ADMIN_IDS)} admin(s), "
    f"{len(MODERATOR_IDS)} moderator(s), "
    f"{len(STATISTICIAN_IDS)} statistician(s)"
)


# =========================
# BOT
# =========================

bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(
        parse_mode=ParseMode.HTML
    )
)

dp = Dispatcher(storage=MemoryStorage())


class SlowUpdateLoggerMiddleware(BaseMiddleware):
    def __init__(self, threshold_seconds: float):
        self.threshold_seconds = threshold_seconds

    async def __call__(
        self,
        handler: Callable[[Any, dict[str, Any]], Awaitable[Any]],
        event: Any,
        data: dict[str, Any],
    ) -> Any:
        started_at = time.perf_counter()
        try:
            return await handler(event, data)
        finally:
            duration = time.perf_counter() - started_at
            if duration >= self.threshold_seconds:
                event_type = type(event).__name__
                user_id = getattr(getattr(event, "from_user", None), "id", None)
                chat_id = getattr(getattr(event, "chat", None), "id", None)
                details = ""

                if isinstance(event, Message):
                    text = event.text or event.caption or ""
                    details = text[:120]
                elif isinstance(event, CallbackQuery):
                    chat_id = getattr(getattr(event.message, "chat", None), "id", None)
                    details = event.data or ""

                logger.warning(
                    "🐢 Slow update handled: "
                    f"type={event_type}, user_id={user_id}, chat_id={chat_id}, "
                    f"duration_ms={duration * 1000:.0f}, details={details!r}"
                )


slow_update_logger = SlowUpdateLoggerMiddleware(SLOW_UPDATE_LOG_SECONDS)
dp.message.middleware(slow_update_logger)
dp.callback_query.middleware(slow_update_logger)

if REDIS_URL:
    logger.info("📝 Redis URL configured (will use for production)")
else:
    logger.warning("⚠️ REDIS_URL not set, using MemoryStorage (development mode)")


# =========================
# DATABASE
# =========================

DB_SCHEMA = "quality_bot"
INIT_SCHEMA_SQL = Path(__file__).resolve().parent / "db" / "init_schema.sql"

# Преобразуем DATABASE_URL для asyncpg если нужно
# Railway отправляет postgresql://, нам нужен postgresql+asyncpg://
if DATABASE_URL and "postgresql+asyncpg" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)
    logger.info("🔄 Converted DATABASE_URL to use asyncpg")

engine = create_async_engine(
    DATABASE_URL,
    echo=False,
)

async_session = async_sessionmaker(
    engine,
    expire_on_commit=False,
)


class Base(DeclarativeBase):
    pass


# =========================
# MODELS
# =========================

class User(Base):
    __tablename__ = "users"
    __table_args__ = {"schema": DB_SCHEMA}

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    telegram_id: Mapped[int] = mapped_column(
        BigInteger,
        unique=True,
    )

    full_name: Mapped[str] = mapped_column(
        String(255)
    )

    city: Mapped[str] = mapped_column(
        String(255),
        default="",
    )

    is_whitelisted: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
        onupdate=func.now(),
    )

    sessions: Mapped[list["TestSession"]] = relationship(
        back_populates="user"
    )

    answers: Mapped[list["Answer"]] = relationship(
        back_populates="user"
    )


class Test(Base):
    __tablename__ = "tests"
    __table_args__ = (
        UniqueConstraint("test_date", name="uq_quality_bot_tests_test_date"),
        {"schema": DB_SCHEMA},
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    title: Mapped[str] = mapped_column(
        String(255)
    )

    test_date: Mapped[date] = mapped_column(
        Date
    )

    active: Mapped[bool] = mapped_column(
        Boolean,
        default=True,
    )

    approved: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    approved_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
        onupdate=func.now(),
    )

    questions: Mapped[list["Question"]] = relationship(
        back_populates="test"
    )

    sessions: Mapped[list["TestSession"]] = relationship(
        back_populates="test"
    )


class Question(Base):
    __tablename__ = "questions"
    __table_args__ = (
        UniqueConstraint(
            "test_id",
            "test_type",
            "question_order",
            name="uq_quality_bot_questions_order",
        ),
        {"schema": DB_SCHEMA},
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    test_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.tests.id")
    )

    test_type: Mapped[str] = mapped_column(
        String(50),
        default="control",
    )

    question_order: Mapped[int] = mapped_column(
        Integer
    )

    image_file_id: Mapped[str] = mapped_column(
        String(500)
    )

    correct_answer: Mapped[int] = mapped_column(
        Integer
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
        onupdate=func.now(),
    )

    test: Mapped["Test"] = relationship(
        back_populates="questions"
    )

    answers: Mapped[list["Answer"]] = relationship(
        back_populates="question"
    )


class Answer(Base):
    __tablename__ = "answers"
    __table_args__ = (
        UniqueConstraint(
            "attempt_id",
            "question_id",
            name="uq_quality_bot_answers_attempt_question",
        ),
        {"schema": DB_SCHEMA},
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    attempt_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.test_attempts.id")
    )

    user_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.users.id")
    )

    question_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.questions.id")
    )

    answer: Mapped[int] = mapped_column(
        Integer
    )

    is_correct: Mapped[bool] = mapped_column(
        Boolean
    )

    correct_answer_at_time: Mapped[int] = mapped_column(
        Integer
    )

    answered_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    attempt: Mapped["TestAttempt"] = relationship(
        back_populates="answers"
    )

    user: Mapped["User"] = relationship(
        back_populates="answers"
    )

    question: Mapped["Question"] = relationship(
        back_populates="answers"
    )


class TestSession(Base):
    __tablename__ = "test_sessions"
    __table_args__ = (
        UniqueConstraint(
            "user_id",
            "test_id",
            "test_type",
            name="uq_quality_bot_test_sessions_user_test_type",
        ),
        {"schema": DB_SCHEMA},
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    user_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.users.id")
    )

    test_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.tests.id")
    )

    test_type: Mapped[str] = mapped_column(
        String(50),
        default="control",
    )

    current_question: Mapped[int] = mapped_column(
        Integer,
        default=0,
    )

    completed: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    started_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    completed_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    attempts: Mapped[int] = mapped_column(
        Integer,
        default=1,
    )

    user: Mapped["User"] = relationship(
        back_populates="sessions"
    )

    test: Mapped["Test"] = relationship(
        back_populates="sessions"
    )

    attempts_history: Mapped[list["TestAttempt"]] = relationship(
        back_populates="session"
    )


class TestAttempt(Base):
    __tablename__ = "test_attempts"
    __table_args__ = (
        UniqueConstraint(
            "session_id",
            "attempt_number",
            name="uq_quality_bot_test_attempts_session_number",
        ),
        {"schema": DB_SCHEMA},
    )

    id: Mapped[int] = mapped_column(BigInteger, primary_key=True)

    session_id: Mapped[int] = mapped_column(
        BigInteger,
        ForeignKey(f"{DB_SCHEMA}.test_sessions.id")
    )

    attempt_number: Mapped[int] = mapped_column(
        Integer,
        default=1,
    )

    started_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
    )

    completed_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    passed: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    session: Mapped["TestSession"] = relationship(
        back_populates="attempts_history"
    )

    answers: Mapped[list["Answer"]] = relationship(
        back_populates="attempt"
    )


# =========================
# STATES
# =========================

class QuestionAddStates(StatesGroup):
    choosing_test_type = State()
    waiting_photo = State()
    waiting_answer = State()
    confirming_question = State()


class UserRegistrationStates(StatesGroup):
    waiting_name = State()
    waiting_city = State()


class TestSelectionStates(StatesGroup):
    choosing_test_type = State()


class StatisticsStates(StatesGroup):
    waiting_date = State()
    waiting_period_start = State()
    waiting_period_end = State()


class AdminEditStates(StatesGroup):
    editing_image = State()
    editing_answer = State()


class TestStates(StatesGroup):
    passing_test = State()


# =========================
# KEYBOARDS
# =========================

user_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📝 Пройти тест")],
    ],
    resize_keyboard=True,
)

moderator_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📝 Пройти тест")],
        [KeyboardButton(text="➕ Добавить вопрос")],
        [KeyboardButton(text="📢 Уведомление")],
    ],
    resize_keyboard=True,
)

statistician_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📊 Статистика")],
        [KeyboardButton(text="📁 Экспорт Excel")],
    ],
    resize_keyboard=True,
)

admin_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📝 Пройти тест")],
        [KeyboardButton(text="➕ Добавить вопрос")],
        [KeyboardButton(text="👀 Просмотреть тест")],
        [KeyboardButton(text="📊 Статистика")],
        [KeyboardButton(text="📁 Экспорт Excel")],
        [KeyboardButton(text="📢 Уведомление")],
    ],
    resize_keyboard=True,
)


# =========================
# HELPERS
# =========================

async def create_db():
    logger.info("📦 Creating database tables...")
    try:
        sql_script = INIT_SCHEMA_SQL.read_text(encoding="utf-8")

        async with engine.begin() as conn:
            for statement in sql_script.split(";"):
                statement = statement.strip()
                if statement:
                    await conn.execute(text(statement))
        logger.info("✅ Database tables ready")
    except Exception as e:
        logger.error(f"❌ Database creation error: {e}")
        raise


async def is_admin(user_id: int):
    return user_id in ADMIN_IDS


async def is_moderator(user_id: int):
    return user_id in MODERATOR_IDS


async def is_statistician(user_id: int):
    return user_id in STATISTICIAN_IDS


async def is_admin_or_moderator(user_id: int):
    return user_id in ADMIN_IDS or user_id in MODERATOR_IDS


async def can_view_reports(user_id: int):
    return user_id in ADMIN_IDS or user_id in STATISTICIAN_IDS


async def get_user_by_tg_id(
    session: AsyncSession,
    telegram_id: int,
):
    result = await session.execute(
        select(User).where(
            User.telegram_id == telegram_id
        )
    )

    return result.scalar_one_or_none()


async def get_active_test(session: AsyncSession):

    result = await session.execute(
        select(Test).where(
            Test.test_date == date.today(),
            Test.active == True,
            Test.approved == True,
        )
    )

    return result.scalar_one_or_none()


async def get_test_for_passing(session: AsyncSession):
    """Получить тест для прохождения (без проверки одобрения)"""
    
    result = await session.execute(
        select(Test).where(
            Test.test_date == date.today(),
            Test.active == True,
        )
    )

    return result.scalars().first()


async def get_today_questions(
    session: AsyncSession,
    test_id: int,
    test_type: str = None,
):
    query = select(Question).where(Question.test_id == test_id)
    
    if test_type:
        query = query.where(Question.test_type == test_type)
    
    query = query.order_by(Question.question_order)
    
    result = await session.execute(query)

    return result.scalars().all()


async def already_completed_today(
    session: AsyncSession,
    user_id: int,
    test_id: int,
    test_type: str = "control",
):
    result = await session.execute(
        select(TestSession).where(
            TestSession.user_id == user_id,
            TestSession.test_id == test_id,
            TestSession.test_type == test_type,
            TestSession.completed == True,
        )
    )

    return result.scalar_one_or_none() is not None


async def get_test_session(
    session: AsyncSession,
    user_id: int,
    test_id: int,
    test_type: str,
):
    result = await session.execute(
        select(TestSession).where(
            TestSession.user_id == user_id,
            TestSession.test_id == test_id,
            TestSession.test_type == test_type,
        )
    )

    return result.scalar_one_or_none()


async def get_current_attempt(
    session: AsyncSession,
    test_session: TestSession,
):
    result = await session.execute(
        select(TestAttempt)
        .where(
            TestAttempt.session_id == test_session.id,
            TestAttempt.completed_at.is_(None),
        )
        .order_by(TestAttempt.attempt_number.desc())
    )

    attempt = result.scalars().first()

    if attempt:
        return attempt

    attempt = TestAttempt(
        session_id=test_session.id,
        attempt_number=test_session.attempts or 1,
        passed=False,
    )
    session.add(attempt)
    await session.flush()

    return attempt


async def get_or_create_test_session(
    session: AsyncSession,
    user_id: int,
    test_id: int,
    test_type: str,
):
    test_session = await get_test_session(
        session,
        user_id,
        test_id,
        test_type,
    )

    if test_session:
        if not test_session.completed:
            await get_current_attempt(session, test_session)
        return test_session

    test_session = TestSession(
        user_id=user_id,
        test_id=test_id,
        test_type=test_type,
        current_question=0,
        completed=False,
        attempts=1,
    )
    session.add(test_session)
    await session.flush()

    session.add(
        TestAttempt(
            session_id=test_session.id,
            attempt_number=1,
            passed=False,
        )
    )

    return test_session


async def is_test_approved(
    session: AsyncSession,
    test_id: int,
) -> bool:
    result = await session.execute(
        select(Test.approved).where(
            Test.id == test_id
        )
    )

    return bool(result.scalar_one_or_none())


def format_russian_date(value: date) -> str:
    months = {
        1: "января",
        2: "февраля",
        3: "марта",
        4: "апреля",
        5: "мая",
        6: "июня",
        7: "июля",
        8: "августа",
        9: "сентября",
        10: "октября",
        11: "ноября",
        12: "декабря",
    }

    return f"{value.day} {months[value.month]} {value.year} года"


def get_test_type_name(test_type: str) -> str:
    return "Контроль" if test_type == "control" else "Опыт"


# =========================
# START
# =========================

@dp.message(Command("start"))
async def start_handler(message: Message, state: FSMContext):
    logger.info(f"👤 User {message.from_user.id} started bot")

    async with async_session() as session:

        user = await get_user_by_tg_id(
            session,
            message.from_user.id,
        )

        if not user:

            user = User(
                telegram_id=message.from_user.id,
                full_name="Unknown",
                is_whitelisted=False,
                city="",
            )

            session.add(user)
            await session.commit()

        if (
            not user.is_whitelisted
            and message.from_user.id not in ADMIN_IDS
            and message.from_user.id not in MODERATOR_IDS
            and message.from_user.id not in STATISTICIAN_IDS
        ):
            await state.clear()
            await message.answer(
                "⛔ У вас нет доступа"
            )
            return

        # Если имя неполное или пусто - запрашиваем
        if not user.full_name or user.full_name == "Unknown":
            await state.set_state(UserRegistrationStates.waiting_name)
            msg = await message.answer("👤 Как вас зовут?")
            await state.update_data(instruction_message_id=msg.message_id)
            return

        # Если город не заполнен - запрашиваем
        if not user.city:
            await state.set_state(UserRegistrationStates.waiting_city)
            msg = await message.answer("🏙️ Из какого вы города?")
            await state.update_data(instruction_message_id=msg.message_id)
            return

        if message.from_user.id in ADMIN_IDS:
            keyboard = admin_keyboard
        elif message.from_user.id in MODERATOR_IDS:
            keyboard = moderator_keyboard
        elif message.from_user.id in STATISTICIAN_IDS:
            keyboard = statistician_keyboard
        else:
            keyboard = user_keyboard

        await message.answer(
            "✅ Добро пожаловать",
            reply_markup=keyboard,
        )


# =========================
# USER REGISTRATION
# =========================

@dp.message(UserRegistrationStates.waiting_name)
async def handle_user_name(message: Message, state: FSMContext):
    if not message.text:
        return

    await state.update_data(user_name=message.text)
    await state.set_state(UserRegistrationStates.waiting_city)

    data = await state.get_data()
    if data.get("instruction_message_id"):
        try:
            await message.bot.delete_message(
                message.chat.id,
                data["instruction_message_id"]
            )
        except:
            pass

    await message.delete()

    msg = await message.answer("🏙️ Из какого вы города?")
    await state.update_data(instruction_message_id=msg.message_id)


@dp.message(UserRegistrationStates.waiting_city)
async def handle_user_city(message: Message, state: FSMContext):
    if not message.text:
        return

    data = await state.get_data()
    user_name = data.get("user_name", "")
    city = message.text

    async with async_session() as session:
        user = await get_user_by_tg_id(
            session,
            message.from_user.id,
        )

        if user_name:
            user.full_name = user_name
        user.city = city

        await session.commit()

    await state.clear()
    await message.delete()

    if data.get("instruction_message_id"):
        try:
            await message.bot.delete_message(
                message.chat.id,
                data["instruction_message_id"]
            )
        except:
            pass

    if message.from_user.id in ADMIN_IDS:
        keyboard = admin_keyboard
    elif message.from_user.id in MODERATOR_IDS:
        keyboard = moderator_keyboard
    elif message.from_user.id in STATISTICIAN_IDS:
        keyboard = statistician_keyboard
    else:
        keyboard = user_keyboard

    await message.answer(
        "✅ Добро пожаловать",
        reply_markup=keyboard,
    )


# =========================
# BUTTONS
# =========================

@dp.message(F.text == "📝 Пройти тест")
async def start_test_button(
    message: Message,
    state: FSMContext,
):
    await start_test(message, state)


@dp.message(F.text == "➕ Добавить вопрос")
async def add_question_button(
    message: Message,
    state: FSMContext,
):
    await add_question_start(message, state)


@dp.message(F.text == "📊 Статистика")
async def results_button(message: Message, state: FSMContext):
    if not await can_view_reports(message.from_user.id):
        await message.answer("⛔ У вас нет доступа")
        return

    await state.clear()
    await show_statistics_menu(message)


@dp.message(F.text == "📁 Экспорт Excel")
async def export_button(message: Message):
    await export_handler(message)


@dp.message(F.text == "📢 Уведомление")
async def notification_button(message: Message):
    await notification_handler(message)


# =========================
# TEST TYPE SELECTION
# =========================

async def choose_test_type(message: Message, state: FSMContext):
    await state.set_state(TestSelectionStates.choosing_test_type)

    kb = InlineKeyboardBuilder()
    kb.button(text="📋 Контроль", callback_data="test_type:control")
    kb.button(text="📚 Опыт", callback_data="test_type:experience")
    kb.adjust(1)

    msg = await message.answer(
        "Выберите тип теста:",
        reply_markup=kb.as_markup()
    )
    await state.update_data(test_selection_message_id=msg.message_id)


@dp.callback_query(F.data.startswith("test_type:"))
async def handle_test_type(callback: CallbackQuery, state: FSMContext):
    test_type = callback.data.split(":")[1]

    await callback.answer()
    await callback.message.delete()

    await state.update_data(test_type=test_type)
    await state.set_state(TestStates.passing_test)

    # Используем callback.from_user.id для получения пользователя
    # и отправляем новое сообщение
    await start_test_from_callback(callback, state, test_type)


async def start_test_from_callback(
    callback: CallbackQuery,
    state: FSMContext,
    test_type: str = "control",
):
    async with async_session() as session:

        user = await get_user_by_tg_id(
            session,
            callback.from_user.id,
        )

        if (
            not user
            or (
                not user.is_whitelisted
                and callback.from_user.id not in ADMIN_IDS
            )
        ):
            await callback.message.answer(
                "⛔ Нет доступа"
            )
            return

        test = await get_active_test(session)

        if not test:
            await callback.message.answer(
               "❌ На сегодня тест отсутствует или не одобрен"
            )
            return

        completed = await already_completed_today(
            session,
            user.id,
            test.id,
            test_type,
        )

        if completed:
            await callback.message.answer(
                f"❌ Вы уже проходили этот тест ({test_type}) сегодня"
            )
            return

        questions = await get_today_questions(
            session,
            test.id,
            test_type,
        )

        if not questions:
            await callback.message.answer(
                "❌ Вопросы отсутствуют"
            )
            return

        new_session = await get_or_create_test_session(
            session,
            user_id=user.id,
            test_id=test.id,
            test_type=test_type,
        )

        await session.commit()

        # Инициализируем прогресс для этого типа теста
        # Сохраняем test_type в состояние для использования в send_question
        await state.update_data(test_type=test_type)

        await send_question(
            callback.from_user.id,
            state,
        )


# =========================
# ALLOW
# =========================

@dp.message(Command("allow"))
async def allow_user(message: Message):

    if not await is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа")
        return

    args = message.text.split()

    if len(args) != 2:
        await message.answer(
            "Использование:\n/allow TELEGRAM_ID"
        )
        return

    telegram_id = int(args[1])

    async with async_session() as session:

        user = await get_user_by_tg_id(
            session,
            telegram_id,
        )

        if not user:
            await message.answer(
                "Пользователь ещё не запускал бота"
            )
            return

        user.is_whitelisted = True

        await session.commit()

        await message.answer(
            "✅ Пользователь добавлен"
        )


# =========================
# ADD QUESTION
# =========================

@dp.message(Command("add_question"))
@dp.message(F.text == "➕ Добавить вопрос")
async def add_question_start(
    message: Message,
    state: FSMContext,
):

    if not await is_admin_or_moderator(message.from_user.id):
        await message.answer("⛔ У вас нет доступа")
        return

    async with async_session() as session:

        test = await get_test_for_passing(session)

        if not test:

            test = Test(
                title=f"Тест {date.today()}",
                test_date=date.today(),
                active=True,
            )

            session.add(test)
            await session.commit()

        if test.approved:
            await state.clear()
            await message.answer(
                "🔐 Тест уже опубликован, добавление вопросов закрыто"
            )
            return

        await state.update_data(
            test_id=test.id
        )

    # Запрашиваем тип теста
    await state.set_state(QuestionAddStates.choosing_test_type)

    kb = InlineKeyboardBuilder()
    kb.button(text="📋 Контроль", callback_data="add_test_type:control")
    kb.button(text="📚 Опыт", callback_data="add_test_type:experience")
    kb.adjust(1)

    msg = await message.answer(
        "Выберите тип теста для добавления вопроса:",
        reply_markup=kb.as_markup()
    )

    await state.update_data(instruction_message_id=msg.message_id)


@dp.callback_query(F.data.startswith("add_test_type:"))
async def handle_add_test_type(callback: CallbackQuery, state: FSMContext):
    test_type = callback.data.split(":")[1]

    await callback.answer()
    await callback.message.delete()

    await state.update_data(question_test_type=test_type)
    await state.set_state(QuestionAddStates.waiting_photo)

    msg = await callback.message.answer(
        "📷 Отправьте изображение"
    )

    await state.update_data(
        instruction_message_id=msg.message_id
    )


@dp.callback_query(F.data.startswith("confirm_question:"))
async def handle_confirm_question(callback: CallbackQuery, state: FSMContext):
    action = callback.data.split(":")[1]
    
    await callback.answer()
    await callback.message.delete()
    
    data = await state.get_data()
    
    if action == "no":
        # Переснять - возвращаемся к выбору фото
        await state.set_state(QuestionAddStates.waiting_photo)
        msg = await callback.message.answer(
            "📷 Отправьте новое изображение"
        )
        await state.update_data(
            instruction_message_id=msg.message_id
        )
        return
    
    # action == "yes" - подтверждаем и сохраняем вопрос
    async with async_session() as session:
        if await is_test_approved(session, data["test_id"]):
            await state.clear()
            await callback.message.answer(
                "🔐 Тест уже опубликован, добавление вопросов закрыто"
            )
            return

        result = await session.execute(
            select(Question).where(
                Question.test_id == data["test_id"],
                Question.test_type == data.get("question_test_type", "control"),
            )
        )
        
        questions = result.scalars().all()
        next_order = len(questions) + 1
        
        question = Question(
            test_id=data["test_id"],
            test_type=data.get("question_test_type", "control"),
            question_order=next_order,
            image_file_id=data["photo_id"],
            correct_answer=data["correct_answer"],
        )
        
        session.add(question)
        await session.commit()
    
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Добавить еще вопрос", callback_data="add_question_more")
    kb.button(text="✅ Закрыть", callback_data="add_question_finish")
    kb.adjust(1)

    await state.update_data(
        test_id=data["test_id"],
        question_test_type=data.get("question_test_type", "control"),
    )

    await callback.message.answer(
        "✅ Вопрос сохранён",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data == "add_question_more")
async def add_question_more(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    await callback.message.delete()

    data = await state.get_data()

    async with async_session() as session:
        if await is_test_approved(session, data["test_id"]):
            await state.clear()
            await callback.message.answer(
                "🔐 Тест уже опубликован, добавление вопросов закрыто"
            )
            return

    await state.set_state(QuestionAddStates.waiting_photo)

    msg = await callback.message.answer(
        "📷 Отправьте изображение следующего вопроса"
    )

    await state.update_data(
        instruction_message_id=msg.message_id
    )


@dp.callback_query(F.data == "add_question_finish")
async def add_question_finish(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    await callback.message.delete()

    await state.clear()

    await callback.message.answer(
        "✅ Добавление вопросов завершено"
    )


# =========================
# PHOTO HANDLER
# =========================

@dp.message(F.photo)
async def photo_handler(
    message: Message,
    state: FSMContext,
):

    current_state = await state.get_state()

    # ДОБАВЛЕНИЕ ВОПРОСА
    if current_state == QuestionAddStates.waiting_photo.state:

        photo = message.photo[-1]
        data = await state.get_data()

        async with async_session() as session:
            if await is_test_approved(session, data["test_id"]):
                await state.clear()
                await message.answer(
                    "🔐 Тест уже опубликован, добавление вопросов закрыто"
                )
                return

        await state.update_data(
            photo_id=photo.file_id
        )

        await state.set_state(QuestionAddStates.waiting_answer)

        await message.delete()

        # Удаляем предыдущее инструкционное сообщение
        if data.get("instruction_message_id"):
            try:
                await message.bot.delete_message(
                    message.chat.id,
                    data["instruction_message_id"]
                )
            except:
                pass

        msg = await message.answer(
            "Введите правильный ответ числом"
        )

        await state.update_data(
            instruction_message_id=msg.message_id
        )

        return

    # РЕДАКТИРОВАНИЕ КАРТИНКИ
    if current_state == AdminEditStates.editing_image.state:

        photo = message.photo[-1]

        data = await state.get_data()

        async with async_session() as session:

            result = await session.execute(
                select(Question, Test)
                .join(Test, Test.id == Question.test_id)
                .where(Question.id == data["edit_question_id"])
            )

            row = result.first()

            if not row:
                await state.clear()
                await message.answer("❌ Вопрос не найден")
                return

            q, test = row

            if test.approved:
                await state.clear()
                await message.answer(
                    "🔐 Тест уже опубликован, редактирование закрыто"
                )
                return

            q.image_file_id = photo.file_id

            await session.commit()

        await state.clear()

        await message.delete()

        # Удаляем инструкционное сообщение
        if data.get("instruction_message_id"):
            try:
                await message.bot.delete_message(
                    message.chat.id,
                    data["instruction_message_id"]
                )
            except:
                pass

        await message.answer(
            "✅ Картинка обновлена"
        )

        return





# =========================
# TEST
# =========================

@dp.message(Command("test"))
async def start_test_command(
    message: Message,
    state: FSMContext,
):
    await start_test(message, state)


async def start_test(
    message: Message,
    state: FSMContext,
    test_type: str | None = None,
):
    requested_test_type = test_type
    logger.info(f"📝 User {message.from_user.id} starting test")

    async with async_session() as session:

        user = await get_user_by_tg_id(
            session,
            message.from_user.id,
        )

        if (
            not user
            or (
                not user.is_whitelisted
                and message.from_user.id not in ADMIN_IDS
            )
        ):
            await message.answer(
                "⛔ Нет доступа"
            )
            return

        test = await get_active_test(session)

        if not test:
            await message.answer(
                "❌ На сегодня тест отсутствует или не одобрен"
            )
            return

        if requested_test_type is None:
            for candidate_type in ("control", "experience"):
                completed = await already_completed_today(
                    session,
                    user.id,
                    test.id,
                    candidate_type,
                )
                questions = await get_today_questions(
                    session,
                    test.id,
                    candidate_type,
                )

                if questions and not completed:
                    test_type = candidate_type
                    break

            if test_type is None:
                await message.answer(
                    "✅ Вы уже прошли все доступные тесты сегодня"
                )
                return

        logger.info(f"📝 User {message.from_user.id} starting {test_type} test")

        completed = await already_completed_today(
            session,
            user.id,
            test.id,
            test_type,
        )

        if completed:
            await message.answer(
                f"❌ Вы уже проходили этот тест ({test_type}) сегодня"
            )
            return

        questions = await get_today_questions(
            session,
            test.id,
            test_type,
        )

        if not questions:
            await message.answer(
                "❌ Вопросы отсутствуют"
            )
            return

        new_session = await get_or_create_test_session(
            session,
            user_id=user.id,
            test_id=test.id,
            test_type=test_type,
        )

        await session.commit()

        # Инициализируем прогресс для этого типа теста
        # Сохраняем test_type в состояние для использования в send_question
        await state.update_data(test_type=test_type)

        await state.set_state(
            TestStates.passing_test
        )

        if test_type == "control":
            await message.answer("📋 Начинаем контроль")
        else:
            await message.answer("📚 Начинаем опыт")

        await send_question(
            message.from_user.id,
            state,
        )


# =========================
# SEND QUESTION
# =========================

async def send_question(
    user_tg_id: int,
    state: FSMContext,
):

    async with async_session() as session:

        test = await get_test_for_passing(session)

        if not test:
            return

        data = await state.get_data()
        test_type = data.get("test_type", "control")

        user = await get_user_by_tg_id(
            session,
            user_tg_id,
        )

        if not user:
            return

        test_session = await get_test_session(
            session,
            user.id,
            test.id,
            test_type,
        )

        if not test_session or test_session.completed:
            return

        attempt = await get_current_attempt(session, test_session)

        questions = await get_today_questions(
            session,
            test.id,
            test_type,
        )

        index = test_session.current_question or 0

        if index >= len(questions):
            answers_result = await session.execute(
                select(Answer, Question)
                .join(Question, Question.id == Answer.question_id)
                .where(Answer.attempt_id == attempt.id)
                .where(Question.test_id == test.id)
                .where(Question.test_type == test_type)
            )
            answers_data = answers_result.all()

            all_correct = (
                len(answers_data) == len(questions)
                and all(answer.is_correct for answer, _ in answers_data)
            )

            now = datetime.now(timezone.utc)

            if not all_correct:
                failed_attempt_number = attempt.attempt_number
                wrong_answers = [
                    (answer, question)
                    for answer, question in answers_data
                    if not answer.is_correct
                ]

                attempt.completed_at = now
                attempt.passed = False
                test_session.attempts = (test_session.attempts or 1) + 1
                test_session.current_question = 0

                session.add(
                    TestAttempt(
                        session_id=test_session.id,
                        attempt_number=test_session.attempts,
                        passed=False,
                    )
                )

                await session.commit()

                data = await state.get_data()
                if data.get("last_question_message_id"):
                    try:
                        await bot.delete_message(
                            user_tg_id,
                            data["last_question_message_id"]
                        )
                    except:
                        pass

                logger.warning(f"⚠️ User {user_tg_id} failed test {test.id} ({test_type}), attempt {test_session.attempts}")

                await bot.send_message(
                    user_tg_id,
                    f"❌ Тест не пройден\n\n"
                    f"{get_test_type_name(test_type)}\n"
                    f"Провальная попытка #{failed_attempt_number}\n\n"
                    f"Начнём заново! 🔄"
                )

                if failed_attempt_number == 5 and wrong_answers:
                    await bot.send_message(
                        user_tg_id,
                        "Ниже вопросы, где были ошибки. Правильные ответы не показываю."
                    )

                    for wrong_answer, wrong_question in wrong_answers:
                        await bot.send_photo(
                            user_tg_id,
                            wrong_question.image_file_id,
                            caption=(
                                f"❌ {get_test_type_name(test_type)}\n"
                                f"Вопрос {wrong_question.question_order}\n"
                                f"Ваш ответ: {wrong_answer.answer}"
                            ),
                        )

                await state.update_data(test_type=test_type)

                await send_question(
                    user_tg_id,
                    state,
                )
                return

            attempt.completed_at = now
            attempt.passed = True
            test_session.completed = True
            test_session.completed_at = now
            test_session.current_question = len(questions)

            await session.commit()

            data = await state.get_data()
            if data.get("last_question_message_id"):
                try:
                    await bot.delete_message(
                        user_tg_id,
                        data["last_question_message_id"]
                    )
                except:
                    pass

            logger.info(f"✅ User {user_tg_id} completed test {test.id} ({test_type}) in {test_session.attempts} attempt(s)")

            if test_type == "control":
                experience_questions = await get_today_questions(
                    session,
                    test.id,
                    "experience",
                )
                experience_completed = await already_completed_today(
                    session,
                    user.id,
                    test.id,
                    "experience",
                )

                if experience_questions and not experience_completed:
                    await get_or_create_test_session(
                        session,
                        user.id,
                        test.id,
                        "experience",
                    )
                    await session.commit()

                    await bot.send_message(
                        user_tg_id,
                        f"✅ Контроль завершён успешно\n\n"
                        f"Количество попыток: {test_session.attempts}\n\n"
                        f"📚 Переходим к опыту"
                    )

                    await state.update_data(test_type="experience")
                    await send_question(
                        user_tg_id,
                        state,
                    )
                    return

            await bot.send_message(
                user_tg_id,
                f"✅ {get_test_type_name(test_type)} завершён успешно!\n\n"
                f"Количество попыток: {test_session.attempts}"
            )

            await state.clear()
            return

        question = questions[index]

        data = await state.get_data()
        if data.get("last_question_message_id"):
            try:
                await bot.delete_message(
                    user_tg_id,
                    data["last_question_message_id"]
                )
            except:
                pass

        sent_message = await bot.send_photo(
            user_tg_id,
            question.image_file_id,
            caption=(
                f"{get_test_type_name(test_type)}\n"
                f"Вопрос {index + 1} "
                f"из {len(questions)}\n\n"
                f"Введите ответ сообщением"
            ),
        )

        await state.update_data(
            last_question_message_id=sent_message.message_id,
            current_question_id=question.id,
            test_type=test_type,
        )


# =========================
# TEST VIEW
# =========================

@dp.message(F.text == "👀 Просмотреть тест")
async def view_test_button(message: Message, state: FSMContext):
    
    if not await is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа")
        return

    # Запрашиваем тип теста
    await state.set_state(TestSelectionStates.choosing_test_type)

    kb = InlineKeyboardBuilder()
    kb.button(text="📋 Контроль", callback_data="view_test_type:control")
    kb.button(text="📚 Опыт", callback_data="view_test_type:experience")
    kb.adjust(1)

    msg = await message.answer(
        "Выберите тип теста для просмотра:",
        reply_markup=kb.as_markup()
    )


@dp.callback_query(F.data.startswith("view_test_type:"))
async def handle_view_test_type(callback: CallbackQuery, state: FSMContext):
    test_type = callback.data.split(":")[1]

    await callback.answer()
    await callback.message.delete()

    async with async_session() as session:

        test = await get_test_for_passing(session)

        if not test:
            await callback.message.answer(
                "❌ Тест не найден"
            )
            return

        questions = await get_today_questions(
            session,
            test.id,
            test_type,
        )

        if not questions:
            await callback.message.answer(
                f"❌ Вопросов нет в типе '{test_type}'"
            )
            return

        kb = InlineKeyboardBuilder()

        for q in questions:
            kb.button(
                text=f"Вопрос {q.question_order}",
                callback_data=f"admin_q:{q.id}"
            )

        kb.adjust(3)

        await callback.message.answer(
            f"📋 Вопросы теста ({test_type}):",
            reply_markup=kb.as_markup()
        )


@dp.callback_query(F.data.startswith("preview_test_type:"))
async def handle_preview_test_type(callback: CallbackQuery):
    test_type = callback.data.split(":")[1]

    await callback.answer()

    async with async_session() as session:

        test = await get_test_for_passing(session)

        if not test:
            await callback.message.answer(
                "❌ Тест не найден"
            )
            return

        questions = await get_today_questions(
            session,
            test.id,
            test_type,
        )

        if not questions:
            await callback.message.answer(
                f"❌ Вопросов нет в типе '{test_type}'"
            )
            return

        kb = InlineKeyboardBuilder()

        for q in questions:
            kb.button(
                text=f"Вопрос {q.question_order}",
                callback_data=f"admin_q:{q.id}"
            )

        kb.adjust(3)

        await callback.message.answer(
            f"📋 Проверка теста ({test_type}):",
            reply_markup=kb.as_markup()
        )


# =========================
# ADMIN QUESTION VIEW
# =========================

@dp.callback_query(F.data.startswith("admin_q:"))
async def admin_view_question(
    callback: CallbackQuery
):

    if not await is_admin_or_moderator(callback.from_user.id):
        return

    question_id = int(
        callback.data.split(":")[1]
    )

    async with async_session() as session:

        result = await session.execute(
            select(Question).where(
                Question.id == question_id
            )
        )

        q = result.scalar_one()

        test_result = await session.execute(
            select(Test).where(
                Test.id == q.test_id
            )
        )
        test = test_result.scalar_one()

    kb = InlineKeyboardBuilder()

    if not test.approved:
        kb.button(
            text="✏️ Картинка",
            callback_data=f"edit_img:{q.id}"
        )

        kb.button(
            text="✏️ Ответ",
            callback_data=f"edit_ans:{q.id}"
        )

        kb.adjust(1)

    await callback.answer()

    await callback.message.answer_photo(
        q.image_file_id,
        caption=(
            f"📌 Вопрос {q.question_order}\n"
            f"🔒 Ответ: {q.correct_answer}"
            + ("\n\n🔐 Тест уже опубликован, редактирование закрыто" if test.approved else "")
        ),
        reply_markup=kb.as_markup() if not test.approved else None
    )


# =========================
# PROCESS ANSWER
# =========================

async def save_test_answer(
    user_tg_id: int,
    state: FSMContext,
    question_id: int,
    answer: int,
) -> bool:
    async with async_session() as session:

        user = await get_user_by_tg_id(
            session,
            user_tg_id,
        )

        question_result = await session.execute(
            select(Question).where(
                Question.id == question_id
            )
        )
        question = question_result.scalar_one_or_none()

        if not user or not question:
            return False

        test_session = await get_test_session(
            session,
            user.id,
            question.test_id,
            question.test_type,
        )

        if not test_session or test_session.completed:
            return False

        attempt = await get_current_attempt(session, test_session)

        existing_result = await session.execute(
            select(Answer).where(
                Answer.attempt_id == attempt.id,
                Answer.question_id == question.id,
            )
        )
        existing_answer = existing_result.scalar_one_or_none()

        is_correct = answer == question.correct_answer

        if existing_answer:
            existing_answer.answer = answer
            existing_answer.is_correct = is_correct
            existing_answer.correct_answer_at_time = question.correct_answer
        else:
            session.add(
                Answer(
                    attempt_id=attempt.id,
                    user_id=user.id,
                    question_id=question.id,
                    answer=answer,
                    is_correct=is_correct,
                    correct_answer_at_time=question.correct_answer,
                )
            )
            test_session.current_question = (test_session.current_question or 0) + 1

        await session.commit()
        await state.update_data(test_type=question.test_type)

    return True


@dp.callback_query(F.data.startswith("answer:"))
async def process_answer(
    callback: CallbackQuery,
    state: FSMContext,
):

    parts = callback.data.split(":")

    question_id = int(parts[1])

    answer = int(parts[2])

    saved = await save_test_answer(
        callback.from_user.id,
        state,
        question_id,
        answer,
    )

    if not saved:
        await callback.answer()
        return

    await callback.message.delete()

    await callback.answer()

    await send_question(
        callback.from_user.id,
        state,
    )


@dp.callback_query(F.data.startswith("confirm_test_answer:"))
async def confirm_test_answer(callback: CallbackQuery, state: FSMContext):
    action = callback.data.split(":")[1]

    await callback.answer()
    await callback.message.delete()

    data = await state.get_data()

    if action == "no":
        await state.update_data(
            pending_question_id=None,
            pending_answer=None,
        )
        await callback.message.answer(
            "Введите ответ еще раз"
        )
        return

    question_id = data.get("pending_question_id")
    answer = data.get("pending_answer")

    if question_id is None or answer is None:
        await callback.message.answer(
            "❌ Ответ не найден. Введите ответ еще раз"
        )
        return

    saved = await save_test_answer(
        callback.from_user.id,
        state,
        int(question_id),
        int(answer),
    )

    await state.update_data(
        pending_question_id=None,
        pending_answer=None,
    )

    if not saved:
        await callback.message.answer(
            "❌ Не удалось сохранить ответ. Попробуйте еще раз"
        )
        return

    await send_question(
        callback.from_user.id,
        state,
    )


# =========================
# IMAGE EDIT
# =========================

@dp.callback_query(F.data.startswith("edit_img:"))
async def edit_image_start(
    callback: CallbackQuery,
    state: FSMContext,
):

    if not await is_admin(callback.from_user.id):
        await callback.answer(
            "⛔ У вас нет доступа к редактированию. Напишите администратору @dezztape",
            show_alert=True,
        )
        return

    question_id = int(
        callback.data.split(":")[1]
    )

    async with async_session() as session:
        result = await session.execute(
            select(Question, Test)
            .join(Test, Test.id == Question.test_id)
            .where(Question.id == question_id)
        )
        row = result.first()

        if not row:
            await callback.answer(
                "Вопрос не найден",
                show_alert=True,
            )
            return

        _, test = row

        if test.approved:
            await callback.answer(
                "Тест уже опубликован, редактирование закрыто",
                show_alert=True,
            )
            return

    await state.update_data(
        edit_question_id=question_id
    )

    await state.set_state(
        AdminEditStates.editing_image
    )

    await callback.answer()

    msg = await callback.message.answer(
        "📸 Отправьте новую картинку"
    )

    await state.update_data(
        instruction_message_id=msg.message_id
    )


# =========================
# ANSWER EDIT
# =========================

@dp.callback_query(F.data.startswith("edit_ans:"))
async def edit_answer_start(
    callback: CallbackQuery,
    state: FSMContext,
):

    if not await is_admin(callback.from_user.id):
        await callback.answer(
            "⛔ У вас нет доступа к редактированию. Напишите администратору @dezztape",
            show_alert=True,
        )
        return

    question_id = int(
        callback.data.split(":")[1]
    )

    async with async_session() as session:
        result = await session.execute(
            select(Question, Test)
            .join(Test, Test.id == Question.test_id)
            .where(Question.id == question_id)
        )
        row = result.first()

        if not row:
            await callback.answer(
                "Вопрос не найден",
                show_alert=True,
            )
            return

        _, test = row

        if test.approved:
            await callback.answer(
                "Тест уже опубликован, редактирование закрыто",
                show_alert=True,
            )
            return

    await state.update_data(
        edit_question_id=question_id
    )

    await state.set_state(
        AdminEditStates.editing_answer
    )

    await callback.answer()

    msg = await callback.message.answer(
        "✏️ Введите новый правильный ответ числом"
    )

    await state.update_data(
        instruction_message_id=msg.message_id
    )


# =========================
# RESULTS
# =========================

@dp.callback_query(F.data.startswith("stats_type:"))
async def handle_stats_type(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    test_type = callback.data.split(":")[1]
    
    await callback.answer()
    await callback.message.delete()
    
    await show_statistics(callback.message, test_type)


@dp.message(Command("results"))
async def results_command(message: Message, state: FSMContext):
    if not await can_view_reports(message.from_user.id):
        await message.answer("⛔ У вас нет доступа")
        return

    await state.clear()
    await show_statistics_menu(message)


async def show_statistics_menu(message: Message):
    kb = InlineKeyboardBuilder()
    kb.button(text="👤 По сотруднику", callback_data="stats_by_employee")
    kb.button(text="📅 Сегодня", callback_data="stats_range:today")
    kb.button(text="↩️ Вчера", callback_data="stats_range:yesterday")
    kb.button(text="🗓 Выбрать день", callback_data="stats_custom_date")
    kb.button(text="📆 Выбрать период", callback_data="stats_custom_period")
    kb.button(text="📌 Текущий месяц", callback_data="stats_range:month")
    kb.button(text="🌐 Все дни", callback_data="stats_range:all")
    kb.adjust(1)

    await message.answer(
        "Выберите период для просмотра статистики:",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data == "stats_by_employee")
async def handle_stats_by_employee(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer("У вас нет доступа", show_alert=True)
        return

    await callback.answer()
    await callback.message.delete()

    excluded_stats_ids = set(ADMIN_IDS + STATISTICIAN_IDS)
    async with async_session() as session:
        users_result = await session.execute(
            select(User)
            .where(User.is_whitelisted == True)
            .order_by(User.city, User.full_name)
        )
        employees = [
            user for user in users_result.scalars().all()
            if user.telegram_id not in excluded_stats_ids
        ]

    if not employees:
        await callback.message.answer("❌ Сотрудники не найдены")
        return

    kb = InlineKeyboardBuilder()
    for employee in employees:
        employee_name = escape(employee.full_name or "Без имени")
        employee_city = escape(employee.city or "Без города")
        kb.button(
            text=f"{employee_name} — {employee_city}",
            callback_data=f"stats_employee:{employee.id}",
        )
    kb.adjust(1)

    await callback.message.answer(
        "Выберите сотрудника для просмотра статистики за текущий месяц:",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data.startswith("stats_employee:"))
async def handle_stats_employee(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer("У вас нет доступа", show_alert=True)
        return

    employee_id = int(callback.data.split(":", 1)[1])
    await callback.answer()
    await callback.message.delete()

    today = date.today()
    month_start = today.replace(day=1)
    daily_stats = defaultdict(lambda: {"total": 0, "correct": 0, "attempts": 0})

    async with async_session() as session:
        employee_result = await session.execute(
            select(User).where(User.id == employee_id)
        )
        employee = employee_result.scalar_one_or_none()
        if not employee:
            await callback.message.answer("❌ Сотрудник не найден")
            return

        sessions_result = await session.execute(
            select(TestSession, Test)
            .join(Test, Test.id == TestSession.test_id)
            .where(TestSession.user_id == employee_id)
            .where(Test.test_date >= month_start)
            .where(Test.test_date <= today)
        )

        for test_session, test in sessions_result.all():
            day_stats = daily_stats[test.test_date]
            day_stats["attempts"] += test_session.attempts or 1

            attempts_result = await session.execute(
                select(TestAttempt)
                .where(TestAttempt.session_id == test_session.id)
            )
            attempts = attempts_result.scalars().all()

            for attempt in attempts:
                answers_result = await session.execute(
                    select(Answer).where(Answer.attempt_id == attempt.id)
                )
                answers = answers_result.scalars().all()
                day_stats["total"] += len(answers)
                day_stats["correct"] += sum(
                    int(answer.is_correct) for answer in answers
                )

    employee_name = escape(employee.full_name or "Без имени")
    employee_city = escape(employee.city or "Без города")
    text = (
        f"📊 <b>{employee_name}</b>\n"
        f"🏙 {employee_city}\n"
        f"📅 Статистика за {today.strftime('%m.%Y')}\n"
        f"{'=' * 38}\n\n"
    )

    active_days = []
    current_day = month_start
    while current_day <= today:
        stats = daily_stats.get(current_day)
        if stats and stats["total"]:
            accuracy = stats["correct"] / stats["total"] * 100
            errors = stats["total"] - stats["correct"]
            text += (
                f"📌 {current_day.strftime('%d.%m.%Y')}: "
                f"точность {accuracy:.0f}%, "
                f"ошибок {errors}, "
                f"попыток {stats['attempts']}\n"
            )
            active_days.append((current_day, accuracy, errors, stats["attempts"]))
        else:
            text += f"📌 {current_day.strftime('%d.%m.%Y')}: нет прохождения\n"
        current_day += timedelta(days=1)

    if active_days:
        first_day = active_days[0]
        last_day = active_days[-1]
        accuracy_delta = last_day[1] - first_day[1]
        errors_delta = last_day[2] - first_day[2]
        attempts_delta = last_day[3] - first_day[3]
        trend = "📈 Есть прогресс" if accuracy_delta > 0 else (
            "📉 Точность снизилась" if accuracy_delta < 0 else "➡️ Точность без изменений"
        )
        text += (
            f"\n<b>Динамика обучения</b>\n"
            f"{trend}\n"
            f"Точность: {accuracy_delta:+.0f} п.п.\n"
            f"Ошибки: {errors_delta:+d}\n"
            f"Попытки: {attempts_delta:+d}\n"
        )
    else:
        text += "\nЗа текущий месяц сотрудник тесты не проходил."

    chunks = []
    current_chunk = ""
    for line in text.splitlines(keepends=True):
        if current_chunk and len(current_chunk) + len(line) > 3900:
            chunks.append(current_chunk)
            current_chunk = ""
        current_chunk += line
    if current_chunk:
        chunks.append(current_chunk)

    for chunk in chunks:
        await callback.message.answer(chunk)


def parse_stats_date(value: str) -> date | None:
    for date_format in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), date_format).date()
        except ValueError:
            continue
    return None


async def show_statistics_for_range(
    message: Message,
    start_date: date,
    end_date: date,
):
    await show_statistics(
        message,
        start_date=start_date,
        end_date=end_date,
    )


@dp.callback_query(F.data.startswith("stats_range:"))
async def handle_stats_range(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    await callback.answer()
    await callback.message.delete()

    period = callback.data.split(":", 1)[1]
    today = date.today()
    if period == "today":
        start_date = end_date = today
    elif period == "yesterday":
        start_date = end_date = today - timedelta(days=1)
    elif period == "month":
        start_date = today.replace(day=1)
        end_date = today
    else:
        start_date = date.min
        end_date = date.max

    await show_statistics_for_range(callback.message, start_date, end_date)


@dp.callback_query(F.data == "stats_custom_date")
async def handle_stats_custom_date(callback: CallbackQuery, state: FSMContext):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer("У вас нет доступа", show_alert=True)
        return

    await callback.answer()
    await callback.message.delete()
    await state.set_state(StatisticsStates.waiting_date)
    await callback.message.answer(
        "Введите дату в формате ДД.ММ.ГГГГ, например 17.09.2026:"
    )


@dp.message(StatisticsStates.waiting_date)
async def handle_stats_date_input(message: Message, state: FSMContext):
    selected_date = parse_stats_date(message.text or "")
    if not selected_date:
        await message.answer("Не удалось распознать дату. Используйте формат ДД.ММ.ГГГГ.")
        return

    await state.clear()
    await show_statistics_for_range(message, selected_date, selected_date)


@dp.callback_query(F.data == "stats_custom_period")
async def handle_stats_custom_period(callback: CallbackQuery, state: FSMContext):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer("У вас нет доступа", show_alert=True)
        return

    await callback.answer()
    await callback.message.delete()
    await state.set_state(StatisticsStates.waiting_period_start)
    await callback.message.answer(
        "Введите начальную дату периода в формате ДД.ММ.ГГГГ:"
    )


@dp.message(StatisticsStates.waiting_period_start)
async def handle_stats_period_start(message: Message, state: FSMContext):
    start_date = parse_stats_date(message.text or "")
    if not start_date:
        await message.answer("Не удалось распознать дату. Используйте формат ДД.ММ.ГГГГ.")
        return

    await state.update_data(stats_start_date=start_date.isoformat())
    await state.set_state(StatisticsStates.waiting_period_end)
    await message.answer("Введите конечную дату периода в формате ДД.ММ.ГГГГ:")


@dp.message(StatisticsStates.waiting_period_end)
async def handle_stats_period_end(message: Message, state: FSMContext):
    end_date = parse_stats_date(message.text or "")
    data = await state.get_data()
    start_value = data.get("stats_start_date")
    start_date = date.fromisoformat(start_value) if start_value else None

    if not end_date or not start_date:
        await message.answer("Не удалось распознать дату. Используйте формат ДД.ММ.ГГГГ.")
        return
    if end_date < start_date:
        await message.answer("Конечная дата не может быть раньше начальной.")
        return

    await state.clear()
    await show_statistics_for_range(message, start_date, end_date)


@dp.callback_query(F.data == "stats_available_days")
async def handle_stats_available_days(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    await callback.answer()
    await callback.message.delete()

    async with async_session() as session:
        result = await session.execute(
            select(Test)
            .order_by(Test.test_date.desc())
            .limit(30)
        )
        tests = result.scalars().all()

    if not tests:
        await callback.message.answer(
            "❌ Доступных дней для статистики нет"
        )
        return

    kb = InlineKeyboardBuilder()
    for test in tests:
        kb.button(
            text=format_russian_date(test.test_date),
            callback_data=f"stats_test:{test.id}",
        )
    kb.adjust(1)

    await callback.message.answer(
        "Выберите день для просмотра статистики:",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data.startswith("stats_test:"))
async def handle_stats_test(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    test_id = int(callback.data.split(":")[1])

    await callback.answer()
    await callback.message.delete()

    await show_statistics(callback.message, test_id=test_id)


async def show_statistics(
    message: Message,
    test_type: str | None = None,
    test_id: int | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
):

    async with async_session() as session:
        if test_id is not None:
            tests_result = await session.execute(
                select(Test).where(Test.id == test_id)
            )
        else:
            if start_date is None:
                start_date = date.today()
            if end_date is None:
                end_date = start_date
            tests_result = await session.execute(
                select(Test)
                .where(Test.test_date >= start_date)
                .where(Test.test_date <= end_date)
                .order_by(Test.test_date)
            )

        tests = tests_result.scalars().all()
        if not tests:
            await message.answer(
                "❌ За выбранный период тестов нет"
            )
            return

        test_ids = [item.id for item in tests]
        excluded_stats_ids = set(ADMIN_IDS + STATISTICIAN_IDS)

        # Получаем сотрудников без админов и роли просмотра статистики
        users_result = await session.execute(
            select(User)
            .where(User.is_whitelisted == True)
            .order_by(User.full_name)
        )
        all_users = [
            user
            for user in users_result.scalars().all()
            if user.telegram_id not in excluded_stats_ids
        ]

        if not all_users:
            await message.answer(
                "❌ Сотрудников не найдено"
            )
            return

        test_types = (
            [(test_type, "Контроль" if test_type == "control" else "Опыт")]
            if test_type
            else [
                ("control", "Контроль"),
                ("experience", "Опыт"),
            ]
        )

        user_stats = {
            user.id: {
                "user": user,
                "types": {
                    current_type: {
                        "total_questions": 0,
                        "correct_answers": 0,
                        "attempts": 0,
                        "completed_tests": 0,
                        "status": "❌ Не начат",
                    }
                    for current_type, _ in test_types
                },
            }
            for user in all_users
        }

        for current_type, _ in test_types:
            sessions_result = await session.execute(
                select(TestSession, Test)
                .join(Test, Test.id == TestSession.test_id)
                .where(TestSession.test_id.in_(test_ids))
                .where(TestSession.test_type == current_type)
            )
            test_sessions = sessions_result.all()

            for ts, current_test in test_sessions:
                if ts.user_id not in user_stats:
                    continue

                stats = user_stats[ts.user_id]["types"][current_type]
                stats["status"] = "✅ Завершен" if ts.completed else "⏳ В процессе"
                stats["attempts"] += ts.attempts or 1
                if ts.completed:
                    stats["completed_tests"] += 1

                # Для ошибок учитываем ответы всех попыток, а не только
                # последней успешной. Иначе 6 попыток с ошибками в первых
                # пяти выглядели бы как «6 попыток, 0 ошибок».
                attempt_query = select(TestAttempt).where(
                    TestAttempt.session_id == ts.id,
                )
                attempt_result = await session.execute(
                    attempt_query.order_by(TestAttempt.attempt_number)
                )
                attempts = attempt_result.scalars().all()

                for attempt in attempts:
                    answers_result = await session.execute(
                        select(Answer, Question)
                        .join(Question, Question.id == Answer.question_id)
                        .where(Answer.attempt_id == attempt.id)
                        .where(Question.test_id == current_test.id)
                        .where(Question.test_type == current_type)
                    )
                    answers_data = answers_result.all()

                    for answer, question in answers_data:
                        stats["total_questions"] += 1
                        if answer.is_correct:
                            stats["correct_answers"] += 1

        selected_label = (
            format_russian_date(tests[0].test_date)
            if len(tests) == 1
            else f"{format_russian_date(tests[0].test_date)} — {format_russian_date(tests[-1].test_date)}"
        )
        text = "📊 Статистика\n" if test_type is None else f"📊 Статистика - {test_types[0][1]}\n"
        text += f"Период: {selected_label}\n"
        text += "=" * 40 + "\n\n"

        users_by_city = defaultdict(list)
        for user_data in user_stats.values():
            users_by_city[user_data["user"].city or "Без города"].append(user_data)

        for city in sorted(users_by_city):
            text += f"🏙 <b>{escape(city)}</b>\n"
            for user_data in sorted(users_by_city[city], key=lambda item: item["user"].full_name):
                user = user_data["user"]
                text += f"   👤 {escape(user.full_name)}\n"

                for current_type, type_name in test_types:
                    stats = user_data["types"][current_type]
                    total = stats["total_questions"]
                    correct = stats["correct_answers"]
                    errors = total - correct
                    percent = (correct / total * 100) if total > 0 else 0
                    type_icon = "📋" if current_type == "control" else "🧪"

                    text += f"      {type_icon} <b>{type_name}</b>: точность {percent:.0f}%, ошибок {errors}, попыток {stats['attempts']}\n"
                    text += f"      Завершено тестов: {stats['completed_tests']}\n"

                    text += f"      {stats['status']}\n"

            text += "-" * 40 + "\n"

        # Telegram ограничивает сообщение 4096 символами. Разбиваем по строкам,
        # чтобы не обрезать HTML-тег <b> посередине и не получать Bad Request.
        chunks = []
        current_chunk = ""
        for line in text.splitlines(keepends=True):
            if current_chunk and len(current_chunk) + len(line) > 3900:
                chunks.append(current_chunk)
                current_chunk = ""
            current_chunk += line
        if current_chunk:
            chunks.append(current_chunk)

        for chunk in chunks:
            await message.answer(chunk)




# =========================

async def notification_handler(message: Message):

    if not await is_admin_or_moderator(message.from_user.id):
        await message.answer(
            "⛔ У вас нет доступа"
        )
        return

    async with async_session() as session:

        # Получаем сегодняшний тест без проверки approved
        result = await session.execute(
            select(Test).where(
                Test.test_date == date.today(),
                Test.active == True,
            )
        )
        test = result.scalars().first()

        if not test:
            await message.answer(
                "❌ Тест не найден"
            )
            return

        # Проверяем, одобрен ли тест
        if test.approved:
            await message.answer(
                "ℹ️ Тест уже одобрен и опубликован"
            )
            return

        kb = InlineKeyboardBuilder()
        kb.button(text="👀 Проверить контроль", callback_data="preview_test_type:control")
        kb.button(text="👀 Проверить опыт", callback_data="preview_test_type:experience")
        kb.button(text="📢 Отправить уведомление", callback_data="publish_test_notification")
        kb.button(text="❌ Отмена", callback_data="cancel_test_notification")
        kb.adjust(1)

        await message.answer(
            "Перед отправкой уведомления проверьте тест.\n\n"
            "После отправки уведомления редактирование вопросов будет закрыто.",
            reply_markup=kb.as_markup(),
        )


@dp.callback_query(F.data == "cancel_test_notification")
async def cancel_test_notification(callback: CallbackQuery):
    await callback.answer()
    await callback.message.delete()


@dp.callback_query(F.data == "publish_test_notification")
async def publish_test_notification(callback: CallbackQuery):
    if not await is_admin_or_moderator(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    await callback.answer()
    await callback.message.delete()

    async with async_session() as session:
        result = await session.execute(
            select(Test).where(
                Test.test_date == date.today(),
                Test.active == True,
            )
        )
        test = result.scalars().first()

        if not test:
            await callback.message.answer(
                "❌ Тест не найден"
            )
            return

        if test.approved:
            await callback.message.answer(
                "ℹ️ Тест уже одобрен и опубликован"
            )
            return

        test.approved = True
        test.approved_at = datetime.now(timezone.utc)
        await session.commit()

        service_role_ids = set(
            ADMIN_IDS + MODERATOR_IDS + STATISTICIAN_IDS
        )

        users_result = await session.execute(
            select(User)
        )
        users = [
            user
            for user in users_result.scalars().all()
            if user.telegram_id not in service_role_ids
        ]

        success_count = 0
        for user in users:
            try:
                await bot.send_message(
                    user.telegram_id,
                    f"🔔 Тесты открыты для прохождения!\n\n"
                    f"📋 Контроль качества №{test.id}\n"
                    f"Нажмите '📝 Пройти тест' чтобы начать"
                )
                success_count += 1
            except:
                pass

        service_success_count = 0
        for telegram_id in service_role_ids:
            try:
                await bot.send_message(
                    telegram_id,
                    "🔔 Тесты открыты для прохождения!"
                )
                service_success_count += 1
            except:
                pass

        await callback.message.answer(
            f"✅ Тест одобрен!\n\n"
            f"📨 Уведомления отправлены: {success_count} сотрудникам\n"
            f"🔔 Служебные уведомления отправлены: {service_success_count}"
        )


# =========================

@dp.message(Command("export"))
@dp.message(F.text == "📁 Экспорт Excel")
async def export_handler(message: Message):

    if not await can_view_reports(message.from_user.id):
        await message.answer(
            "⛔ У вас нет доступа"
        )
        return

    kb = InlineKeyboardBuilder()
    kb.button(text="📅 Сегодня", callback_data="export_date:today")
    kb.button(text="↩️ Вчера", callback_data="export_date:yesterday")
    kb.button(text="🗓 Доступные дни", callback_data="export_available_days")
    kb.adjust(1)

    await message.answer(
        "Выберите день для выгрузки Excel:",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data.startswith("export_date:"))
async def handle_export_date(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    period = callback.data.split(":")[1]
    export_date = date.today()

    if period == "yesterday":
        export_date = date.today() - timedelta(days=1)

    await callback.answer()
    await callback.message.delete()

    async with async_session() as session:
        result = await session.execute(
            select(Test).where(
                Test.test_date == export_date,
            )
        )
        test = result.scalar_one_or_none()

    if not test:
        await callback.message.answer(
            f"❌ Тест за {format_russian_date(export_date)} не найден"
        )
        return

    await export_test_results(callback.message, test.id)


@dp.callback_query(F.data == "export_available_days")
async def handle_export_available_days(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    await callback.answer()
    await callback.message.delete()

    async with async_session() as session:
        result = await session.execute(
            select(Test)
            .order_by(Test.test_date.desc())
            .limit(30)
        )
        tests = result.scalars().all()

    if not tests:
        await callback.message.answer(
            "❌ Доступных дней для экспорта нет"
        )
        return

    kb = InlineKeyboardBuilder()
    for test in tests:
        kb.button(
            text=format_russian_date(test.test_date),
            callback_data=f"export_test:{test.id}",
        )
    kb.adjust(1)

    await callback.message.answer(
        "Выберите день для выгрузки Excel:",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data.startswith("export_test:"))
async def handle_export_test(callback: CallbackQuery):
    if not await can_view_reports(callback.from_user.id):
        await callback.answer(
            "У вас нет доступа",
            show_alert=True,
        )
        return

    test_id = int(callback.data.split(":")[1])

    await callback.answer()
    await callback.message.delete()

    await export_test_results(callback.message, test_id)


async def export_test_results(message: Message, test_id: int):
    async with async_session() as session:
        result = await session.execute(
            select(Test).where(
                Test.id == test_id,
            )
        )
        test = result.scalar_one_or_none()

        if not test:
            await message.answer(
                "❌ Тест не найден"
            )
            return

        # Получаем только завершённые тестовые сессии
        sessions_result = await session.execute(
            select(TestSession).where(
                TestSession.test_id == test.id,
                TestSession.completed == True,
            )
        )
        test_sessions = sessions_result.scalars().all()

        if not test_sessions:
            await message.answer(
                "❌ Нет данных"
            )
            return

        # Структурируем данные для листов сотрудников и сводки ошибок
        user_data = {}
        mistakes_summary = {}

        for ts in test_sessions:
            result_user = await session.execute(
                select(User).where(
                    User.id == ts.user_id,
                    User.is_whitelisted == True,
                )
            )
            user = result_user.scalar_one_or_none()
            if not user:
                continue

            if user.id not in user_data:
                user_data[user.id] = {
                    'user': user,
                    'sessions': {},
                    'attempts': {},
                    'attempt_history': {},
                }

            test_type = ts.test_type

            # Пропускаем, если уже обработали этот тип для этого пользователя
            if test_type in user_data[user.id]['sessions']:
                continue

            user_data[user.id]['sessions'][test_type] = {}
            user_data[user.id]['attempts'][test_type] = ts.attempts
            user_data[user.id]['attempt_history'][test_type] = {}

            attempts_result = await session.execute(
                select(TestAttempt)
                .where(TestAttempt.session_id == ts.id)
                .order_by(TestAttempt.attempt_number)
            )
            attempts = attempts_result.scalars().all()

            passed_attempt = None

            for attempt in attempts:
                if attempt.passed:
                    passed_attempt = attempt

                answers_result = await session.execute(
                    select(Answer, Question)
                    .join(Question, Question.id == Answer.question_id)
                    .where(Answer.attempt_id == attempt.id)
                    .where(Question.test_id == test.id)
                    .where(Question.test_type == test_type)
                    .order_by(Question.question_order)
                )
                answers = answers_result.all()

                attempt_answers = {}
                for answer, question in answers:
                    q_order = question.question_order
                    attempt_answers[q_order] = {
                        "answer": answer.answer,
                        "is_correct": answer.is_correct,
                    }

                    summary_key = (question.test_type, q_order)
                    if summary_key not in mistakes_summary:
                        mistakes_summary[summary_key] = {
                            "test_type": question.test_type,
                            "question_order": q_order,
                            "mistakes": 0,
                            "total": 0,
                            "users": set(),
                        }

                    mistakes_summary[summary_key]["total"] += 1
                    if not answer.is_correct:
                        mistakes_summary[summary_key]["mistakes"] += 1
                        mistakes_summary[summary_key]["users"].add(
                            user.full_name
                        )

                user_data[user.id]['attempt_history'][test_type][
                    attempt.attempt_number
                ] = attempt_answers

            if passed_attempt:
                final_answers_result = await session.execute(
                    select(Answer, Question)
                    .join(Question, Question.id == Answer.question_id)
                    .where(Answer.attempt_id == passed_attempt.id)
                    .where(Question.test_id == test.id)
                    .where(Question.test_type == test_type)
                    .order_by(Question.question_order)
                )
                final_answers = final_answers_result.all()

                for answer, question in final_answers:
                    q_order = question.question_order
                    user_data[user.id]['sessions'][test_type][q_order] = answer.answer

        # История по всем доступным дням нужна для настоящего тренда,
        # а не только для сравнения попыток внутри одного теста.
        trend_rows = []
        trend_result = await session.execute(
            select(TestSession, Test, User)
            .join(Test, Test.id == TestSession.test_id)
            .join(User, User.id == TestSession.user_id)
            .where(TestSession.completed == True)
            .where(User.is_whitelisted == True)
            .order_by(Test.test_date, User.full_name, TestSession.test_type)
        )
        service_ids = set(ADMIN_IDS + STATISTICIAN_IDS)
        for trend_session, trend_test, trend_user in trend_result.all():
            if trend_user.telegram_id in service_ids:
                continue

            trend_attempt_result = await session.execute(
                select(TestAttempt)
                .where(
                    TestAttempt.session_id == trend_session.id,
                    TestAttempt.passed == True,
                )
                .order_by(TestAttempt.attempt_number.desc())
            )
            trend_attempt = trend_attempt_result.scalars().first()
            if not trend_attempt:
                continue

            trend_answers_result = await session.execute(
                select(Answer).where(Answer.attempt_id == trend_attempt.id)
            )
            trend_answers = trend_answers_result.scalars().all()
            total = len(trend_answers)
            correct = sum(int(answer.is_correct) for answer in trend_answers)
            trend_rows.append({
                "date": trend_test.test_date,
                "city": trend_user.city or "Без города",
                "user": trend_user.full_name or "Без имени",
                "type": get_test_type_name(trend_session.test_type),
                "accuracy": correct / total if total else 0,
                "errors": total - correct,
                "attempts": trend_session.attempts or 1,
            })

        # Создаем Excel файл
        if not user_data:
            await message.answer(
                "вќЊ РќРµС‚ РґР°РЅРЅС‹С…"
            )
            return

        filename = f"results_{test.id}_{test.test_date.strftime('%d_%m_%Y')}.xlsx"
        test_date_label = format_russian_date(test.test_date)
        wb = Workbook()
        # Удаляем пустой лист только если будут созданы новые
        default_sheet = wb.active
        sheets_created = False

        # Стили
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        center_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        final_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        attempts_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        attempt_title_fill = PatternFill(start_color="B4A7D6", end_color="B4A7D6", fill_type="solid")
        correct_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        incorrect_fill = PatternFill(start_color="D99694", end_color="D99694", fill_type="solid")
        summary_warning_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        summary_ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        def make_unique_sheet_name(base_name: str) -> str:
            clean_name = "".join(
                ch for ch in base_name if ch not in r'[]:*?/\\'
            ).strip() or "User"
            clean_name = clean_name[:31]
            if clean_name not in wb.sheetnames:
                return clean_name

            counter = 2
            while True:
                suffix = f" {counter}"
                candidate = f"{clean_name[:31 - len(suffix)]}{suffix}"
                if candidate not in wb.sheetnames:
                    return candidate
                counter += 1

        def write_header_row(ws, row_number: int, headers: list[str]):
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=row_number, column=col)
                cell.value = header
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_alignment
                cell.border = border

        # Создаем лист для каждого пользователя
        for user_id in sorted(user_data.keys()):
            user = user_data[user_id]['user']
            sessions = user_data[user_id]['sessions']

            # Ограничиваем длину названия листа до 31 символа
            sheet_name = make_unique_sheet_name(
                user.full_name if user.full_name else "User"
            )
            
            # Если это первый лист, используем дефолтный, иначе создаём новый
            if not sheets_created:
                ws = default_sheet
                ws.title = sheet_name
                sheets_created = True
            else:
                ws = wb.create_sheet(title=sheet_name)

            # Заголовок
            ws['A1'] = f"Контроль качества №{test.id} {test_date_label}"
            ws['A2'] = user.city if user.city else ""

            # Объединяем ячейки для заголовка
            ws.merge_cells('A1:D1')
            ws.merge_cells('A2:D2')
            ws['A1'].font = Font(bold=True, size=12)
            ws['A2'].font = Font(bold=True, size=11)

            # Заголовки таблицы
            write_header_row(
                ws,
                4,
                ["№", "Контроль", "Опыт", "Результат"],
            )

            # Данные
            row = 5
            control_answers = sessions.get('control', {})
            experience_answers = sessions.get('experience', {})

            all_questions = sorted(set(
                list(control_answers.keys()) + 
                list(experience_answers.keys())
            ))

            for q_num in all_questions:
                ws.cell(row=row, column=1, value=q_num)
                
                control_answer = control_answers.get(q_num, "")
                exp_answer = experience_answers.get(q_num, "")

                ws.cell(row=row, column=2, value=control_answer if control_answer else "")
                ws.cell(row=row, column=3, value=exp_answer if exp_answer else "")
                
                # Результат - пусто для каждой строки
                ws.cell(row=row, column=4, value="")

                # Форматирование
                for col in range(1, 5):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = center_alignment
                    cell.border = border

                row += 1

            # Строка "Итог"
            itog_row = row
            last_question_row = row - 1
            ws.cell(row=itog_row, column=1, value="Итог")
            
            # Сумма Контроля
            ws.cell(row=itog_row, column=2, value=f"=SUM(B5:B{last_question_row})")
            # Сумма Опыта
            ws.cell(row=itog_row, column=3, value=f"=SUM(C5:C{last_question_row})")
            
            # Результат в строке Итог = сумма опыта / сумма контроля
            result_formula = f"=IF(SUM(B5:B{last_question_row})=0,0,SUM(C5:C{last_question_row})/SUM(B5:B{last_question_row}))"
            ws.cell(row=itog_row, column=4, value=result_formula)

            # Форматирование строки Итог
            ws.cell(row=itog_row, column=1).font = Font(bold=True)
            for col in range(1, 5):
                cell = ws.cell(row=itog_row, column=col)
                cell.fill = final_fill
                cell.border = border
                cell.alignment = center_alignment

            attempts_row = itog_row + 1
            ws.cell(row=attempts_row, column=1, value="Попыток")
            ws.cell(row=attempts_row, column=2, value=user_data[user_id]['attempts'].get('control', ""))
            ws.cell(row=attempts_row, column=3, value=user_data[user_id]['attempts'].get('experience', ""))
            ws.cell(row=attempts_row, column=4, value="")

            for col in range(1, 5):
                cell = ws.cell(row=attempts_row, column=col)
                cell.fill = attempts_fill
                cell.border = border
                cell.alignment = center_alignment
                cell.font = Font(bold=True)

            attempt_history = user_data[user_id]['attempt_history']
            all_attempt_numbers = sorted(set(
                list(attempt_history.get('control', {}).keys()) +
                list(attempt_history.get('experience', {}).keys())
            ))

            row = attempts_row + 3
            for attempt_number in all_attempt_numbers:
                ws.merge_cells(
                    start_row=row,
                    start_column=1,
                    end_row=row,
                    end_column=3,
                )
                title_cell = ws.cell(row=row, column=1)
                title_cell.value = f"Попытка #{attempt_number}"
                title_cell.fill = attempt_title_fill
                title_cell.alignment = center_alignment
                title_cell.border = border
                for col in range(2, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = attempt_title_fill
                    cell.border = border

                row += 1
                write_header_row(ws, row, ["№", "Контроль", "Опыт"])

                control_attempt = attempt_history.get('control', {}).get(
                    attempt_number,
                    {},
                )
                experience_attempt = attempt_history.get('experience', {}).get(
                    attempt_number,
                    {},
                )
                attempt_questions = sorted(set(
                    list(control_attempt.keys()) +
                    list(experience_attempt.keys())
                ))

                row += 1
                for q_num in attempt_questions:
                    ws.cell(row=row, column=1, value=q_num)

                    control_data = control_attempt.get(q_num)
                    experience_data = experience_attempt.get(q_num)

                    if control_data:
                        cell = ws.cell(
                            row=row,
                            column=2,
                            value=control_data["answer"],
                        )
                        cell.fill = correct_fill if control_data["is_correct"] else incorrect_fill

                    if experience_data:
                        cell = ws.cell(
                            row=row,
                            column=3,
                            value=experience_data["answer"],
                        )
                        cell.fill = correct_fill if experience_data["is_correct"] else incorrect_fill

                    for col in range(1, 4):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = center_alignment
                        cell.border = border

                    row += 1

                row += 2

            # Ширина колонок
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 12

        # Аналитика по городам и регрессия по попыткам каждого лаборанта.
        analytics_sheet_name = make_unique_sheet_name("Статистика")
        analytics_ws = wb.create_sheet(title=analytics_sheet_name, index=0)
        analytics_ws.merge_cells("A1:E1")
        analytics_ws["A1"] = f"Статистика по городам и лаборантам: тест №{test.id}"
        analytics_ws["A1"].font = Font(bold=True, size=13)

        write_header_row(
            analytics_ws,
            3,
            ["Город", "Лаборант", "Тип", "Точность", "Ошибки", "Попытки"],
        )

        metric_rows = []
        for user_id in sorted(user_data):
            report_user = user_data[user_id]["user"]
            for current_type, type_label in (("control", "Контроль"), ("experience", "Опыт")):
                history = user_data[user_id]["attempt_history"].get(current_type, {})
                total_answers = 0
                correct_answers = 0
                for attempt_answers in history.values():
                    for answer_data in attempt_answers.values():
                        total_answers += 1
                        correct_answers += int(answer_data["is_correct"])

                if not history and not user_data[user_id]["sessions"].get(current_type):
                    continue

                metric_rows.append({
                    "city": report_user.city or "Без города",
                    "user": report_user.full_name or "Без имени",
                    "type": type_label,
                    "accuracy": correct_answers / total_answers if total_answers else 0,
                    "answers": total_answers,
                    "errors": total_answers - correct_answers,
                    "attempts": len(history) or user_data[user_id]["attempts"].get(current_type, 0),
                })

        analytics_row = 4
        for item in metric_rows:
            values = [
                item["city"], item["user"], item["type"], item["accuracy"],
                item["errors"], item["attempts"],
            ]
            for column, value in enumerate(values, 1):
                cell = analytics_ws.cell(row=analytics_row, column=column, value=value)
                cell.border = border
                cell.alignment = center_alignment
            analytics_ws.cell(row=analytics_row, column=4).number_format = "0%"
            analytics_row += 1

        for column, width in (("A", 18), ("B", 28), ("C", 14), ("D", 12), ("E", 12), ("F", 12)):
            analytics_ws.column_dimensions[column].width = width

        # Сводка по городам и график всех городов.
        city_summary_start = analytics_row + 2
        analytics_ws.cell(row=city_summary_start, column=1, value="Сводка по городам")
        analytics_ws.cell(row=city_summary_start, column=1).font = Font(bold=True, size=12)
        write_header_row(
            analytics_ws,
            city_summary_start + 1,
            ["Город", "Точность", "Ошибки", "Попытки"],
        )
        city_groups = defaultdict(list)
        for item in metric_rows:
            city_groups[item["city"]].append(item)

        city_row = city_summary_start + 2
        for city in sorted(city_groups):
            items = city_groups[city]
            total_answers = sum(item["answers"] for item in items)
            errors = sum(item["errors"] for item in items)
            attempts = sum(item["attempts"] for item in items)
            accuracy = 1 - errors / total_answers if total_answers else 0
            for column, value in enumerate([city, accuracy, errors, attempts], 1):
                cell = analytics_ws.cell(row=city_row, column=column, value=value)
                cell.border = border
                cell.alignment = center_alignment
            analytics_ws.cell(row=city_row, column=2).number_format = "0%"
            city_row += 1

        if city_row >= city_summary_start + 3:
            city_chart = BarChart()
            city_chart.type = "bar"
            city_chart.style = 10
            city_chart.title = "Точность по городам"
            city_chart.y_axis.title = "Город"
            city_chart.x_axis.title = "Точность"
            city_chart_data = Reference(
                analytics_ws,
                min_col=2,
                min_row=city_summary_start + 1,
                max_row=city_row - 1,
            )
            city_chart_categories = Reference(
                analytics_ws,
                min_col=1,
                min_row=city_summary_start + 2,
                max_row=city_row - 1,
            )
            city_chart.add_data(city_chart_data, titles_from_data=True)
            city_chart.set_categories(city_chart_categories)
            city_chart.height = 7
            city_chart.width = 14
            analytics_ws.add_chart(city_chart, "H3")

        # Линии регрессии по каждой попытке лаборанта.
        regression_start = city_row + 3
        analytics_ws.cell(row=regression_start, column=1, value="Регрессия по попыткам лаборантов")
        analytics_ws.cell(row=regression_start, column=1).font = Font(bold=True, size=12)
        regression_row = regression_start + 2
        chart_row = 20
        for user_id in sorted(user_data):
            report_user = user_data[user_id]["user"]
            attempt_metrics = defaultdict(lambda: {"total": 0, "correct": 0})
            for history in user_data[user_id]["attempt_history"].values():
                for attempt_number, answers in history.items():
                    for answer_data in answers.values():
                        attempt_metrics[attempt_number]["total"] += 1
                        attempt_metrics[attempt_number]["correct"] += int(answer_data["is_correct"])

            if not attempt_metrics:
                continue

            write_header_row(
                analytics_ws,
                regression_row,
                ["Лаборант", "Попытка", "Точность", "Ошибки", "Попыток"],
            )
            data_start = regression_row + 1
            for attempt_number in sorted(attempt_metrics):
                item = attempt_metrics[attempt_number]
                accuracy = item["correct"] / item["total"] if item["total"] else 0
                values = [
                    report_user.full_name,
                    attempt_number,
                    accuracy,
                    item["total"] - item["correct"],
                    attempt_number,
                ]
                for column, value in enumerate(values, 1):
                    cell = analytics_ws.cell(row=regression_row + 1, column=column, value=value)
                    cell.border = border
                    cell.alignment = center_alignment
                analytics_ws.cell(row=regression_row + 1, column=3).number_format = "0%"
                regression_row += 1

            data_end = regression_row
            if data_end >= data_start:
                regression_chart = LineChart()
                regression_chart.title = f"Регрессия — {report_user.full_name}"
                regression_chart.y_axis.title = "Значение"
                regression_chart.x_axis.title = "Попытка"
                regression_data = Reference(
                    analytics_ws,
                    min_col=3,
                    max_col=5,
                    min_row=data_start - 1,
                    max_row=data_end,
                )
                regression_categories = Reference(
                    analytics_ws,
                    min_col=2,
                    min_row=data_start,
                    max_row=data_end,
                )
                regression_chart.add_data(regression_data, titles_from_data=True)
                regression_chart.set_categories(regression_categories)
                regression_chart.height = 7
                regression_chart.width = 14
                analytics_ws.add_chart(regression_chart, f"H{chart_row}")
                chart_row += 16

            regression_row += 2

        trend_sheet_name = make_unique_sheet_name("Тренды")
        trend_ws = wb.create_sheet(title=trend_sheet_name, index=1)
        trend_ws.merge_cells("A1:G1")
        trend_ws["A1"] = "Регрессия показателей по датам"
        trend_ws["A1"].font = Font(bold=True, size=13)
        write_header_row(
            trend_ws,
            3,
            ["Дата", "Город", "Лаборант", "Тип", "Точность", "Ошибки", "Попытки"],
        )

        trend_row = 4
        for item in trend_rows:
            values = [
                item["date"], item["city"], item["user"], item["type"],
                item["accuracy"], item["errors"], item["attempts"],
            ]
            for column, value in enumerate(values, 1):
                cell = trend_ws.cell(row=trend_row, column=column, value=value)
                cell.border = border
                cell.alignment = center_alignment
            trend_ws.cell(row=trend_row, column=1).number_format = "dd.mm.yyyy"
            trend_ws.cell(row=trend_row, column=5).number_format = "0%"
            trend_row += 1

        for column, width in (("A", 14), ("B", 18), ("C", 28), ("D", 14), ("E", 12), ("F", 12), ("G", 12)):
            trend_ws.column_dimensions[column].width = width

        trend_users = defaultdict(list)
        for item in trend_rows:
            trend_users[item["user"]].append(item)

        chart_row = 3
        for user_name in sorted(trend_users):
            user_items = sorted(trend_users[user_name], key=lambda item: (item["date"], item["type"]))
            if len(user_items) < 2:
                continue

            chart_ws = trend_ws
            data_start = trend_row + 2
            chart_ws.cell(row=data_start - 1, column=9, value=f"Данные: {user_name}")
            chart_ws.cell(row=data_start - 1, column=9).font = Font(bold=True)
            chart_ws.cell(row=data_start - 1, column=10, value="Точность")
            chart_ws.cell(row=data_start - 1, column=11, value="Ошибки")
            chart_ws.cell(row=data_start - 1, column=12, value="Попытки")
            trend_data_start = data_start
            for item in user_items:
                chart_ws.cell(row=data_start, column=9, value=item["date"])
                chart_ws.cell(row=data_start, column=10, value=item["accuracy"])
                chart_ws.cell(row=data_start, column=11, value=item["errors"])
                chart_ws.cell(row=data_start, column=12, value=item["attempts"])
                chart_ws.cell(row=data_start, column=9).number_format = "dd.mm.yyyy"
                chart_ws.cell(row=data_start, column=10).number_format = "0%"
                data_start += 1

            chart = LineChart()
            chart.title = f"Регрессия — {user_name}"
            chart.y_axis.title = "Показатель"
            chart.x_axis.title = "Дата"
            chart_data = Reference(
                chart_ws,
                min_col=10,
                max_col=12,
                min_row=trend_data_start - 1,
                max_row=data_start - 1,
            )
            chart_categories = Reference(
                chart_ws,
                min_col=9,
                min_row=trend_data_start,
                max_row=data_start - 1,
            )
            chart.add_data(chart_data, titles_from_data=True)
            chart.set_categories(chart_categories)
            chart.height = 7
            chart.width = 14
            trend_ws.add_chart(chart, f"O{chart_row}")
            chart_row += 16
            trend_row = data_start

        summary_sheet_name = make_unique_sheet_name("Ошибки по вопросам")
        ws = wb.create_sheet(title=summary_sheet_name)

        ws.merge_cells('A1:F1')
        ws['A1'] = f"Сводка ошибок по тесту №{test.id} {test_date_label}"
        ws['A1'].font = Font(bold=True, size=12)
        ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

        write_header_row(
            ws,
            3,
            ["№ вопроса", "Тип", "Ошибок", "Всего ответов", "% ошибок", "Кто ошибся"],
        )

        summary_rows = sorted(
            mistakes_summary.values(),
            key=lambda item: (
                -item["mistakes"],
                item["test_type"],
                item["question_order"],
            )
        )

        row = 4
        for item in summary_rows:
            total = item["total"]
            mistakes = item["mistakes"]
            percent = mistakes / total if total else 0
            users = ", ".join(sorted(item["users"]))

            ws.cell(row=row, column=1, value=item["question_order"])
            ws.cell(row=row, column=2, value=get_test_type_name(item["test_type"]))
            ws.cell(row=row, column=3, value=mistakes)
            ws.cell(row=row, column=4, value=total)
            ws.cell(row=row, column=5, value=percent)
            ws.cell(row=row, column=6, value=users)

            fill = summary_warning_fill if mistakes else summary_ok_fill
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.fill = fill
                cell.border = border
                cell.alignment = center_alignment if col < 6 else Alignment(wrap_text=True, vertical="center")
            ws.cell(row=row, column=5).number_format = "0%"
            row += 1

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 50

        wb.save(filename)

        document = FSInputFile(filename)

        await message.answer_document(
            document=document,
            caption=f"📊 Отчет по тесту №{test.id}",
        )


# =========================
# TEXT HANDLER
# =========================

@dp.message()
async def text_handler(
    message: Message,
    state: FSMContext,
):

    current_state = await state.get_state()

    if not current_state:
        return

    if current_state == TestStates.passing_test.state:
        if not message.text:
            return

        if not message.text.isdigit():
            await message.answer(
                "Введите ответ числом"
            )
            return

        data = await state.get_data()
        question_id = data.get("current_question_id")

        if not question_id:
            await message.answer(
                "❌ Не удалось определить текущий вопрос. Нажмите «Пройти тест» еще раз"
            )
            return

        answer = int(message.text)

        await state.update_data(
            pending_question_id=question_id,
            pending_answer=answer,
        )

        await message.delete()

        kb = InlineKeyboardBuilder()
        kb.button(text="✅ Подтвердить", callback_data="confirm_test_answer:yes")
        kb.button(text="❌ Изменить", callback_data="confirm_test_answer:no")
        kb.adjust(2)

        await message.answer(
            f"Ваш ответ: <b>{answer}</b>\n\nПодтвердить?",
            reply_markup=kb.as_markup(),
            parse_mode=ParseMode.HTML,
        )

        return

    # ДОБАВЛЕНИЕ ВОПРОСА
    if current_state == QuestionAddStates.waiting_answer.state:

        if not message.text:
            return

        if not message.text.isdigit():
            data = await state.get_data()
            if data.get("instruction_message_id"):
                try:
                    await message.bot.delete_message(
                        message.chat.id,
                        data["instruction_message_id"]
                    )
                except:
                    pass

            msg = await message.answer("Введите число")
            await state.update_data(
                instruction_message_id=msg.message_id
            )
            return

        answer = int(message.text)

        data = await state.get_data()

        async with async_session() as session:
            if await is_test_approved(session, data["test_id"]):
                await state.clear()
                await message.answer(
                    "🔐 Тест уже опубликован, добавление вопросов закрыто"
                )
                return
        
        # Сохраняем ответ в state и переходим на подтверждение
        await state.update_data(correct_answer=answer)
        await state.set_state(QuestionAddStates.confirming_question)
        
        await message.delete()
        
        # Удаляем инструкционное сообщение
        if data.get("instruction_message_id"):
            try:
                await message.bot.delete_message(
                    message.chat.id,
                    data["instruction_message_id"]
                )
            except:
                pass
        
        # Показываем превью вопроса с фото и ответом
        kb = InlineKeyboardBuilder()
        kb.button(text="✅ Подтвердить", callback_data="confirm_question:yes")
        kb.button(text="❌ Переснять", callback_data="confirm_question:no")
        kb.adjust(2)
        
        await message.answer_photo(
            photo=data["photo_id"],
            caption=f"📋 Правильный ответ: <b>{answer}</b>\n\n"
                   f"Всё правильно?",
            reply_markup=kb.as_markup(),
            parse_mode=ParseMode.HTML
        )

        return

    # РЕДАКТИРОВАНИЕ ОТВЕТА
    if current_state == AdminEditStates.editing_answer.state:

        if not message.text:
            return

        if not message.text.isdigit():
            data = await state.get_data()
            if data.get("instruction_message_id"):
                try:
                    await message.bot.delete_message(
                        message.chat.id,
                        data["instruction_message_id"]
                    )
                except:
                    pass

            msg = await message.answer("Введите число")
            await state.update_data(
                instruction_message_id=msg.message_id
            )
            return

        answer = int(message.text)

        data = await state.get_data()

        async with async_session() as session:

            result = await session.execute(
                select(Question, Test)
                .join(Test, Test.id == Question.test_id)
                .where(Question.id == data["edit_question_id"])
            )

            row = result.first()

            if not row:
                await state.clear()
                await message.answer("❌ Вопрос не найден")
                return

            q, test = row

            if test.approved:
                await state.clear()
                await message.answer(
                    "🔐 Тест уже опубликован, редактирование закрыто"
                )
                return

            q.correct_answer = answer

            await session.commit()

        await state.clear()

        await message.delete()

        # Удаляем инструкционное сообщение
        if data.get("instruction_message_id"):
            try:
                await message.bot.delete_message(
                    message.chat.id,
                    data["instruction_message_id"]
                )
            except:
                pass

        await message.answer(
            "✅ Ответ обновлён"
        )

        return


# =========================
# MAIN
# =========================

async def setup_bot_commands():
    commands = [
        BotCommand(command="start", description="Начать"),
    ]

    for attempt in range(1, 4):
        try:
            await bot.set_my_commands(commands)
            return
        except Exception as e:
            logger.warning(
                "⚠️ Failed to set bot commands "
                f"(attempt {attempt}/3): {e}"
            )
            await asyncio.sleep(attempt)

    logger.warning(
        "⚠️ Bot commands were not set, continuing startup"
    )


async def main():

    await create_db()
    await setup_bot_commands()

    retry_delay = 5

    while True:
        try:
            logger.info("🚀 Bot started, polling messages...")
            await dp.start_polling(bot)
            break
        except (TelegramNetworkError, TelegramServerError) as e:
            logger.warning(
                "⚠️ Telegram server error during polling startup: "
                f"{e}. Retrying in {retry_delay} seconds..."
            )
            await asyncio.sleep(retry_delay)
            retry_delay = min(retry_delay * 2, 60)


if __name__ == "__main__":
    asyncio.run(main())
