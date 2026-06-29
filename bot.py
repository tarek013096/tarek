import logging
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt, RGBColor
from telegram import (
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)


logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
LOGGER = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "").rstrip("/")
PORT = int(os.environ.get("PORT", "10000"))

MENU_KEYBOARD = ReplyKeyboardMarkup(
    [
        [KeyboardButton("Lab Report Generator")],
        [KeyboardButton("Assignment Generator")],
    ],
    resize_keyboard=True,
)

FIELDS = [
    ("university_name", "University name likhun:"),
    ("course_code", "Course code likhun:"),
    ("course_title", "Course title likhun:"),
    ("work_no", "Assignment/Report no likhun:"),
    ("student_name", "Student er name likhun:"),
    ("student_id", "Student ID likhun:"),
    ("program", "Program likhun:"),
    ("batch_year_semester", "Batch / year / semester likhun:"),
    ("teacher_name", "Teacher er name likhun:"),
    ("teacher_designation", "Teacher er designation likhun:"),
    ("teacher_department", "Teacher er department likhun:"),
]

ASK_FIELD, ASK_LOGO = range(2)


def clean_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_. -]+", "", value).strip()
    return value[:70] or "cover-page"


def add_spacer(document: Document, points: int) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(points)


def add_centered_text(
    document: Document,
    text: str,
    size: int,
    bold: bool = False,
    color: RGBColor | None = None,
    space_after: int = 6,
) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(space_after)
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def add_label_value(paragraph, label: str, value: str) -> None:
    label_run = paragraph.add_run(label)
    label_run.bold = True
    label_run.font.name = "Times New Roman"
    label_run.font.color.rgb = RGBColor(0, 128, 64)

    value_run = paragraph.add_run(value)
    value_run.bold = True
    value_run.font.name = "Times New Roman"
    value_run.font.color.rgb = RGBColor(0, 0, 0)


def add_info_box(document: Document, heading: str, items: list[tuple[str, str]]) -> None:
    table = document.add_table(rows=1, cols=1)
    table.autofit = True
    cell = table.cell(0, 0)
    cell.text = ""

    heading_paragraph = cell.paragraphs[0]
    heading_paragraph.paragraph_format.space_after = Pt(4)
    heading_run = heading_paragraph.add_run(heading)
    heading_run.bold = True
    heading_run.font.name = "Times New Roman"
    heading_run.font.size = Pt(16)
    heading_run.font.color.rgb = RGBColor(0, 128, 64)

    for label, value in items:
        paragraph = cell.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.line_spacing = 1.0
        add_label_value(paragraph, label, value)


def create_cover_page(data: dict, logo_path: Path | None) -> Path:
    document = Document()

    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.start_type = WD_SECTION_START.NEW_PAGE

    green = RGBColor(0, 176, 80)
    accent = RGBColor(22, 122, 143)

    add_centered_text(
        document,
        data["university_name"].upper(),
        size=20,
        bold=True,
        color=green,
        space_after=10,
    )

    if logo_path and logo_path.exists():
        logo_paragraph = document.add_paragraph()
        logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_run = logo_paragraph.add_run()
        try:
            logo_run.add_picture(str(logo_path), width=Inches(1.45))
        except Exception:
            LOGGER.exception("Could not insert uploaded logo")
        logo_paragraph.paragraph_format.space_after = Pt(18)
    else:
        add_spacer(document, 36)

    title = "LAB REPORT" if data["document_type"] == "lab" else "ASSIGNMENT"
    add_centered_text(document, title, size=26, bold=True, color=accent, space_after=12)
    add_centered_text(
        document,
        f"{title.title()} No: {data['work_no']}",
        size=16,
        bold=True,
        color=RGBColor(0, 0, 0),
        space_after=18,
    )

    course = document.add_paragraph()
    course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_label_value(course, "COURSE CODE: ", data["course_code"])
    course.paragraph_format.space_after = Pt(4)

    course_title = document.add_paragraph()
    course_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_label_value(course_title, "COURSE TITLE: ", data["course_title"])
    course_title.paragraph_format.space_after = Pt(24)

    box_table = document.add_table(rows=1, cols=2)
    box_table.autofit = True
    left_cell = box_table.cell(0, 0)
    right_cell = box_table.cell(0, 1)
    left_cell.text = ""
    right_cell.text = ""

    submitted_to = [
        ("NAME: ", data["teacher_name"]),
        ("DESIGNATION: ", data["teacher_designation"]),
        ("DEPARTMENT: ", data["teacher_department"]),
    ]
    submitted_by = [
        ("NAME: ", data["student_name"]),
        ("ID: ", data["student_id"]),
        ("PROGRAM: ", data["program"]),
        ("BATCH/YEAR/SEMESTER: ", data["batch_year_semester"]),
    ]

    fill_cell_box(left_cell, "SUBMITTED TO:", submitted_to)
    fill_cell_box(right_cell, "SUBMITTED BY:", submitted_by)

    add_spacer(document, 32)
    add_centered_text(document, "SUBMISSION DATE:", size=16, bold=True, color=green, space_after=4)
    today = datetime.now(ZoneInfo("Asia/Dhaka")).strftime("%B %d, %Y")
    add_centered_text(document, today, size=14, bold=True, color=RGBColor(0, 0, 0), space_after=0)

    output_dir = Path(tempfile.gettempdir()) / "cover_page_bot"
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = clean_filename(f"{title.lower()}_{data['student_id']}_{data['work_no']}")
    output_path = output_dir / f"{base_name}.docx"
    document.save(output_path)
    return output_path


def fill_cell_box(cell, heading: str, items: list[tuple[str, str]]) -> None:
    heading_paragraph = cell.paragraphs[0]
    heading_paragraph.paragraph_format.space_after = Pt(6)
    heading_run = heading_paragraph.add_run(heading)
    heading_run.bold = True
    heading_run.font.name = "Times New Roman"
    heading_run.font.size = Pt(16)
    heading_run.font.color.rgb = RGBColor(0, 176, 80)

    for label, value in items:
        paragraph = cell.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(2)
        add_label_value(paragraph, label, value)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "Cover page generator ready. Nicher menu theke ekta option choose korun.",
        reply_markup=MENU_KEYBOARD,
    )
    return ConversationHandler.END


async def choose_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip().lower()
    if "lab" in text:
        context.user_data["document_type"] = "lab"
    elif "assignment" in text:
        context.user_data["document_type"] = "assignment"
    else:
        await update.message.reply_text("Menu theke Lab Report Generator ba Assignment Generator choose korun.")
        return ConversationHandler.END

    context.user_data["field_index"] = 0
    await update.message.reply_text(FIELDS[0][1], reply_markup=ReplyKeyboardRemove())
    return ASK_FIELD


async def collect_field(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    field_index = context.user_data.get("field_index", 0)
    field_name, _ = FIELDS[field_index]
    value = update.message.text.strip()

    if not value:
        await update.message.reply_text("Ei field ta khali rakha jabe na. Abar likhun:")
        return ASK_FIELD

    context.user_data[field_name] = value
    field_index += 1
    context.user_data["field_index"] = field_index

    if field_index < len(FIELDS):
        await update.message.reply_text(FIELDS[field_index][1])
        return ASK_FIELD

    await update.message.reply_text(
        "University logo upload korun. Photo hisebe pathate paren, ba image file document hisebe pathate paren."
    )
    return ASK_LOGO


async def collect_logo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.chat.send_action(ChatAction.UPLOAD_DOCUMENT)

    logo_path = None
    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        logo_path = Path(tempfile.gettempdir()) / f"logo_{update.effective_user.id}.jpg"
    elif update.message.document and (update.message.document.mime_type or "").startswith("image/"):
        file_id = update.message.document.file_id
        extension = Path(update.message.document.file_name or "logo.png").suffix or ".png"
        logo_path = Path(tempfile.gettempdir()) / f"logo_{update.effective_user.id}{extension}"
    else:
        await update.message.reply_text("Please ekta image/logo upload korun.")
        return ASK_LOGO

    telegram_file = await context.bot.get_file(file_id)
    await telegram_file.download_to_drive(custom_path=str(logo_path))

    output_path = create_cover_page(dict(context.user_data), logo_path)
    caption = "Cover page ready. DOCX file ta download kore nin."
    with output_path.open("rb") as document_file:
        await update.message.reply_document(document=document_file, filename=output_path.name, caption=caption)

    await update.message.reply_text("Abar generate korte chaile menu theke choose korun.", reply_markup=MENU_KEYBOARD)
    context.user_data.clear()
    return ConversationHandler.END


async def ask_logo_again(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Logo hisebe ekta image upload korun, tarpor DOCX generate hobe.")
    return ASK_LOGO


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Generation cancel kora hoyeche.", reply_markup=MENU_KEYBOARD)
    return ConversationHandler.END


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Menu theke Lab Report Generator ba Assignment Generator choose korun. "
        "Tarpor bot step by step info and logo nibe."
    )


def build_application() -> Application:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable missing.")

    app = Application.builder().token(BOT_TOKEN).build()
    conversation = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex("^(Lab Report Generator|Assignment Generator)$"), choose_type),
        ],
        states={
            ASK_FIELD: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_field)],
            ASK_LOGO: [
                MessageHandler((filters.PHOTO | filters.Document.ALL) & ~filters.COMMAND, collect_logo),
                MessageHandler(filters.ALL & ~filters.COMMAND, ask_logo_again),
            ],
        },
        fallbacks=[
            CommandHandler("start", start),
            CommandHandler("cancel", cancel),
        ],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(conversation)
    return app


def main() -> None:
    app = build_application()
    if WEBHOOK_URL:
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            url_path=BOT_TOKEN,
            webhook_url=f"{WEBHOOK_URL}/{BOT_TOKEN}",
        )
    else:
        app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
