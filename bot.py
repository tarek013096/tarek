import asyncio
import logging
import os
import re
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from telegram import KeyboardButton, ReplyKeyboardMarkup, Update
from telegram.constants import ChatAction
from telegram.error import TelegramError
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

LAB_BUTTON = "Lab Report Generator"
ASSIGNMENT_BUTTON = "Assignment Generator"
BACK_BUTTON = "Back"
CLEAR_BUTTON = "Clear"
CANCEL_BUTTON = "Cancel"

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [
        [KeyboardButton(LAB_BUTTON), KeyboardButton(ASSIGNMENT_BUTTON)],
        [KeyboardButton(CLEAR_BUTTON)],
    ],
    resize_keyboard=True,
    one_time_keyboard=False,
)

FORM_KEYBOARD = ReplyKeyboardMarkup(
    [
        [KeyboardButton(BACK_BUTTON), KeyboardButton(CLEAR_BUTTON)],
        [KeyboardButton(CANCEL_BUTTON)],
    ],
    resize_keyboard=True,
    one_time_keyboard=False,
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
    ("submission_date", "Submission date likhun, example: June 29, 2026"),
]

ASK_FIELD, ASK_LOGO = range(2)


def clean_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_. -]+", "", value).strip()
    return value[:70] or "cover-page"


def remember_message(context: ContextTypes.DEFAULT_TYPE, message_id: int | None) -> None:
    if not message_id:
        return
    ids = context.chat_data.setdefault("message_ids", [])
    if message_id not in ids:
        ids.append(message_id)


async def send_tracked(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str, reply_markup=None):
    message = await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=text,
        reply_markup=reply_markup,
    )
    remember_message(context, message.message_id)
    return message


async def track_incoming(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_message:
        remember_message(context, update.effective_message.message_id)


async def delete_tracked_messages(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat_id = update.effective_chat.id
    message_ids = list(context.chat_data.get("message_ids", []))
    context.chat_data["message_ids"] = []
    for message_id in message_ids:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=message_id)
        except TelegramError:
            pass


def set_document_background(document: Document, color: str = "F7FBFC") -> None:
    background = document._element.find(qn("w:background"))
    if background is None:
        background = OxmlElement("w:background")
        document._element.insert(0, background)
    background.set(qn("w:color"), color)


def set_page_border(section) -> None:
    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:pgBorders"))
    if existing is not None:
        sect_pr.remove(existing)

    borders = OxmlElement("w:pgBorders")
    borders.set(qn("w:offsetFrom"), "page")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "thinThickThinMediumGap")
        border.set(qn("w:sz"), "22")
        border.set(qn("w:space"), "18")
        border.set(qn("w:color"), "167A8F")
        borders.append(border)
    sect_pr.append(borders)


def set_cell_width(cell, width_twips: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    width = tc_pr.find(qn("w:tcW"))
    if width is None:
        width = OxmlElement("w:tcW")
        tc_pr.append(width)
    width.set(qn("w:w"), str(width_twips))
    width.set(qn("w:type"), "dxa")


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_border(cell, color: str = "167A8F", size: str = "14", style: str = "single") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    for side in ("top", "left", "bottom", "right"):
        border = borders.find(qn(f"w:{side}"))
        if border is None:
            border = OxmlElement(f"w:{side}")
            borders.append(border)
        border.set(qn("w:val"), style)
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)


def clear_cell_border(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    for side in ("top", "left", "bottom", "right"):
        border = borders.find(qn(f"w:{side}"))
        if border is None:
            border = OxmlElement(f"w:{side}")
            borders.append(border)
        border.set(qn("w:val"), "nil")


def set_paragraph_bottom_border(paragraph, color: str = "167A8F") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = p_bdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        p_bdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "8")
    bottom.set(qn("w:color"), color)


def add_text(
    paragraph,
    text: str,
    size: int,
    bold: bool = False,
    color: RGBColor | None = None,
    font: str = "Aptos Display",
):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = font
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return run


def add_centered_text(
    document: Document,
    text: str,
    size: int,
    bold: bool = False,
    color: RGBColor | None = None,
    space_after: int = 6,
    font: str = "Aptos Display",
) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(space_after)
    add_text(paragraph, text, size=size, bold=bold, color=color, font=font)
    return paragraph


def add_spacer(document: Document, points: int) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(points)


def add_course_line(document: Document, course_code: str, course_title: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.left_indent = Cm(1.0)
    paragraph.paragraph_format.right_indent = Cm(1.0)
    paragraph.paragraph_format.space_after = Pt(5)
    add_text(paragraph, "COURSE CODE: ", 13, True, RGBColor(0, 128, 64))
    add_text(paragraph, course_code, 13, True, RGBColor(20, 20, 20), "Aptos")

    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.left_indent = Cm(1.0)
    title_paragraph.paragraph_format.right_indent = Cm(1.0)
    title_paragraph.paragraph_format.space_after = Pt(20)
    add_text(title_paragraph, "COURSE TITLE: ", 13, True, RGBColor(0, 128, 64))
    add_text(title_paragraph, course_title, 13, True, RGBColor(20, 20, 20), "Aptos")
    set_paragraph_bottom_border(title_paragraph)


def add_label_value(paragraph, label: str, value: str, show_label: bool = True) -> None:
    if show_label:
        add_text(paragraph, label, 10, True, RGBColor(0, 128, 64), "Aptos")
    add_text(paragraph, value, 10, True, RGBColor(22, 31, 38), "Aptos")


def fill_info_box(cell, heading: str, items: list[tuple[str, str, bool]]) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_shading(cell, "F1FAF7")
    set_cell_border(cell, "167A8F", "16", "single")

    heading_paragraph = cell.paragraphs[0]
    heading_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading_paragraph.paragraph_format.space_after = Pt(7)
    add_text(heading_paragraph, heading, 14, True, RGBColor(0, 128, 64), "Aptos Display")

    for label, value, show_label in items:
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.left_indent = Cm(0.12)
        paragraph.paragraph_format.right_indent = Cm(0.12)
        paragraph.paragraph_format.space_after = Pt(3)
        paragraph.paragraph_format.line_spacing = 1.0
        add_label_value(paragraph, label, value, show_label=show_label)


def add_submission_boxes(document: Document, data: dict) -> None:
    wrapper = document.add_table(rows=1, cols=3)
    wrapper.autofit = False
    left_margin_cell = wrapper.cell(0, 0)
    content_cell = wrapper.cell(0, 1)
    right_margin_cell = wrapper.cell(0, 2)
    for cell, width in ((left_margin_cell, 350), (content_cell, 8350), (right_margin_cell, 350)):
        cell.text = ""
        set_cell_width(cell, width)
        clear_cell_border(cell)
        set_cell_shading(cell, "F7FBFC")
        if cell.paragraphs:
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            cell.paragraphs[0].paragraph_format.space_before = Pt(0)

    table = content_cell.add_table(rows=1, cols=3)
    table.autofit = False
    left_cell = table.cell(0, 0)
    gap_cell = table.cell(0, 1)
    right_cell = table.cell(0, 2)

    for cell, width in ((left_cell, 3900), (gap_cell, 550), (right_cell, 3900)):
        cell.text = ""
        set_cell_width(cell, width)

    clear_cell_border(gap_cell)
    set_cell_shading(gap_cell, "F7FBFC")

    submitted_by = [
        ("NAME: ", data["student_name"], True),
        ("ID: ", data["student_id"], True),
        ("PROGRAM: ", data["program"], True),
        ("", data["batch_year_semester"], False),
    ]
    submitted_to = [
        ("NAME: ", data["teacher_name"], True),
        ("DESIGNATION: ", data["teacher_designation"], True),
        ("DEPARTMENT: ", data["teacher_department"], True),
    ]

    fill_info_box(left_cell, "SUBMITTED BY:", submitted_by)
    fill_info_box(right_cell, "SUBMITTED TO:", submitted_to)


def create_cover_page(data: dict, logo_path: Path | None) -> Path:
    document = Document()
    set_document_background(document)

    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.35)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.35)
    section.right_margin = Cm(1.35)
    section.start_type = WD_SECTION_START.NEW_PAGE
    set_page_border(section)

    green = RGBColor(0, 176, 80)
    teal = RGBColor(22, 122, 143)
    dark = RGBColor(22, 31, 38)

    add_spacer(document, 8)
    university_paragraph = add_centered_text(document, data["university_name"].upper(), 19, True, green, 8)
    university_paragraph.paragraph_format.left_indent = Cm(0.7)
    university_paragraph.paragraph_format.right_indent = Cm(0.7)

    if logo_path and logo_path.exists():
        logo_paragraph = document.add_paragraph()
        logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_paragraph.paragraph_format.space_after = Pt(15)
        try:
            logo_paragraph.add_run().add_picture(str(logo_path), width=Inches(1.35))
        except Exception:
            LOGGER.exception("Could not insert uploaded logo")
    else:
        add_spacer(document, 34)

    title = "LAB REPORT" if data["document_type"] == "lab" else "ASSIGNMENT"
    add_centered_text(document, title, 27, True, teal, 8)
    add_centered_text(document, f"{title.title()} No: {data['work_no']}", 15, True, dark, 14, "Aptos")
    add_course_line(document, data["course_code"], data["course_title"])
    add_submission_boxes(document, data)

    add_spacer(document, 28)
    add_centered_text(document, "SUBMISSION DATE:", 15, True, green, 3)
    add_centered_text(document, data["submission_date"], 13, True, dark, 6, "Aptos")

    output_dir = Path(tempfile.gettempdir()) / "cover_page_bot"
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = clean_filename(f"{title.lower()}_{data['student_id']}_{data['work_no']}")
    output_path = output_dir / f"{base_name}.docx"
    document.save(output_path)
    return output_path


def reset_form(context: ContextTypes.DEFAULT_TYPE, document_type: str | None = None) -> None:
    context.user_data.clear()
    if document_type:
        context.user_data["document_type"] = document_type
        context.user_data["field_index"] = 0


async def show_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> int:
    await send_tracked(update, context, text, MAIN_KEYBOARD)
    return ConversationHandler.END


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    reset_form(context)
    return await show_main_menu(
        update,
        context,
        "Cover page bot ready. Menu theke Lab Report ba Assignment choose korun.",
    )


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    reset_form(context)
    await delete_tracked_messages(update, context)
    message = await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Cleared. Notun kore start korte menu use korun.",
        reply_markup=MAIN_KEYBOARD,
    )
    remember_message(context, message.message_id)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    reset_form(context)
    return await show_main_menu(update, context, "Generation cancel kora hoyeche. Menu theke abar choose korun.")


async def choose_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    text = update.effective_message.text.strip()
    if text == LAB_BUTTON:
        reset_form(context, "lab")
    elif text == ASSIGNMENT_BUTTON:
        reset_form(context, "assignment")
    else:
        return await show_main_menu(update, context, "Menu theke valid option choose korun.")

    await send_tracked(update, context, FIELDS[0][1], FORM_KEYBOARD)
    return ASK_FIELD


async def back_field(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    field_index = int(context.user_data.get("field_index", 0))
    if field_index <= 0:
        await send_tracked(update, context, "Eita prothom step. Menu theke notun option choose korte paren.", FORM_KEYBOARD)
        return ASK_FIELD

    field_index -= 1
    field_name, prompt = FIELDS[field_index]
    context.user_data.pop(field_name, None)
    context.user_data["field_index"] = field_index
    await send_tracked(update, context, f"Back kora holo. {prompt}", FORM_KEYBOARD)
    return ASK_FIELD


async def collect_field(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    text = update.effective_message.text.strip()

    if text in {LAB_BUTTON, ASSIGNMENT_BUTTON}:
        return await choose_type(update, context)

    field_index = int(context.user_data.get("field_index", 0))
    field_name, _ = FIELDS[field_index]

    if not text:
        await send_tracked(update, context, "Ei field ta khali rakha jabe na. Abar likhun:", FORM_KEYBOARD)
        return ASK_FIELD

    context.user_data[field_name] = text
    field_index += 1
    context.user_data["field_index"] = field_index

    if field_index < len(FIELDS):
        await send_tracked(update, context, FIELDS[field_index][1], FORM_KEYBOARD)
        return ASK_FIELD

    await send_tracked(
        update,
        context,
        "University logo upload korun. Photo hisebe ba image document hisebe pathan.",
        FORM_KEYBOARD,
    )
    return ASK_LOGO


async def back_from_logo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    context.user_data["field_index"] = len(FIELDS) - 1
    context.user_data.pop("submission_date", None)
    await send_tracked(update, context, f"Back kora holo. {FIELDS[-1][1]}", FORM_KEYBOARD)
    return ASK_FIELD


async def collect_logo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action=ChatAction.UPLOAD_DOCUMENT)

    logo_path = None
    file_id = None
    if update.effective_message.photo:
        file_id = update.effective_message.photo[-1].file_id
        logo_path = Path(tempfile.gettempdir()) / f"logo_{update.effective_user.id}.jpg"
    elif update.effective_message.document and (update.effective_message.document.mime_type or "").startswith("image/"):
        file_id = update.effective_message.document.file_id
        extension = Path(update.effective_message.document.file_name or "logo.png").suffix or ".png"
        logo_path = Path(tempfile.gettempdir()) / f"logo_{update.effective_user.id}{extension}"

    if not file_id or not logo_path:
        await send_tracked(update, context, "Logo hisebe ekta image upload korun. Vul hole Back click korun.", FORM_KEYBOARD)
        return ASK_LOGO

    telegram_file = await context.bot.get_file(file_id)
    await telegram_file.download_to_drive(custom_path=str(logo_path))

    data = dict(context.user_data)
    missing = [field for field, _ in FIELDS if not data.get(field)]
    if missing:
        context.user_data["field_index"] = 0
        await send_tracked(update, context, "Kichu input missing. Notun kore first step theke shuru korun.", FORM_KEYBOARD)
        return ASK_FIELD

    output_path = create_cover_page(data, logo_path)
    with output_path.open("rb") as document_file:
        sent_doc = await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=document_file,
            filename=output_path.name,
            caption="Cover page ready. DOCX file ta download kore nin.",
            reply_markup=MAIN_KEYBOARD,
        )
    remember_message(context, sent_doc.message_id)
    reset_form(context)
    await send_tracked(update, context, "Abar generate korte chaile menu theke choose korun.", MAIN_KEYBOARD)
    return ConversationHandler.END


async def ask_logo_again(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await track_incoming(update, context)
    await send_tracked(update, context, "Logo image upload korun, ba Back diye submission date edit korun.", FORM_KEYBOARD)
    return ASK_LOGO


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await track_incoming(update, context)
    await send_tracked(
        update,
        context,
        "Menu theke generator choose korun. Protiti step-e Back/Clear ache. Clear old bot messages delete korar try korbe.",
        MAIN_KEYBOARD,
    )


def build_application() -> Application:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable missing.")

    app = Application.builder().token(BOT_TOKEN).build()
    conversation = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex(f"^({re.escape(LAB_BUTTON)}|{re.escape(ASSIGNMENT_BUTTON)})$"), choose_type),
        ],
        states={
            ASK_FIELD: [
                MessageHandler(filters.Regex(f"^{re.escape(BACK_BUTTON)}$"), back_field),
                MessageHandler(filters.Regex(f"^{re.escape(CLEAR_BUTTON)}$"), clear),
                MessageHandler(filters.Regex(f"^{re.escape(CANCEL_BUTTON)}$"), cancel),
                MessageHandler(filters.TEXT & ~filters.COMMAND, collect_field),
            ],
            ASK_LOGO: [
                MessageHandler(filters.Regex(f"^{re.escape(BACK_BUTTON)}$"), back_from_logo),
                MessageHandler(filters.Regex(f"^{re.escape(CLEAR_BUTTON)}$"), clear),
                MessageHandler(filters.Regex(f"^{re.escape(CANCEL_BUTTON)}$"), cancel),
                MessageHandler((filters.PHOTO | filters.Document.ALL) & ~filters.COMMAND, collect_logo),
                MessageHandler(filters.ALL & ~filters.COMMAND, ask_logo_again),
            ],
        },
        fallbacks=[
            CommandHandler("start", start),
            CommandHandler("cancel", cancel),
            MessageHandler(filters.Regex(f"^{re.escape(CLEAR_BUTTON)}$"), clear),
        ],
        allow_reentry=True,
    )

    app.add_handler(conversation)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.Regex(f"^{re.escape(CLEAR_BUTTON)}$"), clear))
    return app


def main() -> None:
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())

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
