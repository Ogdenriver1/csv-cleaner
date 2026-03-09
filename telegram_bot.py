"""
CSV Cleaner Telegram Bot
Send messy CSV data or a .csv file, get back a Numbers-style pivot table Excel file!

Setup:
1. pip install -r requirements.txt
2. Set your environment variables:
   export TELEGRAM_BOT_TOKEN="your-token-from-botfather"
   export ANTHROPIC_API_KEY="your-anthropic-key"
3. python telegram_bot.py
"""

import os
import io
import csv
import json
import asyncio
import logging
import pandas as pd
from telegram import Bot, Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

def send_file(file):
    """Send a file to a Telegram chat (synchronous, for use from data_cleaner.py)."""
    bot = Bot(token=TOKEN)
    import asyncio
    async def _send():
        with open(file, "rb") as f:
            await bot.send_document(chat_id=CHAT_ID, document=f)
    asyncio.run(_send())

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

anthropic_client = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY")
)

# Apple Numbers color palette
NUMBERS_HEADER_BG   = "1B6EC2"   # Numbers blue header
NUMBERS_HEADER_TEXT = "FFFFFF"
NUMBERS_ROW_ALT     = "EEF4FB"   # Light blue alternating row
NUMBERS_ROW_WHITE   = "FFFFFF"
NUMBERS_TOTAL_BG    = "D0E4F5"   # Slightly darker blue for totals row
NUMBERS_TOTAL_TEXT  = "0D3C6B"
NUMBERS_BORDER      = "C5D8EE"   # Soft blue border
NUMBERS_TITLE_BG    = "0D3C6B"   # Dark navy for pivot title bar
NUMBERS_TITLE_TEXT  = "FFFFFF"


def clean_and_plan_pivot(messy_data: str) -> dict:
    """
    Ask Claude to clean the data AND return a pivot table plan.
    Returns { "csv": "...", "pivot": { "index": [...], "columns": [...], "values": [...], "aggfunc": "..." } }
    """
    message = anthropic_client.messages.create(
        model="claude-3-haiku-20240307",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": f"""You are a data analyst. Do two things:

1. Clean this messy data into proper CSV format:
   - Standardize names to Title Case
   - Fix inconsistent delimiters (|, -, tabs → commas)
   - Make numbers plain (no $, no k abbreviations: 50k → 50000)
   - Add a header row if missing
   - Remove empty rows, trim whitespace

2. Suggest the best pivot table configuration for this data, like Apple Numbers would create.
   Choose which columns make sense as:
   - index: row grouping fields (categorical columns, e.g. Department, Region, Category)
   - columns: column grouping field (optional, one column max, e.g. Year, Quarter, Status) — use null if none fits
   - values: numeric columns to aggregate (e.g. Sales, Salary, Amount)
   - aggfunc: "sum", "mean", or "count" — pick whichever makes most sense for the data

Return ONLY valid JSON in this exact format, nothing else:
{{
  "csv": "header1,header2,...\\nval1,val2,...\\n...",
  "pivot": {{
    "index": ["ColumnA"],
    "columns": "ColumnB",
    "values": ["ColumnC"],
    "aggfunc": "sum"
  }}
}}

Data:
{messy_data}"""
            }
        ]
    )

    raw = message.content[0].text.strip()
    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def build_pivot_excel(csv_data: str, pivot_config: dict) -> io.BytesIO:
    """Build a Numbers-style pivot table Excel file from clean CSV + pivot config."""

    # --- Parse CSV into DataFrame robustly ---
    df = pd.read_csv(io.StringIO(csv_data.strip()), on_bad_lines='skip')
    df.columns = [str(c).strip() for c in df.columns]

    # Convert numeric value columns
    valid_values = []
    for col in pivot_config["values"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(r'[$,]', '', regex=True).str.strip(),
                errors='coerce'
            )
            valid_values.append(col)

    # Fall back to any numeric column if Claude's suggestion doesn't match
    if not valid_values:
        valid_values = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    pivot_config["values"] = valid_values

    # Filter index/columns to only existing columns
    pivot_config["index"] = [c for c in pivot_config["index"] if c in df.columns]
    if not pivot_config["index"]:
        pivot_config["index"] = [df.columns[0]]

    if pivot_config.get("columns") and pivot_config["columns"] not in df.columns:
        pivot_config["columns"] = None

    # --- Build pivot table ---
    pivot_kwargs = {
        "index": pivot_config["index"],
        "values": pivot_config["values"],
        "aggfunc": pivot_config["aggfunc"],
        "margins": True,
        "margins_name": "Total",
    }
    if pivot_config.get("columns"):
        pivot_kwargs["columns"] = pivot_config["columns"]

    pivot = pd.pivot_table(df, **pivot_kwargs)

    # Flatten multi-level column index if present
    if isinstance(pivot.columns, pd.MultiIndex):
        pivot.columns = [" · ".join(str(c) for c in col if str(c) != 'nan').strip() for col in pivot.columns]
    pivot = pivot.reset_index()

    # For the raw sheet, use the DataFrame rows
    raw_rows = [list(df.columns)] + df.values.tolist()

    # --- Write to Excel with Numbers styling ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pivot Table"

    def border(color=NUMBERS_BORDER):
        s = Side(style='thin', color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    def fill(hex_color):
        return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    num_cols = len(pivot.columns)
    num_rows = len(pivot)

    # --- Title bar (row 1) ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1)
    agg_label = {"sum": "Sum", "mean": "Average", "count": "Count"}.get(pivot_config["aggfunc"], "Summary")
    value_label = ", ".join(pivot_config["values"])
    title_cell.value = f"Pivot Table  —  {agg_label} of {value_label}"
    title_cell.font = Font(bold=True, color=NUMBERS_TITLE_TEXT, size=13, name="Helvetica Neue")
    title_cell.fill = fill(NUMBERS_TITLE_BG)
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28

    # --- Header row (row 2) ---
    for col_idx, col_name in enumerate(pivot.columns, 1):
        cell = ws.cell(row=2, column=col_idx, value=str(col_name))
        cell.font = Font(bold=True, color=NUMBERS_HEADER_TEXT, size=11, name="Helvetica Neue")
        cell.fill = fill(NUMBERS_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border()
    ws.row_dimensions[2].height = 22

    # --- Data rows ---
    for row_idx, (_, row) in enumerate(pivot.iterrows(), 3):
        is_total = str(row.iloc[0]) == "Total"
        is_alt = (row_idx % 2 == 0)

        if is_total:
            row_fill = fill(NUMBERS_TOTAL_BG)
            row_font_color = NUMBERS_TOTAL_TEXT
            row_bold = True
        elif is_alt:
            row_fill = fill(NUMBERS_ROW_ALT)
            row_font_color = "000000"
            row_bold = False
        else:
            row_fill = fill(NUMBERS_ROW_WHITE)
            row_font_color = "000000"
            row_bold = False

        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)

            # Format numbers nicely
            if pd.isna(value):
                cell.value = "—"
            elif isinstance(value, float):
                if pivot_config["aggfunc"] == "mean":
                    cell.value = round(value, 2)
                    cell.number_format = '#,##0.00'
                else:
                    cell.value = int(value) if value == int(value) else round(value, 2)
                    cell.number_format = '#,##0'
            else:
                cell.value = str(value)

            cell.font = Font(bold=row_bold, color=row_font_color, size=11, name="Helvetica Neue")
            cell.fill = row_fill
            cell.border = border()

            # Right-align numbers, left-align text
            if col_idx > len(pivot_config["index"]):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws.row_dimensions[row_idx].height = 20

    # --- Auto column widths ---
    for col_idx in range(1, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, num_rows + 4):
            cell = ws.cell(row=row_idx, column=col_idx)
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    # --- Freeze header rows ---
    ws.freeze_panes = "A3"

    # --- Also add raw cleaned data on a second sheet ---
    ws2 = wb.create_sheet(title="Cleaned Data")
    for row_idx, row in enumerate(raw_rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=str(value).strip() if row_idx > 1 else value)
            if row_idx == 1:
                cell.font = Font(bold=True, color=NUMBERS_HEADER_TEXT)
                cell.fill = fill(NUMBERS_HEADER_BG)
                cell.alignment = Alignment(horizontal="center")

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Welcome to CSV Cleaner Bot!*\n\n"
        "Send me messy data or a `.csv` file and I'll return a *Numbers-style pivot table* Excel file.\n\n"
        "Just paste your data or upload a CSV and I'll handle the rest! ✨",
        parse_mode='Markdown'
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "*📖 Help*\n\n"
        "Send messy CSV data (pasted text or `.csv` file).\n"
        "I'll clean it and build a pivot table just like Apple Numbers.\n\n"
        "*/start* — Welcome\n"
        "*/help* — This message",
        parse_mode='Markdown'
    )


async def process_and_respond(update: Update, messy_data: str, output_filename: str):
    processing_msg = await update.message.reply_text("⏳ Building your pivot table...")

    try:
        result = await asyncio.get_event_loop().run_in_executor(
            None, clean_and_plan_pivot, messy_data
        )
        clean_csv = result["csv"]
        pivot_config = result["pivot"]

        excel_file = build_pivot_excel(clean_csv, pivot_config)

        await processing_msg.delete()

        agg = {"sum": "Sum", "mean": "Average", "count": "Count"}.get(pivot_config["aggfunc"], pivot_config["aggfunc"])
        caption = (
            f"✅ *Pivot Table ready!*\n\n"
            f"📊 *Rows:* {', '.join(pivot_config['index'])}\n"
            + (f"📋 *Columns:* {pivot_config['columns']}\n" if pivot_config.get('columns') else "")
            + f"🔢 *Values:* {', '.join(pivot_config['values'])} ({agg})"
        )

        await update.message.reply_document(
            document=excel_file,
            filename=output_filename,
            caption=caption,
            parse_mode='Markdown'
        )

    except Exception as e:
        logger.error(f"Error: {e}")
        await processing_msg.edit_text(f"❌ Error: {str(e)}\n\nPlease try again.")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    messy_data = update.message.text
    if len(messy_data) < 10:
        await update.message.reply_text("Please send more data — I need at least a few rows.")
        return
    await process_and_respond(update, messy_data, "pivot_table.xlsx")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.lower().endswith('.csv'):
        await update.message.reply_text("Please send a `.csv` file, or paste your data as text!")
        return
    try:
        file = await context.bot.get_file(document.file_id)
        file_bytes = await file.download_as_bytearray()
        messy_data = file_bytes.decode('utf-8')
        output_name = document.file_name.replace('.csv', '_pivot.xlsx')
        await process_and_respond(update, messy_data, output_name)
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        await update.message.reply_text(f"❌ Error: {str(e)}\n\nPlease try again.")


def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        print("❌ TELEGRAM_BOT_TOKEN not set!")
        return
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("❌ ANTHROPIC_API_KEY not set!")
        return

    print("🤖 Starting CSV Cleaner Bot (Pivot Table mode)...")
    print("Press Ctrl+C to stop")

    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
