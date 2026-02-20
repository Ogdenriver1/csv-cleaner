"""
CSV Cleaner Telegram Bot
Send messy CSV data or a .csv file, get back a clean Excel file!

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
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# API clients
anthropic_client = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY")
)

def clean_data_with_claude(messy_data: str) -> str:
    """Send messy data to Claude and get clean CSV back."""
    message = anthropic_client.messages.create(
        model="claude-3-haiku-20240307",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": f"""Clean this messy data into a proper CSV format.

Rules:
- Standardize names to Title Case
- Standardize all values appropriately
- Make all numbers plain (no $, no k abbreviations)
- Use commas as separators
- Add an appropriate header row if missing
- Fix any inconsistent delimiters (|, -, tabs, etc.)
- Remove empty rows
- Trim whitespace

Data:
{messy_data}

Return ONLY the clean CSV, nothing else. No explanations, no markdown code blocks."""
            }
        ]
    )
    return message.content[0].text.strip()


def csv_to_excel(csv_data: str) -> io.BytesIO:
    """Convert CSV string to a formatted Excel file."""
    # Parse CSV
    lines = csv_data.strip().split('\n')
    reader = csv.reader(lines)
    rows = list(reader)
    
    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cleaned Data"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Add data
    for row_idx, row in enumerate(rows, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value.strip())
            cell.border = border
            cell.alignment = Alignment(horizontal='left')
            
            # Style header row
            if row_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 50)
    
    # Save to bytes
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /start command."""
    welcome_message = """👋 **Welcome to CSV Cleaner Bot!**

I clean messy data and turn it into beautiful Excel files.

**How to use:**
1️⃣ Send me a `.csv` file, OR
2️⃣ Paste your messy data directly

**I can fix:**
• Inconsistent formats (john smith → John Smith)
• Mixed delimiters (commas, pipes, tabs)
• Messy numbers ($50k → 50000)
• Missing headers
• And more!

**Try it!** Paste this:
```
john smith, sales, 50k
jane doe - marketing - 62000
bob wilson | engineering | $75,000
```

I'll send you back a clean Excel file! ✨"""
    
    await update.message.reply_text(welcome_message, parse_mode='Markdown')


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /help command."""
    help_text = """**📖 Help**

**Commands:**
/start - Welcome message
/help - This help message

**Supported formats:**
• CSV files (.csv)
• Pasted text with any delimiter

**Tips:**
• The more data you send, the better I understand the pattern
• I work best with tabular data (rows and columns)
• Send data in any messy format - I'll figure it out!

**Questions?** Contact the developer."""
    
    await update.message.reply_text(help_text, parse_mode='Markdown')


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle text messages with messy data."""
    messy_data = update.message.text
    
    if len(messy_data) < 10:
        await update.message.reply_text("Please send more data! I need at least a few rows to clean.")
        return
    
    # Send "processing" message
    processing_msg = await update.message.reply_text("🔄 Cleaning your data...")
    
    try:
        # Clean with Claude
        clean_csv = clean_data_with_claude(messy_data)
        
        # Convert to Excel
        excel_file = csv_to_excel(clean_csv)
        
        # Send the file
        await update.message.reply_document(
            document=excel_file,
            filename="cleaned_data.xlsx",
            caption="✨ Here's your cleaned data!"
        )
        
        # Also send preview
        preview_lines = clean_csv.split('\n')[:5]
        preview = '\n'.join(preview_lines)
        if len(clean_csv.split('\n')) > 5:
            preview += f"\n... and {len(clean_csv.split(chr(10))) - 5} more rows"
        
        await update.message.reply_text(f"**Preview:**\n```\n{preview}\n```", parse_mode='Markdown')
        
        # Delete processing message
        await processing_msg.delete()
        
    except Exception as e:
        logger.error(f"Error processing message: {e}")
        await processing_msg.edit_text(f"❌ Error: {str(e)}\n\nPlease try again with different data.")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle uploaded CSV files."""
    document = update.message.document
    
    # Check if it's a CSV
    if not document.file_name.lower().endswith('.csv'):
        await update.message.reply_text("Please send a .csv file, or paste your data as text!")
        return
    
    # Send "processing" message
    processing_msg = await update.message.reply_text("🔄 Processing your CSV file...")
    
    try:
        # Download file
        file = await context.bot.get_file(document.file_id)
        file_bytes = await file.download_as_bytearray()
        messy_data = file_bytes.decode('utf-8')
        
        # Clean with Claude
        clean_csv = clean_data_with_claude(messy_data)
        
        # Convert to Excel
        excel_file = csv_to_excel(clean_csv)
        
        # Generate output filename
        output_name = document.file_name.replace('.csv', '_cleaned.xlsx')
        
        # Send the file
        await update.message.reply_document(
            document=excel_file,
            filename=output_name,
            caption="✨ Here's your cleaned data!"
        )
        
        # Send preview
        preview_lines = clean_csv.split('\n')[:5]
        preview = '\n'.join(preview_lines)
        if len(clean_csv.split('\n')) > 5:
            preview += f"\n... and {len(clean_csv.split(chr(10))) - 5} more rows"
        
        await update.message.reply_text(f"**Preview:**\n```\n{preview}\n```", parse_mode='Markdown')
        
        # Delete processing message
        await processing_msg.delete()
        
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        await processing_msg.edit_text(f"❌ Error: {str(e)}\n\nPlease try again.")


def main():
    """Start the bot."""
    # Get token
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        print("❌ Error: TELEGRAM_BOT_TOKEN not set!")
        print("Run: export TELEGRAM_BOT_TOKEN='your-token-here'")
        return
    
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("❌ Error: ANTHROPIC_API_KEY not set!")
        print("Run: export ANTHROPIC_API_KEY='your-key-here'")
        return
    
    print("🤖 Starting CSV Cleaner Bot...")
    print("Press Ctrl+C to stop")
    
    # Create application
    app = Application.builder().token(token).build()
    
    # Add handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    # Start polling
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
