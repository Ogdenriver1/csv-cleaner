# CSV Cleaner Bot 🧹✨

A Telegram bot that cleans messy CSV data using AI and sends back beautiful Excel files.

![Demo](demo.gif)

## Features

- 📤 **Send CSV files** - Upload any messy .csv file
- 📝 **Paste data** - Or just paste messy data directly
- 🤖 **AI-powered** - Uses Claude AI to understand and fix your data
- 📊 **Excel output** - Get back a formatted .xlsx file
- ✨ **Auto-fixes**:
  - Inconsistent names → Title Case
  - Mixed delimiters (commas, pipes, tabs) → Clean CSV
  - Messy numbers ($50k) → Plain numbers (50000)
  - Missing headers → Adds appropriate headers
  - Extra whitespace → Trimmed

## Quick Start

### 1. Create Your Telegram Bot

1. Open Telegram and search for **@BotFather**
2. Send `/newbot`
3. Choose a name (e.g., "My CSV Cleaner")
4. Choose a username ending in `bot` (e.g., `my_csv_cleaner_bot`)
5. Save the token BotFather gives you

### 2. Get an Anthropic API Key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up / Log in
3. Go to API Keys → Create Key
4. Save your key

### 3. Install & Run

```bash
# Clone the repo
git clone https://github.com/Ogdenriver1/csv-cleaner.git
cd csv-cleaner

# Install dependencies
pip install -r requirements.txt

# Set your API keys
export TELEGRAM_BOT_TOKEN="your-telegram-token"
export ANTHROPIC_API_KEY="your-anthropic-key"

# Run the bot!
python telegram_bot.py
```

### 4. Use Your Bot

1. Open Telegram
2. Search for your bot's username
3. Send `/start`
4. Send messy data or upload a CSV file
5. Get back a clean Excel file! 🎉

## Example

**You send:**
```
john smith, sales, 50k
jane doe - marketing - 62000
bob wilson | engineering | $75,000
```

**Bot returns:** A formatted Excel file with:

| Name | Department | Salary |
|------|------------|--------|
| John Smith | Sales | 50000 |
| Jane Doe | Marketing | 62000 |
| Bob Wilson | Engineering | 75000 |

## Running 24/7 (Optional)

To keep your bot running all the time:

### On your Mac:
```bash
# Run in background
nohup python telegram_bot.py > bot.log 2>&1 &
```

### On a server (recommended):
Deploy to Railway, Render, or any VPS. Set the environment variables and run `python telegram_bot.py`.

## Files

- `telegram_bot.py` - The Telegram bot
- `data_cleaner.py` - Standalone CLI version
- `requirements.txt` - Python dependencies

## Cost

- **Telegram**: Free
- **Anthropic API**: ~$0.001 per request (very cheap)

## License

MIT - Do whatever you want with it!

---

Made with 🦞 by [Ogdenriver1](https://github.com/Ogdenriver1)
