# CSV Cleaner 🧹

A simple tool that uses Claude AI to clean messy CSV data.

## What it does

- Standardizes names to Title Case
- Normalizes salaries (removes $, converts "50k" to "50000")
- Fixes mixed delimiters (commas, pipes, dashes → commas)
- Adds proper headers

## Installation

```bash
git clone https://github.com/Ogdenriver1/csv-cleaner.git
cd csv-cleaner
pip install -r requirements.txt
```

## Setup

1. Get an API key from [Anthropic Console](https://console.anthropic.com/)
2. Set your API key:
   ```bash
   export ANTHROPIC_API_KEY="your-key-here"
   ```

## Usage

```bash
python data_cleaner.py
```

Edit the `messy_data` variable in the script with your own data.

## Example

**Input:**
```
john smith, sales, 50k
jane doe - marketing - 62000
bob wilson | engineering | $75,000
```

**Output:**
```csv
name,department,salary
John Smith,Sales,50000
Jane Doe,Marketing,62000
Bob Wilson,Engineering,75000
```

## License

MIT
