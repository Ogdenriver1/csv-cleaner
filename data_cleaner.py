import anthropic
import os
import pandas as pd
from telegram_bot import send_file

# Your messy data
messy_data = """
john smith, sales, 50k
jane doe - marketing - 62000
bob wilson | engineering | $75,000
"""

# Connect to Claude
client = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY")
)

# Ask Claude to clean it
message = client.messages.create(
    model="claude-3-haiku-20240307",
    max_tokens=1024,
    messages=[
        {
            "role": "user",
            "content": f"""Clean this messy data into a proper CSV format.
Rules:
- Standardize names to Title Case
- Make all salaries plain numbers (no $, no k)
- Use commas as separators
- Add a header row: name,department,salary

Data:
{messy_data}

Return ONLY the clean CSV, nothing else."""
        }
    ]
)

clean_csv = message.content[0].text.strip()

print("✨ Cleaned data:")
print(clean_csv)

# Save cleaned CSV
output_path = "cleaned_output.csv"
with open(output_path, "w") as f:
    f.write(clean_csv)

# --- Pivot table ---
def make_pivot(input_file, output_file):
    df = pd.read_csv(input_file)

    pivot = pd.pivot_table(
        df,
        index=df.columns[0],
        columns=df.columns[1],
        values=df.columns[2],
        aggfunc="sum"
    )

    pivot.to_excel(output_file)
    print(f"📊 Pivot table saved to {output_file}")

clean_file = output_path
pivot_file = output_path.replace(".csv", "_pivot.xlsx")

make_pivot(clean_file, pivot_file)

# --- Send pivot to Telegram ---
send_file(pivot_file)
