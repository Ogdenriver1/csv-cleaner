import anthropic
import os

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

print("✨ Cleaned data:")
print(message.content[0].text)
