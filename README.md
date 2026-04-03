# Roblox Group Member Exporter

A python script which fetches members of a specific role from a Roblox group via the roblox api and exports them to an Excel Spreadsheet with pagination support

## Features

- Fetches all members of a specific role from any Roblox group
- Handles pagination automatically for large groups
- Exports data to a styled `.xlsx` spreadsheet including:
  - User ID, username, display name, and profile URL
  - Alternating row colors and freeze panes for readability
  - Auto-filter and metadata (export date, total member count)

## Requirements

- Python 3.10+
- `requests`
- `openpyxl`

Install dependencies:

```bash
pip install requests openpyxl
```

## Usage

1. Open `exporter.py` and set the config variables at the top:

```python
GROUP_ID  = 123456789   # Your Roblox group ID
ROLE_NAME = "ROLE NAME" # The exact role name to export
OUTPUT_FILE = "members.xlsx" # Output filename
```

2. Run the script:

```bash
python exporter.py
```

3. The spreadsheet will be saved to the path specified in `OUTPUT_FILE`.

## Output

| # | User ID | Username | Display Name | Profile URL |
|---|---------|----------|--------------|-------------|
| 1 | 123456 | Builderman | Builderman | https://www.roblox.com/users/123456/profile |

## License

MIT
