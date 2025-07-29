# Usage with uv

This project is set up to work with [uv](https://docs.astral.sh/uv/) for fast, predictable dependency management across Linux, macOS, and Windows.

**Requirements**: Python 3.9 or higher

## Installing uv

### macOS and Linux
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Windows
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### Alternative: Using pip
```bash
pip install uv
```

## Running the Application

### Direct execution (recommended)
```bash
# Run the main application
uv run excel-aggregator.py

# Run the Excel to CSV converter standalone  
uv run excel_to_csv.py
```
