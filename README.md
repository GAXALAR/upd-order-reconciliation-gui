# Enterprise Universal Transfer Document (UPD) & Order Reconciliation Tool

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)
![Python](https://img.shields.io/badge/Python-3.9%2B-green.svg)
![GUI](https://img.shields.io/badge/UI-WinForms%2FTkinter-orange.svg)

Desktop automation application engineered to resolve cross-system SKU mismatches between supplier procurement orders and formal Universal Transfer Documents (UPD / УПД).

## 🏢 Business Problem
In B2B supply chains, vendors (e.g., Komus) frequently issue transfer documents using internal system identifiers that differ from the client's master catalog numbers. Manual reconciliation leads to high error rates and delayed ERP entries.

## 🛠 Key Features
* **Automated SKU Mapping:** Parses order files, builds internal code lookup maps, and updates line item SKUs across complex spreadsheet layouts.
* **Order Number Injection:** Scans document structure for transfer base fields ("Основание передачи") and automatically injects order metadata.
* **Python/Tkinter Primary Application:** Uses OpenPyXL for workbook processing without requiring Microsoft Excel; PowerShell/WinForms remains an alternative that uses Excel COM Automation.
* **Real-time Execution Logging:** Built-in UI console tracking line-by-line status changes and error handling.

## Running

The Windows batch launcher starts the primary Python version (Python 3.9 or newer):

```powershell
.\komus_upd.bat
```

Install Python dependencies before the first launch:

```powershell
python -m pip install -r python\requirements.txt
```

To start the PowerShell alternative directly (requires Microsoft Excel):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\powershell\upd_editor.ps1
```

## Testing

The synthetic reconciliation tests use fictional workbooks. From the `python` directory:

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

## Workbook layout

Both versions read the shared `config.json` in the repository root. It defines
1-based sheet indexes, Excel column letters, first data rows, labels, and the
row offset/column where the order number is written. If the selected order
layout contains no mappings or the configured UPD layout produces no matches,
processing stops with a configuration-oriented message.
