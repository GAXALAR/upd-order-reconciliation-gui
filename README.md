# Enterprise Universal Transfer Document (UPD) & Order Reconciliation Tool

![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.1%2B-green.svg)
![GUI](https://img.shields.io/badge/UI-Tkinter-orange.svg)

Desktop automation application engineered to resolve cross-system SKU mismatches between supplier procurement orders and formal Universal Transfer Documents (UPD / УПД).

## 🏢 Business Problem
In B2B supply chains, vendors (e.g., Komus) frequently issue transfer documents using internal system identifiers that differ from the client's master catalog numbers. Manual reconciliation leads to high error rates and delayed ERP entries.

## 🛠 Key Features
* **Automated SKU Mapping:** Parses order files, builds internal code lookup maps, and updates line item SKUs across complex spreadsheet layouts.
* **Order Number Injection:** Scans document structure for transfer base fields ("Основание передачи") and automatically injects order metadata.
* **Python Desktop Application:** Uses Tkinter for the interface and OpenPyXL to read and update workbooks without requiring Microsoft Excel.
* **Real-time Execution Logging:** Built-in UI console tracking line-by-line status changes and error handling.

## Running

Requires Python 3.9 or newer. Install dependencies:

```console
python -m pip install -r python\requirements.txt
```

Start the application with the Windows launcher:

```bat
.\komus_upd.bat
```

Or run it directly:

```console
python .\python\upd_editor.py
```

## Testing

The synthetic reconciliation tests use fictional workbooks. Run from the `python` directory:

```console
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

## Workbook layout

The app reads the shared `config.json` in the repository root. It defines
1-based sheet indexes, Excel column letters, first data rows, labels, and the
row offset/column where the order number is written. If the selected order
layout contains no mappings or the configured UPD layout produces no matches,
processing stops with a configuration-oriented message.
