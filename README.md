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
* **COM Interop & OpenPyXL Support:** Employs COM Automation for native MS Excel rendering in PowerShell, plus a platform-agnostic Python/Tkinter alternative.
* **Real-time Execution Logging:** Built-in UI console tracking line-by-line status changes and error handling.

## 🚀 Execution
Run via batch wrapper (`komus_upd.bat`) or directly via PowerShell:
```powershell
.\powershell\upd_editor.ps1