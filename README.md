# PsSPM - Printer Status Monitoring via SNMP

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![SharpSnmpLib](https://img.shields.io/badge/SNMP-SharpSnmpLib-orange.svg)](https://github.com/lextudio/sharpsnmplib)

**PsSPM** is a PowerShell-based tool for monitoring printer status via SNMP. It checks toner levels, page counters, device information, and generates CSV/HTML reports. The tool supports both **Console** and **Full GUI (WPF)** modes, with optional email reporting.

> **SNMP Library used:** [SharpSnmpLib](https://github.com/lextudio/sharpsnmplib)

---

## ✨ Features

- ✅ Query any SNMP-enabled printer
- 📊 Generate **CSV** and **HTML** reports (toner levels, counters, device info)
- 📧 Send reports via email (Console mode only)
- 🖥️ Two interface modes:
  - **Console** – for automation and scripting
  - **FullGui** – WPF-based interactive GUI
- 🔧 Easily extendable – add new printer models and OIDs

---

## 📦 Supported Printers (Partial List)

| Brand    | Models |
|----------|--------|
| **Xerox** | C400/405, C600/605, B400, B600/B605/B610, 3330, 3325, 3655, 8045, 7025, 6510, 6600, 6700 |
| **HP**    | M402, M428, M252n, T790/730 (dummy) |
| **Lexmark** | MX611de, MX622ade |
| **Kyocera** | ECOSYS MA4500fx, ECOSYS PA3500cx |

> Need another model? See [Adding a New Printer](#-adding-a-new-printer).

---

## 🚀 Usage Examples

### Console Mode

```powershell
# Basic console mode with IP list file
powershell.exe -File .\PsSPM_0.3.5b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.5b\IP*.txt

# Console mode with email sending (custom user/pass)
powershell.exe -File .\PsSPM_0.3.5b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.5b\IP*.txt -MailSend $false -MailUser User -MailPass Pass
