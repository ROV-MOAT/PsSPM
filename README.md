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
| **Xerox** | C400/405, C600/605, B400/405, B600/605/610, 3325, 3330, 3655, 4622, 6510, 6600, 6700, 7025, 8045 |
| **HP**    | M402, M428, M252n / T790/730 (dummy) |
| **Lexmark** | MX611de, MX622ade |
| **Kyocera** | ECOSYS MA4500fx, ECOSYS PA3500cx |

> Need another model? See [Adding a New Printer](#%EF%B8%8F-configuration).

---

## 🚀 Usage Examples

### Console Mode

```powershell
# Basic console mode with IP list file
powershell.exe -File .\PsSPM_0.3.6b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.6b\IP*.txt

# Console mode with email sending (custom user/pass)
powershell.exe -File .\PsSPM_0.3.6b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.6b\IP*.txt -MailSend $false -MailUser User -MailPass Pass
```

### Full GUI Mode

```powershell
powershell.exe -File .\PsSPM_0.3.6b.ps1 -InterfaceMode FullGui
```

Note: Email settings are only available in Console mode (no mail configuration in GUI).

---

## 📧 Email Reporting (Console Mode)

Email sending is convenient for automation. Example:

```powershell
-MailSend $true -MailUser "monitoring@example.com" -MailPass "your_password"
```

⚠️ Security Warning: Do not store your account password directly in the script. Use environment variables or secure strings where possible.

---

## ⚙️ Configuration
Adding a New Printer Model:

Open Lib\PsSPM_oid.psd1 - Add your printer model and corresponding OIDs

Edit the $modelPatterns table in the main script to include pattern matching for your model

---

## 🛠️ Requirements
Windows PowerShell 5.1 or PowerShell 7+

SNMP enabled on target printers

.NET Framework 4.7.2+ (for WPF GUI)

---

## 📄 License
This project is licensed under the MIT License – see the LICENSE file for details.
