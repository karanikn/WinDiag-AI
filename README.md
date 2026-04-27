# WinDiag-AI

[![GitHub release](https://img.shields.io/badge/version-1.0-blue?style=flat-square)](https://github.com/karanikn/WinDiag-AI)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?style=flat-square&logo=powershell)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey?style=flat-square&logo=windows)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-GPL--3.0-blue?style=flat-square)](LICENSE)
[![AI Assisted](https://img.shields.io/badge/built%20with-Claude%20AI-orange?style=flat-square&logo=anthropic)](https://claude.ai)

> **Windows system diagnostics with local AI analysis — all in a single PowerShell script.**  
> Scan → Analyze → Report. No cloud. No telemetry. Everything stays on your machine.

---

## 📸 Screenshots

| Main Window | Settings — AI Models |
|---|---|
| ![Main Window](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-Main.png) | ![Settings AI](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-SettingsAI.png) |

| Settings — General | HTML Report |
|---|---|
| ![Settings General](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-SettingsGeneral.png) | ![Report](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-Report.png) |

---

## ✨ Overview

**WinDiag-AI** is a professional Windows diagnostics tool built as a WPF GUI in PowerShell. It collects **28 categories of system data** — hardware, disk health, event logs, security, network, BSOD logs, driver status, and more — then sends the results to a **locally running AI model (Ollama)** for intelligent analysis.

The AI produces a structured report with:
- Executive summary of system health
- Critical issues requiring immediate attention
- Ready-to-paste repair commands (PowerShell / CMD)
- Optimization suggestions and preventive measures

Everything runs locally. No data leaves your machine. No subscriptions. No API keys.

Designed for **IT administrators, help desk engineers, and power users** who need fast, actionable diagnostics without manual log hunting.

---

## 🚀 Quick Start

### Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / Windows 11 |
| PowerShell | 5.1 or later |
| .NET Framework | 4.7.2+ (included with Windows 10/11) |
| Ollama | Installed and running — [ollama.com](https://ollama.com/download/windows) |
| Privileges | **Administrator** (for full diagnostics) |

### Run

```powershell
# Right-click WinDiag-AI.ps1 → Run with PowerShell
# The script auto-elevates to Administrator

# Or from an elevated PowerShell prompt (any folder):
powershell -ExecutionPolicy Bypass -File "C:\path\to\WinDiag-AI.ps1"
```

### Install Ollama (if not already installed)

```powershell
winget install Ollama.Ollama
```

Or download from [ollama.com/download](https://ollama.com/download/windows). Start Ollama, then launch WinDiag-AI — it will detect it automatically.

---

## 🖥️ Interface

The application features a two-panel layout:

- **Left panel** — 28 diagnostic check toggles, All/None selection
- **Right panel (top)** — Live log output with color-coded `[INFO]` / `[OK]` / `[WARN]` / `[ERROR]` / `[SCAN]` / `[DEBUG]` entries
- **Right panel (bottom)** — AI analysis output
- **Toolbar** — Scan System, AI Analysis, Save Report, Ollama, Cleanup, Verbose
- **Model bar** — AI model selector with info tooltip and download button
- **Theme selector** — Auto (follows Windows setting), Light, Dark

### Toolbar Buttons

| Button | Description |
|---|---|
| **Scan System** | Runs all selected diagnostic checks in background |
| **AI Analysis** | Sends scan results to local Ollama for analysis |
| **Save Report** | Exports full HTML report (all data + AI analysis) |
| **Ollama** | Install, start, or check Ollama status |
| **Cleanup** | Delete downloaded AI models to free disk space |
| **Verbose** | Show detailed debug output for every check |

---

## 📋 Diagnostic Checks

| # | Check | What it collects |
|---|---|---|
| 1 | **System Information** | OS, CPU, RAM, uptime, last boot, BIOS, manufacturer, serial number |
| 2 | **Hardware Details** | RAM module specs, GPU name and VRAM |
| 3 | **Battery Info** | Charge %, health %, chemistry, estimated runtime |
| 4 | **Disk Health + S.M.A.R.T** | All logical and physical disks, free space, health status |
| 5 | **↳ SMART Details** | Temperature, wear level, power-on hours, read/write errors |
| 6 | **Event Logs** | System & Application critical/error/warning events (configurable timeframe) |
| 7 | **Services (Stopped Auto)** | Auto-start services that are not currently running |
| 8 | **Network + DNS** | Adapters, IPs, gateways, DNS, established connections, ARP table, routing table, mapped drives |
| 9 | **Security Status** | Windows Defender, real-time protection, definition date, UAC, firewall profiles |
| 10 | **Windows Updates** | Pending updates with title, KB number, and severity |
| 11 | **Top Processes** | Top 15 processes by CPU and by RAM usage |
| 12 | **Startup Programs** | Registry run keys + startup folders (per-user and common) |
| 13 | **Scheduled Tasks** | Active non-Microsoft scheduled tasks |
| 14 | **Autoruns (Registry/Shell)** | Shell icon overlays, browser helper objects, non-system drivers |
| 15 | **System Integrity (SFC)** | SFC results from CBS.log + DISM log, reboot pending status |
| 16 | **Windows Search/Indexing** | WSearch service status, index size, item count |
| 17 | **Installed Software** | Up to 100 installed applications (registry-based) |
| 18 | **User Accounts** | Local accounts, enabled status, last logon, active sessions, user folders |
| 19 | **Hosts File** | All active entries (supports large adblock files with 8000+ entries) |
| 20 | **External IP** | Public IP address via ifconfig.me / ipify.org |
| 21 | **Installed Hotfixes** | Last 30 installed KBs with date and description |
| 22 | **Remote Tools** | TeamViewer, AnyDesk, RustDesk — installed and service status |
| 23 | **Chkdsk Logs** | Previous chkdsk results from Windows Event Log |
| 24 | **Battery Report (powercfg)** | Design capacity, full charge capacity, cycle count |
| 25 | **Driver Check** | Problem devices with error codes (Win32_PnPEntity + Win32_PnPSignedDriver) |
| 26 | **BSOD / Crash Logs** | BugCheck events (ID 1001/6008), minidumps, MEMORY.DMP, crash registry config |
| 27 | **RAM Test Logs** | mdsched/MemoryDiagnostics results, WHEA memory error count |
| 28 | **Performance Report** | Live CPU avg/peak, RAM, Disk I/O, Network I/O, page faults, context switches, top processes |

---

## 🤖 AI Models

Download and manage models directly from **File → Settings → AI Models**.

| Model | Size | Quality | Min RAM |
|---|---|---|---|
| llama3.2:1b | ~1.3 GB | Basic — very fast | 4 GB |
| llama3.2 | ~2.0 GB | Good balance | 6 GB |
| phi3:mini | ~2.3 GB | Very good | 6 GB |
| phi3.5 | ~2.2 GB | Very good+ | 6 GB |
| gemma2:2b | ~1.6 GB | Compact — very fast | 4 GB |
| qwen2.5:3b | ~1.9 GB | Good multilingual | 6 GB |
| qwen2.5:7b | ~4.7 GB | Strong multilingual | 8 GB |
| mistral | ~4.1 GB | Strong | 8 GB |
| llama3.1:8b | ~4.7 GB | Detailed | 10 GB |
| deepseek-r1:8b | ~4.9 GB | Reasoning / chain-of-thought | 10 GB |
| granite3.2:8b | ~4.9 GB | Enterprise IT | 10 GB |

Custom models from your local Ollama installation are also supported.  
Model downloads show **real-time progress**: percentage, MB/s speed, and a **Cancel** button.

---

## ⚙️ Settings

Access via **File → Settings**

### General Tab

| Setting | Default | Description |
|---|---|---|
| Ollama URL | `http://localhost:11434` | Supports remote Ollama (e.g. `http://192.168.1.100:11434`) |
| Report Output Folder | Script folder | Where HTML reports are saved |
| Log File Path | `WinDiag-AI.log` | Full activity log next to the script (empty to disable) |
| AI Temperature | 0.3 | 0.1 = precise/technical, 0.9 = creative |
| AI Max Tokens | 4096 | Response length: 1024 (short) to 8192 (very detailed) |
| Events Timeframe | 24 hours | How far back to look in Event Logs (12h / 24h / 48h / 7d) |
| Auto-Save Report | Off | Automatically generate HTML report after each scan |
| Custom Checks Folder | `custom_checks\` | Folder with user-defined `.ps1` diagnostic scripts |

### AI Models Tab

- Full table of recommended models: size, quality, speed, installed status, URL
- Download and delete models from within the app
- Open the Ollama models folder directly

---

## 📄 HTML Report

The exported HTML report includes:

- **Summary cards** — Critical / Error / Warning counts, service issues, pending updates
- **Live network test** — google.com, debian.org, karanik.gr (runs in browser on open)
- **All 28 diagnostic sections** as collapsible panels with sortable tables
- **BSOD section** — highlighted crash events and dump files
- **Performance snapshot** — CPU, RAM, Disk, Network metrics
- **AI Analysis** — full structured output with repair commands
- Standalone HTML file — no dependencies, works offline

---

## 🧩 Custom Checks

Place `.ps1` scripts in the `custom_checks` folder (create via **File → Custom Checks**). Each script runs automatically during scan; output is captured in both the log and the HTML report.

```powershell
# Example: custom_checks/my_check.ps1
Write-Output "Checking VPN connectivity..."
$ping = Test-Connection 10.0.0.1 -Count 1 -Quiet
Write-Output "VPN gateway: $(if($ping){'Reachable'}else{'Unreachable'})"
```

---

## 🗂️ File Structure

```
WinDiag-AI/
├── WinDiag-AI.ps1              # Main script — run this
├── WinDiag-AI.log              # Activity log (auto-created in script folder)
├── custom_checks/              # Your .ps1 diagnostic scripts
│   └── _example.ps1
└── README.md
```

All output files (reports, logs) are created **in the same folder as the script** by default. Paths are fully configurable via Settings.

---

## 🏗️ Architecture

| Component | Details |
|---|---|
| **Language** | PowerShell 5.1+ |
| **UI** | WPF (Windows Presentation Foundation) via `[System.Windows.Markup.XamlReader]` |
| **Threading** | `[PowerShell]::Create()` + dedicated Runspace; `ConcurrentQueue<string>` for thread-safe UI updates via `DispatcherTimer` |
| **AI integration** | HTTP POST to local Ollama `/api/chat`; temperature and max tokens configurable |
| **Model downloads** | Streaming `HttpWebRequest` to Ollama `/api/pull`; line-by-line JSON for real-time progress + cancel flag |
| **Scan data** | PowerShell hashtables → JSON; split into base + extended payload to avoid PS 5.1 `ConvertTo-Json` depth limits |
| **Report generation** | Self-contained HTML with embedded CSS/JS; sortable tables; dark-themed |

---

## 🔒 Privacy

All processing is **100% local**:

- Diagnostic data never leaves your machine
- AI analysis runs on your local Ollama instance
- No telemetry, no analytics, no cloud APIs
- Internet only used for: Ollama model downloads, external IP check (`ifconfig.me`), and browser connectivity test in HTML report

---

## 👤 Author

**Nikolaos Karanikolas**  
🌐 [karanik.gr](https://karanik.gr)

---

## 🤖 AI Assistance

This project was developed with the assistance of **[Claude](https://claude.ai)** (Anthropic AI). The architecture, WPF GUI, threading model, async patterns, scan pipeline, HTML report generator, and all PowerShell code were designed and iterated collaboratively between the developer and Claude over an extended development session.

---

## ⚠️ Disclaimer

This tool executes PowerShell scripts with Administrator privileges. Always review scripts before running them in production environments. The author takes no responsibility for data loss or system damage resulting from use of this tool.
