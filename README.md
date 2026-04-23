# WinDiag-AI

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D4?logo=windows)
![Ollama](https://img.shields.io/badge/AI-Ollama-black?logo=ollama)
![License](https://img.shields.io/badge/License-GPL--3.0-green)
![Platform](https://img.shields.io/badge/Platform-x64-lightgrey)

**Windows Diagnostics with AI Analysis** — A WPF GUI tool that collects comprehensive system diagnostics and sends them to a local AI (Ollama) for intelligent analysis and actionable repair recommendations.

> 🔧 by [Nikolaos Karanikolas](https://karanik.gr) | [karanik.gr](https://karanik.gr)

---

## 📸 Screenshots

> Dark theme with scan results, AI analysis, and HTML report output.

---

## ✨ Features

- **28 diagnostic checks** covering hardware, software, security, network, and system health
- **AI-powered analysis** via local Ollama — no cloud, no data leaving your machine
- **Dark / Light / Auto theme** with full dialog support
- **HTML report** export with sortable tables, network connectivity test, and AI summary
- **Streaming model download** with real-time progress, speed, and cancel support
- **Custom checks** — run your own `.ps1` scripts during scan
- **BSOD & Crash log detection** — events, minidumps, MEMORY.DMP
- **Performance snapshot** — CPU, RAM, Disk I/O, Network I/O, top processes
- **S.M.A.R.T details** — disk health, temperature, wear, power-on hours
- **Driver check** — problem devices with error codes
- **Console output** — all actions mirrored to PowerShell window with colors
- **File logging** — all activity saved to `WinDiag-AI.log` in script folder
- **Background worker** — non-blocking UI via `[PowerShell]::Create()` + Runspace

---

## 🖥️ Requirements

| Component | Requirement |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1+ (comes with Windows) |
| .NET Framework | 4.7.2+ |
| Ollama | Installed & running ([ollama.com](https://ollama.com)) |
| Privileges | **Administrator** (for full diagnostics) |

---

## 🚀 Quick Start

### 1. Download

```powershell
# Option A: Download from GitHub
# Download WinDiag-AI.ps1 and run it

# Option B: Clone
git clone https://github.com/karanikn/WinDiag-AI.git
cd WinDiag-AI
```

### 2. Install Ollama (if not already installed)

```powershell
winget install Ollama.Ollama
```

Or download from [ollama.com/download](https://ollama.com/download/windows)

### 3. Run

```powershell
# Right-click → Run with PowerShell (as Administrator)
# OR from an elevated PowerShell prompt:
powershell -ExecutionPolicy Bypass -File WinDiag-AI.ps1
```

> The script auto-elevates to Administrator if not already running as one.

---

## 📋 Diagnostic Checks

| # | Check | Details |
|---|---|---|
| 1 | **System Information** | OS, CPU, RAM, uptime, BIOS, manufacturer, serial |
| 2 | **Hardware Details** | RAM modules, GPU info |
| 3 | **Battery Info** | Charge %, health %, chemistry, runtime |
| 4 | **Disk Health + S.M.A.R.T** | Logical & physical disks, free space, status |
| 5 | **↳ SMART Details** | Temperature, wear, power-on hours, read/write errors |
| 6 | **Event Logs** | System & Application errors/warnings (configurable timeframe) |
| 7 | **Services (Stopped Auto)** | Auto-start services that are not running |
| 8 | **Network + DNS** | Adapters, IPs, gateways, DNS, established connections, ARP, routes |
| 9 | **Security Status** | Defender, real-time protection, definitions, UAC, firewall |
| 10 | **Windows Updates** | Pending updates with severity |
| 11 | **Top Processes** | Top 15 by CPU and RAM |
| 12 | **Startup Programs** | Registry + startup folder entries |
| 13 | **Scheduled Tasks** | Non-Microsoft active tasks |
| 14 | **Autoruns (Registry/Shell)** | Shell extensions, BHOs, non-system drivers |
| 15 | **System Integrity (SFC)** | CBS.log SFC results, DISM health, reboot pending |
| 16 | **Windows Search/Indexing** | Service status, index size, item count |
| 17 | **Installed Software** | Up to 100 installed applications |
| 18 | **User Accounts** | Local accounts, enabled status, last logon, sessions |
| 19 | **Hosts File** | Active entries (up to 8000+) |
| 20 | **External IP** | Public IP via ifconfig.me / ipify |
| 21 | **Installed Hotfixes** | Last 30 KBs |
| 22 | **Remote Tools** | TeamViewer, AnyDesk, RustDesk detection |
| 23 | **Chkdsk Logs** | Previous chkdsk results from event log |
| 24 | **Battery Report (powercfg)** | Design capacity, full charge capacity, cycle count |
| 25 | **Driver Check** | Problem devices with error codes (Win32_PnPEntity) |
| 26 | **BSOD / Crash Logs** | BugCheck events, unexpected shutdowns, minidumps, MEMORY.DMP |
| 27 | **RAM Test Logs** | mdsched/MemoryDiagnostics results, WHEA errors |
| 28 | **Performance Report** | Live CPU/RAM/Disk/Network counters, page faults, top processes |

---

## 🤖 Supported AI Models

Download and manage models directly from the app's Settings → AI Models tab.

| Model | Size | Quality | Min RAM |
|---|---|---|---|
| llama3.2:1b | ~1.3 GB | Basic — fast | 4 GB |
| llama3.2 | ~2.0 GB | Good balance | 6 GB |
| phi3:mini | ~2.3 GB | Very good | 6 GB |
| phi3.5 | ~2.2 GB | Very good+ | 6 GB |
| gemma2:2b | ~1.6 GB | Compact | 4 GB |
| qwen2.5:3b | ~1.9 GB | Multilingual | 6 GB |
| qwen2.5:7b | ~4.7 GB | Strong multilingual | 8 GB |
| mistral | ~4.1 GB | Strong | 8 GB |
| llama3.1:8b | ~4.7 GB | Detailed | 10 GB |
| deepseek-r1:8b | ~4.9 GB | Reasoning | 10 GB |
| granite3.2:8b | ~4.9 GB | Enterprise | 10 GB |

> You can also use any custom model available in your Ollama installation.

---

## ⚙️ Settings

Access via **File → Settings**

### General Tab

| Setting | Default | Description |
|---|---|---|
| Ollama URL | `http://localhost:11434` | Remote Ollama support (e.g. `http://192.168.1.100:11434`) |
| Report Output Folder | Script folder | Where HTML reports are saved |
| Log File Path | `WinDiag-AI.log` | Full activity log (empty = disable) |
| AI Temperature | 0.3 | 0.1 = precise, 0.9 = creative |
| AI Max Tokens | 4096 | Response length (1024–8192) |
| Events Timeframe | 24 hours | How far back to look in Event Logs |
| Auto-Save Report | Off | Auto-generate HTML report after scan |
| Custom Checks Folder | `custom_checks\` | Folder with user `.ps1` diagnostic scripts |

### AI Models Tab

- View all recommended models with size, quality, speed, and status
- Download / delete models from within the app
- Real-time download progress with MB/s, percentage, and cancel button

---

## 📄 HTML Report

The exported report includes:

- **Executive summary cards** — Critical/Error/Warning counts, service issues, pending updates
- **Live network connectivity test** — google.com, debian.org, karanik.gr (runs in browser)
- **All diagnostic sections** as collapsible `<details>` panels
- **Sortable tables** — click any column header to sort
- **AI Analysis** section with structured repair commands
- Dark-themed design matching the app

---

## 🧩 Custom Checks

Place `.ps1` scripts in the `custom_checks` folder (create via **File → Custom Checks**). Each script runs during scan and its output is captured in the report.

```powershell
# Example: custom_checks/check_disk_temp.ps1
Write-Output "Custom disk temperature check"
Write-Output "Status: OK"
```

---

## 📁 File Structure

```
WinDiag-AI/
├── WinDiag-AI.ps1          # Main script
├── WinDiag-AI.log          # Activity log (auto-created)
├── custom_checks/          # Your custom .ps1 diagnostic scripts
│   └── _example.ps1
└── README.md
```

---

## 🏗️ Architecture

```
WinDiag-AI.ps1
├── WPF GUI (STA thread)
│   ├── Toolbar + Model bar + Theme selector
│   ├── 28-check panel (left)
│   ├── Log output RichTextBox (right-top)
│   └── AI Output RichTextBox (right-bottom)
├── Background Workers ([PowerShell]::Create + Runspace)
│   ├── Scan Worker → JSON via ConcurrentQueue
│   ├── AI Worker → streaming chat → queue
│   └── Download Worker → streaming HTTP → queue + cancel flag
├── DispatcherTimer (200ms) → UI updates from queue
└── Report Generator → standalone HTML file
```

**Data flow:**
1. Scan Worker collects data → serializes to JSON (split: base + extended to avoid PS 5.1 limits)
2. `EXT:` message carries extended data (Hosts, Drivers, BSOD, RAM, Perfmon, etc.)
3. `DONE:SCAN:` triggers UI summary + optional auto-save
4. AI Worker receives formatted diagnostic text → streams response
5. Report is generated from `$Global:DiagData` + `$Global:ExtJson`

---

## 🔒 Privacy

All processing is **100% local**:
- Diagnostic data never leaves your machine
- AI analysis runs on your local Ollama instance
- No telemetry, no cloud APIs, no internet required (except for model downloads and external IP check)

---

## 📝 License

[GNU General Public License v3.0](LICENSE)

---

## 👤 Author

**Nikolaos Karanikolas**
- Website: [karanik.gr](https://karanik.gr)
- GitHub: [@karanikn](https://github.com/karanikn)

---

## 🔗 Related Projects

- [karanik_WinMaintenance_Online](https://github.com/karanikn/karanik_WinMaintenance_Online) — Windows maintenance automation
- [KeePassNetworkChecker](https://github.com/karanikn/KeePassNetworkChecker) — KeePass plugin for network checks
