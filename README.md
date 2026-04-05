# ⚡ Windows Power & Sleep Manager

![CI](https://github.com/AniketDeshmane/SleepProcessManager/actions/workflows/ci.yml/badge.svg)
![Release](https://github.com/AniketDeshmane/SleepProcessManager/actions/workflows/release.yml/badge.svg)

A **real-time monitoring dashboard** for diagnosing Windows sleep and wake issues. Built with Python and PyQt5.

![Dashboard Preview](https://img.shields.io/badge/Platform-Windows-blue?style=for-the-badge&logo=windows)
![Python](https://img.shields.io/badge/Python-3.10%2B-green?style=for-the-badge&logo=python)

---

## 🎯 Problem

Your Windows laptop won't sleep or wakes up randomly? This tool shows you **exactly what's blocking sleep** in real time and lets you fix it with one click.

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔄 **Real-Time Monitoring** | Polls `powercfg /requests`, `/lastwake`, `/waketimers` every 5 seconds |
| 🚦 **Traffic Light** | Green = system clear, Pulsing Red = blockers active |
| ⛔ **Process Killer** | Terminate blocking processes directly from the dashboard |
| 🛡️ **Sleep Override** | Apply `powercfg /requestsoverride` with one click |
| 🖥️ **Device Manager** | Quick shortcut to manage wake-capable devices |
| ⚙️ **Power Options** | Direct access to Windows Power settings |
| 📋 **Event Log** | Timestamped history of all scans and actions |
| 🔒 **Admin Check** | Automatically requests elevation on startup |

## 📥 Quick Install (No Python Needed)

1. Go to [**Releases**](https://github.com/AniketDeshmane/SleepProcessManager/releases/latest)
2. Download `SleepProcessManager-x.x.x.exe`
3. **Right-click → Run as Administrator**

## 🛠️ Run from Source

```powershell
# Clone the repo
git clone https://github.com/AniketDeshmane/SleepProcessManager.git
cd SleepProcessManager

# Install dependencies
pip install -r requirements.txt

# Run as Administrator
python sleep_manager.py
```

## 🏗️ Build EXE Locally

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed --name "SleepProcessManager" sleep_manager.py
# Output: dist/SleepProcessManager.exe
```

## 🔄 CI/CD Pipeline

| Workflow | Trigger | What it does |
|----------|---------|-------------|
| **CI** | Push to `master` / PRs | Syntax check → Build EXE → Upload artifact |
| **Release** | Push tag `v*` | Build versioned EXE → Generate release notes → Publish GitHub Release |

### Creating a Release

```powershell
git tag v1.0.0
git push origin v1.0.0
```

This automatically builds the `.exe` and creates a GitHub Release with download links.

## 📜 License

MIT License — use freely.
