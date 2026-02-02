# 📊 Monthly Report Generator - Setup Guide

This application automatically generates monthly reports from Excel files.

---

## 🖥️ System Requirements

- **Windows 10/11** or **macOS** or **Linux**
- **Python 3.10 or higher** (Follow Step 1 if not installed)

---

## 📥 Step 1: Install Python (Skip if already installed)

### Windows:
1. Visit [Python Official Website](https://www.python.org/downloads/)
2. Click "Download Python 3.x.x"
3. Run the downloaded installer
4. **IMPORTANT**: Check "Add Python to PATH"
5. Click "Install Now"

### macOS:
1. Open Terminal
2. Run the following command:
```bash
brew install python3
```
(If you don't have Homebrew, install it from [here](https://brew.sh/))

### Verify Installation:
Open Command Prompt (Windows) or Terminal (macOS/Linux) and run:
```bash
python --version
```
You should see `Python 3.10.x` or higher

---

## 📂 Step 2: Extract Files

Extract the ZIP file you received to any location.

Example: `C:\MonthlyReport\` or `~/MonthlyReport/`

---

## 🔧 Step 3: Install Required Libraries

### Windows:
1. Open **Command Prompt** (Windows Key + R → type `cmd`)
2. Navigate to the extracted folder:
```cmd
cd C:\MonthlyReport\WeeklyExcel-Report
```
3. Install required libraries:
```cmd
pip install -r requirements.txt
```

### macOS/Linux:
1. Open **Terminal**
2. Navigate to the extracted folder:
```bash
cd ~/MonthlyReport/WeeklyExcel-Report
```
3. Install required libraries:
```bash
pip3 install -r requirements.txt
```

**Note**: Installation may take a few minutes.

---

## 🚀 Step 4: Start the Application

### Windows:
```cmd
python run_api.py
```

### macOS/Linux:
```bash
python3 run_api.py
```

If successful, you should see:
```
INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
INFO:     Started server process
INFO:     Waiting for application startup.
INFO:     Application startup complete.
```

**Do NOT close this window!** (The app is running)

---

## 🌐 Step 5: Access in Browser

1. Open your browser (Chrome, Edge, Firefox, etc.)
2. Type in the address bar:
```
http://localhost:8000
```
3. The Monthly Report Generator interface will appear

---

## 📊 Step 6: How to Generate Reports

### Monthly Report Generation:
1. **Report Type**: Select "Monthly"
2. **Year**: Enter the report year (e.g., 2025)
3. **Month**: Enter the report month (e.g., 12)
4. **Choose Excel File**: Select your Excel file
5. **Generate Web View**: Click to generate
6. **Download PDF**: Download the generated PDF

### SLA Breach Data (Optional):
- Check the checkbox and enter data
- Will be automatically included in the report

### Development Efforts Data (Optional):
- Check the checkbox and enter data
- Will be automatically included in the report

---

## ⚠️ Troubleshooting

### Error: `ModuleNotFoundError`
→ Re-run the library installation from Step 3

### Error: `Address already in use`
→ Port 8000 is already in use. Run:
```cmd
# Kill existing process
taskkill /F /IM python.exe
# Start again
python run_api.py
```

### Cannot connect in browser
→ Firewall might be blocking it
→ Try `http://127.0.0.1:8000` instead

---

## 🛑 How to Stop the Application

Press **Ctrl+C** in the Command Prompt/Terminal window

---

## 📧 Support

If you encounter issues, please provide:
- Screenshot of error message
- Python version (`python --version` output)
- Operating System (Windows 10, macOS, etc.)

---

## 📁 File Structure

```
WeeklyExcel-Report/
├── app/                    # Application code
│   ├── api/               # API endpoints
│   ├── services/          # Business logic
│   └── infra/             # Infrastructure (chart generation, etc.)
├── static/                 # Static files (HTML, generated files)
├── data/                   # Data files (SLA, Dev Efforts)
├── requirements.txt        # Required libraries list
├── run_api.py             # Startup script
└── SETUP_GUIDE.md         # This file
```

---

**Happy Reporting! 📊✨**
