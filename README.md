# 📊 IFS Weekly Report Generator

A Streamlit web application for generating weekly ServiceNow reports for NAFTA Marelli USA.

## 🚀 Features

- 📁 Upload Excel files (NA Daily work.xlsx and Template)
- 📅 Select date range for report generation
- 📊 Automatic ticket and user data processing
- 📥 Download generated reports
- 🎨 Formatted Excel output with styling and auto-filters

## 🌐 Live Demo

**App URL**: [Coming Soon]

## 🛠️ Local Development

### Prerequisites

- Python 3.8+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/tatsuogoya/IFS-Weekly-Report.git
cd IFS-Weekly-Report

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

## 📖 Usage

1. **Upload Source Data**: Upload your `NA Daily work.xlsx` file
2. **Upload Template**: Upload the `SNOW_report_Template.xlsx` file
3. **Select Date**: Choose a reference date for the report
4. **Generate**: Click "Generate Report" button
5. **Download**: Download the generated Excel report

## 📋 Requirements

- streamlit
- pandas
- openpyxl
- python-dateutil

## 🔒 Privacy

This app is **public** and can be accessed by anyone with the URL. No sensitive data is stored on the server.

## 📝 License

For internal use at NAFTA Marelli USA.

## 👤 Author

Tatsuo Goya

## 🆘 Support

For issues or questions, please contact the development team.
