# 💊 Balance Sheet — PDF to Excel Converter

A Streamlit web app that automatically extracts pharmacy balance sheet data from a monthly PDF and converts it into a formatted Excel file ready for editing.

## ✨ Features
- 📤 Upload any monthly PDF balance sheet
- 🔄 Auto-extracts all table data (drug code, name, balance, prices)
- 📊 Preview extracted data before downloading
- 📥 Download a styled, RTL Arabic Excel file
- ➕ Totals row auto-calculated
- 🌐 Works in the browser — no installation needed

## 🚀 Deploy on Streamlit Cloud (Free)

1. Fork this repository on GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your forked repo
4. Set **Main file path** to `app.py`
5. Click **Deploy** — done! 🎉

## 🖥️ Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📁 Project Structure

```
pdf-to-excel/
├── app.py            # Main Streamlit application
├── requirements.txt  # Python dependencies
└── README.md
```

## 📋 Expected PDF Format

The PDF should contain a table with these columns (Arabic):

| كود الدواء | اسم الدواء | الرصيد الحالي | اخر سعر توريد | سعر البيع للمنتفع | اخر سعر متحرك |
|---|---|---|---|---|---|
| Drug Code | Drug Name | Current Balance | Last Supply Price | Selling Price | Moving Price |
