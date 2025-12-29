# Insurance Card Mailer

A desktop application for generating **medical insurance e-cards (PDF)** from Excel data and **automatically emailing them to members** with the card attached.

Built with **Python, Tkinter (ttkbootstrap), ReportLab, Pandas, and SMTP**.

---

## ✨ Features

- 📄 Generate medical insurance cards as **PDF**
- 📊 Read member data from **Excel (.xlsx)**
- 🆔 Auto-generate **unique card IDs**
- 🖼 Supports **logos & backside image**
- 📧 Send personalized **HTML emails with PDF attachment**
- 👥 CC support (default + per-member)
- ✅ Selective email sending using `SendEmail` column
- 🖥 User-friendly **desktop GUI**
- 🔐 Custom SMTP email configuration

---

## 🧩 Technologies Used

- Python 3.x
- Tkinter + ttkbootstrap
- ReportLab
- Pandas
- SMTP (Gmail supported)
- Excel (.xlsx)

---

## 📁 Project Structure


insurance-card-mailer/
│
├── main.py # Main application
├── README.md
├── requirements.txt
├── assets/
│ ├── logo_left.png
│ ├── logo_middle.png
│ ├── logo_right.png
│ └── card_back.png



---

## 📄 Excel Template Columns

Required columns:

| Column Name | Description |
|------------|------------|
| CardNumber | Member card number |
| InsuredName | Member name |
| Policyholder | Company name |
| PolicyNumber | Policy number |
| StartDate | Policy start date |
| ExpiryDate | Policy expiry date |
| DOB | Date of birth |
| Gender | Gender |
| Email | Member email |
| SendEmail | Yes / No |
| UniqueID | Auto-generated if empty |

Use **Download Template** button in the app to get a ready-to-use Excel file.

---

## 🚀 How to Run

### 1️⃣ Install dependencies
```bash
pip install -r requirements.txt


