# 📘 Excel Column Consistency Checker

This project provides a **Python script** to check **column consistency** across multiple Excel files stored in **Google Drive**.  
It automatically validates column values against predefined **patterns** and generates a detailed **Excel report**.

---

## 🚀 Features

- Process **multiple Excel files** from Google Drive.
- Read **all sheets** in each file.
- Automatically **detect column patterns** (Arabic, English, Numbers, Dates).
- Special handling for:
  - **Address** columns
  - **Request Number** columns
  - **Note/Description** columns
- **Length consistency** check for numeric fields.
- Ignore **null/empty** values.
- Generate a **single Excel report** summarizing:
  - File name
  - Sheet name
  - Column name
  - Detected pattern
  - Numeric average length & allowed range
  - Consistency percentage
  - Count of consistent & inconsistent values
  - List of inconsistent values

---

## 📥 Input

- Multiple **Excel files** stored in Google Drive.
- Each file may contain **multiple sheets**.

---

## 📏 Pattern Detection Rules

The **first non-empty record** in each column defines its pattern  
(❗except for `Address`, `Request Number`, and `Note`, which have explicit rules).

| Pattern            | Rule                                                                 | Notes                                                                 |
| ------------------ | -------------------------------------------------------------------- | --------------------------------------------------------------------- |
| **Arabic words**   | `[\u0600-\u06FF]` + symbols `() _ - \ / .`                          | First non-empty record defines pattern                                |
| **English words**  | `A–Z`, `a–z`, spaces, `_ - .`                                       | First non-empty record defines pattern                                |
| **Numbers**        | `[0-9]+`                                                            | First non-empty record defines pattern; length check (see below)      |
| **Dates**          | Recognizable formats (`YYYY-MM-DD`, `DD/MM/YYYY`, etc.)             | First non-empty record defines pattern                                |
| **Address**        | Column name contains `"address"` → Arabic/English/digits + symbols. Must contain at least one letter. Numbers-only values are inconsistent. | Independent of first record |
| **Request Number** | Column name = `"REQ_NO"` or `"REQUEST_NO"` → Must be **number > 0** | Independent of first record                                           |
| **Note**           | Column name contains `"note"`, `"desc"`, `"desca"` → Any value OK   | Independent of first record                                           |

---

## 🔢 Length Consistency (Numbers Only)

- Compute **average length** of all non-empty numeric values.
- Allowed range = `average ± 25%`.
- **Exception:** If average length `< 2`, skip length check.

---

## ⚖️ Consistency Rules

A record is **consistent** if:

- It matches the column’s detected/assigned pattern.
- If numeric → its length is within `average ± 25%` (unless average length < 2).
- If Request Number → numeric and > 0.
- If Address → contains valid chars **and at least one letter**.
- If Note → any non-null value is valid.

Otherwise → **inconsistent**.

---

## 📊 Output Report

The generated **Excel report** includes:

- File name
- Sheet/Table name
- Column name
- Detected pattern
- Average length and allowed range (for numbers)
- Consistency percentage
- Count of consistent and inconsistent values
- List of inconsistent values

---

## 💾 Output File Format

- Single **Excel file** (`.xlsx`) summarizing all processed files.
- Stored in **Google Drive**.

---

## 🛠️ Tech Stack

- **Python**
- **pandas** (Excel handling & analysis)
- **openpyxl** (Excel writing)
- **re** (pattern matching)
- **Google Drive API / PyDrive** (file access)
- colab

---
## ▶️ Usage in Google Colab

1. Open a new cell code in **Google Colab**:  
  copy and paste the code

2. Mount your Google Drive:  
  
   from google.colab import drive
   drive.mount('/content/drive')

 3. Locate the files you want to analy
    in your drive(your_location var)

 4. run the code
   The report will be saved to /content/drive/MyDrive/cleaned_consistency_report.xlsx
    



