
# 📘 Python Script Requirements: Excel Column Consistency Checker

## 1. Input

* Process multiple Excel files stored in Google Drive.
* For each file, read **all sheets (tables)**.

---

## 2. Pattern Detection Rules

For each column, the **first non-empty record** defines its pattern
(❗except for Address, Request Number, and Note, which have explicit rules).

### Allowed Patterns

| Pattern            | Rule                                                                                                                                                                                                                         | Notes                                                                 |
| ------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| **Arabic words**   | Allowed chars: Arabic letters `[\u0600-\u06FF]` + symbols `() _ - \\ / .`                                                                                                                                                    | First non-empty record defines pattern                                |
| **English words**  | Allowed chars: `A–Z`, `a–z`, spaces, `_`, `-`, `.`                                                                                                                                                                           | First non-empty record defines pattern                                |
| **Numbers**        | Digits only `[0-9]+`                                                                                                                                                                                                         | First non-empty record defines pattern; special length check (see §3) |
| **Dates**          | Recognizable formats (e.g., `YYYY-MM-DD`, `DD/MM/YYYY`, etc.)                                                                                                                                                                | First non-empty record defines pattern                                |
| **Address**        | Applied if column name contains `"address"` (case-insensitive). Allowed chars: Arabic, English, digits, spaces, `, . - / #`. **Must contain at least one letter** (Arabic or English). Numbers-only values are inconsistent. | Independent of first record                                           |
| **Request Number** | Applied if column name = `"REQ_NO"` or `"REQUEST_NO"` (case-insensitive). Consistent only if value is a **number > 0**.                                                                                                      | Independent of first record                                           |
| **Note**           | Applied if column name contains `"note"`, `"desc"`, or `"desca"` (case-insensitive). Any value is accepted.                                                                                                                  | Independent of first record                                           |

---

## 3. Length Consistency

* Applies **only to Numbers**.
* Compute average length of all non-empty numeric values.
* Allowed range = `average ± 25%`.
* **Exception:** If average length `< 2`, **skip length check** (i.e., all valid numbers are accepted regardless of length).

---

## 4. Null Values

* Ignore null/empty values.
* If entire column = null → skip column.

---

## 5. Consistency Rules

A record is **consistent** if:

* It matches the column’s detected/assigned pattern.
* If numeric → its length is within `average ± 25%` (unless average length < 2, then skip).
* If Request Number → it is numeric and > 0.
* If Address → it has allowed characters and **at least one letter**.
* If Note → any non-null value is consistent.

Otherwise → inconsistent.

---

## 6. Report Output

For each column, report should include:

* File name
* Sheet/Table name
* Column name
* Detected pattern
* Average length and allowed range (if numeric; optional for others)
* Consistency percentage (`consistent ÷ total non-empty × 100`)
* Count of consistent and inconsistent values
* List of all inconsistent values

---

## 7. File Format

* Save as a single Excel file (`.xlsx`) summarizing all processed files.
* Store in Google Drive. --> write this as read me file for github 
