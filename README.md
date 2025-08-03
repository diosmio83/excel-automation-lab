# excel-automation-lab
# 🧬 LIMS Data Import & Validation Automation

## 🧭 Project Overview
Automated Excel processing for LIMS data using VBA and Power Automate Desktop (optional enhancement).

---

## 📌 Features
- Match and validate data using `Raw Data ID`
- Format, insert, highlight, and document import results
- Optional: Launch VBA via Power Automate Desktop flow
- Color-coded result visualization
- Logging unmatched entries

---

## 🧰 Technologies Used
- Excel VBA (`mod_ProcessingImport.bas`)
- Power Automate Desktop (UI automation flow)
- Excel dummy files (anonymized sample)

---

## 🧠 Workflow Overview

### 🔸 Part 1 – VBA-Only Version

1. User opens the Excel macro-enabled workbook
2. User clicks a button or runs the macro manually
3. File picker opens (source file selection)
4. Matching, transfer, formatting, coloring, and logging occurs

📁 Main script: `mod_ProcessingImport.bas`

---

### 🔸 Part 2 – With Power Automate Desktop (Optional)

1. User launches a PAD flow
2. PAD opens Excel
3. PAD triggers the macro (via UI)
4. PAD waits for macro to finish
5. Fully automated process: no Excel interaction required

📁 Flow screenshots: `/screenshots/`  
📁 Flow export (PDF): `PowerAutomate_UIFlow.pdf` *(optional)*

---

## 📂 Project Structure

