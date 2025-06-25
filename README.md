# 🏗️ Ground Mount Quote Tool

This Python-based automation tool generates accurate and optimized bills of materials (BOM) for ground-mounted solar structures. Designed to minimize quoting time and eliminate manual Excel work, it allows engineers and sales teams to produce cost-effective system quotes using only basic panel and project inputs.

---

## 🔧 Features

- ✅ Inputs: 
  - Panel width, height, number of panels
  - Row configuration, tilt angle, and bay configuration
  - Location-specific wind loading consideration
- 🧮 Auto-calculates:
  - Quantity of structural members and brackets
  - Post and foundation requirements
  - Total number of modules and layout spacing
- 📤 Outputs:
  - A structured CSV file ready for import into **SAGE Intacct**
  - Customizable and editable quote sheet format
- ⏱️ **Reduces quoting time** from 30–60 minutes to **under 15 minutes**

---

## 🧠 Technologies Used

- Python
- `pandas` (data manipulation)
- `math` (calculation logic)
- Excel/CSV I/O

---

## 📁 Folder Usage

- `/input_files/` → Store your template or incoming panel data files
- `/output/` → Receives the CSV BOM files compatible with SAGE
- `/K8 Ground Mount Quote Tool 1.3.py` → Main script

---

## 📈 Business Impact

- Eliminated manual Excel formulas and layout adjustments
- Enabled faster client turnaround for quotations
- Improved consistency in bill of materials and quote accuracy
- CSV outputs are directly importable into SAGE Intacct, skipping manual line entry

---

## 🚀 How to Use

1. Open the script in Python 3
2. Input your project parameters when prompted (panel count, bay layout, tilt, etc.)
3. The tool will process the calculation and save the result in `output/quote.csv`
4. Import this CSV into SAGE or continue modifying it in Excel

---

## 📌 Notes

- Originally developed during employment at **Lumax Energy** for internal quoting automation
- This version is sanitized and shared for demonstration purposes
- Designed with simplicity and modularity in mind

---

## 📷 Screenshot Suggestion (Optional)

> Add a screenshot here showing a sample Excel output or a CSV snippet if desired
