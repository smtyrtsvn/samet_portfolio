# 🚀 Supply Chain & Merchandising Automation Portfolio

Welcome to my automation portfolio! This repository contains a collection of Python-based tools and scripts I developed to optimize daily operations in retail, merchandising, and supply chain management. 

The primary goal of these projects is to eliminate manual data entry, ensure 100% data accuracy across PLM/ERP systems, and accelerate the critical path from product development to final delivery.

## 🗂️ Projects Overview

### 1. 📦 Carton Label and PL Generator (Automated Logistics Documentation)
**Description:** A unified automation suite designed to handle the two most time-consuming paperwork tasks in the logistics phase: **Carton Labels (Koliüstü)** and **Packing Lists (Checklists)**.
**Business Impact:** By reading complex Excel-based packaging and model data, the suite generates formatted, print-ready shipment documents. It handles complex size matrices (from XS to Onesize) and automatically identifies "MIX" cartons to prevent warehouse sorting errors. This ensures 100% consistency between the physical shipment and the system data.
* **Tech Stack:** Python, `openpyxl` (Dynamic cell formatting and data manipulation), `glob`, `os`.

### 2. ⚡ PLM_line_sheet_fill_2026 (Data Entry Automation)
**Description:** An automation script designed to rapidly populate mandatory fields within PLM systems by extracting raw data from supplier Excel files.
**Business Impact:** Minimizes human error and drastically reduces the time spent on administrative data entry. Accelerates the Go-To-Market (GTM) process by ensuring technical data is synced instantly across design and sourcing teams.
* **Tech Stack:** Python, `pyautogui`, `pyperclip`, `openpyxl`, `xlrd`.

### 3. 🏷️ Barcode_sticker_generator_2026
**Description:** A dynamic tool that generates high-quality Code128 barcodes with brand-specific logo placement (e.g., Williams Sonoma, Pottery Barn).
**Business Impact:** Streamlines inventory management. Automated tag generation ensures that every SKU is accurately labeled with its specific department, vendor, and price information, reducing warehouse discrepancies.
* **Tech Stack:** Python, `pandas`, `PIL` (Pillow), `python-barcode`.

### 4. 🪪 Business_Card_Generator_2026
**Description:** A bulk generator for corporate identity assets featuring embedded vCard QR codes for instant digital contact sharing.
**Business Impact:** Automates the design work for HR and corporate branding. It intelligently handles text-wrapping for long names and titles, outputting print-ready files for large teams in seconds.
* **Tech Stack:** Python, `pandas`, `matplotlib`, `reportlab` (QR generation), `PIL`.

---

## 💡 Why This Matters
In fast-paced retail and manufacturing environments, agility is everything. By leveraging these Python automation tools, I have successfully:
- **Eliminated hours of manual data entry** in ERP, PLM, and logistics workflows.
- **Prevented operational bottlenecks** in packaging, labeling, and shipment documentation.
- **Ensured 100% Data Integrity** across the supply chain, from the first sample to the final delivery.

## 🛠️ Skills Demonstrated
`Python` | `Data Automation (Pandas, Openpyxl)` | `Logistics & Warehouse Documentation` | `Process Optimization`

---
*Feel free to explore the individual folders for source code and specific usage instructions.*
