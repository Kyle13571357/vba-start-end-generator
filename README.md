# vba-start-end-generator
# VBA Transportation Start-End Autofill Tool

This VBA macro automates the process of filling in start and end locations for transportation reimbursement forms based on employee information and predefined route mappings. It supports two terminal types and handles direction logic dynamically.

---

## 🚀 Features

- Automatically looks up route information based on employee ID
- Dynamically determines start and end locations using a predefined reference table
- Supports direction-aware logic: home → terminal or terminal → home
- Works across multiple worksheets (e.g., Sheet 6 to Sheet 95)
- Includes error handling, user prompts, and status messages

---

## 📋 How It Works

- The reference table is stored in **Sheet "WorkSheet1"**:
  - `A3:E100` for Terminal 1 lookups
  - `I3:M100` for Terminal 2 lookups

- The macro scans **employee worksheets**:
  - Sheets 6–49: Terminal 1
  - Sheets 50–95: Terminal 2

- For each employee (rows 4 to 26):
  - It checks the employee’s ID (Column C)
  - Performs a `VLOOKUP` in the respective reference table
  - Fills the result in Column F as the start/end location
  - If the cell `L13` contains a home address, it inserts the direction:
    - e.g., `Home → Terminal1` or `Terminal1 → Home`

---

## 🛠 Example Sheet Structure

| A | B | C (Employee ID) | ... | F (Start–End Output) | 
|---|---|------------------|-----|-----------------------|

---

## 📂 File Structure

- `GenerateStartEndLocations.bas` — the core macro module
- `README.md` — this documentation
- *(optional)* `docs/example_layout.png` — visual layout of the expected worksheet

---

## 🧠 Engineering Design Highlights

- **Data-driven**: Logic depends on the reference tables, not hardcoded rules
- **Low maintenance**: Update the mapping sheets as needed; no code modification required
- **Practical automation**: Reduces manual entry across 90+ sheets

---

## 📄 License

This tool is released for non-commercial, internal use and educational demonstration.  
MIT or similar license may apply.
