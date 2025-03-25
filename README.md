# Excel Automation Script

## 📌 Overview
This script automates the process of updating an Excel file (`updated_excel2.xlsx`) by merging new data from another Excel file (`Excel1.xlsx`). It ensures data consistency, removes duplicates, and maintains a structured format.

## 🚀 Features
- Loads data from `Excel1.xlsx` and updates `updated_excel2.xlsx`
- Ensures column consistency
- Converts timestamps to the format: `YYYY-MM-DDTHH:MM:SS`
- Removes duplicate entries
- Saves the updated file automatically

## 📂 File Structure
```
ExcelAutomation/
│-- Excel1.xlsx             # Source Excel file
│-- updated_excel2.xlsx     # Updated Excel file (output)
│-- script.py               # Main Python script
│-- README.md               # Project documentation
```

## 🛠️ Installation & Setup
1. **Clone the repository**
   ```sh
   git clone https://github.com/<your-username>/Excel-Automation.git
   cd Excel-Automation
   ```
2. **Install dependencies**
   ```sh
   pip install pandas openpyxl
   ```
3. **Run the script**
   ```sh
   python script.py
   ```

## ⚡ How It Works
1. The script loads `Excel1.xlsx` and renames columns to match a predefined structure.
2. It converts the `TimeStamp` and `SourceTimeStamp` columns to `YYYY-MM-DDTHH:MM:SS` format.
3. It checks for an existing `updated_excel2.xlsx` file:
   - If it exists, it loads the data and appends new data from `Excel1.xlsx`.
   - If it doesn’t exist, it creates a new one.
4. Duplicate entries are removed, ensuring only unique records are saved.
5. The final data is saved in `updated_excel2.xlsx`.

## 📝 Columns in the Final Output
| Column          | Description |
|----------------|-------------|
| Name           | Well ID |
| TimeStamp      | Date of Sampling |
| SourceTimeStamp | Same as TimeStamp |
| Area          | Location Area |
| Latitude       | (Future implementation) |
| Longitude      | (Future implementation) |
| Type           | Parameter tested |
| Unit           | Measurement unit |
| Value          | Result value |
| Relationships  | (Future implementation) |

## 📌 Notes
- Make sure `Excel1.xlsx` is in the correct directory before running the script.
- The script automatically handles missing columns and formats timestamps correctly.

## 🤝 Contributing
Feel free to submit issues or contribute improvements via pull requests!

