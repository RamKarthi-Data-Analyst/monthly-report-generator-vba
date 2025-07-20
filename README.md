# monthly-report-generator-vba
Excel VBA macro to automate monthly order report generation
# Monthly Order Report Generator (Excel VBA)

This project uses Excel VBA to automate the generation of monthly order reports from customer data.

## 🧩 Features
- Removes duplicate dates
- Filters data by month using AutoFilter
- Creates separate sheets for each month
- Saves hours of manual work!

## 📂 Files Included
- `Generate Monthly Reports.xlsm`: Macro-enabled Excel workbook [Generate Monthly Report](https://github.com/RamKarthi-Data-Analyst/monthly-report-generator-vba/blob/main/Generate%20Monthly%20Reports.xlsm)
- `Customer report.xlsx`: Sample input data[Customer_Report](https://github.com/RamKarthi-Data-Analyst/monthly-report-generator-vba/blob/main/Customer%20report.xlsx)
- `VBA_Monthly_Report_Project_Summary.pdf`: Project summary [PDF](https://github.com/RamKarthi-Data-Analyst/monthly-report-generator-vba/blob/main/VBA_Monthly_Report_Project_Summary.pdf)

## 💡 Code Snippet
```vba
For Each cell In wsDates.Range("A2:A13")
    wsData.AutoFilterMode = False
    wsData.Range("A1").AutoFilter Field:=1, Criteria1:=cell.Value
    wsData.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = Format(cell.Value, "MMM_YYYY")
    ActiveSheet.Paste
Next cell

