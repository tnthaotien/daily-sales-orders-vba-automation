# daily-sales-orders-vba-automation
**1. PROJECT OVERVIEW**

This repository showcases Excel VBA automation projects that streamline daily sales operations and reduce repetitive tasks. By leveraging modular VBA functions, these tools improve reporting accuracy, save time, and support better decision-making.

The repository contains two main automation projects:
- Import Orders by Date and Category – Consolidates order data from multiple sheets into a formatted report.
- Map Data to New Sheet – Maps raw sales data into a standardized structure for reporting and analysis.

**1.1. Features**

- Automated Sales Reporting – Generates daily-to-monthly reports with minimal manual work.

- Data Mapping & Transformation – Standardizes raw sales data into clean, structured output.

- Error Reduction – Cuts manual error checks by ~40%.

- Workflow Efficiency – Optimized VBA workflows save up to 30% preparation time.

- Reusable Modular Functions – Each task (inputs, mapping, formatting) is broken into functions for easier reuse.

**1.2. Tech Stack**

- MS Excel

- VBA (Visual Basic for Applications)

**2. PROJECT DETAILS**

**2.1. Import Orders by Date and Category**

- Prompts for date, category, and file path.

- Consolidates orders from multiple sheets (e.g., Snack, Confectionery, Instand Noodle).

- Builds a new worksheet with clean formatting and sequential order numbers.

```
Sub ImportOrders_ByDateAndCategory_Modular()
    targetDate = GetInputDate()
    category = GetInputCategory()
    CopyOrdersByDate wbSource, wsTarget, sheetNames, targetDate
    FormatOrderSheet wsTarget
    MsgBox "Orders imported successfully!"
End Sub
```
**2.2. Map Data to New Sheet**

- Maps raw sales data headers into a standardized reporting format.

- Handles column mismatches with dictionary mapping.

- Applies correct number formats (%, MT, integers).

```
Sub MapDataToNewSheet_Modular()
    sheetName = InputBox("Enter the source sheet name:")
    Set wsSource = GetWorksheet(ThisWorkbook, sheetName)
    Set colDict = MapHeaders(wsSource, headers, sourceCols)
    CopyAndFormatData wsSource, wsNew, headers, colDict
    MsgBox "Data mapping completed successfully!"
End Sub
```
**3. USE CASE**

Automating monthly or daily sales reporting.

**4. IMPACT**

- 40% fewer manual error checks through automation.

- 30% faster report preparation by optimizing VBA workflows.

- Improved cross-team visibility with cleaner, standardized reports

**Author: Tran Ngoc Thao Tien**

- LinkedIn: [linkedin.com/in/tientranngocthao](https://www.linkedin.com/in/tientranngocthao/)

- Email: tientrank44.ueh@gmail.com
