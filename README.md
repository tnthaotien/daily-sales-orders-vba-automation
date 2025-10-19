**1. PROJECT OVERVIEW**

This repository contains a **modular Excel VBA automation tool** designed to streamline daily sales reporting. The macro consolidates sales orders by date and category from multiple source sheets into a single, cleanly formatted worksheet—eliminating repetitive copy-paste and reducing manual errors.

**1.1. Features**

- **Automated Reporting**: Generates daily consolidated reports in seconds with just a few prompts.
- **User-Friendly Inputs**: Prompts for date, category, and file path—no manual setup needed.
- **Data Consolidation**: Combines orders from multiple sheets (e.g., Snacks, Confectionery, Noodle) based on date and category.
- **Professional Formatting**: Clean numbering, bold headers, Arial 11, auto-fit columns, all ready for analysis.
- **Reusable Modular Code**: Each task is a dedicated function/sub for easy maintenance and reusability.

**1.2. Tech Stack**

- MS Excel

- VBA (Visual Basic for Applications)
  

**2. PROJECT DETAILS**

**Import Orders by Date and Category**

- Prompts for date, category, and file path.

- Consolidates orders from multiple sheets (e.g., Snack, Confectionery, Instant Noodle).

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
_DEMO_
<img width="2738" height="722" alt="image" src="https://github.com/user-attachments/assets/2a21d27c-0172-43f6-96f7-ed1dc4d7721a" />

**4. USE CASE**

Automating daily sales reporting.


**5. IMPACT**

- 40% fewer manual error checks through automation.

- 30% faster report preparation by optimizing VBA workflows.

- Improved cross-team visibility with cleaner, standardized reports

**Author: Tran Ngoc Thao Tien**

- LinkedIn: [linkedin.com/in/tientranngocthao](https://www.linkedin.com/in/tientranngocthao/)

- Email: tientrank44.ueh@gmail.com
