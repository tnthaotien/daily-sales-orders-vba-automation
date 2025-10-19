Attribute VB_Name = "Module1"
Sub ImportOrders_ByDateAndCategory_Modular()
    Dim targetDate As Date, category As String, filePath As String
    Dim sheetNames As Variant, newSheetName As String
    Dim wbSource As Workbook, wsTarget As Worksheet

    targetDate = GetInputDate()
    category = GetInputCategory()
    sheetNames = Array("Snack", "Conf", "Noodle")
    filePath = GetInputFilePath()
    If filePath = "False" Then Exit Sub

    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)

    newSheetName = BuildSheetName(targetDate, category)
    Set wsTarget = PrepareTargetSheet(ThisWorkbook, newSheetName)

    CopyOrdersByDate wbSource, wsTarget, sheetNames, targetDate

    wbSource.Close SaveChanges:=False

    FormatOrderSheet wsTarget

    MsgBox "Orders imported to sheet: " & newSheetName, vbInformation
End Sub

Function GetInputDate() As Date
    Dim inputDateStr As String
    inputDateStr = InputBox("Enter the target date (dd/mm/yyyy):", "Select Date")
    If inputDateStr = "" Or Not IsDate(inputDateStr) Then
        MsgBox "Invalid date!", vbCritical: End
    End If
    GetInputDate = CDate(inputDateStr)
End Function

Function GetInputCategory() As String
    Dim category As String
    category = InputBox("Enter category name (e.g. Oils, Flours, Rice):", "Category")
    If category = "" Then MsgBox "No category entered!", vbExclamation: End
    GetInputCategory = category
End Function

Function GetInputFilePath() As String
    GetInputFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select sales file")
End Function

Function BuildSheetName(targetDate As Date, category As String) As String
    BuildSheetName = "Orders_" & Format(targetDate, "yyyymmdd") & "_" & category
End Function

Function PrepareTargetSheet(wb As Workbook, newSheetName As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets(newSheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set PrepareTargetSheet = wb.Sheets.Add
    PrepareTargetSheet.Name = newSheetName
End Function

Sub CopyOrdersByDate(wbSource As Workbook, wsTarget As Worksheet, _
    sheetNames As Variant, targetDate As Date)
    Dim wsSource As Worksheet, lastRow As Long, lastCol As Long
    Dim i As Long, r As Long, tgtRow As Long, headerCopied As Boolean

    tgtRow = 2: headerCopied = False
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set wsSource = wbSource.Sheets(sheetNames(i))
        lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

        If Not headerCopied Then
            wsTarget.Cells(1, 1).value = "No."
            wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol)).Copy wsTarget.Cells(1, 2)
            If wsTarget.Cells(1, 2).value = "FF" Then wsTarget.Cells(1, 2).value = "Date of SD"
            headerCopied = True
        End If

        For r = 2 To lastRow
            If IsDate(wsSource.Cells(r, 1).value) Then
                If CDate(wsSource.Cells(r, 1).value) = targetDate Then
                    wsTarget.Cells(tgtRow, 1).value = tgtRow - 1
                    wsSource.Range(wsSource.Cells(r, 1), wsSource.Cells(r, lastCol)).Copy wsTarget.Cells(tgtRow, 2)
                    tgtRow = tgtRow + 1
                End If
            End If
        Next r
    Next i
End Sub

Sub FormatOrderSheet(wsTarget As Worksheet)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    lastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column

    If lastRow > 1 Then
        Dim dataRange As Range
        Set dataRange = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastRow, lastCol + 1))

        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        With dataRange.Font
            .Name = "Arial"
            .Size = 11
        End With

        With dataRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        wsTarget.Rows(1).Font.Bold = True
        dataRange.EntireColumn.AutoFit
        dataRange.EntireRow.AutoFit
    End If
End Sub


