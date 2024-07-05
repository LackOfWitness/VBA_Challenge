Sub TickerPriceMacro()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range
    Dim ticker As String
    Dim minDate As Date
    Dim maxDate As Date
    Dim openPrice As Double
    Dim closePrice As Double
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Step 1: Create the CONCAT column
        ws.Range("H1").Value = "CONCAT"
        For i = 2 To lastRow
            ws.Cells(i, "H").FormulaR1C1 = ws.Cells(i, "A").Value & Format(ws.Cells(i, "B").Value, "mm/dd/yyyy")
        Next i
        
        ' Step 2: Copy unique tickers to column J
        ws.Columns("A:A").Copy
        ws.Range("J1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        ws.Range("$J$1:$J$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo
        
        ' Step 3: Calculate earliest and latest dates
        ws.Range("K1").Value = "<open date>"
        ws.Range("L1").Value = "<close date>"
        ws.Range("M1").Value = "<open price>"
        ws.Range("N1").Value = "<close price>"
        ws.Range("O1").Value = "Quarterly Change"
        ws.Range("P1").Value = "Percent Change"
        ws.Range("Q1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
            ticker = ws.Cells(i, "J").Value
            minDate = Application.WorksheetFunction.MinIfs(ws.Range("B:B"), ws.Range("A:A"), ticker)
            maxDate = Application.WorksheetFunction.MaxIfs(ws.Range("B:B"), ws.Range("A:A"), ticker)
            
            ws.Cells(i, "K").Value = Format(minDate, "mm/dd/yyyy")
            ws.Cells(i, "L").Value = Format(maxDate, "mm/dd/yyyy")
            
            openPrice = Application.WorksheetFunction.Index(ws.Range("C:C"), Application.WorksheetFunction.Match(ticker & Format(minDate, "mm/dd/yyyy"), ws.Range("H:H"), 0))
            closePrice = Application.WorksheetFunction.Index(ws.Range("F:F"), Application.WorksheetFunction.Match(ticker & Format(maxDate, "mm/dd/yyyy"), ws.Range("H:H"), 0))
            
            ws.Cells(i, "M").Value = openPrice
            ws.Cells(i, "N").Value = closePrice
            ws.Cells(i, "O").Value = closePrice - openPrice
            ws.Cells(i, "P").Value = IIf(openPrice <> 0, (closePrice - openPrice) / openPrice, 0)
            ws.Cells(i, "Q").Value = Application.WorksheetFunction.SumIfs(ws.Range("G:G"), ws.Range("A:A"), ticker)
        Next i
        
        ' Step 4: Apply conditional formatting to columns O and P
        Set rng = ws.Range("O2:O" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)
        rng.FormatConditions.Delete
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$O2>0")
            .Interior.Color = RGB(0, 255, 0) ' Green
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$O2<0")
            .Interior.Color = RGB(255, 0, 0) ' Red
        End With
        
        Set rng = ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)
        rng.FormatConditions.Delete
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$P2>0")
            .Interior.Color = RGB(0, 255, 0) ' Green
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$P2<0")
            .Interior.Color = RGB(255, 0, 0) ' Red
        End With
        
        ' Step 5: Summary statistics
        ws.Range("S2").Value = "Greatest % Increase"
        ws.Range("S3").Value = "Greatest % Decrease"
        ws.Range("S4").Value = "Greatest Total Volume"
        ws.Range("T1").Value = "Ticker"
        ws.Range("U1").Value = "Value"
        
        ws.Range("U2:U3").NumberFormat = "0.00%"
        ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row).NumberFormat = "0.00%"
        
        ' Step 6: Hide intermediate columns
        ws.Columns("H:H").EntireColumn.Hidden = True
        ws.Columns("K:N").EntireColumn.Hidden = True

        ' Step 7: Extend columns J, O, P, Q, S, T, U
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        ws.Columns("S:S").EntireColumn.AutoFit
        ws.Columns("T:T").EntireColumn.AutoFit
        ws.Columns("U:U").EntireColumn.AutoFit

    Next ws


    
    MsgBox "Ticker Price Macro has completed successfully for all worksheets!"
End Sub
