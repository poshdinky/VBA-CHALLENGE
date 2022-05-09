Attribute VB_Name = "Module1"
Sub GenerateData()


Dim LR As Long, LineT As Long, LineStart As Long, i As Long
Dim Sh As Worksheet

'--- Optimize the code

Application.ScreenUpdating = False

'Loop trough all the sheet in the workbook

For Each Sh In Worksheets

    'Activate the sheet to become the active sheet
    
    Sh.Activate
    'Clean the results col
    Range("I:Z").ClearContents
    
    'Put the column's headers
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % increase"
    Range("O3") = "Greatest % decrease"
    Range("O4") = "Greatest total volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    'Initialize variables
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    LineStart = 2
    LineT = 2
    
    'Loop trough all stocks
    
    For i = 2 To LR
        'Check if the ticker change
        If Range("A" & i + 1) <> Range("A" & i) Then
            'Get the ticker
            Range("I" & LineT) = Range("A" & LineStart)
            'Get the yearly change
            Range("J" & LineT) = Range("F" & i) - Range("C" & LineStart)
            '-------- Change color : + green, - red
            If Range("J" & LineT) >= 0 Then
                Range("J" & LineT).Interior.Color = vbGreen
            Else
                Range("J" & LineT).Interior.Color = vbRed
            End If
            
            
            'Get the percent change
            If Range("C" & LineT) = 0 Then
            
                Range("K" & LineT) = Format(0, "Percent")
            Else
                Range("K" & LineT).Value = (Range("F" & i) - Range("C" & LineStart)) / Range("C" & LineStart) '(LastClose-FirstOpen)/FirstOpen
            End If
            'Get the total volume
            Range("L" & LineT) = Application.WorksheetFunction.Sum(Range("G" & LineStart & ":G" & i))
            
            'next ticker line in result
            LineT = LineT + 1
            'Jump to the next ticker in the col A
            LineStart = i + 1
        
        End If
    
    Next
    
    'Get the indicators
    Range("Q2") = Application.WorksheetFunction.Max(Range("K2:K" & LR))
    Range("P2") = Range("I" & Application.WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & LR), 0) + 1)
    
    Range("Q3") = Application.WorksheetFunction.Min(Range("K2:K" & LR))
    Range("P3") = Range("I" & Application.WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & LR), 0) + 1)
    
    Range("Q4") = Application.WorksheetFunction.Max(Range("L2:L" & LR))
    Range("P4") = Range("I" & Application.WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & LR), 0) + 1)

    
    'Autofit columns
    Range("K:K").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    Columns("A:Q").AutoFit
Next

Application.ScreenUpdating = True

'Inform the end
MsgBox ("Stocks Report has been generated successfuly !")
End Sub



