Attribute VB_Name = "Module1"
Sub VBA_Worksheets()

' Found how to update all WS from this website:
'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html

    Dim WS As Worksheet
    Application.ScreenUpdating = False
    For Each WS In Worksheets
        WS.Select
        Call Stock_VBA
    Next
    Application.ScreenUpdating = True
 
End Sub

Sub Stock_VBA()

    ' Set initial variables
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Volume As LongLong
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
        
    'Set up Summary Table Output
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change ($)"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Set up dynamic last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all stocks
    For i = 2 To LastRow
    
        ' Check if Ticker doesn't match previous Ticker
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            'Grab Ticker and put into summary table
            Ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = Ticker
            
            'Grab Open price for calculations
            Open_Price = Cells(i, 3).Value
            'Range("M" & Summary_Table_Row).Value = Open_Price
            
            'Grab Volume
            Volume = Cells(i, 7)
         
         
         'Else if Ticker doesn't match next Ticker
         ElseIf Cells(i, 1) <> Cells(i + 1, 1) Then
            
            'Grab Closeing price for calculations
            Close_Price = Cells(i, 6).Value
            'Range("N" & Summary_Table_Row).Value = Close_Price
            
            'Calculate Yearly Change
            Yearly_Change = Close_Price - Open_Price
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            ' Conditional formatting on Yearly Change
                If Yearly_Change >= 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else: Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
            'Calculate Percent Change
                If Open_Price > 0 Then
                    Percent_Change = ((Close_Price - Open_Price) / Open_Price)
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Else: Range("K" & Summary_Table_Row).Value = ""
                End If
            
            'Add to Volume and print into summary table
            Volume = Volume + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = Volume
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
        'Else if ticker matches next ticker
        Else: Cells(i - 1, 1).Value = Cells(i, 1).Value
        
        'Add to Volume
            Volume = Volume + Cells(i, 7).Value
            
            
        End If
        
            
    Next i
    
    ' Set up Bonus summary table
    
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total As LongLong
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Found max formula from google search and this site:
    'https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba/42633375
    'Finding Greatest % Increase and inputting into summary table
    Greatest_Increase = Application.WorksheetFunction.Max(Range("k:k"))
    Range("Q2").Value = Greatest_Increase
    Range("Q2").NumberFormat = "0.00%"
    
    'Finding Greatest % Decrease and inputting into summary table
    Greatest_Decrease = Application.WorksheetFunction.Min(Range("k:k"))
    Range("Q3").Value = Greatest_Decrease
    Range("Q3").NumberFormat = "0.00%"
    
    'Finding Greatest Total Volume
    Greatest_Volume = Application.WorksheetFunction.Max(Range("l:l"))
    Range("Q4").Value = Greatest_Volume
    
    'tying tickers back to Max and Mins
    For i = 2 To LastRow
        
        If Cells(i, 11).Value = Greatest_Increase Then
            Range("P2").Value = Cells(i, 9)
        ElseIf Cells(i, 11).Value = Greatest_Decrease Then
            Range("P3").Value = Cells(i, 9)
        End If
        
        If Cells(i, 12).Value = Greatest_Volume Then
            Range("P4").Value = Cells(i, 9)
        End If
        
    Next i
    
    ' Adjusting column width
    ' Found via google search and this site:
    'https://stackoverflow.com/questions/24058774/excel-vba-auto-adjust-column-width-after-pasting-data
    Columns("A:Q").Select
    Selection.EntireColumn.AutoFit

'MsgBox ("VBA done.")

End Sub

