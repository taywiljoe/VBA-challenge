VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()
    
    Dim ticker As String
    'Dim opendt, closeddt As Single
    Dim ws As Worksheet
    Dim j As Integer
    Dim i As Long
    Dim total_volume As Double
    Dim close_price As Double
    Dim open_price As Double
    Dim yearly_change As Double
    
    open_price = Cells(2, 3).Value
    
    
    
    
'The for loop that does the calculations
    
    For Each ws In ThisWorkbook.Worksheets
    j = 0
    total_volume = 0
    
'Labeling columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Loop through the ticker symbol to find the earlies date open and latest date closed

    'opendt = WorksheetFunction.Min(Range("B:B"))
    'Debug.Print opendt
    
    'closeddt = WorksheetFunction.Max(Range("B:B"))
    'Debug.Print closeddt
    
'Loop through each unique ticker and get the opening adn closing price for the year
    finalrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To finalrow
   
        'check to see if the cell in the next row is different
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            
            
            'output print results
            ws.Cells(2 + j, 9).Value = ws.Cells(i, 1).Value
           
            
            ws.Cells(2 + j, 12).Value = total_volume
            
            
        
            'yearly change
            'yearly_change = (close_price - open_price)
            
            'print yearly change
            'Range("J" & total_volume).Value = yearly_change
            
       
            
            j = j + 1
            total_volume = 0
        End If
        
        
        
        total_volume = total_volume + ws.Cells(i, 7).Value
    
    Next i
Next ws

End Sub
