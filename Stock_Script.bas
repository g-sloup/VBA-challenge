Attribute VB_Name = "Module1"
Sub Stock_Market()
    
    'Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    'Add Summary Table
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    'Set Variables
    Dim Ticker_Symbol As String
    Dim Total_Volume As Double
        Total_Volume = 0
    Dim Opening_Day As Double
    Dim Closing_Day As Double
    Dim Percent_Change As Double
    Dim Yearly_Change As Double
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
    Dim i As Long

    'Determine the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through stock data
    For i = 2 To lastrow
        
        'Determine Opening Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
            
            'Add to the Total Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            'Print Ticker & Volume
            ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
            ws.Range("M" & Summary_Table_Row).Value = Total_Volume
            
            'Reset the Total Volume
            Total_Volume = 0
            
            'Set Closing Day Value
            Closing_Day = ws.Cells(i, 6).Value
            


            If Opening_Day = 0 Then
                Yearly_Change = 0
                Percent_Change = 0
    
            Else
                'Calculate Change
                Yearly_Change = Closing_Day - Opening_Day
                Percent_Change = (Closing_Day - Opening_Day) / Opening_Day
            
            End If

        'Print Yearly and Percent Change
        ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("L" & Summary_Table_Row).Value = Percent_Change
        ws.Range("L" & Summary_Table_Row).Style = "Percent"
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"

        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
        Opening_Day = ws.Cells(i, 3)
    

    Else
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If


    Next i

    'Format Color
    For i = 2 To lastrow

    If ws.Range("K" & i).Value > 0 Then
        ws.Range("K" & i).Interior.ColorIndex = 4

        ElseIf ws.Range("K" & i).Value < 0 Then
        ws.Range("K" & i).Interior.ColorIndex = 3
        
    End If

    Next i
    
Next ws


End Sub
'-----------
         
        
