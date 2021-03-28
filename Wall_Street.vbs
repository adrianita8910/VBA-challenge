Attribute VB_Name = "Module1"
Sub Wall_Street()

    'Loop to loop trough all of the worksheets
    For Each ws In Worksheets
    
     'Grab the value of the last row of the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Name of the worksheet
    Dim WorksheetName As String

    'Grab the Worksheet name
    WorksheetName = ws.Name
 
    'Variable to hold the Ticker Symbol
    Dim Ticker As String
    
     
    'Total for Yearly
    Dim Yearly As Double
    Yearly = 0
    
    'Variable to hold percentage change
    Dim Percentage As Double
    Percentage = 0
    
    'Variable to hold the volumen
    Dim Volumen As Long
    Volume = 0
    
    
    'Variable to hold the summary table row
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    For Row = 2 To lastRow
    

    
        'Check to see if we have changed ticker symbol
        If (ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value) Then
        
        'Change the ticker Symbol
        Ticker = ws.Cells(Row, 1).Value
        
        'Add to the yearly change
        Yearly = Yearly + ws.Cells(Row, 3).Value - ws.Cells(Row, 6).Value
        
        'Add to the Percetage change
        Percentage = ws.Cells(Row, 3).Value - ws.Cells(Row, 6).Value
              
        'Add to the volume
        Volume = Volume + ws.Cells(Row, 7).Value
        
        'Update the summary table
        'Add the ticker Symbol
        ws.Range("I" & summaryTableRow).Value = Ticker
        
        'Add Yearly change
        ws.Range("J" & summaryTableRow).Value = Yearly
        
        'Add Percantage Change
        ws.Range("K" & summaryTableRow).Value = Percentage
     
        'Add Volume Change
        ws.Range("L" & summaryTableRow).Value = Volume
        
         'Add the name of Ticker
        ws.Range("I1").Value = "Ticker"
        
         'Add the name of Yearly Change
        ws.Range("J1").Value = "Yearly Change"
        
         'Add the name of Percentage Change
        ws.Range("K1").Value = "Percentage Change"
        
         'Add the name of Total Stock Volume
        ws.Range("L1").Value = "Total Stock Volume"
        
        
            
        'Reset the yearly change
        Yearly = 0
        
         'Reset the Percentage change
        Percetange = 0
        
         'Reset the Volume change
       Volume = 0
        
        'Add one to the summary Table row count
        summaryTableRow = summaryTableRow + 1
        
        Else
        'If the ticker symbol is still the same
        'Add on to the current Yearly Change
        Yearly = Yearly + ws.Cells(Row, 3).Value - ws.Cells(Row, 6).Value
        
        'Add on to the current Percentage Change
        Percentage = ws.Cells(Row, 3).Value - ws.Cells(Row, 6).Value
        
         'Add on to the current Volume Change
       Volume = Volume + ws.Cells(Row, 7).Value
        
        
        End If
    
    
    Next Row

      Exit For
    
    Next ws
    
End Sub







