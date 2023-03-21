Attribute VB_Name = "Module1"
Sub StockAnalysis()

'define and set worksheet

'Defining variable

Dim Ticker As String
Dim Openvalue As Double
Dim Closevalue As Double
Dim Yearlychange As Double
Dim PercentChange As Double
Dim Volume As Long
Dim volumetotal As Double
Dim summarytable As Integer
Dim Lastrow As Long
Dim start As Long


'loop through all the worksheets

For Each ws In ActiveWorkbook.Worksheets


'Assigning values to variable

volumetotal = 0
summarytable = 2
start = 2


'Create the column heading

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percent change"
    ws.Range("L1").Value = "Total Stock Volume"


     ws.Range("O2").Value = "Greatest Percentage Increase"
     ws.Range("O3").Value = "Greatest Percentage Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"


'create headins for column P & Q

     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"


'determining the last row to run the function

    Lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    

'loop through all tickers
    
    For i = 2 To Lastrow
    
        If i = Lastrow + 1 Then
        
    
    End If
    

'set ticker name by checking if all values are same, by adding +1 value

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
'Ticker Value
    
    Ticker = ws.Cells(i, 1).Value


'adding total volume
    
    volumetotal = volumetotal + ws.Cells(i, 7).Value

'Opening Value
    
    Openvalue = ws.Cells(start, 3).Value

'Closing value
    
    Closevalue = ws.Cells(i, 6).Value
    
'caluculating yearly change

    Yearlychange = Closevalue - Openvalue
    
'Calculating change in percentage

    PercentChange = Yearlychange / Openvalue
    
    start = i + 1
    
'Adding header titles

    ws.Cells(summarytable, 9).Value = Ticker
    ws.Cells(summarytable, 10).Value = Yearlychange
    ws.Cells(summarytable, 11).Value = FormatPercent(PercentChange, 2)
    ws.Cells(summarytable, 12).Value = volumetotal
    
    
        'Yearly Change
        If Yearlychange >= 0 Then
        
        ws.Cells(summarytable, 10).Interior.Color = vbGreen
        
        Else
        
        ws.Cells(summarytable, 10).Interior.Color = vbRed
        
        End If
        
    'Adding +1 to the summartable to go to next value
    
    summarytable = summarytable + 1
    
    
    'Reset Volume
    
    volumetotal = 0
    
    Else
    
    'Adding value to the total volume
    
    volumetotal = volumetotal + ws.Cells(i, 7).Value
    
    End If
    
    
    Next i
    
    
 'Setting variable back to 0 to go onto next run
 
    Ticker = ""
    volumetotal = 0
    
    
    
 ' Same process to next worksheet
 
    Next ws
    
    


End Sub
