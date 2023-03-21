Attribute VB_Name = "Module2"
Sub Bonus_Rangevalue()

'Define Variables

Dim Lastrow As Double
Dim Percentincrease As Double
Dim Percentdecrease As Double
Dim Totalvolume As Double



'looping through allt he worksheets

For Each ws In ActiveWorkbook.Worksheets


'Setting last row command

     Lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    
    'finding Maximum range or increase in value
    
    maxvalue = WorksheetFunction.Max(ws.Range("K2:K" & Lastrow))
    maxindex = WorksheetFunction.Match(maxvalue, ws.Range("K2:K" & Lastrow), 0)
    
    ws.Range("Q2") = "%" & maxvalue * 100
    ws.Range("P2") = Cells(maxindex + 1, 9)
    
    
    'finding Minimum range or decrease in value
    
    
    minvalue = WorksheetFunction.Min(ws.Range("K2:K" & Lastrow))
    minindex = WorksheetFunction.Match(maxvalue, ws.Range("K2:K" & Lastrow), 0)
    
    ws.Range("Q3") = "%" & minvalue * 100
    ws.Range("P3") = Cells(minindex + 1, 9)
    
    
    
    'finding volume
    
    
    maxvolvalue = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))
    maxindex = WorksheetFunction.Match(maxvolvalue, ws.Range("L2:L" & Lastrow), 0)
    
    ws.Range("Q4") = maxvolvalue
    ws.Range("P4") = Cells(maxindex + 1, 9)
    
    
Next ws



End Sub
