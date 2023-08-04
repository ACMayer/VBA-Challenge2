Attribute VB_Name = "Module1"
Sub TestData()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
    On Error Resume Next
    
     Dim SummaryTable As Integer
     SummaryTable = 2
     Dim Brand_Total As Double
     Dim Yearly_Change As Double
     Dim Percentage_Change As Double
     Dim Last_Row As Long
     Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
     Dim StartValue As Double
     StartValue = 2
    Dim Brand_Name As String
     Dim Greatest_Increase As Double
       Dim Greatest_Decrease As Double
       Dim Greatest_Volume As Double
       Cells(2, 16).Value = "Greatest % Increase"
       Cells(3, 16).Value = "Greatest % Decrease"
       Cells(4, 16).Value = "Greatest Total Volume"
       Cells(1, 17).Value = "Ticker"
       Cells(1, 18).Value = "Value"
       Cells(1, 9).Value = "Ticker"
       Cells(1, 10).Value = "Yearly Change"
       Cells(1, 11).Value = "Percentage Change"
       Cells(1, 12).Value = "Total Stock Volume"
       Dim i As Long
       
       
               
       For i = 2 To Last_Row
        
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               Brand_Name = Cells(i, 1).Value
               Brand_Total = Brand_Total + Cells(i, 7).Value
               SummaryTable = SummaryTable + 1
               CloseValueEnd = Cells(i, 6).Value
               OpenValue = Cells(StartValue, 3).Value
               Yearly_Change = CloseValueEnd - OpenValue
               Percentage_Change = Yearly_Change / OpenValue
               Range("I" & SummaryTable).Value = Brand_Name
               Range("J" & SummaryTable).Value = Yearly_Change
               Range("K" & SummaryTable).Value = Percentage_Change
               Range("K" & SummaryTable).NumberFormat = "0.00%"
               Range("L" & SummaryTable).Value = Brand_Total

                Brand_Total = 0
                StartValue = i + 1
                 
                 If Yearly_Change > 0 Then
                 Range("J" & SummaryTable).Interior.ColorIndex = 4
                 ElseIf Yearly_Change < 0 Then
                 Range("J" & SummaryTable).Interior.ColorIndex = 3
                 End If
                 
                 
                    
                
         Else
               Brand_Total = Brand_Total + Cells(i, 7).Value
              
               
               End If
                
               
               Next i
               
               Cells(2, 18).Value = Application.WorksheetFunction.Max(Range("K3:K3002"))
               Cells(3, 18).Value = Application.WorksheetFunction.Min(Range("K3:K3002"))
               Cells(4, 18).Value = Application.WorksheetFunction.Max(Range("L3:L3002"))
               Range("R2:R3").NumberFormat = "0.00%"
               
                
                Greatest_Increase = Cells(2, 18).Value
                Greatest_Decrease = Cells(3, 18).Value
                Greatest_Volume = Cells(4, 18).Value
               
    
      
      
    If Err.Number <> 0 Then
        Debug.Print "Error on Sheet: " & ws.Name
        Err.Clear
    End If
    Set myrange = Range("I3:L3002")
    
    GIncrease = WorksheetFunction.Match(WorksheetFunction.Max(Range("K3:K3002" & RowCount)), Range("K3:K3002" & RowCount), 0)
    GDecrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K3:K3002" & RowCount)), Range("K3:K3002" & RowCount), 0)
    GVolume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L3:L3002" & RowCount)), Range("L3:L3002" & RowCount), 0)
    
    Range("Q2") = Cells(GIncrease + 2, 9)
    Range("Q3") = Cells(GDecrease + 2, 9)
    Range("Q4") = Cells(GVolume + 2, 9)
        
        
    
     ws.Activate
Next


  
End Sub




