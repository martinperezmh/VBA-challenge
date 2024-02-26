Attribute VB_Name = "Module"
Sub Loop_OneYear()
Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("I1:L1").WrapText = True
    
    
        Dim lastrow As Long
        Dim Ticker As String
        Dim RecallRow As Long
            RecallRow = 2
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & RecallRow).Value = Ticker
                RecallRow = RecallRow + 1
            End If
        Next i
        
    
        Dim Yearly_Change As Double
            Yearly_Change = 0
        Dim Opening As Double
            Opening = 0
        Dim Closing As Double
            Closing = 0
        Dim RecallRow1 As Long
            RecallRow1 = 2
        
        Dim Yearly_Change_Percent As Double
        
        Opening = ws.Cells(2, 3).Value
        
        Dim Max_Percent As Double
            Max_Percent = 0
        Dim Min_Percent As Double
            Min_Percent = 0
        Dim Max_Ticker As String
            Max_Ticker = ""
        Dim Min_Ticker As String
            Min_Ticker = ""
    
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Closing = ws.Cells(i, 6).Value
                Yearly_Change = Closing - Opening
                
                    If (Yearly_Change > 0) Then
                        ws.Range("J" & RecallRow1).Interior.ColorIndex = 4
                    ElseIf (Yearly_Change <= 0) Then
                        ws.Range("J" & RecallRow1).Interior.ColorIndex = 3
                    End If
                
                    If Opening <> 0 Then
                        Yearly_Change_Percent = ((Closing - Opening) / Opening)
                        
                    End If
                
                    If (Yearly_Change_Percent > Max_Percent) Then
                        Max_Percent = Yearly_Change_Percent
                        Max_Ticker = Ticker
                    ElseIf (Yearly_Change_Percent < Min_Percent) Then
                        Min_Percent = Yearly_Change_Percent
                        Min_Ticker = Ticker
                    End If
                
                RecallRow1 = RecallRow1 + 1
                 
                Opening = ws.Cells(i + 1, 3).Value
                
                ws.Range("Q2").Value = Max_Percent
                ws.Range("Q3").Value = Min_Percent
                ws.Range("Q2:Q3").NumberFormat = "0.00%"
                ws.Range("P2").Value = Max_Ticker
                ws.Range("P3").Value = Min_Ticker
                
                
                ws.Range("J" & RecallRow1).Value = Yearly_Change
                ws.Range("K" & RecallRow1).Value = Yearly_Change_Percent
                ws.Range("K" & RecallRow1).NumberFormat = "0.00%"
                
                
                
                Yearly_Change_Percent = 0
            End If
        Next i
        
    
        Dim RecallRow2 As Long
            RecallRow2 = 2
            Total_Vol = 0
            Total_Vol = ws.Cells(2, 7).Value
        
        Dim Max_Volume As Long
            Max_Volume = 0
        Dim Total_Volume As Long
            'Total_Volume = 0
            
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = Cells(i, 1).Value
                Total_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            
                ws.Range("L" & RecallRow2).Value = Total_Stock_Vol
                RecallRow2 = RecallRow2 + 1
                
                
            
                If (Total_Vol > Max_Volume) Then
                    Total_Vol = ws.Cells(i, 12).Value
                    'Total_Volume = Total_Vol
                    Max_Volume = Total_Vol
                    Max_Ticker = Ticker
                End If
                
                Total_Stock_Vol = 0
                
                ws.Range("Q4").Value = Max_Volume
                ws.Range("P4").Value = Max_Ticker
                
                
                
            Else
               Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                
            End If
               
        Next i
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
     Next ws
End Sub