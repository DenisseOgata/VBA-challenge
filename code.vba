     Option Explicit
Sub Test_Stocks()
     
    Dim LastRow As Long
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Change As Double
    Dim PrvTicker As String
    Dim Ticker As String
    Dim NextTicker As String
    Dim Printer As Integer
    Dim i As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim TotStcVol As Double
    Dim PChange As Double
    Dim MaxChange As Double
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim MinChange As String
    Dim ws As Worksheet
    Dim Cell As Range
    Dim j As Integer
    
    
    
                    
                    For Each ws In Worksheets
                    
            
                     
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row  'shows the last row at the end of column and go 1 cell up
    
  
    
    
        j = 0
        
        Printer = 2
        
        Change = 0
        
        MaxChange = 0

    
    For i = 2 To LastRow
        
            PrvTicker = ws.Cells(i - 1, 1).Value 'Title in the variable
                Ticker = ws.Cells(i, 1).Value 'Stock name starts
                 NextTicker = ws.Cells(i + 1, 1) 'Next Tciker
            
                If PrvTicker <> Ticker Then
                    OpenValue = ws.Cells(i, 3).Value
                    TotStcVol = ws.Cells(i, 7).Value
            

    ws.Range("O2").Value = "Greates % Increase"
    ws.Range("O3").Value = "Greates % Decrease"
    ws.Range("O4").Value = "Greates Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Total Stock Value"
                
            ElseIf Ticker <> NextTicker Then
            
                    TotStcVol = ws.Cells(i, 7).Value + TotStcVol
                    ws.Cells(Printer, 12).Value = TotStcVol
                    ws.Cells(3, 16).Value = TotStcVol
                
                    ws.Cells(Printer, 9).Value = Ticker
                    CloseValue = ws.Cells(i, 6).Value
                    Change = CloseValue - OpenValue
                    ws.Cells(Printer, 10).Value = Change
                    
                    PChange = Change / OpenValue
                    ws.Cells(Printer, 11).Value = PChange
                    ws.Cells(Printer, 11).NumberFormat = "0.00%"
                    
                    
                Select Case Change
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                        
                End Select
                    j = j + 1
               
                
                    Printer = Printer + 1
                
                TotStcVol = 0
        
                Else
                    TotStcVol = ws.Cells(i, 7).Value + TotStcVol
                
                
            
            End If
        
             Next i
     Next ws
        
              

    MsgBox ("Well Done!!!")
End Sub

                    
