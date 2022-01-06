Sub Stock()
  Dim i, j, n, resultcount, subtotal, Results, Closeprice, Openprice, top, bottom As Long
  Dim WS As Worksheet
  
  'run through each worksheet
    For Each WS In ThisWorkbook.Worksheets

  Results = 2
  n = WS.Cells(Rows.Count, 1).End(xlUp).Row
  subtotal = 0
  
  WS.Cells(1, 9).Value = "Ticker"
  WS.Cells(1, 10).Value = "Chg"
  WS.Cells(1, 11).Value = "% Chg"
  WS.Cells(1, 12).Value = "Volume"
  WS.Cells(1, 14).Value = "Close Price"
  WS.Cells(1, 15).Value = "Open Price"


    For i = 2 To WS.Cells(Rows.Count, 1).End(xlUp).Row
        ' add to the total volume
        subtotal = subtotal + WS.Cells(i, 7).Value
        ' iterate on each stock
        If LCase(WS.Cells(i + 1, 1).Value) <> LCase(WS.Cells(i, 1).Value) Then
        'Look for Opening Pricepoint
        top = WS.Range("A:A").Find(What:=WS.Cells(i, 1).Value).Row
        opened = WS.Cells(top, 3).Value
        'Get Close Price point
        'bottom = WS.Range("A:A").Find(What:=Ws.cells(i, 1).Value, SearchDirection:=xlPrevious).Row
        closed = WS.Cells(i, 6).Value
       
            ' store the name of the ticker being finished
            WS.Cells(Results, 9).Value = LCase(WS.Cells(i, 1))
            'Store Open/Close Pricepoint
            WS.Cells(Results, 15).Value = opened
            WS.Cells(Results, 14).Value = closed
            
            'Calcuate Change
            chg = closed - opened
            WS.Cells(Results, 10) = chg
                 If opened <> 0 Then
                        Pctchg = closed / opened - 1
                 Else
                        Pctchg = "0"
                End If
            
                WS.Cells(Results, 11) = Pctchg
            
            
             'chg Coloring
                 If Pctchg > 0 Then
                        WS.Cells(Results, 10).Interior.ColorIndex = 4
                Else
                        WS.Cells(Results, 10).Interior.ColorIndex = 3
                End If
            ' store the volume total
            WS.Cells(Results, 12).Value = subtotal
            ' increment results
            Results = Results + 1
        
            
            ' reset the total
            subtotal = 0
        End If
      
    Next i

    WS.Range("K:K").NumberFormat = "0.00%"
    WS.Range("L:L").NumberFormat = "#,##0"
 Next WS
 
 End Sub
