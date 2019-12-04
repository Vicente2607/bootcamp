Attribute VB_Name = "Módulo1"
Sub stock()
 Dim i As Double
 Dim j As Integer
 Dim ticker1 As String
 Dim ticker2 As String
 Dim ticker3 As String
 Dim precio1 As Double
 Dim precio2 As Double
 
 Dim dif_precio_anual As Double
 Dim dif_porcentual As Double
 Dim total_acciones As Double
 Dim lastrow As Double
 
 
 For Each ws In Worksheets
    ws.Select


j = 2
Range("i1") = "Ticker"
Range("j1") = "Yearly change"
Range("k1") = "Percent change"
Range("l1") = "Total stock volume"


Range("a1").Select
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


ticker = Range("a2").Value
precio1 = Range("c2").Value
total_acciones = Range("g2").Value




For i = 2 To lastrow
    If Cells(i, 1) = Cells(i + 1, 1) Then
            total_acciones = total_acciones + Cells(i + 1, 7)
    Else
            Cells(j, 9).Value = Cells(i, 1)
            Cells(j, 10).Value = Cells(i, 6) - precio1
            If Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                 Cells(j, 10).Interior.ColorIndex = 3
            End If
            If precio1 = 0 Then
                Cells(j, 11).Value = 0
            Else
            Cells(j, 11).Value = Format(Cells(j, 10) / precio1, "percent")
            End If
            Cells(j, 12).Value = total_acciones
            total_acciones = Cells(i + 1, 7)
            precio1 = Cells(i + 1, 4)
            j = j + 1
    End If
Next i
    
    Range("n2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("n4") = "Gratest Total Volume"
    Range("o1") = "Ticker"
    Range("p1") = "Value"
    
    
    Dim GI As Double
    Dim GD As Double
    Dim GV As Double
    
    
  lastrow = Cells(Rows.Count, 9).End(xlUp).Row
  GI = Range("k2").Value
  GD = GI
  GV = Range("l2").Value
  ticker1 = Range("i2")
  ticker2 = ticker1
  ticker3 = ticker1
  
  
  For i = 2 To lastrow
   If GI > Cells(i + 1, 11) Then
   Else
        GI = Cells(i + 1, 11)
        ticker1 = Cells(i, 9)
    End If
    
    If GD < Cells(i + 1, 11) Then
    Else
        GD = Cells(i + 1, 11)
        ticker1 = Cells(i, 9)
    End If
    
   If GV > Cells(i + 1, 12) Then
   Else
        GV = Cells(i + 1, 12)
        ticker3 = Cells(i, 9)
    End If
  Next i
  
    Range("o2").Value = ticker1
    Range("o3").Value = ticker2
    Range("o4").Value = ticker3
    
    Range("p2").Value = Format(GI, "percent")
   
    Range("p3").Value = Format(GD, "percent")
    Range("p4").Value = GV
    
Next ws
End Sub
