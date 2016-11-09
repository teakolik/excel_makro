' Excel Sütun Karşılaştırma
' TEAkolik.com
' @TEAkolik

Sub TEAkolik()
Dim hcr As Range, j1 As Integer, j2 As Integer
On Error Resume Next
For i = 1 To [A65536].End(xlUp).Row
  x = Columns("B:B").Find(Cells(i, "A"), lookat:=xlWhole).Row
    If Err.Number <> 0 Then
       j1 = j1 + 1
       Cells(j1, "D").Value = Cells(i, "A")
      Err.Clear
        Else
          j2 = j2 + 1
          Cells(j2, "C").Value = Cells(i, "A")
    End If
Next
End Sub
