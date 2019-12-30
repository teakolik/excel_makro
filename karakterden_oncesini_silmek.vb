'Aşağıda "@" olarak belirtmiş olduğum 9. satırdaki karakteri arar. Bulduğu zaman solunda ne varsa siler. 

Sub Sil()
Set s = ActiveSheet
For i = 1 To 65000
Bak = s.Cells(i, 1)
For j = 1 To Len(s.Cells(i, 1))

If Mid(s.Cells(i, 1), j, 1) = "@" Then

Bak = Right(s.Cells(i, 1), Len(s.Cells(i, 1)) - j)
End If
Next
s.Cells(i, 1) = Bak
Next
End Sub
