'Aşağıdaki makro belirli bir kriteri arar ve bulduğu zaman o satırı komple siler! TEAkolik

Option Compare Text
Sub KelimeBulveSatirSil()

    Dim son As Long, deg, i As Long, durum As Boolean, j As Integer

    son = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Aradığımız kelimeler buraya yazılacak
    deg = Array("*@gmail.com*", "*@hotmail.com*", "*@mynet.com*", "*@mail.com.tr*", "*@icloud.com*")

    Application.ScreenUpdating = False

    For i = son To 1 Step -1
        durum = False
        For j = 0 To UBound(deg)
            If Cells(i, "A") Like deg(j) Then durum = True
            If durum = True Then Exit For
        Next j
        If durum = True Then Rows(i).Delete Shift:=xlUp
    Next i

    Application.ScreenUpdating = True

End Sub
