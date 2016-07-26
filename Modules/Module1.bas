Attribute VB_Name = "Module1"
Sub TempCall()
Dim RNGNum As Integer
RNGNum = Int((300 - -300 + 1) * Rnd + -300)
Worksheets("LemonData").Cells(2, 12).Formula = RNGNum / 10
End Sub


Sub weathercall1()

Dim RNGNum As Integer

RNGNum = Int((5 - 1 + 1) * Rnd + 1)

If RNGNum = 1 Or RNGNum = 2 Then
Worksheets("LemonData").Cells(2, 11) = "Sunny"
End If

If RNGNum = 3 Or RNGNum = 4 Then
Worksheets("LemonData").Cells(2, 11) = "Cloudy"
End If

If RNGNum = 5 Then
    If Worksheets("LemonData").Cells(2, 12) > 0 Then
    Worksheets("LemonData").Cells(2, 11) = "Rainy"
    Else
    Worksheets("LemonData").Cells(2, 11) = "Snowy"
    End If
End If

End Sub


Sub maineee()
Call TempCall
Call weathercall1
End Sub
