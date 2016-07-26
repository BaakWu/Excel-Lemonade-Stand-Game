VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Store 
   Caption         =   "Not Loblaws"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   OleObjectBlob   =   "Store.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub LemonB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'only allows numeric inputs
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub SugarB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'only allows numeric inputs
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub IceB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'only allows numeric inputs
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub BuyB_Click()
Dim Calculation As Double

If LemonB = "" Then 'assumes you want 0 if left blank field
LemonB.Text = 0
End If

If SugarB = "" Then 'assumes you want 0 if left blank field
SugarB.Text = 0
End If

If IceB = "" Then 'assumes you want 0 if left blank field
IceB.Text = 0
End If

If CupB = "" Then 'assumes you want 0 if left blank field
CupB.Text = 0
End If

Calculation = Worksheets("LemonData").Cells(2, 1) - ((CInt(LemonB) * Worksheets("LemonData").Cells(2, 6)) + (CInt(SugarB) * Worksheets("LemonData").Cells(2, 7)) + (CInt(IceB) * Worksheets("LemonData").Cells(2, 8)) + (CInt(CupB) * Worksheets("LemonData").Cells(2, 10)))
' Money calculation taking the sum cost of all products against your cash reservers.
If Calculation < 0 Then 'says insufficient funds if there is not enough money
MsgBox ("insufficient funds")
Exit Sub
End If


Worksheets("LemonData").Cells(2, 2) = Worksheets("LemonData").Cells(2, 2) + CInt(LemonB) 'adds lemons
Worksheets("LemonData").Cells(2, 3) = Worksheets("LemonData").Cells(2, 3) + CInt(SugarB) 'adds sugar
Worksheets("LemonData").Cells(2, 4) = Worksheets("LemonData").Cells(2, 4) + CInt(IceB) * 50 'adds ice
Worksheets("LemonData").Cells(2, 9) = Worksheets("LemonData").Cells(2, 9) + CInt(CupB) * 100 'adds cups
Worksheets("LemonData").Cells(2, 1) = Round(Worksheets("LemonData").Cells(2, 1) - ((CInt(LemonB) * Worksheets("LemonData").Cells(2, 6)) + (CInt(SugarB) * Worksheets("LemonData").Cells(2, 7)) + (CInt(IceB) * Worksheets("LemonData").Cells(2, 8)) + (CInt(CupB) * Worksheets("LemonData").Cells(2, 10))), 2)
'reduces cash by cost of goods

LemonB.Text = "" 'clears the fields for another input
SugarB.Text = ""
IceB.Text = ""
CupB.Text = ""
Call RefreshData 'refreshes the data for the main page, putting it up to date

GameMain.CashI.Caption = "$" & Worksheets("LemonData").Cells(2, 1) 'refreshes the captions putting them up to date after purchase
GameMain.LemonI.Caption = Worksheets("LemonData").Cells(2, 2)
GameMain.SugarI.Caption = Worksheets("LemonData").Cells(2, 3)
GameMain.IceI.Caption = Worksheets("LemonData").Cells(2, 4)
GameMain.CupI.Caption = Worksheets("LemonData").Cells(2, 9)

End Sub




Private Sub CupB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'only allows numeric inputs
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub UserForm_Initialize()
Call RefreshData 'brings data up to date
LemonP.Caption = "$" & Worksheets("LemonData").Cells(2, 6) 'price caption for all prices in the store
SugarP.Caption = "$" & Worksheets("LemonData").Cells(2, 7)
IceP.Caption = "$" & Worksheets("LemonData").Cells(2, 8)
CupP.Caption = "$" & Worksheets("LemonData").Cells(2, 10)

Me.LemonB.MaxLength = 4 'maximum amount to buy is 9999 lemons otherwise overflow occurs
Me.SugarB.MaxLength = 4 'maximum amount to buy is 9999 lemons otherwise overflow occurs
Me.IceB.MaxLength = 4 'maximum amount to buy is 9999 lemons otherwise overflow occurs
Me.CupB.MaxLength = 4 'maximum amount to buy is 9999 lemons otherwise overflow occurs
End Sub

Public Sub RefreshData() 'refreshes data for up to date inventory
CashI.Caption = "$" & Worksheets("LemonData").Cells(2, 1)
LemonI.Caption = Worksheets("LemonData").Cells(2, 2)
SugarI.Caption = Worksheets("LemonData").Cells(2, 3)
IceI.Caption = Worksheets("LemonData").Cells(2, 4)
CupI.Caption = Worksheets("LemonData").Cells(2, 9)

End Sub

