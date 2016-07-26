VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameMain 
   Caption         =   "Lemonade Empire"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   OleObjectBlob   =   "GameMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub IceR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'restricts input to only numeric
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub LemonR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'restricts input to only numeric
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub Location_Change() 'changes the data based on location change in dropdown box
If Location = "Neighborhood" Then
Worksheets("LemonData").Cells(2, 13) = Location
Worksheets("LemonData").Cells(2, 14) = "Neighbors"
Worksheets("LemonData").Cells(2, 15) = "Low"
Worksheets("LemonData").Cells(2, 16) = 0
Call RefreshLocation
End If

If Location = "Mall" Then
Worksheets("LemonData").Cells(2, 13) = Location
Worksheets("LemonData").Cells(2, 14) = "Seniors"
Worksheets("LemonData").Cells(2, 15) = "Medium"
Worksheets("LemonData").Cells(2, 16) = 30
Call RefreshLocation
End If

If Location = "Park" Then
Worksheets("LemonData").Cells(2, 13) = Location
Worksheets("LemonData").Cells(2, 14) = "Kids"
Worksheets("LemonData").Cells(2, 15) = "High"
Worksheets("LemonData").Cells(2, 16) = 10
Call RefreshLocation
End If

If Location = "Football Stadium" Then
Worksheets("LemonData").Cells(2, 13) = Location
Worksheets("LemonData").Cells(2, 14) = "Adults"
Worksheets("LemonData").Cells(2, 15) = "Very High"
Worksheets("LemonData").Cells(2, 16) = 40
Call RefreshLocation
End If
End Sub

Sub RefreshLocation() 'refreshes the location data on the userform
Weather.Caption = Worksheets("LemonData").Cells(2, 11)
Demographic.Caption = Worksheets("LemonData").Cells(2, 14)
Activity.Caption = Worksheets("LemonData").Cells(2, 15)
Rent.Caption = "$" & Worksheets("LemonData").Cells(2, 16)
End Sub



Private Sub PriceR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'allows a "." in price but does not allow two, numeric inputs only for price
If (KeyAscii = 46) Then
    If InStr(PriceR, ".") <> 0 Then
    KeyAscii = 0
    Else
    KeyAscii = KeyAscii
    End If
End If

If (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub SaveBut_Click()
ActiveWorkbook.Save
End Sub

Private Sub Start_Click()
If LemonR = "" Or SugarR = "" Or IceR = "" Or PriceR = "" Then 'error checks if fields left blank
MsgBox ("Some parts of the recipe are empty")
Exit Sub
End If

If Location = "" Then 'error checks if fields left blank
MsgBox ("The location is empty")
Exit Sub
End If

If Worksheets("LemonData").Cells(2, 16) > Worksheets("LemonData").Cells(2, 1) Then 'if rent is more than cash on hand then cannot use location
MsgBox ("There is not enough cash to pay for rent")
Exit Sub
End If


DayReport.Show 'generates the simulation
Worksheets("LemonData").Cells(2, 5) = CInt(Worksheets("LemonData").Cells(2, 5).Value) + 1 'afterwards sets the day to the next day

 Worksheets("LemonData").Cells(2, 6) = Round(0.4 + (Int((2 - -2 + 1) * Rnd + -2) / 10), 2) 'reprices store
 Worksheets("LemonData").Cells(2, 7) = Round(0.4 + (Int((2 - -2 + 1) * Rnd + -2) / 10), 2)
 Worksheets("LemonData").Cells(2, 8) = Round(1 + (Int((5 - -5 + 1) * Rnd + -5) / 10), 2)
 Worksheets("LemonData").Cells(2, 10) = Round(1 + (Int((5 - -5 + 1) * Rnd + -5) / 10), 2)

Call weathercall 'gets a new weather call
Call TempCall ' gets a new temperature call

Dim BillyRNG As Integer
Dim BillyRNGQ As Integer
BillyRNG = Int((5 - 1 + 1) * Rnd + 1)

If BillyRNG = 1 Then 'Billy RNG event causes a loss of inventory to prevent hoarding on low prices
BillyRNGQ = Int((4 - 1 + 1) * Rnd + 1)
    If BillyRNGQ = 1 Then
    MsgBox ("Billy from Iced Tea Tycoon stole all your lemons!")
    Worksheets("LemonData").Cells(2, 2) = 40
    End If
    If BillyRNGQ = 2 Then
    MsgBox ("Billy from Iced Tea Tycoon stole all your sugar!")
    Worksheets("LemonData").Cells(2, 3) = 0
    End If
    If BillyRNGQ = 3 Then
    MsgBox ("Billy from Iced Tea Tycoon melted all your ice!")
    Worksheets("LemonData").Cells(2, 4) = 0
    End If
    If BillyRNGQ = 4 Then
    MsgBox ("Billy from Iced Tea Tycoon stole all cups!")
    Worksheets("LemonData").Cells(2, 9) = 0
    End If
End If
Call RefreshData 'refreshes the data
End Sub

Private Sub StoreButton_Click() 'brings up the store
Store.Show
End Sub




Private Sub SugarR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' only allows numeric inputs in the recipie field
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub UserForm_Initialize()
Call RefreshData 'refreshes all the data to bring the form back up to present data
Location.AddItem "Neighborhood" 'adds the locations to the dropdown box
Location.AddItem "Mall"
Location.AddItem "Park"
Location.AddItem "Football Stadium"

End Sub

Public Sub RefreshData() 'Refreshes all the main data on the page based on saved lemon data
CashI.Caption = "$" & Worksheets("LemonData").Cells(2, 1)
LemonI.Caption = Worksheets("LemonData").Cells(2, 2)
SugarI.Caption = Worksheets("LemonData").Cells(2, 3)
IceI.Caption = Worksheets("LemonData").Cells(2, 4)
DayI.Caption = Worksheets("LemonData").Cells(2, 5)
CupI.Caption = Worksheets("LemonData").Cells(2, 9)
Weather.Caption = Worksheets("LemonData").Cells(2, 11)
Temperature.Caption = Worksheets("LemonData").Cells(2, 12) & "c"
GameMain.CupsC.Caption = Worksheets("LemonData").Cells(2, 17)
GameMain.RevenueC.Caption = Worksheets("LemonData").Cells(2, 18)
End Sub
Public Sub TempCall() 'same temp call as before
Dim RNGNum As Integer
RNGNum = Int((300 - -300 + 1) * Rnd + -300)
Worksheets("LemonData").Cells(2, 12).Formula = RNGNum / 10
End Sub

Public Sub weathercall() 'same weather call as before
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

