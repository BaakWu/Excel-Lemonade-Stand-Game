VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Intro 
   Caption         =   "Lemon Intro"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8940
   OleObjectBlob   =   "Intro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub LoadSave_Click()
Unload Me
GameMain.Show
End Sub

Private Sub NewGame_Click()
Message = MsgBox("Are you sure you want to start a new game? All data from previous game will be lost.", vbYesNo) 'comfirms to user if he really wants to clear all data

If Message = vbYes Then 'clears all data

Worksheets("LemonData").Cells(2, 1) = 40
Worksheets("LemonData").Cells(2, 2) = 0
Worksheets("LemonData").Cells(2, 3) = 0
Worksheets("LemonData").Cells(2, 4) = 0
Worksheets("LemonData").Cells(2, 9) = 0
Worksheets("LemonData").Cells(2, 5) = 1
Worksheets("LemonData").Cells(2, 17) = 0
Worksheets("LemonData").Cells(2, 18) = 0


 Worksheets("LemonData").Cells(2, 6) = 0.4 'reprices store
 Worksheets("LemonData").Cells(2, 7) = 0.4
 Worksheets("LemonData").Cells(2, 8) = 1
 Worksheets("LemonData").Cells(2, 10) = 1
Call TempCall
Call weathercall
Unload Me
GameMain.Show

End If

End Sub


Public Sub TempCall() 'creates a random number for temeprature between 30.0 to -30.0 c
Dim RNGNum As Integer
RNGNum = Int((300 - -300 + 1) * Rnd + -300)
Worksheets("LemonData").Cells(2, 12).Formula = RNGNum / 10
End Sub

Public Sub weathercall() 'randomly generates weather 2/5 chance for sunny, 2/5 chance for cloudy and 1/5 chance for rain or snow depending on the temperature
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

Private Sub UserForm_Click()

End Sub
