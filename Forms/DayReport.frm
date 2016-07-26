VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DayReport 
   Caption         =   "End of Day Report"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   OleObjectBlob   =   "DayReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DayReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
Day.Caption = "Day " & Worksheets("LemonData").Cells(2, 5) 'captions all relevent data
Weather.Caption = Worksheets("LemonData").Cells(2, 11)
Demographic.Caption = Worksheets("LemonData").Cells(2, 14)
Activity.Caption = Worksheets("LemonData").Cells(2, 15)
Rent.Caption = "$" & Worksheets("LemonData").Cells(2, 16)
Temperature.Caption = Worksheets("LemonData").Cells(2, 12) & "c"
Location.Caption = Worksheets("LemonData").Cells(2, 13)

Dim RecD As Double 'Distance of used recuipie from optimal recipie
Dim WeatherD As Double ' Weather factor
Dim PriceD As Double ' Price factor


Dim CupsSold As Double
Dim IceTemp As Double ' Variable for optimal ice in drink
IceTemp = Worksheets("LemonData").Cells(2, 12) / 5 ' Optimal ice in drink is 1 ice cub for every 5 degrees above 0, if negative it induces less sales

If Worksheets("LemonData").Cells(2, 11) = "Sunny" Then 'increase sales by 50% if it is sunny
WeatherD = 1.5
End If

If Worksheets("LemonData").Cells(2, 11) = "Cloudy" Or Worksheets("LemonData").Cells(2, 13) = "Mall" Then 'sales are regular if it is cloudy or in a mall
WeatherD = 1
End If

If Worksheets("LemonData").Cells(2, 11) = "Rainy" Or Worksheets("LemonData").Cells(2, 11) = "Snowy" Then 'sales are halved if it is raining or snowing
WeatherD = 0.5
End If


If Worksheets("LemonData").Cells(2, 13) = "Neighborhood" Then
RecD = Abs((CInt(GameMain.LemonR) - 6) + (CInt(GameMain.SugarR) - 3) + (CInt(GameMain.IceR) - IceTemp)) 'absolute value difference between the recipie proposed by user and the optimal recipie for demographic
PriceD = CDbl(GameMain.PriceR) - 1.25 'finds the distance between optimal price and the price proposed by user
If PriceD <= 0 Then ' if the price is less than optimal give all sales like if it was optimal
PriceD = 0
End If
CupsSold = Round((30 - (2 ^ RecD) - (20 * PriceD)) * WeatherD, 0) 'cups sold formula
If CupsSold < 0 Then 'if you sell less than 0 then it is 0
CupsSold = 0
End If
End If


If Worksheets("LemonData").Cells(2, 13) = "Mall" Then 'same as above
RecD = Abs((CInt(GameMain.LemonR) - 9) + (CInt(GameMain.SugarR) - 3) + (CInt(GameMain.IceR) - 4))
PriceD = CDbl(GameMain.PriceR) - 1.75
If PriceD <= 0 Then
PriceD = 0
End If
CupsSold = Round((50 - (2 ^ RecD) - (30 * PriceD)) * WeatherD, 0)
If CupsSold < 0 Then
CupsSold = 0
End If
End If

If Worksheets("LemonData").Cells(2, 13) = "Park" Then  'same as above
RecD = Abs((CInt(GameMain.LemonR) - 6) + (CInt(GameMain.SugarR) - 6) + (CInt(GameMain.IceR) - IceTemp))
PriceD = CDbl(GameMain.PriceR) - 1.75
If PriceD <= 0 Then
PriceD = 0
End If
CupsSold = Round((60 - (2 ^ RecD) - (35 * PriceD)) * WeatherD, 0)
If CupsSold < 0 Then
CupsSold = 0
End If
End If

If Worksheets("LemonData").Cells(2, 13) = "Football Stadium" Then  'same as above
RecD = Abs((CInt(GameMain.LemonR) - 6) + (CInt(GameMain.SugarR) - 3) + (CInt(GameMain.IceR) - IceTemp))
PriceD = CDbl(GameMain.PriceR) - 2.5
If PriceD <= 0 Then
PriceD = 0
End If
CupsSold = Round((100 - (2 ^ RecD) - (55 * PriceD)) * WeatherD, 0)
If CupsSold < 0 Then
CupsSold = 0
End If
End If

Dim CupsUsed As Integer

If CupsUsed > Worksheets("LemonData").Cells(2, 9) Then 'If there are more cups used than there are cups then there are only as many cups used as there are cups
CupsUsed = Worksheets("LemonData").Cells(2, 9)
End If

Do While Worksheets("LemonData").Cells(2, 2) - CInt(GameMain.LemonR) >= 0 And Worksheets("LemonData").Cells(2, 3) - CInt(GameMain.SugarR) >= 0 And Worksheets("LemonData").Cells(2, 4) - (CInt(GameMain.IceR) * 12) >= 0
'Loops until the inventory is depleated of available resources
If CupsUsed >= CupsSold Then ' if the amount of cups used is more than the cups sold then there are as many cups used as cups sold
CupsUsed = CupsSold
Exit Do
End If
Worksheets("LemonData").Cells(2, 2) = Worksheets("LemonData").Cells(2, 2) - CInt(GameMain.LemonR)
Worksheets("LemonData").Cells(2, 3) = Worksheets("LemonData").Cells(2, 3) - CInt(GameMain.SugarR)
Worksheets("LemonData").Cells(2, 4) = Worksheets("LemonData").Cells(2, 4) - (CInt(GameMain.IceR) * 12)
CupsUsed = CupsUsed + 12 'adds 12 cups used after using the required lemons, sugar and ice to sell the product.
Loop

Worksheets("LemonData").Cells(2, 1) = Worksheets("LemonData").Cells(2, 1) - Worksheets("LemonData").Cells(2, 16) 'Cash - Rent
Worksheets("LemonData").Cells(2, 1) = Worksheets("LemonData").Cells(2, 1) + (CupsUsed * CDbl(GameMain.PriceR))
Worksheets("LemonData").Cells(2, 9) = Worksheets("LemonData").Cells(2, 9) - CupsUsed

CupsS.Caption = CupsUsed 'updates caption
PriceS.Caption = GameMain.PriceR
RevS.Caption = (CupsUsed * CDbl(GameMain.PriceR))
CashS.Caption = Worksheets("LemonData").Cells(2, 1)
Worksheets("LemonData").Cells(2, 17) = Worksheets("LemonData").Cells(2, 17) + CupsUsed ' Total of Career cups sold
Worksheets("LemonData").Cells(2, 18) = Worksheets("LemonData").Cells(2, 18) + (CupsUsed * CDbl(GameMain.PriceR)) 'total revenue over career
GameMain.CupsC.Caption = Worksheets("LemonData").Cells(2, 17)  'updates caption of career cups and revenue
GameMain.RevenueC.Caption = Worksheets("LemonData").Cells(2, 18)
End Sub
