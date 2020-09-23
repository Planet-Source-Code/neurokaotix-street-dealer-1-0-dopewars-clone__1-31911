Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Public User As String 'Name when they login to Food Wars
Public History(1 To 365, 1 To 13)
Public Foods(1 To 13) As String 'Foods available to buy
Public Prices(1 To 13) As Long 'Price of foods
Public Qty(1 To 13) As Integer 'Quantity of items purchased
Public Quantity(1 To 13) As Integer 'Quantity of items purchased
Public Weapon(1 To 8) As String 'Available weapons
Public YW(1 To 8) As Boolean 'The weapons you have
Public Credit As Double 'your money
Public Quality(1 To 4) As String 'your money
Public Debit As Long 'no need for double cant go past 4,000,000
Public Avg(1 To 13) As Double 'Avg of bought prices
Public iSpace As Integer 'iSpace remaining
Public Used As Integer 'Used truck spaces
Public TotalSpace As Integer 'Total truck space
Public Caught As Byte 'how many times caught
Public Stole As Boolean 'If items stolen before
Public SDir As String 'Directory with sound files
Public Truck As String 'Directory with Truck wav
Public Sold As Boolean 'Whether you have traded
Public Attacked As Boolean 'Whether you have attacked anyone with a weapon
Public Health As Integer
Public HSL As New HighScoreList
Public FinalScore As Double 'Final score.
Dim conitm As ListItem

Public Sub AddWeapons()
    Weapon(1) = "Brass Knuckles"
    Weapon(2) = "Switch Blade"
    Weapon(3) = ".22 Marlin"
    Weapon(4) = "TEK 9mm"
    Weapon(5) = "Beretta 92F"
    Weapon(6) = "MAC 10"
    Weapon(7) = "9mm UZI"
    Weapon(8) = "AR15"
End Sub

Public Sub AddFoods()
    Dim i As Byte 'Counter
    Foods(1) = "Cocaine"
    Foods(2) = "Heroine"
    Foods(3) = "Weed"
    Foods(4) = "Crystal Meth"
    Foods(5) = "PCP"
    Foods(6) = "Ecstacy"
    Foods(7) = "Crack"
    Foods(8) = "Acid"
    Foods(9) = "Mushrooms"
    Foods(10) = "Special K"
    Foods(11) = "Hashish"
    Foods(12) = "MDA"
    Foods(13) = "NO2"
    For i = 1 To 13
        frmMain.lstFoods.ListItems.Add , , Foods(i)
    Next
    
End Sub

Public Sub AddPrices()
    Dim Temp(1 To 13) As Integer 'Store actual prices before randomized
    Dim iRnd As Byte
    Dim j As Integer
    Dim i As Byte 'Counter
    Dim Temp2(1 To 13) As Integer 'Store actual quantities
    Dim i2 As Byte 'Counter
    Dim i3 As Byte 'Counter
    Randomize
    Temp(1) = 20000
    Temp(2) = 8000
    Temp(3) = 40
    Temp(4) = 500
    Temp(5) = 150
    Temp(6) = 20
    Temp(7) = 2600
    Temp(8) = 2800
    Temp(9) = 1100
    Temp(10) = 400
    Temp(11) = 600
    Temp(12) = 3000
    Temp(13) = 100
    For i = 1 To 13
        Do
            Prices(i) = Int((Temp(i) / 0.7) * Rnd) + 1
        Loop Until Prices(i) > Temp(i) - Temp(i) * 0.3
    Next
    
    Randomize
    Temp2(1) = 70
    Temp2(2) = 100
    Temp2(3) = 200
    Temp2(4) = 150
    Temp2(5) = 150
    Temp2(6) = 300
    Temp2(7) = 120
    Temp2(8) = 110
    Temp2(9) = 140
    Temp2(10) = 180
    Temp2(11) = 160
    Temp2(12) = 120
    Temp2(13) = 200
    For i2 = 1 To 13
        Do
            Qty(i2) = Int((Temp2(i2) / 0.9) * Rnd) + 1
        Loop Until Qty(i2) > Temp2(i2) - Temp2(i2) * 0.3
    Next


    iRnd = Int(24 * Rnd) + 1
    If iRnd = 2 Then
        j = Int(Rnd * 13) + 1
        Prices(j) = Prices(j) * 5
        frmMain.lblNews.Caption = ""
        frmMain.lblNews.Caption = "There are shortages of " & Foods(j) & " and the prices have gone up!"
    End If
    If iRnd = 4 Then
        j = Int(Rnd * 13) + 1
        Prices(j) = Prices(j) / 5
        frmMain.lblNews.Caption = ""
        frmMain.lblNews.Caption = "Prices on " & Foods(j) & " have gone down!"
    End If
    For i = 1 To 13
        frmMain.lstFoods.ListItems(i).ListSubItems.Clear
    Next

    For i = 1 To 13
        History(frmMain.Day, i) = Prices(i)
        frmMain.lstFoods.ListItems(i).ListSubItems.Add , , "$" & Format(Prices(i), "###,###,###")
        frmMain.lstFoods.ListItems(i).ListSubItems.Add , , Qty(i)
        If Prices(i) > Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(1).Index
        End If
        If Prices(i) < Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(2).Index
        End If
        If Prices(i) = Temp(i) Then
            frmMain.lstFoods.ListItems(i).SmallIcon = frmMain.imgList.ListImages(3).Index
        End If
    Next

End Sub

