VERSION 5.00
Begin VB.Form frmBuy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buy"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBuy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar scr 
      Height          =   135
      Left            =   2160
      TabIndex        =   9
      Top             =   760
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "frmBuy.frx":0442
      Left            =   720
      List            =   "frmBuy.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Nevermind"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&Buy"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the amount you would like to buy"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   120
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   3420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   720
      X2              =   4080
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   720
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ounce(s)"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   120
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblInform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quality:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   120
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "frmBuy.frx":0446
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HowMuch As Long
Private Sub cmdBuy_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    j = 1
    'If cash < than the amount of money this will cost don't go ahead
    If Credit < frmMain.lstFoods.SelectedItem.ListSubItems(1).text * txtQty.text Then Exit Sub
    frmMain.lstItems.ListItems.Clear
    Credit = Credit - frmMain.lstFoods.SelectedItem.ListSubItems(1).text * txtQty.text 'Update cash remaining
    frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
    If Quantity(frmMain.lstFoods.SelectedItem.Index) > 0 Then 'if selected foods has been purchased before
        Quantity(frmMain.lstFoods.SelectedItem.Index) = Quantity(frmMain.lstFoods.SelectedItem.Index) + txtQty.text 'Add bought units onto already got units
        Else
        Quantity(frmMain.lstFoods.SelectedItem.Index) = txtQty.text 'Haven't been bought before so just add quantity bought
    End If
    If Avg(frmMain.lstFoods.SelectedItem.Index) > 0 Or Not Quantity(frmMain.lstFoods.SelectedItem.Index) = 0 Then  'If You have already bought the item or if you have stolen
        Avg(frmMain.lstFoods.SelectedItem.Index) = ((Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text) * (Avg(frmMain.lstFoods.SelectedItem.Index)) + txtQty.text * frmMain.lstFoods.SelectedItem.ListSubItems(1)) / (txtQty.text + (Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text)) 'Get Avererage.
        Else
        Avg(frmMain.lstFoods.SelectedItem.Index) = frmMain.lstFoods.SelectedItem.ListSubItems(1).text 'It hasn't been bought before so just set the average as the current price
    End If
    'Avg is a double variable because if the user bought
    '180 units at 40 cash and then bought one by one until they reached 20 at 50 cash
    'the average would remain 40 instead of 41 or 42 or something becuase  Long
    'Integers don't contain decimals and if the result returned was 40.25 it would
    'round it to 40 so no matter how many units you went up by one it would remain 40
    'unless you increased by a higher number in one turn like 30 or 40, because
    'even if there is still decimal places the number would change the average
    'by several whole numbers. With Double or Single variables if it recieved 40.25
    'it would keep the value 40.25 and below express it rounded off but actually keep
    'the decimal stored so if you did it again and it got 0.40 it would add on to
    'that 40.25 making it 40.65 and then rounding it off to 41 where if it was a long
    'it would have made 40.40 and rounded it off to 40 because the previous decimal
    'wasn't stored. I hope you understand anyway.
    For i = 1 To 13
        If Quantity(i) > 0 Then 'Add all foods from 1 to 17 that have been purchased (Quantity is how many purchased, if quantity = 0 none have been purchased)
            frmMain.lstItems.ListItems.Add j, , frmMain.lstFoods.ListItems(i)
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , "$" & Format(Round(Avg(i), 0), "###,###,###")
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , Quantity(i)
            j = j + 1 'Done, now goto next j
        End If
    Next
    iSpace = iSpace - txtQty.text 'iSpace left from 200 spaces (standard)
    Used = Used + txtQty.text 'How many spaces used
    frmMain.lblItems = "Items: " & Used & " of " & TotalSpace 'Update label
    PlaySound SDir & "cashreg.wav", 0, 3
    Unload frmBuy
End Sub

Private Sub cmdCancel_Click()
    Unload frmBuy
End Sub
Private Sub AddQualities()
If frmMain.lstFoods.SelectedItem = "Weed" Then
Quality(1) = "Chronic"
Quality(2) = "Dank"
Quality(3) = "Swag"
Quality(4) = "Weed?"
For q = 1 To 4
Combo1.AddItem Quality(q)
Next q
Exit Sub
End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Buy " & frmMain.lstFoods.SelectedItem
    AddQualities
    scr.Min = 1
    If Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text) > iSpace Then
        scr.Max = iSpace
        scr.Value = iSpace
        txtQty.SelLength = Len(txtQty)
        Exit Sub
    End If
    scr.Max = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    scr.Value = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    txtQty.text = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    txtQty.SelLength = Len(txtQty)
    Label3.Caption = "$" & Int(frmMain.lstFoods.SelectedItem.ListSubItems(1).text * scr.Value)
End Sub

Private Sub scr_Change()
    txtQty.text = scr.Value
    Label3.Caption = "$" & Int(frmMain.lstFoods.SelectedItem.ListSubItems(1).text * scr.Value)
End Sub

