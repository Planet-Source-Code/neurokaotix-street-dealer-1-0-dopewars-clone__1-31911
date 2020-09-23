VERSION 5.00
Begin VB.Form frmDoctor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.HScrollBar scr 
      Height          =   135
      Left            =   3480
      Min             =   1
      TabIndex        =   4
      Top             =   2040
      Value           =   1
      Width           =   855
   End
   Begin VB.Label lblPoints 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "How many doctor points do you wish to buy?"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3165
   End
   Begin VB.Label lblDoctor 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmDoctor
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If (txtQty.text * 10000) > Credit Then Exit Sub 'You cannot afford it
    If txtQty.text < 1 Then Exit Sub
    Credit = Credit - (txtQty.text * 10000)
    Health = Health + txtQty.text
    frmMain.pbHealth.Value = Health
    frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
    Unload frmDoctor
End Sub

Private Sub Form_Load()
    Dim MaxVal As Integer
    MaxVal = Int(Credit / 10000)
    If MaxVal > (100 - Health) Then MaxVal = (100 - Health)
    scr.Value = MaxVal
    scr.Max = MaxVal
    txtQty.text = MaxVal
    txtQty.SelLength = Len(txtQty.text)
    lblDoctor.Caption = "For each percentage of health you want to restore it will cost 10,000 money units. Your health is on " & Health & "% You would need " & CDbl(100 - Health) & " doctor points costing " & Format(CDbl(100 - Health) * 10000, "###,###,###") & " to restore your health. If each doctor point represents one percentage of your health you want to restore enter how many doctor points you want to buy"
End Sub

Private Sub scr_Change()
    txtQty.text = scr.Value
End Sub
