VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFinances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Finances"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmFinances.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtDebit 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin MSComctlLib.ListView lstLoan 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Amount"
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Interest"
         Text            =   "Interest"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.TextBox txtCash 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "&Pay"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoan 
      Caption         =   "&Loan"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblLoan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Loan"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   435
   End
   Begin VB.Label lblCash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "frmFinances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoan_Click()
    On Error Resume Next
    Dim lInt As Double
    lInt = Int(InputBox("How much money would you like to borrow?", "Bank Loan", 1000))
    If lInt < 1 Then Exit Sub
    If lInt + Debit > 500000 Then MsgBox "You cannot loan more than 500,000.": Exit Sub
    Debit = Debit + lInt
    lstLoan.ListItems.Clear
    lstLoan.ListItems.Add , , Format(Debit, "###,###,###")
    lstLoan.ListItems(1).ListSubItems.Add , , "2%"
    Credit = Credit + lInt
    txtCash.text = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    txtDebit.text = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
    cmdExit.SetFocus
End Sub

Private Sub cmdPay_Click()
    On Error Resume Next
    Dim lInt As Double
    Dim Pay As Double
    If Debit > Credit Then Pay = Credit
    If Debit < Credit Then Pay = Debit
    lInt = Int(InputBox("How much of the loan would you like to pay off?", "Bank Loan", Pay))
    If lInt < 1 Then Exit Sub
    If lInt > Debit Then Exit Sub
    Debit = Debit - lInt
    Credit = Credit - lInt
    lstLoan.ListItems.Clear
    If Debit <> 0 Then
        lstLoan.ListItems.Add , , Format(Debit, "###,###,###")
        lstLoan.ListItems(1).ListSubItems.Add , , "2%"
        lstLoan.ListItems(1).ListSubItems.Add , , Format(lInt, "###,###,###")
    End If
    txtCash.text = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    txtDebit.text = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
    cmdExit.SetFocus
End Sub

Private Sub Form_Load()
    txtCash.text = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    txtDebit.text = IIf(Debit <> 0, Format(Debit, "###,###,###"), 0)
    If Debit > 0 Then
        lstLoan.ListItems.Add , , Format(Debit, "###,###,###")
        lstLoan.ListItems(1).ListSubItems.Add , , "2%"
    End If
End Sub
