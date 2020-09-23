VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price History"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvPrice 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Item"
         Text            =   "Item"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.ComboBox cmbDay 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHistory.frx":0442
      Left            =   120
      List            =   "frmHistory.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iDay As Integer

Public Sub Start()
    Dim I As Integer
    Dim j As Integer
    For I = 1 To frmMain.Day
        cmbDay.AddItem "Day " & I
    Next
    cmbDay.ListIndex = cmbDay.ListCount - 1
End Sub

Private Sub cmbDay_Click()
    Dim I As Integer
    lvPrice.ListItems.Clear
    For I = 0 To 16
        iDay = cmbDay.ListIndex + 1
    Next
    For I = 1 To 13
        lvPrice.ListItems.Add , , Foods(I)
        lvPrice.ListItems(I).ListSubItems.Add , , History(iDay, I)
    Next
End Sub

Private Sub cmdClose_Click()
    Unload frmHistory
End Sub

