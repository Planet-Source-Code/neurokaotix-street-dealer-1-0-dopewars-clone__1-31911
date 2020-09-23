VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Street Dealer"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Play"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   4365
      X2              =   4365
      Y1              =   2760
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4380
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   2760
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4380
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "frmName.frx":030A
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
Load frmMain
frmMain.Show
frmName.Hide
End Sub

