VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002C1E19&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Street Dealer - Buy low, sell high... or steal"
   ClientHeight    =   6240
   ClientLeft      =   615
   ClientTop       =   615
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   600
      ScaleWidth      =   8040
      TabIndex        =   38
      Top             =   2280
      Width           =   8040
      Begin VB.Label lblNews 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing really worth mentioning..."
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   7785
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Word on the streets:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   1800
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Saved Games "".sav|.SAV"""
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   1320
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   4800
      Picture         =   "frmMain.frx":0E7C
      ScaleHeight     =   2040
      ScaleWidth      =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3360
      Begin VB.CommandButton cmdColes 
         Caption         =   "&Compton"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdMaxi 
         Caption         =   "&The Valley"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSafeway 
         Caption         =   "&Inglewood"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdIGA 
         Caption         =   "&Lakewood"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdFranklins 
         Caption         =   "&San Pedro"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdBILO 
         Caption         =   "&Long Beach"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compton"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   570
         TabIndex        =   31
         Top             =   1070
         Width           =   630
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "San Pedro"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   480
         TabIndex        =   30
         Top             =   1590
         Width           =   810
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Valley"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   2025
         TabIndex        =   29
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lakewood"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   2100
         TabIndex        =   28
         Top             =   1070
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Long Beach"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   2040
         TabIndex        =   27
         Top             =   1590
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inglewood"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   480
         TabIndex        =   26
         Top             =   510
         Width           =   810
      End
      Begin VB.Image imgCity6 
         Height          =   405
         Left            =   1725
         Picture         =   "frmMain.frx":1E18
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Image imgCity3 
         Height          =   405
         Left            =   135
         Picture         =   "frmMain.frx":21A7
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Image imgCity5 
         Height          =   405
         Left            =   1725
         Picture         =   "frmMain.frx":2536
         Top             =   900
         Width           =   1500
      End
      Begin VB.Image imgCity2 
         Height          =   405
         Left            =   120
         Picture         =   "frmMain.frx":28C5
         Top             =   900
         Width           =   1500
      End
      Begin VB.Image imgCity4 
         Height          =   405
         Left            =   1725
         Picture         =   "frmMain.frx":2C54
         Top             =   360
         Width           =   1500
      End
      Begin VB.Image imgCity1 
         Height          =   405
         Left            =   135
         Picture         =   "frmMain.frx":2FE3
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblPlace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You are now in:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1350
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3372
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3427
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":347D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5282
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A41
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6254
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":697E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      Picture         =   "frmMain.frx":7139
      ScaleHeight     =   2040
      ScaleWidth      =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3360
      Begin StreetDealer.ProgBar pbHealth 
         Height          =   165
         Left            =   960
         Top             =   1680
         Width           =   2265
         _extentx        =   3995
         _extenty        =   291
         backcolor       =   255
         barcolor        =   65280
         borderstyle     =   0
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   920
         TabIndex        =   10
         Top             =   600
         Width           =   2310
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   920
         TabIndex        =   9
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label dspBank 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   120
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Day:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label lblCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   920
         TabIndex        =   5
         Top             =   120
         Width           =   2325
      End
      Begin VB.Label dspHealth 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Health:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label dspCash 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   120
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   450
      End
      Begin VB.Label dspDebit 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Debt:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   450
      End
   End
   Begin MSComctlLib.ListView lstFoods 
      Height          =   2910
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   5133
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   16777215
      BackColor       =   2891289
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Food"
         Text            =   "Item"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qty."
         Object.Width           =   1288
      EndProperty
      Picture         =   "frmMain.frx":80D5
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   12
      Top             =   6030
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   370
      Style           =   1
      SimpleText      =   "Name"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstItems 
      Height          =   2910
      Left            =   4800
      TabIndex        =   14
      Top             =   3000
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   5133
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   2891289
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Food"
         Text            =   "Item"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Qty"
         Text            =   "Qty."
         Object.Width           =   1288
      EndProperty
      Picture         =   "frmMain.frx":9587
   End
   Begin VB.Image imgDown 
      Height          =   405
      Left            =   5640
      Picture         =   "frmMain.frx":AA39
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Image imgUp 
      Height          =   405
      Left            =   5160
      Picture         =   "frmMain.frx":ADC8
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Image imgOrigDwn 
      Height          =   405
      Left            =   5400
      Picture         =   "frmMain.frx":B157
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Image imgOrigUp 
      Height          =   405
      Left            =   5280
      Picture         =   "frmMain.frx":B49C
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<< Drop"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3840
      TabIndex        =   24
      Top             =   4680
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3885
      TabIndex        =   23
      Top             =   5685
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clinic"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3840
      TabIndex        =   22
      Top             =   5175
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steal >>"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3795
      TabIndex        =   21
      Top             =   4170
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<< Sell"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3825
      TabIndex        =   20
      Top             =   3660
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buy >>"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3855
      TabIndex        =   19
      Top             =   3165
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3960
      TabIndex        =   18
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgBank 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":B7E1
      Top             =   1760
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3675
      TabIndex        =   17
      Top             =   1375
      Width           =   900
   End
   Begin VB.Image imgCalc 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":BB26
      Top             =   1212
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Hist."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3645
      TabIndex        =   16
      Top             =   830
      Width           =   990
   End
   Begin VB.Image imgPriceHist 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":BE6B
      Top             =   666
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loans"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   3900
      TabIndex        =   15
      Top             =   285
      Width           =   450
   End
   Begin VB.Image imgLoans 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":C1B0
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgBuy 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":C4F5
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Image imgSell 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":C83A
      Top             =   3510
      Width           =   1140
   End
   Begin VB.Image imgSteal 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":CB7F
      Top             =   4005
      Width           =   1140
   End
   Begin VB.Image imgClinic 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":CEC4
      Top             =   5010
      Width           =   1140
   End
   Begin VB.Image imgStore 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":D209
      Top             =   5520
      Width           =   1140
   End
   Begin VB.Image imgDrop 
      Height          =   405
      Left            =   3555
      Picture         =   "frmMain.frx":D54E
      Top             =   4515
      Width           =   1140
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Progress"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Game"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit Game"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFinances 
         Caption         =   "&Finances"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewScores 
         Caption         =   "&High Scores"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "&Price History"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSounds 
         Caption         =   "&Enable Sounds"
         Begin VB.Menu mnuTruck 
            Caption         =   "&Truck sound"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOther 
            Caption         =   "&Other sounds"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const BM_SETSTYLE = &HF4
Private Const BS_SOLID = 0
Public Day As Integer
Dim LoadedData As String

Private Sub cmdBank_Click()

End Sub

Private Sub cmdBilo_Click()
If cmdBILO.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 365 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Long Beach"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdBILO.Enabled = False
    EnableOthers cmdColes, cmdSafeway, cmdIGA, cmdFranklins, cmdMaxi
    Call AddPrices
End Sub


Private Sub cmdColes_Click()
If cmdColes.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 365 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Compton"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdColes.Enabled = False
    Call EnableOthers(cmdFranklins, cmdIGA, cmdMaxi, cmdSafeway, cmdBILO)
    Call AddPrices
End Sub

Private Sub cmdDrop_Click()

End Sub

Private Sub cmdFinances_Click()

End Sub

Private Sub cmdFranklins_Click()
If cmdFranklins.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 365 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in San Pedro"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdFranklins.Enabled = False
    EnableOthers cmdIGA, cmdMaxi, cmdColes, cmdSafeway, cmdBILO
    Call AddPrices
End Sub

Private Sub cmdHistory_Click()

End Sub

Private Sub cmdIGA_click()
If cmdIGA.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 35 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Lakewood"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdIGA.Enabled = False
    EnableOthers cmdFranklins, cmdColes, cmdMaxi, cmdBILO, cmdSafeway
    Call AddPrices
End Sub

Private Sub cmdMaxi_Click()
If cmdMaxi.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 365 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    frmMain.lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in The Valley"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdMaxi.Enabled = False
    EnableOthers cmdBILO, cmdColes, cmdFranklins, cmdIGA, cmdSafeway
    Call AddPrices
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSafeway_Click()
If cmdSafeway.Enabled = False Then MsgBox "You're already there dumbass!", vbInformation: Exit Sub
    Dim Exec As Boolean
    Exec = CheckScore
    If Exec = False Then Exit Sub
    If Day = 365 Then Exit Sub
    Day = Day + 1
    Debit = Debit + (Debit * 0.2)
    lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblPlace.Caption = "You are now in Inglewood"
    lblDay.Caption = "Day: " & Day & " of 365"
    cmdSafeway.Enabled = False
    EnableOthers cmdBILO, cmdColes, cmdFranklins, cmdIGA, cmdMaxi
    Call AddPrices
End Sub

Private Sub cmdSteal_Click()
    If iSpace > 0 Then frmSteal.Show vbModal
End Sub

Private Sub cmdSell_Click()

End Sub



Private Sub Form_Load()
StatusBar1.SimpleText = frmName.Text1.text
NewStyle
    If App.PrevInstance = True Then End
    If Right(App.Path, 1) = "\" Then
        SDir = App.Path & "Sounds\"
        Else
        SDir = App.Path & "\Sounds\"
    End If
    Truck = SDir & "Truck.wav"
    CButton cmdColes
    CButton cmdMaxi
    CButton cmdSafeway
    CButton cmdIGA
    CButton cmdFranklins
    CButton cmdBILO
    Call Init
End Sub

Public Sub Init()
    Dim Rand As Byte
    Randomize
    Sound
    EnableControls
    ResetCaptions
    YW(1) = False
    YW(2) = False
    YW(3) = False
    YW(4) = False
    Call HSL.SetListCount(20)
    Call HSL.FileName(HSL.DefaultFileName)
    FinalScore = 0
    Credit = 1000
    Debit = 3000
    Day = 1
    iSpace = 200
    Used = 0
    TotalSpace = 200
    Caught = 0
    Health = 100
    Stole = False
    Attacked = False
    Sold = False
    cmdColes.Enabled = False
    EnableOthers cmdFranklins, cmdBILO, cmdIGA, cmdSafeway, cmdMaxi
    lblDay.Caption = "Day: " & Day & " of 365"
    lblPlace.Caption = "You are now in Compton"
    lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
    lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
    lblItems.Caption = "Items: 0 of 200"
    pbHealth.Value = Health
    lstFoods.ListItems.Clear
    lstItems.ListItems.Clear
    Erase Avg
    Erase Quantity
    Call AddFoods
    Call AddPrices
    Call AddWeapons
End Sub


Private Sub imgLoans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLoans.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgLoans_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLoans.Picture = imgOrigUp.Picture
    frmFinances.Show vbModal
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLoans.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLoans.Picture = imgOrigUp.Picture
    frmFinances.Show vbModal
End Sub

Private Sub imgPriceHist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPriceHist.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgPriceHist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPriceHist.Picture = imgOrigUp.Picture
    frmHistory.Start
    frmHistory.Show vbModal
End Sub

Private Sub imgCity2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity2.Picture = imgDown.Picture
End Sub
Private Sub imgCity2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity2.Picture = imgUp.Picture
cmdColes_Click
End Sub

Private Sub imgCity3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity3.Picture = imgDown.Picture
End Sub
Private Sub imgCity3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity3.Picture = imgUp.Picture
cmdFranklins_Click
End Sub

Private Sub imgCity4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity4.Picture = imgDown.Picture
End Sub
Private Sub imgCity4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity4.Picture = imgUp.Picture
cmdMaxi_Click
End Sub

Private Sub imgCity5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity5.Picture = imgDown.Picture
End Sub
Private Sub imgCity5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity5.Picture = imgUp.Picture
cmdIGA_click
End Sub

Private Sub imgCity6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity6.Picture = imgDown.Picture
End Sub
Private Sub imgCity6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity6.Picture = imgUp.Picture
cmdBilo_Click
End Sub
Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity6.Picture = imgDown.Picture
End Sub
Private Sub Label14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity6.Picture = imgUp.Picture
cmdBilo_Click
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity5.Picture = imgDown.Picture
End Sub
Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity5.Picture = imgUp.Picture
cmdIGA_click
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity4.Picture = imgDown.Picture
End Sub
Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity4.Picture = imgUp.Picture
cmdMaxi_Click
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity3.Picture = imgDown.Picture
End Sub
Private Sub Label17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity3.Picture = imgUp.Picture
cmdFranklins_Click
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity2.Picture = imgDown.Picture
End Sub
Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity2.Picture = imgUp.Picture
cmdColes_Click
End Sub



Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPriceHist.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPriceHist.Picture = imgOrigUp.Picture
    frmHistory.Start
    frmHistory.Show vbModal
End Sub

Private Sub imgCalc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCalc.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgCalc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCalc.Picture = imgOrigUp.Picture
    frmCalc.Show vbModal
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCalc.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCalc.Picture = imgOrigUp.Picture
    frmCalc.Show vbModal
End Sub

Private Sub imgBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBank.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBank.Picture = imgOrigUp.Picture
End Sub

Private Sub imgCity1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity1.Picture = imgDown.Picture
End Sub
Private Sub imgCity1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity1.Picture = imgUp.Picture
cmdSafeway_Click
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity1.Picture = imgDown.Picture
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCity1.Picture = imgUp.Picture
cmdSafeway_Click
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBank.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBank.Picture = imgOrigUp.Picture
End Sub

Private Sub imgBuy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBuy.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgBuy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBuy.Picture = imgOrigUp.Picture
    If iSpace = 0 Then Exit Sub
    If lstFoods.SelectedItem.ListSubItems(1).text > Credit Then
        MsgBox "You can't afford it, borrow some money if you really want it!", vbExclamation
        Else
        frmBuy.Show vbModal
    End If
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBuy.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgBuy.Picture = imgOrigUp.Picture
    If iSpace = 0 Then Exit Sub
    If lstFoods.SelectedItem.ListSubItems(1).text > Credit Then
        MsgBox "You can't afford it, borrow some money if you really want it!", vbExclamation
        Else
        frmBuy.Show vbModal
    End If
End Sub

Private Sub imgSell_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSell.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgSell_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSell.Picture = imgOrigUp.Picture
    If Used = 0 Then
        MsgBox "You have nothing to sell", vbInformation
        Else
        frmSell.Show vbModal
    End If
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSell.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSell.Picture = imgOrigUp.Picture
    If Used = 0 Then
        MsgBox "You have nothing to sell", vbInformation
        Else
        frmSell.Show vbModal
    End If
End Sub

Private Sub imgSteal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSteal.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgSteal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSteal.Picture = imgOrigUp.Picture
If iSpace > 0 Then frmSteal.Show vbModal
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSteal.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSteal.Picture = imgOrigUp.Picture
If iSpace > 0 Then frmSteal.Show vbModal
End Sub

Private Sub imgDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDrop.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgDrop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDrop.Picture = imgOrigUp.Picture
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDrop.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDrop.Picture = imgOrigUp.Picture
End Sub

Private Sub imgClinic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClinic.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgClinic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClinic.Picture = imgOrigUp.Picture
    If Health = 100 Then MsgBox "You don't need a doctor at the moment": Exit Sub
    If Credit < 10000 Then MsgBox "You can't afford a doctor": Exit Sub
    PlaySound SDir & "Doctor.wav", 0, 3
    frmDoctor.Show vbModal
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClinic.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgClinic.Picture = imgOrigUp.Picture
    If Health = 100 Then MsgBox "You don't need a doctor at the moment": Exit Sub
    If Credit < 10000 Then MsgBox "You can't afford a doctor": Exit Sub
    PlaySound SDir & "Doctor.wav", 0, 3
    frmDoctor.Show vbModal
End Sub

Private Sub imgStore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStore.Picture = imgOrigDwn.Picture
End Sub
Private Sub imgStore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStore.Picture = imgOrigUp.Picture
    frmStore.Show vbModal
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStore.Picture = imgOrigDwn.Picture
End Sub
Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStore.Picture = imgOrigUp.Picture
    frmStore.Show vbModal
End Sub

Private Sub lblDay_Change()
    Dim iRnd As Byte
    Dim Ans As Byte
    Dim j As Byte, i As Byte
    Randomize
    If Day > 1 Then
        PlaySound Truck, 0, 3
        iRnd = Int(Rnd * 70)
        Select Case iRnd
            Case 5, 32, 18
            'Checks to see if you own the 200 space truck and you can afford it.
            If Credit >= 600 And TotalSpace = 200 Then
                Ans = MsgBox("A guy on the bus offers you a duffelbag that can hold 400 items, do you want to buy this for $600?", vbYesNo + vbQuestion)
                If Ans = vbYes Then
                    Credit = Credit - 600
                    TotalSpace = 400
                    iSpace = iSpace + 200
                    lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                    lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                End If
            End If
            
            Case 13
            If iSpace >= 10 And Quantity(13) = 0 Then
                MsgBox "You found 4 cans of NO2 on a dead dude on the bus!", vbInformation
                lstItems.ListItems.Clear
                Quantity(13) = 4
                Avg(13) = 0
                j = 1
                Used = Used + 10
                iSpace = iSpace - 10
                lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                For i = 1 To 13
                    If Quantity(i) > 0 Then
                        frmMain.lstItems.ListItems.Add j, , frmMain.lstFoods.ListItems(i)
                        frmMain.lstItems.ListItems(j).ListSubItems.Add , , Round(Avg(i), 0)
                        frmMain.lstItems.ListItems(j).ListSubItems.Add , , Quantity(i)
                        j = j + 1
                    End If
                Next
            End If
            
            Case 30
            If Credit >= 10 Then
                MsgBox "Two guys jumped you on the bus! They took " & Format(Round(Credit / 3, 2), "###,###,###") & " and ran!", vbExclamation
                Credit = Credit - (Credit / 3)
                lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
            End If
            
            Case 41
            If Credit > 1000 And TotalSpace = 400 Then
                Ans = MsgBox("A guy on the bus offers you a duffelbag that can hold 600 items, do you want to buy this for $1,000?", vbYesNo + vbQuestion)
                If Ans = vbYes Then
                    Credit = Credit - 1000
                    TotalSpace = 600
                    iSpace = iSpace + 200
                    lblItems.Caption = "Items: " & Used & " of " & TotalSpace
                    lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                End If
            End If
            
            Case 53
            If Credit >= 10 Then
                MsgBox "The man busts you for j-walking! You have to pay a fine!", vbExclamation
                Credit = Credit - (Credit * 0.1) 'or credit/10 of course
                lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
            End If
            
            Case 9, 22, 68, 27, 48, 15, 42, 54
            If Sold = True Then
                frmManager.Show vbModal
            End If
            
            Case 3, 26, 34, 38, 45, 1
            If Attacked = True Or Stole = True Then
                PlaySound SDir & "Police.wav", 0, 3
                frmPolice.Show vbModal
            End If
        End Select
    End If
End Sub

Private Sub lblDebit_Change()
    If Debit > 3500000 And Debit < 4000000 Then MsgBox "The bank want their money, they don't trust you with that huge debit. Pay up or else they will take you to court", vbExclamation
    If Debit > 4000000 Then
        If Credit > Debit Then
            Dim Temp As Long
            MsgBox "The bank wants their money, you have enough money to pay off this loan. You have to pay it now", vbInformation
            Do
                frmFinances.Show vbModal
            Loop Until Debit = 0
            lblDebit.Caption = IIf(Debit <> 0, "$" & Format(Debit, "###,###,###"), 0)
            lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
        End If
        If Credit < Debit Then
            Dim Ret As Byte
            MsgBox "You don't have enough money to pay of this loan. Next time try to keep track of your debit", vbCritical
            Ret = MsgBox("Do you want to play again?", vbYesNo + vbQuestion)
            If Ret = vbYes Then
                frmMain.mnuNew_Click
                Else
                End
            End If
        End If
    End If
End Sub




Private Sub mnuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuFinances_Click()
    frmFinances.Show
End Sub

Private Sub mnuHelp_Click()
    MsgBox "I didn't bother adding help since this game shouldn't be hard to understand. If there is a bug or you cant understand how to use a feature on this game email adz8@softhome.net"
End Sub

Private Sub mnuHistory_Click()
    frmHistory.Start
    frmHistory.Show vbModal
End Sub

Private Sub mnuLoad_Click()
'CD1.FileName = ""
'CD1.DialogTitle = "Load Saved Progress"
'CD1.FilterIndex = 1
'CD1.flags = cdlOFNOverwritePrompt + cdlOFNNoChangeDir + cdlOFNLongNames
'CD1.CancelError = False
'CD1.DialogTitle = "Open"
'CD1.InitDir = App.Path
'CD1.Filter = "Saved Progress File (*.sav)|*.sav|"
'CD1.ShowOpen
'If CD1.FileName = "" Then Exit Sub
'CD1.FileName = LoadedData
'Open LoadedData For Output As #1

'Close #1
End Sub
Public Function Parse(sIn As String, sDel As String) As Variant
    Dim i As Integer, X As Integer, s As Integer, t As Integer
    i = 1: s = 1: t = 1: X = 1
    ReDim tArr(1 To X) As Variant


    If InStr(1, sIn, sDel) <> 0 Then


        Do
            ReDim Preserve tArr(1 To X) As Variant
            tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            X = X + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(1 To X) As Variant
        tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(1) = sIn
    End If
    Parse = tArr
End Function
Public Sub mnuNew_Click()
    Call Init
End Sub

Private Function EnableOthers(cmd1 As CommandButton, cmd2 As CommandButton, cmd3 As CommandButton, cmd4 As CommandButton, cmd5 As CommandButton)
    cmd1.Enabled = True
    cmd2.Enabled = True
    cmd3.Enabled = True
    cmd4.Enabled = True
    cmd5.Enabled = True
End Function

Private Function CButton(Button As CommandButton) As Long
    SendMessage Button.hWnd, BM_SETSTYLE, BS_SOLID, 1
End Function

Private Sub mnuOther_Click()
    If mnuOther.Checked = True Then
        mnuOther.Checked = False
        Else
        mnuOther.Checked = True
    End If
    SaveSetting "Food Wars", "Options", "Sounds", mnuOther.Checked
    Sound
End Sub

Private Sub mnuTruck_Click()
    If mnuTruck.Checked = True Then
        mnuTruck.Checked = False
        Else
        mnuTruck.Checked = True
    End If
    SaveSetting "Food Wars", "Options", "Truck", mnuTruck.Checked
    Sound
End Sub

Private Sub mnuViewScores_Click()
    Call HSL.FillScoreList(frmScores.lstScores)
    frmScores.Show vbModal
End Sub

Private Function CheckScore() As Boolean
    Dim i As Integer
    Dim Pass As Boolean
    Dim Cnt As Byte
    Dim Success As Boolean
    Success = False
    Cnt = 0
    CheckScore = False
    If Day = 364 Then CheckScore = True: ChangeCaptions
    If Day <> 365 Then CheckScore = True: Exit Function
    FinalScore = Credit
    For i = 1 To 13
        If Quantity(i) > 0 Then
            Cnt = Cnt + 1
        End If
    Next
    If Cnt > 0 Then MsgBox "You have to get rid of all your drugs before you can finish off!"
    If Cnt = 0 Then
        DisableControls
        Call HSL.FillScoreList(frmScores.lstScores)
        Success = HSL.AddHighScore(frmScores.lstScores, Int(FinalScore))
        CheckScore = True
        If Success = True Then
            frmScores.Show vbModal
            Exit Function
            Else
            i = MsgBox("Your time has run out! Play a new game or quit?")
        End If
    End If
End Function

Private Sub ChangeCaptions()
    cmdColes.Caption = "FINISH"
    cmdMaxi.Caption = "FINISH"
    cmdSafeway.Caption = "FINISH"
    cmdIGA.Caption = "FINISH"
    cmdBILO.Caption = "FINISH"
    cmdFranklins.Caption = "FINISH"
End Sub

Private Sub ResetCaptions()
    cmdColes.Caption = "&Coles"
    cmdMaxi.Caption = "&Maxi"
    cmdSafeway.Caption = "&Safeway"
    cmdIGA.Caption = "&IGA"
    cmdBILO.Caption = "&BI-LO"
    cmdFranklins.Caption = "&Franklins"
End Sub

Private Sub DisableControls()
    'This sub will be called when you approach the end of the days
    'first it disables all the COMMAND BUTTONS and then enabled the few needed
    'so the user can still view high scores, start a new game and view prices from
    'other days (this is more efficent than disabling only the command buttons
    'needed)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        If TypeOf Controls(i) Is CommandButton Then
            Controls(i).Enabled = False
        End If
    Next
    'cmdHistory.Enabled = True 'So they can view prices
    mnuFinances.Enabled = False 'so they cant mess around with their previous game
End Sub

Private Sub EnableControls()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Controls.Count - 1
        If TypeOf Controls(i) Is CommandButton Then
            Controls(i).Enabled = True
        End If
    Next
    mnuFinances.Enabled = True
End Sub

Private Sub Sound()
    mnuTruck.Checked = GetSetting("Food Wars", "Options", "Truck", True)
    mnuOther.Checked = GetSetting("Food Wars", "Options", "Sounds", True)
    If mnuOther.Checked = False Then
        'SDir is needed for the program to know where the sounds are and resetting
        'SDir to nothing is easier than making IF statements for every time a
        'sound is going to be played.
        SDir = ""
        Else
        If Right(App.Path, 1) = "\" Then
            SDir = App.Path & "Sounds\"
            Else
            SDir = App.Path & "\Sounds\"
        End If
    End If
    If mnuTruck.Checked = False Then
        Truck = vbNullString
        Else
        If Right(App.Path, 1) = "\" Then
            Truck = App.Path & "Sounds\"
            Else
            Truck = App.Path & "\Sounds\"
        End If
        Truck = Truck & "Truck.wav"
    End If
End Sub

Private Sub sckClient_Connect()
sckClient.SendData "Name" & vbTab & User.Name & vbCrLf
frmMain.Show
Unload frmName
End Sub
Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
Dim Data As String, MainData() As String, SplitData() As String
sckClient.GetData Data, vbString
MainData = Split(Data, vbCrLf)
For X = LBound(MainData) To UBound(MainData) - 1
SplitData = Split(MainData(X), vbTab)
Select Case SplitData(0)

Case "Message"
Message SplitData(1)

Case "Kicked"
frmLogin.Show
Unload Me

End Select
Next X
End Sub
Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "BAH COULDN'T CONNECT!!!", vbExclamation
End Sub

Private Sub Timer1_Timer()
Label13.Caption = "Time: " & Time
End Sub
