VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLACK-MART"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Guns/Weapons"
      TabPicture(0)   =   "frmStore.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstWeapons"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Slightly Illegal"
      TabPicture(1)   =   "frmStore.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstIllegal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Miscellaneous"
      TabPicture(2)   =   "frmStore.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstMisc"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView lstWeapons 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   2990
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Weapon"
            Text            =   "Item"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "Price"
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstIllegal 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   2990
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Weapon"
            Text            =   "Item"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "Price"
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstMisc 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   2990
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Weapon"
            Text            =   "Item"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "Price"
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty."
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Add to shopping basket"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Buy"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Add to shopping basket"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstBought 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Weapon"
         Text            =   "Weapon"
         Object.Width           =   6438
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Items:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1020
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price(1 To 8) As Long

Private Sub cmdAdd_Click()
    Dim i As Byte
    If YW(lstWeapons.SelectedItem.Index) = True Then Exit Sub
    If lstWeapons.SelectedItem.ListSubItems(1).text > Credit Then
        MsgBox "You cannot afford to pay for this"
        Else
        Credit = Credit - lstWeapons.SelectedItem.ListSubItems(1).text
        YW(lstWeapons.SelectedItem.Index) = True
        frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
        lstBought.ListItems.Clear
        For i = 1 To UBound(Weapon)
        If YW(i) = True Then
            lstBought.ListItems.Add , , Weapon(i)
        End If
    Next
    End If
    If YW(1) = True And YW(2) = True And YW(3) = True And YW(4) = True Then cmdAdd.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload frmStore
End Sub

Private Sub Form_Load()
    Dim i As Byte
    Price(1) = 20
    Price(2) = 35
    Price(3) = 110
    Price(4) = 195
    Price(5) = 410
    Price(6) = 795
    Price(7) = 995
    Price(8) = 4550
    If YW(1) = True And YW(2) = True And YW(3) = True And YW(4) = True Then cmdAdd.Enabled = False
    For i = 1 To UBound(Weapon)
        lstWeapons.ListItems.Add i, , Weapon(i)
        lstWeapons.ListItems(i).ListSubItems.Add , , "$" + Format(Price(i), "###,###,###") + ".00"
    Next
    For i = 1 To 8
        If YW(i) = True Then
            lstBought.ListItems.Add , , Weapon(i)
        End If
    Next
End Sub

