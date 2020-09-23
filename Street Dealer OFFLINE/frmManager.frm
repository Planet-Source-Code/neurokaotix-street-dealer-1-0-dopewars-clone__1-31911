VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You got jumped!"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstWeapons 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2355
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Weapon"
         Text            =   "Weapons"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "UD"
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.CommandButton cmdStay 
      Caption         =   "&Stay"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "&Fight"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin StreetDealer.ProgBar pbHealth 
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      BackColor       =   255
      BarColor        =   65280
      BorderStyle     =   0
   End
   Begin VB.Label lblDetails 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmManager.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5205
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFight_Click()
    Dim i As Integer
    Dim RndHit As Byte
    Dim rndWeapon As Byte
    Dim rndHealth As Byte
    Dim rHit As Byte
    Randomize
    Attacked = True
    If lstWeapons.SelectedItem.ListSubItems(1) = 1 Then
        rHit = Int(5 * Rnd) + 1
        If rHit = 2 Then
            lblStatus.Caption = "You killed him!"
            iDisabled
            PlaySound SDir & "9mmHit.wav", 0, 0
            iEnabled
            Unload frmManager
            Exit Sub
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "9mmMiss.wav", 0, 0
            iEnabled
        End If
    End If
    If lstWeapons.SelectedItem.ListSubItems(1) = 2 Then
        rHit = Int(3 * Rnd) + 1
        If rHit = 2 Then
            lblStatus.Caption = "You killed him!"
            iDisabled
            PlaySound SDir & "mgnmhit.wav", 0, 0
            iEnabled
            Unload frmManager
            Exit Sub
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "mgnmmiss.wav", 0, 0
            iEnabled
        End If
    End If
    If lstWeapons.SelectedItem.ListSubItems(1) = 3 Then
        rHit = Int(5 * Rnd) + 1
        If rHit = 1 Or rHit = 2 Or rHit = 3 Or rHit = 5 Then '20% chance of miss
            lblStatus.Caption = "You killed him!"
            iDisabled
            PlaySound SDir & "mghit.wav", 0, 0
            iEnabled
            Unload frmManager
            Exit Sub
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "mgmiss.wav", 0, 0
            iEnabled
        End If
    End If
    If lstWeapons.SelectedItem.ListSubItems(1) = 4 Then
        rHit = Int(7 * Rnd) + 1
        If rHit = 1 Or rHit = 2 Or rHit = 3 Or rHit = 5 Or rHit = 6 Or rHit = 7 Then '14% chance of miss
            lblStatus.Caption = "You killed him!"
            iDisabled
            PlaySound SDir & "rlHit.wav", 0, 0
            iEnabled
            Unload frmManager
            Exit Sub
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "rlMiss.wav", 0, 0
            iEnabled
        End If
    End If
    'His turn to fight now
    rndWeapon = Int(2 * Rnd) + 1
    RndHit = Int(5 * Rnd) + 1
    Do
        rndHealth = Int(15 * Rnd) + 1
    Loop Until rndHealth > 7
    If RndHit = 2 Or RndHit = 4 Then
        If rndWeapon = 1 Then 'Chainsaw hit
            lblStatus.Caption = "The chainsaw got you!"
            iDisabled
            PlaySound SDir & "csHit.wav", 0, 0
            Health = Health - rndHealth
            lblStatus.Caption = vbNullString
            iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
        End If
        If rndWeapon = 2 Then 'Can hit
            lblStatus.Caption = "He fired and hit you!"
            iDisabled
            PlaySound SDir & "fiveseven-1.wav", 0, 0
            Health = Health - rndHealth
            lblStatus.Caption = vbNullString
            iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
        End If
    End If
    If RndHit = 3 Or RndHit = 1 Or RndHit = 5 Then
        If rndWeapon = 2 Then
            lblStatus.Caption = "He fired and missed!"
            iDisabled
            PlaySound SDir & "ric2.wav", 0, 0  'can miss
            lblStatus.Caption = vbNullString
            iEnabled
        End If
        If rndWeapon = 1 Then
            lblStatus.Caption = "He tried to get you with a chainsaw and missed!"
            iDisabled
            PlaySound SDir & "csmiss.wav", 0, 0   'chainsaw miss
            lblStatus.Caption = vbNullString
            iEnabled
        End If
    End If
    frmMain.pbHealth.Value = Health
    pbHealth.Value = Health
    cmdFight.Default = True
End Sub

Private Sub cmdRun_Click()
    Dim RndHit As Byte
    Dim rndWeapon As Byte
    Dim rndHealth As Byte
    Randomize
    rndWeapon = Int(2 * Rnd) + 1
    RndHit = Int(7 * Rnd) + 1
    Do
        rndHealth = Int(15 * Rnd) + 1
    Loop Until rndHealth > 7
    If RndHit = 3 Or RndHit = 4 Then
        If rndWeapon = 1 Then 'Chainsaw hit
            lblStatus.Caption = "The chainsaw got you!"
            iDisabled
            PlaySound SDir & "csHit.wav", 0, 0
            Health = Health - rndHealth
            lblStatus.Caption = vbNullString
            iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
        End If
        If rndWeapon = 2 Then 'Can hit
            lblStatus.Caption = "He fired and hit you!"
            iDisabled
            PlaySound SDir & "fiveseven-1.wav", 0, 0
            Health = Health - rndHealth
            lblStatus.Caption = vbNullString
            iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
        End If
    End If
    If RndHit = 6 Or RndHit = 1 Or RndHit = 5 Then
        If rndWeapon = 2 Then
            lblStatus.Caption = "He fired but missed!"
            iDisabled
            PlaySound SDir & "ric2.wav", 0, 0  'miss with can
            lblStatus.Caption = vbNullString
            iEnabled
        End If
        If rndWeapon = 1 Then
            lblStatus.Caption = "He tried to get you with a chainsaw and missed!"
            iDisabled
            PlaySound SDir & "csmiss.wav", 0, 0   'miss with chainsaw
            lblStatus.Caption = vbNullString
            iEnabled
        End If
    End If
    If RndHit = 2 Or RndHit = 7 Then
        lblStatus.Caption = "You lost him behind a dumpster!"
        iDisabled
        PlaySound SDir & "whew.wav", 0, 0
        lblStatus.Caption = vbNullString
        iEnabled
        Unload frmManager
    End If
    frmMain.pbHealth.Value = Health
    pbHealth.Value = Health
    cmdRun.Default = True
End Sub

Private Sub iEnabled()
    If lstWeapons.ListItems.Count > 0 Then cmdFight.Enabled = True
    cmdRun.Enabled = True
    cmdStay.Enabled = True
    Refresh
End Sub

Private Sub iDisabled()
    cmdFight.Enabled = False
    cmdRun.Enabled = False
    cmdStay.Enabled = False
    Refresh
End Sub

Private Sub cmdStay_Click()
    Dim RndHit As Byte
    Dim rndWeapon As Byte
    Dim rndHealth As Byte
    Dim i As Byte
    Dim j As Byte
    For i = 1 To 13
        If Quantity(i) > 0 Then j = j + 1
    Next
    If j > 0 Then
        Randomize
        rndWeapon = Int(2 * Rnd) + 1
        RndHit = Int(5 * Rnd) + 1
        MsgBox "You stand still hoping for mercy but he is still attacks you!"
        Do
            rndHealth = Int(15 * Rnd) + 1
        Loop Until rndHealth > 7
        If RndHit = 3 Or RndHit = 4 Then
            If rndWeapon = 1 Then 'Chainsaw hit
                lblStatus.Caption = "The chainsaw got you!"
                iDisabled
                PlaySound SDir & "csHit.wav", 0, 0
                Health = Health - rndHealth
                lblStatus.Caption = vbNullString
                iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
            End If
            If rndWeapon = 2 Then 'Can hit
                lblStatus.Caption = "He fired and hit you!"
                iDisabled
                PlaySound SDir & "fiveseven-1.wav", 0, 0
                Health = Health - rndHealth
                lblStatus.Caption = vbNullString
                iEnabled
            If Health <= 0 Then MsgBox "He wasted your ass!", vbExclamation: Unload frmManager
            End If
        End If
        If RndHit = 2 Or RndHit = 1 Or RndHit = 5 Then
            If rndWeapon = 2 Then
                lblStatus.Caption = "He fired but missed!"
                iDisabled
                PlaySound SDir & "ric2.wav", 0, 0  'miss with can
                lblStatus.Caption = vbNullString
                iEnabled
            End If
            If rndWeapon = 1 Then
                lblStatus.Caption = "He tried to get you with a chainsaw and missed!"
                iDisabled
                PlaySound SDir & "csmiss.wav", 0, 0   'miss with chainsaw
                lblStatus.Caption = vbNullString
                iEnabled
            End If
        End If
        frmMain.pbHealth.Value = Health
        pbHealth.Value = Health
        cmdStay.Default = True
        Exit Sub
    Else
        MsgBox "He feels sorry for you because you're so weak and decides to let you go!"
        Unload frmManager
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    Dim j As Integer
    Dim i As Integer
    j = 1
    pbHealth.Value = Health
    For i = 1 To 4
        If YW(i) = True Then
            lstWeapons.ListItems.Add j, , Weapon(i)
            lstWeapons.ListItems(j).ListSubItems.Add , , i
            j = j + 1
        End If
    Next
    If lstWeapons.ListItems.Count = 0 Then cmdFight.Enabled = False
End Sub

