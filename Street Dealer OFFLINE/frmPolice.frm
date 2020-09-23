VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPolice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Its the fuzz!"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdFight 
      Caption         =   "&Fight"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstWeapons 
      Height          =   1335
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
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
   Begin StreetDealer.ProgBar pbHealth 
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      BackColor       =   255
      BarColor        =   65280
      BorderStyle     =   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5205
   End
   Begin VB.Label lblDetails 
      BackStyle       =   0  'Transparent
      Caption         =   "xx cops are after you. You better run while you still can."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmPolice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PCount As Integer 'Amount of officers after you
Dim Sel As Byte

Private Sub cmdFight_Click()
    Dim i As Integer
    Dim RndHit As Byte 'Whether it hits or not
    Dim rndHealth As Byte 'How much health will be removed if the cops get you
    Dim rndKill As Byte 'How many cops it took out
    Dim rndGain As Long 'How much you steel of the cops
    Dim rHit As Byte 'Chance of miss
    Randomize
    Attacked = True
    If lstWeapons.SelectedItem.ListSubItems(1) = 1 Then
        rHit = Int(5 * Rnd) + 1
        If rHit = 2 Then
            Do
                rndGain = Int((Credit * 0.2) * Rnd + 1)
            Loop Until rndGain > (Credit * 0.05)
            lblStatus.Caption = "You killed one!"
            iDisabled
            PlaySound SDir & "9mmHit.wav", 0, 0
            iEnabled
            PCount = PCount - 1
            If PCount <= 0 Then
                Credit = Credit + rndGain
                MsgBox "You killed the cops and got " & Format(rndGain, "###,###,###") & " from them"
                frmMain.pbHealth.Value = Health
                frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                Unload frmPolice
                Exit Sub
            End If
            lblDetails.Caption = PCount & " cop(s) are after you. You better run while you still can."
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "9mmMiss.wav", 0, 0
            iEnabled
        End If
    End If
    If lstWeapons.SelectedItem.ListSubItems(1) = 2 Then
        rHit = Int(2 * Rnd) + 1
        If rHit = 2 Then
            Do
                rndGain = Int((Credit * 0.2) * Rnd + 1)
            Loop Until rndGain > (Credit * 0.05)
            lblStatus.Caption = "You killed one!"
            iDisabled
            PlaySound SDir & "mgnmhit.wav", 0, 0
            iEnabled
            PCount = PCount - 1
            If PCount <= 0 Then
                Credit = Credit + rndGain
                MsgBox "You killed the cops and got " & Format(rndGain, "###,###,###") & " from them"
                frmMain.pbHealth.Value = Health
                frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                Unload frmPolice
                Exit Sub
            End If
            lblDetails.Caption = PCount & " cop(s) are after you. You better run while you still can."
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
            Do
                rndGain = Int((Credit * 0.2) * Rnd + 1)
            Loop Until rndGain > (Credit * 0.05)
            rndKill = Int(3 * Rnd) + 1
            If rndKill > 1 And PCount > 1 Then
                lblStatus.Caption = "You killed a few cops"
                Else
                lblStatus.Caption = "You killed a cop"
            End If
            iDisabled
            PlaySound SDir & "mghit.wav", 0, 0
            iEnabled
            PCount = PCount - rndKill
            If PCount <= 0 Then
                Credit = Credit + rndGain
                MsgBox "You killed the cops and got " & Format(rndGain, "###,###,###") & " from them"
                frmMain.pbHealth.Value = Health
                frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                Unload frmPolice
                Exit Sub
            End If
            lblDetails.Caption = PCount & " cop(s) are after you. You better run while you still can."
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
            Do
                rndGain = Int((Credit * 0.2) * Rnd + 1)
            Loop Until rndGain > (Credit * 0.05)
            rndKill = Int(5 * Rnd) + 1
            If rndKill > 1 And PCount > 1 Then
                lblStatus.Caption = "You killed a few cops"
                Else
                lblStatus.Caption = "You killed a cop"
            End If
            iDisabled
            PlaySound SDir & "rlHit.wav", 0, 0
            iEnabled
            PCount = PCount - rndKill
            If PCount <= 0 Then
                Credit = Credit + rndGain
                MsgBox "You killed the cops and got " & Format(rndGain, "###,###,###") & " from them"
                frmMain.pbHealth.Value = Health
                frmMain.lblCash.Caption = IIf(Credit <> 0, "$" & Format(Credit, "###,###,###"), 0)
                Unload frmPolice
                Exit Sub
            End If
            lblDetails.Caption = PCount & " cop(s) are after you. You better run while you still can."
            Else
            lblStatus.Caption = "You missed!"
            iDisabled
            PlaySound SDir & "rlMiss.wav", 0, 0
            iEnabled
        End If
    End If
    'His turn to fight
    RndHit = Int(3 * Rnd) + 1
    Do
        rndHealth = Int(15 * Rnd) + 1
    Loop Until rndHealth > 5
    
    If RndHit = 1 Then
        lblStatus.Caption = "One cop took a shot and got you!"
        iDisabled
        PlaySound SDir & "cophit.wav", 0, 0
        Health = Health - rndHealth
        lblStatus.Caption = vbNullString
        iEnabled
    End If
    
    If RndHit = 2 Or RndHit = 3 Then
        lblStatus.Caption = "One cop fired and just missed"
        iDisabled
        PlaySound SDir & "copmiss.wav", 0, 0
        lblStatus.Caption = vbNullString
        iEnabled
    End If
    pbHealth.Value = Health
    cmdFight.Default = True
    lstWeapons.SetFocus
End Sub

Private Sub cmdRun_Click()
    On Error Resume Next
    Dim RndHit As Byte
    Dim rndHealth As Byte
    Randomize
    Sel = lstWeapons.SelectedItem.Index
    RndHit = Int(4 * Rnd) + 1
    Do
        rndHealth = Int(15 * Rnd) + 1
    Loop Until rndHealth > 5
    
    If RndHit = 1 Then
        lblStatus.Caption = "One cop took a shot and got you!"
        iDisabled
        PlaySound SDir & "cophit.wav", 0, 0
        Health = Health - rndHealth
        lblStatus.Caption = vbNullString
        iEnabled
    End If
    
    If RndHit = 2 Or RndHit = 3 Then
        lblStatus.Caption = "One cop fired and just missed"
        iDisabled
        PlaySound SDir & "copmiss.wav", 0, 0
        lblStatus.Caption = vbNullString
        iEnabled
    End If
    
    If RndHit = 4 Then
        lblStatus.Caption = "You lost them!"
        iDisabled
        PlaySound SDir & "Whew.wav", 0, 0
        lblStatus.Caption = vbNullString
        iEnabled
        frmMain.pbHealth.Value = Health
        Unload frmPolice
    End If
    
    pbHealth.Value = Health
    cmdRun.Default = True
End Sub

Private Sub Form_Activate()
    Dim i As Byte
    Dim j As Byte
    Randomize
    j = 1
    Do
        PCount = Int(Rnd * 10) + 1
    Loop Until PCount > 1
    lblDetails.Caption = PCount & " cop(s) are after you. You better run while you still can."
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

Private Sub iEnabled()
    If lstWeapons.ListItems.Count > 0 Then cmdFight.Enabled = True
    cmdRun.Enabled = True
    Refresh
End Sub

Private Sub iDisabled()
    cmdFight.Enabled = False
    cmdRun.Enabled = False
    Refresh
End Sub
