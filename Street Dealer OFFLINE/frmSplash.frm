VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Binary Cipher"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   DrawStyle       =   6  'Inside Solid
   DrawWidth       =   2
   FillStyle       =   0  'Solid
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   939.815
   ScaleMode       =   0  'User
   ScaleWidth      =   2857.988
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   71.006
      X2              =   2769.231
      Y1              =   444.445
      Y2              =   444.445
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTime As Date

Sub NewGradient(R1 As Integer, G1 As Integer, B1 As Integer, R2 As Integer, G2 As Integer, B2 As Integer, obj As Form, WhichWay As Integer, TopOrBottom As Integer)
    Vert = 0
    Horz = 1
    If WhichWay = Vert Then pixels = obj.Height
    If WhichWay = Horz Then pixels = obj.Width
    If R1 < R2 Then
        tempR1 = R1
        tempR2 = R2
        R1 = tempR2
        R2 = tempR1
    End If


    If G1 < G2 Then
        tempG1 = G1
        tempG2 = G2
        G1 = tempG2
        G2 = tempG1
    End If


    If B1 < B2 Then
        tempB1 = B1
        tempB2 = B2
        B1 = tempB2
        B2 = tempB1
    End If
    'Set the Value for how much the Red, Blu
    '     e, and Green will go
    'up each time
    If (R1 - R2) <> 0 Then nRStep = (R1 - R2) / pixels
    If (G1 - G2) <> 0 Then nGStep = (G1 - G2) / pixels
    If (B1 - B2) <> 0 Then nBStep = (B1 - B2) / pixels
    'Fill in Gradient


    For X = 1 To pixels
        'Set Red, Green, and Blue values. Light
        '     on top
        'Darker as you go down


        If TopOrBottom = 0 Then
            nR = nR + nRStep
            nG = nG + nGStep
            nB = nB + nBStep
            r = R1 - nR
            G = G1 - nG
            b = B1 - nB
        End If
        'Set Red, Green, and Blue values. Dark o
        '     n Top,
        'Lighter as you go down


        If TopOrBottom = 1 Then
            r = r + nRStep
            G = G + nGStep
            b = b + nBStep
        End If
        'Make sure R, G, or B don't go less then
        '     zero,
        'Because this would cause an erro
        If r < 0 Then r = 0
        If G < 0 Then G = 0
        If b < 0 Then b = 0
        'If WhichWay = Vert then draw Horizontal
        '     line
        If WhichWay = Vert Then obj.Line (1, X)-(obj.Width, X), RGB(r, G, b), BF
        'If WhichWay = Horz then draw Vertical l
        '     ine
        If WhichWay = Horz Then obj.Line (X, 1)-(X, obj.Height), RGB(r, G, b), BF
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload frmSplash
End Sub

Private Sub Form_Load()
    NewGradient 10, 0, 30, 40, 192, 324, frmSplash, 0, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload frmSplash
End Sub
