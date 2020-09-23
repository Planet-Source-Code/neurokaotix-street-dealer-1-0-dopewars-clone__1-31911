VERSION 5.00
Begin VB.UserControl ProgBar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3690
   ToolboxBitmap   =   "Progbar.ctx":0000
End
Attribute VB_Name = "ProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************
' ProgBar V1.0  (C)1998 David Crowell
' You may use this source code within
' your own applications.  You may not
' distribute it on a website without my express
' permission.
'
' http://www.qtm.net/~davidc
' davidc@qtm.net
'********************************

Public Enum BorderStyles    ' BorderStyles for the control
    bdNone
    bdFixedSingle
End Enum

Public Event Click()            ' yup I coded a click event

'********************************
' Here are the private variables
' that contain the properties
'********************************
Private mBackColor As Long
Private mBarColor As Long
Private mVertical As Boolean
Private mMin As Long
Private mMax As Long
Private mValue As Long
Private mBorderStyle As Long

'********************************
' All properties are read/write
'********************************
' If you get an error here, go to project references, and be
' sure that OLE Automation is selected.  If you don't want
' to do that, change the OLE_COLOR to Long.  It will work,
' but you won't get the pretty color picker in the properties
' window.
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Color of the control under the bar."
Attribute BackColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    mBackColor = NewColor
    UserControl.BackColor = NewColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BarColor(ByVal NewColor As OLE_COLOR)
Attribute BarColor.VB_Description = "Color of the progress bar."
Attribute BarColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    mBarColor = NewColor
    UserControl_Paint
    PropertyChanged "BarColor"
End Property
Public Property Get BarColor() As OLE_COLOR
    BarColor = mBarColor
End Property

Public Property Let Vertical(ByVal val As Boolean)
Attribute Vertical.VB_Description = "If true, the progress bar advances from bottom to top, otherwise, left to right."
Attribute Vertical.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    mVertical = val
    UserControl_Resize
    PropertyChanged "Vertical"
End Property
Public Property Get Vertical() As Boolean
    Vertical = mVertical
End Property

Public Property Let Max(ByVal val As Long)
Attribute Max.VB_Description = "The upper limit for value."
Attribute Max.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    If val < 1 Then val = 1
    If val <= mMin Then val = mMin + 1
    mMax = val
    If Value > mMax Then Value = mMax
    UserControl_Resize
    PropertyChanged "Max"
End Property
Public Property Get Max() As Long
    Max = mMax
End Property

Public Property Let Min(ByVal val As Long)
Attribute Min.VB_Description = "The lower limit for value."
Attribute Min.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    If val >= mMax Then val = Max - 1
    If val < 0 Then val = 0
    mMin = val
    If Value < mMin Then Value = mMin
    UserControl_Resize
    PropertyChanged "Min"
End Property
Public Property Get Min() As Long
    Min = mMin
End Property

Public Property Let Value(ByVal val As Long)
Attribute Value.VB_Description = "The value determines the position of the progress bar."
Attribute Value.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    If val > mMax Then val = Max
    If val < mMin Then val = mMin
    mValue = val
    UserControl_Paint
    PropertyChanged "Value"
End Property
Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let BorderStyle(ByVal val As BorderStyles)
Attribute BorderStyle.VB_Description = "Will the control have a visible border?"
Attribute BorderStyle.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    If val < 0 Then val = 0
    If val > 1 Then val = 1
    mBorderStyle = val
    UserControl.BorderStyle = mBorderStyle
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property
Public Property Get BorderStyle() As BorderStyles
    BorderStyle = mBorderStyle
End Property

'********************************
' Set up the defaults
'********************************
Private Sub UserControl_InitProperties()
    BackColor = vbButtonFace
    BarColor = vbHighlight
    Vertical = False
    Max = 100
    Min = 0
    Value = 50
    BorderStyle = 1
End Sub

'********************************
' Reload design-time settings
'********************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    BarColor = PropBag.ReadProperty("BarColor", vbHighlight)
    Vertical = PropBag.ReadProperty("Vertical", False)
    Max = PropBag.ReadProperty("Max", 100)
    Min = PropBag.ReadProperty("Min", 0)
    Value = PropBag.ReadProperty("Value", 50)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

'********************************
' Save design-time settings
'********************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
    PropBag.WriteProperty "BarColor", BarColor, vbHighlight
    PropBag.WriteProperty "Vertical", Vertical, False
    PropBag.WriteProperty "Max", Max, 100
    PropBag.WriteProperty "Min", Min, 0
    PropBag.WriteProperty "Value", Value, 50
    PropBag.WriteProperty "BorderStyle", BorderStyle, 1
End Sub

'********************************
' The bulk of the work is this small little
' sub.  It does the drawing.
'********************************
Private Sub UserControl_Paint()
    Dim w As Long           ' I'm storing some properties
    Dim h As Long           ' in variables to improve performance
    Dim v As Long
    v = mValue - mMin
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    If mVertical Then                                                     ' is this a vertical control?
        UserControl.Line (0, 0)-(w, h - v), mBackColor, BF  ' draw the background color
        If v > 0 Then   ' only draw the bar if there is one to draw
            UserControl.Line (0, h)-(w, h - v), mBarColor, BF   ' draw the bar
        End If
    Else
        UserControl.Line (v, 0)-(w, h), mBackColor, BF          ' this is the same code as above
        If v > 0 Then
            UserControl.Line (0, 0)-(v, h), mBarColor, BF         ' but for horizontal controls'        End If
        End If
    End If
End Sub

'********************************
' There is a little more work to be done
' if the control is resized
'********************************
Private Sub UserControl_Resize()
    On Error Resume Next        ' just in case
    UserControl.ScaleWidth = mMax - mMin
    UserControl.ScaleHeight = mMax - mMin
    UserControl_Paint               ' repaint the control
End Sub

'********************************
' This is really simple.  Catch the click event
' in the usercontrol, and pass it on to the
' container form.
'********************************
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

