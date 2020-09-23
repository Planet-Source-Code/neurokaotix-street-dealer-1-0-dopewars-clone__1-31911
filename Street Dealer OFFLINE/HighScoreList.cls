VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HighScoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim MaxPeople As Integer '= 100
Dim ChrCount As Integer '= 10
Dim Separator As String
Dim TheFile As String

Public Sub FillScoreList(TheList As ListBox)
    Dim i As Integer, Item As String
    Dim NewFileNeeded As Boolean, TheLen As Long
    ' Set the Separator
    Separator = String(3, Chr(9))
    If MaxPeople = 0 Then
        ' The default amount of items to load
        MaxPeople = 100
    End If
    If TheFile = "" Then
        ' The default file to save to
        TheFile = FixPath & "HighScores.ini"
    End If
    If Len(Dir(TheFile)) = 0 Then
        NewFileNeeded = True
    Else
        Open TheFile For Input As #1
            ' take out a big chunk
            TheLen = LOF(1)
            If TheLen > 32000 Then TheLen = 32000
            Item = Input(TheLen, #1)
            ' see if file has too few lines
            If CountLines(Item) < MaxPeople Then
                NewFileNeeded = True
            Else
                NewFileNeeded = False
            End If
        Close #1
    End If
    
    If NewFileNeeded = True Then
        ' the file doesn't exist or wasn't any good
        ' so fill the list and file up with blanks
        Open TheFile For Output As #1
            For i = 0 To MaxPeople - 1
                Print #1, Encrypt("Blank" & Separator & "0")
                TheList.AddItem (i + 1) & ". " & FixName("Blank", TheList) & Separator & "0"
            Next
        Close #1
    Else
        ' the file does exist
        Open TheFile For Input As #1
            For i = 0 To MaxPeople - 1
                Line Input #1, Item
                Item = Encrypt(Item)
                TheList.AddItem (i + 1) & ". " & GetPersonsName(Item, TheList) & Separator & GetPersonsScore(Item)
            Next
        Close #1
    End If
End Sub

Public Sub SetSeparator(NewString As String)
    ' This sub lets you set the string that will separate the names
    ' from the scores. I recommend you only use tabs for this.
    Separator = NewString
End Sub

Public Sub SetListCount(NumPeople As Integer)
    ' This sub lets you set how many names you want in the list.
    MaxPeople = NumPeople
End Sub

Public Sub SetNameLength(Length As Integer)
    ' This sub lets you set how long you want names in the list to be
    ChrCount = Length
End Sub

Public Function GetNameLength() As Integer
    ' Returns the maximum length as name can be
    If ChrCount = 0 Then
        GetNameLength = 10
    Else
        GetNameLength = ChrCount
    End If
End Function

Private Function ShouldScoreGoIn(TheList As ListBox, Score As Long) As Integer
    ' This function checks to see if the new entry should be added to the list.
    ' If it should, the function will return the INDEX in which it should go.
    ' If it shouldn't, the function will return -1.
    Dim i As Integer
    ShouldScoreGoIn = -1
    For i = 0 To MaxPeople - 1
        
        If CLng(GetPersonsScore(TheList.List(i))) < Score Then
            ShouldScoreGoIn = i
            Exit For
        End If
    Next
End Function

Public Function AddHighScore(TheList As ListBox, PName As String, PScore As Long) As Boolean
    ' Call this sub when you want to see if an item should be added
    ' to the highscore list. If it should, it will add the item,
    ' save the list, and the function will return as true. If it
    ' shouldn't be added the function will return as false.
    Dim Index As Integer, i As Integer
    Index = ShouldScoreGoIn(TheList, PScore)
    If Index = -1 Then
        AddHighScore = False
    Else
        For i = TheList.ListCount - 1 To Index Step -1
            TheList.List(i) = (i + 1) & ". " & Mid(TheList.List(i - 1), InStr(1, TheList.List(i - 1), ".") + 2)
        Next
        TheList.List(Index) = (Index + 1) & ". " & FixName(PName, TheList) & Separator & PScore
        
        ' save file
        If TheFile = "" Then
            ' The default file to save to
            TheFile = FixPath & "HighScores.ini"
        End If
        Open TheFile For Output As #1
            For i = 0 To MaxPeople - 1
                Print #1, Encrypt(Mid(TheList.List(i), InStr(1, TheList.List(i), ".") + 2))
            Next
        Close #1
        AddHighScore = True
    End If
End Function

Private Function CountLines(txt As String) As Integer
    ' Counts the number of lines in a string.
    Dim i As Long, X As Long
    Do While InStr(i + 1, txt, Chr(10))
        i = InStr(i + 1, txt, Chr(10))
        X = X + 1
    Loop
    CountLines = X + 1
End Function

Public Function Encrypt(text As String) As String
    ' This is the function that this example uses to encrypt text.
    ' It's not the best encrypting function, but it works pretty good.
    ' Call it once to encrypt text, and again to unencrypt it.
    Dim Chrlist1 As String, Chrlist2 As String, i As Integer
   Dim Pos As Integer, NewText As String
    Chrlist1 = " '?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥" & Chr(9)
    Chrlist2 = " '¿¡@#$%^&*()_+|10²³457698ÀbÁdÂÃghÄjklmÅÒÓqÔÕÖÙvwÛÜz.,-~AàáâãFGHäJKåMNØ¶QR§TÚVWX¥Z?!23acefinoprstuxyBCDEILOPSUY" & Chr(9)
    For i = 1 To Len(text)
        Pos = InStr(Chrlist1, Mid(text, i, 1))
        If Not Pos = 0 Then
            NewText = NewText + Mid(Chrlist2, Pos, 1)
        End If
    Next
    Encrypt = NewText
End Function

Private Function FixPath() As String
    ' Returns the application's file path and makes sure it's set up right.
    Dim FPath As String
    FPath = App.Path
    If Right(FPath, 1) = "\" Then
        FixPath = FPath
    Else
        FixPath = FPath & "\"
    End If
End Function

Private Function FixName(txt As String, TheList As ListBox) As String
    ' This function makes sure everyone's name is of the right length
    Dim OldFontFace As String, OldFontSize As Integer
    If ChrCount = 0 Then
        ChrCount = 10
    End If
    If TheList.Parent.TextWidth(txt) < TheList.Parent.TextWidth(String(ChrCount, "N")) Then
        ' store the old font values of the form in variables
        OldFontFace = TheList.Parent.Font
        OldFontSize = TheList.Parent.FontSize
        ' set the form's font values to that of the listbox
        TheList.Parent.Font = TheList.Font
        TheList.Parent.FontSize = TheList.FontSize
        ' loop until the text is of the right length
        Do While TheList.Parent.TextWidth(txt) < TheList.Parent.TextWidth(String(ChrCount, "N"))
            txt = txt & " "
        Loop
        ' change the form's font values back
        TheList.Parent.Font = OldFontFace
        TheList.Parent.FontSize = OldFontSize
        FixName = txt
    Else
        FixName = txt
    End If
End Function

Private Function GetPersonsName(txt As String, TheList As ListBox) As String
    ' Returns the person's name from the list item.
    GetPersonsName = FixName(Mid(txt, 1, InStr(1, txt, Separator) - 1), TheList)
End Function

Private Function GetPersonsScore(txt As String) As String
    ' Returns the person's score from the list item.
    GetPersonsScore = Mid(txt, InStr(1, txt, Separator) + Len(Separator))
End Function

Public Sub FileName(TheName As String)
    ' set the name of the file to save to
    TheFile = TheName
End Sub

Public Function DefaultFileName() As String
    ' a default file name you can use
    DefaultFileName = FixPath & "HighScores.ini"
End Function
