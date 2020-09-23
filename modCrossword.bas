Attribute VB_Name = "modCrossword"
Option Explicit
'Constants
Private Const NumberFontName As String = "Trebuchet MS"
Private Const WordFontName As String = "Trebuchet MS"

'Global Variables
Public Crossword As CrosswordFile
Public currGridNumber As Integer

'File Definition
Public Enum eDifficultyLevel
    ForChildren = 0
    VeryEasy = 1
    Easy = 2
    Medium = 3
    Hard = 4
    VeryHard = 5
    ForGenii = 6
End Enum
Public Enum eOrientation
    Across = 0
    Down = 1
End Enum
Public Type Word
    X As Integer
    Y As Integer
    Answer As String
    Orientation As eOrientation
    Number As Integer
    Clue As String
    Visible As Boolean
End Type
Public Type cGrid
    Empty As Boolean
    Title As String
    Subject As String
    Difficulty As eDifficultyLevel
    GridSizeX As Integer
    GridSizeY As Integer
    Words() As Word
End Type
Public Type CrosswordFile
    Grids() As cGrid
    BeingEdited As Boolean
End Type
'Subs
Public Sub AddGrid()
    ReDim Preserve Crossword.Grids(UBound(Crossword.Grids) + 1)
    ResetGrid (UBound(Crossword.Grids))
End Sub
Public Sub RemoveGrid(ID As Integer)
    Dim i As Integer
    For i = 0 To UBound(Crossword.Grids) - 1
        Crossword.Grids(i) = Crossword.Grids(i + 1)
    Next
    'remove
    ReDim Preserve Crossword.Grids(UBound(Crossword.Grids) - 1)
End Sub
Private Sub ResetGrid(ID As Integer)
    ReDim Crossword.Grids(ID).Words(0)
    Crossword.Grids(ID).Title = "Crossword " & (ID + 1)
    Crossword.Grids(ID).Empty = True
    Crossword.Grids(ID).Difficulty = Easy
    Crossword.Grids(ID).Subject = "General"
    Crossword.Grids(ID).GridSizeX = 16
    Crossword.Grids(ID).GridSizeY = 16
End Sub
Public Sub ResetCrossword()
    'remove the words
    ReDim Crossword.Grids(0)
    ResetGrid 0
End Sub
Public Function GreaterOf(a As Integer, b As Integer) As Integer
    If a > b Then GreaterOf = a Else GreaterOf = b
End Function
Public Sub DrawEditMode(VGuide As Control, HGuide As Control, V As Control)
    VGuide.AutoRedraw = True
    VGuide.ScaleMode = vbPixels
    HGuide.AutoRedraw = True
    HGuide.ScaleMode = vbPixels
    VGuide.Cls
    HGuide.Cls
    Dim SquareWidth As Double
    Dim SquareHeight As Double
    SquareWidth = V.ScaleWidth / Crossword.Grids(currGridNumber).GridSizeX
    SquareHeight = V.ScaleHeight / Crossword.Grids(currGridNumber).GridSizeY
    Dim i As Integer
    'set fon
    HGuide.FontSize = 8
    HGuide.FontName = WordFontName
    VGuide.FontSize = 8
    VGuide.FontName = WordFontName
    'V Guide
    For i = 0 To Crossword.Grids(currGridNumber).GridSizeY
        VGuide.CurrentX = 0
        VGuide.CurrentY = (i * SquareHeight) + (SquareHeight / 2) - (UCase$(VGuide.TextHeight(i)) / 2)
        VGuide.Print i
    Next
    'H Guide
    For i = 0 To Crossword.Grids(currGridNumber).GridSizeX
        HGuide.CurrentX = (i * SquareWidth) + (SquareWidth / 2) - (UCase$(HGuide.TextWidth(i)) / 2)
        HGuide.CurrentY = 0
        HGuide.Print i
    Next
    'grid of red lines (H)
    For i = 0 To Crossword.Grids(currGridNumber).GridSizeY + 1
        V.Line (0, (i * SquareHeight))-(V.ScaleWidth, (i * SquareHeight)), vbRed, BF
    Next
    'grid of red lines (v)
    For i = 0 To Crossword.Grids(currGridNumber).GridSizeX + 1
        V.Line ((i * SquareWidth), 0)-((i * SquareWidth), V.ScaleHeight), vbRed, BF
    Next
End Sub
Public Sub ResizeCrossword(X As Integer, Y As Integer)
    'confirm
    Dim ans As Integer
    ans = MsgBox("Warning. Changing the size of the crossword may result in word loss. Continue?", vbYesNo + vbQuestion, "Confirmation Required")
    If ans = vbNo Then Exit Sub
    'delete Y/N array
    Dim Delete() As Boolean
    Dim DeleteO() As eOrientation
    Dim DeleteN() As Integer
    ReDim Delete(UBound(Crossword.Grids(currGridNumber).Words))
    ReDim DeleteO(UBound(Crossword.Grids(currGridNumber).Words))
    ReDim DeleteN(UBound(Crossword.Grids(currGridNumber).Words))
    Dim i As Integer
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        If Crossword.Grids(currGridNumber).Words(i).Orientation = Across And Crossword.Grids(currGridNumber).Words(i).X + Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1 > X Then Delete(i) = True: DeleteO(i) = Across: DeleteN(i) = Crossword.Grids(currGridNumber).Words(i).Number
        If Crossword.Grids(currGridNumber).Words(i).Orientation = Down And Crossword.Grids(currGridNumber).Words(i).Y + Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1 > Y Then Delete(i) = True: DeleteO(i) = Down: DeleteN(i) = Crossword.Grids(currGridNumber).Words(i).Number
    Next
    'delete all bad words
    For i = 0 To UBound(Delete)
        If Delete(i) Then RemoveWord DeleteO(i), DeleteN(i)
    Next
    'resize
    Crossword.Grids(currGridNumber).GridSizeX = X
    Crossword.Grids(currGridNumber).GridSizeY = Y
    DrawCrossword frmEdit.picView
End Sub
Private Sub UpdateNumbers()
    If Crossword.Grids(currGridNumber).Empty Then Exit Sub
    Dim i As Integer, j As Integer
    Dim Num As Integer
    Dim ChangedOne As Boolean
    Num = 1
    For j = 0 To Crossword.Grids(currGridNumber).GridSizeY
        For i = 0 To Crossword.Grids(currGridNumber).GridSizeX
            ChangedOne = False
            If GetWordIndex(i, j, Across) <> -1 Then Crossword.Grids(currGridNumber).Words(GetWordIndex(i, j, Across)).Number = Num: ChangedOne = True
            If GetWordIndex(i, j, Down) <> -1 Then Crossword.Grids(currGridNumber).Words(GetWordIndex(i, j, Down)).Number = Num: ChangedOne = True
            If ChangedOne Then Num = Num + 1
        Next
    Next
End Sub
Public Sub OpenFile(f As String)
    'Files are saved as ASCII.
    ResetCrossword
    Dim FF As String
    Dim FN As Integer
    FN = FreeFile
    Open f For Input As #FN
        Dim i As Integer
        Dim j As Integer
        Line Input #FN, FF: ReDim Crossword.Grids(FF)
        For i = 0 To UBound(Crossword.Grids)
            Line Input #FN, FF: Crossword.Grids(i).Difficulty = FF
            Line Input #FN, FF:  Crossword.Grids(i).Empty = FF
            Line Input #FN, FF:  Crossword.Grids(i).GridSizeX = FF
            Line Input #FN, FF:  Crossword.Grids(i).GridSizeY = FF
            Line Input #FN, FF:  Crossword.Grids(i).Subject = FF
            Line Input #FN, FF:  Crossword.Grids(i).Title = FF
            Line Input #FN, FF:  ReDim Preserve Crossword.Grids(i).Words(FF)
            For j = 0 To UBound(Crossword.Grids(i).Words)
                Line Input #FN, FF: Crossword.Grids(i).Words(j).Answer = FF
                Line Input #FN, FF: Crossword.Grids(i).Words(j).Clue = FF
                Line Input #FN, FF: Crossword.Grids(i).Words(j).Orientation = FF
                Line Input #FN, FF: Crossword.Grids(i).Words(j).X = FF
                Line Input #FN, FF: Crossword.Grids(i).Words(j).Y = FF
            Next
        Next
    Close #FN
End Sub
Public Sub DrawCrossword(V As Control)
    UpdateNumbers
    V.Cls
    'set up the control to support our methods
    V.AutoRedraw = True
    V.ScaleMode = vbPixels
    Dim i As Integer
    Dim j As Integer
    Dim SquareWidth As Double
    Dim SquareHeight As Double
    Dim TargX1 As Double
    Dim TargY1 As Double
    Dim TargX2 As Double
    Dim TargY2 As Double
    'crosswords are stretched to size of control. Get square sizes...
    SquareWidth = V.ScaleWidth / Crossword.Grids(currGridNumber).GridSizeX
    SquareHeight = V.ScaleHeight / Crossword.Grids(currGridNumber).GridSizeY
    'draw the black background square
    V.Line (0, 0)-(V.ScaleWidth, V.ScaleHeight), vbBlack, BF
    'exit if no words
    If Crossword.Grids(0).Empty Then GoTo EmptyCW
    'Now, for every word, draw it.
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        'For each letter...
        For j = 0 To Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1
            'DRAW THE WHITE SQUARE-----------------------------------------
            'across word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Across Then
                TargX1 = ((Crossword.Grids(currGridNumber).Words(i).X + j) * SquareWidth) + 1
                TargY1 = (Crossword.Grids(currGridNumber).Words(i).Y * SquareHeight) + 1
            End If
            'down word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Down Then
                TargX1 = ((Crossword.Grids(currGridNumber).Words(i).X) * SquareWidth) + 1
                TargY1 = ((Crossword.Grids(currGridNumber).Words(i).Y + j) * SquareHeight) + 1
            End If
            'common
            TargX2 = TargX1 + SquareWidth - 2
            TargY2 = TargY1 + SquareHeight - 2
            'draw the white square
            V.Line (TargX1, TargY1)-(TargX2, TargY2), vbWhite, BF

        Next
    Next
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        'For each letter...
        For j = 0 To Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1
            'DRAW THE WHITE SQUARE-----------------------------------------
            'across word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Across Then
                TargX1 = ((Crossword.Grids(currGridNumber).Words(i).X + j) * SquareWidth) + 1
                TargY1 = (Crossword.Grids(currGridNumber).Words(i).Y * SquareHeight) + 1
            End If
            'down word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Down Then
                TargX1 = ((Crossword.Grids(currGridNumber).Words(i).X) * SquareWidth) + 1
                TargY1 = ((Crossword.Grids(currGridNumber).Words(i).Y + j) * SquareHeight) + 1
            End If
            'common
            TargX2 = TargX1 + SquareWidth - 2
            TargY2 = TargY1 + SquareHeight - 2
            '--------------------------------------------------------------
            'if on the first letter, draw the number
            If j = 0 Then
                V.CurrentX = TargX1 + GetNumberIndent(SquareWidth)
                V.CurrentY = TargY1 + GetNumberIndent(SquareHeight)
                V.FontSize = GetNumberFontSize(SquareWidth)
                V.FontName = NumberFontName
                V.Print Crossword.Grids(currGridNumber).Words(i).Number
            End If
            'draw the letter if the word is visible or we are in edit mode
            If Crossword.Grids(currGridNumber).Words(i).Visible Or Crossword.BeingEdited Then
                V.FontSize = GetWordFontSize(SquareWidth)
                V.FontName = WordFontName
                V.CurrentX = TargX1 + (SquareWidth / 2) - (UCase$(V.TextWidth(Mid$(Crossword.Grids(currGridNumber).Words(i).Answer, j + 1, 1))) / 2)
                V.CurrentY = TargY1 + (SquareHeight / 2) - (UCase$(V.TextHeight(Mid$(Crossword.Grids(currGridNumber).Words(i).Answer, j + 1, 1))) / 2)
                V.Print UCase$(Mid$(Crossword.Grids(currGridNumber).Words(i).Answer, j + 1, 1))
            End If
        Next
    Next
EmptyCW:
    If Crossword.BeingEdited Then
        'non portable code
        DrawEditMode frmEdit.picNumbersDown, frmEdit.picNumbersAcross, frmEdit.picView
        frmEdit.txtTitle = Crossword.Grids(currGridNumber).Title
        frmEdit.txtSubject = Crossword.Grids(currGridNumber).Subject
        frmEdit.txtXSize = Crossword.Grids(currGridNumber).GridSizeX
        frmEdit.txtYSize = Crossword.Grids(currGridNumber).GridSizeY
        If Crossword.Grids(currGridNumber).Difficulty = ForChildren Then frmEdit.optDifficulty(0).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = VeryEasy Then frmEdit.optDifficulty(1).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = Easy Then frmEdit.optDifficulty(2).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = Medium Then frmEdit.optDifficulty(3).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = Hard Then frmEdit.optDifficulty(4).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = VeryHard Then frmEdit.optDifficulty(5).Value = True
        If Crossword.Grids(currGridNumber).Difficulty = ForGenii Then frmEdit.optDifficulty(6).Value = True
    End If
End Sub
Public Sub RemoveWord(Orientation As eOrientation, Number As Integer, Optional arrIndex As Integer = -1)
    Dim Index As Integer
    If arrIndex = -1 Then
        Dim i As Integer
        'Get the words index
        For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Orientation And Crossword.Grids(currGridNumber).Words(i).Number = Number Then Index = i
        Next
    Else
        i = arrIndex
    End If
    Dim j As Integer
    'reorganize, (erase)
    For j = i To UBound(Crossword.Grids(currGridNumber).Words) - 1
        Crossword.Grids(currGridNumber).Words(j) = Crossword.Grids(currGridNumber).Words(j + 1)
    Next
    'resize the array
    ReDim Preserve Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words) - 1)
End Sub
'Functions
Public Function AddWord(X As Integer, Y As Integer, Answer As String, Orientation As eOrientation, Clue As String, Visible As Boolean) As Boolean
    'CHECK TO SEE IF THE WORD WILL PHYSICALLY FIT-----------------------------------
    Dim OK As Boolean
    OK = True
    'validate co-ordinates
    If X < 0 Or Y < 0 Or X > Crossword.Grids(currGridNumber).GridSizeX Or Y > Crossword.Grids(currGridNumber).GridSizeY Then OK = False
    'validate word length
    If Orientation = Across Then If X + Len(Answer) > Crossword.Grids(currGridNumber).GridSizeX Then OK = False
    If Orientation = Down Then If Y + Len(Answer) > Crossword.Grids(currGridNumber).GridSizeY Then OK = False
    'return
    If OK = False Then
        MsgBox "The word '" & Answer & "' will not fit at (" & X & "," & Y & "). The word may be too long or the coordinates may be invalid.", vbCritical, "Error"
        Exit Function
    End If
    '-------------------------------------------------------------------------------
    'CHECK TO SEE IF THE WORDS LETTERS MATCH WHATS ALREADY THERE--------------------
    OK = True
    Dim i As Integer
    Dim checkX As Integer
    Dim checkY As Integer
    For i = 0 To Len(Answer) - 1
        'get coordinate of letter
        If Orientation = Across Then
            checkX = X + i
            checkY = Y
        Else
            checkX = X
            checkY = Y + i
        End If
        'check
        If UCase$(Mid$(Answer, i + 1, 1)) <> UCase$(GetSquareContents(checkX, checkY)) And Len(GetSquareContents(checkX, checkY)) = 1 Then
            OK = False
        End If
    Next
    'return
    If OK = False Then
        MsgBox "The word '" & Answer & "' will not fit at (" & X & "," & Y & "). The shared letters do not match.", vbCritical, "Error"
        Exit Function
    End If
    '------------------------------------------------------------------------------
    'CHECK TO SEE THAT THE PREVIOUS SQUARE IS BLANK IF IT EXISTS-------------------
    OK = True
    If Orientation = Across Then
        If X - 1 > 0 Then If Len(GetSquareContents(X - 1, Y)) <> 0 Then OK = False
    End If
    If Orientation = Down Then
        If Y - 1 > 0 Then If Len(GetSquareContents(X, Y - 1)) <> 0 Then OK = False
    End If
    If OK = False Then
        MsgBox "The word cannot be added. The previous square is not blank.", vbCritical, "Error"
        Exit Function
    End If
    '------------------------------------------------------------------------------
    'The word must be at least 3 chars long
    If Len(Answer) < 3 Then
        MsgBox "The word cannot be added. The minimum word length is 3.", vbCritical, "Error"
        Exit Function
    End If
    
    'CHECK to see that the word does not cross the start or end of any other words
    Dim PreStartOfWordX As Integer, PreStartOfWordY As Integer, PostEndOfWordX As Integer, PostEndOfWordY As Integer
    OK = True
    Dim Check As Boolean
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        If Orientation <> Crossword.Grids(currGridNumber).Words(i).Orientation Then
            Check = True
            'get co-ords
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Across Then
                If Crossword.Grids(currGridNumber).Words(i).X > 0 Then
                    PreStartOfWordX = Crossword.Grids(currGridNumber).Words(i).X - 1
                    PreStartOfWordY = Crossword.Grids(currGridNumber).Words(i).Y
                Else
                    Check = False
                End If
                If Crossword.Grids(currGridNumber).Words(i).X + Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1 < Crossword.Grids(currGridNumber).GridSizeX Then
                    PostEndOfWordX = Crossword.Grids(currGridNumber).Words(i).X + Len(Crossword.Grids(currGridNumber).Words(i).Answer)
                    PostEndOfWordY = Crossword.Grids(currGridNumber).Words(i).Y
                Else
                    Check = False
                End If
            Else
                If Crossword.Grids(currGridNumber).Words(i).Y > 0 Then
                    PreStartOfWordX = Crossword.Grids(currGridNumber).Words(i).X
                    PreStartOfWordY = Crossword.Grids(currGridNumber).Words(i).Y - 1
                Else
                    Check = False
                End If
                If Crossword.Grids(currGridNumber).Words(i).Y + Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1 < Crossword.Grids(currGridNumber).GridSizeY Then
                    PostEndOfWordX = Crossword.Grids(currGridNumber).Words(i).X
                    PostEndOfWordY = Crossword.Grids(currGridNumber).Words(i).Y + Len(Crossword.Grids(currGridNumber).Words(i).Answer)
                Else
                    Check = False
                End If
            End If
            'check.
            If Orientation = Across Then
                'other word is down, so compare its x value
                If Y = PreStartOfWordY And X <= PreStartOfWordX And X + Len(Answer) - 1 >= PostEndOfWordX Then OK = False
                If Y = PostEndOfWordY And X <= PreStartOfWordX And X + Len(Answer) - 1 >= PostEndOfWordX Then OK = False
            End If
            If Orientation = Down Then
                'other word is acros, so compare its y value
                If X = PreStartOfWordX And Y <= PreStartOfWordY And Y + Len(Answer) - 1 >= PostEndOfWordY Then OK = False
                If X = PostEndOfWordX And Y <= PreStartOfWordY And Y + Len(Answer) - 1 >= PostEndOfWordY Then OK = False
            End If
        End If
    Next
    If OK = False Then
        MsgBox "This word would interfere with another word.", vbCritical, "Error"
        Exit Function
    End If
    'To get here, the word is fine.
    'resize the words array if empty
    If Crossword.Grids(currGridNumber).Empty = False Then ReDim Preserve Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words) + 1)
    Crossword.Grids(currGridNumber).Empty = False
    'assign the properties
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).Answer = Answer
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).Clue = Clue
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).Orientation = Orientation
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).Visible = Visible
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).X = X
    Crossword.Grids(currGridNumber).Words(UBound(Crossword.Grids(currGridNumber).Words)).Y = Y
    AddWord = True
End Function
Public Function GetWordIndex(X As Integer, Y As Integer, Orientation As eOrientation) As Integer
    Dim i As Integer
    GetWordIndex = -1
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        If Crossword.Grids(currGridNumber).Words(i).Orientation = Orientation Then
            If Crossword.Grids(currGridNumber).Words(i).X = X And Crossword.Grids(currGridNumber).Words(i).Y = Y Then
                GetWordIndex = i
            End If
        End If
    Next
End Function
Private Function GetSquareContents(X As Integer, Y As Integer) As String
    GetSquareContents = vbNullString
    'make an array and fill it with all the letters
    Dim Letters() As String
    ReDim Letters(Crossword.Grids(currGridNumber).GridSizeX, Crossword.Grids(currGridNumber).GridSizeY)
    Dim i As Integer, j As Integer
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        For j = 0 To Len(Crossword.Grids(currGridNumber).Words(i).Answer) - 1
            'across word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Across Then Letters(Crossword.Grids(currGridNumber).Words(i).X + j, Crossword.Grids(currGridNumber).Words(i).Y) = UCase$(Mid$(Crossword.Grids(currGridNumber).Words(i).Answer, j + 1, 1))
            'Down word
            If Crossword.Grids(currGridNumber).Words(i).Orientation = Down Then Letters(Crossword.Grids(currGridNumber).Words(i).X, Crossword.Grids(currGridNumber).Words(i).Y + j) = UCase$(Mid$(Crossword.Grids(currGridNumber).Words(i).Answer, j + 1, 1))
        Next
    Next
    'return
    GetSquareContents = Letters(X, Y)
End Function
Private Function GetNumberIndent(Dimension As Double) As Double
    'Indent is 10%
    GetNumberIndent = Dimension * 0.02
End Function
Private Function GetNumberFontSize(SquareWidth As Double) As Integer
    'Fontsize is 10%
    GetNumberFontSize = SquareWidth * 0.15
End Function
Private Function GetWordFontSize(SquareWidth As Double) As Integer
    'Fontsize is 10%
    GetWordFontSize = SquareWidth * 0.4
End Function

