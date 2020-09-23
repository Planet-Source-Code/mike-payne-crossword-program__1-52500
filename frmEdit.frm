VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEdit 
   Caption         =   "Crossword Editor"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGridAdder 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   7920
      ScaleHeight     =   1095
      ScaleWidth      =   1695
      TabIndex        =   35
      Top             =   6360
      Width           =   1695
      Begin VB.CommandButton cmdRemoveGrid 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdAddGrid 
         Caption         =   "+"
         Height          =   375
         Left            =   960
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmbCrosswords 
         Height          =   315
         ItemData        =   "frmEdit.frx":038A
         Left            =   360
         List            =   "frmEdit.frx":038C
         TabIndex        =   36
         Top             =   0
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Picture         =   "frmEdit.frx":038E
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4680
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "cwd"
   End
   Begin VB.PictureBox picAddWord 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   5520
      Width           =   7815
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   5280
         TabIndex        =   39
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtWord 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtClue 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton optDown 
         Caption         =   "Down"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdOKAddWord 
         Caption         =   "Update"
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optAcross 
         Caption         =   "Across"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Word Editor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Word:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clue:"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   11
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   195
         Index           =   4
         Left            =   5880
         TabIndex        =   10
         Top             =   840
         Width           =   150
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   195
         Index           =   5
         Left            =   6840
         TabIndex        =   9
         Top             =   840
         Width           =   150
      End
   End
   Begin VB.PictureBox picNumbersDown 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picNumbersAcross 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   7455
      TabIndex        =   14
      Top             =   0
      Width           =   7455
   End
   Begin VB.PictureBox picView 
      Height          =   5055
      Left            =   480
      ScaleHeight     =   4995
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   615
         Left            =   3360
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox picCWEditor 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   8040
      ScaleHeight     =   6375
      ScaleWidth      =   1695
      TabIndex        =   16
      Top             =   -120
      Width           =   1695
      Begin VB.OptionButton optDifficulty 
         Caption         =   "For Genii"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   5640
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "Very Hard"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   5400
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "Hard"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   5160
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "Medium"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   4920
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "Easy"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   4680
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "Very Easy"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton optDifficulty 
         Caption         =   "For Kids"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdateProperties 
         Caption         =   "Update"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox txtYSize 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtXSize 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difficulty:"
         Height          =   195
         Index           =   11
         Left            =   450
         TabIndex        =   27
         Top             =   3840
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   25
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   23
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y Size:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X Size:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Crossword Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   290
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "&Refresh"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Refreshing As Boolean 'used to prevent an eternal loop when populating the grid list.
Private Sub cmbCrosswords_Change()
    'a valid grid MUST be selected
    If cmbCrosswords.ListIndex < 0 Then cmbCrosswords.ListIndex = 0
End Sub
Private Sub cmbCrosswords_Click()
    'change to the selected grid
    cmbCrosswords_Change 'make sure the selected number is valid
    currGridNumber = cmbCrosswords.ListIndex 'set the current number to the selected number
    If Refreshing = False Then RefreshView 'draw the crossword if we aren't already
End Sub
Private Sub cmdAddGrid_Click()
    'add a grid to the crossword
    AddGrid 'add the grid
    currGridNumber = UBound(Crossword.Grids) 'set the current number to the new number
    RefreshView 'draw the crossword
End Sub
Private Sub cmdDelete_Click()
    'Delete the current word
    'Does the word already exist?
    'get the orientation we are talking about
    Dim bo As eOrientation
    If optAcross.Value = True Then bo = Across Else bo = Down
    'get the index of the word that currently exists there (-1 if no word is there)
    Dim Index As Integer
    Index = GetWordIndex(txtX.Text, txtY.Text, bo)
    If Index <> -1 Then
        'if a word is currently there, remove it
        RemoveWord 0, 0, Index
        'blank the boxes
        txtWord.Text = ""
        txtClue.Text = ""
        optAcross.Value = True
        txtX.Text = 0
        txtY.Text = 0
    End If
    RefreshView 'draw the crossword
End Sub
Private Sub cmdOKAddWord_Click()
    'NON BLANK VALIDATION
    If Trim$(txtWord.Text) = "" Or Trim$(txtClue.Text) = "" Then
        MsgBox "The word and clue must be non-blank.", vbCritical, "Error"
        Exit Sub
    End If
    'NUMERIC VALIDATION
    Dim OK As Boolean
    OK = True
    If IsNumeric(txtX.Text) Then
        If txtX.Text < 0 Then OK = False
    Else
        OK = False
    End If
    If IsNumeric(txtY.Text) Then
        If txtY.Text < 0 Then OK = False
    Else
        OK = False
    End If
    'return
    If OK = False Then
        MsgBox "Please enter valid numbers for Number, X, and Y.", vbCritical, "Error"
        Exit Sub
    End If
    'Does the word already exist?
    Dim bo As eOrientation
    If optAcross.Value = True Then bo = Across Else bo = Down
    If GetWordIndex(txtX.Text, txtY.Text, bo) <> -1 Then
       'The word must be at least 3 chars long
        If Len(txtWord) < 3 Then
            MsgBox "The word cannot be added. The minimum word length is 3.", vbCritical, "Error"
            Exit Sub
        End If
        'just update the word
        Crossword.Grids(currGridNumber).Words(GetWordIndex(txtX.Text, txtY.Text, Across)).Answer = UCase$(txtWord)
        Crossword.Grids(currGridNumber).Words(GetWordIndex(txtX.Text, txtY.Text, Across)).Clue = txtClue
        'blank the boxes
        txtWord.Text = ""
        txtClue.Text = ""
        optAcross.Value = True
        txtX.Text = 0
        txtY.Text = 0
    Else
        'TRY TO ADD, BLANK BOXES IF SUCCESSFUL
        Dim Orientation As eOrientation
        If optAcross.Value = True Then Orientation = Across Else Orientation = Down
        If AddWord(txtX, txtY, txtWord, Orientation, txtClue, False) Then
            'blank the boxes
            txtWord.Text = ""
            txtClue.Text = ""
            optAcross.Value = True
            txtX.Text = 0
            txtY.Text = 0
        End If
    End If
    'Draw the crossword
    RefreshView
End Sub
Private Sub cmdRemoveGrid_Click()
    'remove current grid
    If UBound(Crossword.Grids) = 0 Then
        MsgBox "A file must contain at least one crossword.", vbCritical, "Error"
        Exit Sub
    End If
    RemoveGrid currGridNumber
    'change to the next grid down, (0 if on first grid)
    currGridNumber = GreaterOf(currGridNumber - 1, 0)
    RefreshView 'draw crossword
End Sub
Private Sub cmdUpdateProperties_Click()
    'NUMERIC VALIDATION
    Dim OK As Boolean
    OK = True
    If IsNumeric(txtXSize.Text) Then
        If txtXSize.Text < 5 Then OK = False
    Else
        OK = False
    End If
    If IsNumeric(txtYSize.Text) Then
        If txtYSize.Text < 5 Then OK = False
    Else
        OK = False
    End If
    'return
    If OK = False Then
        MsgBox "Please enter valid numbers for X and Y (>=5).", vbCritical, "Error"
        Exit Sub
    End If
    Crossword.Grids(currGridNumber).Title = txtTitle
    Crossword.Grids(currGridNumber).Subject = txtSubject
    If optDifficulty(0).Value Then Crossword.Grids(currGridNumber).Difficulty = ForChildren
    If optDifficulty(1).Value Then Crossword.Grids(currGridNumber).Difficulty = VeryEasy
    If optDifficulty(2).Value Then Crossword.Grids(currGridNumber).Difficulty = Easy
    If optDifficulty(3).Value Then Crossword.Grids(currGridNumber).Difficulty = Medium
    If optDifficulty(4).Value Then Crossword.Grids(currGridNumber).Difficulty = Hard
    If optDifficulty(5).Value Then Crossword.Grids(currGridNumber).Difficulty = VeryHard
    If optDifficulty(6).Value Then Crossword.Grids(currGridNumber).Difficulty = ForGenii
    'resize
    ResizeCrossword txtXSize, txtYSize
    'Draw the crossword
    RefreshView
End Sub
Private Sub Form_Load()
    'set up for use
    ResetCrossword
    Crossword.BeingEdited = True
End Sub
Private Sub RefreshView()
    'draw crossword
    Refreshing = True
    DrawCrossword picView
    'fill list of crosswords
    cmbCrosswords.Clear
    Dim i  As Integer
    For i = 0 To UBound(Crossword.Grids)
        cmbCrosswords.AddItem Crossword.Grids(i).Title
    Next
    cmbCrosswords.ListIndex = currGridNumber
    Refreshing = False
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'various positioning, and a redraw
    picView.Height = Me.ScaleHeight - 148
    picView.Width = Me.ScaleWidth - 152
    picAddWord.Top = picView.Top + picView.Height + 8
    picCWEditor.Left = picView.Left + picView.Width + 8
    picNumbersAcross.Width = picView.Width
    picNumbersDown.Height = picView.Height
    picGridAdder.Left = Me.ScaleWidth - picGridAdder.Width - 16
    picGridAdder.Top = picView.Height + 64 + picView.Top
    RefreshView
    Shape1.Visible = False 'if we dont do this, the shape does not match the size of the squares. try it.
End Sub
Private Sub mnuFileNew_Click()
    'ask for save confirmation, then create a new crossword
    Dim ans As Integer
    ans = MsgBox("Save changes to current crossword?", vbYesNoCancel + vbQuestion, "Confirmation Required")
    If ans = vbCancel Then Exit Sub
    If ans = vbYes Then mnuFileSave_Click
    ResetCrossword
    Crossword.BeingEdited = True
    RefreshView
End Sub
Private Sub mnuFileOpen_Click()
    On Error GoTo No
    dlg.ShowOpen
    'open the file
    OpenFile dlg.FileName
    RefreshView
No:
End Sub
Private Sub mnuFileSave_Click()
    On Error GoTo No
    dlg.ShowSave
    'save the file#
    SaveFile dlg.FileName
No:
End Sub
Private Sub SaveFile(f As String)
    'save all data, one bit at a time.
    Dim FN As Integer
    FN = FreeFile
    Open f For Output As #FN
        Dim i As Integer
        Dim j As Integer
        Print #FN, UBound(Crossword.Grids)
        For i = 0 To UBound(Crossword.Grids)
            Print #FN, Crossword.Grids(i).Difficulty
            Print #FN, Crossword.Grids(i).Empty
            Print #FN, Crossword.Grids(i).GridSizeX
            Print #FN, Crossword.Grids(i).GridSizeY
            Print #FN, Crossword.Grids(i).Subject
            Print #FN, Crossword.Grids(i).Title
            Print #FN, UBound(Crossword.Grids(i).Words)
            For j = 0 To UBound(Crossword.Grids(i).Words)
                Print #FN, Crossword.Grids(i).Words(j).Answer
                Print #FN, Crossword.Grids(i).Words(j).Clue
                Print #FN, Crossword.Grids(i).Words(j).Orientation
                Print #FN, Crossword.Grids(i).Words(j).X
                Print #FN, Crossword.Grids(i).Words(j).Y
            Next
        Next
    Close #FN
End Sub
Private Sub mnuHelp_Click()
    frmHelp.Show 0, Me
End Sub
Private Sub mnuRefresh_Click()
    RefreshView
End Sub
Private Sub picView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'calculate what square we are on
    Dim SquareWidth As Double
    Dim SquareHeight As Double
    SquareWidth = picView.ScaleWidth / Crossword.Grids(currGridNumber).GridSizeX
    SquareHeight = picView.ScaleHeight / Crossword.Grids(currGridNumber).GridSizeY
    Dim XSquare As Integer
    Dim YSquare As Integer
    Shape1.Move Int(X / SquareWidth) * SquareWidth, Int(Y / SquareHeight) * SquareHeight, SquareWidth, SquareHeight
    Shape1.Visible = True
    txtX.Text = Int(X / SquareWidth)
    txtY.Text = Int(Y / SquareHeight)
    'load word if exists
    Dim bo As eOrientation
    Dim Loaded As Boolean
    If GetWordIndex(txtX.Text, txtY.Text, Across) <> -1 Then bo = Across: Loaded = True
    If GetWordIndex(txtX.Text, txtY.Text, Down) <> -1 Then bo = Down: Loaded = True
    If Loaded Then
        txtWord.Text = Crossword.Grids(currGridNumber).Words(GetWordIndex(txtX.Text, txtY.Text, bo)).Answer
        txtClue.Text = Crossword.Grids(currGridNumber).Words(GetWordIndex(txtX.Text, txtY.Text, bo)).Clue
        If bo = Across Then
            optAcross.Value = True
        Else
            optDown.Value = True
        End If
    End If
End Sub

