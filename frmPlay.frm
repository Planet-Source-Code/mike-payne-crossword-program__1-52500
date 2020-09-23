VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPlay 
   AutoRedraw      =   -1  'True
   Caption         =   "Crossword Player"
   ClientHeight    =   9570
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   638
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picClues 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   629
      TabIndex        =   3
      Top             =   5040
      Width           =   9495
      Begin VB.TextBox txtClues 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmPlay.frx":038A
         Top             =   0
         Width           =   9285
      End
   End
   Begin VB.ComboBox cmbCrosswords 
      Height          =   315
      ItemData        =   "frmPlay.frx":0390
      Left            =   1200
      List            =   "frmPlay.frx":0392
      TabIndex        =   2
      Top             =   150
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4320
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "cwd"
   End
   Begin VB.PictureBox picView 
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   615
         Left            =   960
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select Puzzle:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
   End
   Begin VB.Menu mnuReset 
      Caption         =   "Reset"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Guess  As String 'used for guessing a word, lol
Private Sub cmbCrosswords_Change()
    'make sure a valid grid is selected
    If cmbCrosswords.ListIndex < 0 Then cmbCrosswords.ListIndex = 0
End Sub
Private Sub cmbCrosswords_Click()
    'validate then selected the specified grid
    cmbCrosswords_Change
    currGridNumber = cmbCrosswords.ListIndex
    RefreshView
End Sub
Private Sub RefreshView()
    'title
    Me.Caption = Crossword.Grids(currGridNumber).Title & " (" & Crossword.Grids(currGridNumber).Subject & ")"
    If Crossword.Grids(currGridNumber).Difficulty = ForChildren Then Me.Caption = Me.Caption & " - For Kids"
    If Crossword.Grids(currGridNumber).Difficulty = VeryEasy Then Me.Caption = Me.Caption & " - Very Easy"
    If Crossword.Grids(currGridNumber).Difficulty = Easy Then Me.Caption = Me.Caption & " - Easy"
    If Crossword.Grids(currGridNumber).Difficulty = Medium Then Me.Caption = Me.Caption & " - Medium"
    If Crossword.Grids(currGridNumber).Difficulty = Hard Then Me.Caption = Me.Caption & " - Hard"
    If Crossword.Grids(currGridNumber).Difficulty = VeryHard Then Me.Caption = Me.Caption & " - Very Hard"
    If Crossword.Grids(currGridNumber).Difficulty = ForGenii Then Me.Caption = Me.Caption & " - For Genii"
    
    
    Crossword.BeingEdited = False
    'draw the crossword
    DrawCrossword picView
    If Crossword.Grids(currGridNumber).Empty Then
        'if it's empty, quit now.
        txtClues.Text = "This crossword is empty."
        Exit Sub
    End If
    'fill the clues textbox
    txtClues.Text = "ACROSS" & vbNewLine
    Dim i As Integer
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        If Crossword.Grids(currGridNumber).Words(i).Orientation = Across Then
            txtClues.Text = txtClues.Text & Crossword.Grids(currGridNumber).Words(i).Number & " - " & Crossword.Grids(currGridNumber).Words(i).Clue

            txtClues.Text = txtClues.Text & vbNewLine
        End If
    Next
    txtClues.Text = txtClues.Text & vbNewLine
    'clues
    txtClues.Text = txtClues.Text & "DOWN" & vbNewLine
    For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
        If Crossword.Grids(currGridNumber).Words(i).Orientation = Down Then
            txtClues.Text = txtClues.Text & Crossword.Grids(currGridNumber).Words(i).Number & " - " & Crossword.Grids(currGridNumber).Words(i).Clue
            txtClues.Text = txtClues.Text & vbNewLine
        End If
    Next
End Sub
Private Sub Form_Load()
    ResetCrossword
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    picView.Height = Me.ScaleHeight - picClues.Height
    picView.Width = Me.ScaleWidth
    picClues.Width = Me.ScaleWidth
    picClues.Top = picView.Height
    txtClues.Width = picClues.Width - 14
    RefreshView
End Sub
Private Sub GetGuess()
    frmGuess.Show 1, Me
End Sub
Private Sub mnuFileOpen_Click()
    'Open a file
    On Error GoTo No
    dlg.ShowOpen
    ResetCrossword
    OpenFile dlg.FileName
    'populate grid list
    cmbCrosswords.Clear
    Dim i  As Integer
    For i = 0 To UBound(Crossword.Grids)
        cmbCrosswords.AddItem Crossword.Grids(i).Title
    Next
    cmbCrosswords.ListIndex = currGridNumber
No:
RefreshView
End Sub
Private Sub mnuHelp_Click()
    frmHelp.Show 0, Me
End Sub
Private Sub mnuReset_Click()
    'resets the crossword so you can play it again.
    Dim i As Integer, j As Integer
    For i = 0 To UBound(Crossword.Grids)
        For j = 0 To UBound(Crossword.Grids(i).Words)
            Crossword.Grids(i).Words(j).Visible = False
        Next
    Next
    RefreshView
End Sub
Private Sub picView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'find out what square we are on
    Dim SquareWidth As Double
    Dim SquareHeight As Double
    SquareWidth = picView.ScaleWidth / Crossword.Grids(currGridNumber).GridSizeX
    SquareHeight = picView.ScaleHeight / Crossword.Grids(currGridNumber).GridSizeY
    Dim XSquare As Integer
    Dim YSquare As Integer
    Shape1.Move Int(X / SquareWidth) * SquareWidth, Int(Y / SquareHeight) * SquareHeight, SquareWidth, SquareHeight
    Shape1.Visible = True
    'load word if exists
    Dim bo As eOrientation
    Dim Loaded As Boolean
    'if an across word exists, set up frmGuess to accept guessing of an across word
    If GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Across) <> -1 Then
        If Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Across)).Visible = False Then
            bo = Across
            Loaded = True
            frmGuess.optAcross.Visible = True
            frmGuess.optAcross.Value = True
            frmGuess.optAcross.Caption = Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Across)).Number & " Across - " & Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Across)).Clue
        End If
    End If
    'if a down word exists, set up frmGuess to accept guessing of a down word
    If GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Down) <> -1 Then
        If Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Down)).Visible = False Then
            bo = Down
            Loaded = True
            frmGuess.optDown.Caption = Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Down)).Number & " Down - " & Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), Down)).Clue
            frmGuess.optDown.Value = True
            frmGuess.optDown.Visible = True
        End If
    End If
    If Loaded Then
        'if we found a word, and its not already guessed...
        If Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), bo)).Visible = False Then
            'get the guess
            GetGuess
            If frmGuess.GuessedAcross Then bo = Across Else bo = Down
            'test the guess
            If Guess = UCase$(Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), bo)).Answer) Then
                MsgBox "Correct!", vbInformation, "Good!"
                Crossword.Grids(currGridNumber).Words(GetWordIndex(Int(X / SquareWidth), Int(Y / SquareHeight), bo)).Visible = True
            Else
                MsgBox "Your guess was incorrect!", vbCritical, "Too bad!"
            End If
            Guess = ""
        End If
        Shape1.Visible = False
        RefreshView
        'check to see if we've won
        Dim i As Integer
        Dim Won As Boolean
        Won = True
        For i = 0 To UBound(Crossword.Grids(currGridNumber).Words)
            If Crossword.Grids(currGridNumber).Words(i).Visible = False Then Won = False
        Next
        If Won Then MsgBox "You win!", vbInformation, "Well done!"
    End If
    
End Sub
