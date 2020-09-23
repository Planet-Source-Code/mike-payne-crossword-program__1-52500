VERSION 5.00
Begin VB.Form frmGuess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guess"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optDown 
      Caption         =   "Down"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.OptionButton optAcross 
      Caption         =   "Across"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtGuess 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your guess:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmGuess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GuessedAcross As Boolean
Private Sub cmdOK_Click()
    'send the guess to frmPlay
    frmPlay.Guess = UCase$(txtGuess)
    If optAcross.Value And optAcross.Visible Then GuessedAcross = True Else GuessedAcross = False
    Unload Me
End Sub

