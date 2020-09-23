VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crossword Program"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":038A
   ScaleHeight     =   5460
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Crossword Player"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Crossword Editor"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "by Mike Payne 2004"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEdit_Click()
    frmEdit.Show
    Me.Hide
End Sub
Private Sub cmdPlay_Click()
    frmPlay.Show
    Me.Hide
End Sub
