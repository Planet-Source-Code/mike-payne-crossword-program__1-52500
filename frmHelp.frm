VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":038A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   435
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   5115
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Playing Crosswords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   165
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   1470
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":041E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1155
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   5115
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editing Crosswords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   165
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0644
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   795
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   5115
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The File Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   165
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Crossword Functionality provided by Mike Payne (guitarsidekick-win@guitarplaying.com)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   435
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   3675
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0726
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   915
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5115
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":08BE
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   555
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5115
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "This program was created by Mike Payne, 19th-20th March 2004. To contact me, use mike.payne@icdonline.co.uk. NO MAILING LISTS."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5115
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About The Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

