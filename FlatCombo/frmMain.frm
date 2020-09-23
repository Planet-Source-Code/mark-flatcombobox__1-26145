VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin Progetto1.UserControl1 UserControl3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      _ExtentX        =   8493
      _ExtentY        =   556
      Enable          =   0   'False
   End
   Begin Progetto1.UserControl1 UserControl2 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
   End
   Begin Progetto1.UserControl1 UserControl1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin VB.Label Label4 
      Caption         =   "mmark@tiscalinet.it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Dropped"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Enable"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Disable"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 30
  UserControl1.AddItem "Stringa testo prova " & i
  UserControl2.AddItem "Stringa testo prova " & i
  UserControl3.AddItem "Stringa testo prova " & i
Next i
End Sub
