VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ss As Integer

Private Sub Form_Load()
   On Error Resume Next
    ss = FreeFile
    Open "c:\settings.ini" For Input As ss

    Text1.Text = Input$
    Close #ss
End Sub
