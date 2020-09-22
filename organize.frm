VERSION 5.00
Begin VB.Form organize 
   BorderStyle     =   0  'None
   Caption         =   "Organize Favorites - WTMWVue"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   Picture         =   "organize.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000005&
      Height          =   2955
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image22 
      Height          =   300
      Left            =   720
      Picture         =   "organize.frx":3644
      Top             =   4080
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   720
      Picture         =   "organize.frx":4F0F
      Top             =   4080
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   315
      Left            =   2580
      Picture         =   "organize.frx":66A0
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   2580
      Picture         =   "organize.frx":7BD7
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   2160
      Picture         =   "organize.frx":90A3
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   2160
      Picture         =   "organize.frx":A5C7
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Organize Favorites"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   1935
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   240
      Picture         =   "organize.frx":BA22
      Top             =   480
      Width           =   825
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   240
      Picture         =   "organize.frx":D64E
      Top             =   480
      Width           =   825
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   1320
      Picture         =   "organize.frx":EEC7
      Top             =   4080
      Width           =   825
   End
   Begin VB.Image Image8 
      Height          =   300
      Left            =   1320
      Picture         =   "organize.frx":10B7D
      Top             =   4080
      Width           =   825
   End
End
Attribute VB_Name = "organize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FN As Integer
Dim FT As Integer
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   On Error Resume Next
    FN = FreeFile
    Open App.Path & "\data\" & "favorites.dat" For Input As FN


    Do Until EOF(FN)
        Line Input #FN, Nextline$
        List1.AddItem Nextline$
    Loop
    Close #FN
   On Error Resume Next
    FT = FreeFile
    Open App.Path & "\data\" & "favoritest.dat" For Input As FT


    Do Until EOF(FT)
        Line Input #FT, Nextline$
        List2.AddItem Nextline$
    Loop
    Close #FT
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = False
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Dim nEntryNum As Integer
nEntryNum = List2.ListCount
Do While nEntryNum > 0
    nEntryNum = nEntryNum - 1
    If List2.Selected(nEntryNum) = True Then
    List2.RemoveItem nEntryNum
    List1.RemoveItem nEntryNum
    End If
Loop
End Sub


Private Sub Image22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = False
End Sub

Private Sub Image22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = True
Kill App.Path & "\data\" & "favorites.dat"
Kill App.Path & "\data\" & "favoritest.dat"
    FN = FreeFile
    Open App.Path & "\data\" & "favorites.dat" For Append As FN
    Print #FN, List1.List(ListIndex)
    Close #FT
    FN = FreeFile
    Open App.Path & "\data\" & "favoritest.dat" For Append As FT
    Print #FT, List2.List(ListIndex)
    Close #FT
    Unload Me
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Me.WindowState = 1
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Unload Me
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub
