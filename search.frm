VERSION 5.00
Begin VB.Form search 
   BorderStyle     =   0  'None
   Caption         =   "Search - WTMWVue"
   ClientHeight    =   900
   ClientLeft      =   1260
   ClientTop       =   1200
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   Picture         =   "search.frx":0000
   ScaleHeight     =   900
   ScaleWidth      =   4485
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Enter Search Here"
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   3840
      Picture         =   "search.frx":2D2D
      Top             =   480
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   3840
      Picture         =   "search.frx":4620
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Search - WTMWVue"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   3720
      Picture         =   "search.frx":5E96
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   3720
      Picture         =   "search.frx":73BA
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4130
      Picture         =   "search.frx":8815
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4130
      Picture         =   "search.frx":9D4C
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Command1_Click()
SRH$ = Text1.Text
Webfrm.WebBrowser1.Navigate "http://www.lycos.com/srch/?lpv=1&loc=searchhp&query=" & SRH$ & ""
Unload Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Unload Me
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Me.WindowState = 1
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
SRH$ = Text1.Text
Webfrm.WebBrowser1.Navigate "http://www.lycos.com/srch/?lpv=1&loc=meta_index&query=" & SRH$ & ""
Unload Me
End Sub

Private Sub Title_Click()
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub
