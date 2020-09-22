VERSION 5.00
Begin VB.Form location 
   BorderStyle     =   0  'None
   Caption         =   "Open Location - WTMWVue"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   Picture         =   "openlocation.frx":0000
   ScaleHeight     =   915
   ScaleWidth      =   4515
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "http://"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   3840
      Picture         =   "openlocation.frx":2D2D
      Top             =   460
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3840
      Picture         =   "openlocation.frx":45F8
      Top             =   450
      Width           =   525
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   195
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   3720
      Picture         =   "openlocation.frx":5D89
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   3720
      Picture         =   "openlocation.frx":72AD
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4160
      Picture         =   "openlocation.frx":8708
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4160
      Picture         =   "openlocation.frx":9C3F
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Location"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

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
Image4.Visible = True
Me.WindowState = 1
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
URL$ = Text1.Text
Webfrm.WebBrowser1.Navigate URL$
Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub
