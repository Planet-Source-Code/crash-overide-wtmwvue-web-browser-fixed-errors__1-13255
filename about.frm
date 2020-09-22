VERSION 5.00
Begin VB.Form abouts 
   BorderStyle     =   0  'None
   Caption         =   "About - WTMWVue"
   ClientHeight    =   2400
   ClientLeft      =   7125
   ClientTop       =   1035
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   Picture         =   "about.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   4515
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Version + Build: Beta 1, Build 192"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "WTMWGaming"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "WTMWNetwork"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Websites:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"about.frx":3E71
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   3600
      Picture         =   "about.frx":3F86
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3600
      Picture         =   "about.frx":5851
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   280
      Left            =   3720
      Picture         =   "about.frx":6FE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   280
      Left            =   3720
      Picture         =   "about.frx":8506
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   280
      Left            =   4140
      Picture         =   "about.frx":9961
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   4140
      Picture         =   "about.frx":AE98
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "abouts"
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

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()
URL$ = "http://www.wtmw.net"
Webfrm.WebBrowser1.Navigate URL$
End Sub

Private Sub Label5_Click()
URL$ = "http://www.wtmwgaming.com"
Webfrm.WebBrowser1.nacigate URL$
End Sub
