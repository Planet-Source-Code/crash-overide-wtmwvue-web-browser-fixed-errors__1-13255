VERSION 5.00
Begin VB.Form plist 
   BorderStyle     =   0  'None
   Caption         =   "Playlist - WTMWVue"
   ClientHeight    =   4500
   ClientLeft      =   7290
   ClientTop       =   1485
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "plist.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   3000
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000009&
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image12 
      Height          =   300
      Left            =   2280
      Picture         =   "plist.frx":3644
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   2280
      Picture         =   "plist.frx":5133
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   1560
      Picture         =   "plist.frx":6B59
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   1560
      Picture         =   "plist.frx":8680
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image8 
      Height          =   300
      Left            =   840
      Picture         =   "plist.frx":A0FA
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   840
      Picture         =   "plist.frx":BC0A
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   120
      Picture         =   "plist.frx":D658
      Top             =   4080
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   120
      Picture         =   "plist.frx":F15C
      Top             =   4080
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Playlist"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   2160
      Picture         =   "plist.frx":10B94
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   2160
      Picture         =   "plist.frx":1215A
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2580
      Picture         =   "plist.frx":13624
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   380
      Left            =   2640
      Picture         =   "plist.frx":14C55
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "plist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = False
End Sub

Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = True
Webfrm.MediaPlayer1.Stop
Webfrm.sngtitle.Caption = ""
List1.Clear
List2.Clear
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Me.Hide
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
Me.WindowState = 1
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
addfile.Show
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = True
adddir.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub

Private Sub List2_DblClick()
Filindex = List2.ListIndex
List1.ListIndex = Filindex
Webfrm.MediaPlayer1.Filename = List1.Text
Webfrm.sngtitle.Caption = List2.Text
Webfrm.Slider1.Max = Webfrm.MediaPlayer1.Duration
End Sub
