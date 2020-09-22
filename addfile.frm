VERSION 5.00
Begin VB.Form addfile 
   BorderStyle     =   0  'None
   Caption         =   "Add MP3 - WTMWVue"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "addfile.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   120
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ListBox Titlelist 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      ItemData        =   "addfile.frx":3644
      Left            =   120
      List            =   "addfile.frx":3646
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ListBox Pathlist 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   960
      ItemData        =   "addfile.frx":3648
      Left            =   240
      List            =   "addfile.frx":364A
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add MP3"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   75
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2640
      Picture         =   "addfile.frx":364C
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   2160
      Picture         =   "addfile.frx":4C7D
      Top             =   0
      Width           =   420
   End
   Begin VB.Image imgOK 
      Height          =   300
      Left            =   1080
      Picture         =   "addfile.frx":6243
      Top             =   4080
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   1080
      Picture         =   "addfile.frx":7B0E
      Top             =   4080
      Width           =   525
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   2160
      Picture         =   "addfile.frx":929F
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   2640
      Picture         =   "addfile.frx":A769
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "addfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Option Explicit
Dim mp3 As New ID3retr
Dim filnavn As String
Dim Z As Integer
Dim Box, Rounded As Long
Dim HoldNed As Boolean
Dim xPos, yPos As Integer



Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
err:
End Sub



Private Sub File1_DblClick()
On Error GoTo Ingenting
filnavn = Dir1.Path & "\" & File1.Filename
sjekkTittel (filnavn)
Pathlist.AddItem mp3.Filename
Ingenting:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim n As Integer
Titlelist.Visible = False
For n = 0 To Titlelist.ListCount - 1
Titlelist.ListIndex = n
Pathlist.ListIndex = n
plist.List2.AddItem Titlelist.Text
plist.List1.AddItem Pathlist.Text
Next
Unload Me
End If
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
Image4.Visible = True
Me.WindowState = 1
End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOK.Visible = False
End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOK.Visible = True
Dim n As Integer
Titlelist.Visible = False
For n = 0 To Titlelist.ListCount - 1
Titlelist.ListIndex = n
Pathlist.ListIndex = n
plist.List2.AddItem Titlelist.Text
plist.List1.AddItem Pathlist.Text
Next
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

Private Sub Titlelist_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Filindex As Integer
Select Case KeyCode
Case vbKeyDelete
Filindex = Titlelist.ListIndex
Titlelist.RemoveItem Filindex
Pathlist.RemoveItem Filindex
End Select
End Sub

