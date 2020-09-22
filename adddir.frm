VERSION 5.00
Begin VB.Form adddir 
   BorderStyle     =   0  'None
   Caption         =   "Add Directory - WTMWVue"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "adddir.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2715
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000009&
      Height          =   1200
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Directory"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   2160
      Picture         =   "adddir.frx":3644
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   2160
      Picture         =   "adddir.frx":4C0A
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   2560
      Picture         =   "adddir.frx":60D4
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   2560
      Picture         =   "adddir.frx":7705
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1200
      Picture         =   "adddir.frx":8C80
      Top             =   4080
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1200
      Picture         =   "adddir.frx":A54B
      Top             =   4080
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MP3's in this folder:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "adddir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        
        
        
        If Len(Dir1.Path) > 3 Then
            Form1.List1.AddItem Dir1.Path & "\" & File1.Filename
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        Form1.List1.AddItem Dir1.Path & File1.Filename
        End If
    Next tel
            Unload Me
Else
    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    Unload Me
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub


Private Sub Form_Load()

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Unload Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        
        
        
        If Len(Dir1.Path) > 3 Then
            plist.List1.AddItem Dir1.Path & "\" & File1.Filename
            plist.List2.AddItem File1.Filename
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        plist.List1.AddItem Dir1.Path & File1.Filename
        End If
    Next tel
            Unload Me
Else
    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    Unload Me
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub
