VERSION 5.00
Begin VB.Form boptions 
   BorderStyle     =   0  'None
   Caption         =   "Browser Options - WTMWVue"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "boptions.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00800000&
      Caption         =   "Save settings"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00800000&
      Caption         =   "Work Offline"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Work Online"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtHP 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "http://"
      Top             =   720
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800000&
      Caption         =   "Allow Popup Windows"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   3600
      Picture         =   "boptions.frx":3E71
      Top             =   2000
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3600
      Picture         =   "boptions.frx":573C
      Top             =   2000
      Width           =   525
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Page:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   4140
      Picture         =   "boptions.frx":6ECD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   3720
      Picture         =   "boptions.frx":8404
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Options"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   25
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   4140
      Picture         =   "boptions.frx":9928
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   3720
      Picture         =   "boptions.frx":ADF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "boptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HP As Integer
Dim AP As Integer
Dim WON As Integer
Dim WOF As Integer
Dim ERRL As String
Dim EL As Integer
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Form_Load()
ERRL = App.Path & "\data\" & "home.dat"
   On Error GoTo ELF
    HP = FreeFile

    Open App.Path & "\data\" & "home.dat" For Input As HP


    Do Until EOF(HP)
        Line Input #HP, Nextline$
        txtHP.Text = Nextline$
    Loop
    Close #HP
ERRL = App.Path & "\data\" & "id22.dat"
    On Error GoTo ELF
    AP = FreeFile
    
    Open App.Path & "\data\" & "id22.dat" For Input As AP
    Do Until EOF(AP)
        Line Input #AP, Nextline$
        Check1.Value = Nextline$
    Loop
    Close #AP
ELF:
MsgBox "The file " & ERRL & " was not found, please re-install the program. Thank you.", vbExclamation, "Missing file"
MsgBox "The program will now exit.", vbInformation, "Error has occurred"
Unload plist
Unload Me
EL = FreeFile
Open App.Path & "\Logs\" & "Errors.log" For Input As EL
Print #EL, ERRL & " was not found, please re-install."
Close #EL
End
End Sub

Private Sub Image2_Click()
Me.Hide
End Sub

Private Sub Image4_Click()
Me.WindowState = 1
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
    HP = FreeFile
    Open App.Path & "\data\" & "home.dat" For Append As HP
    Print #HP, txtHP.Text
    Close #HP
Me.Hide
    If Check2.Value Then
    AP = FreeFile
    Open App.Path & "\data\" & "id22.dat" For Append As AP
    Print #AP, Check1.Value
    Close #AP
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Webfrm.WebBrowser1.Offline = False
Option2.Value = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Webfrm.WebBrowser1.Offline = True
Option1.Value = False
End If
End Sub
