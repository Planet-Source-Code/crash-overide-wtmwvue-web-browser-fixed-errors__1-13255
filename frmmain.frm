VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Webfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Web Browser"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   12000
   Begin VB.ListBox lstfavstit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   3930
      Left            =   7200
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      Picture         =   "frmmain.frx":A985
      ScaleHeight     =   510
      ScaleWidth      =   3015
      TabIndex        =   9
      Top             =   875
      Width           =   3015
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&File"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Help"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   280
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Options"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   280
         Width           =   615
      End
      Begin VB.Image cmdstop 
         Height          =   300
         Left            =   1920
         Picture         =   "frmmain.frx":B01E
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   1920
         Picture         =   "frmmain.frx":CA02
         Top             =   0
         Width           =   600
      End
      Begin VB.Image cmdrefresh 
         Height          =   300
         Left            =   1320
         Picture         =   "frmmain.frx":E354
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   1320
         Picture         =   "frmmain.frx":FDBB
         Top             =   0
         Width           =   600
      End
      Begin VB.Image cmdback 
         Height          =   300
         Left            =   120
         Picture         =   "frmmain.frx":117B6
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   120
         Picture         =   "frmmain.frx":131AF
         Top             =   0
         Width           =   600
      End
      Begin VB.Image cmdforward 
         Height          =   300
         Left            =   720
         Picture         =   "frmmain.frx":14B15
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   720
         Picture         =   "frmmain.frx":164F9
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.ListBox lstFavs 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   3930
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   12000
      ExtentX         =   21167
      ExtentY         =   12303
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6975
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12303
      _Version        =   393217
      TextRTF         =   $"frmmain.frx":17E5C
   End
   Begin VB.PictureBox cmdgo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5040
      Picture         =   "frmmain.frx":17F3E
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   500
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5020
      Picture         =   "frmmain.frx":19557
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   500
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1680
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   100
      ScaleHeight     =   180
      ScaleWidth      =   9495
      TabIndex        =   6
      Top             =   60
      Width           =   9495
      Begin VB.Label Title 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome! - WTMWVue"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12120
      End
   End
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Text            =   "http://www.wtmwgaming.com"
      Top             =   520
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   960
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      Begin VB.Label sngtitle 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   6255
      Left            =   10560
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   11033
      _Version        =   393216
      Orientation     =   1
      TickStyle       =   3
   End
   Begin VB.Timer Timer3 
      Interval        =   60
      Left            =   11160
      Top             =   5280
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   10920
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.TextBox sngpath 
      Height          =   285
      Left            =   10560
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.OptionButton optres 
      Caption         =   "Option1"
      Height          =   255
      Left            =   10560
      TabIndex        =   20
      Top             =   2760
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image39 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":1AB7B
      Top             =   3240
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image38 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":1C5BF
      Top             =   3240
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image37 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":1DE0F
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image36 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":1F7DA
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image35 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":2110A
      Top             =   2520
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image34 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":22A49
      Top             =   2520
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image33 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":2401B
      Top             =   2160
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image32 
      Height          =   300
      Left            =   11040
      Picture         =   "frmmain.frx":25931
      Top             =   2160
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11040
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image31 
      Height          =   6940
      Left            =   10320
      Picture         =   "frmmain.frx":26E77
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   45
      Width           =   1695
   End
   Begin VB.Image Command4 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":2BB59
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image30 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":2D705
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Command3 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":2F085
      Top             =   1320
      Width           =   900
   End
   Begin VB.Image Image29 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":30C32
      Top             =   1320
      Width           =   900
   End
   Begin VB.Image Command2 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":325AD
      Top             =   1680
      Width           =   900
   End
   Begin VB.Image Command1 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":3427B
      Top             =   2040
      Width           =   900
   End
   Begin VB.Image Image18 
      Height          =   300
      Left            =   2880
      Picture         =   "frmmain.frx":35F10
      Top             =   2040
      Width           =   900
   End
   Begin VB.Image backhtml 
      Height          =   300
      Left            =   8280
      Picture         =   "frmmain.frx":379F5
      Top             =   8130
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Image backhtmlon 
      Height          =   300
      Left            =   8280
      Picture         =   "frmmain.frx":3A4E9
      Top             =   8130
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Image htmledit 
      Height          =   435
      Left            =   11040
      Picture         =   "frmmain.frx":3CDFC
      Top             =   8050
      Width           =   855
   End
   Begin VB.Image htmlediton 
      Height          =   435
      Left            =   11040
      Picture         =   "frmmain.frx":3EDAC
      Top             =   8050
      Width           =   855
   End
   Begin VB.Image blue 
      Height          =   2880
      Index           =   1
      Left            =   0
      Picture         =   "frmmain.frx":40B5A
      Stretch         =   -1  'True
      Top             =   5055
      Width           =   2730
   End
   Begin VB.Image hidelinks 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":41E95
      Top             =   4750
      Width           =   2730
   End
   Begin VB.Image Image28 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":44631
      Top             =   4750
      Width           =   2730
   End
   Begin VB.Image cool11 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":469CC
      Top             =   4450
      Width           =   2730
   End
   Begin VB.Image Image27 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":48FE4
      Top             =   4450
      Width           =   2730
   End
   Begin VB.Image cool10 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":4B193
      Top             =   4155
      Width           =   2730
   End
   Begin VB.Image Image26 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":4D77F
      Top             =   4150
      Width           =   2730
   End
   Begin VB.Image cool9 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":4F90E
      Top             =   3850
      Width           =   2730
   End
   Begin VB.Image Image25 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":51EA2
      Top             =   3850
      Width           =   2730
   End
   Begin VB.Image cool8 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":53FB6
      Top             =   3560
      Width           =   2730
   End
   Begin VB.Image Image24 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":56729
      Top             =   3540
      Width           =   2730
   End
   Begin VB.Image cool7 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":58AA4
      Top             =   3240
      Width           =   2730
   End
   Begin VB.Image Image23 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":5AE1E
      Top             =   3240
      Width           =   2730
   End
   Begin VB.Image cool6 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":5CCC8
      Top             =   2940
      Width           =   2730
   End
   Begin VB.Image Image22 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":5F5FC
      Top             =   2940
      Width           =   2730
   End
   Begin VB.Image cool5 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":61B9D
      Top             =   2640
      Width           =   2730
   End
   Begin VB.Image Image21 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":64389
      Top             =   2640
      Width           =   2730
   End
   Begin VB.Image cool4 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":66780
      Top             =   2360
      Width           =   2730
   End
   Begin VB.Image Image20 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":68F81
      Top             =   2340
      Width           =   2730
   End
   Begin VB.Image cool2 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":6B38E
      Top             =   1740
      Width           =   2730
   End
   Begin VB.Image showlinks 
      Height          =   300
      Left            =   9000
      Picture         =   "frmmain.frx":6DB3C
      Top             =   480
      Width           =   675
   End
   Begin VB.Image Image17 
      Height          =   300
      Left            =   9000
      Picture         =   "frmmain.frx":6F635
      Top             =   480
      Width           =   675
   End
   Begin VB.Image cool3 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":71055
      Top             =   2040
      Width           =   2730
   End
   Begin VB.Image Image16 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":73A06
      Top             =   2040
      Width           =   2730
   End
   Begin VB.Image Image15 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":76018
      Top             =   1740
      Width           =   2730
   End
   Begin VB.Image cool1 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":783BA
      Top             =   1440
      Width           =   2730
   End
   Begin VB.Image Image14 
      Height          =   300
      Left            =   0
      Picture         =   "frmmain.frx":7AB8F
      Top             =   1440
      Width           =   2730
   End
   Begin VB.Image Image13 
      Height          =   300
      Left            =   11280
      Picture         =   "frmmain.frx":7CF68
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image12 
      Height          =   300
      Left            =   11280
      Picture         =   "frmmain.frx":7E48C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image11 
      Height          =   285
      Left            =   11640
      Picture         =   "frmmain.frx":7F8E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   11640
      Picture         =   "frmmain.frx":80E1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image cmdtxsz 
      Height          =   285
      Left            =   8320
      Picture         =   "frmmain.frx":822EA
      Top             =   480
      Width           =   675
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   8320
      Picture         =   "frmmain.frx":83D8B
      Top             =   480
      Width           =   675
   End
   Begin VB.Image cmdhome 
      Height          =   285
      Left            =   7800
      Picture         =   "frmmain.frx":8573A
      Top             =   480
      Width           =   525
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   7800
      Picture         =   "frmmain.frx":8700B
      Top             =   480
      Width           =   525
   End
   Begin VB.Image cmdfav 
      Height          =   285
      Left            =   7120
      Picture         =   "frmmain.frx":8883B
      Top             =   480
      Width           =   675
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   7120
      Picture         =   "frmmain.frx":8A300
      Top             =   480
      Width           =   675
   End
   Begin VB.Image cmdsearch 
      Height          =   285
      Left            =   6600
      Picture         =   "frmmain.frx":8BD22
      Top             =   480
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   6600
      Picture         =   "frmmain.frx":8D615
      Top             =   480
      Width           =   525
   End
   Begin VB.Image cmdadd 
      Height          =   285
      Left            =   5400
      Picture         =   "frmmain.frx":8EE8B
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   5400
      Picture         =   "frmmain.frx":90E7C
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   520
      Width           =   495
   End
   Begin VB.Label BrwsStat 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   10695
   End
   Begin VB.Label Status 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Document Finished"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   10320
      TabIndex        =   3
      Top             =   520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Go To:"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   525
      Width           =   615
   End
   Begin VB.Image Image19 
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "frmmain.frx":92D71
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2730
   End
End
Attribute VB_Name = "Webfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim FN As Integer
    Dim allowpopup As Boolean
    Dim FT As Integer
    Dim ERRL As String
    Dim hmp As String
    Dim EL As Integer


Private Sub backhtml_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
backhtml.Visible = False
End Sub

Private Sub backhtml_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
backhtml.Visible = True
RichTextBox1.Visible = True
    Do Until WebBrowser1.Width <= 2900
        WebBrowser1.Width = WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
backhtml.Visible = False
backhtmlon.Visible = False
End Sub

Private Sub cmdadd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdadd.Visible = False
End Sub

Private Sub cmdadd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdadd.Visible = True
    FN = FreeFile
    Open App.Path & "\data\" & "favorites.dat" For Append As FN
    Print #FN, txtUrl.Text
    Close #FN
    FT = FreeFile
    Open App.Path & "\data\" & "favoritest.dat" For Append As FT
    Print #FT, WebBrowser1.LocationName
    Close #FT
End Sub

Private Sub cmdback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdback.Visible = False
End Sub

Private Sub cmdback_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdback.Visible = True
    On Error Resume Next
    WebBrowser1.GoBack
End Sub

Private Sub cmdfav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdfav.Visible = False
End Sub

Private Sub cmdfav_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdfav.Visible = True
   On Error Resume Next
    FN = FreeFile
    Open App.Path & "\data\" & "favorites.dat" For Input As FN
    lstFavs.Visible = True

    Do Until EOF(FN)
        Line Input #FN, Nextline$
        lstFavs.AddItem Nextline$
    Loop
    Close #FN
    FT = FreeFile
    Open App.Path & "\data\" & "favoritest.dat" For Input As FT
    lstfavstit.Visible = True
    Do Until EOF(FT)
        Line Input #FT, nextline1$
        lstfavstit.AddItem nextline1$
    Loop
    Close #FT
End Sub

Private Sub cmdforward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdforward.Visible = False
End Sub

Private Sub cmdforward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdforward.Visible = True
    On Error Resume Next
    WebBrowser1.GoForward
End Sub

Private Sub cmdgo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdgo.Visible = False
End Sub

Private Sub cmdgo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdgo.Visible = True
    URL$ = txtUrl.Text
    WebBrowser1.Navigate URL$
End Sub

Private Sub cmdReload_Click()

End Sub


Private Sub cmdhome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdhome.Visible = False
End Sub

Private Sub cmdhome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdhome.Visible = True
    WebBrowser1.GoHome
    Pause 2
    URL$ = WebBrowser1.LocationURL
    txtUrl.Text = URL$
End Sub

Private Sub cmdrefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdrefresh.Visible = False
End Sub

Private Sub cmdrefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdrefresh.Visible = True
    WebBrowser1.Refresh
End Sub

Private Sub cmdsearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdsearch.Visible = False
End Sub

Private Sub cmdsearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdsearch.Visible = True
search.Show
End Sub

Private Sub cmdstop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdstop.Visible = False
End Sub

Private Sub cmdstop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdstop.Visible = True
    WebBrowser1.Stop
End Sub

Private Sub cmdtxsz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdtxsz.Visible = False
End Sub

Private Sub cmdtxsz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdtxsz.Visible = True
menus.PopupMenu menus.txtsize
End Sub





Private Sub Command5_Click()
    Do Until WebBrowser1.Width <= 2900
        WebBrowser1.Width = WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
Command5.Visible = False
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Visible = False
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Visible = True
Open App.Path & "\preview.html" For Output As #1
Print #1, RichTextBox1.Text
Close #1
    Do Until WebBrowser1.Width >= 12000
        WebBrowser1.Width = WebBrowser1.Width + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
WebBrowser1.Navigate App.Path & "\preview.html"
backhtml.Visible = True
backhtmlon.Visible = True
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.Visible = False
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.Visible = True
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.Filename <> "" Then
    Open CommonDialog1.Filename For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Visible = False
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Visible = True
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowOpen
If CommonDialog1.Filename <> "" Then
    Open CommonDialog1.Filename For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    RichTextBox1.Text = alltext$
    Loop
    Close #1
End If

End Sub

Private Sub cool1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool1.Visible = False
End Sub

Private Sub cool1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool1.Visible = True
URL$ = "http://www.wtmw.net"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool10.Visible = False
End Sub

Private Sub cool10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool10.Visible = True
URL$ = "http://www.mamamedia.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool11.Visible = False
End Sub

Private Sub cool11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool11.Visible = True
URL$ = "http://www.gamesages.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool2.Visible = False
End Sub

Private Sub cool2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool2.Visible = True
URL$ = "http://www.wtmwgaming.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool3.Visible = False
End Sub

Private Sub cool3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool3.Visible = True
URL$ = "http://www.gokounetwork.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool4.Visible = False
End Sub

Private Sub cool4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool4.Visible = True
URL$ = "http://kissmyglass.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool5.Visible = False
End Sub

Private Sub cool5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool5.Visible = True
URL$ = "http://deathmetal.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool6.Visible = False
End Sub

Private Sub cool6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool6.Visible = True
URL$ = "http://planetsourcecode.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool7.Visible = False
End Sub

Private Sub cool7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool7.Visible = True
URL$ = "http://www.fox.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool8.Visible = False
End Sub

Private Sub cool8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool8.Visible = True
URL$ = "http://www.thesimpsons.com"
WebBrowser1.Navigate URL$
End Sub

Private Sub cool9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool9.Visible = False
End Sub

Private Sub cool9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cool9.Visible = True
URL$ = "http://www.gamers.com"
WebBrowser1.Navigate URL$
End Sub


Private Sub Form_Click()
lstfavstit.Visible = False
lstFavs.Visible = False
lstfavstit.Clear
lstFavs.Clear
End Sub

Private Sub Form_Load()
    HP = FreeFile
    Open App.Path & "\data\" & "home.dat" For Input As HP


    Do Until EOF(HP)
        Line Input #HP, Nextline$
        hmp = Nextline$
    Loop
    Close #HP
RichTextBox1.Text = "<HTML>" & vbCrLf & "<META DESCRIPTION=This page was auto generated by WTMWVue>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "Auto Generated Page</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>" & vbCrLf & ""
Set ShapeThePB = New clsTransForm

ShapeThePB.ShapeMe RGB(255, 255, 255), True, , Picture2
    URL$ = hmp
    txtUrl.Text = hmp
    WebBrowser1.Navigate URL$
End Sub




Private Sub hidelinks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
hidelinks.Visible = False
End Sub

Private Sub hidelinks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
hidelinks.Visible = True
showlinks.Visible = True
menus.scl.Caption = "Show Cool Links"
    Do Until WebBrowser1.Left <= 0
        WebBrowser1.Left = WebBrowser1.Left - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
    Do Until WebBrowser1.Width >= 12000
        WebBrowser1.Width = WebBrowser1.Width + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
End Sub

Private Sub htmledit_Click()
htmledit.Visible = False
RichTextBox1.Visible = True
    Do Until WebBrowser1.Width <= 2900
        WebBrowser1.Width = WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = False
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
End
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.Visible = False
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.Visible = True
Me.WindowState = 1
End Sub

Private Sub Image33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image33.Visible = False
End Sub

Private Sub Image33_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image33.Visible = True
If MediaPlayer1.Filename = "" Then
If plist.List2.ListCount = 0 Then
addfile.Show
Exit Sub
End If
plist.List2.ListIndex = 0
plist.List1.ListIndex = 0
MediaPlayer1.Filename = plist.List1.Text
sngtitle.Caption = plist.List2.Text '
Slider1.Max = MediaPlayer1.Duration
Slider1.Value = MediaPlayer1.CurrentPosition
End If
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.CurrentPosition = 0
Slider1.Max = MediaPlayer1.Duration
Slider1.Value = MediaPlayer1.CurrentPosition
Exit Sub
End If
Filindex = plist.List2.ListIndex
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration
Slider1.Value = MediaPlayer1.CurrentPosition
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Image35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image35.Visible = False
End Sub

Private Sub Image35_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image35.Visible = True
MediaPlayer1.Stop
End Sub

Private Sub Image37_Click()
Image37.Visible = False
MediaPlayer1.Pause
End Sub

Private Sub Image39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image39.Visible = False
End Sub

Private Sub Image39_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image39.Visible = True
plist.Show
End Sub

Private Sub Label3_Click()
menus.PopupMenu menus.file
End Sub

Private Sub Label4_Click()
menus.PopupMenu menus.help
End Sub

Private Sub Label5_Click()
menus.PopupMenu menus.Options
End Sub

Private Sub List2_Click()
sngtitle.Caption = List2.Text
End Sub




Private Sub Text1_Change()
    On Error Resume Next
        txtUrl.Text = URL$

    If KeyAscii = 13 Then
        URL$ = txtUrl.Text
        WebBrowser1.Navigate URL$
    End If
End Sub

Private Sub medium_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
End Sub



Private Sub lstfavstit_Click()
lstFavs.ListIndex = lstfavstit.ListIndex
    txtUrl.Text = lstFavs.List(lstFavs.ListIndex)
    URL$ = lstFavs.Text
    WebBrowser1.Navigate URL$
    lstFavs.Visible = False
    lstfavstit.Visible = False
    lstFavs.Clear
    lstfavstit.Clear
    Close #FN
    Close #FT
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
    If optres.Value = True Then
        On Error GoTo err:
        plist.List1.ListIndex = List1.ListIndex + 1
        plist.List2.ListIndex = List1.ListIndex
        sngpath.Text = plist.List1.Text
        MediaPlayer1.Filename = sngpath.Text
        Slider1.Value = MediaPlayer1.CurrentPosition
        Slider1.Max = MediaPlayer1.Duration
        MediaPlayer1.Play
    End If
err:
End Sub

Private Sub RichTextBox1_Change()
RichTextBox1.RightMargin = RichTextBox1.Width
End Sub

Private Sub showlinks_Click()
showlinks.Visible = False
menus.scl.Caption = "Hide Cool Links"
    Do Until WebBrowser1.Width <= 9500
        WebBrowser1.Width = WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
    Do Until WebBrowser1.Left >= 2400
        WebBrowser1.Left = WebBrowser1.Left + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
End Sub

Private Sub Slider1_Click()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Timer1_Timer()
Title.Left = Title.Left - 70
sngtitle.Left = sngtitle.Left - 70
If Title.Left < -Title.Width Then
If sngtitle.Left < -sngtitle.Width Then
Title.Left = Picture1.ScaleWidth
sngtitle.Left = Picture4.ScaleWidth
End If
End If
End Sub

Private Sub Timer2_Timer()
Label6.Caption = Now
End Sub

Private Sub Timer3_Timer()
tinseconden = MediaPlayer1.CurrentPosition
    Dim Min As Integer
    Dim Sec As Integer
    Min = tinseconden \ 60
    Sec = tinseconden - (Min * 60)
    If Sec = "-1" Then Sec = "00"
    Label7.Caption = "0" & Min & ":" & Sec & ""
Slider1.Value = MediaPlayer1.CurrentPosition
End Sub

Private Sub Timer4_Timer()
End Sub

Private Sub txtUrl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
WebBrowser1.Navigate txtUrl.Text
End If

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Status.Caption = "Document Finished"
End Sub

Private Sub WebBrowser1_DownloadBegin()
Status.Caption = "Downloading"
End Sub

Private Sub WebBrowser1_DownloadComplete()
Status.Caption = "Finished"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
txtUrl.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If boptions.Check1.Value Then
    Cancel = False
    Dim NewInstance As New FrmNew
    Load NewInstance
    NewInstance.Show
    NewInstance.WebBrowser1.Navigate LocationURL
Else
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
BrwsStat.Caption = Text
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
TTL$ = WebBrowser1.LocationName
Title.Caption = "" & TTL$ & " - WTMWVue"
Me.Caption = "" & TTL$ & " - WTMWVue"
End Sub

Private Sub Command1_Click()
    Do Until WebBrowser1.Width >= 12000
        WebBrowser1.Width = WebBrowser1.Width + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
htmledit.Visible = True
RichTextBox1.Visible = False
End Sub

