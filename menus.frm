VERSION 5.00
Begin VB.Form menus 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu txtsize 
      Caption         =   "txtsize"
      Begin VB.Menu smallest 
         Caption         =   "Smallest"
      End
      Begin VB.Menu smaller 
         Caption         =   "Smaller"
      End
      Begin VB.Menu medium 
         Caption         =   "Medium"
      End
      Begin VB.Menu larger 
         Caption         =   "Larger"
      End
      Begin VB.Menu largest 
         Caption         =   "Largest"
      End
   End
   Begin VB.Menu file 
      Caption         =   "&file"
      Begin VB.Menu nw 
         Caption         =   "New"
         Begin VB.Menu newin 
            Caption         =   "New Window"
         End
      End
      Begin VB.Menu oll 
         Caption         =   "&Open Link Location"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "option"
      Begin VB.Menu boption 
         Caption         =   "Browser Options"
      End
      Begin VB.Menu skdoak 
         Caption         =   "-"
      End
      Begin VB.Menu orgfav 
         Caption         =   "Favorites"
      End
      Begin VB.Menu scl 
         Caption         =   "Show Cool Links"
      End
      Begin VB.Menu mp3 
         Caption         =   "MP3 Player"
         Begin VB.Menu smp 
            Caption         =   "Show MP3 Player"
         End
         Begin VB.Menu hmp 
            Caption         =   "Hide MP3 Player"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "&help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu favorites 
      Caption         =   "favorites"
      Begin VB.Menu fvlst 
         Caption         =   "fvlist"
      End
   End
End
Attribute VB_Name = "menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
abouts.Show
End Sub

Private Sub boption_Click()
boptions.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub hmp_Click()
    Do Until Webfrm.WebBrowser1.Width >= 12000
        Webfrm.WebBrowser1.Width = Webfrm.WebBrowser1.Width + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
Webfrm.Picture4.Visible = False
Webfrm.Label7.Visible = False
Webfrm.Slider1.Visible = False
Webfrm.Image31.Visible = False
Webfrm.Image32.Visible = False
Webfrm.Image33.Visible = False
Webfrm.Image34.Visible = False
Webfrm.Image35.Visible = False
Webfrm.Image36.Visible = False
Webfrm.Image37.Visible = False
Webfrm.Image38.Visible = False
Webfrm.Image39.Visible = False
End Sub

Private Sub medium_Click()
Webfrm.WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
End Sub

Private Sub newin_Click()
Dim NewInstance As New FrmNew
Load NewInstance
NewInstance.Show
End Sub

Private Sub oll_Click()
location.Show
End Sub

Private Sub orgfav_Click()
organize.Show
End Sub

Private Sub quit_Click()
End
End Sub

Private Sub scl_Click()
If scl.Caption = "Show Cool Links" Then
Webfrm.showlinks.Visible = False
    Do Until Webfrm.WebBrowser1.Width <= 9500
        Webfrm.WebBrowser1.Width = Webfrm.WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
    Do Until Webfrm.WebBrowser1.Left >= 2400
        Webfrm.WebBrowser1.Left = Webfrm.WebBrowser1.Left + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
scl.Caption = "Hide Cool Links"
Else
Webfrm.showlinks.Visible = True
    Do Until Webfrm.WebBrowser1.Width >= 12000
        Webfrm.WebBrowser1.Width = Webfrm.WebBrowser1.Width + 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
    Do Until Webfrm.WebBrowser1.Left <= 0
        Webfrm.WebBrowser1.Left = Webfrm.WebBrowser1.Left - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
scl.Caption = "Show Cool Links"
End If
End Sub

Private Sub smaller_Click()
Webfrm.WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
End Sub

Private Sub smallest_Click()
Webfrm.WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
End Sub

Private Sub larger_Click()
Webfrm.WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
End Sub

Private Sub largest_Click()
Webfrm.WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
End Sub

Private Sub smp_Click()
Webfrm.Picture4.Visible = True
Webfrm.Label7.Visible = True
Webfrm.Slider1.Visible = True
Webfrm.Image31.Visible = True
Webfrm.Image32.Visible = True
Webfrm.Image33.Visible = True
Webfrm.Image34.Visible = True
Webfrm.Image35.Visible = True
Webfrm.Image36.Visible = True
Webfrm.Image37.Visible = True
Webfrm.Image38.Visible = True
Webfrm.Image39.Visible = True
    Do Until Webfrm.WebBrowser1.Width <= 10320
        Webfrm.WebBrowser1.Width = Webfrm.WebBrowser1.Width - 380
        For i = 1 To 1000
            DoEvents
        Next i
    Loop
End Sub
