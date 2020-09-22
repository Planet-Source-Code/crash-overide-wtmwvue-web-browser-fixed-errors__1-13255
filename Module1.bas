Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Sub sjekkTittel(fil As String)
Dim sArtist, sTittel As String
mp3.Filename = fil
sArtist = mp3.Artist
sTittel = mp3.Title
If sArtist = "" Then
    If sTittel = "" Then
      Titlelist.AddItem File1.Filename
      Exit Sub
    End If
Titlelist.AddItem "ukjent artist - " & sTittel
Exit Sub
End If
If sTittel = "" Then
Titlelist.AddItem sArtist & " - Track" & Z
Z = Z + 1
Exit Sub
End If
Titlelist.AddItem sArtist & " - " & sTittel
End Sub

Public Sub Pause(Duration As Double)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub Select_Text(TextBoxName As Variant)
    TextBoxName.SelStart = 0
    TextBoxName.SelLength = Len(TextBoxName.Text)
End Sub
