VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ID3retr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sFilename As String


Private Type Info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sComment As String * 30
    sYear As String * 4
    End Type
    Private MP3Info As Info


Public Property Get Filename() As String
  Filename = sFilename
End Property



Public Property Let Filename(ByVal sPassFilename As String)

    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 127
    ' Store the filename (for "Get Filename" property)
    sFilename = sPassFilename
    
    ' Clear the info variables
    
    MP3Info.sTitle = ""
    MP3Info.sArtist = ""
    MP3Info.sAlbum = ""
    MP3Info.sYear = ""
    MP3Info.sComment = ""
    
    ' Ensure the MP3 file exists
    On Error GoTo err:
    If Dir(sFilename) = "" Then Exit Property
    
    ' Retrieve the info data from the MP3
    
    iFreefile = FreeFile
    lFilePos = FileLen(sFilename) - 127
    Open sFilename For Binary As #iFreefile
    Get #iFreefile, lFilePos, sData
    Close #iFreefile
    
    ' Populate the info variables
    


    If Left(sData, 3) = "TAG" Then
        MP3Info.sTitle = Mid(sData, 4, 30)
        MP3Info.sArtist = Mid(sData, 34, 30)
        MP3Info.sAlbum = Mid(sData, 64, 30)
        MP3Info.sYear = Mid(sData, 94, 4)
        MP3Info.sComment = Mid(sData, 98, 30)
    End If
Exit Property
err:
End Property



Public Property Get title() As String

    title = RTrim(MP3Info.sTitle)
End Property



Public Property Get Artist() As String

    Artist = RTrim(MP3Info.sArtist)
End Property



Public Property Get Album() As String

    Album = RTrim(MP3Info.sAlbum)
End Property



Public Property Get Year() As String

    Year = MP3Info.sYear
End Property



Public Property Get Comment() As String

    Comment = RTrim(MP3Info.sComment)
End Property



