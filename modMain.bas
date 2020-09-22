Attribute VB_Name = "modMain"
'---------------------------------------------------------------
'
'      Rock Away
'
'      Written by Dirk Rijmenants 2005
'
'---------------------------------------------------------------

Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Global SoundBuffer As String

Public Scores(10, 2) As String

Sub BeginPlaySound(ByVal ResourceId As Integer)
SoundBuffer = StrConv(LoadResData(ResourceId, "Geluiden"), vbUnicode)
Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
End Sub


