Attribute VB_Name = "Module5"
'sound play

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long




Public Sub PlaySound(ByVal strFileName As String)
    Call sndPlaySound(strFileName, SND_ASYNC)
End Sub 'Note: You must use an "Unload". Otherwi
'     se, the music continues when the applica
'     tion is closed. Which could be used to t
'     he benefit of a prank or annoyance. lol.
'     .. To do that just add this line:


Public Sub stopsound(sound As String)
    temp = mciSendString("stop " & sound, 0&, 0, 0)
End Sub
'To make a "Play" button (I used a label
'     , use what you wish):



