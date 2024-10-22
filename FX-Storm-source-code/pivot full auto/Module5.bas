Attribute VB_Name = "Module5"

'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before doing anything with this code

'Disclaimer: This is illegal if excuted on real people and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education pupose only.
'For more educational source codes please visit us http://www.digi77.com/training.html
'Dr Jeeni Founder of www.oman0.net & www.digi77.com wishes you good luck :).

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
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



