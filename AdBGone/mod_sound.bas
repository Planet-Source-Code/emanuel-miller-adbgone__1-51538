Attribute VB_Name = "mod_Sound"
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
      "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
      Long) As Long
   Const SND_SYNC = &H0
   Const SND_ASYNC = &H1
   Const SND_NODEFAULT = &H2
   Const SND_LOOP = &H8
   Const SND_NOSTOP = &H10


Public Function play_sound(file)

    On Error Resume Next
    
    Dim SoundName As String
    SoundName$ = file
    wFlags% = SND_ASYNC
    x = sndPlaySound(SoundName$, wFlags%)

End Function
