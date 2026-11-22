Attribute VB_Name = "modSound"
'==============================================================
' Grid Gobbler - retro sound effects
' Square-wave WAV data is generated in memory at startup and
' played asynchronously with PlaySound, so no files are needed.
'==============================================================
Option Explicit

Private Declare Function PlaySoundMem Lib "winmm.dll" Alias "PlaySoundA" _
    (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_MEMORY As Long = &H4

Public Enum SfxId
    sndGobble = 0
    sndWrong = 1
    sndZap = 2
    sndLevel = 3
    sndWarn = 4
    sndLife = 5
    sndStart = 6
    sndGameOver = 7
End Enum

Private Const SFX_COUNT As Long = 8
Private Const RATE As Long = 11025
Private Const AMP As Double = 34

Public gSoundOn As Boolean

Private mBuf() As Byte
Private mOfs(0 To SFX_COUNT - 1) As Long
Private mReady As Boolean

' Each effect is a list of  frequency:milliseconds  notes (0 = rest)
Public Sub InitSounds()
    Dim spec(0 To SFX_COUNT - 1) As String
    Dim i As Long, total As Long, p As Long

    spec(sndGobble) = "523:30,784:30,1047:50"
    spec(sndWrong) = "220:110,165:220"
    spec(sndZap) = "880:40,660:40,440:50,330:60,220:120"
    spec(sndLevel) = "523:80,659:80,784:80,1047:80,0:40,1047:220"
    spec(sndWarn) = "1400:40,0:40,1400:40"
    spec(sndLife) = "784:60,988:60,1175:60,1568:180"
    spec(sndStart) = "392:70,523:70,659:70,784:150"
    spec(sndGameOver) = "392:160,370:160,349:160,330:360"

    total = 0
    For i = 0 To SFX_COUNT - 1
        total = total + 44 + SampleCount(spec(i))
    Next
    ReDim mBuf(0 To total - 1)

    p = 0
    For i = 0 To SFX_COUNT - 1
        mOfs(i) = p
        WriteWave spec(i), p
    Next
    mReady = True
    gSoundOn = True
End Sub

Public Sub PlaySfx(ByVal id As SfxId)
    If Not gSoundOn Or Not mReady Then Exit Sub
    PlaySoundMem mBuf(mOfs(id)), 0&, SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
End Sub

Public Sub StopSfx()
    PlaySoundMem ByVal 0&, 0&, 0&
End Sub

Private Function NoteSamples(ByVal ms As Long) As Long
    NoteSamples = (((RATE * ms) \ 1000) \ 2) * 2
End Function

Private Function SampleCount(ByVal spec As String) As Long
    Dim notes() As String, parts() As String, i As Long
    notes = Split(spec, ",")
    For i = 0 To UBound(notes)
        parts = Split(notes(i), ":")
        SampleCount = SampleCount + NoteSamples(CLng(Val(parts(1))))
    Next
End Function

Private Sub WriteWave(ByVal spec As String, p As Long)
    Dim n As Long, notes() As String, parts() As String
    Dim i As Long, k As Long, ns As Long, v As Long
    Dim freq As Double, env As Double, ph As Double

    n = SampleCount(spec)
    PutStr p, "RIFF"
    PutLong p, 36 + n
    PutStr p, "WAVE"
    PutStr p, "fmt "
    PutLong p, 16          ' fmt chunk size
    PutWord p, 1           ' PCM
    PutWord p, 1           ' mono
    PutLong p, RATE        ' sample rate
    PutLong p, RATE        ' bytes per second
    PutWord p, 1           ' block align
    PutWord p, 8           ' bits per sample
    PutStr p, "data"
    PutLong p, n

    notes = Split(spec, ",")
    For i = 0 To UBound(notes)
        parts = Split(notes(i), ":")
        freq = Val(parts(0))
        ns = NoteSamples(CLng(Val(parts(1))))
        For k = 0 To ns - 1
            If freq <= 0 Then
                v = 128
            Else
                env = 1 - (k / ns) * 0.85
                ph = k * freq / RATE
                ph = ph - Int(ph)
                If ph < 0.5 Then
                    v = 128 + CLng(AMP * env)
                Else
                    v = 128 - CLng(AMP * env)
                End If
            End If
            mBuf(p) = CByte(v)
            p = p + 1
        Next
    Next
End Sub

Private Sub PutStr(p As Long, ByVal s As String)
    Dim i As Long
    For i = 1 To Len(s)
        mBuf(p) = Asc(Mid$(s, i, 1))
        p = p + 1
    Next
End Sub

Private Sub PutWord(p As Long, ByVal v As Long)
    mBuf(p) = v And &HFF&
    mBuf(p + 1) = (v \ &H100&) And &HFF&
    p = p + 2
End Sub

Private Sub PutLong(p As Long, ByVal v As Long)
    mBuf(p) = v And &HFF&
    mBuf(p + 1) = (v \ &H100&) And &HFF&
    mBuf(p + 2) = (v \ &H10000) And &HFF&
    mBuf(p + 3) = (v \ &H1000000) And &HFF&
    p = p + 4
End Sub
