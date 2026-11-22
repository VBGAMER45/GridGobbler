Attribute VB_Name = "modScores"
'==============================================================
' Grid Gobbler - Hall of Fame (top 10, saved next to the EXE)
'==============================================================
Option Explicit

Public Type HiScore
    PName As String
    Score As Long
    Level As Long
    GameMode As String
End Type

Public gHi(1 To 10) As HiScore

Private Function ScoreFile() As String
    Dim p As String
    p = App.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    ScoreFile = p & "gobbler.hi"
End Function

Private Sub SetDefaults()
    Dim names As Variant, i As Long
    names = Array("BOLT", "SPARKY", "PIXEL", "WIDGET", "GADGET", "CHIP", "SPROCKET", "DOODAD", "TINKER", "COGS")
    For i = 1 To 10
        gHi(i).PName = names(i - 1)
        gHi(i).Score = (11 - i) * 200
        gHi(i).Level = (12 - i) \ 2
        gHi(i).GameMode = "Mixed"
    Next
End Sub

Public Sub LoadScores()
    Dim f As Integer, ln As String, parts() As String, i As Long
    SetDefaults
    On Error GoTo Fail
    If Len(Dir$(ScoreFile())) = 0 Then Exit Sub
    f = FreeFile
    Open ScoreFile() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        parts = Split(ln, "|")
        If UBound(parts) >= 3 And i < 10 Then
            i = i + 1
            gHi(i).PName = parts(0)
            gHi(i).Score = CLng(Val(parts(1)))
            gHi(i).Level = CLng(Val(parts(2)))
            gHi(i).GameMode = parts(3)
        End If
    Loop
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Public Sub SaveScores()
    Dim f As Integer, i As Long
    On Error GoTo Fail
    f = FreeFile
    Open ScoreFile() For Output As #f
    For i = 1 To 10
        Print #f, gHi(i).PName & "|" & gHi(i).Score & "|" & gHi(i).Level & "|" & gHi(i).GameMode
    Next
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

' Returns the rank (1-10) a score would earn, or 0 if it does not qualify
Public Function ScoreRank(ByVal sc As Long) As Long
    Dim i As Long
    For i = 1 To 10
        If sc > gHi(i).Score Then
            ScoreRank = i
            Exit Function
        End If
    Next
End Function

Public Sub InsertScore(ByVal rank As Long, ByVal nm As String, ByVal sc As Long, ByVal lvl As Long, ByVal md As String)
    Dim i As Long
    If rank < 1 Or rank > 10 Then Exit Sub
    For i = 10 To rank + 1 Step -1
        gHi(i) = gHi(i - 1)
    Next
    gHi(rank).PName = nm
    gHi(rank).Score = sc
    gHi(rank).Level = lvl
    gHi(rank).GameMode = md
End Sub

'--------------------------------------------------------------
' Settings (difficulty + sound), saved next to the EXE
'--------------------------------------------------------------
Public Function AppFolder() As String
    Dim p As String
    p = App.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    AppFolder = p
End Function

Public Sub LoadSettings(diff As Long, soundOn As Boolean)
    Dim f As Integer, ln As String, p As Long
    diff = 1
    soundOn = True
    On Error GoTo Fail
    If Len(Dir$(AppFolder() & "gobbler.cfg")) = 0 Then Exit Sub
    f = FreeFile
    Open AppFolder() & "gobbler.cfg" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        p = InStr(ln, "=")
        If p > 1 Then
            Select Case LCase$(Trim$(Left$(ln, p - 1)))
            Case "difficulty"
                diff = CLng(Val(Mid$(ln, p + 1)))
            Case "sound"
                soundOn = (Val(Mid$(ln, p + 1)) <> 0)
            End Select
        End If
    Loop
    Close #f
    If diff < 0 Or diff > 2 Then diff = 1
    Exit Sub
Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Public Sub SaveSettings(ByVal diff As Long, ByVal soundOn As Boolean)
    Dim f As Integer
    On Error GoTo Fail
    f = FreeFile
    Open AppFolder() & "gobbler.cfg" For Output As #f
    Print #f, "difficulty=" & diff
    Print #f, "sound=" & IIf(soundOn, 1, 0)
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub
