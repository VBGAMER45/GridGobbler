VERSION 5.00
Begin VB.Form frmGame
   AutoRedraw      =   -1  'True
   BackColor       =   &H002C120E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Gobbler"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   BeginProperty Font
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   0  'Manual
   Begin VB.Timer tmrGame
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8
      Top             =   8
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' Grid Gobbler
' Steer Bolt the robot around the grid and gobble every square
' that fits the rule - but keep away from the Glitches!
'==============================================================
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

' ---- Layout (pixels) ----
Private Const SCR_W As Long = 640
Private Const SCR_H As Long = 480
Private Const COLS As Long = 6
Private Const ROWS As Long = 5
Private Const CELL_W As Long = 100
Private Const CELL_H As Long = 72
Private Const GRID_X As Long = 20
Private Const GRID_Y As Long = 64
Private Const HUD_Y As Long = 430

Private Const MAX_GLITCH As Long = 6
Private Const MAX_FLOAT As Long = 12

' ---- Game states ----
Private Const ST_TITLE As Long = 0
Private Const ST_PLAY As Long = 1
Private Const ST_MSG As Long = 2
Private Const ST_PAUSE As Long = 3
Private Const ST_SCORES As Long = 4
Private Const ST_NAME As Long = 5
Private Const ST_ABOUT As Long = 6

' ---- What happens when a message box is closed ----
Private Const AFTER_RESUME As Long = 0
Private Const AFTER_NEXTLEVEL As Long = 1
Private Const AFTER_GAMEOVER As Long = 2
Private Const AFTER_BONUSGO As Long = 3
Private Const AFTER_BONUSEND As Long = 4

' ---- Bonus rounds (every 3rd level) ----
Private Const BONUS_NONE As Long = 0
Private Const BONUS_DASH As Long = 1        ' grab batteries, dodge Zippers
Private Const BONUS_LIGHTNING As Long = 2   ' timed gobbling, no Glitches
Private Const BONUS_MEMORY As Long = 3      ' board flips face-down

' ---- Glitch kinds ----
Private Const GK_ZIPPER As Long = 0      ' runs straight across
Private Const GK_WOBBLER As Long = 1     ' wanders at random
Private Const GK_HUNTER As Long = 2      ' chases the player
Private Const GK_SCRAMBLER As Long = 3   ' refills empty squares

Private Type Glitch
    Active As Boolean
    Kind As Long
    X As Long
    Y As Long
    FromX As Long
    FromY As Long
    T As Single           ' slide progress 0..1
    Speed As Single       ' slide progress per tick
    DirX As Long
    DirY As Long
    MoveTimer As Long
    MoveEvery As Long     ' ticks between moves
    Life As Long          ' moves left before it wanders off
    Warn As Long          ' >0 = still waiting off the grid
    Leaving As Boolean
End Type

Private Type Floater
    Active As Boolean
    X As Single
    Y As Single
    Txt As String
    Clr As Long
    Life As Long
End Type

Private Type PowerUp
    Active As Boolean
    X As Long
    Y As Long
    Life As Long
End Type

' ---- State ----
Private mState As Long
Private mAfter As Long
Private mFrame As Long

Private mMsgTitle As String
Private mMsgLines(0 To 5) As String
Private mMsgCount As Long
Private mMsgColor As Long
Private mMsgDelay As Long

Private mMode As Long
Private mDiff As Long            ' 0 easy, 1 normal, 2 hard
Private mLevel As Long
Private mScore As Long
Private mLives As Long
Private mNextLife As Long
Private mStreak As Long          ' right answers in a row
Private mLevelMistakes As Long
Private mRight As Long
Private mWrong As Long
Private mRule As GameRule
Private mGrid(0 To COLS - 1, 0 To ROWS - 1) As GridCell
Private mFitText(0 To COLS - 1, 0 To ROWS - 1) As String
Private mFitSize(0 To COLS - 1, 0 To ROWS - 1) As Single

' Player (Bolt)
Private pX As Long, pY As Long
Private pFromX As Long, pFromY As Long
Private pT As Single
Private pFace As Long          ' 0=down 1=up 2=left 3=right
Private pChomp As Long
Private pHurt As Long

Private mGl(0 To MAX_GLITCH - 1) As Glitch
Private mFl(0 To MAX_FLOAT - 1) As Floater
Private mSpawnTimer As Long

' Freeze power-up
Private mPow As PowerUp
Private mPowTimer As Long
Private mFreeze As Long          ' ticks the Glitches stay frozen

' Bonus rounds
Private mBonus As Long
Private mBonusCount As Long
Private mBonusScore As Long
Private mBonusTotal As Long      ' length of the round (ms)
Private mBonusEnd As Long        ' GetTickCount value when it ends
Private mBonusShown As Long      ' ms left in the current phase (for the bar)
Private mBonusPhase As Long      ' length of the current phase (ms)
Private mMemHidden As Boolean
Private mMemReveal As Long       ' how long the board is shown (ms)
Private mMemRevealEnd As Long
Private mPauseAt As Long
Private mPractice As Boolean
Private mPracticeIdx As Long

Private mNameEntry As String
Private mNewRank As Long
Private mCustomCats As Long
Private mStarX(0 To 59) As Single
Private mStarY(0 To 59) As Single

'==============================================================
' Form events
'==============================================================
Private Sub Form_Load()
    Dim i As Long
    Randomize
    Me.ScaleMode = vbPixels
    SizeClient SCR_W, SCR_H
    Me.AutoRedraw = True
    Me.FontName = "Arial"
    Me.FontBold = True
    Me.FontTransparent = True
    Me.FillStyle = vbFSTransparent

    InitRules
    mCustomCats = LoadPacks(AppFolder() & "packs\")
    InitSounds
    LoadScores
    LoadSettings mDiff, gSoundOn

    For i = 0 To 59
        mStarX(i) = Rnd * SCR_W
        mStarY(i) = Rnd * SCR_H
    Next

    mState = ST_TITLE
    tmrGame.Interval = 50
    tmrGame.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrGame.Enabled = False
    StopSfx
    SaveSettings mDiff, gSoundOn
End Sub

' Make the inside of the window exactly w x h pixels and centre it
Private Sub SizeClient(ByVal w As Long, ByVal h As Long)
    Dim bw As Single, bh As Single, ww As Single, hh As Single
    bw = Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX
    bh = Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY
    ww = bw + w * Screen.TwipsPerPixelX
    hh = bh + h * Screen.TwipsPerPixelY
    Me.Move (Screen.Width - ww) / 2, (Screen.Height - hh) / 2, ww, hh
End Sub

Private Sub tmrGame_Timer()
    mFrame = mFrame + 1
    If pHurt > 0 Then pHurt = pHurt - 1

    Select Case mState
    Case ST_PLAY
        UpdatePlay
        UpdateFloaters
    Case ST_MSG
        If mMsgDelay > 0 Then mMsgDelay = mMsgDelay - 1
        If pChomp > 0 Then pChomp = pChomp - 1
        UpdateFloaters
    End Select

    Render
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyM And mState <> ST_NAME Then
        gSoundOn = Not gSoundOn
        If Not gSoundOn Then StopSfx
        Exit Sub
    End If

    Select Case mState
    Case ST_TITLE
        Select Case KeyCode
        Case vbKey1, vbKeyNumpad1: StartGame MODE_MATH
        Case vbKey2, vbKeyNumpad2: StartGame MODE_WORDS
        Case vbKey3, vbKeyNumpad3: StartGame MODE_MIXED
        Case vbKeyD, vbKeyRight
            mDiff = (mDiff + 1) Mod 3
        Case vbKeyLeft
            mDiff = (mDiff + 2) Mod 3
        Case vbKeyB
            StartPractice
        Case vbKeyH
            mNewRank = 0
            mState = ST_SCORES
        Case vbKeyA
            PlaySfx sndLevel
            mState = ST_ABOUT
        Case vbKeyEscape
            Unload Me
        End Select

    Case ST_PLAY
        Select Case KeyCode
        Case vbKeyLeft, vbKeyNumpad4: MovePlayer -1, 0, 2
        Case vbKeyRight, vbKeyNumpad6: MovePlayer 1, 0, 3
        Case vbKeyUp, vbKeyNumpad8: MovePlayer 0, -1, 1
        Case vbKeyDown, vbKeyNumpad2: MovePlayer 0, 1, 0
        Case vbKeySpace, vbKeyReturn: DoGobble
        Case vbKeyP, vbKeyEscape
            mPauseAt = GetTickCount
            mState = ST_PAUSE
        End Select

    Case ST_PAUSE
        Select Case KeyCode
        Case vbKeyP, vbKeySpace, vbKeyReturn
            ' push the bonus clock forward by however long we were paused
            mBonusEnd = mBonusEnd + (GetTickCount - mPauseAt)
            mMemRevealEnd = mMemRevealEnd + (GetTickCount - mPauseAt)
            mState = ST_PLAY
        Case vbKeyEscape
            StopSfx
            mBonus = BONUS_NONE
            mState = ST_TITLE
        End Select

    Case ST_MSG
        If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
            If mMsgDelay = 0 Then DismissMsg
        End If

    Case ST_SCORES
        If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
            mNewRank = 0
            mState = ST_TITLE
        End If

    Case ST_ABOUT
        If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
            mState = ST_TITLE
        End If

    Case ST_NAME
        Select Case KeyCode
        Case vbKeyReturn
            SaveName
        Case vbKeyBack
            If Len(mNameEntry) > 0 Then mNameEntry = Mid$(mNameEntry, 1, Len(mNameEntry) - 1)
        End Select
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mState = ST_NAME Then
        If KeyAscii >= 32 And KeyAscii <= 126 And KeyAscii <> 124 Then
            If Not (KeyAscii = 32 And Len(mNameEntry) = 0) Then
                If Len(mNameEntry) < 12 Then mNameEntry = mNameEntry & Chr$(KeyAscii)
            End If
        End If
    End If
    KeyAscii = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, bx As Single
    Select Case mState
    Case ST_TITLE
        If Y >= 312 And Y <= 348 Then
            For i = 0 To 2
                bx = 50 + i * 185
                If X >= bx And X <= bx + 170 Then
                    StartGame i + 1
                    Exit Sub
                End If
            Next
        ElseIf Y >= 356 And Y <= 378 Then
            For i = 0 To 2
                bx = 280 + i * 70
                If X >= bx - 32 And X <= bx + 32 Then mDiff = i
            Next
        End If
    Case ST_MSG
        If mMsgDelay = 0 Then DismissMsg
    Case ST_SCORES
        mNewRank = 0
        mState = ST_TITLE
    Case ST_ABOUT
        mState = ST_TITLE
    End Select
End Sub

'==============================================================
' Difficulty
'==============================================================
Private Function DiffName() As String
    Select Case mDiff
    Case 0: DiffName = "Easy"
    Case 2: DiffName = "Hard"
    Case Else: DiffName = "Normal"
    End Select
End Function

' Points are worth more on harder settings
Private Function DiffMult() As Long
    DiffMult = mDiff + 1
End Function

' How tough the rules are for the current level
Private Function RuleLevel() As Long
    Select Case mDiff
    Case 0: RuleLevel = (mLevel + 1) \ 2
    Case 2: RuleLevel = mLevel + 3
    Case Else: RuleLevel = mLevel
    End Select
End Function

'==============================================================
' Game flow
'==============================================================
Private Sub StartGame(ByVal gm As Long)
    mMode = gm
    mPractice = False
    mLevel = 1
    mScore = 0
    If mDiff = 0 Then mLives = 5 Else mLives = 3
    mNextLife = 2000 * DiffMult()
    mNewRank = 0
    mStreak = 0
    mRight = 0
    mWrong = 0
    mBonusCount = 0
    SaveSettings mDiff, gSoundOn
    ResetRuleHistory
    SetupLevel
End Sub

' Try the bonus rounds from the title screen (B key)
Private Sub StartPractice()
    mMode = MODE_MIXED
    mPractice = True
    mLevel = 2
    mScore = 0
    mLives = 3
    mNextLife = 2000000000
    mStreak = 0
    mPracticeIdx = (mPracticeIdx Mod 3) + 1
    StartBonus mPracticeIdx
End Sub

Private Function PickMode() As Long
    If mMode = MODE_MIXED Then
        If mBonus <> BONUS_NONE Then
            PickMode = RandRange(MODE_MATH, MODE_WORDS)
        ElseIf (mLevel Mod 2) = 0 Then
            PickMode = MODE_WORDS
        Else
            PickMode = MODE_MATH
        End If
    Else
        PickMode = mMode
    End If
End Function

Private Sub ResetBoardState()
    ClearGlitches
    ClearFloaters
    mFreeze = 0
    mPow.Active = False
    mPowTimer = 250 + RandRange(0, 200)
    pX = 2: pY = 2
    pFromX = 2: pFromY = 2
    pT = 1
    pFace = 0
    pChomp = 0
End Sub

Private Sub SetupLevel()
    mBonus = BONUS_NONE
    mRule = MakeRule(RuleLevel(), PickMode())
    FillGrid COLS * ROWS, RandRange(9, 12)
    ResetBoardState
    mLevelMistakes = 0
    mSpawnTimer = FirstSpawnDelay()
    PlaySfx sndStart
    ShowMsg "LEVEL " & mLevel, RGB(255, 210, 50), AFTER_RESUME, _
            "Find and gobble:", "*" & mRule.Title, RuleHint(mRule)
End Sub

' Fill nFilled squares (the rest stay empty); nGood of them fit the rule
Private Sub FillGrid(ByVal nFilled As Long, ByVal nGood As Long)
    Dim order(0 To 29) As Long
    Dim total As Long, i As Long, j As Long, tmp As Long
    Dim x As Long, y As Long, tries As Long

    total = COLS * ROWS
    If mRule.Kind = RK_CATEGORY Then
        If nGood > CategorySize(mRule.Cat) Then nGood = CategorySize(mRule.Cat)
    End If
    For i = 0 To total - 1
        order(i) = i
        mGrid(i Mod COLS, i \ COLS).Filled = False
        mGrid(i Mod COLS, i \ COLS).Txt = ""
        mFitText(i Mod COLS, i \ COLS) = ""
    Next
    ' shuffle the squares; the first nGood get right answers
    For i = total - 1 To 1 Step -1
        j = RandRange(0, i)
        tmp = order(i): order(i) = order(j): order(j) = tmp
    Next
    For i = 0 To nFilled - 1
        x = order(i) Mod COLS
        y = order(i) \ COLS
        For tries = 1 To 20
            MakeCell mRule, (i < nGood), mGrid(x, y)
            If Not TextUsed(mGrid(x, y).Txt, order(i)) Then Exit For
        Next
    Next
End Sub

Private Function TextUsed(ByVal s As String, ByVal skipIdx As Long) As Boolean
    Dim k As Long
    For k = 0 To COLS * ROWS - 1
        If k <> skipIdx Then
            If mGrid(k Mod COLS, k \ COLS).Filled Then
                If mGrid(k Mod COLS, k \ COLS).Txt = s Then
                    TextUsed = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function CountGood() As Long
    Dim x As Long, y As Long
    For y = 0 To ROWS - 1
        For x = 0 To COLS - 1
            If mGrid(x, y).Filled And mGrid(x, y).Good Then CountGood = CountGood + 1
        Next
    Next
End Function

Private Sub ShowMsg(ByVal msgTitle As String, ByVal clr As Long, ByVal afterAct As Long, ParamArray msgLines() As Variant)
    Dim i As Long
    mMsgTitle = msgTitle
    mMsgColor = clr
    mAfter = afterAct
    mMsgCount = 0
    For i = LBound(msgLines) To UBound(msgLines)
        If mMsgCount <= 5 Then
            mMsgLines(mMsgCount) = CStr(msgLines(i))
            mMsgCount = mMsgCount + 1
        End If
    Next
    mMsgDelay = 8
    mState = ST_MSG
End Sub

Private Sub DismissMsg()
    Select Case mAfter
    Case AFTER_NEXTLEVEL
        mLevel = mLevel + 1
        If ((mLevel - 1) Mod 3) = 0 Then
            mBonusCount = mBonusCount + 1
            StartBonus ((mBonusCount - 1) Mod 3) + 1
        Else
            SetupLevel
        End If
    Case AFTER_BONUSGO
        If mBonus = BONUS_MEMORY Then
            mMemHidden = False
            mMemRevealEnd = GetTickCount + mMemReveal
            mBonusEnd = mMemRevealEnd + mBonusTotal
        Else
            mBonusEnd = GetTickCount + mBonusTotal
        End If
        mState = ST_PLAY
    Case AFTER_BONUSEND
        mBonus = BONUS_NONE
        If mPractice Then
            mState = ST_TITLE
        Else
            SetupLevel
        End If
    Case AFTER_GAMEOVER
        mBonus = BONUS_NONE
        mNewRank = ScoreRank(mScore)
        If mNewRank > 0 Then
            mNameEntry = ""
            mState = ST_NAME
        Else
            mState = ST_SCORES
        End If
    Case Else
        mState = ST_PLAY
    End Select
End Sub

Private Sub SaveName()
    Dim nm As String
    nm = Trim$(mNameEntry)
    If Len(nm) = 0 Then nm = "BOLT"
    InsertScore mNewRank, nm, mScore, mLevel, ModeName(mMode) & " " & Mid$("ENH", mDiff + 1, 1)
    SaveScores
    mState = ST_SCORES
End Sub

'==============================================================
' Bonus rounds
'==============================================================
Private Function BonusName(ByVal kind As Long) As String
    Select Case kind
    Case BONUS_DASH: BonusName = "Battery Dash"
    Case BONUS_LIGHTNING: BonusName = "Lightning Round"
    Case Else: BonusName = "Memory Round"
    End Select
End Function

Private Sub StartBonus(ByVal kind As Long)
    Dim x As Long, y As Long
    mBonus = kind
    mBonusScore = 0
    mMemHidden = False
    ResetBoardState
    mPowTimer = 1000000

    Select Case kind
    Case BONUS_DASH
        For y = 0 To ROWS - 1
            For x = 0 To COLS - 1
                mGrid(x, y).Filled = Not (x = pX And y = pY)
                mGrid(x, y).Good = True
                mGrid(x, y).Txt = ""
            Next
        Next
        mRule.Kind = RK_CATEGORY
        mRule.Title = "Grab the batteries!"
        mBonusTotal = 20000
        mSpawnTimer = 12
        ShowMsg "BONUS ROUND", RGB(120, 220, 255), AFTER_BONUSGO, "*Battery Dash", _
                "Roll over the batteries to grab them - no need to press Space.", _
                "Zippers come thick and fast. Getting zapped ends the round,", _
                "but it does not cost a life."

    Case BONUS_LIGHTNING
        mRule = MakeRule(RuleLevel(), PickMode())
        FillGrid COLS * ROWS, RandRange(10, 13)
        mBonusTotal = 25000
        ShowMsg "BONUS ROUND", RGB(120, 220, 255), AFTER_BONUSGO, "*Lightning Round", _
                "No Glitches! Gobble as fast as you can:  " & mRule.Title, _
                RuleHint(mRule), _
                "A wrong answer costs 3 seconds instead of a life."

    Case Else
        mRule = MakeRule(RuleLevel(), PickMode())
        FillGrid 12, 6
        mBonusTotal = 20000
        Select Case mDiff
        Case 0: mMemReveal = 8000
        Case 2: mMemReveal = 4500
        Case Else: mMemReveal = 6000
        End Select
        ShowMsg "BONUS ROUND", RGB(120, 220, 255), AFTER_BONUSGO, "*Memory Round", _
                "The rule is:  " & mRule.Title, _
                "Study the board - then the squares flip face down!", _
                "Gobble the right ones from memory. One miss ends the round."
    End Select

    If kind = BONUS_MEMORY Then mBonusPhase = mMemReveal Else mBonusPhase = mBonusTotal
    mBonusShown = mBonusPhase
    PlaySfx sndLife
End Sub

Private Sub UpdateBonus()
    Dim nowT As Long
    nowT = GetTickCount
    If mBonus = BONUS_MEMORY And Not mMemHidden Then
        mBonusShown = mMemRevealEnd - nowT
        If mBonusShown <= 0 Then
            mMemHidden = True
            mBonusPhase = mBonusTotal
            mBonusShown = mBonusTotal
            PlaySfx sndWarn
        End If
        Exit Sub
    End If
    mBonusShown = mBonusEnd - nowT
    If mBonusShown <= 0 Then
        mBonusShown = 0
        EndBonus "TIME'S UP!"
    End If
End Sub

Private Sub BonusAdd(ByVal pts As Long)
    mBonusScore = mBonusScore + pts
    AddScore pts
End Sub

Private Sub EndBonus(ByVal msgTitle As String, Optional ByVal extra As String = "")
    ClearGlitches
    mMemHidden = False
    If Len(extra) > 0 Then
        ShowMsg msgTitle, RGB(120, 220, 255), AFTER_BONUSEND, extra, "Bonus round score:", "*" & mBonusScore & " points"
    Else
        PlaySfx sndLevel
        ShowMsg msgTitle, RGB(120, 220, 255), AFTER_BONUSEND, "Bonus round score:", "*" & mBonusScore & " points"
    End If
End Sub

'==============================================================
' Player actions
'==============================================================
Private Sub MovePlayer(ByVal dx As Long, ByVal dy As Long, ByVal face As Long)
    Dim nx As Long, ny As Long, pts As Long
    pFace = face
    nx = pX + dx
    ny = pY + dy
    If Not InGrid(nx, ny) Then Exit Sub
    pFromX = pX: pFromY = pY
    pX = nx: pY = ny
    pT = 0

    ' Battery Dash: batteries are picked up just by rolling over them
    If mBonus = BONUS_DASH Then
        If mGrid(pX, pY).Filled Then
            mGrid(pX, pY).Filled = False
            pts = 20 * DiffMult()
            PlaySfx sndGobble
            AddFloater "+" & pts, CellCX(pX), CellTop(pY) + 10, RGB(120, 255, 140)
            BonusAdd pts
            If CountGood() = 0 Then
                BonusAdd 500 * DiffMult()
                EndBonus "CLEAN SWEEP!"
            End If
        End If
        Exit Sub
    End If

    ' Snowflake: freezes every Glitch for a while
    If mPow.Active Then
        If mPow.X = pX And mPow.Y = pY Then
            mPow.Active = False
            mPowTimer = 450 + RandRange(0, 250)
            Select Case mDiff
            Case 0: mFreeze = 150
            Case 2: mFreeze = 80
            Case Else: mFreeze = 110
            End Select
            PlaySfx sndLife
            AddFloater "FREEZE!", CellCX(pX), CellTop(pY) + 10, RGB(150, 210, 255)
        End If
    End If
End Sub

Private Function ComboMult() As Long
    Dim cm As Long
    cm = 1 + (mStreak - 1) \ 4
    If cm < 1 Then cm = 1
    If cm > 4 Then cm = 4
    ComboMult = cm
End Function

Private Sub DoGobble()
    Dim pts As Long, msg As String, s As String
    pT = 1
    pChomp = 6
    If mBonus = BONUS_DASH Then Exit Sub
    If mBonus = BONUS_MEMORY And Not mMemHidden Then Exit Sub
    If Not mGrid(pX, pY).Filled Then Exit Sub

    If mGrid(pX, pY).Good Then
        mGrid(pX, pY).Filled = False
        PlaySfx sndGobble
        Select Case mBonus
        Case BONUS_LIGHTNING
            pts = 20 * DiffMult()
            AddFloater "+" & pts, CellCX(pX), CellTop(pY) + 10, RGB(120, 255, 140)
            BonusAdd pts
            If CountGood() = 0 Then FillGrid COLS * ROWS, RandRange(10, 13)

        Case BONUS_MEMORY
            pts = 50 * DiffMult()
            AddFloater "+" & pts, CellCX(pX), CellTop(pY) + 10, RGB(120, 255, 140)
            BonusAdd pts
            If CountGood() = 0 Then
                BonusAdd 300 * DiffMult()
                EndBonus "PERFECT MEMORY!"
            End If

        Case Else
            mRight = mRight + 1
            mStreak = mStreak + 1
            pts = 10 * mLevel * DiffMult() * ComboMult()
            s = "+" & pts
            If ComboMult() > 1 Then s = s & "  x" & ComboMult()
            AddFloater s, CellCX(pX), CellTop(pY) + 10, RGB(120, 255, 140)
            AddScore pts
            If CountGood() = 0 Then LevelClear
        End Select
    Else
        PlaySfx sndWrong
        msg = WrongMessage(mRule, mGrid(pX, pY))
        Select Case mBonus
        Case BONUS_LIGHTNING
            mBonusEnd = mBonusEnd - 3000
            AddFloater "-3 sec", CellCX(pX), CellTop(pY) + 10, RGB(255, 120, 90)
        Case BONUS_MEMORY
            EndBonus "OOPS!", msg
        Case Else
            mWrong = mWrong + 1
            mStreak = 0
            mLevelMistakes = mLevelMistakes + 1
            LoseLife "OOPS!", RGB(255, 130, 90), msg, "The rule is:  " & mRule.Title
        End Select
    End If
End Sub

Private Sub AddScore(ByVal pts As Long)
    mScore = mScore + pts
    Do While mScore >= mNextLife
        mNextLife = mNextLife + 2000 * DiffMult()
        If mLives < 6 Then
            mLives = mLives + 1
            PlaySfx sndLife
            AddFloater "1UP!", CellCX(pX), CellTop(pY) - 10, RGB(255, 220, 80)
        End If
    Loop
End Sub

Private Sub LevelClear()
    Dim bonus As Long, perfect As Long
    bonus = 50 * mLevel * DiffMult()
    ClearGlitches
    If mLevelMistakes = 0 Then perfect = 100 * DiffMult()
    AddScore bonus + perfect
    PlaySfx sndLevel
    If perfect > 0 Then
        ShowMsg "LEVEL " & mLevel & " CLEAR!", RGB(120, 255, 150), AFTER_NEXTLEVEL, _
                "Level bonus:  " & bonus & " points", "*Perfect level!  +" & perfect, "No mistakes and no zaps."
    Else
        ShowMsg "LEVEL " & mLevel & " CLEAR!", RGB(120, 255, 150), AFTER_NEXTLEVEL, _
                "Great gobbling!", "Level bonus:  " & bonus & " points"
    End If
End Sub

Private Sub PlayerHit(ByVal kind As Long)
    PlaySfx sndZap
    If mBonus = BONUS_DASH Then
        pHurt = 24
        EndBonus "ZAPPED!", "A Zipper ended the round - but you keep all your lives."
        Exit Sub
    End If
    mStreak = 0
    mLevelMistakes = mLevelMistakes + 1
    LoseLife "ZAPPED!", RGB(255, 90, 90), "A " & GlitchName(kind) & " caught Bolt!", "Steer clear of the Glitches."
End Sub

Private Sub LoseLife(ByVal msgTitle As String, ByVal clr As Long, ByVal line1 As String, ByVal line2 As String)
    Dim pct As Long
    mLives = mLives - 1
    pHurt = 24
    ClearGlitches
    mFreeze = 0
    mSpawnTimer = FirstSpawnDelay()
    If mLives <= 0 Then
        mLives = 0
        If mRight + mWrong > 0 Then pct = (mRight * 100) \ (mRight + mWrong)
        PlaySfx sndGameOver
        ShowMsg msgTitle, clr, AFTER_GAMEOVER, line1, "*GAME OVER", _
                "Right answers: " & mRight & "     Wrong: " & mWrong & "     Accuracy: " & pct & "%"
    Else
        ShowMsg msgTitle, clr, AFTER_RESUME, line1, line2
    End If
End Sub

'==============================================================
' Glitches
'==============================================================
Private Sub UpdatePlay()
    Dim i As Long
    If pT < 1 Then
        pT = pT + 0.34
        If pT > 1 Then pT = 1
    End If
    If pChomp > 0 Then pChomp = pChomp - 1

    If mBonus <> BONUS_NONE Then
        UpdateBonus
        If mState <> ST_PLAY Then Exit Sub
        If mBonus <> BONUS_DASH Then Exit Sub     ' no Glitches in these rounds
    End If

    ' snowflake power-up comes and goes
    If mPow.Active Then
        mPow.Life = mPow.Life - 1
        If mPow.Life <= 0 Then
            mPow.Active = False
            mPowTimer = 400 + RandRange(0, 250)
        End If
    Else
        mPowTimer = mPowTimer - 1
        If mPowTimer <= 0 Then SpawnPower
    End If

    If mFreeze > 0 Then
        mFreeze = mFreeze - 1
    Else
        mSpawnTimer = mSpawnTimer - 1
        If mSpawnTimer <= 0 Then
            If ActiveGlitches() < MaxGlitches() Then SpawnGlitch
            mSpawnTimer = SpawnInterval()
        End If
        For i = 0 To MAX_GLITCH - 1
            If mGl(i).Active Then UpdateGlitch i
        Next
    End If

    CheckCollisions
End Sub

Private Sub SpawnPower()
    Dim tries As Long, x As Long, y As Long
    For tries = 1 To 12
        x = RandRange(0, COLS - 1)
        y = RandRange(0, ROWS - 1)
        If Not mGrid(x, y).Filled Then
            If Not (x = pX And y = pY) Then
                mPow.Active = True
                mPow.X = x
                mPow.Y = y
                mPow.Life = 170
                Exit Sub
            End If
        End If
    Next
    mPowTimer = 60      ' no empty square yet - try again soon
End Sub

Private Function ActiveGlitches() As Long
    Dim i As Long
    For i = 0 To MAX_GLITCH - 1
        If mGl(i).Active Then ActiveGlitches = ActiveGlitches + 1
    Next
End Function

Private Function MaxGlitches() As Long
    Dim v As Long
    If mBonus = BONUS_DASH Then
        MaxGlitches = 3 + mDiff
        Exit Function
    End If
    v = 1 + mLevel \ 2
    If mDiff = 0 Then v = v - 1
    If mDiff = 2 Then v = v + 1
    If v < 1 Then v = 1
    If v > 4 + mDiff \ 2 Then v = 4 + mDiff \ 2
    MaxGlitches = v
End Function

Private Function SpawnInterval() As Long
    Dim v As Long
    If mBonus = BONUS_DASH Then
        SpawnInterval = 26 - mDiff * 5 + RandRange(0, 12)
        Exit Function
    End If
    v = 150 - mLevel * 10
    If v < 40 Then v = 40
    If mDiff = 0 Then v = (v * 3) \ 2
    If mDiff = 2 Then v = (v * 7) \ 10
    SpawnInterval = v + RandRange(0, 40)
End Function

Private Function FirstSpawnDelay() As Long
    Dim v As Long
    v = 120 - mLevel * 8
    If v < 40 Then v = 40
    If mDiff = 0 Then v = v + 60
    FirstSpawnDelay = v
End Function

Private Function GlitchMoveEvery(ByVal kind As Long) As Long
    Dim baseT As Long, v As Long
    If mBonus = BONUS_DASH Then
        GlitchMoveEvery = 8 - mDiff
        Exit Function
    End If
    baseT = 28 - mLevel * 2
    If mDiff = 0 Then baseT = baseT + 8
    If mDiff = 2 Then baseT = baseT - 5
    If baseT < 8 - mDiff Then baseT = 8 - mDiff
    Select Case kind
    Case GK_ZIPPER: v = (baseT * 6) \ 10
    Case GK_HUNTER: v = (baseT * 12) \ 10
    Case GK_SCRAMBLER: v = (baseT * 9) \ 10
    Case Else: v = baseT
    End Select
    If v < 4 Then v = 4
    GlitchMoveEvery = v
End Function

Private Function PickGlitchKind() As Long
    Dim hunterLvl As Long, scramLvl As Long
    If mBonus = BONUS_DASH Then
        PickGlitchKind = GK_ZIPPER
        Exit Function
    End If
    ' Hunters and Scramblers show up earlier on harder settings
    hunterLvl = 5 - mDiff * 2
    scramLvl = 6 - mDiff * 2
    If mLevel >= scramLvl And Rnd < 0.22 Then
        PickGlitchKind = GK_SCRAMBLER
    ElseIf mLevel >= hunterLvl And Rnd < 0.3 Then
        PickGlitchKind = GK_HUNTER
    ElseIf Rnd < 0.5 Then
        PickGlitchKind = GK_ZIPPER
    Else
        PickGlitchKind = GK_WOBBLER
    End If
End Function

Private Sub SpawnGlitch()
    Dim i As Long, tries As Long, st As Long
    For i = 0 To MAX_GLITCH - 1
        If Not mGl(i).Active Then Exit For
    Next
    If i >= MAX_GLITCH Then Exit Sub

    ' pick an edge to come in from (not right where Bolt is standing)
    For tries = 1 To 8
        Select Case RandRange(0, 3)
        Case 0
            mGl(i).X = -1: mGl(i).Y = RandRange(0, ROWS - 1): mGl(i).DirX = 1: mGl(i).DirY = 0
        Case 1
            mGl(i).X = COLS: mGl(i).Y = RandRange(0, ROWS - 1): mGl(i).DirX = -1: mGl(i).DirY = 0
        Case 2
            mGl(i).X = RandRange(0, COLS - 1): mGl(i).Y = -1: mGl(i).DirX = 0: mGl(i).DirY = 1
        Case Else
            mGl(i).X = RandRange(0, COLS - 1): mGl(i).Y = ROWS: mGl(i).DirX = 0: mGl(i).DirY = -1
        End Select
        If Not (mGl(i).X + mGl(i).DirX = pX And mGl(i).Y + mGl(i).DirY = pY) Then Exit For
    Next

    mGl(i).Active = True
    mGl(i).Kind = PickGlitchKind()
    mGl(i).Leaving = False
    mGl(i).FromX = mGl(i).X
    mGl(i).FromY = mGl(i).Y
    mGl(i).T = 1
    If mBonus = BONUS_DASH Then mGl(i).Warn = 16 Else mGl(i).Warn = 30
    mGl(i).MoveTimer = 0
    mGl(i).MoveEvery = GlitchMoveEvery(mGl(i).Kind)
    st = mGl(i).MoveEvery - 1
    If st > 5 Then st = 5
    If st < 2 Then st = 2
    mGl(i).Speed = 1 / st
    mGl(i).Life = 8 + RandRange(0, 5) + mLevel
    PlaySfx sndWarn
End Sub

Private Sub UpdateGlitch(ByVal i As Long)
    ' still flashing a warning at the edge?
    If mGl(i).Warn > 0 Then
        mGl(i).Warn = mGl(i).Warn - 1
        If mGl(i).Warn = 0 Then
            mGl(i).FromX = mGl(i).X
            mGl(i).FromY = mGl(i).Y
            mGl(i).X = mGl(i).X + mGl(i).DirX
            mGl(i).Y = mGl(i).Y + mGl(i).DirY
            mGl(i).T = 0
            mGl(i).MoveTimer = 0
        End If
        Exit Sub
    End If

    If mGl(i).T < 1 Then
        mGl(i).T = mGl(i).T + mGl(i).Speed
        If mGl(i).T >= 0.999 Then
            mGl(i).T = 1
            OnGlitchArrive i
            If Not mGl(i).Active Then Exit Sub
        End If
    End If

    mGl(i).MoveTimer = mGl(i).MoveTimer + 1
    If mGl(i).MoveTimer >= mGl(i).MoveEvery And mGl(i).T >= 1 Then
        mGl(i).MoveTimer = 0
        ChooseGlitchMove i
    End If
End Sub

Private Sub OnGlitchArrive(ByVal i As Long)
    Dim x As Long, y As Long
    x = mGl(i).X
    y = mGl(i).Y
    If Not InGrid(x, y) Then
        mGl(i).Active = False
        Exit Sub
    End If
    If mGl(i).Kind = GK_SCRAMBLER And Not mGl(i).Leaving Then
        If Not mGrid(x, y).Filled And Rnd < 0.7 Then
            If Not (mPow.Active And mPow.X = x And mPow.Y = y) Then
                MakeCell mRule, (Rnd < 0.5), mGrid(x, y)
            End If
        End If
    End If
End Sub

Private Sub ChooseGlitchMove(ByVal i As Long)
    Dim dx As Long, dy As Long, ddx As Long, ddy As Long

    If mGl(i).Kind = GK_ZIPPER Or mGl(i).Leaving Then
        dx = mGl(i).DirX
        dy = mGl(i).DirY
    Else
        mGl(i).Life = mGl(i).Life - 1
        If mGl(i).Life <= 0 Then
            ExitDirection mGl(i).X, mGl(i).Y, dx, dy
            mGl(i).Leaving = True
        ElseIf mGl(i).Kind = GK_HUNTER And Rnd < 0.7 Then
            ddx = pX - mGl(i).X
            ddy = pY - mGl(i).Y
            If Abs(ddx) > Abs(ddy) Or (Abs(ddx) = Abs(ddy) And Rnd < 0.5) Then
                dx = Sgn(ddx)
            Else
                dy = Sgn(ddy)
            End If
            If dx = 0 And dy = 0 Then RandomDir i, dx, dy
        Else
            RandomDir i, dx, dy
        End If
    End If

    If dx <> 0 Or dy <> 0 Then
        mGl(i).DirX = dx
        mGl(i).DirY = dy
    End If
    mGl(i).FromX = mGl(i).X
    mGl(i).FromY = mGl(i).Y
    mGl(i).X = mGl(i).X + dx
    mGl(i).Y = mGl(i).Y + dy
    mGl(i).T = 0
End Sub

Private Sub RandomDir(ByVal i As Long, dx As Long, dy As Long)
    Dim tries As Long, nx As Long, ny As Long
    For tries = 1 To 16
        dx = 0: dy = 0
        Select Case RandRange(0, 3)
        Case 0: dx = 1
        Case 1: dx = -1
        Case 2: dy = 1
        Case Else: dy = -1
        End Select
        nx = mGl(i).X + dx
        ny = mGl(i).Y + dy
        If InGrid(nx, ny) Then
            If Not GlitchAt(nx, ny, i) Then
                If tries > 8 Then Exit Sub
                If Not (dx = -mGl(i).DirX And dy = -mGl(i).DirY) Then Exit Sub
            End If
        End If
    Next
    dx = 0: dy = 0
End Sub

Private Sub ExitDirection(ByVal x As Long, ByVal y As Long, dx As Long, dy As Long)
    Dim dL As Long, dR As Long, dT As Long, dB As Long
    dL = x + 1: dR = COLS - x: dT = y + 1: dB = ROWS - y
    dx = 0: dy = 0
    If dL <= dR And dL <= dT And dL <= dB Then
        dx = -1
    ElseIf dR <= dT And dR <= dB Then
        dx = 1
    ElseIf dT <= dB Then
        dy = -1
    Else
        dy = 1
    End If
End Sub

Private Function GlitchAt(ByVal x As Long, ByVal y As Long, ByVal skipIdx As Long) As Boolean
    Dim k As Long
    For k = 0 To MAX_GLITCH - 1
        If k <> skipIdx Then
            If mGl(k).Active And mGl(k).Warn = 0 Then
                If mGl(k).X = x And mGl(k).Y = y Then
                    GlitchAt = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Sub CheckCollisions()
    Dim i As Long, rx As Single, ry As Single, gx As Single, gy As Single, pts As Long
    rx = pFromX + (pX - pFromX) * pT
    ry = pFromY + (pY - pFromY) * pT
    For i = 0 To MAX_GLITCH - 1
        If mGl(i).Active And mGl(i).Warn = 0 Then
            gx = mGl(i).FromX + (mGl(i).X - mGl(i).FromX) * mGl(i).T
            gy = mGl(i).FromY + (mGl(i).Y - mGl(i).FromY) * mGl(i).T
            If Abs(gx - rx) < 0.55 And Abs(gy - ry) < 0.55 Then
                If mFreeze > 0 And mBonus = BONUS_NONE Then
                    ' frozen Glitches shatter when Bolt bumps them
                    mGl(i).Active = False
                    pts = 25 * DiffMult()
                    PlaySfx sndGobble
                    AddFloater "ZAP!  +" & pts, CellCX(pX), CellTop(pY) + 10, RGB(150, 210, 255)
                    AddScore pts
                Else
                    PlayerHit mGl(i).Kind
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Private Sub ClearGlitches()
    Dim i As Long
    For i = 0 To MAX_GLITCH - 1
        mGl(i).Active = False
    Next
End Sub

Private Function InGrid(ByVal x As Long, ByVal y As Long) As Boolean
    InGrid = (x >= 0 And x < COLS And y >= 0 And y < ROWS)
End Function

Private Function CellCX(ByVal x As Long) As Single
    CellCX = GRID_X + x * CELL_W + CELL_W / 2
End Function

Private Function CellTop(ByVal y As Long) As Single
    CellTop = GRID_Y + y * CELL_H
End Function

Private Function GlitchName(ByVal kind As Long) As String
    Select Case kind
    Case GK_ZIPPER: GlitchName = "Zipper"
    Case GK_WOBBLER: GlitchName = "Wobbler"
    Case GK_HUNTER: GlitchName = "Hunter"
    Case Else: GlitchName = "Scrambler"
    End Select
End Function

Private Function GlitchColor(ByVal kind As Long) As Long
    Select Case kind
    Case GK_ZIPPER: GlitchColor = RGB(255, 70, 80)
    Case GK_WOBBLER: GlitchColor = RGB(255, 160, 40)
    Case GK_HUNTER: GlitchColor = RGB(200, 110, 255)
    Case Else: GlitchColor = RGB(40, 225, 200)
    End Select
End Function

'==============================================================
' Floating "+30" style text
'==============================================================
Private Sub AddFloater(ByVal s As String, ByVal x As Single, ByVal y As Single, ByVal c As Long)
    Dim i As Long
    For i = 0 To MAX_FLOAT - 1
        If Not mFl(i).Active Then
            mFl(i).Active = True
            mFl(i).Txt = s
            mFl(i).X = x
            mFl(i).Y = y
            mFl(i).Clr = c
            mFl(i).Life = 22
            Exit Sub
        End If
    Next
End Sub

Private Sub UpdateFloaters()
    Dim i As Long
    For i = 0 To MAX_FLOAT - 1
        If mFl(i).Active Then
            mFl(i).Y = mFl(i).Y - 1.5
            mFl(i).Life = mFl(i).Life - 1
            If mFl(i).Life <= 0 Then mFl(i).Active = False
        End If
    Next
End Sub

Private Sub ClearFloaters()
    Dim i As Long
    For i = 0 To MAX_FLOAT - 1
        mFl(i).Active = False
    Next
End Sub

'==============================================================
' Drawing
'==============================================================
Private Sub Render()
    Me.Cls
    Select Case mState
    Case ST_TITLE
        DrawTitle
    Case ST_SCORES
        DrawScores
    Case ST_ABOUT
        DrawAbout
    Case ST_NAME
        DrawNameEntry
    Case Else
        DrawPlayfield
        If mState = ST_MSG Then
            DrawPanel mMsgTitle, mMsgColor, MsgFooter(), mMsgLines, mMsgCount
        ElseIf mState = ST_PAUSE Then
            DrawPausePanel
        End If
    End Select
    Me.Refresh
End Sub

Private Function MsgFooter() As String
    If mMsgDelay > 0 Then Exit Function
    If ((mFrame \ 6) Mod 4) = 3 Then Exit Function
    MsgFooter = "Press SPACE to continue"
End Function

Private Sub DrawPlayfield()
    Dim x As Long, y As Long, i As Long, tint As Long
    Dim rx As Single, ry As Single, gx As Single, gy As Single

    For y = 0 To ROWS - 1
        For x = 0 To COLS - 1
            DrawCell x, y
        Next
    Next

    If mPow.Active Then
        If mPow.Life > 40 Or ((mFrame \ 3) Mod 2) = 0 Then
            DrawSnowflake CellCX(mPow.X), CellTop(mPow.Y) + CELL_H / 2
        End If
    End If

    For i = 0 To MAX_GLITCH - 1
        If mGl(i).Active And mGl(i).Warn > 0 Then DrawWarning i
    Next

    rx = pFromX + (pX - pFromX) * pT
    ry = pFromY + (pY - pFromY) * pT
    DrawRobot GRID_X + rx * CELL_W + CELL_W / 2, GRID_Y + ry * CELL_H + 30, 0.9, _
              pChomp > 0, pFace, (pHurt > 0) And ((mFrame Mod 4) < 2)

    For i = 0 To MAX_GLITCH - 1
        If mGl(i).Active And mGl(i).Warn = 0 Then
            gx = mGl(i).FromX + (mGl(i).X - mGl(i).FromX) * mGl(i).T
            gy = mGl(i).FromY + (mGl(i).Y - mGl(i).FromY) * mGl(i).T
            tint = -1
            If mFreeze > 0 Then
                ' frozen: icy blue, flickering back to normal just before they thaw
                If mFreeze > 30 Or ((mFrame \ 3) Mod 2) = 0 Then tint = RGB(130, 180, 255)
            End If
            DrawGlitch GRID_X + gx * CELL_W + CELL_W / 2, GRID_Y + gy * CELL_H + CELL_H / 2, _
                       mGl(i).Kind, 1, IIf(mFreeze > 0, i * 7, mFrame + i * 7), tint
        End If
    Next

    DrawFloaters
    ' bars are drawn last so Glitches appear to slide in from underneath
    DrawTopBar
    DrawHud
End Sub

Private Sub DrawCell(ByVal x As Long, ByVal y As Long)
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    Dim s As String, c As Long, txtClr As Long

    x1 = GRID_X + x * CELL_W + 2
    y1 = GRID_Y + y * CELL_H + 2
    x2 = x1 + CELL_W - 5
    y2 = y1 + CELL_H - 5

    If x = pX And y = pY Then
        c = RGB(48, 62, 130)
    ElseIf ((x + y) Mod 2) = 0 Then
        c = RGB(26, 34, 80)
    Else
        c = RGB(32, 42, 92)
    End If
    FillBox x1, y1, x2, y2, c, RGB(66, 86, 168)

    If Not mGrid(x, y).Filled Then Exit Sub

    If mBonus = BONUS_DASH Then
        DrawBattery (x1 + x2) / 2, (y1 + y2) / 2
        Exit Sub
    End If

    s = mGrid(x, y).Txt
    txtClr = RGB(238, 242, 255)
    If mBonus = BONUS_MEMORY And mMemHidden Then
        s = "?"
        txtClr = RGB(120, 150, 230)
    End If

    Me.FontBold = True
    If x = pX And y = pY Then
        ' Bolt is standing here - show the word small, under the robot
        FitFont s, CELL_W - 16, 10
        TextC s, (x1 + x2) / 2, y2 - 9, RGB(255, 250, 190), False
    Else
        If mFitText(x, y) <> s Then
            FitFont s, CELL_W - 16, 14
            mFitText(x, y) = s
            mFitSize(x, y) = Me.FontSize
        Else
            Me.FontSize = mFitSize(x, y)
        End If
        TextC s, (x1 + x2) / 2, (y1 + y2) / 2, txtClr, True
    End If
End Sub

Private Sub DrawBattery(ByVal cx As Single, ByVal cy As Single)
    Dim bob As Single
    bob = Sin((mFrame + cx) * 0.25) * 2
    cy = cy + bob
    FillBox cx - 4, cy - 20, cx + 4, cy - 16, RGB(190, 200, 220), -1
    FillBox cx - 11, cy - 16, cx + 11, cy + 18, RGB(36, 50, 84), RGB(190, 200, 220)
    FillBox cx - 8, cy - 6, cx + 8, cy + 15, RGB(70, 230, 110), -1
    Me.DrawWidth = 2
    Me.Line (cx + 3, cy - 12)-(cx - 4, cy + 2), RGB(255, 230, 70)
    Me.Line -(cx + 4, cy + 2), RGB(255, 230, 70)
    Me.Line -(cx - 3, cy + 13), RGB(255, 230, 70)
    Me.DrawWidth = 1
End Sub

Private Sub DrawSnowflake(ByVal cx As Single, ByVal cy As Single)
    Dim i As Long, a As Single
    FillDisc cx, cy, 16, RGB(40, 110, 220)
    FillDisc cx, cy, 13, RGB(70, 150, 245)
    Me.DrawWidth = 2
    For i = 0 To 2
        a = i * 1.0472 + mFrame * 0.05
        Me.Line (cx - Cos(a) * 10, cy - Sin(a) * 10)-(cx + Cos(a) * 10, cy + Sin(a) * 10), vbWhite
    Next
    Me.DrawWidth = 1
End Sub

Private Sub DrawWarning(ByVal i As Long)
    Dim ex As Long, ey As Long, x1 As Single, y1 As Single, c As Long
    Dim bdx As Single, bdy As Single
    ex = mGl(i).X + mGl(i).DirX
    ey = mGl(i).Y + mGl(i).DirY
    If Not InGrid(ex, ey) Then Exit Sub
    If ((mFrame \ 3) Mod 2) = 1 Then Exit Sub

    c = GlitchColor(mGl(i).Kind)
    x1 = GRID_X + ex * CELL_W
    y1 = GRID_Y + ey * CELL_H
    Me.DrawWidth = 3
    Me.Line (x1 + 4, y1 + 4)-(x1 + CELL_W - 6, y1 + CELL_H - 6), c, B
    Me.DrawWidth = 1

    ' "!" badge on the side the Glitch is coming from
    bdx = x1 + CELL_W / 2 - mGl(i).DirX * (CELL_W / 2 - 14)
    bdy = y1 + CELL_H / 2 - mGl(i).DirY * (CELL_H / 2 - 14)
    FillDisc bdx, bdy, 9, c
    Me.FontSize = 11
    TextC "!", bdx, bdy, RGB(0, 0, 0), False
End Sub

Private Sub DrawTopBar()
    Dim frac As Single, barC As Long
    FillBox 0, 0, SCR_W, GRID_Y - 8, RGB(24, 30, 88), -1
    Me.FontBold = True

    If mBonus <> BONUS_NONE Then
        ' the stripe under the title is the bonus clock
        frac = 0
        If mBonusPhase > 0 Then frac = mBonusShown / mBonusPhase
        If frac > 1 Then frac = 1
        If frac < 0 Then frac = 0
        If mBonus = BONUS_MEMORY And Not mMemHidden Then barC = RGB(90, 220, 255) Else barC = RGB(255, 196, 40)
        FillBox 0, GRID_Y - 9, SCR_W, GRID_Y - 4, RGB(90, 30, 40), -1
        If frac > 0 Then FillBox 0, GRID_Y - 9, SCR_W * frac, GRID_Y - 4, barC, -1
        Me.FontSize = 8
        TextC "BONUS ROUND  -  " & UCase$(BonusName(mBonus)), SCR_W / 2, 10, RGB(120, 220, 255), False
        Me.FontSize = 18
        TextR CStr((mBonusShown + 999) \ 1000), SCR_W - 18, 30, barC
    Else
        FillBox 0, GRID_Y - 8, SCR_W, GRID_Y - 5, RGB(255, 196, 40), -1
        Me.FontSize = 8
        TextC "GOBBLE EVERY SQUARE THAT FITS THE RULE", SCR_W / 2, 10, RGB(140, 160, 235), False
    End If

    FitFont mRule.Title, SCR_W - 120, 22
    TextC mRule.Title, SCR_W / 2, 33, RGB(255, 228, 90), True
End Sub

Private Sub DrawHud()
    Dim i As Long, s As String
    FillBox 0, HUD_Y, SCR_W, SCR_H, RGB(18, 22, 60), -1
    FillBox 0, HUD_Y, SCR_W, HUD_Y + 2, RGB(66, 86, 168), -1
    Me.FontBold = True

    Me.FontSize = 8
    TextL "SCORE", 20, HUD_Y + 14, RGB(140, 160, 235)
    If mBonus = BONUS_NONE And ComboMult() > 1 Then
        TextL "COMBO x" & ComboMult(), 70, HUD_Y + 14, RGB(255, 170, 60)
    End If
    TextL "LEVEL", 170, HUD_Y + 14, RGB(140, 160, 235)
    TextL "LIVES", 262, HUD_Y + 14, RGB(140, 160, 235)
    If mBonus <> BONUS_NONE Then
        s = "Bonus round:  +" & mBonusScore
    ElseIf mFreeze > 0 Then
        s = "Glitches frozen!"
    Else
        s = ModeName(mMode) & "  /  " & DiffName()
    End If
    TextR s, SCR_W - 20, HUD_Y + 14, RGB(140, 160, 235)
    TextR "P = pause    M = sound " & IIf(gSoundOn, "on", "off"), SCR_W - 20, HUD_Y + 34, RGB(110, 125, 190)

    Me.FontSize = 15
    TextL CStr(mScore), 20, HUD_Y + 34, RGB(255, 255, 255)
    TextL CStr(mLevel), 170, HUD_Y + 34, RGB(255, 255, 255)

    For i = 1 To mLives
        DrawRobot 272 + (i - 1) * 24, HUD_Y + 37, 0.42, False, 0, False
    Next
End Sub

Private Sub DrawFloaters()
    Dim i As Long
    Me.FontBold = True
    Me.FontSize = 14
    For i = 0 To MAX_FLOAT - 1
        If mFl(i).Active Then TextC mFl(i).Txt, mFl(i).X, mFl(i).Y, mFl(i).Clr, True
    Next
End Sub

Private Sub DrawPanel(ByVal panelTitle As String, ByVal clr As Long, ByVal footer As String, _
                      panelLines() As String, ByVal lineCount As Long)
    Dim w As Single, h As Single, x1 As Single, y1 As Single, yy As Single
    Dim i As Long, s As String

    w = 520
    h = 78 + lineCount * 28 + 34
    x1 = (SCR_W - w) / 2
    y1 = (SCR_H - h) / 2

    FillBox x1 + 8, y1 + 8, x1 + w + 8, y1 + h + 8, RGB(4, 6, 16), -1
    FillBox x1, y1, x1 + w, y1 + h, RGB(22, 28, 72), -1
    Me.DrawWidth = 3
    Me.Line (x1, y1)-(x1 + w, y1 + h), clr, B
    Me.DrawWidth = 1
    Me.Line (x1 + 6, y1 + 6)-(x1 + w - 6, y1 + h - 6), RGB(60, 70, 140), B

    Me.FontBold = True
    FitFont panelTitle, w - 40, 24
    TextC panelTitle, SCR_W / 2, y1 + 36, clr, True

    yy = y1 + 82
    For i = 0 To lineCount - 1
        s = panelLines(i)
        If Mid$(s, 1, 1) = "*" Then
            s = Mid$(s, 2)
            FitFont s, w - 40, 18
            TextC s, SCR_W / 2, yy, RGB(255, 228, 90), True
        Else
            FitFont s, w - 40, 12
            TextC s, SCR_W / 2, yy, RGB(230, 235, 255), False
        End If
        yy = yy + 28
    Next

    If Len(footer) > 0 Then
        Me.FontSize = 10
        TextC footer, SCR_W / 2, y1 + h - 20, RGB(150, 165, 220), False
    End If
End Sub

Private Sub DrawPausePanel()
    Dim pl(0 To 2) As String
    pl(0) = "Take a break - everything is frozen."
    pl(1) = "Press P to keep playing."
    pl(2) = "Press Esc to quit to the title screen."
    DrawPanel "PAUSED", RGB(120, 200, 255), "", pl, 3
End Sub

Private Sub DrawStars()
    Dim i As Long
    Me.DrawWidth = 2
    For i = 0 To 59
        If ((i + mFrame \ 3) Mod 11) = 0 Then
            Me.PSet (mStarX(i), mStarY(i)), RGB(255, 255, 255)
        Else
            Me.PSet (mStarX(i), mStarY(i)), RGB(70, 80, 140)
        End If
    Next
    Me.DrawWidth = 1
End Sub

Private Sub DrawTitle()
    Dim w1 As Single, w2 As Single, x0 As Single, bob As Single
    Dim i As Long, bx As Single, edgeC As Long, fillC As Long, s As String
    Dim labels(0 To 2) As String

    DrawStars

    Me.FontBold = True
    Me.FontSize = 40
    w1 = Me.TextWidth("GRID ")
    w2 = Me.TextWidth("GOBBLER")
    x0 = (SCR_W - w1 - w2) / 2
    TextL "GRID ", x0 + 3, 47, RGB(0, 0, 0)
    TextL "GRID ", x0, 44, RGB(255, 210, 50)
    TextL "GOBBLER", x0 + w1 + 3, 47, RGB(0, 0, 0)
    TextL "GOBBLER", x0 + w1, 44, RGB(90, 220, 255)

    Me.FontSize = 11
    TextC "Help Bolt the robot gobble every square that fits the rule!", SCR_W / 2, 94, RGB(200, 210, 255)

    bob = Sin(mFrame * 0.18) * 5
    DrawRobot 150, 204 + bob, 2.2, ((mFrame \ 7) Mod 4) = 0, (mFrame \ 25) Mod 4, False

    DrawGlitch 312, 136, GK_ZIPPER, 0.75, mFrame
    DrawGlitch 312, 178, GK_WOBBLER, 0.75, mFrame + 5
    DrawGlitch 312, 220, GK_HUNTER, 0.75, mFrame + 10
    DrawGlitch 312, 262, GK_SCRAMBLER, 0.75, mFrame + 15

    Me.FontSize = 11
    TextL "ZIPPER", 346, 128, GlitchColor(GK_ZIPPER)
    TextL "WOBBLER", 346, 170, GlitchColor(GK_WOBBLER)
    TextL "HUNTER", 346, 212, GlitchColor(GK_HUNTER)
    TextL "SCRAMBLER", 346, 254, GlitchColor(GK_SCRAMBLER)
    Me.FontSize = 9
    TextL "Dashes straight across the grid.", 346, 145, RGB(170, 180, 220)
    TextL "Wanders around at random.", 346, 187, RGB(170, 180, 220)
    TextL "Chases Bolt!", 346, 229, RGB(170, 180, 220)
    TextL "Drops new squares onto the board.", 346, 271, RGB(170, 180, 220)
    TextL "Grab a snowflake to freeze the Glitches!", 346, 296, RGB(150, 210, 255)
    DrawSnowflake 312, 296

    labels(0) = "1   MATH"
    labels(1) = "2   WORDS"
    labels(2) = "3   MIXED"
    For i = 0 To 2
        bx = 50 + i * 185
        If ((mFrame \ 8) Mod 3) = i Then edgeC = RGB(255, 210, 50) Else edgeC = RGB(90, 110, 200)
        FillBox bx, 316, bx + 170, 348, RGB(30, 40, 100), edgeC
        Me.FontSize = 14
        TextC labels(i), bx + 85, 332, RGB(255, 255, 255), True
    Next

    ' difficulty picker
    Me.FontSize = 9
    TextR "DIFFICULTY", 236, 367, RGB(170, 185, 235)
    labels(0) = "EASY"
    labels(1) = "NORMAL"
    labels(2) = "HARD"
    For i = 0 To 2
        bx = 280 + i * 70
        If i = mDiff Then
            fillC = RGB(255, 196, 40)
            FillBox bx - 32, 356, bx + 32, 378, fillC, RGB(255, 255, 255)
            TextC labels(i), bx, 367, RGB(20, 20, 50)
        Else
            FillBox bx - 32, 356, bx + 32, 378, RGB(26, 34, 80), RGB(70, 90, 170)
            TextC labels(i), bx, 367, RGB(150, 165, 220)
        End If
    Next
    TextL "D = change", 462, 367, RGB(120, 135, 200)

    Me.FontSize = 9
    TextC "Arrows = move     Space = gobble     P = pause     M = sound (" & IIf(gSoundOn, "on", "off") & ")", _
          SCR_W / 2, 394, RGB(170, 185, 235)
    TextC "Bonus every 3 levels (B = try one)     H = Hall of Fame     A = About     Esc = quit", _
          SCR_W / 2, 411, RGB(170, 185, 235)
    s = CategoryCount() & " word categories"
    If mCustomCats > 0 Then s = s & "  (" & mCustomCats & " from the packs folder)"
    TextC s & "        Top gobbler:  " & gHi(1).PName & "  " & gHi(1).Score, SCR_W / 2, 432, RGB(255, 210, 50)
    If ((mFrame \ 8) Mod 3) <> 0 Then
        Me.FontSize = 12
        TextC "Press 1, 2 or 3 to start", SCR_W / 2, 460, RGB(255, 255, 255)
    End If
End Sub

Private Sub DrawScores()
    Dim i As Long, y As Single, c As Long, hdr As Long
    DrawStars
    Me.FontBold = True
    Me.FontSize = 30
    TextC "HALL OF FAME", SCR_W / 2, 44, RGB(255, 210, 50), True

    hdr = RGB(140, 160, 235)
    Me.FontSize = 10
    TextL "RANK", 64, 92, hdr
    TextL "NAME", 140, 92, hdr
    TextR "SCORE", 420, 92, hdr
    TextC "LEVEL", 485, 92, hdr
    TextC "MODE", 556, 92, hdr
    FillBox 60, 104, 580, 105, RGB(70, 90, 170), -1

    For i = 1 To 10
        y = 104 + i * 28
        If i = mNewRank Then
            FillBox 60, y - 12, 580, y + 12, RGB(40, 50, 110), -1
            If ((mFrame \ 4) Mod 2) = 0 Then c = RGB(255, 240, 120) Else c = RGB(255, 160, 60)
        Else
            c = RGB(225, 230, 255)
        End If
        Me.FontSize = 12
        TextL i & ".", 72, y, c
        TextL gHi(i).PName, 140, y, c
        TextR CStr(gHi(i).Score), 420, y, c
        TextC CStr(gHi(i).Level), 485, y, c
        TextC gHi(i).GameMode, 556, y, c
    Next

    Me.FontSize = 9
    TextC "E / N / H after the mode = Easy, Normal or Hard", SCR_W / 2, 424, RGB(120, 135, 200)
    If ((mFrame \ 8) Mod 3) <> 0 Then
        Me.FontSize = 11
        TextC "Press SPACE to return", SCR_W / 2, 452, RGB(150, 165, 220)
    End If
End Sub

Private Sub DrawAbout()
    Dim i As Long, y As Single, hdr As Long, soft As Long
    Dim s As String, ch As String, x As Single, tw As Single
    Dim lap As Long, px0 As Single, gx As Single, goRight As Boolean
    Dim who(0 To 4) As String, role(0 To 4) As String

    DrawStars
    hdr = RGB(140, 160, 235)
    soft = RGB(200, 210, 255)

    Me.FontBold = True
    Me.FontSize = 28
    TextC "ABOUT", SCR_W / 2, 36, RGB(90, 220, 255), True
    Me.FontSize = 10
    TextC "Grid Gobbler   -   version " & App.Major & "." & App.Minor, SCR_W / 2, 70, soft

    ' ---- the big wavy, colour-cycling name ----
    TextC "DREAMED UP AND CREATED BY", SCR_W / 2, 102, hdr
    s = "JONATHAN VALENTIN"
    Me.FontSize = 24
    tw = Me.TextWidth(s)
    x = (SCR_W - tw) / 2
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        y = 138 + Sin((mFrame + i * 3) * 0.22) * 5
        TextL ch, x + 2, y + 2, RGB(0, 0, 0)
        TextL ch, x, y, NameColor(mFrame \ 2 + i)
        x = x + Me.TextWidth(ch)
    Next

    ' ---- the cast ----
    Me.FontSize = 10
    TextC "STARRING", SCR_W / 2, 180, hdr
    FillBox 150, 190, 490, 191, RGB(70, 90, 170), -1

    who(0) = "BOLT":      role(0) = "as himself (always hungry)"
    who(1) = "ZIPPER":    role(1) = "does all his own stunts"
    who(2) = "WOBBLER":   role(2) = "still looking for the exit"
    who(3) = "HUNTER":    role(3) = "just wants a hug. Honest."
    who(4) = "SCRAMBLER": role(4) = "tidies up by making a mess"

    For i = 0 To 4
        y = 208 + i * 24
        If i = 0 Then
            DrawRobot 150, y, 0.42, ((mFrame \ 6) Mod 3) = 0, 0, False
            Me.FontSize = 11
            TextL who(i), 176, y, RGB(255, 210, 50)
        Else
            DrawGlitch 150, y, i - 1, 0.42, mFrame + i * 5
            Me.FontSize = 11
            TextL who(i), 176, y, GlitchColor(i - 1)
        End If
        Me.FontSize = 9
        TextL role(i), 290, y, RGB(190, 200, 235)
    Next

    ' ---- small print ----
    Me.FontSize = 9
    TextC "Written in Visual Basic 6.  Yes, really.  On purpose.", SCR_W / 2, 336, soft
    TextC "No squares were harmed in the making of this game.  (They were gobbled.)", SCR_W / 2, 354, soft
    TextC "Free and open source under the MIT License.  (c) 2026 Jonathan Valentin", SCR_W / 2, 372, RGB(255, 210, 50)

    ' ---- the chase: Glitches chase Bolt one way, then Bolt chases them back ----
    lap = (mFrame \ 150) Mod 2
    goRight = (lap = 0)
    If goRight Then
        px0 = -340 + (mFrame Mod 150) * 7
    Else
        px0 = SCR_W + 60 - (mFrame Mod 150) * 7
    End If
    y = 414 + Abs(Sin(mFrame * 0.4)) * -4
    If goRight Then
        ' Bolt out in front, running for it
        DrawRobot px0 + 280, y, 0.9, ((mFrame \ 3) Mod 2) = 0, 3, False
        For i = 0 To 3
            gx = px0 + i * 56
            DrawGlitch gx, 414 + Sin((mFrame + i * 4) * 0.5) * 3, 3 - i, 0.7, mFrame + i * 5
        Next
    Else
        ' the tables have turned: frozen blue Glitches flee from Bolt
        For i = 0 To 3
            gx = px0 + i * 56
            DrawGlitch gx, 414 + Sin((mFrame + i * 4) * 0.5) * 3, i, 0.7, mFrame + i * 5, RGB(120, 190, 255)
        Next
        DrawRobot px0 + 280, y, 0.9, ((mFrame \ 3) Mod 2) = 0, 2, False
        DrawSnowflake px0 + 320, 392
    End If

    If ((mFrame \ 8) Mod 3) <> 0 Then
        Me.FontSize = 11
        TextC "Press SPACE to return", SCR_W / 2, 458, RGB(150, 165, 220)
    End If
End Sub

Private Function NameColor(ByVal n As Long) As Long
    Select Case n Mod 6
    Case 0: NameColor = RGB(255, 210, 50)
    Case 1: NameColor = RGB(255, 150, 60)
    Case 2: NameColor = RGB(255, 100, 120)
    Case 3: NameColor = RGB(200, 130, 255)
    Case 4: NameColor = RGB(90, 220, 255)
    Case Else: NameColor = RGB(120, 240, 150)
    End Select
End Function

Private Sub DrawNameEntry()
    Dim s As String, c As Long
    DrawStars
    Me.FontBold = True
    Me.FontSize = 30
    If ((mFrame \ 4) Mod 2) = 0 Then c = RGB(255, 230, 80) Else c = RGB(255, 150, 60)
    TextC "NEW HIGH SCORE!", SCR_W / 2, 80, c, True

    Me.FontSize = 13
    TextC "You scored " & mScore & " points and reached level " & mLevel & ".", SCR_W / 2, 140, RGB(220, 230, 255)
    TextC "That's #" & mNewRank & " in the Hall of Fame!", SCR_W / 2, 166, RGB(220, 230, 255)

    Me.FontSize = 12
    TextC "Type your name, then press ENTER:", SCR_W / 2, 214, RGB(160, 180, 240)
    FillBox 170, 236, 470, 280, RGB(10, 14, 36), RGB(255, 200, 40)
    s = mNameEntry
    If ((mFrame \ 5) Mod 2) = 0 Then s = s & "_"
    Me.FontSize = 18
    TextL s, 184, 258, RGB(255, 255, 255)

    DrawRobot SCR_W / 2, 372 + Sin(mFrame * 0.3) * 5, 1.6, ((mFrame \ 5) Mod 2) = 0, 0, False
End Sub

' Bolt the robot. (cx,cy) is roughly the middle of the head/body.
Private Sub DrawRobot(ByVal cx As Single, ByVal cy As Single, ByVal sc As Single, _
                      ByVal chomp As Boolean, ByVal face As Long, ByVal hurt As Boolean)
    Dim headC As Long, bodyC As Long, edgeC As Long, boltC As Long, lightC As Long
    Dim ex As Single, ey As Single, dw As Long

    headC = RGB(150, 220, 255)
    bodyC = RGB(255, 196, 40)
    edgeC = RGB(30, 50, 90)
    boltC = RGB(170, 180, 200)
    If hurt Then
        headC = RGB(255, 110, 110)
        bodyC = RGB(255, 150, 150)
    End If

    ' antenna
    dw = CLng(sc * 2)
    If dw < 1 Then dw = 1
    Me.DrawWidth = dw
    Me.Line (cx, cy - 19 * sc)-(cx, cy - 27 * sc), RGB(190, 200, 220)
    Me.DrawWidth = 1
    If ((mFrame \ 6) Mod 2) = 0 Then
        FillDisc cx, cy - 29 * sc, 3.5 * sc, RGB(255, 70, 70)
    Else
        FillDisc cx, cy - 29 * sc, 3.5 * sc, RGB(150, 20, 20)
    End If

    ' ear bolts and head
    FillBox cx - 19 * sc, cy - 13 * sc, cx - 15 * sc, cy - 5 * sc, boltC, edgeC
    FillBox cx + 15 * sc, cy - 13 * sc, cx + 19 * sc, cy - 5 * sc, boltC, edgeC
    FillBox cx - 16 * sc, cy - 20 * sc, cx + 16 * sc, cy + 3 * sc, headC, edgeC

    ' eyes look the way Bolt is facing
    Select Case face
    Case 1: ex = 0: ey = -1.6 * sc
    Case 2: ex = -1.6 * sc: ey = 0
    Case 3: ex = 1.6 * sc: ey = 0
    Case Else: ex = 0: ey = 1.2 * sc
    End Select
    FillDisc cx - 7 * sc, cy - 11 * sc, 4.5 * sc, vbWhite
    FillDisc cx + 7 * sc, cy - 11 * sc, 4.5 * sc, vbWhite
    FillDisc cx - 7 * sc + ex, cy - 11 * sc + ey, 2.2 * sc, RGB(10, 10, 30)
    FillDisc cx + 7 * sc + ex, cy - 11 * sc + ey, 2.2 * sc, RGB(10, 10, 30)

    ' mouth
    If chomp Then
        FillBox cx - 10 * sc, cy - 5 * sc, cx + 10 * sc, cy + 2 * sc, RGB(70, 0, 20), edgeC
        FillBox cx - 8 * sc, cy - 5 * sc, cx - 5 * sc, cy - 3 * sc, vbWhite, -1
        FillBox cx - 1.5 * sc, cy - 5 * sc, cx + 1.5 * sc, cy - 3 * sc, vbWhite, -1
        FillBox cx + 5 * sc, cy - 5 * sc, cx + 8 * sc, cy - 3 * sc, vbWhite, -1
    Else
        FillBox cx - 8 * sc, cy - 3 * sc, cx + 8 * sc, cy - 1 * sc, edgeC, -1
    End If

    ' body with a blinking chest light
    FillBox cx - 12 * sc, cy + 4 * sc, cx + 12 * sc, cy + 17 * sc, bodyC, RGB(120, 80, 0)
    If ((mFrame \ 4) Mod 2) = 0 Then lightC = RGB(60, 220, 100) Else lightC = RGB(30, 120, 60)
    FillBox cx - 4 * sc, cy + 8 * sc, cx + 4 * sc, cy + 12 * sc, lightC, -1

    ' wheels
    FillDisc cx - 8 * sc, cy + 19 * sc, 3.5 * sc, RGB(60, 60, 75)
    FillDisc cx + 8 * sc, cy + 19 * sc, 3.5 * sc, RGB(60, 60, 75)
End Sub

' A spiky Glitch bug. tint <> -1 overrides its colour (used when frozen).
Private Sub DrawGlitch(ByVal cx As Single, ByVal cy As Single, ByVal kind As Long, ByVal sc As Single, _
                       ByVal fr As Long, Optional ByVal tint As Long = -1)
    Dim c As Long, i As Long, a As Single, r1 As Single, r2 As Single, dw As Long, spin As Single

    c = GlitchColor(kind)
    If tint <> -1 Then c = tint
    dw = CLng(sc * 2)
    If dw < 1 Then dw = 1
    If (kind Mod 2) = 0 Then spin = 0.12 Else spin = -0.12

    Me.DrawWidth = dw
    r1 = 12 * sc
    For i = 0 To 9
        a = i * 0.6283185 + fr * spin
        If ((i + fr \ 3) Mod 2) = 0 Then r2 = 25 * sc Else r2 = 19 * sc
        Me.Line (cx + Cos(a) * r1, cy + Sin(a) * r1)-(cx + Cos(a) * r2, cy + Sin(a) * r2), c
    Next
    Me.DrawWidth = 1
    FillDisc cx, cy, 15 * sc, c

    ' eyes (Hunters have red ones)
    FillDisc cx - 6 * sc, cy - 3 * sc, 4.5 * sc, vbWhite
    FillDisc cx + 6 * sc, cy - 3 * sc, 4.5 * sc, vbWhite
    If kind = GK_HUNTER Then
        FillDisc cx - 5 * sc, cy - 2 * sc, 2.2 * sc, RGB(220, 0, 0)
        FillDisc cx + 5 * sc, cy - 2 * sc, 2.2 * sc, RGB(220, 0, 0)
    Else
        FillDisc cx - 5 * sc, cy - 2 * sc, 2.2 * sc, vbBlack
        FillDisc cx + 5 * sc, cy - 2 * sc, 2.2 * sc, vbBlack
    End If

    ' angry brows and zig-zag mouth
    Me.DrawWidth = dw
    Me.Line (cx - 11 * sc, cy - 10 * sc)-(cx - 3 * sc, cy - 7 * sc), vbBlack
    Me.Line (cx + 11 * sc, cy - 10 * sc)-(cx + 3 * sc, cy - 7 * sc), vbBlack
    Me.Line (cx - 7 * sc, cy + 6 * sc)-(cx - 3.5 * sc, cy + 9 * sc), vbBlack
    Me.Line -(cx, cy + 6 * sc), vbBlack
    Me.Line -(cx + 3.5 * sc, cy + 9 * sc), vbBlack
    Me.Line -(cx + 7 * sc, cy + 6 * sc), vbBlack
    Me.DrawWidth = 1

    ' Scramblers have two sparks orbiting them
    If kind = GK_SCRAMBLER Then
        a = fr * 0.3
        FillDisc cx + Cos(a) * 21 * sc, cy + Sin(a) * 21 * sc, 2.5 * sc, vbWhite
        FillDisc cx - Cos(a) * 21 * sc, cy - Sin(a) * 21 * sc, 2.5 * sc, vbWhite
    End If
End Sub

'==============================================================
' Low-level drawing helpers
'==============================================================
Private Sub FillBox(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                    ByVal fillC As Long, ByVal edgeC As Long)
    Me.Line (x1, y1)-(x2, y2), fillC, BF
    If edgeC <> -1 Then Me.Line (x1, y1)-(x2, y2), edgeC, B
End Sub

Private Sub FillDisc(ByVal x As Single, ByVal y As Single, ByVal r As Single, ByVal c As Long)
    If r < 0.5 Then r = 0.5
    Me.FillStyle = vbFSSolid
    Me.FillColor = c
    Me.Circle (x, y), r, c
    Me.FillStyle = vbFSTransparent
End Sub

' Shrink the font until s fits in maxW pixels
Private Sub FitFont(ByVal s As String, ByVal maxW As Single, ByVal startSize As Single)
    Dim sz As Single
    sz = startSize
    Me.FontSize = sz
    Do While Me.TextWidth(s) > maxW And sz > 7
        sz = sz - 1
        Me.FontSize = sz
    Loop
End Sub

Private Sub TextC(ByVal s As String, ByVal cx As Single, ByVal cy As Single, ByVal clr As Long, _
                  Optional ByVal shadow As Boolean = False)
    Dim tw As Single, th As Single
    tw = Me.TextWidth(s)
    th = Me.TextHeight(s)
    If shadow Then
        Me.ForeColor = RGB(0, 0, 0)
        Me.CurrentX = cx - tw / 2 + 2
        Me.CurrentY = cy - th / 2 + 2
        Me.Print s
    End If
    Me.ForeColor = clr
    Me.CurrentX = cx - tw / 2
    Me.CurrentY = cy - th / 2
    Me.Print s
End Sub

Private Sub TextL(ByVal s As String, ByVal x As Single, ByVal cy As Single, ByVal clr As Long)
    Me.ForeColor = clr
    Me.CurrentX = x
    Me.CurrentY = cy - Me.TextHeight(s) / 2
    Me.Print s
End Sub

Private Sub TextR(ByVal s As String, ByVal x As Single, ByVal cy As Single, ByVal clr As Long)
    Me.ForeColor = clr
    Me.CurrentX = x - Me.TextWidth(s)
    Me.CurrentY = cy - Me.TextHeight(s) / 2
    Me.Print s
End Sub
