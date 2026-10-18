Attribute VB_Name = "Typewriter"
' Typewriter: write forward only. Works in Word for Windows and Word for Mac.
' Start it from the Typewriter tab on the ribbon, or with Ctrl+Alt+T
' (Control+Option+T on a Mac).
' While it is on you can only add text at the end of the document:
' no backspace, delete, cut, undo, replace, overtyping, drag-and-drop,
' AutoCorrect, or spelling/grammar fixes. You can still scroll up to reread.
' With a time limit set, it stays on until the timer runs out.
Option Explicit

Private Const TYPEWRITER_VERSION As String = "0.1.0"
Private Const REPO_URL As String = "https://github.com/skytheacademic/typewriter"
Private Const LATEST_RELEASE_API As String = _
    "https://api.github.com/repos/skytheacademic/typewriter/releases/latest"

Private Const TICK_SECONDS As Long = 1
Private Const DEFAULT_MINUTES As Long = 25

' Modifier codes. Word for Windows and Word for Mac number them differently
' (the Mac uses the classic Mac OS values), so each platform has its own key
' list. The Windows keys are stored in the file when it is built; the Mac keys
' are added when Word starts on a Mac.
Private Const WIN_SHIFT As Long = 256
Private Const WIN_CTRL As Long = 512
Private Const WIN_ALT As Long = 1024
Private Const MAC_CMD As Long = 256
Private Const MAC_SHIFT As Long = 512
Private Const MAC_OPT As Long = 2048
Private Const MAC_CTRL As Long = 4096
Private Const MAC_FWD_DELETE As Long = 127   ' fn+delete, tried alongside wdKeyDelete

Private macKeyErrors As String

Private twOn As Boolean
Private twDoc As Document
Private twEvents As TypewriterEvents
Private twRibbon As IRibbonUI
Private pinning As Boolean

' Timer state
Private twMinutes As Long        ' chosen session length, 0 = no limit
Private prefsLoaded As Boolean
Private twStartTime As Date
Private twEndTime As Date        ' 0 when there is no limit
Private twStartWords As Long
Private tickPending As Boolean
Private lastTimerError As String

' Word settings changed while typewriter mode is on, restored when it ends
Private origReplaceText As Boolean
Private origDragDrop As Boolean
Private origSpellAsYouType As Boolean
Private origGrammarAsYouType As Boolean
Private origOvertype As Boolean
Private origShowSpelling As Boolean
Private origShowGrammar As Boolean
Private origCaption As String

' ---------- setup ----------

' Run once after importing: assigns the keyboard shortcuts in this template.
Public Sub InstallTypewriter()
    BindKeys
    MsgBox "Typewriter is installed." & vbCr & vbCr & _
        "Now close Word completely and run add-ribbon.ps1 to add the Typewriter tab.", _
        vbInformation, "Typewriter"
End Sub

' The key list for this platform, shared by the key binding and the
' diagnostics. Each entry is "key code|macro"; each macro does the
' platform's normal thing when typewriter mode is off.
Private Function KeyList() As Variant
    #If Mac Then
        KeyList = Array( _
            wdKeyBackspace & "|TypewriterBackspace", _
            (MAC_SHIFT + wdKeyBackspace) & "|TypewriterBackspace", _
            (MAC_CMD + wdKeyBackspace) & "|TypewriterCtrlBackspace", _
            (MAC_OPT + wdKeyBackspace) & "|TypewriterAltBackspace", _
            wdKeyDelete & "|TypewriterDelete", _
            MAC_FWD_DELETE & "|TypewriterDelete", _
            (MAC_SHIFT + wdKeyDelete) & "|TypewriterDelete", _
            (MAC_CMD + wdKeyDelete) & "|TypewriterCtrlDelete", _
            (MAC_CMD + MAC_FWD_DELETE) & "|TypewriterCtrlDelete", _
            (MAC_OPT + wdKeyDelete) & "|TypewriterAltDelete", _
            (MAC_OPT + MAC_FWD_DELETE) & "|TypewriterAltDelete", _
            (MAC_CTRL + MAC_OPT + wdKeyT) & "|ToggleTypewriter")
    #Else
        KeyList = Array( _
            wdKeyBackspace & "|TypewriterBackspace", _
            (WIN_SHIFT + wdKeyBackspace) & "|TypewriterBackspace", _
            (WIN_CTRL + wdKeyBackspace) & "|TypewriterCtrlBackspace", _
            (WIN_ALT + wdKeyBackspace) & "|TypewriterAltBackspace", _
            wdKeyDelete & "|TypewriterDelete", _
            (WIN_CTRL + wdKeyDelete) & "|TypewriterCtrlDelete", _
            (WIN_ALT + wdKeyDelete) & "|TypewriterAltDelete", _
            (WIN_CTRL + WIN_ALT + wdKeyT) & "|ToggleTypewriter")
    #End If
End Function

' On a Mac, adds any Mac shortcut that isn't bound yet. The Windows key codes
' stored in the file mean different keys there (or none), so these replace
' them. Runs once per Word session; failures are kept for the diagnostics.
Private Sub EnsureMacKeys()
    #If Mac Then
        Static done As Boolean
        Dim k As Variant, parts() As String, cmd As String
        If done Then Exit Sub
        done = True
        macKeyErrors = ""
        On Error Resume Next
        CustomizationContext = MacroContainer
        ' Never fall back to changing someone's Normal template
        If Err.Number <> 0 Or CustomizationContext Is NormalTemplate Then
            macKeyErrors = "couldn't select Typewriter for key changes: error " & _
                Err.Number & " (" & Err.Description & ")" & vbCr
            Exit Sub
        End If
        For Each k In KeyList
            parts = Split(k, "|")
            cmd = ""
            cmd = FindKey(CLng(parts(0))).Command
            If InStr(1, cmd, parts(1), vbTextCompare) = 0 Then
                Err.Clear
                KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:=parts(1), _
                    KeyCode:=CLng(parts(0))
                If Err.Number <> 0 Then
                    macKeyErrors = macKeyErrors & "key code " & parts(0) & ": error " & _
                        Err.Number & " (" & Err.Description & ")" & vbCr
                End If
            End If
        Next k
        MacroContainer.Saved = True
    #End If
End Sub

Private Sub BindKeys()
    Dim k As Variant, parts() As String
    CustomizationContext = MacroContainer
    For Each k In KeyList
        parts = Split(k, "|")
        KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:=parts(1), _
            KeyCode:=CLng(parts(0))
    Next k
    MacroContainer.Save
End Sub

' Runs automatically when Word loads this template from the Startup folder.
Public Sub AutoExec()
    HookEvents
    LoadPrefs
    EnsureMacKeys
End Sub

' Runs automatically when Word quits.
Public Sub AutoExit()
    If twOn Then StopTypewriter
End Sub

Private Sub HookEvents()
    If twEvents Is Nothing Then
        Set twEvents = New TypewriterEvents
        Set twEvents.App = Application
    End If
End Sub

Private Sub LoadPrefs()
    If prefsLoaded Then Exit Sub
    twMinutes = DEFAULT_MINUTES
    On Error Resume Next
    twMinutes = CLng(GetSetting("Typewriter", "Prefs", "Minutes", CStr(DEFAULT_MINUTES)))
    On Error GoTo 0
    prefsLoaded = True
End Sub

Private Sub SetStatus(text As String)
    On Error Resume Next
    Application.StatusBar = text
End Sub

' ---------- on / off ----------

Public Sub ToggleTypewriter()
    If Documents.Count = 0 Then Exit Sub
    CheckTimeUp
    If Not twOn Then
        StartTypewriter ActiveDocument
    ElseIf twEndTime > 0 Then
        MsgBox "Still " & TimeText & " to go. Keep writing!" & vbCr & vbCr & _
            "Typewriter turns off by itself when the timer ends.", _
            vbInformation, "Typewriter"
    ElseIf MsgBox("Leave typewriter mode and allow editing again?", _
            vbYesNo + vbQuestion + vbDefaultButton2, "Typewriter") = vbYes Then
        StopTypewriter
    End If
    RefreshRibbon
End Sub

Private Sub StartTypewriter(doc As Document)
    HookEvents
    LoadPrefs
    Set twDoc = doc

    ' Each setting is changed on its own so one missing on a platform
    ' doesn't stop the rest.
    On Error Resume Next
    origReplaceText = AutoCorrect.ReplaceText
    origDragDrop = Options.AllowDragAndDrop
    origSpellAsYouType = Options.CheckSpellingAsYouType
    origGrammarAsYouType = Options.CheckGrammarAsYouType
    origOvertype = Options.Overtype
    origShowSpelling = doc.ShowSpellingErrors
    origShowGrammar = doc.ShowGrammaticalErrors
    AutoCorrect.ReplaceText = False
    Options.AllowDragAndDrop = False
    Options.CheckSpellingAsYouType = False
    Options.CheckGrammarAsYouType = False
    Options.Overtype = False
    ' Word for Mac keeps marking an open document unless it's told directly
    doc.ShowSpellingErrors = False
    doc.ShowGrammaticalErrors = False
    origCaption = doc.Windows(1).Caption
    On Error GoTo 0

    On Error Resume Next
    twStartWords = doc.ComputeStatistics(wdStatisticWords)
    On Error GoTo 0
    twStartTime = Now
    If twMinutes > 0 Then
        twEndTime = twStartTime + TimeSerial(0, twMinutes, 0)
    Else
        twEndTime = 0
    End If

    twOn = True
    PinCursor
    RefreshDisplay
    ScheduleTick
End Sub

Private Sub StopTypewriter()
    On Error Resume Next
    AutoCorrect.ReplaceText = origReplaceText
    Options.AllowDragAndDrop = origDragDrop
    Options.CheckSpellingAsYouType = origSpellAsYouType
    Options.CheckGrammarAsYouType = origGrammarAsYouType
    Options.Overtype = origOvertype
    twDoc.ShowSpellingErrors = origShowSpelling
    twDoc.ShowGrammaticalErrors = origShowGrammar
    twDoc.Windows(1).Caption = origCaption
    On Error GoTo 0

    twOn = False
    twEndTime = 0
    Set twDoc = Nothing
    SetStatus "Typewriter mode is off."
    RefreshRibbon
End Sub

' Timer ran out: unlock and report how the session went.
Private Sub FinishSession()
    Dim written As Long
    On Error Resume Next
    written = twDoc.ComputeStatistics(wdStatisticWords) - twStartWords
    On Error GoTo 0
    Dim mins As Long
    mins = (DateDiff("s", twStartTime, Now) + 30) \ 60
    StopTypewriter
    MsgBox "Time's up! You wrote " & Format(written, "#,##0") & " words in " & _
        mins & IIf(mins = 1, " minute.", " minutes.") & vbCr & vbCr & _
        "Editing is unlocked.", vbInformation, "Typewriter"
End Sub

' ---------- timer ----------

' Schedules a macro in this module to run after a few seconds. Word versions
' differ in the forms they accept, so try each in turn. Returns False (and
' records the error) if none works.
Private Function ScheduleMacro(seconds As Long, macroName As String) As Boolean
    Dim t As Date
    t = Now + TimeSerial(0, 0, seconds)
    On Error Resume Next
    Err.Clear
    Application.OnTime When:=t, Name:="Typewriter." & macroName
    If Err.Number <> 0 Then
        lastTimerError = "error " & Err.Number & " (" & Err.Description & ")"
        Err.Clear
        Application.OnTime When:=t, Name:=macroName
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime When:=Format(t, "hh:nn:ss"), Name:=macroName
    End If
    ScheduleMacro = (Err.Number = 0)
    If ScheduleMacro Then lastTimerError = ""
End Function

' Word can't cancel an OnTime call, so only one tick is ever queued and it
' simply stops rescheduling itself once typewriter mode is off.
Private Sub ScheduleTick()
    If tickPending Then Exit Sub
    tickPending = ScheduleMacro(TICK_SECONDS, "TypewriterTick")
End Sub

Public Sub TypewriterTick()
    tickPending = False
    If Not twOn Then Exit Sub
    If twEndTime > 0 And Now >= twEndTime Then
        FinishSession
        Exit Sub
    End If
    RefreshDisplay
    ScheduleTick
End Sub

' Backup for computers where the timer can't run: ends a timed session the
' next time you type a blocked key, click, or press the button.
Public Sub CheckTimeUp()
    If twOn And twEndTime > 0 Then
        If Now >= twEndTime Then FinishSession
    End If
End Sub

' "12:34" left when a limit is set, otherwise time elapsed.
Private Function TimeText() As String
    Dim secs As Long
    If twEndTime > 0 Then
        secs = DateDiff("s", Now, twEndTime)
        If secs < 0 Then secs = 0
    Else
        secs = DateDiff("s", twStartTime, Now)
    End If
    TimeText = (secs \ 60) & ":" & Format(secs Mod 60, "00")
End Function

Private Function StatusText() As String
    If Not twOn Then
        StatusText = "Ready"
    ElseIf twEndTime > 0 Then
        StatusText = TimeText & " left"
    Else
        StatusText = TimeText & " elapsed"
    End If
End Function

Private Sub RefreshDisplay()
    On Error Resume Next
    twDoc.Windows(1).Caption = origCaption & "  [TYPEWRITER " & StatusText & "]"
    On Error GoTo 0
    SetStatus "Typewriter: " & StatusText & ". Keep writing."
    If Not twRibbon Is Nothing Then twRibbon.InvalidateControl "twTime"
End Sub

Private Sub RefreshRibbon()
    If Not twRibbon Is Nothing Then twRibbon.Invalidate
End Sub

' ---------- ribbon tab (callbacks named in customUI14.xml) ----------

Private Function Durations() As Variant
    Durations = Array(0, 5, 10, 15, 20, 25, 30, 45, 60, 90)
End Function

Public Sub TWRibbonLoad(ribbon As IRibbonUI)
    Set twRibbon = ribbon
    HookEvents
    LoadPrefs
    EnsureMacKeys
End Sub

Public Sub TWToggleAction(control As IRibbonControl, pressed As Boolean)
    ToggleTypewriter
End Sub

Public Sub TWTogglePressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = twOn
End Sub

Public Sub TWToggleLabel(control As IRibbonControl, ByRef returnedVal)
    If twOn Then returnedVal = "Writing..." Else returnedVal = "Start Writing"
End Sub

Public Sub TWDurationIndex(control As IRibbonControl, ByRef returnedVal)
    Dim d As Variant, i As Long
    LoadPrefs
    d = Durations
    returnedVal = 0
    For i = 0 To UBound(d)
        If d(i) = twMinutes Then returnedVal = i
    Next i
End Sub

Public Sub TWDurationAction(control As IRibbonControl, id As String, index As Integer)
    twMinutes = Durations()(index)
    prefsLoaded = True
    On Error Resume Next
    SaveSetting "Typewriter", "Prefs", "Minutes", CStr(twMinutes)
End Sub

Public Sub TWDurationEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not twOn
End Sub

Public Sub TWTimeLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = StatusText
End Sub

Public Sub TWDiagnosticsAction(control As IRibbonControl)
    If twOn Then
        MsgBox "Finish your writing session first, then run Diagnostics.", _
            vbInformation, "Typewriter"
        Exit Sub
    End If
    Dim timerLine As String
    If ScheduleMacro(3, "TypewriterTimerCheck") Then
        timerLine = "PASS  Timer scheduled. A ""Timer check"" box should appear in a " & _
            "few seconds; if it doesn't, the countdown can't run here."
    Else
        timerLine = "FAIL  Timer: " & lastTimerError & ". Timed sessions still end, " & _
            "but only when you next type, click or press the button."
    End If
    MsgBox TypewriterDiagnostics() & timerLine, vbInformation, "Typewriter Diagnostics"
End Sub

Public Sub TypewriterTimerCheck()
    MsgBox "Timer check: PASS. Word ran the scheduled timer.", vbInformation, "Typewriter Diagnostics"
End Sub

' ---------- updates ----------

' Check for Updates button. Only runs when clicked; nothing is sent in the
' background.
Public Sub TWUpdateAction(control As IRibbonControl)
    Dim latest As String, errText As String
    #If Mac Then
        MsgBox "You have Typewriter " & TYPEWRITER_VERSION & "." & vbCr & vbCr & _
            "Word for Mac can't check automatically. Newer versions are listed at:" & _
            vbCr & REPO_URL & "/releases", vbInformation, "Typewriter"
        Exit Sub
    #End If
    latest = LatestReleaseTag(errText)
    If latest = "" Then
        MsgBox "Couldn't check for updates: " & errText & "." & vbCr & vbCr & _
            "Releases are listed at:" & vbCr & REPO_URL & "/releases", _
            vbExclamation, "Typewriter"
    ElseIf IsNewer(latest, TYPEWRITER_VERSION) Then
        If MsgBox("Typewriter " & latest & " is available (you have " & _
                TYPEWRITER_VERSION & ")." & vbCr & vbCr & "Open the download page?", _
                vbYesNo + vbQuestion, "Typewriter") = vbYes Then
            OpenUrl REPO_URL & "/releases/latest"
        End If
    Else
        MsgBox "You have the latest version (" & TYPEWRITER_VERSION & ").", _
            vbInformation, "Typewriter"
    End If
End Sub

' Asks GitHub for the newest release tag, e.g. "v0.2.0". Returns "" and sets
' errText when it can't tell. Windows only: Word for Mac's VBA has no way to
' make web requests.
Private Function LatestReleaseTag(ByRef errText As String) As String
    #If Mac Then
        errText = "not available on Mac"
    #Else
        Dim http As Object, body As String, p As Long, q As Long
        On Error GoTo Failed
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", LATEST_RELEASE_API, False
        http.SetRequestHeader "User-Agent", "Typewriter-for-Word/" & TYPEWRITER_VERSION
        http.SetRequestHeader "Accept", "application/vnd.github+json"
        http.Send
        If http.Status = 404 Then errText = "no releases published yet": Exit Function
        If http.Status <> 200 Then errText = "GitHub answered " & http.Status: Exit Function
        body = http.ResponseText
        p = InStr(body, """tag_name""")
        If p = 0 Then errText = "unexpected reply from GitHub": Exit Function
        p = InStr(p + 10, body, """")      ' opening quote of the value
        q = InStr(p + 1, body, """")
        LatestReleaseTag = Mid(body, p + 1, q - p - 1)
        Exit Function
Failed:
        errText = "couldn't reach GitHub (" & Err.Description & ")"
    #End If
End Function

' True if release tag "v1.2.3" is newer than version "1.2.2".
Private Function IsNewer(tag As String, current As String) As Boolean
    Dim a() As String, b() As String, i As Long, x As Long, y As Long
    a = Split(Replace(LCase(tag), "v", ""), ".")
    b = Split(current, ".")
    For i = 0 To 2
        x = 0: y = 0
        If i <= UBound(a) Then x = Val(a(i))
        If i <= UBound(b) Then y = Val(b(i))
        If x <> y Then IsNewer = (x > y): Exit Function
    Next i
End Function

Private Sub OpenUrl(url As String)
    On Error Resume Next
    CreateObject("Shell.Application").ShellExecute url
End Sub

' ---------- diagnostics ----------

' Checks the parts that can differ between computers and returns a report.
' Runs a short typewriter session in a scratch document it closes afterwards.
' Any error is written into the report with the check it happened in, and
' the remaining checks still run.
Public Function TypewriterDiagnostics() As String
    Dim r As String, k As Variant, parts() As String, cmd As String
    Dim stage As String, keyName As String
    On Error GoTo Trouble

    stage = "setup"
    #If Mac Then
        r = "Typewriter " & TYPEWRITER_VERSION & ", Word " & Application.Version & " on Mac" & vbCr
    #Else
        r = "Typewriter " & TYPEWRITER_VERSION & ", Word " & Application.Version & " on Windows" & vbCr
    #End If
    HookEvents
    EnsureMacKeys
    If macKeyErrors <> "" Then r = r & "FAIL  Adding Mac shortcuts: " & macKeyErrors
    r = r & Mark(Not twEvents Is Nothing) & "Word events connected" & vbCr
    r = r & Mark(Not twRibbon Is Nothing) & "Ribbon tab connected" & vbCr

    ' Keys: what each one is bound to, as Word names the key on this computer
    For Each k In KeyList
        parts = Split(k, "|")
        stage = "key code " & parts(0) & " (" & parts(1) & ")"
        keyName = stage
        keyName = KeyString(CLng(parts(0))) & " key"
        cmd = ""
        cmd = FindKey(CLng(parts(0))).Command
        r = r & Mark(InStr(1, cmd, parts(1), vbTextCompare) > 0) & keyName & vbCr
    Next k

    If twOn Then
        TypewriterDiagnostics = r & "(Skipped the live test: a session is running.)" & vbCr
        Exit Function
    End If

    ' Live test in a scratch document
    Dim doc As Document, before As String, ok As Boolean, savedMinutes As Long
    stage = "live test"
    Set doc = Documents.Add
    doc.Content.Text = "Diagnostics test sentence."
    before = doc.Content.Text
    savedMinutes = twMinutes
    twMinutes = 0
    StartTypewriter doc
    twMinutes = savedMinutes

    r = r & Mark(twOn) & "Session starts" & vbCr
    stage = "checking paused settings"
    r = r & Mark(Not AutoCorrect.ReplaceText) & "AutoCorrect paused" & vbCr
    r = r & Mark(Not Options.CheckSpellingAsYouType) & "Spellcheck as you type paused" & vbCr
    r = r & Mark(Not doc.ShowSpellingErrors And Not doc.ShowGrammaticalErrors) & _
        "Spelling and grammar marks hidden" & vbCr
    r = r & Mark(Not Options.AllowDragAndDrop) & "Drag-and-drop paused" & vbCr
    stage = "live test"

    doc.Range(0, 4).Select
    DoEvents
    ok = (Selection.Start = doc.Content.End - 1 And Selection.End = Selection.Start)
    r = r & Mark(ok) & "Cursor snaps back to the end" & vbCr
    PinCursor

    Call TypewriterBackspace
    Call TypewriterCtrlBackspace
    Call TypewriterAltBackspace
    Call TypewriterDelete
    Call TypewriterCtrlDelete
    Call TypewriterAltDelete
    Call EditClear
    Call DeleteWord
    Call DeleteBackWord
    Call EditCut
    Call EditUndo
    r = r & Mark(doc.Content.Text = before) & "Delete, cut and undo blocked" & vbCr

    Selection.TypeText " More."
    r = r & Mark(InStr(doc.Content.Text, "sentence. More.") > 0) & "Typing still adds text at the end" & vbCr

    stage = "stopping the test session"
    StopTypewriter
    ok = (AutoCorrect.ReplaceText = origReplaceText And _
        Options.CheckSpellingAsYouType = origSpellAsYouType And _
        Options.AllowDragAndDrop = origDragDrop)
    r = r & Mark(Not twOn And ok) & "Session stops and settings come back" & vbCr

    stage = "closing the test document"
    If twOn Then StopTypewriter
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    TypewriterDiagnostics = r
    Exit Function

Trouble:
    r = r & "ERROR " & Err.Number & " in " & stage & ": " & Err.Description & vbCr
    Resume Next
End Function

Private Function Mark(passed As Boolean) As String
    If passed Then Mark = "PASS  " Else Mark = "FAIL  "
End Function

' ---------- test hook ----------
' Takes an argument so it stays out of the Macros list (no escape hatch).
' action: "install", "start", "stop", "seconds" (start a session of arg
' seconds), "state"
Public Function TypewriterTest(action As String, Optional arg As Long = 0) As String
    Select Case action
        Case "install"
            BindKeys
        Case "start"
            If Not twOn Then StartTypewriter ActiveDocument
        Case "seconds"
            If Not twOn Then StartTypewriter ActiveDocument
            twEndTime = twStartTime + TimeSerial(0, 0, arg)
            RefreshDisplay
        Case "stop"
            If twOn Then StopTypewriter
        Case "update"
            Dim e As String, latest As String
            latest = LatestReleaseTag(e)
            TypewriterTest = "version=" & TYPEWRITER_VERSION & " latest=" & latest & _
                " error=" & e & " newer=" & (latest <> "" And IsNewer(latest, TYPEWRITER_VERSION))
            Exit Function
        Case "versions"
            TypewriterTest = Mark(IsNewer("v0.2.0", "0.1.0")) & Mark(IsNewer("v1.0.0", "0.9.9")) & _
                Mark(IsNewer("v0.1.10", "0.1.9")) & Mark(Not IsNewer("v0.1.0", "0.1.0")) & _
                Mark(Not IsNewer("v0.0.9", "0.1.0")) & Mark(Not IsNewer("v0.1", "0.1.0"))
            Exit Function
    End Select
    TypewriterTest = "on=" & twOn & " minutes=" & twMinutes & " status=" & StatusText & _
        " replaceText=" & AutoCorrect.ReplaceText & " dragDrop=" & Options.AllowDragAndDrop & _
        " spell=" & Options.CheckSpellingAsYouType & " grammar=" & Options.CheckGrammarAsYouType & _
        " overtype=" & Options.Overtype & " noEvents=" & (twEvents Is Nothing) & _
        " noRibbon=" & (twRibbon Is Nothing) & " tickPending=" & tickPending
End Function

' ---------- helpers used by TypewriterEvents ----------

Public Function LocksDocument(doc As Document) As Boolean
    If twOn Then LocksDocument = (doc Is twDoc)
End Function

' Called when a document closes; ends typewriter mode if it was that one.
Public Sub ReleaseDocument(doc As Document)
    If LocksDocument(doc) Then StopTypewriter
End Sub

' Moves the cursor back to the end of the document and clears any selection.
Public Sub PinCursor()
    If Not twOn Or pinning Then Exit Sub
    pinning = True
    On Error Resume Next
    Dim endPos As Long
    endPos = twDoc.Content.End - 1
    If Selection.StoryType <> wdMainTextStory Or Selection.Start <> endPos _
            Or Selection.End <> endPos Then
        twDoc.Range(endPos, endPos).Select
    End If
    On Error GoTo 0
    pinning = False
End Sub

Private Function Locked() As Boolean
    If Documents.Count > 0 Then Locked = LocksDocument(ActiveDocument)
End Function

Private Sub Blocked()
    CheckTimeUp
    If Not twOn Then Exit Sub
    Beep
    SetStatus "Typewriter: no editing yet. Keep going! (" & StatusText & ")"
End Sub

' ---------- blocked keys ----------
' Bound to keys by BindKeys. When typewriter mode is off each one does what
' that key normally does on this platform.

' Backspace / Shift+Backspace (the Mac "delete" key)
Public Sub TypewriterBackspace()
    If Locked Then Blocked Else Selection.TypeBackspace
End Sub

' Windows Ctrl+Backspace: delete previous word. Mac Cmd+delete: delete to line start.
Public Sub TypewriterCtrlBackspace()
    If Locked Then Blocked: Exit Sub
    #If Mac Then
        If Selection.Type = wdSelectionIP Then Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        Selection.Delete
    #Else
        Selection.Delete Unit:=wdWord, Count:=-1
    #End If
End Sub

' Windows Alt+Backspace: undo. Mac Option+delete: delete previous word.
Public Sub TypewriterAltBackspace()
    If Locked Then Blocked: Exit Sub
    #If Mac Then
        Selection.Delete Unit:=wdWord, Count:=-1
    #Else
        On Error Resume Next
        ActiveDocument.Undo
    #End If
End Sub

' Delete (Windows) / fn+delete (Mac)
Public Sub TypewriterDelete()
    If Locked Then Blocked Else Selection.Delete
End Sub

' Windows Ctrl+Delete: delete next word. Mac Cmd+fn+delete: delete to line end.
Public Sub TypewriterCtrlDelete()
    If Locked Then Blocked: Exit Sub
    #If Mac Then
        If Selection.Type = wdSelectionIP Then Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Delete
    #Else
        Selection.Delete Unit:=wdWord, Count:=1
    #End If
End Sub

' Mac Option+fn+delete: delete next word. (Alt+Delete does nothing on Windows.)
Public Sub TypewriterAltDelete()
    If Locked Then Blocked: Exit Sub
    #If Mac Then
        Selection.Delete Unit:=wdWord, Count:=1
    #End If
End Sub

' ---------- blocked commands ----------
' A macro with the same name as a built-in Word command replaces it, so these
' catch the ribbon buttons, right-click menu and any other shortcut for them.
' When typewriter mode is off they do the normal thing.

Public Sub EditClear()
    If Locked Then Blocked Else Selection.Delete
End Sub

Public Sub DeleteWord()
    If Locked Then Blocked Else Selection.Delete Unit:=wdWord, Count:=1
End Sub

Public Sub DeleteBackWord()
    If Locked Then Blocked Else Selection.Delete Unit:=wdWord, Count:=-1
End Sub

' Ctrl+X / Cmd+X
Public Sub EditCut()
    If Locked Then Blocked Else Selection.Cut
End Sub

' Ctrl+Z / Cmd+Z
Public Sub EditUndo()
    If Locked Then
        Blocked
    Else
        On Error Resume Next
        ActiveDocument.Undo
    End If
End Sub

' Find and Replace
Public Sub EditReplace()
    If Locked Then
        Blocked
    Else
        On Error Resume Next
        Dialogs(wdDialogEditReplace).Show
    End If
End Sub
