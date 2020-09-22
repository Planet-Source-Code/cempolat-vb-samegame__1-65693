Attribute VB_Name = "main"
Public Statistics(1 To 5) As Integer
Public whited As Boolean
Public Const N_ROWS = 20
Public Const N_COLS = 10
Public Const mCELLHEIGHT = 400
Public Const mCELLWIDTH = 400
Public Const MARKTEXT = "MARK: "
Public Const POINTTEXT = "POINT: "
Public Const SCORETEXT = "SCORE: "
Public cellcnt As Integer
Public ncells As Integer
Public thisscore As Integer
Public undoscore As Long

Public SelectedPrinterName As String
Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private m_lApiReturn As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260
Private m_oEventSink As Object

Public Enum edialog
    enSaveFile = 0
    enOpenFile = 1
End Enum

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "comdlg32" _
    Alias "GetOpenFileNameA" (File As OPENFILENAME) As Long
    
Public Declare Function GetSaveFileName Lib "comdlg32" _
    Alias "GetSaveFileNameA" (File As OPENFILENAME) As Long

    
Private Declare Function GetFileTitle Lib "comdlg32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum



Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0&
    CDERR_STRUCTSIZE = &H1&
    CDERR_INITIALIZATION = &H2&
    CDERR_NOTEMPLATE = &H3&
    CDERR_NOHINSTANCE = &H4&
    CDERR_LOADSTRFAILURE = &H5&
    CDERR_FINDRESFAILURE = &H6&
    CDERR_LOADRESFAILURE = &H7&
    CDERR_LOCKRESFAILURE = &H8&
    CDERR_MEMALLOCFAILURE = &H9&
    CDERR_MEMLOCKFAILURE = &HA&
    CDERR_NOHOOK = &HB&
    CDERR_REGISTERMSGFAIL = &HC&

    PDERR_PRINTERCODES = &H1000&
    PDERR_SETUPFAILURE = &H1001&
    PDERR_PARSEFAILURE = &H1002&
    PDERR_RETDEFFAILURE = &H1003&
    PDERR_LOADDRVFAILURE = &H1004&
    PDERR_GETDEVMODEFAIL = &H1005&
    PDERR_INITFAILURE = &H1006&
    PDERR_NODEVICES = &H1007&
    PDERR_NODEFAULTPRN = &H1008&
    PDERR_DNDMMISMATCH = &H1009&
    PDERR_CREATEICFAILURE = &H100A&
    PDERR_PRINTERNOTFOUND = &H100B&
    PDERR_DEFAULTDIFFERENT = &H100C&

    CFERR_CHOOSEFONTCODES = &H2000&
    CFERR_NOFONTS = &H2001&
    CFERR_MAXLESSTHANMIN = &H2002&

    FNERR_FILENAMECODES = &H3000&
    FNERR_SUBCLASSFAILURE = &H3001&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_BUFFERTOOSMALL = &H3003&

    CCERR_CHOOSECOLORCODES = &H5000&
End Enum

' Hook and notification support:
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
'// Structure used for all file based OpenFileName notifications
Private Type OFNOTIFY
    hdr As NMHDR
    lpOFN As Long           ' Long pointer to OFN structure
    pszFile As String ';        // May be NULL
End Type

'// Structure used for all object based OpenFileName notifications
Private Type OFNOTIFYEX
    hdr As NMHDR
    lpOFN As Long       ' Long pointer to OFN structure
    psf As Long
    LPVOID As Long          '// May be NULL
End Type

Private Type OFNOTIFYshort
    hdr As NMHDR
    lpOFN As Long
End Type

' Messages:
Private Const WM_INITDIALOG = &H110
Private Const WM_NOTIFY = &H4E
Private Const WM_USER = &H400
Private Const WM_GETDLGCODE = &H87
Private Const WM_NCDESTROY = &H82


' Notification codes:
Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_LAST = (H_MAX - 699)

'// Notifications when Open or Save dialog status changes
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)

Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_LAST = (WM_USER + 200)

Private Const DWL_MSGRESULT = 0
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public SmartRand As Boolean

Public Enum eCreate
    eNew = 0
    eLoad = 1
End Enum

Public Enum eLoadType
    eUndo = 0
    eReplay = 1
End Enum

Public Type tttype
    Score As Long
    Date As String
    Name As String
End Type

Public Type cellattr
    visible As Boolean
    left As Integer
    top As Integer
    caption As String
End Type

Public basestring(1 To N_ROWS * N_COLS) As String
Public undostring(1 To N_ROWS * N_COLS) As cellattr
Public replaystring(1 To N_ROWS * N_COLS) As cellattr

Public iTopTen(1 To 10) As tttype
Public Sub updatestatistics()
Dim total As Integer
For i = 1 To 5
    Statistics(i) = 0
Next i

For i = 1 To ncells
    If Form1.Cells(i).visible = True Then
        total = Asc(Form1.Cells(i).caption) - 64
        Statistics(total) = Statistics(total) + 1
    End If
Next i
End Sub

Public Function CheckScore(iscore As Long) As Boolean
Dim intopten As Boolean
intopten = False

For i = 1 To 10

    If iscore > iTopTen(i).Score Then
        thisscore = i
        For j = 9 To i Step -1
            iTopTen(j + 1).Name = iTopTen(j).Name
            iTopTen(j + 1).Date = iTopTen(j).Date
            iTopTen(j + 1).Score = iTopTen(j).Score
        Next j
        
        iTopTen(i).Name = InputBox("Enter your name", "You have achived a high score", " ")
        iTopTen(i).Score = iscore
        iTopTen(i).Date = "" & Date
        intopten = True
        Exit For
    End If

Next i
SaveTopTen
CheckScore = intopten
End Function

Public Sub ReadTopTen()
Dim val As Variant

For i = 1 To 10

    val = GetSetting(appname:="SameGame", section:="TopTen", Key:="Score" & i, Default:="XXX")
    If val <> "XXX" Then
        iTopTen(i).Score = val
    Else
        iTopTen(i).Score = 0
    End If

    val = GetSetting(appname:="SameGame", section:="TopTen", Key:="Date" & i, Default:="XXX")
    If val <> "XXX" Then
        iTopTen(i).Date = val
    Else
        iTopTen(i).Date = ""
    End If

    val = GetSetting(appname:="SameGame", section:="TopTen", Key:="Name" & i, Default:="XXX")
    If val <> "XXX" Then
        iTopTen(i).Name = val
    Else
        iTopTen(i).Name = 0
    End If

Next i

End Sub
Public Sub SaveTopTen()

For i = 1 To 10
    SaveSetting "SameGame", "TopTen", "Score" & i, iTopTen(i).Score
    SaveSetting "SameGame", "TopTen", "Name" & i, iTopTen(i).Name
    SaveSetting "SameGame", "TopTen", "Date" & i, iTopTen(i).Date
Next i

End Sub

Public Function basecolor(letter As String) As OLE_COLOR
Select Case letter
    Case "A"
        basecolor = vbBlue
    Case "B"
        basecolor = vbRed
    Case "C"
        basecolor = vbMagenta
    Case "D"
        basecolor = vbYellow
    Case "E"
        basecolor = vbCyan
    Case Else
    '    MsgBox letter
End Select

End Function

Public Function DialogHook( _
        ByVal hDlg As Long, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
    )
Dim tNMH As NMHDR
Dim tOFNs As OFNOTIFYshort
Dim tOF As OPENFILENAME

    If Not (m_oEventSink Is Nothing) Then
        Select Case msg
        Case WM_INITDIALOG
            DialogHook = m_oEventSink.InitDialog(hDlg)
        Case WM_NOTIFY
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            Select Case tNMH.code
            Case CDN_SELCHANGE
                ' Changed selected file:
                DialogHook = m_oEventSink.FileChange(hDlg)
            Case CDN_FOLDERCHANGE
                ' Changed folder:
                DialogHook = m_oEventSink.FolderChange(hDlg)
            Case CDN_FILEOK
                ' Clicked OK:
                If Not m_oEventSink.ConfirmOK() Then
                    SetWindowLong hDlg, DWL_MSGRESULT, 1
                    DialogHook = 1
                Else
                    SetWindowLong hDlg, DWL_MSGRESULT, 0
                End If
            Case CDN_HELP
                ' Help clicked
            Case CDN_TYPECHANGE
                DialogHook = m_oEventSink.TypeChange(hDlg)
            Case CDN_INCLUDEITEM
                ' Hmmm
            End Select
        Case WM_NCDESTROY
            m_oEventSink.DialogClose
        End Select
    End If
End Function


Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = m_lApiReturn
End Property

Private Function lHookAddress(lPtr As Long) As Long
    'Debug.Print lPtr
    lHookAddress = lPtr
End Function
Private Function StrZToStr(s As String) As String
    StrZToStr = left$(s, lstrlen(s))
End Function

Function VBGetFileName(DialogType As edialog, _
                           Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long, _
                           Optional Hook As Boolean = False, _
                           Optional EventSink As Object _
                        ) As Boolean
            
    Dim opfile As OPENFILENAME, s As String

    m_lApiReturn = 0


With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    If (Hook) Then
        ''HookedDialog = Me
        '.lpfnHook = lHookAddress(AddressOf DialogHookFunction)
        '.flags = .flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
        ''Set m_oEventSink = EventSink
    End If
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    If DialogType = enOpenFile Then
        m_lApiReturn = GetOpenFileName(opfile)
    Else
        m_lApiReturn = GetSaveFileName(opfile)
    End If
    
    Set m_oEventSink = Nothing
    'ClearHookedDialog
    Select Case m_lApiReturn
    Case 1
        VBGetFileName = True
        Filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
        ' Cancelled:
        VBGetFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    Case Else
        ' Extended error:
        VBGetFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

