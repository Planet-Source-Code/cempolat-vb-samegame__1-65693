VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Same Game"
   ClientHeight    =   4290
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4005
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   503
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "MARK:"
            TextSave        =   "MARK:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8943
            MinWidth        =   5292
            Text            =   "POINT:"
            TextSave        =   "POINT:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "SCORE:"
            TextSave        =   "SCORE:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Cells 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   400
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu NewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Replay 
         Caption         =   "&Restart Game"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mOpenGame 
         Caption         =   "&Open Game"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSaveGame 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu mUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Score 
         Caption         =   "&Score"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mInformation 
         Caption         =   "&Information"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ssa 
         Caption         =   "-"
      End
      Begin VB.Menu nExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Same Game by Cem POLAT 2004
'
' cempolat@mailcity.com
'
Private marks As Long
Private points As Long
Private Scores As Long
Private Filename As String
Private Sub GenRandLetters()
Dim val As Integer
Dim norder(1 To 5) As Integer

SmartRand = True

If SmartRand = True Then
       
        For i = 1 To 5
ignoree:
            Randomize
            val = Int((5 * Rnd) + 1)
            
            For j = 1 To 5
                If norder(j) = val Then GoTo ignoree
            Next j
            norder(i) = val
        Next i
    
    
    For i = 1 To ncells
        Randomize
        val = Int((100 * Rnd) + 1)
            
        Select Case val
        
        Case 1 To 16
            basestring(i) = Chr(64 + norder(1))
        Case 17 To 32
            basestring(i) = Chr(64 + norder(2))
        Case 33 To 70
            basestring(i) = Chr(64 + norder(3))
        Case 71 To 81
            basestring(i) = Chr(64 + norder(4))
        Case Else
            basestring(i) = Chr(64 + norder(5))
        End Select
        
    Next i
 
Else
    For i = 1 To ncells
        Randomize
        val = Int((5 * Rnd) + 1)
        basestring(i) = Chr(64 + val)
    Next i
End If

End Sub

Private Sub loadgame()
Dim x As Byte
Dim y As Integer
Dim fsize As Integer

y = 1
fsize = ncells / 2

Open Filename For Binary Access Read As #1 'Len = 101

For i = 1 To fsize
    Get #1, i, x
    basestring(y) = Chr((x Mod 8) + 65)
    y = y + 1
    basestring(y) = Chr(Int(x / 8) + 65)
    y = y + 1
Next

Close #1

End Sub
Private Sub savegame()
Dim x As Byte
Dim y As Integer

y = 1
Open Filename For Binary Access Write As #1 'Len = 101

For i = 1 To ncells Step 2
    x = Asc(replaystring(i).caption) - 65
    x = x + 8 * (Asc(replaystring(i + 1).caption) - 65)
    Put #1, y, x
    y = y + 1
Next

Close #1

End Sub

Private Sub CreateBoard(kind As eCreate)
Scores = 0
marks = 0
points = 0
whited = False
undoscore = 0

mUndo.Enabled = False

Select Case kind
    Case eLoad

    Case eNew
        GenRandLetters
End Select

For i = 1 To N_COLS
    For j = 1 To N_ROWS
        idcell = ((i - 1) * N_ROWS) + j
        Cells(idcell).left = Cells(0).Width * (j - 1)
        Cells(idcell).top = Cells(0).Height * (i - 1)
        Cells(idcell).caption = basestring(idcell)
        Cells(idcell).BackColor = basecolor(basestring(idcell))
        Cells(idcell).visible = True
    Next j
Next i
savemove eReplay
End Sub

Private Sub savemove(Optional loadtype As eLoadType = eUndo)

For i = 1 To ncells
    
    If loadtype = eUndo Then
        undostring(i).caption = Cells(i).caption
        undostring(i).left = Cells(i).left
        undostring(i).top = Cells(i).top
        undostring(i).visible = Cells(i).visible
    Else
        replaystring(i).caption = Cells(i).caption
        replaystring(i).left = Cells(i).left
        replaystring(i).top = Cells(i).top
        replaystring(i).visible = Cells(i).visible
    End If

Next i


End Sub
Private Sub loadmove(Optional loadtype As eLoadType = eUndo)

For i = 1 To ncells
    
    If loadtype = eUndo Then
        Cells(i).caption = undostring(i).caption
        Cells(i).BackColor = basecolor(undostring(i).caption)
        Cells(i).left = undostring(i).left
        Cells(i).top = undostring(i).top
        Cells(i).visible = undostring(i).visible
    Else
        Cells(i).caption = replaystring(i).caption
        Cells(i).BackColor = basecolor(replaystring(i).caption)
        Cells(i).left = replaystring(i).left
        Cells(i).top = replaystring(i).top
        Cells(i).visible = replaystring(i).visible
    End If
    
Next i

End Sub

Private Sub scorecalc()
Scores = Scores + points
mUndo.Enabled = True
points = 0
marks = 0
End Sub
Private Sub pointcalc()
points = (marks * (marks - 3)) + 4
End Sub
Private Sub whitecnt()
Dim i As Integer
cellcnt = 0

For i = 1 To ncells
    If Cells(i).BackColor = vbWhite Then
        cellcnt = cellcnt + 1
    End If
Next i
marks = cellcnt
pointcalc
End Sub
Private Sub updatestat()
Status.Panels(1).Text = MARKTEXT & marks
Status.Panels(2).Text = POINTTEXT & points
Status.Panels(3).Text = SCORETEXT & Scores
End Sub

Private Sub CheckValidMove()
Dim i As Integer
Dim j As Integer
Dim idcell As Integer

For i = 1 To N_COLS
    For j = 1 To N_ROWS
        idcell = ((i - 1) * N_ROWS) + j
        If Cells(idcell).top < -2 Or Cells(idcell).left < -2 Then GoTo ignoore
        If ValidMoveTest(idcell) Then Exit Sub
ignoore:
    Next j
Next i
updatestat
If CheckScore(Scores) = False Then
    MsgBox "Game Over"
Else
    ScoreBest.Show vbModal, Me
End If
undoscore = 0
mUndo.Enabled = False
End Sub
Private Function ValidMoveTest(Index As Integer) As Boolean

Dim cellx As Integer
Dim celly As Integer
Dim CELLWIDTH As Integer
Dim CELLHEIGHT As Integer
Dim lefti As Integer
Dim topi As Integer
Dim colori As OLE_COLOR
Dim cellcolor As OLE_COLOR
Dim i As Integer
Dim celletter As String * 1
Dim letteri As String * 1

cellx = Cells(Index).left
celly = Cells(Index).top
CELLWIDTH = 2 * Cells(Index).Width
CELLHEIGHT = 2 * Cells(Index).Height
celletter = Cells(Index).caption

For i = 1 To Cells.Count - 1
    lefti = Cells(i).left
    topi = Cells(i).top
    letteri = Cells(i).caption
    If i = Index Then GoTo ignoree
    If Abs(lefti - cellx) < CELLWIDTH And topi = celly Or Abs(topi - celly) < CELLHEIGHT And lefti = cellx Then
        If celletter = letteri Then
            ValidMoveTest = True
            Exit Function
        End If
    End If
    
ignoree:
Next i

ValidMoveTest = False
End Function

Private Sub ExplodeBatch()
Dim i As Integer
Dim j As Integer
Dim idcell As Integer

For i = 1 To N_COLS
    For j = 1 To N_ROWS
        idcell = ((i - 1) * N_ROWS) + j
        
        If Cells(idcell).BackColor = vbWhite Then
            Explode idcell
        End If
    Next j
Next i
whited = False

End Sub
Private Sub Explode(Index As Integer)
Dim i As Integer
Dim hh As Long
Dim ww As Long
Dim lastrow As Boolean

Cells(Index).BackColor = basecolor(Cells(Index).caption)
Cells(Index).ForeColor = vbBlack
Cells(Index).visible = False

hh = Cells(Index).top
ww = Cells(Index).left

Cells(Index).top = -2000
Cells(Index).left = -2000

lastrow = False

For i = 1 To ncells
    If i = Index Then GoTo ignoree
    If Cells(i).left = ww Then
    If Cells(i).top < hh And Cells(i).top > -2 Then
        Cells(i).top = Cells(i).top + mCELLHEIGHT

        lastrow = True
    End If
    End If
ignoree:
Next i
    
If lastrow = False Then
    If hh >= (9 * mCELLHEIGHT) Then
   
       For i = 1 To ncells
          If Cells(i).top > -2 And Cells(i).left > ww Then
            Cells(i).left = Cells(i).left - mCELLWIDTH
          End If
       Next i
        
    End If
End If
End Sub
Private Sub CancelExplode()
Dim i As Integer
Dim j As Integer
Dim idcell As Integer

For i = 1 To N_COLS
    For j = 1 To N_ROWS
        idcell = ((i - 1) * N_ROWS) + j
       If Cells(idcell).BackColor = vbWhite Then
            Cells(idcell).BackColor = basecolor(Cells(idcell).caption)
            Cells(idcell).ForeColor = vbBlack
        End If
    Next j
Next i
whited = False
marks = 0
points = 0
updatestat
End Sub

Private Sub onload()
Dim i As Integer
Dim j As Integer
Dim idcell As Integer

ncells = N_ROWS * N_COLS

If Cells.Count = 1 Then
    For i = 1 To ncells
        Load Cells(i)
    Next i
    Cells(0).visible = False
End If

CreateBoard eNew
updatestat
End Sub

Private Sub About_Click()
MsgBox "Same Game" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "By Cem POLAT 2004"
End Sub

Private Sub Cells_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    If whited = True Then
        CancelExplode
    End If
Else

    If Cells(Index).BackColor = vbWhite Then
        savemove
        undoscore = Scores
        ExplodeBatch
        scorecalc
        CheckValidMove
    Else
        
        If whited = True Then
            CancelExplode
        Else
            neighbourtest (Index)
            whitecnt
        End If
    End If
updatestat
End If
End Sub

Private Sub Form_Load()
Cells(0).Height = mCELLHEIGHT
Cells(0).Width = mCELLWIDTH
ReadTopTen
onload
End Sub

Private Sub mInformation_Click()
Load Information
Information.Show vbModal, Me
End Sub

Private Sub mOpenGame_Click()
Dim fname As String

VBGetFileName enOpenFile, fname, , False, "Samegame Files|*.sav|All Files|*.*", , CurDir, "Open Game File", "*.sav;*.*", Me.hwnd

If Len(fname) > 0 Then
    Filename = fname
    loadgame
    CreateBoard eLoad
End If

End Sub

Private Sub mSaveGame_Click()
Dim fname As String

VBGetFileName enSaveFile, fname, , False, "Samegame Files|*.sav|All Files|*.*", , CurDir, "Save Game File", "*.sav;*.*", Me.hwnd

If Len(fname) > 0 Then
    Filename = fname
    savegame
End If

End Sub

Private Sub mUndo_Click()
loadmove
Scores = undoscore
undoscore = 0
mUndo.Enabled = False
updatestat
End Sub

Private Sub NewGame_Click()
CreateBoard eNew
updatestat
End Sub

Private Sub nExit_Click()
End
End Sub


Private Sub Replay_Click()
Scores = 0
updatestat
loadmove eReplay
End Sub

Private Sub Score_Click()
Load ScoreBest
ScoreBest.Show vbModal, Me
End Sub

Sub neighbourtest(Index As Integer)
Dim cellx As Long
Dim celly As Long
Dim CELLWIDTH As Long
Dim CELLHEIGHT As Long
Dim lefti As Long
Dim topi As Long
Dim colori As OLE_COLOR
Dim i As Integer

Dim celletter As String * 1
Dim letteri As String * 1

Dim firstmatch As Boolean

firstmatch = False
' Recursive search algorithm for valid neighbours
cellx = Me.Cells(Index).left
celly = Me.Cells(Index).top
CELLWIDTH = 2 * Me.Cells(Index).Width
CELLHEIGHT = 2 * Me.Cells(Index).Height
celletter = Me.Cells(Index).caption

For i = 0 To Me.Cells.Count - 1
    lefti = Me.Cells(i).left
    topi = Me.Cells(i).top
    colori = Me.Cells(i).BackColor
    letteri = Me.Cells(i).caption
    If colori = vbWhite Then GoTo ignoree
    If i = Index Then GoTo ignoree
    If Abs(lefti - cellx) < CELLWIDTH And topi = celly Or Abs(topi - celly) < CELLHEIGHT And lefti = cellx Then
        If celletter = letteri Then
            If Me.Cells(Index).visible = False Then GoTo ignoree
            If firstmatch = False Then
                Me.Cells(Index).BackColor = vbWhite
                Me.Cells(Index).ForeColor = basecolor(Me.Cells(Index).caption)
                firstmatch = True
            End If
            Me.Cells(i).BackColor = vbWhite
            Me.Cells(i).ForeColor = basecolor(Me.Cells(i).caption)
            neighbourtest (i)
        End If
    End If
    
ignoree:
Next i

whited = firstmatch

End Sub
