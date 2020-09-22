VERSION 5.00
Begin VB.Form ScoreBest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score Best 10"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "ScoreBest.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   405
      Left            =   780
      TabIndex        =   2
      Top             =   2955
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   405
      Left            =   2700
      TabIndex        =   0
      Top             =   2955
      Width           =   1455
   End
   Begin VB.Label Names 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Dates 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Scores 
      BackStyle       =   0  'Transparent
      Caption         =   "sd"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   690
      TabIndex        =   3
      Top             =   0
      Width           =   705
   End
   Begin VB.Label Ids 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1 > "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "ScoreBest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub onload()

For i = 1 To 10
    If i = thisscore Then
        Scores(i).ForeColor = vbBlue
    End If
    Ids(i).caption = i & "-"
    Names(i).caption = iTopTen(i).Name
    Scores(i).caption = iTopTen(i).Score
    Dates(i).caption = iTopTen(i).Date
Next i

thisscore = 0

End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
For i = 1 To 10
    iTopTen(i).Date = ""
    iTopTen(i).Name = ""
    iTopTen(i).Score = 0
Next i
SaveTopTen
onload
End Sub

Private Sub Form_Load()

For i = 1 To 10
    Load Ids(i)
    Load Scores(i)
    Load Dates(i)
    Load Names(i)
    Ids(i).top = Ids(i - 1).top + Ids(i - 1).Height
    Scores(i).top = Scores(i - 1).top + Scores(i - 1).Height
    Dates(i).top = Dates(i - 1).top + Dates(i - 1).Height
    Names(i).top = Names(i - 1).top + Names(i - 1).Height
    Ids(i).visible = True
    Names(i).visible = True
    Dates(i).visible = True
    Scores(i).visible = True
Next i

    Ids(0).visible = False
    Names(0).visible = False
    Dates(0).visible = False
    Scores(0).visible = False

onload
End Sub

