VERSION 5.00
Begin VB.Form Information 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INFORMATION"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   ControlBox      =   0   'False
   Icon            =   "Information.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2835
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   570
      TabIndex        =   1
      Top             =   2130
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   350
      TabIndex        =   5
      Top             =   1530
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   350
      TabIndex        =   4
      Top             =   1176
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   350
      TabIndex        =   3
      Top             =   824
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   350
      TabIndex        =   2
      Top             =   472
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   350
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
updatestatistics
Label1(0).Caption = "(A) Blue   ..... "
Label1(1).Caption = "(B) Red    ..... "
Label1(2).Caption = "(C) Purple ..... "
Label1(3).Caption = "(D) Yellow ..... "
Label1(4).Caption = "(E) Cyan   ..... "

For i = 0 To 4
    Label1(i).Caption = Label1(i).Caption & Statistics(i + 1)
Next i
End Sub
