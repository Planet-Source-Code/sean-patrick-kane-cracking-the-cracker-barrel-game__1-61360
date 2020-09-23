VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cracking the Cracker Barrel Game"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Initial Peg Missing"
      Height          =   2655
      Left            =   2295
      TabIndex        =   8
      Top             =   720
      Width           =   4575
      Begin VB.OptionButton optTriangle 
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   2160
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "2"
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   22
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   21
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "4"
         Height          =   195
         Index           =   4
         Left            =   3060
         TabIndex        =   20
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "5"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   19
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "6"
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   18
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "7"
         Height          =   195
         Index           =   7
         Left            =   1725
         TabIndex        =   17
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "8"
         Height          =   195
         Index           =   8
         Left            =   2595
         TabIndex        =   16
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "9"
         Height          =   195
         Index           =   9
         Left            =   3480
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "10"
         Height          =   195
         Index           =   10
         Left            =   1320
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "11"
         Height          =   195
         Index           =   11
         Left            =   2220
         TabIndex        =   13
         Top             =   1200
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "12"
         Height          =   195
         Index           =   12
         Left            =   3120
         TabIndex        =   12
         Top             =   1200
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "13"
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "14"
         Height          =   195
         Index           =   14
         Left            =   2640
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "15"
         Height          =   195
         Index           =   15
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      Height          =   375
      Left            =   4627
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3322
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstFinal 
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   4440
      Width           =   9015
   End
   Begin VB.Label Label3 
      Caption         =   "Tip: Double click a line to copy it to the clipboard"
      Height          =   255
      Left            =   2775
      TabIndex        =   6
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Label cntFinal 
      Caption         =   "0"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label cntBoards 
      Caption         =   "0"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Boards that have been completed:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Boards that still have moves left:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function BeginGames()
'This is the function that's responsible for parsing through lstBoards
Dim curarray As board, curpath As String, i As Integer, j As Integer

Do
    DoEvents 'Let the program breathe

    'Convert the flag into a board array
    curarray = FlagToArray(lstBoards(0).flag)
    curpath = lstBoards(0).past
    
    'Now that curarray is prepared...let's send all the arguments to EvalBoard
    If lstBoards(0).delete = False Then EvalBoard curarray, curpath
    
    'Remove item 0
    tmpBoards = lstBoards
    If UBound(lstBoards) <> 0 Then
        ReDim lstBoards(UBound(lstBoards) - 1)
    Else
        ReDim lstBoards(0)
    End If
    For i = 1 To UBound(tmpBoards)
        lstBoards(i - 1) = tmpBoards(i)
    Next i
    
    cntBoards = UBound(lstBoards)
    If UBound(lstBoards) = 0 And lstBoards(0).past = "" Then Exit Function
    
    'Check for duplicates
    For i = 1 To UBound(lstBoards)
        If lstBoards(i).flag = lstBoards(0).flag Then
            lstBoards(i).delete = True
        End If
    Next i
Loop
End Function

Private Sub cmdInstructions_Click()
frmInstructions.Show
End Sub

Private Sub cmdStart_Click()
Dim i As Integer, tmpcount As Long

lstFinal.Clear
cntFinal = 0

For i = 1 To 15
    If optTriangle(i).Value = False Then tmpcount = tmpcount + (2 ^ i)
Next i

ReDim lstBoards(0)
lstBoards(0).flag = tmpcount 'This command alone will start a chain reaction that will attempt every single possible game with this initial configuration

cntBoards = UBound(lstBoards)
BeginGames
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstFinal_DblClick()
Clipboard.SetText lstFinal.List(lstFinal.ListIndex)
End Sub

