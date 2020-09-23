VERSION 5.00
Begin VB.Form frmInstructions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmInstructions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3045
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   $"frmInstructions.frx":030A
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label Label4 
      Caption         =   "->Double click the listbox entries you like to copy them onto your clipboard."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   $"frmInstructions.frx":03A7
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   $"frmInstructions.frx":0463
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmInstructions.frx":052C
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
