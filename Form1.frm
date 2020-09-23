VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   7440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   13361
      _Version        =   393217
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As XMLTree.XMLToTree
On Error GoTo ErrHandler
Set x = New XMLTree.XMLToTree
d1.CancelError = True
d1.ShowOpen
x.PlantTree TreeView1, d1.FileName
ErrHandler:
End Sub

Private Sub Command2_Click()
TreeView1.Nodes.Clear
End Sub
