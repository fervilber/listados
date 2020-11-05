VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   255
      TabIndex        =   3
      Top             =   270
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   5010
      TabIndex        =   2
      Top             =   1440
      Width           =   1065
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   2430
      TabIndex        =   1
      Top             =   930
      Width           =   2115
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   225
      TabIndex        =   0
      Top             =   780
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmPath.Show
End Sub
