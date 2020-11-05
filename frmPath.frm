VERSION 5.00
Begin VB.Form Lee_Dir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lee_directorio"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "frmPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnLimpia 
      Caption         =   "Limpiafichero"
      Height          =   330
      Left            =   3630
      TabIndex        =   6
      Top             =   3495
      Width           =   1485
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   2715
      TabIndex        =   3
      Top             =   -15
      Width           =   2580
   End
   Begin VB.DirListBox dirPath 
      Height          =   3015
      Left            =   -15
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.DriveListBox drvPath 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton cmdFichero 
      Height          =   315
      Left            =   3615
      Picture         =   "frmPath.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4305
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Para leer el directorio actual pulsar el boton"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   4350
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   $"frmPath.frx":067C
      Height          =   780
      Left            =   45
      TabIndex        =   4
      Top             =   3525
      Width           =   3540
   End
End
Attribute VB_Name = "Lee_Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnLimpia_Click()
Dim NFI As Integer
Dim NomFichero As String

NomFichero = "Listados.txt"
NFI = FreeFile
Open App.Path & "\" & NomFichero For Output As #NFI
Close #NFI

End Sub


Private Sub cmdFichero_Click()
Dim NFI As Integer
Dim NomFichero As String
Dim i As Integer, k As Integer
Dim tamano 'Tamaño de los archivos

NomFichero = "Listados.txt"
NFI = FreeFile

Open App.Path & "\" & NomFichero For Append As #NFI


    For i = 0 To Me.File1.ListCount - 1
        
        Me.File1.ListIndex = i
        tamano = FileLen(Me.dirPath.Path & "\" & Me.File1.FileName)
        tamano = Format(tamano / 1024, "0.00")
        FILEDATE()
        
        'Escribe en el fichero el nombre la ruta y el tamaño del archivo
        
        Print #NFI, Me.File1.FileName & Chr(9) & Me.dirPath.Path & Chr(9) & tamano
    Next i

Close NFI
End Sub



Private Sub Command1_Click()

End Sub

Private Sub dirPath_Change()
File1.Path = dirPath.Path
File1.Refresh
'MsgBox Me.dirPath.ListCount
End Sub

Private Sub drvPath_Change()
  dirPath.Path = drvPath.Drive
  dirPath.Refresh
End Sub

Private Sub Form_Activate()
  'txtFichero.Text = Path
End Sub

Private Sub Form_Load()
  'txtFichero.Text = Path
'  mSoloDirectorio = False
'  bStatus = True
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  'Evento.ValorChange dirPath.Path & IIf(Right(dirPath.Path, 1) = "\", "", "\") & txtFichero.Text
'  Hide
'  bStatus = False
'End Sub


