VERSION 5.00
Begin VB.Form Renomeia 
   Caption         =   "Renomeia"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Pastas 
      Height          =   5715
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Renomear"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox inicial 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Prefixo 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.FileListBox arquivos 
      Height          =   4770
      Left            =   3000
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inicial"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prefixo"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Renomeia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ct = inicial
For i = 0 To arquivos.ListCount - 1
    If arquivos.Selected(i) Then
        Nome_old = Pastas.Path & "\" & arquivos.List(i)
        extensao = Right(arquivos.List(i), 4)
        Nome_new = Pastas.Path & "\" & Prefixo & Format(ct, "000") & extensao
        'MsgBox Nome_old & vbCrLf & Nome_new
        ct = ct + 1
        Name Nome_old As Nome_new
    End If
Next
arquivos.Refresh
End Sub

Private Sub Drive_Change()
Pastas.Path = Drive.Drive
End Sub

Private Sub Pastas_Click()
arquivos.Path = Pastas.Path
End Sub
