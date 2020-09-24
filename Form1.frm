VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Search, Prompt, & Replace Within ANY file(s)"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search && Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose File to Search"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "Good"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtSeek 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Evil"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Replace With:"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Search For:"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Example which will
'Find & Replace a String within Any file(s)
'Binaries  included ...
'
'by Pbryan^(2K And 2)
'

Dim ChangeThisFile As String

Private Sub Command1_Click() ' Choose File
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    
        ChangeThisFile = CommonDialog1.FileName
        Label3.Caption = "File to Examine = " & ChangeThisFile
        Command2.Enabled = True: txtSeek.Enabled = True
        txtNew.Enabled = True
            Exit Sub
    End If
ErrHandler:
    Exit Sub

End Sub

Private Sub Command2_Click() ' Find And Replace
    MakeChanges ChangeThisFile, txtSeek, txtNew
    MsgBox "There were " & NumFound & " Replacement(s)!", vbOKOnly, "Search Complete!"
    End
End Sub

Private Sub Command3_Click() ' quit
    End
End Sub
