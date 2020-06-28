VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "..."
      Height          =   255
      Left            =   3996
      TabIndex        =   2
      Top             =   372
      Width           =   255
   End
   Begin VB.TextBox txtExeFile 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "実行ファイル名"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    SaveSetting "NCArray", _
                "Settings", _
                "Editor", _
                txtExeFile.Text

    Unload Me

End Sub

Private Sub cmdSee_Click()

    frmFileList.Show vbModal
    With frmFileList
        If .filList.FileName <> "" Then
            txtExeFile.Text = _
                .dirList.Path & "\" & .filList.FileName
        End If
    End With
    Unload frmFileList

End Sub

Private Sub Form_Load()

    Top = frmMain.Top + 300
    Left = frmMain.Left + 300

    Caption = "エディタの設定"

    txtExeFile = GetSetting("NCArray", _
                         "Settings", _
                         "Editor")

End Sub
