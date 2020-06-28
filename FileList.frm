VERSION 5.00
Begin VB.Form frmFileList 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "FileList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.DriveListBox drvList 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.DirListBox dirList 
      Height          =   2610
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.FileListBox filList 
      Height          =   3150
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    'dirList.Path が現在選択されているディレクトリと
    '異なる場合は、dirList.Path を更新します。
    '同じ場合は、実行します。
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub
    End If

    Me.Visible = False

End Sub

Private Sub dirList_Change()


   ' ファイル リスト ボックスをディレクトリ リスト ボックスと連動して更新
   ' します。
   filList.Path = dirList.Path

End Sub

Private Sub drvList_Change()

   On Error GoTo DriveHandler

   ' 新しいドライブが選択された場合は、Dirボックスの
   ' 表示を更新します。
   dirList.Path = drvList.Drive
   Exit Sub

   ' エラーが発生した場合は、drvList.Drive の値を
   ' dirList.Path のドライブに戻します。
DriveHandler:
   drvList.Drive = dirList.Path
   Exit Sub

End Sub

Private Sub Form_Load()

    Caption = "ファイルの選択"

    With frmEdit
        Top = .Top + 300
        Left = .Left + 300
    End With

    cmdOK.Default = True 'デフォルトボタンの設定

End Sub
