VERSION 5.00
Begin VB.Form frmFileList 
   BorderStyle     =   3  '�Œ��޲�۸�
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
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.DriveListBox drvList 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��ݾ�"
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

    'dirList.Path �����ݑI������Ă���f�B���N�g����
    '�قȂ�ꍇ�́AdirList.Path ���X�V���܂��B
    '�����ꍇ�́A���s���܂��B
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub
    End If

    Me.Visible = False

End Sub

Private Sub dirList_Change()


   ' �t�@�C�� ���X�g �{�b�N�X���f�B���N�g�� ���X�g �{�b�N�X�ƘA�����čX�V
   ' ���܂��B
   filList.Path = dirList.Path

End Sub

Private Sub drvList_Change()

   On Error GoTo DriveHandler

   ' �V�����h���C�u���I�����ꂽ�ꍇ�́ADir�{�b�N�X��
   ' �\�����X�V���܂��B
   dirList.Path = drvList.Drive
   Exit Sub

   ' �G���[�����������ꍇ�́AdrvList.Drive �̒l��
   ' dirList.Path �̃h���C�u�ɖ߂��܂��B
DriveHandler:
   drvList.Drive = dirList.Path
   Exit Sub

End Sub

Private Sub Form_Load()

    Caption = "�t�@�C���̑I��"

    With frmEdit
        Top = .Top + 300
        Left = .Left + 300
    End With

    cmdOK.Default = True '�f�t�H���g�{�^���̐ݒ�

End Sub
