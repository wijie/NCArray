VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  '�Ȃ�
   Caption         =   "�ʕt���N"
   ClientHeight    =   4110
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5910
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  '���
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows �̊���l
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�㑵��
      Height          =   360
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnOpen"
            Object.ToolTipText     =   "�t�@�C�����J��"
            Object.Tag             =   "imgOpen"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   3960
         TabIndex        =   41
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.TextBox txtMoji 
      Height          =   264
      Left            =   960
      MaxLength       =   11
      TabIndex        =   21
      ToolTipText     =   "�ԕ���"
      Top             =   3360
      Width           =   1572
   End
   Begin VB.TextBox txtSosu 
      Height          =   270
      Left            =   960
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "�w��"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtTotal 
      Height          =   270
      Index           =   1
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   19
      ToolTipText     =   "Y�̑S��"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   0
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   7
      ToolTipText     =   "Y��WBS"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   7
      TabIndex        =   6
      ToolTipText     =   "X��WBS"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   7
      TabIndex        =   18
      ToolTipText     =   "X�̑S��"
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "�I��(&Q)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   39
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "�ر(&C)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   38
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "���s(&X)"
      Height          =   375
      Left            =   3000
      TabIndex        =   37
      Top             =   3600
      Width           =   852
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   3960
      TabIndex        =   36
      ToolTipText     =   "�o�̓t�@�C����"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox cmbStack 
      Height          =   300
      ItemData        =   "frmMain.frx":08CA
      Left            =   960
      List            =   "frmMain.frx":08CC
      TabIndex        =   23
      ToolTipText     =   "�X�^�b�N�ʒu"
      Top             =   3720
      Width           =   1692
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   2
      Left            =   960
      MaxLength       =   7
      TabIndex        =   12
      ToolTipText     =   "X�̖ʕt���s�b�`"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   3
      Left            =   960
      MaxLength       =   4
      TabIndex        =   15
      ToolTipText     =   "X�̖ʕt����"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   3
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   16
      ToolTipText     =   "Y�̖ʕt����"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   2
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   13
      ToolTipText     =   "Y�̖ʕt���s�b�`"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   1
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   10
      ToolTipText     =   "Y�̐��i���@"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame fraIdou 
      Caption         =   "�ړ���(&I)"
      Height          =   1095
      Left            =   2760
      TabIndex        =   29
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtIdou 
         Height          =   270
         Index           =   1
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   34
         ToolTipText     =   "Y�̈ړ���"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtIdou 
         Height          =   270
         Index           =   0
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   33
         ToolTipText     =   "X�̈ړ���"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTestHole 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   31
         ToolTipText     =   "�������̈ړ���"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblIdou 
         Caption         =   "���ʔ�"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTestHole 
         Caption         =   "������"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraTCode 
      Caption         =   "T����(&T)"
      Height          =   1095
      Left            =   2760
      TabIndex        =   24
      Top             =   360
      Width           =   1815
      Begin VB.TextBox txtMojiTCode 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   26
         ToolTipText     =   "�ԕ�����T�R�[�h"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtSGAG 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   28
         ToolTipText     =   "SG/AG��T�R�[�h"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblMojiTCode 
         Caption         =   "�ԕ���"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSGAG 
         Caption         =   "SG/AG"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   1
      Left            =   960
      MaxLength       =   7
      TabIndex        =   9
      ToolTipText     =   "X�̐��i���@"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame fraDataType 
      Caption         =   "�ް�����(&D)"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
      Begin VB.OptionButton optTHNT 
         Caption         =   "TH"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTHNT 
         Caption         =   "NT"
         Height          =   180
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDDE 
      BorderStyle     =   1  '����
      Caption         =   "DDE�p"
      Height          =   255
      Left            =   4800
      TabIndex        =   42
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblWBSX 
      Caption         =   "�v�a�r"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblFileName 
      Caption         =   "̧�ٖ�(&N):"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblStack 
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Caption         =   "�S��"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblMoji 
      Caption         =   "�ԕ���"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblMenzuke 
      Caption         =   "�ʕt��"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblPitch 
      Caption         =   "�ʕt�߯�"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblSize 
      Caption         =   "���i���@"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblSosu 
      Caption         =   "�w��"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "�J��(&O)"
      End
      Begin VB.Menu mnuFileStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�I��(&Q)"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "����(&O)"
      Begin VB.Menu mnuFormatTombo 
         Caption         =   "����"
      End
      Begin VB.Menu mnuFormatMoji 
         Caption         =   "�ԕ���"
      End
      Begin VB.Menu mnuFormatMBE 
         Caption         =   "�O�H�����"
      End
      Begin VB.Menu mnuFormatSG 
         Caption         =   "SG"
      End
      Begin VB.Menu mnuFormatAG 
         Caption         =   "AG"
      End
      Begin VB.Menu mnuFormatTestHole 
         Caption         =   "������"
      End
      Begin VB.Menu mnuFormatT50 
         Caption         =   "�t��Ėh�~(T50)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "°�(&T)"
      Begin VB.Menu mnuEdit 
         Caption         =   "��ި��ݒ�"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "�ް�ޮݏ��(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' �^�C�g���o�[�ɕ\������^�C�g���ݒ�p�萔
Const strTitleText As String = "�ʕt���N"
Private mstrEditor As String
Private mudtCurrentNC As DrillData
Private typNC(1) As DrillData
Private txtWBS(1) As TextBox ' ���[�N�{�[�h�T�C�Y
Private txtSize(1) As TextBox ' ���i�T�C�Y
Private txtPitch(1) As TextBox ' �ʕt���s�b�`
Private txtNumber(1) As TextBox ' �ʕt����

'*********************************************************
' �p  �r: �t�@�C�����J���_�C�A���O��\������
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub GetInputFile()

    ' CancelError�̐ݒ�͐^(True)�ł��B
    On Error GoTo ErrHandler

    With CommonDialog1
        ' �t�@�C���̑I����@��ݒ肵�܂��B
        .Filter = "���ׂẴt�@�C�� (*.*)|*.*|" & _
                  "NC�t�@�C�� (*.nc)|*.nc|" & _
                  "�f�[�^�t�@�C�� (*.dat)|*.dat"

        ' ����̑I����@���w�肵�܂��B
        .FilterIndex = 1

        ' [�ǂݎ���p�t�@�C���Ƃ��ĊJ��]�`�F�b�N�{�b�N�X��\�����Ȃ�
        ' �����̃t�@�C�����������͂ł��Ȃ��悤�ɂ���
        .Flags = cdlOFNHideReadOnly Or _
                 cdlOFNFileMustExist

        ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂��B
        .ShowOpen

        If txtFileName.Text <> "" Then
            cmdExec.Enabled = True
        End If

        Caption = strTitleText & " - " & .FileName
    '    ChDir .InitDir
        Exit Sub
    End With

ErrHandler:
    ' ���[�U�[��[�L�����Z��] �{�^�����N���b�N���܂����B
    If Err.Number = cdlCancel Then
        If Mid(Caption, Len(strTitleText) + 4) = "" Then
            Caption = strTitleText
            cmdExec.Enabled = False
        End If
    End If

End Sub

'*********************************************************
' �p  �r: �ϐ��ɒl���Z�b�g����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sSetENV()

    Set txtWBS(X) = txtInputX(0)
    Set txtWBS(Y) = txtInputY(0)
    Set txtSize(X) = txtInputX(1)
    Set txtSize(Y) = txtInputY(1)
    Set txtPitch(X) = txtInputX(2)
    Set txtPitch(Y) = txtInputY(2)
    Set txtNumber(X) = txtInputX(3)
    Set txtNumber(Y) = txtInputY(3)

    With mudtCurrentNC
        ' ���̓t�@�C�������Z�b�g����Ă�����t�@�C�������L������
        If Mid(Caption, Len(strTitleText) + 4) <> "" Then
            .strInFile = Mid(Caption, Len(strTitleText) + 4)
        End If
        If optTHNT(TH).Value = True Then
            .intNCType = TH
        Else
            .intNCType = NT
        End If

        .intSosu = Val(txtSosu) ' �w��
        .lngWBS(X) = Val(txtWBS(X)) * 100 ' X��WBS
        .lngWBS(Y) = Val(txtWBS(Y)) * 100 ' Y��WBS
        .lngSize(X) = Val(txtSize(X)) * 100 ' ���i�T�C�YX
        .lngSize(Y) = Val(txtSize(Y)) * 100 ' ���i�T�C�YY
        .lngPitch(X) = Val(txtPitch(X)) * 100 ' �ʕt���s�b�`X
        .lngPitch(Y) = Val(txtPitch(Y)) * 100 ' �ʕt���s�b�`Y
        .intNumber(X) = Val(txtNumber(X)) ' �ʕt����X
        .intNumber(Y) = Val(txtNumber(Y)) ' �ʕt����Y
        .lngTotalSize(X) = Val(txtTotal(X)) * 100 ' ���i�S��X
        .lngTotalSize(Y) = Val(txtTotal(Y)) * 100 ' ���i�S��Y
        .strMoji = txtMoji.Text ' �ԕ���

        ' �X�^�b�N�ʒu
        If cmbStack.Text = "��������/����" Then
            .lngStack = .lngWBS(Y) / 2
            .strStart = "Stack"
        ElseIf .intSosu > 2 Then ' ���w��
            .lngStack = Val(cmbStack.Text) * 100
            .strStart = "Stack"
        Else ' ���ʔ�
            .lngStack = Val(cmbStack.Text) * 100
            .strStart = "Machine"
        End If

        .intMojiTool = Val(txtMojiTCode) ' �ԕ�����T�R�[�h
        .intSGAG = Val(txtSGAG.Text) ' SGAG��T�R�[�h
        .lngTestHole = Val(txtTestHole.Text) * 100 ' �������̈ړ���
        .lngIdou(X) = Val(txtIdou(X).Text) * 100 ' ���ʔ̈ړ���X
        .lngIdou(Y) = Val(txtIdou(Y).Text) * 100 ' ���ʔ̈ړ���Y
        .strOutFile = txtFileName.Text ' �o�̓t�@�C����

        ' �g���{�̏o�͂̐ݒ�
        If mnuFormatTombo.Checked = True Then
            .blnTombo = True
        Else
            .blnTombo = False
        End If

        ' �ԕ����̏o�͂̐ݒ�
        If mnuFormatMoji.Checked = True Then
            .blnMoji = True
        Else
            .blnMoji = False
        End If

        ' MBE�N�[�|���̏o�͂̐ݒ�
        If mnuFormatMBE.Checked = True Then
            .blnMBE = True
        Else
            .blnMBE = False
        End If

        ' SG�̏o�͂̐ݒ�
        If mnuFormatSG.Checked = True Then
            .blnSG = True
        Else
            .blnSG = False
        End If

        ' AG�̏o�͂̐ݒ�
        If mnuFormatAG.Checked = True Then
            .blnAG = True
        Else
            .blnAG = False
        End If

        ' �������̏o�͂̐ݒ�
        If mnuFormatTestHole.Checked = True Then
            .blnTestHole = True
        Else
            .blnTestHole = False
        End If

        ' �t�Z�b�g�h�~�f�[�^(T50)�̏o�͂̐ݒ�
        If mnuFormatT50.Checked = True Then
            .blnT50 = True
        Else
            .blnT50 = False
        End If
    End With

End Sub

'*********************************************************
' �p  �r: NC�̎d�l�ɉ����ă��j���[�̏�Ԃ�ύX����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub SetMnuFormat()

    Dim intSosu As Integer

    ' �ϐ��̐ݒ�
    intSosu = Val(txtSosu)

    ' ���j���[�̏�����
    mnuFormatMBE.Checked = False
    mnuFormatTombo.Checked = False
    mnuFormatMoji.Checked = False
    mnuFormatSG.Checked = False
    mnuFormatAG.Checked = False
    mnuFormatTestHole.Checked = False
    mnuFormatT50.Checked = False

    If optTHNT(TH).Value = True Then
        If intSosu > 2 Then
            If InStr(1, txtMoji.Text, "AMS", 1) = 1 Or _
               InStr(1, txtFileName.Text, "AMS", 1) = 1 Then
                mnuFormatT50.Checked = True
                mnuFormatMoji.Checked = True
                mnuFormatTestHole.Checked = True
            Else
                mnuFormatMBE.Checked = True
                mnuFormatTombo.Checked = True
                mnuFormatMoji.Checked = True
                mnuFormatSG.Checked = True
                mnuFormatAG.Checked = True
                mnuFormatTestHole.Checked = True
            End If
        ElseIf intSosu <> 0 Then
            If InStr(1, txtMoji.Text, "AMS", 1) = 1 Or _
               InStr(1, txtFileName.Text, "AMS", 1) = 1 Then
                mnuFormatMoji.Checked = True
                mnuFormatTestHole.Checked = True
            Else
                mnuFormatTombo.Checked = True
                mnuFormatMoji.Checked = True
                mnuFormatSG.Checked = True
                mnuFormatAG.Checked = True
                mnuFormatTestHole.Checked = True
            End If
        End If
    Else
        mnuFormatTestHole.Checked = True
    End If

End Sub

'*********************************************************
' �p  �r: NC�̎d�l�ɉ����ăX�^�b�N�ʒu���Z�b�g����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sStack()

    Dim intSosu As Integer ' �w��
    Dim lngWBSY As Long ' Y�����[�N�{�[�h�T�C�Y

    Set txtWBS(Y) = txtInputY(0)

    ' �ϐ��̏����ݒ�
    intSosu = Val(txtSosu)
    lngWBSY = Val(txtWBS(Y))

    With cmbStack
        If intSosu <= 2 Then
            If lngWBSY > 500 Then
                .Text = "��������/����"
            ElseIf lngWBSY >= 400 Then
                .Text = "205"
            ElseIf lngWBSY <> 0 Then
                .Text = "180"
            End If

'            If InStr(1, txtMoji.Text, "AMS", vbTextCompare) > 0 Then
'                .Text = "180" ' AMS�i��180�X�^�b�N
'            ElseIf InStr(1, txtFileName.Text, "AMS", vbTextCompare) > 0 Then
'                .Text = "180" ' AMS�i��180�X�^�b�N
'            ElseIf lngWBSY > 500 Then
'                .Text = "��������/����"
'            ElseIf lngWBSY >= 400 Then
'                .Text = "205"
'            ElseIf lngWBSY <> 0 Then
'                .Text = "180"
'            End If
        Else
            .Text = "��������/����"
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���s�{�^����Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdExec_Click()

    Dim strEdit_Arg As String

    ' ���s�{�^���������Ȃ��悤�ɂ���
    cmdExec.Enabled = False
    Me.Enabled = False

    ' �ϐ��ɃZ�b�g����
    Call sSetENV

    ' Men2k.Cfg�ɃZ�[�u����
'    Call sSaveCfg(mudtCurrentNC)

    ' �ϊ������s����
    ProgressBar1.Visible = True
    Call sMakeSub(mudtCurrentNC)
    Call sMakeMain(mudtCurrentNC)
    ProgressBar1.Visible = False

    With mudtCurrentNC
        ' �������̈ړ��ʂ��ύX���ꂽ��������Ȃ��̂ōĐݒ�
        If .lngTestHole <> 0 Then
            txtTestHole.Text = _
                Format(.lngTestHole / 100, "##0.00")
        End If
        ' �t�@�C������W�J����
'        strEdit_Arg = Replace(strEditor, "$OUTFILE", .strOutFile)
    End With

    Me.Enabled = True
    cmdExec.Enabled = True
    cmdExec.SetFocus ' �t�H�[�J�X��߂�

    ' �G�f�B�b�^���ݒ肳��Ă�����N������
    If mstrEditor <> "" Then
        Shell mstrEditor & " " & mudtCurrentNC.strOutFile, vbNormalFocus
    End If

End Sub

'*********************************************************
' �p  �r: ���ʔ̈ړ��ʂ��v�Z����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sIdou()

    Dim sngIdou(1) As Single ' �ړ���

    Set txtWBS(X) = txtInputX(0) ' X��WBS
    Set txtWBS(Y) = txtInputY(0) ' X��WBS

    If txtTotal(X) <> "" Then
        sngIdou(X) = Round((Val(txtWBS(X)) - Val(txtTotal(X))) / 2 - 4, 1)
        txtIdou(X) = Format(sngIdou(X), "##0.00")
    Else
        txtIdou(X) = "" ' �S�����ݒ肳��Ă��Ȃ����͈ړ��ʂ�ݒ肵�Ȃ�
    End If
    If txtTotal(Y) <> "" Then
        sngIdou(Y) = Round((Val(txtWBS(Y)) - Val(txtTotal(Y))) / 2, 1)
        txtIdou(Y) = Format(sngIdou(Y), "##0.00")
    Else
        txtIdou(Y) = "" ' �S�����ݒ肳��Ă��Ȃ����͈ړ��ʂ�ݒ肵�Ȃ�
    End If

End Sub

'*********************************************************
' �p  �r: ���i�̑S�����v�Z����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sTotalSize()

    Dim sngSize(1) As Single ' ���i�T�C�Y
    Dim sngPitch(1) As Single ' �ʕt���s�b�`
    Dim intNumber(1) As Integer ' �ʕt����
    Dim sngTotal(1) As Single ' �S��

    ' �I�u�W�F�N�g�ϐ��̐ݒ�
    Set txtSize(X) = txtInputX(1)
    Set txtSize(Y) = txtInputY(1)
    Set txtPitch(X) = txtInputX(2)
    Set txtPitch(Y) = txtInputY(2)
    Set txtNumber(X) = txtInputX(3)
    Set txtNumber(Y) = txtInputY(3)

    ' �ϐ��̐ݒ�
    sngSize(X) = Val(txtSize(X))
    sngSize(Y) = Val(txtSize(Y))
    sngPitch(X) = Val(txtPitch(X))
    sngPitch(Y) = Val(txtPitch(Y))
    intNumber(X) = Val(txtNumber(X))
    intNumber(Y) = Val(txtNumber(Y))

    ' X�̑S��
    If intNumber(X) > 0 Then
        sngTotal(X) = sngSize(X) + Abs(sngPitch(X)) * (intNumber(X) - 1)
    Else
        sngTotal(X) = sngSize(X)
    End If
    If sngTotal(X) = 0 Then
        txtTotal(X) = ""
    Else
        txtTotal(X) = Format(sngTotal(X), "##0.00")
    End If

    ' Y�̑S��
    If intNumber(Y) > 0 Then
        sngTotal(Y) = sngSize(Y) + Abs(sngPitch(Y)) * (intNumber(Y) - 1)
    Else
        sngTotal(Y) = sngSize(Y)
    End If
    If sngTotal(Y) = 0 Then
        txtTotal(Y) = ""
    Else
        txtTotal(Y) = Format(sngTotal(Y), "##0.00")
    End If

End Sub

'*********************************************************
' �p  �r: �N���A�{�^����Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdClear_Click()

    Dim i As Integer

    For i = 0 To 3
        txtInputX(i).Text = ""
        txtInputY(i).Text = ""
    Next i
    txtSosu.Text = ""
    With txtMoji
        .Text = ""
        .Enabled = True
        .BackColor = &H80000005
    End With
    lblMoji.Enabled = True
    txtTotal(X).Text = ""
    txtTotal(Y).Text = ""
    cmbStack.Text = ""
    With txtMojiTCode
        .Text = ""
        .Enabled = True
        .BackColor = &H80000005
    End With
    lblMojiTCode.Enabled = True
    With txtSGAG
        .Text = ""
        .Enabled = True
        .BackColor = &H80000005
    End With
    lblSGAG.Enabled = True
    With txtTestHole
        .Text = ""
        .Enabled = True
        .BackColor = &H80000005
    End With
    lblTestHole.Enabled = True
    txtIdou(X).Text = ""
    txtIdou(Y).Text = ""
    txtFileName.Text = ""
    mnuFormatMoji.Checked = False
    mnuFormatTombo.Checked = False
    mnuFormatMBE.Checked = False
    mnuFormatSG.Checked = False
    mnuFormatAG.Checked = False
    mnuFormatTestHole.Checked = False
    mnuFormatT50.Checked = False

    ' �ϐ�������������
    Call ClrEnv(mudtCurrentNC)
    Call ClrEnv(typNC(TH))
    Call ClrEnv(typNC(NT))
    Caption = strTitleText
    optTHNT(TH).Value = True
    cmdExec.Enabled = False

    ' �t�H�[�J�X���I�v�V�����{�^���ɖ߂�
    optTHNT(TH).SetFocus

End Sub

'*********************************************************
' �p  �r: �I���{�^����Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdEnd_Click()

    ' �v���O�����̏I��
    Unload Me
    End

End Sub

'*********************************************************
' �p  �r: DDE�ʐM��LinkExecute�C�x���g
' ��  ��: CmdStr: �f�X�e�B�l�[�V�����A�v���P�[�V�����ɂ����
'                 ���M���ꂽ������
'         Cancel: �����񂪎󂯕t����ꂽ���ǂ�����ʒm�����
'                 �̐����l
' �߂�l: ����
'*********************************************************

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

    Cancel = 0

    Select Case UCase(CmdStr)
        Case "SOSU"
            lblDDE.Caption = Val(txtSosu.Text) ' �w��
        Case "WBSX"
            lblDDE.Caption = Val(txtInputX(0).Text) * 100 ' X��WBS
        Case "WBSY"
            lblDDE.Caption = Val(txtInputY(0).Text) * 100 ' Y��WBS
        Case "STACK"
            With cmbStack
                If .Text = "��������/����" Then
                    lblDDE.Caption = Val(txtWBS(Y).Text) * 100 / 2
                Else
                    lblDDE.Caption = Val(.Text) * 100
                End If
            End With
        Case "START"
            With cmbStack
                If .Text = "��������/����" Or Val(txtSosu.Text) > 2 Then
                    lblDDE.Caption = "Stack"
                Else
                    lblDDE.Caption = "Machine"
                End If
            End With
    End Select

End Sub

'*********************************************************
' �p  �r: frmMain��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    Dim intFileNum As Integer
    Dim strInput As String
    Dim strValue() As String
    Dim strFileName() As String
    Dim strHomeDir As String

'    On Error GoTo Trap
'    With lblDDE
'        .LinkMode = 0
'        .LinkTopic = "NCArray|frmMain"
'        .LinkItem = "lblDDE"
'        .LinkMode = 1
'    End With
'Trap:

    ' �O��I�����̈ʒu�𕜌�
    Top = GetSetting("NCArray", _
                     "Position", _
                     "Top", _
                     "0")
    Left = GetSetting("NCArray", _
                      "Position", _
                      "Left", _
                      "0")

    ProgressBar1.Visible = False
    Set ProgressBar = frmMain.ProgressBar1
    lblDDE.Visible = False

    ' �R���{�{�b�N�X�̐ݒ�
    With cmbStack
        .AddItem "180"
        .AddItem "205"
        .AddItem "��������/����"
'        .AddItem "��������"
    End With

    If Command <> "" Then
        mudtCurrentNC.strInFile = Command
'        With SysInfo1
'            If .OSPlatform = 1 And .OSVersion = 4 And .OSBuild = 950 Then
'                ' for Windows95
'                mudtCurrentNC.strInFile = Command
'            Else
'                ' �����̃t�@�C�����������O�p�X�ɕϊ����ĕϐ��ɃZ�b�g����
'                ' (Win95�ł͎g���Ȃ�, Win2000�ł͈Ӗ����Ȃ�:-p)
'                mudtCurrentNC.strInFile = ChangeLongPath(Command)
'            End If
'        End With
        With mudtCurrentNC
            Caption = strTitleText & " - " & .strInFile
            strFileName = Split(Command, "\", -1)
            ' �t�@�C�������폜����
            strFileName(UBound(strFileName)) = ""
            ' �J�����g�f�B���N�g�����ړ�����
            ChDir (Join(strFileName, "\"))
        End With
    End If

    ' �N�����͎��s�{�^���������Ȃ��悤�ɂ���
    cmdExec.Enabled = False

    ' ���W�X�g����ǂ�
    mstrEditor = GetSetting("NCArray", _
                            "Settings", _
                            "Editor")

    ' �z�[���f�B���N�g����NCArray.def���L��Γǂ�
    strHomeDir = Environ("HOME")
    ' ���[�g�f�B���N�g���[���̔��f
    If Right$(strHomeDir, 1&) <> "\" Then
        strHomeDir = strHomeDir & "\"
    End If
    If Dir(strHomeDir & "NCArray.def") <> "" Then
        intFileNum = FreeFile
        Open strHomeDir & "NCArray.def" For Input As #intFileNum
        Do Until EOF(intFileNum)
            Line Input #intFileNum, strInput
            strValue = Split(strInput, "=", -1)
            Select Case UCase(strValue(0))
                Case "MAKEMAIN_CMD"
'                   strMakeMain = strValue(1)
                Case "EDITOR"
                    mstrEditor = strValue(1)
            End Select
        Loop
        Close #intFileNum
    End If

End Sub

'*********************************************************
' �p  �r: frmMain��UnLoad�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)

    ' �I�����̈ʒu�����W�X�g���ɕۑ�
    SaveSetting "NCArray", _
                "Position", _
                "Top", _
                Me.Top
    SaveSetting "NCArray", _
                "Position", _
                "Left", _
                Me.Left

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuEdit)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuEdit_Click()

    frmEdit.Show vbModal

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFileOpen)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFileOpen_Click()

    ' �t�@�C�����J���_�C�A���O��\������
    GetInputFile

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFileQuit)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFileQuit_Click()

    ' �v���O�����̏I��
    Unload Me
    End

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatAG)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatAG_Click()

    ' AG�̃`�F�b�N���g�O������
    With mnuFormatAG
        .Checked = Not .Checked
        ' TextBox�̏�Ԃ�ύX����
        If .Checked = True Then
            With txtSGAG
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblSGAG.Enabled = True
        ElseIf .Checked = False And mnuFormatSG.Checked = False Then
            With txtSGAG
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblSGAG.Enabled = False
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatMBE)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatMBE_Click()

    ' �O�H�N�[�|���̃`�F�b�N���g�O������
    With mnuFormatMBE
        .Checked = Not .Checked
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatMoji)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatMoji_Click()

    ' �ԕ����̃`�F�b�N���g�O������
    With mnuFormatMoji
        .Checked = Not .Checked
        ' TextBox�̏�Ԃ�ύX����
        If .Checked = True Then
            With txtMoji
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblMoji.Enabled = True
            With txtMojiTCode
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblMojiTCode.Enabled = True
        Else
            With txtMoji
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblMoji.Enabled = False
            If mnuFormatTombo.Checked = False Then
                With txtMojiTCode
                    .Enabled = False
                    .BackColor = &H8000000F
                End With
                lblMojiTCode.Enabled = False
            End If
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatSG)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatSG_Click()

    ' SG�̃`�F�b�N���g�O������
    With mnuFormatSG
        .Checked = Not .Checked
        ' TextBox�̏�Ԃ�ύX����
        If .Checked = True Then
            With txtSGAG
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblSGAG.Enabled = True
        ElseIf .Checked = False And mnuFormatAG.Checked = False Then
            With txtSGAG
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblSGAG.Enabled = False
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatT50)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatT50_Click()

    ' �t�Z�b�g�h�~(T50)�̃`�F�b�N���g�O������
    With mnuFormatT50
        .Checked = Not .Checked
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatTestHole)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatTestHole_Click()

    ' �������̃`�F�b�N���g�O������
    With mnuFormatTestHole
        .Checked = Not .Checked
        ' TextBox�̏�Ԃ�ύX����
        If .Checked = True Then
            With txtTestHole
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblTestHole.Enabled = True
        Else
            With txtTestHole
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblTestHole.Enabled = False
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuFormatTombo)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuFormatTombo_Click()

    ' �g���{�̃`�F�b�N���g�O������
    With mnuFormatTombo
        .Checked = Not .Checked
        ' TextBox�̏�Ԃ�ύX����
        If .Checked = True Then
            With txtMojiTCode
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblMojiTCode.Enabled = True
        ElseIf mnuFormatMoji.Checked = False Then
            With txtMojiTCode
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblMojiTCode.Enabled = False
        End If
    End With

End Sub

'*********************************************************
' �p  �r: ���j���[(mnuHelpAbout)��Click�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub mnuHelpAbout_Click()

    ' �ް�ޮݏ��̕\��
    frmAbout.Show vbModal

End Sub

'*********************************************************
' �p  �r: �I�v�V�����{�^��(optTHNT)��Click�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub optTHNT_Click(Index As Integer)

    Static intDataType As Integer ' TH/NT�����l���v���O�����I���܂ŕۑ�

    If intDataType = Index Then Exit Sub ' �����{�^�����N���b�N���ꂽ��I��

    Call sSetENV ' ���݂̐ݒ��mudtCurrentNC�ɃZ�b�g
    With mudtCurrentNC
        If Index = TH Then
            intDataType = TH
            .intNCType = NT ' �N���b�N�����O��NT�������̂�
            typNC(NT) = mudtCurrentNC ' ���݂̐ݒ��NT�Ƃ��ĕۑ�
            mudtCurrentNC = typNC(TH) ' TH�̐ݒ�����݂̐ݒ�ɂ���
            SetMnuFormat ' ���j���[���Đݒ肷��
            lblMoji.Enabled = True ' �ԕ�����True�ɂ���
            With txtMoji
                .Enabled = True
                .BackColor = &H80000005
            End With
            fraTCode.Enabled = True ' T�R�[�h��True�ɂ���
            lblSGAG.Enabled = True
            With txtSGAG
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblMojiTCode.Enabled = True
            With txtMojiTCode
                .Enabled = True
                .BackColor = &H80000005
            End With
            lblStack.Enabled = True
            With cmbStack
                .Enabled = True
                .BackColor = &H80000005
            End With
        ElseIf Index = NT Then
            intDataType = NT
            .intNCType = TH ' �N���b�N�����O��TH�������̂�
            typNC(TH) = mudtCurrentNC ' ���݂̐ݒ��TH�Ƃ��ĕۑ�
            mudtCurrentNC = typNC(NT) ' TH�̐ݒ�����݂̐ݒ�ɂ���
            SetMnuFormat ' ���j���[���Đݒ肷��
            lblMoji.Enabled = False ' �ԕ�����False�ɂ���
            With txtMoji
                .Enabled = False
                .BackColor = &H8000000F
            End With
            ' T�R�[�h��False�ɂ���
            fraTCode.Enabled = False
            lblSGAG.Enabled = False
            With txtSGAG
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblMojiTCode.Enabled = False
            With txtMojiTCode
                .Enabled = False
                .BackColor = &H8000000F
            End With
            lblStack.Enabled = False
            With cmbStack
                .Enabled = False
                .BackColor = &H8000000F
            End With
            If .lngTestHole = 0 And typNC(TH).lngTestHole > 0 Then
                .lngTestHole = _
                    typNC(TH).intLastTool * 500 + typNC(TH).lngTestHole
            ElseIf .lngTestHole = 0 And typNC(TH).lngTestHole < 0 Then
                .lngTestHole = _
                    typNC(TH).intLastTool * -500 + typNC(TH).lngTestHole
            End If
        End If

        ' �t�@�C�������Z�b�g����Ă�����^�C�g���o�[�ɕ\������
        If .strInFile <> "" Then
            Caption = strTitleText & " - " & .strInFile
        Else
            Caption = strTitleText
        End If

        ' �t�@�C������TextBox�ɃZ�b�g����
        txtFileName.Text = .strOutFile

        ' �������̈ړ��ʂ�\������
        If .lngTestHole <> 0 Then
            txtTestHole.Text = Format(.lngTestHole / 100, "##0.00")
        Else
            txtTestHole.Text = ""
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �c�[���o�[(Toolbar1)��ButtonClick�C�x���g
' ��  ��: Button: Button�I�u�W�F�N�g�ւ̎Q��
' �߂�l: ����
'*********************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    ' �t�@�C�����J���_�C�A���O��\������
    GetInputFile

End Sub

'*********************************************************
' �p  �r: �o�̓t�@�C�����p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtFileName_Change()

    ' ���j���[�̍Đݒ������
    ' Call SetMnuFormat
    If Mid(frmMain.Caption, 8) <> "" And txtFileName <> "" Then
        cmdExec.Enabled = True
    Else
        cmdExec.Enabled = False
    End If

End Sub

'*********************************************************
' �p  �r: �o�̓t�@�C�����p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtFileName_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtFileName
        .SelStart = 0
        .SelLength = Len(txtFileName)
    End With

End Sub

'*********************************************************
' �p  �r: �o�̓t�@�C�����p�e�L�X�g�{�b�N�X��LostFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtFileName_LostFocus()

    ' �t�H�[�J�X���ړ����鎞,�t�@�C������啶���ɕϊ�����
    txtFileName.Text = UCase(txtFileName.Text)

End Sub

'*********************************************************
' �p  �r: �ړ��ʓ��͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtIdou_GotFocus(Index As Integer)

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtIdou(Index)
        .SelStart = 0
        .SelLength = Len(txtIdou(Index))
    End With

End Sub

'*********************************************************
' �p  �r: �ړ��ʓ��͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
'         Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtIdou_Validate(Index As Integer, Cancel As Boolean)

    With txtIdou(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �e��X�̒l���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInputX_GotFocus(Index As Integer)

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtInputX(Index)
        .SelStart = 0
        .SelLength = Len(txtInputX(Index))
    End With

End Sub

'*********************************************************
' �p  �r: �e��X�̒l���͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInputX_Change(Index As Integer)

    Call sTotalSize ' ���i�̑S�����v�Z
    If Val(txtSosu) <= 2 Then
        Call sIdou ' ���ʔ̈ړ��ʂ̌v�Z
    End If

End Sub

'*********************************************************
' �p  �r: �e��X�̒l���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
'         Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtInputX_Validate(Index As Integer, Cancel As Boolean)

    With txtInputX(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �e��Y�̒l���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInputY_GotFocus(Index As Integer)

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtInputY(Index)
        .SelStart = 0
        .SelLength = Len(txtInputY(Index))
    End With

End Sub

'*********************************************************
' �p  �r: �e��Y�̒l���͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInputY_Change(Index As Integer)

    Call sTotalSize ' ���i�̑S�����v�Z
    If Val(txtSosu) <= 2 Then
        Call sIdou ' ���ʔ̈ړ��ʂ̌v�Z
        Call sStack ' �X�^�b�N�ʒu�̐ݒ�
    End If

End Sub

'*********************************************************
' �p  �r: �e��Y�̒l���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
'         Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtInputY_Validate(Index As Integer, Cancel As Boolean)

    With txtInputY(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �ԕ������͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtMoji_Change()

    ' ���j���[�̍Đݒ������
    Call SetMnuFormat

    ' �X�^�b�N�ʒu�̍Đݒ������
    Call sStack

End Sub

'*********************************************************
' �p  �r: �ԕ������͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtMoji_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtMoji
        .SelStart = 0
        .SelLength = Len(txtMoji)
    End With

End Sub

'*********************************************************
' �p  �r: �ԕ������͗p�e�L�X�g�{�b�N�X��LostFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtMoji_LostFocus()

    ' �t�H�[�J�X���ړ����鎞,�ԕ�����啶���ɕϊ�����
    txtMoji.Text = UCase(txtMoji.Text)

End Sub

'*********************************************************
' �p  �r: �ԕ�����T�R�[�h���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtMojiTCode_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtMojiTCode
        .SelStart = 0
        .SelLength = Len(txtMojiTCode)
    End With

End Sub

'*********************************************************
' �p  �r: �ԕ�����T�R�[�h���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtMojiTCode_Validate(Cancel As Boolean)

    With txtMojiTCode
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: SG/AG��T�R�[�h���͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtSGAG_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtSGAG
        .SelStart = 0
        .SelLength = Len(txtSGAG)
    End With

End Sub

'*********************************************************
' �p  �r: SG/AG��T�R�[�h���͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtSGAG_Validate(Cancel As Boolean)

    With txtSGAG
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_Change()

    Dim intSosu As Integer ' �w��

    ' �ϐ��̐ݒ�
    intSosu = Abs(Val(txtSosu.Text))

    ' ���j���[�̍Đݒ������
    SetMnuFormat

    ' �X�^�b�N�ʒu�̐ݒ�
    Call sStack

    ' ���ʔ̎��̐ݒ�
    If intSosu <= 2 Then
        lblIdou.Enabled = True
        With txtIdou(X)
            .Enabled = True
            .BackColor = &H80000005
        End With
        With txtIdou(Y)
            .Enabled = True
            .BackColor = &H80000005
        End With
        txtTestHole = "80.00" ' �������̒l
    Else ' ���w�̎��̐ݒ�
        lblIdou.Enabled = False
        With txtIdou(X)
            .Enabled = False
            .BackColor = &H8000000F
            .Text = ""
        End With
        With txtIdou(Y)
            .Enabled = False
            .BackColor = &H8000000F
            .Text = ""
        End With
        txtTestHole = "-20.00" ' �������̒l
    End If

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtSosu
        .SelStart = 0
        .SelLength = Len(txtSosu)
    End With

End Sub

'*********************************************************
' �p  �r: �w�����͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtSosu_Validate(Cancel As Boolean)

    ' �����łȂ��ꍇ�ē��͂�����
    If Not IsNumeric(txtSosu.Text) Then
        Cancel = True
        MsgBox "��������͂��ĉ������B", vbCritical, "�w��"
    End If

End Sub

'*********************************************************
' �p  �r: ���������͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtTestHole_GotFocus()

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtTestHole
        .SelStart = 0
        .SelLength = Len(txtTestHole)
    End With

End Sub

'*********************************************************
' �p  �r: ���������͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtTestHole_Validate(Cancel As Boolean)

    With txtTestHole
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �S�����͗p�e�L�X�g�{�b�N�X��GotFocus�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub txtTotal_GotFocus(Index As Integer)

    ' �e�L�X�g�{�b�N�X��I����Ԃɂ���
    With txtTotal(Index)
        .SelStart = 0
        .SelLength = Len(txtTotal(Index))
    End With

End Sub

'*********************************************************
' �p  �r: �S�����͗p�e�L�X�g�{�b�N�X��Validate�C�x���g
' ��  ��: Cancel: �R���g���[�����t�H�[�J�X���������ǂ�����
'                 ���肷��l
' �߂�l: ����
'*********************************************************

Private Sub txtTotal_Validate(Index As Integer, Cancel As Boolean)

    With txtTotal(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "��������͂��ĉ�����", , .ToolTipText
        End If
    End With

End Sub
