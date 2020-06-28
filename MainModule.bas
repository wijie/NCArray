Attribute VB_Name = "MainModule"
Option Explicit

'�萔�̐錾
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const ConfigFile As String = "MakeMain.Cfg"  'Config�t�@�C���̃t�@�C����

'�\���̂̐錾
Public Type DrillData
    intSosu As Integer '�w��
    lngWBS(1) As Long 'X,Y�̃��[�N�{�[�h�T�C�Y
    lngSize(1) As Long 'X,Y�̐��i�T�C�Y
    lngPitch(1) As Long 'X,Y�̖ʕt���s�b�`
    intNumber(1) As Integer 'X,Y�̖ʕt����
    lngTotalSize(1) As Long 'X,Y�̑S��
    lngIdou(1) As Long 'X,Y�̈ړ���
    intLastTool As Integer '�Ō��T�R�[�h
    strInFile As String '���̓t�@�C����
    strOutFile As String '�o�̓t�@�C����
    lngTestHole As Long '�������̒l
    blnTestHole As Boolean '���������o�͂��邩�ۂ��������t���O
    strMoji As String '�ԕ���
    intMojiTool As Integer '�ԕ�����T�R�[�h
    blnMoji As Boolean '�ԕ������o�͂��邩�ۂ��������t���O
    blnTombo As Boolean '�g���{�o�͂��邩�ۂ��������t���O
    intSGAG As Integer 'SG/AG��T�R�[�h
    blnSG As Boolean 'SG���o�͂��邩�ۂ��������t���O
    blnAG As Boolean 'AG���o�͂��邩�ۂ��������t���O
    lngStack As Long '�X�^�b�N
    blnMBE As Boolean '�O�H�N�[�|�����o�͂��邩�ۂ��������t���O
    blnT50 As Boolean '�t�Z�b�g�h�~�f�[�^(T50)���o�͂��邩�ۂ��������t���O
    strStart As String '�s����X�X�^�[�g���}�V�����_�X�^�[�g��������
    blnT00 As Boolean 'T00�̗L���������t���O
    blnArrayFlag As Boolean
    blnUnArrayFlag As Boolean
    intSubList() As Integer
    intNCType As Integer 'TH����NT
End Type

'�ϐ��̐錾
Public gstrSeparator As String
Public ProgressBar As ProgressBar

'*********************************************************
' �p  �r: �\���̂�����������
' ��  ��: ENV: ����������\����
' �߂�l: ����
'*********************************************************

Public Sub ClrEnv(ByRef ENV As DrillData)

    With ENV
        '�����o�[������������
        .intSosu = 0
        Erase .lngWBS
        Erase .lngSize
        Erase .lngPitch
        Erase .intNumber
        Erase .lngTotalSize
        Erase .lngIdou
        .intLastTool = 0
        .strInFile = ""
        .strOutFile = ""
        .lngTestHole = 0
        .blnTestHole = True
        .strMoji = ""
        .intMojiTool = 0
        .blnMoji = True
        .blnTombo = True
        .intSGAG = 0
        .blnSG = True
        .blnAG = True
        .lngStack = 0
        .blnMBE = True
        .blnT50 = False
        .strStart = ""
        .blnT00 = False
        .blnArrayFlag = False
        .blnUnArrayFlag = False
        Erase .intSubList
        .intNCType = 0
    End With

End Sub

'*********************************************************
' �p  �r: �X�^�[�g�A�b�v
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Sub Main()

    '2�d�N�����`�F�b�N
    If App.PrevInstance Then
        MsgBox "���łɋN������Ă��܂��I"
        End
    End If

    '�Z�p���[�^�̐ݒ�
    gstrSeparator = String(40, " ")

    Load frmMain
    frmMain.Show

End Sub

'Public Sub sSaveCfg(udtCurrentNC As DrillData)
'
'    Dim intFileNum As Integer '�t�@�C��No.
'
'    On Error GoTo FileWriteError
'
'    With udtCurrentNC
'        '�f�[�^�̏�������
'        intFileNum = FreeFile
'        Open ConfigFile For Output As #intFileNum
'        Print #intFileNum, "Sosu="; CStr(.intSosu)
'        Print #intFileNum, "WBS="; CStr(.lngWBS(X)) & "," & CStr(.lngWBS(Y))
'        Print #intFileNum, "Stack="; CStr(.lngStack)
'        Print #intFileNum, "Start="; CStr(.strStart)
'        Close #intFileNum
'        Exit Sub
'    End With
'
'FileWriteError:
'    Close #intFileNum
'    MsgBox "�������݃G���[�ł��B", , "�͒��A�G���[�ł��B"
'
'End Sub

'*********************************************************
' �p  �r: ���s�t�@�C����Path���擾����
' ��  ��: ����
' �߂�l: ���s�t�@�C����Path
'*********************************************************

Public Function fMyPath() As String

    '�v���O�����I���܂Ł@MyPath�@�̓��e��ێ�
    Static MyPath As String
    '�r���Ńf�B���N�g��-���ύX����Ă��N���f�B���N�g��-���m��
    If Len(MyPath) = 0& Then
        MyPath = App.Path         '�f�B���N�g��-���擾
        '���[�g�f�B���N�g���[���̔��f
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath

End Function
