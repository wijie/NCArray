Attribute VB_Name = "ChgLongPath"
Option Explicit

'�w��̃p�X���������O�p�X���ɕϊ�����API
Private Declare Function GetLongPathName Lib "KERNEL32" _
    Alias "GetLongPathNameA" _
    (ByVal lpszShortPath As String, _
     ByVal lpszLongPath As String, _
     ByVal cchBuffer As Long) As Long

'*********************************************************
' �p  �r: �V���[�g�p�X�����烍���O�p�X��������
' ��  ��: strShortPath: �V���[�g�p�X��
' �߂�l: �����O�p�X��
'*********************************************************

'�V���[�g�p�X���������O�p�X���ɕϊ�
Public Function ChangeLongPath(ByVal strShortPath As String) As String

    Dim strLongPath As String   '�����O�t�@�C�������󂯎��o�b�t�@
    Dim lngBuffer As Long       '��,�o�C�g��

    '�Ƃ肠����,�o�b�t�@�̃T�C�Y��260�Ƃ���
    lngBuffer = 260

    'strLongPath�ɂ��炩����Null���i�[
    strLongPath = String$(lngBuffer, vbNullChar)

    '�֐��̎��s(�����O�t�@�C�����ɕϊ�)
    Call GetLongPathName(strShortPath, strLongPath, lngBuffer)

    '�]����Null����菜��
    ChangeLongPath = Left$(strLongPath, InStr(strLongPath, vbNullChar) - 1)

End Function
