Attribute VB_Name = "MakeSub"
Option Explicit

Public gvarUnSubNC() As Variant ' �ʕt�����Ȃ�NC���i�[����z��
Private mblnSubT00 As Boolean ' �ʕt������NC��T00�����邩�����t���O
Private mblnUnSubT00 As Boolean ' �ʕt�����Ȃ�NC��T00�����邩�����t���O
Private mintSubMax As Integer ' �ʕt������NC�̍ő�T�R�[�h
Private mintUnSubMax As Integer ' �ʕt�����Ȃ�NC�̍ő�T�R�[�h

'*********************************************************
' �p  �r: �T�u�v���O���������쐬����
' ��  ��: udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Public Sub sMakeSub(udtCurrentNC As DrillData)

    Dim varDelStr As Variant ' �폜���镶����
    Dim varStr As Variant ' ��������폜���鎞�̃e���|����
    Dim intFileNum As Integer ' �t�@�C��No.
    Dim strNC0 As String ' NC�t�@�C����ǂݍ���Ŋi�[����ϐ�
    Dim strNC1() As String ' T�R�[�h��Split���Ċi�[����z��
    Dim varSub() As Variant ' �ʕt������NC���i�[����z��
    Dim i As Integer ' ���[�v�J�E���^
    Dim j As Integer '     �V
    Dim k As Integer '     �V
    Dim strEnter As String ' ���s�R�[�h�̎�ނ��i�[����ϐ�
    Dim bytBuf() As Byte

    With udtCurrentNC
        ' �ϐ��̏����ݒ�
        .blnArrayFlag = False
        .blnUnArrayFlag = False
        mblnSubT00 = False
        mblnUnSubT00 = False
        mintSubMax = -32767
        mintUnSubMax = -32767
        ReDim .intSubList(0)
        .intSubList(0) = -32767
        ProgressBar.Max = 7 ' �v���O���X�o�[�̍ő�l(�K��)
        ProgressBar.Min = 0 ' �v���O���X�o�[�̍ŏ��l
        ProgressBar.Value = ProgressBar.Min '�v���O���X�o�[�̏����l

        ' NC��ǂݍ���
        intFileNum = FreeFile
        Open .strInFile For Binary As #intFileNum
        ReDim bytBuf(LOF(intFileNum))
        Get #intFileNum, , bytBuf
        Close #intFileNum
        strNC0 = StrConv(bytBuf, vbUnicode)
        ProgressBar.Value = 1 ' �v���O���X�o�[�̌��ݒl

        ' ���s�R�[�h�𒲂ׂ�
        If InStr(strNC0, vbCrLf) > 0 Then
            strEnter = vbCrLf
        ElseIf InStr(strNC0, vbLf) > 0 Then
            strEnter = vbLf
        ElseIf InStr(strNC0, vbCr) > 0 Then
            strEnter = vbCr
        Else
            'MsgBox "�s���ȃt�@�C���ł�"
            Exit Sub
        End If
        ProgressBar.Value = 2 ' �v���O���X�o�[�̌��ݒl

        ' �폜/�ύX���镶�������������
        varDelStr = Array("G25", "M00", "M02", "M99", "%", " ") ' �폜���镶����
        For Each varStr In varDelStr
            strNC0 = Replace(strNC0, varStr, "", 1, -1, vbTextCompare)
        Next
        strNC0 = Replace(strNC0, "*T", "T*", 1, -1, vbTextCompare)
        While InStr(strNC0, strEnter & strEnter) > 0
            strNC0 = Replace(strNC0, strEnter & strEnter, strEnter)
        Wend
        ProgressBar.Value = 3 ' �v���O���X�o�[�̌��ݒl

        ' T�R�[�h��Split����
        strNC1 = Split(strNC0, "T", -1, vbTextCompare)
        ' �ʕt������f�[�^�Ƃ��Ȃ��f�[�^�ɐU�蕪����
        j = 0
        k = 0
        For i = 1 To UBound(strNC1)
            If InStr(1, strNC1(i), "C") > 0 Then ' �h�����a�w���̕����͖�������
                ' �������Ȃ�
            ElseIf InStr(1, strNC1(i), "*") > 0 Then
                ' �ʕt�����Ȃ��f�[�^
                .blnUnArrayFlag = True
                strNC1(i) = Replace(strNC1(i), "*", "")
                ReDim Preserve gvarUnSubNC(j)
                gvarUnSubNC(j) = Split(strNC1(i), strEnter, -1)
                j = j + 1
            Else
                ' �ʕt������f�[�^
                .blnArrayFlag = True
                ReDim Preserve varSub(k)
                varSub(k) = Split(strNC1(i), strEnter, -1)
                k = k + 1
            End If
        Next
        ProgressBar.Value = 4 ' �v���O���X�o�[�̌��ݒl

        ' �ʕt������f�[�^����������
        If .blnArrayFlag = True Then
            Call sSubMemo(varSub, udtCurrentNC)
        End If
        ProgressBar.Min = 5 ' �v���O���X�o�[�̌��ݒl
        ' �ʕt�����Ȃ��f�[�^����������
        If .blnUnArrayFlag = True Then
            Call sUnSubMemo(gvarUnSubNC, udtCurrentNC)
        End If
        ProgressBar.Value = 6 ' �v���O���X�o�[�̌��ݒl

        ' �o�Ă���T�R�[�h�̍ő�ԍ���ݒ肷��
        If mintUnSubMax > mintSubMax Then
            .intLastTool = mintUnSubMax
        ElseIf mblnSubT00 = False And mblnUnSubT00 = True Then
            ' �ʕt������NC��T00������,�ʕt�����Ȃ�NC��T00���L�鎞��,
            ' �ʕt�����Ȃ�NC�ɂ�,�T�u���������܂߂Ȃ�
            .intLastTool = mintSubMax + 1
        Else
            .intLastTool = mintSubMax
        End If
        ' �ϐ����J������
        strNC0 = ""
        Erase strNC1
        ProgressBar.Value = 7 ' �v���O���X�o�[�̌��ݒl
        Exit Sub
    End With

FileReadError:
    Close #intFileNum
    MsgBox "�ǂݍ��݃G���[�ł��B", , "�͒��A�G���[�ł��B"

End Sub

'*********************************************************
' �p  �r: �T�u�v���O�������� Nxx �` M99 �̕������쐬����
' ��  ��: varSub(): �ʕt������NC�f�[�^
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sSubMemo( _
    ByRef varSub() As Variant, _
    ByRef udtCurrentNC As DrillData)

    Dim intFileNum As Integer ' �t�@�C��No.
    Dim i As Integer ' ���[�v�J�E���^
    Dim j As Long ' ���[�v�J�E���^(�����������ƃI�[�o�[�t���[����̂�Long�^)

    With udtCurrentNC
        ReDim .intSubList(UBound(varSub))
        j = UBound(varSub)
        For i = 0 To j
            ' T�R�[�h��00�̏ꍇ,32767�ɕt���ւ���
            If CInt(varSub(i)(0)) = 0 Then
                varSub(i)(0) = 32767
                mblnSubT00 = True
                .blnT00 = True
            End If
        Next
        ' �ʕt������NC��T�R�[�h�Ń\�[�g����
        Call sToolSort(varSub)
        ' �o�͂���
        intFileNum = FreeFile
        Open .strOutFile For Output As #intFileNum
        Print #intFileNum, ""
        Print #intFileNum, gstrSeparator
        Print #intFileNum, "G26"
        Print #intFileNum, gstrSeparator
        For i = 0 To UBound(varSub)
            If i = 0 Then ' �擪��T�R�[�h��
                If CInt(varSub(i)(0)) = 32767 Then ' T00�̏ꍇ
                    Print #intFileNum, "N51"
                    .intSubList(i) = 1
                Else
                    Print #intFileNum, "N" & CInt(varSub(i)(0)) + 50
                    .intSubList(i) = CInt(varSub(i)(0))
                End If
                Print #intFileNum, gstrSeparator
            ElseIf CInt(varSub(i)(0)) <> CInt(varSub(i - 1)(0)) Then
                If CInt(varSub(i)(0)) = 32767 Then
                    Print #intFileNum, "N" & CInt(varSub(i - 1)(0)) + 51
                    .intSubList(i) = CInt(varSub(i - 1)(0)) + 1
                Else
                    Print #intFileNum, "N" & CInt(varSub(i)(0)) + 50
                    .intSubList(i) = CInt(varSub(i)(0))
                End If
                Print #intFileNum, gstrSeparator
            Else ' ����T�R�[�h���A�����Ă���ꍇ
                .intSubList(i) = CInt(varSub(i)(0))
            End If
            For j = 1 To UBound(varSub(i)) - 1
                Print #intFileNum, varSub(i)(j)
            Next
            Print #intFileNum, gstrSeparator
            If i = UBound(varSub) Then
                Print #intFileNum, "M99"
                Print #intFileNum, gstrSeparator
            ElseIf CInt(varSub(i)(0)) <> CInt(varSub(i + 1)(0)) Then
                Print #intFileNum, "M99"
                Print #intFileNum, gstrSeparator
            End If
        Next
        Print #intFileNum, "G25"
        Print #intFileNum, gstrSeparator
        Print #intFileNum, "%"
        Close #intFileNum
        ' �ʕt������NC�̍ő�T�R�[�h���Z�b�g����
        mintSubMax = .intSubList(UBound(.intSubList))
        ' �z����J������
        Erase varSub
    End With

End Sub

'*********************************************************
' �p  �r: �ʕt�����Ȃ�NC�f�[�^�̏���
' ��  ��: varSub(): �ʕt�����Ȃ�NC�f�[�^
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sUnSubMemo( _
    ByRef varUnSub() As Variant, _
    ByRef udtCurrentNC As DrillData)

    Dim i As Integer, j As Integer ' ���[�v�J�E���^

    With udtCurrentNC
        j = UBound(varUnSub)
        For i = 0 To j
            ' T00��32767�ɕt���ւ���
            If CInt(varUnSub(i)(0)) = 0 Then
                varUnSub(i)(0) = 32767
                mblnUnSubT00 = True
                .blnT00 = True
            End If
        Next
        ' �ʕt�����Ȃ�NC��T�R�[�h�Ń\�[�g����
        Call sToolSort(varUnSub)
        ' �ʕt�����Ȃ�NC�̍ő�T�R�[�h�𒲂ׂ�
        For i = j To 0 Step -1
            If CInt(varUnSub(i)(0)) <> 32767 Then
                mintUnSubMax = CInt(varUnSub(i)(0))
                Exit For
            End If
        Next
        If mblnUnSubT00 = True Then
            mintUnSubMax = mintUnSubMax + 1
        End If
    End With

End Sub

'*********************************************************
' �p  �r: NC�f�[�^��T�R�[�h�̏��������ɕ��בւ���
' ��  ��: varNC(): �\�[�g����NC�f�[�^
' �߂�l: ����
'*********************************************************

Private Sub sToolSort( _
    ByRef varNC() As Variant)

    Dim blnSortFlag As Boolean ' ���בւ��������������ۂ��������t���O
    Dim strTempArray() As String ' �z�����ւ��p�e���|�����z��
    Dim i As Integer, j As Integer ' ���[�v�J�E���^

    ' NC��T�R�[�h�Ń\�[�g����
    j = UBound(varNC) - 1
    Do
        blnSortFlag = False
        For i = 0 To j
            If CInt(varNC(i)(0)) > CInt(varNC(i + 1)(0)) Then
                strTempArray = varNC(i + 1)
                varNC(i + 1) = varNC(i)
                varNC(i) = strTempArray
                blnSortFlag = True
            End If
        Next
        j = j - 1
    Loop While blnSortFlag = True

End Sub
