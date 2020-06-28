Attribute VB_Name = "MakeMain"
Option Explicit

Private mblnRmTombo As Boolean ' �g���{�̈ꕔ���폜���邩�ۂ��̃t���O
Private mlngSG(1) As Long ' SG�̋���X
Private mintFileNum As Integer ' �o�̓t�@�C���̃t�@�C��No.

'*********************************************************
' �p  �r: ���C���v���O���������쐬����
' ��  ��: udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Public Sub sMakeMain(udtCurrentNC As DrillData)

    Dim intT As Integer ' ���݂�T�R�[�h
    Dim intScore As Integer ' ���i��/�K�C�h���̗L���������X�R�A

    ProgressBar.Min = 0 ' �v���O���X�o�[�̍ŏ��l
    ProgressBar.Value = ProgressBar.Min ' �v���O���X�o�[�̏����l

    With udtCurrentNC
        ' TH��T�R�[�h�̍ő�l��ݒ肷��
        If .intNCType = TH Then
            If .intLastTool < .intMojiTool Then
                .intLastTool = .intMojiTool
            End If
            If .intLastTool < .intSGAG Then
                .intLastTool = .intSGAG
            End If
        End If

        ProgressBar.Max = .intLastTool ' �v���O���X�o�[�̍ő�l

        ' �g���{���ԕ���/�U�O���Əd�Ȃ�Ȃ����`�F�b�N
        Call sChkTmb(udtCurrentNC)

        mintFileNum = FreeFile
        If .blnArrayFlag = False Then
            Open .strOutFile For Output As #mintFileNum
            Print #mintFileNum, ""
        Else
            Open .strOutFile For Append As #mintFileNum
        End If
        Print #mintFileNum, gstrSeparator

        If .intSosu < 3 And .intNCType = TH Then
            If .strStart = "Machine" And .lngStack - 18000 <> 0 Then
                Print #mintFileNum, "X0Y" & -1 * (.lngStack - 18000)
                Print #mintFileNum, gstrSeparator
            ElseIf .strStart = "Stack" Then
                Print #mintFileNum, "X0Y" & -1 * .lngStack
                Print #mintFileNum, gstrSeparator
            End If
        ElseIf .intSosu > 2 And .intNCType = NT Then
            Print #mintFileNum, "X100Y" & .lngStack
            Print #mintFileNum, gstrSeparator
        End If
        Call sBDD(udtCurrentNC)
        For intT = 1 To .intLastTool
            If intT = .intLastTool And .blnT00 = True Then
                Print #mintFileNum, "T00"
                Print #mintFileNum, gstrSeparator
                Print #mintFileNum, "M00"
            Else
                Print #mintFileNum, "T" & intT
            End If
            Print #mintFileNum, gstrSeparator
            intScore = fArray(intT, udtCurrentNC) ' �߂�l 1 or 0
            Call sMBE(intT, udtCurrentNC)
            intScore = intScore + fTombo(intT, udtCurrentNC) ' �߂�l 2 or 0
            intScore = intScore + fSG(intT, udtCurrentNC) ' �߂�l 4 or 0
            ' fHanaMoji��mlngSG���Q�Ƃ���̂�fSG�̌�Ɏ��s����K�v������
            intScore = intScore + fHanaMoji(intT, udtCurrentNC) ' �߂�l 8 or 0
            intScore = intScore + fAG(intT, udtCurrentNC) ' �߂�l 16 or 0
            Call sTestHole(intT, udtCurrentNC)
            ProgressBar.Value = intT ' �v���O���X�o�[�̌��ݒl
            If intScore = 0 Then
                MsgBox "T" & intT & "�̃f�[�^���L��܂���", vbExclamation, "�v�m�F"
            End If
        Next
        If .intSosu < 3 And .intNCType = TH Then
            If .strStart = "Machine" And .lngStack - 18000 <> 0 Then
                Print #mintFileNum, "X0Y" & .lngStack - 18000
                Print #mintFileNum, gstrSeparator
            ElseIf .strStart = "Stack" Then
                Print #mintFileNum, "X0Y" & .lngStack
                Print #mintFileNum, gstrSeparator
            End If
        ElseIf .intSosu > 2 And .intNCType = NT Then
            Print #mintFileNum, "X-100Y" & -1 * .lngStack
            Print #mintFileNum, gstrSeparator
        End If
        Print #mintFileNum, "M02"
        Print #mintFileNum, gstrSeparator
        Print #mintFileNum, ""
        Close #mintFileNum
    End With

End Sub

'*********************************************************
' �p  �r: ���C���v���O�������̖ʕt������
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Function fArray( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim strM As String ' �T�u�������Ăяo��Mxx
    Dim lngABS(1) As Long ' ���W
    Dim lngPitch(1) As Long ' �ʕt���s�b�`
    Dim blnCurrentToolArray As Boolean ' �T�u�������̗L���������t���O
    Dim blnCurrentToolUnArray As Boolean ' �ʕt�����Ȃ�NC�̗L���������t���O
    Dim varSub As Variant ' �e���|�����ϐ�
    Dim i As Integer, j As Integer ' ���[�v�J�E���^

    With udtCurrentNC
        ' �ϐ��̏����ݒ�
        lngABS(X) = 0
        lngABS(Y) = 0
        lngPitch(X) = .lngPitch(X)
        lngPitch(Y) = .lngPitch(Y)
        If .intNumber(X) = 0 Then
            .intNumber(X) = 1 ' �ʕt������0�̎���1�ɐݒ肷��
        End If
        If .intNumber(Y) = 0 Then
            .intNumber(Y) = 1 ' �ʕt������0�̎���1�ɐݒ肷��
        End If
        strM = "M" & intCurrentTool + 50
        blnCurrentToolArray = False
        blnCurrentToolUnArray = False

        If .blnArrayFlag = False And .blnUnArrayFlag = False Then
            fArray = 0
            Exit Function
        ElseIf .blnArrayFlag = True Then
            ' �T�u�������̗L���𒲂ׂ�
            For Each varSub In .intSubList
                If intCurrentTool Like varSub = True Then
                    blnCurrentToolArray = True
                    Exit For
                End If
            Next
        End If
        If .blnUnArrayFlag = True Then
            ' �ʕt�����Ȃ�NC�̗L���𒲂ׂ�
            For i = 0 To UBound(gvarUnSubNC)
                If intCurrentTool = CInt(gvarUnSubNC(i)(0)) Or _
                   (CInt(gvarUnSubNC(i)(0)) = 32767 And intCurrentTool = .intLastTool) Then
                    blnCurrentToolUnArray = True
                    Exit For
                End If
            Next
        End If

        If blnCurrentToolArray = True Or blnCurrentToolUnArray = True Then
            ' ���ʔ̈ړ���
            If .intSosu < 3 And .lngIdou(X) <> 0 And .lngIdou(Y) <> 0 Then
                Print #mintFileNum, "X" & .lngIdou(X) & "Y" & .lngIdou(Y)
            ElseIf .intSosu > 2 And blnCurrentToolArray = True Then ' ���w��
                Print #mintFileNum, "X0Y0"
            End If
        End If

        If blnCurrentToolArray = True Then ' �ʕt��
            Print #mintFileNum, strM
            If .intNumber(X) >= .intNumber(Y) Then
                For i = 1 To .intNumber(Y)
                    For j = 2 To .intNumber(X)
                        Print #mintFileNum, "X" & lngPitch(X) & "Y0"
                        Print #mintFileNum, strM
                        lngABS(X) = lngABS(X) + lngPitch(X)
                    Next
                    If i < .intNumber(Y) Then
                        Print #mintFileNum, "X0Y" & lngPitch(Y)
                        Print #mintFileNum, strM
                        lngABS(Y) = lngABS(Y) + lngPitch(Y)
                        lngPitch(X) = lngPitch(X) * -1
                    End If
                Next
            Else
                For i = 1 To .intNumber(X)
                    For j = 2 To .intNumber(Y)
                        Print #mintFileNum, "X0Y" & lngPitch(Y)
                        Print #mintFileNum, strM
                        lngABS(Y) = lngABS(Y) + lngPitch(Y)
                    Next
                    If i < .intNumber(X) Then
                        Print #mintFileNum, "X" & lngPitch(X) & "Y0"
                        Print #mintFileNum, strM
                        lngABS(X) = lngABS(X) + lngPitch(X)
                        lngPitch(Y) = lngPitch(Y) * -1
                    End If
                Next
            End If
            If .intSosu > 2 Or lngABS(X) <> 0 Or lngABS(Y) <> 0 Then
                Print #mintFileNum, "X" & lngABS(X) * -1 & "Y" & lngABS(Y) * -1
            End If
            If blnCurrentToolUnArray = True Then
                Print #mintFileNum, gstrSeparator
            End If
        End If

        If blnCurrentToolUnArray = True Then
            ' �ʕt�����Ȃ�NC�̓ǂݍ���
            Call sUnArray(intCurrentTool, udtCurrentNC)
        End If

        If blnCurrentToolArray = True Or blnCurrentToolUnArray = True Then
            ' ���ʔ̈ړ��ʂ̖߂�
            If .intSosu < 3 And .lngIdou(X) <> 0 And .lngIdou(Y) <> 0 Then
                Print #mintFileNum, "X" & -1 * .lngIdou(X) & "Y" & -1 * .lngIdou(Y)
            End If
            Print #mintFileNum, gstrSeparator
        End If
    End With

    If blnCurrentToolArray = False And blnCurrentToolUnArray = False Then
        fArray = 0
    Else
        fArray = 1
    End If

End Function

'*********************************************************
' �p  �r: ���������쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sTestHole( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim intPitch As Integer ' �������̃s�b�`

    With udtCurrentNC
        ' �������s�v�̏ꍇ�͎��s���Ȃ�
        If .blnTestHole = False Then Exit Sub

        ' �ϐ��̏����ݒ�
        If .lngTestHole > 0 Then
            intPitch = 500
        Else
            intPitch = -500
        End If

        Print #mintFileNum, "G81"
        Print #mintFileNum, "X0Y" & .lngTestHole + (intCurrentTool - 1) * intPitch
        Print #mintFileNum, "G80"
        Print #mintFileNum, "X0Y" & (.lngTestHole + (intCurrentTool - 1) * intPitch) * -1
        Print #mintFileNum, gstrSeparator
    End With

End Sub

'*********************************************************
' �p  �r: �ԕ������쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Function fHanaMoji( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim intMojiFNum As Integer ' �ԕ����t�@�C���̃t�@�C��No
    Dim strMoji As String ' �ԕ����̕�����
    Dim strMojiFile0 As String ' �ԕ����t�@�C���e���|����1
    Dim strMojiFile1() As String ' �ԕ����t�@�C���e���|����2
    Dim strMojiFile2() As Variant ' �ԕ����t�@�C���i�[�p�z��
    Dim lngMoji_Idou As Long ' �ԕ����ړ���
    Dim i As Integer, j As Integer ' ���[�v�J�E���^
    Dim bytBuf() As Byte

    With udtCurrentNC
        ' �ԕ���������c�[���łȂ��ꍇ�͎��s���Ȃ�
        If intCurrentTool <> .intMojiTool Or .blnMoji = False Then
            fHanaMoji = 0
            Exit Function
        End If

        ' �ԕ����̈ړ��ʂ̐ݒ�
        If .intSosu > 2 Then
            If .lngTestHole < 0 And mlngSG(Y) + 500 - .lngStack > 7700 Then
                lngMoji_Idou = 7000
            Else
                lngMoji_Idou = (mlngSG(Y) + 500 - .lngStack) + Len(.strMoji) * 1000
                lngMoji_Idou = Int(lngMoji_Idou / 500 + 0.5) * 500
            End If
        Else
            lngMoji_Idou = .lngTestHole - 1000
        End If

        ' �ԕ����f�[�^�t�@�C����ǂ�
        intMojiFNum = FreeFile
        Open fMyPath() & "Moji.dat" For Binary As #intMojiFNum
        ReDim bytBuf(LOF(intMojiFNum))
        Get #intMojiFNum, , bytBuf
        Close #intMojiFNum
        strMojiFile0 = StrConv(bytBuf, vbUnicode)

        strMojiFile1 = Split(strMojiFile0, ";" & vbCrLf, -1, vbTextCompare)
        ReDim strMojiFile2(UBound(strMojiFile1))
        For i = 0 To UBound(strMojiFile1)
            strMojiFile2(i) = Split(strMojiFile1(i), "," & vbCrLf, -1, vbTextCompare)
        Next
        Print #mintFileNum, "X0Y" & lngMoji_Idou
        For i = 1 To Len(.strMoji)
            strMoji = Mid(.strMoji, i, 1)
            For j = 0 To UBound(strMojiFile2)
                If strMoji Like strMojiFile2(j)(0) = True Then
                    Print #mintFileNum, strMojiFile2(j)(1);
                End If
            Next
            If i <> Len(.strMoji) Then
                Print #mintFileNum, "X0Y-1000"
            Else
                Print #mintFileNum, "X0Y" & (Len(.strMoji) - 1) * 1000
            End If
        Next
        Print #mintFileNum, "X0Y" & -1 * lngMoji_Idou
        Print #mintFileNum, gstrSeparator
    End With

    fHanaMoji = 8

End Function

'*********************************************************
' �p  �r: SG���쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: SG���o�͂����ꍇ��1, ���Ȃ��ꍇ��0
'*********************************************************

Private Function fSG( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngWBS(1) As Long ' ���[�N�{�[�h�T�C�Y
    Dim lngStack As Long ' �X�^�b�N�ʒu
    Dim lngDistY As Long ' �g���{�`SG�Ԃ�Y�����̋���

    With udtCurrentNC
        ' �ϐ��̏����ݒ�
        lngWBS(X) = .lngWBS(X)
        If .lngWBS(Y) = 33800 Then ' Y��338mm�̎�340mm�Ƃ��ď���
            lngWBS(Y) = 34000
            lngStack = 17000
        Else
            lngWBS(Y) = .lngWBS(Y)
            lngStack = .lngStack
        End If
        If .lngWBS(X) > 60400 Then ' X��604mm�𒴂�����̂͑S��550mm
            mlngSG(X) = 55000
        Else
            mlngSG(X) = Int((lngWBS(X) - 2000) / 5000) * 5000
        End If
        mlngSG(Y) = Int((lngWBS(Y) - 500) / 5000) * 5000 - 5000
        ' �g���{��SG��3���ڂ��d�Ȃ�ꍇSG��3���ڂ���50mm����
        lngDistY = lngWBS(Y) - (mlngSG(Y) + 500) - (lngWBS(Y) - .lngTotalSize(Y)) / 2 ' �g���{�`SG�Ԃ�Y�����̋���
        If .blnTombo = True And _
            (lngWBS(X) - .lngTotalSize(X)) / 2 - 1000 < 210 And _
            lngDistY > -350 And _
            lngDistY < 2300 Then
                mlngSG(Y) = mlngSG(Y) - 5000
        End If

        ' SG������c�[���łȂ��ꍇ, ����ȍ~�͎��s���Ȃ�
        If intCurrentTool <> .intSGAG Or .blnSG = False Then
            fSG = 0
            Exit Function
        End If

        If .intSosu < 3 Then
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X100Y500"
            Print #mintFileNum, "X" & mlngSG(X) & "Y0"
            Print #mintFileNum, "X-" & mlngSG(X) & "Y" & mlngSG(Y)
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X-100Y-" & (mlngSG(Y) + 500)
            Print #mintFileNum, gstrSeparator
        Else
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X0Y-" & (lngStack - 500)
            Print #mintFileNum, "X" & mlngSG(X) & "Y0"
            Print #mintFileNum, "X-" & mlngSG(X) & "Y" & mlngSG(Y)
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X0Y-" & mlngSG(Y) - (lngStack - 500)
            Print #mintFileNum, gstrSeparator
        End If
    End With

    fSG = 4

End Function

'*********************************************************
' �p  �r: AG���쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: SG���o�͂����ꍇ��1, ���Ȃ��ꍇ��0
'*********************************************************

Private Function fAG( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngWBS(1) As Long ' ���[�N�{�[�h�T�C�Y
    Dim lngStack As Long ' �X�^�b�N�ʒu
    Dim lngX1 As Long ' 1���ځ`2���ڊԂ̋���X
    Dim lngX2 As Long ' ���[�N�Z���^�[�`3���ڊԂ̋���X

    With udtCurrentNC
        ' AG������c�[���łȂ��ꍇ�͎��s���Ȃ�
        If intCurrentTool <> .intSGAG Or .blnAG = False Then
            fAG = 0
            Exit Function
        End If

        ' �ϐ��̏����ݒ�
        lngWBS(X) = .lngWBS(X)
        If .lngWBS(Y) = 33800 Then ' Y��338mm�̎�340mm�Ƃ��ď���
            lngWBS(Y) = 34000
            lngStack = 17000
        Else
            lngWBS(Y) = .lngWBS(Y)
            lngStack = .lngStack
        End If
        If lngWBS(X) > 60400 Then ' X��604mm�𒴂�����̂͑S��550mm
            lngX1 = 55000
        Else
            lngX1 = Int((lngWBS(X) - 2000) / 5000) * 5000
        End If
        If lngWBS(X) = 66000 Then
            lngX2 = 7500
        ElseIf lngWBS(X) = 45700 Then
            lngX2 = 4500
        Else
            lngX2 = Int(lngWBS(X) / 10000) * 1000
        End If

        If .intSosu < 3 Then
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & (lngWBS(X) - lngX1) / 2 - 400 & "Y" & lngWBS(Y) - 500
            Print #mintFileNum, "X" & lngX1 & "Y0"
            Print #mintFileNum, "X-" & lngX1 / 2 - lngX2 & "Y-" & lngWBS(Y) - 1000
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X-" & lngWBS(X) / 2 + lngX2 - 400 & "Y-500"
            Print #mintFileNum, gstrSeparator
        Else
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & (lngWBS(X) - lngX1) / 2 - 500 & "Y" & lngWBS(Y) - lngStack - 500
            Print #mintFileNum, "X" & lngX1 & "Y0"
            Print #mintFileNum, "X-" & lngX1 / 2 - lngX2 & "Y-" & lngWBS(Y) - 1000
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X-" & lngWBS(X) / 2 + lngX2 - 500 & "Y" & lngStack - 500
            Print #mintFileNum, gstrSeparator
        End If
    End With

    fAG = 16

End Function

'*********************************************************
' �p  �r: �g���{���쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: �g���{���o�͂����ꍇ��1, ���Ȃ��ꍇ��0
'*********************************************************

Private Function fTombo( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngPitch As Long ' ���O���ԃs�b�`
    Dim lngSpace(1) As Long ' WB�`���i�Ԃ̗]��X/Y

    With udtCurrentNC
        ' �g���{������c�[���łȂ��ꍇ�͎��s���Ȃ�
        If intCurrentTool <> .intMojiTool Or _
           .blnTombo = False Or _
           (.lngTotalSize(X) = 0 And .lngTotalSize(Y) = 0) Then
            fTombo = 0
            Exit Function
        End If

        ' �ϐ��̏����ݒ�
        lngPitch = .lngWBS(X) - 1000
        lngSpace(X) = Int((lngPitch - .lngTotalSize(X)) / 2)
        lngSpace(Y) = Int((.lngWBS(Y) - .lngTotalSize(Y)) / 2)

        If .intSosu < 3 Then
            Print #mintFileNum, "X" & .lngIdou(X) & "Y" & .lngIdou(Y)
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X-500Y0"
            Print #mintFileNum, "X0Y500"
            Print #mintFileNum, "X0Y" & .lngTotalSize(Y) - 2000
            Print #mintFileNum, "X0Y500"
            Print #mintFileNum, "X" & .lngTotalSize(X) + 1000 & "Y1000"
            Print #mintFileNum, "X0Y-500"
            Print #mintFileNum, "X0Y" & 1000 - .lngTotalSize(Y)
            Print #mintFileNum, "X0Y-500"
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X" & -1 * (.lngTotalSize(X) + 500) & "Y0"
            Print #mintFileNum, "X" & -1 * .lngIdou(X) & "Y" & -1 * .lngIdou(Y)
            Print #mintFileNum, gstrSeparator
        Else
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & lngSpace(X) - 500 & "Y" & lngSpace(Y) - .lngStack
            Print #mintFileNum, "X0Y500"
            Print #mintFileNum, "X0Y" & .lngTotalSize(Y) - 2000
            Print #mintFileNum, "X0Y500"
            Print #mintFileNum, "X" & .lngTotalSize(X) + 1000 & "Y1000"
            Print #mintFileNum, "X0Y-500"
            If mblnRmTombo = True Then
                Print #mintFileNum, "G80"
                Print #mintFileNum, "X" & -1 * (.lngTotalSize(X) + 500 + lngSpace(X)) & _
                                    "Y" & -1 * (.lngWBS(Y) - .lngStack - lngSpace(Y) - 500)
                Print #mintFileNum, gstrSeparator
            Else
                Print #mintFileNum, "X0Y" & -1 * (.lngTotalSize(Y) - 1000)
                Print #mintFileNum, "X0Y-500"
                Print #mintFileNum, "G80"
                Print #mintFileNum, "X" & -1 * (.lngTotalSize(X) + 500 + lngSpace(X)) & _
                                    "Y" & .lngStack - lngSpace(Y)
                Print #mintFileNum, gstrSeparator
            End If
        End If
    End With

    fTombo = 2

End Function

'*********************************************************
' �p  �r: �O�H�N�[�|�����쐬����
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sMBE( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim i As Integer ' ���[�v�J�E���^
    ' �v���V�[�W�����I��������, �ϐ����J������Ă͍���̂�Static
    Static lngOffSet() As Long ' �N�[�|�������炷�ꍇ�̃I�t�Z�b�g��

    With udtCurrentNC
        ' 30�w�ȏ�, ���ʔ̓N�[�|���̎d�l�O�Ȃ̂ŉ������Ȃ�
        If .intSosu > 30 Or .intSosu <= 2 Or .blnMBE = False Then Exit Sub

        If intCurrentTool = 1 And .intMojiTool = 1 Then
            ' �N�[�|�������炷�K�v�����邩���ׂ�
            lngOffSet = fChkMBE(udtCurrentNC)
            ' �g���{���ŏ��a�̏ꍇ�̓��[�h�������o���p���𕪊����Ȃ�
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & .lngWBS(X) - 1500 + lngOffSet(X) & _
                                "Y" & -2500 + lngOffSet(Y)
            Print #mintFileNum, "X0Y-508"
            For i = 1 To .intSosu - 2
                Print #mintFileNum, "X0Y-254"
            Next
            Print #mintFileNum, "X0Y" & (.intSosu * 254) - 8128
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X" & 1500 - lngOffSet(X) - .lngWBS(X) & _
                                "Y" & 10628 - lngOffSet(Y)
            Print #mintFileNum, gstrSeparator
        ElseIf intCurrentTool = 1 Then
            ' �N�[�|�������炷�K�v�����邩���ׂ�
            lngOffSet = fChkMBE(udtCurrentNC)
            ' ���i���̍ŏ��v��
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & .lngWBS(X) - 1500 + lngOffSet(X) & _
                                "Y" & -3008 + lngOffSet(Y)
            For i = 1 To .intSosu - 2
                Print #mintFileNum, "X0Y-254"
            Next
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X" & 1500 - lngOffSet(X) - .lngWBS(X) & _
                                "Y" & 2500 - lngOffSet(Y) + (.intSosu * 254)
            Print #mintFileNum, gstrSeparator
        ElseIf intCurrentTool = .intMojiTool Then
            ' ���[�h�������o���p��
            Print #mintFileNum, "G81"
            Print #mintFileNum, "X" & .lngWBS(X) - 1500 + lngOffSet(X) & _
                                "Y" & -2500 + lngOffSet(Y)
            Print #mintFileNum, "X0Y-8128"
            Print #mintFileNum, "G80"
            Print #mintFileNum, "X" & 1500 - lngOffSet(X) - .lngWBS(X) & _
                                "Y" & 10628 - lngOffSet(Y)
            Print #mintFileNum, gstrSeparator
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �t�Z�b�g�h�~�f�[�^���쐬����
' ��  ��: udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sBDD( _
    udtCurrentNC As DrillData)

    Dim lngMark(1) As Long ' ���O���}�[�N

    With udtCurrentNC
        If .blnT50 = True Then
            lngMark(X) = .lngWBS(X) - 1000
            If .lngWBS(Y) < 31000 Then
                lngMark(Y) = 12000
            Else
                lngMark(Y) = 15000
            End If
            Print #mintFileNum, "T50"
            Print #mintFileNum, gstrSeparator
            Print #mintFileNum, "X" & lngMark(X) & "Y-" & lngMark(Y)
            Print #mintFileNum, "M89"
            Print #mintFileNum, "X-" & lngMark(X) & "Y" & lngMark(Y)
            Print #mintFileNum, gstrSeparator
        End If
    End With

End Sub

'*********************************************************
' �p  �r: �g���{���ԕ���/��Ώ̂Əd�Ȃ�Ȃ����`�F�b�N����
' ��  ��: udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sChkTmb(udtCurrentNC As DrillData)

    Dim intButton As Integer ' MsgBox�̖߂�l
    Dim strReturn As String ' InputBox�̖߂�l
    Dim DistX As Single ' �ԕ����`�g���{/���O���}�[�N�Ԃ̋���X
    Dim DistY As Single ' �ԕ����`�g���{/���O���}�[�N�Ԃ̋���Y
    Dim Dist As Single ' �ԕ����`�g���{/���O���}�[�N�Ԃ̋���
    Dim lngMark(1) As Long ' ���O���}�[�N

    With udtCurrentNC
        ' �ϐ��̏����ݒ�
        mblnRmTombo = False ' �f�t�H���g�͍폜���Ȃ�
        lngMark(X) = .lngWBS(X) - 1000
        If .lngWBS(Y) < 31000 Then
            lngMark(Y) = 10000
        Else
            lngMark(Y) = 15000
        End If

        If .intSosu <= 2 Then
            If udtCurrentNC.blnMoji = False Then Exit Sub ' �ԕ������o�͂��Ȃ��ꍇ�͉������Ȃ�
            DistX = (.lngWBS(X) - .lngTotalSize(X)) / 2
            DistY = .lngTestHole - 1000 - CSng((Len(.strMoji)) - 1) * 1000 - .lngIdou(Y)
            If DistX <= 1400 And DistY <= 1200 Then
                intButton = _
                    MsgBox("�ԕ��������炵�܂����H", _
                           vbYesNo + vbQuestion, _
                           "�g���{�Ɖԕ������d�Ȃ�܂�")
                Select Case intButton
                    Case vbYes
                        strReturn = _
                            InputBox((1200 - DistY + .lngTestHole) / 100 & "mm �ȏ�ɐݒ肵�Ă�������", _
                                     "�ړ��ʂ���͂��ĉ�����", _
                                     Int(((1200 - DistY + .lngTestHole) / 500) + 0.9) * 5)
                        If strReturn = "" Then
                            ' �������Ȃ�
                        Else
                            .lngTestHole = CSng(strReturn) * 100
                        End If
                    Case vbNo
                        ' �������Ȃ�
                End Select
            End If
        Else
            DistX = (.lngWBS(X) - .lngTotalSize(X)) / 2 - 1000
            ' �g���{�O��
            DistY = lngMark(Y) - Abs(.lngStack - (.lngWBS(Y) - .lngTotalSize(Y)) / 2)
            Dist = Round(Sqr(DistX ^ 2 + DistY ^ 2))
            If Dist <= 700 Then
                GoTo Question
            End If
            ' �g���{����
            DistY = lngMark(Y) - Abs(.lngStack - (.lngWBS(Y) - .lngTotalSize(Y)) / 2 - 500)
            Dist = Round(Sqr(DistX ^ 2 + DistY ^ 2))
            If Dist <= 700 Then
                GoTo Question
            End If
        End If
        Exit Sub
    End With

Question:
    intButton = _
        MsgBox("�g���{����Ώ̂���" & Dist / 100 & "mm�ł��B" _
        & Chr(&HD) & Chr(&HA) & "�폜���܂����H", vbYesNo + vbQuestion, "�g���{�Ɣ�Ώ̂��d�Ȃ�܂�")
    Select Case intButton
        Case vbYes
            ' MsgBox "�g���{���폜���܂���", , "�m�F"
            mblnRmTombo = True
        Case vbNo
            mblnRmTombo = False
    End Select

End Sub

'*********************************************************
' �p  �r: �O�H�N�[�|�������炷���₢���킹��
' ��  ��: udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: �N�[�|�������炷��X/Y�̔z��
'*********************************************************

Private Function fChkMBE(udtCurrentNC As DrillData) As Long()

    Dim lngOffSet(1) As Long ' �N�[�|�������炷�ꍇ�̃I�t�Z�b�g��
    Dim intRet As Integer ' MsgBox�֐��̖߂�l
    Dim strRet As String ' InputBox�֐��̖߂�l
    Dim strArray() As String ' Split�֐��̖߂�l

    ' ������
    lngOffSet(X) = 0&
    lngOffSet(Y) = 0&

    With udtCurrentNC
        If .lngWBS(X) - .lngTotalSize(X) < 3000 Then
            intRet = MsgBox("�O�H�N�[�|�������炵�܂���?", _
                            vbYesNo + vbQuestion, _
                            "�]����15mm�ȉ��ł�")
            Select Case intRet
                Case vbYes
                    strRet = InputBox("X,Y�̈ړ��������͂��ĉ�����", _
                                      "��mm���炵�܂���?", _
                                      "2.5,0")
                    If strRet = "" Then
                        ' �������Ȃ�
                    Else
                        strArray = Split(strRet, ",", -1, vbTextCompare)
                        If UBound(strArray) = 1 Then
                            lngOffSet(X) = CLng(CSng(strArray(X)) * 100)
                            lngOffSet(Y) = CLng(CSng(strArray(Y)) * 100)
                        End If
                    End If
                Case vbNo
                    ' �������Ȃ�
            End Select
        End If
    End With

    fChkMBE = lngOffSet

End Function

'*********************************************************
' �p  �r: �ʕt�����Ȃ�NC�f�[�^�����C���v���O�������ɑ}������
' ��  ��: intCurrentTool: ���ݏ�������T�R�[�h�̐����̕���
'         udtCurrentNC: �w��, WBS, etc...�����߂�ꂽ�\����
' �߂�l: ����
'*********************************************************

Private Sub sUnArray( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim i As Integer ' ���[�v�J�E���^
    Dim j As Long ' ���[�v�J�E���^(�����������ƃI�[�o�[�t���[����̂�Long�^)

    With udtCurrentNC
        If .blnUnArrayFlag = False Then Exit Sub

        For i = 0 To UBound(gvarUnSubNC)
            If intCurrentTool = CInt(gvarUnSubNC(i)(0)) Or _
               (CInt(gvarUnSubNC(i)(0)) = 32767 And intCurrentTool = .intLastTool) Then
                For j = 1 To UBound(gvarUnSubNC(i)) - 1
                    Print #mintFileNum, gvarUnSubNC(i)(j)
                Next
                If i < UBound(gvarUnSubNC) Then
                    If CInt(gvarUnSubNC(i)(0)) = CInt(gvarUnSubNC(i + 1)(0)) Then
                        Print #mintFileNum, gstrSeparator
                    End If
                End If
            End If
        Next
    End With

End Sub
