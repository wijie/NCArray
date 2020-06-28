Attribute VB_Name = "MakeMain"
Option Explicit

Private mblnRmTombo As Boolean ' トンボの一部を削除するか否かのフラグ
Private mlngSG(1) As Long ' SGの距離X
Private mintFileNum As Integer ' 出力ファイルのファイルNo.

'*********************************************************
' 用  途: メインプログラム部を作成する
' 引  数: udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Public Sub sMakeMain(udtCurrentNC As DrillData)

    Dim intT As Integer ' 現在のTコード
    Dim intScore As Integer ' 製品穴/ガイド穴の有無を示すスコア

    ProgressBar.Min = 0 ' プログレスバーの最小値
    ProgressBar.Value = ProgressBar.Min ' プログレスバーの初期値

    With udtCurrentNC
        ' THのTコードの最大値を設定する
        If .intNCType = TH Then
            If .intLastTool < .intMojiTool Then
                .intLastTool = .intMojiTool
            End If
            If .intLastTool < .intSGAG Then
                .intLastTool = .intSGAG
            End If
        End If

        ProgressBar.Max = .intLastTool ' プログレスバーの最大値

        ' トンボが花文字/ザグリと重ならないかチェック
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
            intScore = fArray(intT, udtCurrentNC) ' 戻り値 1 or 0
            Call sMBE(intT, udtCurrentNC)
            intScore = intScore + fTombo(intT, udtCurrentNC) ' 戻り値 2 or 0
            intScore = intScore + fSG(intT, udtCurrentNC) ' 戻り値 4 or 0
            ' fHanaMojiはmlngSGを参照するのでfSGの後に実行する必要がある
            intScore = intScore + fHanaMoji(intT, udtCurrentNC) ' 戻り値 8 or 0
            intScore = intScore + fAG(intT, udtCurrentNC) ' 戻り値 16 or 0
            Call sTestHole(intT, udtCurrentNC)
            ProgressBar.Value = intT ' プログレスバーの現在値
            If intScore = 0 Then
                MsgBox "T" & intT & "のデータが有りません", vbExclamation, "要確認"
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
' 用  途: メインプログラム部の面付け処理
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Function fArray( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim strM As String ' サブメモリ呼び出しMxx
    Dim lngABS(1) As Long ' 座標
    Dim lngPitch(1) As Long ' 面付けピッチ
    Dim blnCurrentToolArray As Boolean ' サブメモリの有無を示すフラグ
    Dim blnCurrentToolUnArray As Boolean ' 面付けしないNCの有無を示すフラグ
    Dim varSub As Variant ' テンポラリ変数
    Dim i As Integer, j As Integer ' ループカウンタ

    With udtCurrentNC
        ' 変数の初期設定
        lngABS(X) = 0
        lngABS(Y) = 0
        lngPitch(X) = .lngPitch(X)
        lngPitch(Y) = .lngPitch(Y)
        If .intNumber(X) = 0 Then
            .intNumber(X) = 1 ' 面付け数が0の時は1に設定する
        End If
        If .intNumber(Y) = 0 Then
            .intNumber(Y) = 1 ' 面付け数が0の時は1に設定する
        End If
        strM = "M" & intCurrentTool + 50
        blnCurrentToolArray = False
        blnCurrentToolUnArray = False

        If .blnArrayFlag = False And .blnUnArrayFlag = False Then
            fArray = 0
            Exit Function
        ElseIf .blnArrayFlag = True Then
            ' サブメモリの有無を調べる
            For Each varSub In .intSubList
                If intCurrentTool Like varSub = True Then
                    blnCurrentToolArray = True
                    Exit For
                End If
            Next
        End If
        If .blnUnArrayFlag = True Then
            ' 面付けしないNCの有無を調べる
            For i = 0 To UBound(gvarUnSubNC)
                If intCurrentTool = CInt(gvarUnSubNC(i)(0)) Or _
                   (CInt(gvarUnSubNC(i)(0)) = 32767 And intCurrentTool = .intLastTool) Then
                    blnCurrentToolUnArray = True
                    Exit For
                End If
            Next
        End If

        If blnCurrentToolArray = True Or blnCurrentToolUnArray = True Then
            ' 両面板の移動量
            If .intSosu < 3 And .lngIdou(X) <> 0 And .lngIdou(Y) <> 0 Then
                Print #mintFileNum, "X" & .lngIdou(X) & "Y" & .lngIdou(Y)
            ElseIf .intSosu > 2 And blnCurrentToolArray = True Then ' 多層板
                Print #mintFileNum, "X0Y0"
            End If
        End If

        If blnCurrentToolArray = True Then ' 面付け
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
            ' 面付けしないNCの読み込み
            Call sUnArray(intCurrentTool, udtCurrentNC)
        End If

        If blnCurrentToolArray = True Or blnCurrentToolUnArray = True Then
            ' 両面板の移動量の戻り
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
' 用  途: 試し穴を作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sTestHole( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim intPitch As Integer ' 試し穴のピッチ

    With udtCurrentNC
        ' 試し穴不要の場合は実行しない
        If .blnTestHole = False Then Exit Sub

        ' 変数の初期設定
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
' 用  途: 花文字を作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Function fHanaMoji( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim intMojiFNum As Integer ' 花文字ファイルのファイルNo
    Dim strMoji As String ' 花文字の文字列
    Dim strMojiFile0 As String ' 花文字ファイルテンポラリ1
    Dim strMojiFile1() As String ' 花文字ファイルテンポラリ2
    Dim strMojiFile2() As Variant ' 花文字ファイル格納用配列
    Dim lngMoji_Idou As Long ' 花文字移動量
    Dim i As Integer, j As Integer ' ループカウンタ
    Dim bytBuf() As Byte

    With udtCurrentNC
        ' 花文字を入れるツールでない場合は実行しない
        If intCurrentTool <> .intMojiTool Or .blnMoji = False Then
            fHanaMoji = 0
            Exit Function
        End If

        ' 花文字の移動量の設定
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

        ' 花文字データファイルを読む
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
' 用  途: SGを作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: SGを出力した場合は1, しない場合は0
'*********************************************************

Private Function fSG( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngWBS(1) As Long ' ワークボードサイズ
    Dim lngStack As Long ' スタック位置
    Dim lngDistY As Long ' トンボ～SG間のY方向の距離

    With udtCurrentNC
        ' 変数の初期設定
        lngWBS(X) = .lngWBS(X)
        If .lngWBS(Y) = 33800 Then ' Y側338mmの時340mmとして処理
            lngWBS(Y) = 34000
            lngStack = 17000
        Else
            lngWBS(Y) = .lngWBS(Y)
            lngStack = .lngStack
        End If
        If .lngWBS(X) > 60400 Then ' X側604mmを超えるものは全て550mm
            mlngSG(X) = 55000
        Else
            mlngSG(X) = Int((lngWBS(X) - 2000) / 5000) * 5000
        End If
        mlngSG(Y) = Int((lngWBS(Y) - 500) / 5000) * 5000 - 5000
        ' トンボとSGの3穴目が重なる場合SGの3穴目から50mm引く
        lngDistY = lngWBS(Y) - (mlngSG(Y) + 500) - (lngWBS(Y) - .lngTotalSize(Y)) / 2 ' トンボ～SG間のY方向の距離
        If .blnTombo = True And _
            (lngWBS(X) - .lngTotalSize(X)) / 2 - 1000 < 210 And _
            lngDistY > -350 And _
            lngDistY < 2300 Then
                mlngSG(Y) = mlngSG(Y) - 5000
        End If

        ' SGを入れるツールでない場合, これ以降は実行しない
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
' 用  途: AGを作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: SGを出力した場合は1, しない場合は0
'*********************************************************

Private Function fAG( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngWBS(1) As Long ' ワークボードサイズ
    Dim lngStack As Long ' スタック位置
    Dim lngX1 As Long ' 1穴目～2穴目間の距離X
    Dim lngX2 As Long ' ワークセンター～3穴目間の距離X

    With udtCurrentNC
        ' AGを入れるツールでない場合は実行しない
        If intCurrentTool <> .intSGAG Or .blnAG = False Then
            fAG = 0
            Exit Function
        End If

        ' 変数の初期設定
        lngWBS(X) = .lngWBS(X)
        If .lngWBS(Y) = 33800 Then ' Y側338mmの時340mmとして処理
            lngWBS(Y) = 34000
            lngStack = 17000
        Else
            lngWBS(Y) = .lngWBS(Y)
            lngStack = .lngStack
        End If
        If lngWBS(X) > 60400 Then ' X側604mmを超えるものは全て550mm
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
' 用  途: トンボを作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: トンボを出力した場合は1, しない場合は0
'*********************************************************

Private Function fTombo( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData) As Integer

    Dim lngPitch As Long ' 座グリ間ピッチ
    Dim lngSpace(1) As Long ' WB～製品間の余白X/Y

    With udtCurrentNC
        ' トンボを入れるツールでない場合は実行しない
        If intCurrentTool <> .intMojiTool Or _
           .blnTombo = False Or _
           (.lngTotalSize(X) = 0 And .lngTotalSize(Y) = 0) Then
            fTombo = 0
            Exit Function
        End If

        ' 変数の初期設定
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
' 用  途: 三菱クーポンを作成する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sMBE( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim i As Integer ' ループカウンタ
    ' プロシージャが終了した時, 変数が開放されては困るのでStatic
    Static lngOffSet() As Long ' クーポンをずらす場合のオフセット量

    With udtCurrentNC
        ' 30層以上, 両面板はクーポンの仕様外なので何もしない
        If .intSosu > 30 Or .intSosu <= 2 Or .blnMBE = False Then Exit Sub

        If intCurrentTool = 1 And .intMojiTool = 1 Then
            ' クーポンをずらす必要があるか調べる
            lngOffSet = fChkMBE(udtCurrentNC)
            ' トンボが最小径の場合はリード線引き出し用穴を分割しない
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
            ' クーポンをずらす必要があるか調べる
            lngOffSet = fChkMBE(udtCurrentNC)
            ' 製品内の最小計穴
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
            ' リード線引き出し用穴
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
' 用  途: 逆セット防止データを作成する
' 引  数: udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sBDD( _
    udtCurrentNC As DrillData)

    Dim lngMark(1) As Long ' 座グリマーク

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
' 用  途: トンボが花文字/非対称と重ならないかチェックする
' 引  数: udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sChkTmb(udtCurrentNC As DrillData)

    Dim intButton As Integer ' MsgBoxの戻り値
    Dim strReturn As String ' InputBoxの戻り値
    Dim DistX As Single ' 花文字～トンボ/座グリマーク間の距離X
    Dim DistY As Single ' 花文字～トンボ/座グリマーク間の距離Y
    Dim Dist As Single ' 花文字～トンボ/座グリマーク間の距離
    Dim lngMark(1) As Long ' 座グリマーク

    With udtCurrentNC
        ' 変数の初期設定
        mblnRmTombo = False ' デフォルトは削除しない
        lngMark(X) = .lngWBS(X) - 1000
        If .lngWBS(Y) < 31000 Then
            lngMark(Y) = 10000
        Else
            lngMark(Y) = 15000
        End If

        If .intSosu <= 2 Then
            If udtCurrentNC.blnMoji = False Then Exit Sub ' 花文字を出力しない場合は何もしない
            DistX = (.lngWBS(X) - .lngTotalSize(X)) / 2
            DistY = .lngTestHole - 1000 - CSng((Len(.strMoji)) - 1) * 1000 - .lngIdou(Y)
            If DistX <= 1400 And DistY <= 1200 Then
                intButton = _
                    MsgBox("花文字をずらしますか？", _
                           vbYesNo + vbQuestion, _
                           "トンボと花文字が重なります")
                Select Case intButton
                    Case vbYes
                        strReturn = _
                            InputBox((1200 - DistY + .lngTestHole) / 100 & "mm 以上に設定してください", _
                                     "移動量を入力して下さい", _
                                     Int(((1200 - DistY + .lngTestHole) / 500) + 0.9) * 5)
                        If strReturn = "" Then
                            ' 何もしない
                        Else
                            .lngTestHole = CSng(strReturn) * 100
                        End If
                    Case vbNo
                        ' 何もしない
                End Select
            End If
        Else
            DistX = (.lngWBS(X) - .lngTotalSize(X)) / 2 - 1000
            ' トンボ外側
            DistY = lngMark(Y) - Abs(.lngStack - (.lngWBS(Y) - .lngTotalSize(Y)) / 2)
            Dist = Round(Sqr(DistX ^ 2 + DistY ^ 2))
            If Dist <= 700 Then
                GoTo Question
            End If
            ' トンボ内側
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
        MsgBox("トンボが非対称から" & Dist / 100 & "mmです。" _
        & Chr(&HD) & Chr(&HA) & "削除しますか？", vbYesNo + vbQuestion, "トンボと非対称が重なります")
    Select Case intButton
        Case vbYes
            ' MsgBox "トンボを削除しました", , "確認"
            mblnRmTombo = True
        Case vbNo
            mblnRmTombo = False
    End Select

End Sub

'*********************************************************
' 用  途: 三菱クーポンをずらすか問い合わせる
' 引  数: udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: クーポンをずらす量X/Yの配列
'*********************************************************

Private Function fChkMBE(udtCurrentNC As DrillData) As Long()

    Dim lngOffSet(1) As Long ' クーポンをずらす場合のオフセット量
    Dim intRet As Integer ' MsgBox関数の戻り値
    Dim strRet As String ' InputBox関数の戻り値
    Dim strArray() As String ' Split関数の戻り値

    ' 初期化
    lngOffSet(X) = 0&
    lngOffSet(Y) = 0&

    With udtCurrentNC
        If .lngWBS(X) - .lngTotalSize(X) < 3000 Then
            intRet = MsgBox("三菱クーポンをずらしますか?", _
                            vbYesNo + vbQuestion, _
                            "余白が15mm以下です")
            Select Case intRet
                Case vbYes
                    strRet = InputBox("X,Yの移動距離入力して下さい", _
                                      "何mmずらしますか?", _
                                      "2.5,0")
                    If strRet = "" Then
                        ' 何もしない
                    Else
                        strArray = Split(strRet, ",", -1, vbTextCompare)
                        If UBound(strArray) = 1 Then
                            lngOffSet(X) = CLng(CSng(strArray(X)) * 100)
                            lngOffSet(Y) = CLng(CSng(strArray(Y)) * 100)
                        End If
                    End If
                Case vbNo
                    ' 何もしない
            End Select
        End If
    End With

    fChkMBE = lngOffSet

End Function

'*********************************************************
' 用  途: 面付けしないNCデータをメインプログラム部に挿入する
' 引  数: intCurrentTool: 現在処理中のTコードの数字の部分
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sUnArray( _
    ByVal intCurrentTool As Integer, _
    udtCurrentNC As DrillData)

    Dim i As Integer ' ループカウンタ
    Dim j As Long ' ループカウンタ(穴数が多いとオーバーフローするのでLong型)

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
