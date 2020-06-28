Attribute VB_Name = "MakeSub"
Option Explicit

Public gvarUnSubNC() As Variant ' 面付けしないNCを格納する配列
Private mblnSubT00 As Boolean ' 面付けするNCにT00があるか示すフラグ
Private mblnUnSubT00 As Boolean ' 面付けしないNCにT00があるか示すフラグ
Private mintSubMax As Integer ' 面付けするNCの最大Tコード
Private mintUnSubMax As Integer ' 面付けしないNCの最大Tコード

'*********************************************************
' 用  途: サブプログラム部を作成する
' 引  数: udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Public Sub sMakeSub(udtCurrentNC As DrillData)

    Dim varDelStr As Variant ' 削除する文字列
    Dim varStr As Variant ' 文字列を削除する時のテンポラリ
    Dim intFileNum As Integer ' ファイルNo.
    Dim strNC0 As String ' NCファイルを読み込んで格納する変数
    Dim strNC1() As String ' TコードでSplitして格納する配列
    Dim varSub() As Variant ' 面付けするNCを格納する配列
    Dim i As Integer ' ループカウンタ
    Dim j As Integer '     〃
    Dim k As Integer '     〃
    Dim strEnter As String ' 改行コードの種類を格納する変数
    Dim bytBuf() As Byte

    With udtCurrentNC
        ' 変数の初期設定
        .blnArrayFlag = False
        .blnUnArrayFlag = False
        mblnSubT00 = False
        mblnUnSubT00 = False
        mintSubMax = -32767
        mintUnSubMax = -32767
        ReDim .intSubList(0)
        .intSubList(0) = -32767
        ProgressBar.Max = 7 ' プログレスバーの最大値(適当)
        ProgressBar.Min = 0 ' プログレスバーの最小値
        ProgressBar.Value = ProgressBar.Min 'プログレスバーの初期値

        ' NCを読み込む
        intFileNum = FreeFile
        Open .strInFile For Binary As #intFileNum
        ReDim bytBuf(LOF(intFileNum))
        Get #intFileNum, , bytBuf
        Close #intFileNum
        strNC0 = StrConv(bytBuf, vbUnicode)
        ProgressBar.Value = 1 ' プログレスバーの現在値

        ' 改行コードを調べる
        If InStr(strNC0, vbCrLf) > 0 Then
            strEnter = vbCrLf
        ElseIf InStr(strNC0, vbLf) > 0 Then
            strEnter = vbLf
        ElseIf InStr(strNC0, vbCr) > 0 Then
            strEnter = vbCr
        Else
            'MsgBox "不正なファイルです"
            Exit Sub
        End If
        ProgressBar.Value = 2 ' プログレスバーの現在値

        ' 削除/変更する文字列を処理する
        varDelStr = Array("G25", "M00", "M02", "M99", "%", " ") ' 削除する文字列
        For Each varStr In varDelStr
            strNC0 = Replace(strNC0, varStr, "", 1, -1, vbTextCompare)
        Next
        strNC0 = Replace(strNC0, "*T", "T*", 1, -1, vbTextCompare)
        While InStr(strNC0, strEnter & strEnter) > 0
            strNC0 = Replace(strNC0, strEnter & strEnter, strEnter)
        Wend
        ProgressBar.Value = 3 ' プログレスバーの現在値

        ' TコードでSplitする
        strNC1 = Split(strNC0, "T", -1, vbTextCompare)
        ' 面付けするデータとしないデータに振り分ける
        j = 0
        k = 0
        For i = 1 To UBound(strNC1)
            If InStr(1, strNC1(i), "C") > 0 Then ' ドリル径指示の部分は無視する
                ' 何もしない
            ElseIf InStr(1, strNC1(i), "*") > 0 Then
                ' 面付けしないデータ
                .blnUnArrayFlag = True
                strNC1(i) = Replace(strNC1(i), "*", "")
                ReDim Preserve gvarUnSubNC(j)
                gvarUnSubNC(j) = Split(strNC1(i), strEnter, -1)
                j = j + 1
            Else
                ' 面付けするデータ
                .blnArrayFlag = True
                ReDim Preserve varSub(k)
                varSub(k) = Split(strNC1(i), strEnter, -1)
                k = k + 1
            End If
        Next
        ProgressBar.Value = 4 ' プログレスバーの現在値

        ' 面付けするデータを処理する
        If .blnArrayFlag = True Then
            Call sSubMemo(varSub, udtCurrentNC)
        End If
        ProgressBar.Min = 5 ' プログレスバーの現在値
        ' 面付けしないデータを処理する
        If .blnUnArrayFlag = True Then
            Call sUnSubMemo(gvarUnSubNC, udtCurrentNC)
        End If
        ProgressBar.Value = 6 ' プログレスバーの現在値

        ' 出てきたTコードの最大番号を設定する
        If mintUnSubMax > mintSubMax Then
            .intLastTool = mintUnSubMax
        ElseIf mblnSubT00 = False And mblnUnSubT00 = True Then
            ' 面付けするNCにT00が無く,面付けしないNCにT00が有る時は,
            ' 面付けしないNCには,サブメモリを含めない
            .intLastTool = mintSubMax + 1
        Else
            .intLastTool = mintSubMax
        End If
        ' 変数を開放する
        strNC0 = ""
        Erase strNC1
        ProgressBar.Value = 7 ' プログレスバーの現在値
        Exit Sub
    End With

FileReadError:
    Close #intFileNum
    MsgBox "読み込みエラーです。", , "艦長、エラーです。"

End Sub

'*********************************************************
' 用  途: サブプログラム部の Nxx ～ M99 の部分を作成する
' 引  数: varSub(): 面付けするNCデータ
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sSubMemo( _
    ByRef varSub() As Variant, _
    ByRef udtCurrentNC As DrillData)

    Dim intFileNum As Integer ' ファイルNo.
    Dim i As Integer ' ループカウンタ
    Dim j As Long ' ループカウンタ(穴数が多いとオーバーフローするのでLong型)

    With udtCurrentNC
        ReDim .intSubList(UBound(varSub))
        j = UBound(varSub)
        For i = 0 To j
            ' Tコードが00の場合,32767に付け替える
            If CInt(varSub(i)(0)) = 0 Then
                varSub(i)(0) = 32767
                mblnSubT00 = True
                .blnT00 = True
            End If
        Next
        ' 面付けするNCをTコードでソートする
        Call sToolSort(varSub)
        ' 出力する
        intFileNum = FreeFile
        Open .strOutFile For Output As #intFileNum
        Print #intFileNum, ""
        Print #intFileNum, gstrSeparator
        Print #intFileNum, "G26"
        Print #intFileNum, gstrSeparator
        For i = 0 To UBound(varSub)
            If i = 0 Then ' 先頭のTコードが
                If CInt(varSub(i)(0)) = 32767 Then ' T00の場合
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
            Else ' 同じTコードが連続している場合
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
        ' 面付けするNCの最大Tコードをセットする
        mintSubMax = .intSubList(UBound(.intSubList))
        ' 配列を開放する
        Erase varSub
    End With

End Sub

'*********************************************************
' 用  途: 面付けしないNCデータの処理
' 引  数: varSub(): 面付けしないNCデータ
'         udtCurrentNC: 層数, WBS, etc...が収められた構造体
' 戻り値: 無し
'*********************************************************

Private Sub sUnSubMemo( _
    ByRef varUnSub() As Variant, _
    ByRef udtCurrentNC As DrillData)

    Dim i As Integer, j As Integer ' ループカウンタ

    With udtCurrentNC
        j = UBound(varUnSub)
        For i = 0 To j
            ' T00を32767に付け替える
            If CInt(varUnSub(i)(0)) = 0 Then
                varUnSub(i)(0) = 32767
                mblnUnSubT00 = True
                .blnT00 = True
            End If
        Next
        ' 面付けしないNCをTコードでソートする
        Call sToolSort(varUnSub)
        ' 面付けしないNCの最大Tコードを調べる
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
' 用  途: NCデータをTコードの小さい順に並べ替える
' 引  数: varNC(): ソートするNCデータ
' 戻り値: 無し
'*********************************************************

Private Sub sToolSort( _
    ByRef varNC() As Variant)

    Dim blnSortFlag As Boolean ' 並べ替えが発生したか否かを示すフラグ
    Dim strTempArray() As String ' 配列入れ替え用テンポラリ配列
    Dim i As Integer, j As Integer ' ループカウンタ

    ' NCをTコードでソートする
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
