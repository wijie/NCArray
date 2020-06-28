Attribute VB_Name = "MainModule"
Option Explicit

'定数の宣言
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const ConfigFile As String = "MakeMain.Cfg"  'Configファイルのファイル名

'構造体の宣言
Public Type DrillData
    intSosu As Integer '層数
    lngWBS(1) As Long 'X,Yのワークボードサイズ
    lngSize(1) As Long 'X,Yの製品サイズ
    lngPitch(1) As Long 'X,Yの面付けピッチ
    intNumber(1) As Integer 'X,Yの面付け数
    lngTotalSize(1) As Long 'X,Yの全長
    lngIdou(1) As Long 'X,Yの移動量
    intLastTool As Integer '最後のTコード
    strInFile As String '入力ファイル名
    strOutFile As String '出力ファイル名
    lngTestHole As Long '試し穴の値
    blnTestHole As Boolean '試し穴を出力するか否かを示すフラグ
    strMoji As String '花文字
    intMojiTool As Integer '花文字のTコード
    blnMoji As Boolean '花文字を出力するか否かを示すフラグ
    blnTombo As Boolean 'トンボ出力するか否かを示すフラグ
    intSGAG As Integer 'SG/AGのTコード
    blnSG As Boolean 'SGを出力するか否かを示すフラグ
    blnAG As Boolean 'AGを出力するか否かを示すフラグ
    lngStack As Long 'スタック
    blnMBE As Boolean '三菱クーポンを出力するか否かを示すフラグ
    blnT50 As Boolean '逆セット防止データ(T50)を出力するか否かを示すフラグ
    strStart As String 'ピン上ススタートかマシン原点スタートかを示す
    blnT00 As Boolean 'T00の有無を示すフラグ
    blnArrayFlag As Boolean
    blnUnArrayFlag As Boolean
    intSubList() As Integer
    intNCType As Integer 'TH又はNT
End Type

'変数の宣言
Public gstrSeparator As String
Public ProgressBar As ProgressBar

'*********************************************************
' 用  途: 構造体を初期化する
' 引  数: ENV: 初期化する構造体
' 戻り値: 無し
'*********************************************************

Public Sub ClrEnv(ByRef ENV As DrillData)

    With ENV
        'メンバーを初期化する
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
' 用  途: スタートアップ
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Sub Main()

    '2重起動をチェック
    If App.PrevInstance Then
        MsgBox "すでに起動されています！"
        End
    End If

    'セパレータの設定
    gstrSeparator = String(40, " ")

    Load frmMain
    frmMain.Show

End Sub

'Public Sub sSaveCfg(udtCurrentNC As DrillData)
'
'    Dim intFileNum As Integer 'ファイルNo.
'
'    On Error GoTo FileWriteError
'
'    With udtCurrentNC
'        'データの書き込み
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
'    MsgBox "書き込みエラーです。", , "艦長、エラーです。"
'
'End Sub

'*********************************************************
' 用  途: 実行ファイルのPathを取得する
' 引  数: 無し
' 戻り値: 実行ファイルのPath
'*********************************************************

Public Function fMyPath() As String

    'プログラム終了まで　MyPath　の内容を保持
    Static MyPath As String
    '途中でディレクトリ-が変更されても起動ディレクトリ-を確保
    If Len(MyPath) = 0& Then
        MyPath = App.Path         'ディレクトリ-を取得
        'ルートディレクトリーかの判断
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath

End Function
