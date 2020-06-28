Attribute VB_Name = "ChgLongPath"
Option Explicit

'指定のパス名をロングパス名に変換するAPI
Private Declare Function GetLongPathName Lib "KERNEL32" _
    Alias "GetLongPathNameA" _
    (ByVal lpszShortPath As String, _
     ByVal lpszLongPath As String, _
     ByVal cchBuffer As Long) As Long

'*********************************************************
' 用  途: ショートパス名からロングパス名を引く
' 引  数: strShortPath: ショートパス名
' 戻り値: ロングパス名
'*********************************************************

'ショートパス名をロングパス名に変換
Public Function ChangeLongPath(ByVal strShortPath As String) As String

    Dim strLongPath As String   'ロングファイル名を受け取るバッファ
    Dim lngBuffer As Long       '同,バイト数

    'とりあえず,バッファのサイズを260とする
    lngBuffer = 260

    'strLongPathにあらかじめNullを格納
    strLongPath = String$(lngBuffer, vbNullChar)

    '関数の実行(ロングファイル名に変換)
    Call GetLongPathName(strShortPath, strLongPath, lngBuffer)

    '余分なNullを取り除く
    ChangeLongPath = Left$(strLongPath, InStr(strLongPath, vbNullChar) - 1)

End Function
