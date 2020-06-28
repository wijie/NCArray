VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'なし
   Caption         =   "面付け君"
   ClientHeight    =   4110
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5910
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'ｿｰｽ
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows の既定値
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '上揃え
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
            Object.ToolTipText     =   "ファイルを開く"
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
      ToolTipText     =   "花文字"
      Top             =   3360
      Width           =   1572
   End
   Begin VB.TextBox txtSosu 
      Height          =   270
      Left            =   960
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "層数"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtTotal 
      Height          =   270
      Index           =   1
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   19
      ToolTipText     =   "Yの全長"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   0
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   7
      ToolTipText     =   "YのWBS"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   7
      TabIndex        =   6
      ToolTipText     =   "XのWBS"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   7
      TabIndex        =   18
      ToolTipText     =   "Xの全長"
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "終了(&Q)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   39
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ｸﾘｱ(&C)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   38
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "実行(&X)"
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
      ToolTipText     =   "出力ファイル名"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox cmbStack 
      Height          =   300
      ItemData        =   "frmMain.frx":08CA
      Left            =   960
      List            =   "frmMain.frx":08CC
      TabIndex        =   23
      ToolTipText     =   "スタック位置"
      Top             =   3720
      Width           =   1692
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   2
      Left            =   960
      MaxLength       =   7
      TabIndex        =   12
      ToolTipText     =   "Xの面付けピッチ"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtInputX 
      Height          =   270
      Index           =   3
      Left            =   960
      MaxLength       =   4
      TabIndex        =   15
      ToolTipText     =   "Xの面付け数"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   3
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   16
      ToolTipText     =   "Yの面付け数"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   2
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   13
      ToolTipText     =   "Yの面付けピッチ"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtInputY 
      Height          =   270
      Index           =   1
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   10
      ToolTipText     =   "Yの製品寸法"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame fraIdou 
      Caption         =   "移動量(&I)"
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
         ToolTipText     =   "Yの移動量"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtIdou 
         Height          =   270
         Index           =   0
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   33
         ToolTipText     =   "Xの移動量"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTestHole 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   31
         ToolTipText     =   "試し穴の移動量"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblIdou 
         Caption         =   "両面版"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTestHole 
         Caption         =   "試し穴"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraTCode 
      Caption         =   "Tｺｰﾄﾞ(&T)"
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
         ToolTipText     =   "花文字のTコード"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtSGAG 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   28
         ToolTipText     =   "SG/AGのTコード"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblMojiTCode 
         Caption         =   "花文字"
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
      ToolTipText     =   "Xの製品寸法"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame fraDataType 
      Caption         =   "ﾃﾞｰﾀﾀｲﾌﾟ(&D)"
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
      BorderStyle     =   1  '実線
      Caption         =   "DDE用"
      Height          =   255
      Left            =   4800
      TabIndex        =   42
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblWBSX 
      Caption         =   "ＷＢＳ"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblFileName 
      Caption         =   "ﾌｧｲﾙ名(&N):"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblStack 
      Caption         =   "ｽﾀｯｸ"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Caption         =   "全長"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblMoji 
      Caption         =   "花文字"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblMenzuke 
      Caption         =   "面付数"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblPitch 
      Caption         =   "面付ﾋﾟｯﾁ"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblSize 
      Caption         =   "製品寸法"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblSosu 
      Caption         =   "層数"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "開く(&O)"
      End
      Begin VB.Menu mnuFileStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "終了(&Q)"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "書式(&O)"
      Begin VB.Menu mnuFormatTombo 
         Caption         =   "ﾄﾝﾎﾞ"
      End
      Begin VB.Menu mnuFormatMoji 
         Caption         =   "花文字"
      End
      Begin VB.Menu mnuFormatMBE 
         Caption         =   "三菱ｸｰﾎﾟﾝ"
      End
      Begin VB.Menu mnuFormatSG 
         Caption         =   "SG"
      End
      Begin VB.Menu mnuFormatAG 
         Caption         =   "AG"
      End
      Begin VB.Menu mnuFormatTestHole 
         Caption         =   "試し穴"
      End
      Begin VB.Menu mnuFormatT50 
         Caption         =   "逆ｾｯﾄ防止(T50)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "ﾂｰﾙ(&T)"
      Begin VB.Menu mnuEdit 
         Caption         =   "ｴﾃﾞｨﾀ設定"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' タイトルバーに表示するタイトル設定用定数
Const strTitleText As String = "面付け君"
Private mstrEditor As String
Private mudtCurrentNC As DrillData
Private typNC(1) As DrillData
Private txtWBS(1) As TextBox ' ワークボードサイズ
Private txtSize(1) As TextBox ' 製品サイズ
Private txtPitch(1) As TextBox ' 面付けピッチ
Private txtNumber(1) As TextBox ' 面付け数

'*********************************************************
' 用  途: ファイルを開くダイアログを表示する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub GetInputFile()

    ' CancelErrorの設定は真(True)です。
    On Error GoTo ErrHandler

    With CommonDialog1
        ' ファイルの選択方法を設定します。
        .Filter = "すべてのファイル (*.*)|*.*|" & _
                  "NCファイル (*.nc)|*.nc|" & _
                  "データファイル (*.dat)|*.dat"

        ' 既定の選択方法を指定します。
        .FilterIndex = 1

        ' [読み取り専用ファイルとして開く]チェックボックスを表示しない
        ' 既存のファイル名しか入力できないようにする
        .Flags = cdlOFNHideReadOnly Or _
                 cdlOFNFileMustExist

        ' [ファイルを開く] ダイアログ ボックスを表示します。
        .ShowOpen

        If txtFileName.Text <> "" Then
            cmdExec.Enabled = True
        End If

        Caption = strTitleText & " - " & .FileName
    '    ChDir .InitDir
        Exit Sub
    End With

ErrHandler:
    ' ユーザーが[キャンセル] ボタンをクリックしました。
    If Err.Number = cdlCancel Then
        If Mid(Caption, Len(strTitleText) + 4) = "" Then
            Caption = strTitleText
            cmdExec.Enabled = False
        End If
    End If

End Sub

'*********************************************************
' 用  途: 変数に値をセットする
' 引  数: 無し
' 戻り値: 無し
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
        ' 入力ファイル名がセットされていたらファイル名を記憶する
        If Mid(Caption, Len(strTitleText) + 4) <> "" Then
            .strInFile = Mid(Caption, Len(strTitleText) + 4)
        End If
        If optTHNT(TH).Value = True Then
            .intNCType = TH
        Else
            .intNCType = NT
        End If

        .intSosu = Val(txtSosu) ' 層数
        .lngWBS(X) = Val(txtWBS(X)) * 100 ' X側WBS
        .lngWBS(Y) = Val(txtWBS(Y)) * 100 ' Y側WBS
        .lngSize(X) = Val(txtSize(X)) * 100 ' 製品サイズX
        .lngSize(Y) = Val(txtSize(Y)) * 100 ' 製品サイズY
        .lngPitch(X) = Val(txtPitch(X)) * 100 ' 面付けピッチX
        .lngPitch(Y) = Val(txtPitch(Y)) * 100 ' 面付けピッチY
        .intNumber(X) = Val(txtNumber(X)) ' 面付け数X
        .intNumber(Y) = Val(txtNumber(Y)) ' 面付け数Y
        .lngTotalSize(X) = Val(txtTotal(X)) * 100 ' 製品全長X
        .lngTotalSize(Y) = Val(txtTotal(Y)) * 100 ' 製品全長Y
        .strMoji = txtMoji.Text ' 花文字

        ' スタック位置
        If cmbStack.Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
            .lngStack = .lngWBS(Y) / 2
            .strStart = "Stack"
        ElseIf .intSosu > 2 Then ' 多層版
            .lngStack = Val(cmbStack.Text) * 100
            .strStart = "Stack"
        Else ' 両面版
            .lngStack = Val(cmbStack.Text) * 100
            .strStart = "Machine"
        End If

        .intMojiTool = Val(txtMojiTCode) ' 花文字のTコード
        .intSGAG = Val(txtSGAG.Text) ' SGAGのTコード
        .lngTestHole = Val(txtTestHole.Text) * 100 ' 試し穴の移動量
        .lngIdou(X) = Val(txtIdou(X).Text) * 100 ' 両面板の移動量X
        .lngIdou(Y) = Val(txtIdou(Y).Text) * 100 ' 両面板の移動量Y
        .strOutFile = txtFileName.Text ' 出力ファイル名

        ' トンボの出力の設定
        If mnuFormatTombo.Checked = True Then
            .blnTombo = True
        Else
            .blnTombo = False
        End If

        ' 花文字の出力の設定
        If mnuFormatMoji.Checked = True Then
            .blnMoji = True
        Else
            .blnMoji = False
        End If

        ' MBEクーポンの出力の設定
        If mnuFormatMBE.Checked = True Then
            .blnMBE = True
        Else
            .blnMBE = False
        End If

        ' SGの出力の設定
        If mnuFormatSG.Checked = True Then
            .blnSG = True
        Else
            .blnSG = False
        End If

        ' AGの出力の設定
        If mnuFormatAG.Checked = True Then
            .blnAG = True
        Else
            .blnAG = False
        End If

        ' 試し穴の出力の設定
        If mnuFormatTestHole.Checked = True Then
            .blnTestHole = True
        Else
            .blnTestHole = False
        End If

        ' 逆セット防止データ(T50)の出力の設定
        If mnuFormatT50.Checked = True Then
            .blnT50 = True
        Else
            .blnT50 = False
        End If
    End With

End Sub

'*********************************************************
' 用  途: NCの仕様に応じてメニューの状態を変更する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub SetMnuFormat()

    Dim intSosu As Integer

    ' 変数の設定
    intSosu = Val(txtSosu)

    ' メニューの初期化
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
' 用  途: NCの仕様に応じてスタック位置をセットする
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sStack()

    Dim intSosu As Integer ' 層数
    Dim lngWBSY As Long ' Y側ワークボードサイズ

    Set txtWBS(Y) = txtInputY(0)

    ' 変数の初期設定
    intSosu = Val(txtSosu)
    lngWBSY = Val(txtWBS(Y))

    With cmbStack
        If intSosu <= 2 Then
            If lngWBSY > 500 Then
                .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
            ElseIf lngWBSY >= 400 Then
                .Text = "205"
            ElseIf lngWBSY <> 0 Then
                .Text = "180"
            End If

'            If InStr(1, txtMoji.Text, "AMS", vbTextCompare) > 0 Then
'                .Text = "180" ' AMS品は180スタック
'            ElseIf InStr(1, txtFileName.Text, "AMS", vbTextCompare) > 0 Then
'                .Text = "180" ' AMS品は180スタック
'            ElseIf lngWBSY > 500 Then
'                .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
'            ElseIf lngWBSY >= 400 Then
'                .Text = "205"
'            ElseIf lngWBSY <> 0 Then
'                .Text = "180"
'            End If
        Else
            .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
        End If
    End With

End Sub

'*********************************************************
' 用  途: 実行ボタンのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdExec_Click()

    Dim strEdit_Arg As String

    ' 実行ボタンを押せないようにする
    cmdExec.Enabled = False
    Me.Enabled = False

    ' 変数にセットする
    Call sSetENV

    ' Men2k.Cfgにセーブする
'    Call sSaveCfg(mudtCurrentNC)

    ' 変換を実行する
    ProgressBar1.Visible = True
    Call sMakeSub(mudtCurrentNC)
    Call sMakeMain(mudtCurrentNC)
    ProgressBar1.Visible = False

    With mudtCurrentNC
        ' 試し穴の移動量が変更されたかもしれないので再設定
        If .lngTestHole <> 0 Then
            txtTestHole.Text = _
                Format(.lngTestHole / 100, "##0.00")
        End If
        ' ファイル名を展開する
'        strEdit_Arg = Replace(strEditor, "$OUTFILE", .strOutFile)
    End With

    Me.Enabled = True
    cmdExec.Enabled = True
    cmdExec.SetFocus ' フォーカスを戻す

    ' エディッタが設定されていたら起動する
    If mstrEditor <> "" Then
        Shell mstrEditor & " " & mudtCurrentNC.strOutFile, vbNormalFocus
    End If

End Sub

'*********************************************************
' 用  途: 両面板の移動量を計算する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sIdou()

    Dim sngIdou(1) As Single ' 移動量

    Set txtWBS(X) = txtInputX(0) ' X側WBS
    Set txtWBS(Y) = txtInputY(0) ' X側WBS

    If txtTotal(X) <> "" Then
        sngIdou(X) = Round((Val(txtWBS(X)) - Val(txtTotal(X))) / 2 - 4, 1)
        txtIdou(X) = Format(sngIdou(X), "##0.00")
    Else
        txtIdou(X) = "" ' 全長が設定されていない時は移動量を設定しない
    End If
    If txtTotal(Y) <> "" Then
        sngIdou(Y) = Round((Val(txtWBS(Y)) - Val(txtTotal(Y))) / 2, 1)
        txtIdou(Y) = Format(sngIdou(Y), "##0.00")
    Else
        txtIdou(Y) = "" ' 全長が設定されていない時は移動量を設定しない
    End If

End Sub

'*********************************************************
' 用  途: 製品の全長を計算する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub sTotalSize()

    Dim sngSize(1) As Single ' 製品サイズ
    Dim sngPitch(1) As Single ' 面付けピッチ
    Dim intNumber(1) As Integer ' 面付け数
    Dim sngTotal(1) As Single ' 全長

    ' オブジェクト変数の設定
    Set txtSize(X) = txtInputX(1)
    Set txtSize(Y) = txtInputY(1)
    Set txtPitch(X) = txtInputX(2)
    Set txtPitch(Y) = txtInputY(2)
    Set txtNumber(X) = txtInputX(3)
    Set txtNumber(Y) = txtInputY(3)

    ' 変数の設定
    sngSize(X) = Val(txtSize(X))
    sngSize(Y) = Val(txtSize(Y))
    sngPitch(X) = Val(txtPitch(X))
    sngPitch(Y) = Val(txtPitch(Y))
    intNumber(X) = Val(txtNumber(X))
    intNumber(Y) = Val(txtNumber(Y))

    ' Xの全長
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

    ' Yの全長
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
' 用  途: クリアボタンのClickイベント
' 引  数: 無し
' 戻り値: 無し
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

    ' 変数を初期化する
    Call ClrEnv(mudtCurrentNC)
    Call ClrEnv(typNC(TH))
    Call ClrEnv(typNC(NT))
    Caption = strTitleText
    optTHNT(TH).Value = True
    cmdExec.Enabled = False

    ' フォーカスをオプションボタンに戻す
    optTHNT(TH).SetFocus

End Sub

'*********************************************************
' 用  途: 終了ボタンのClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub cmdEnd_Click()

    ' プログラムの終了
    Unload Me
    End

End Sub

'*********************************************************
' 用  途: DDE通信のLinkExecuteイベント
' 引  数: CmdStr: デスティネーションアプリケーションによって
'                 送信された文字列
'         Cancel: 文字列が受け付けられたかどうかを通知する為
'                 の整数値
' 戻り値: 無し
'*********************************************************

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

    Cancel = 0

    Select Case UCase(CmdStr)
        Case "SOSU"
            lblDDE.Caption = Val(txtSosu.Text) ' 層数
        Case "WBSX"
            lblDDE.Caption = Val(txtInputX(0).Text) * 100 ' X側WBS
        Case "WBSY"
            lblDDE.Caption = Val(txtInputY(0).Text) * 100 ' Y側WBS
        Case "STACK"
            With cmbStack
                If .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Then
                    lblDDE.Caption = Val(txtWBS(Y).Text) * 100 / 2
                Else
                    lblDDE.Caption = Val(.Text) * 100
                End If
            End With
        Case "START"
            With cmbStack
                If .Text = "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ" Or Val(txtSosu.Text) > 2 Then
                    lblDDE.Caption = "Stack"
                Else
                    lblDDE.Caption = "Machine"
                End If
            End With
    End Select

End Sub

'*********************************************************
' 用  途: frmMainのLoadイベント
' 引  数: 無し
' 戻り値: 無し
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

    ' 前回終了時の位置を復元
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

    ' コンボボックスの設定
    With cmbStack
        .AddItem "180"
        .AddItem "205"
        .AddItem "ｾﾝﾀｰｽﾀｯｸ/ｽﾀｰﾄ"
'        .AddItem "ｾﾝﾀｰｽﾀｯｸ"
    End With

    If Command <> "" Then
        mudtCurrentNC.strInFile = Command
'        With SysInfo1
'            If .OSPlatform = 1 And .OSVersion = 4 And .OSBuild = 950 Then
'                ' for Windows95
'                mudtCurrentNC.strInFile = Command
'            Else
'                ' 引数のファイル名をロングパスに変換して変数にセットする
'                ' (Win95では使えない, Win2000では意味がない:-p)
'                mudtCurrentNC.strInFile = ChangeLongPath(Command)
'            End If
'        End With
        With mudtCurrentNC
            Caption = strTitleText & " - " & .strInFile
            strFileName = Split(Command, "\", -1)
            ' ファイル名を削除する
            strFileName(UBound(strFileName)) = ""
            ' カレントディレクトリを移動する
            ChDir (Join(strFileName, "\"))
        End With
    End If

    ' 起動時は実行ボタンを押せないようにする
    cmdExec.Enabled = False

    ' レジストリを読む
    mstrEditor = GetSetting("NCArray", _
                            "Settings", _
                            "Editor")

    ' ホームディレクトリにNCArray.defが有れば読む
    strHomeDir = Environ("HOME")
    ' ルートディレクトリーかの判断
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
' 用  途: frmMainのUnLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)

    ' 終了時の位置をレジストリに保存
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
' 用  途: メニュー(mnuEdit)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuEdit_Click()

    frmEdit.Show vbModal

End Sub

'*********************************************************
' 用  途: メニュー(mnuFileOpen)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFileOpen_Click()

    ' ファイルを開くダイアログを表示する
    GetInputFile

End Sub

'*********************************************************
' 用  途: メニュー(mnuFileQuit)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFileQuit_Click()

    ' プログラムの終了
    Unload Me
    End

End Sub

'*********************************************************
' 用  途: メニュー(mnuFormatAG)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatAG_Click()

    ' AGのチェックをトグルする
    With mnuFormatAG
        .Checked = Not .Checked
        ' TextBoxの状態を変更する
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
' 用  途: メニュー(mnuFormatMBE)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatMBE_Click()

    ' 三菱クーポンのチェックをトグルする
    With mnuFormatMBE
        .Checked = Not .Checked
    End With

End Sub

'*********************************************************
' 用  途: メニュー(mnuFormatMoji)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatMoji_Click()

    ' 花文字のチェックをトグルする
    With mnuFormatMoji
        .Checked = Not .Checked
        ' TextBoxの状態を変更する
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
' 用  途: メニュー(mnuFormatSG)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatSG_Click()

    ' SGのチェックをトグルする
    With mnuFormatSG
        .Checked = Not .Checked
        ' TextBoxの状態を変更する
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
' 用  途: メニュー(mnuFormatT50)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatT50_Click()

    ' 逆セット防止(T50)のチェックをトグルする
    With mnuFormatT50
        .Checked = Not .Checked
    End With

End Sub

'*********************************************************
' 用  途: メニュー(mnuFormatTestHole)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatTestHole_Click()

    ' 試し穴のチェックをトグルする
    With mnuFormatTestHole
        .Checked = Not .Checked
        ' TextBoxの状態を変更する
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
' 用  途: メニュー(mnuFormatTombo)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuFormatTombo_Click()

    ' トンボのチェックをトグルする
    With mnuFormatTombo
        .Checked = Not .Checked
        ' TextBoxの状態を変更する
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
' 用  途: メニュー(mnuHelpAbout)のClickイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub mnuHelpAbout_Click()

    ' ﾊﾞｰｼﾞｮﾝ情報の表示
    frmAbout.Show vbModal

End Sub

'*********************************************************
' 用  途: オプションボタン(optTHNT)のClickイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub optTHNT_Click(Index As Integer)

    Static intDataType As Integer ' TH/NT示す値をプログラム終了まで保存

    If intDataType = Index Then Exit Sub ' 同じボタンがクリックされたら終了

    Call sSetENV ' 現在の設定をmudtCurrentNCにセット
    With mudtCurrentNC
        If Index = TH Then
            intDataType = TH
            .intNCType = NT ' クリックされる前はNTだったので
            typNC(NT) = mudtCurrentNC ' 現在の設定をNTとして保存
            mudtCurrentNC = typNC(TH) ' THの設定を現在の設定にする
            SetMnuFormat ' メニューを再設定する
            lblMoji.Enabled = True ' 花文字をTrueにする
            With txtMoji
                .Enabled = True
                .BackColor = &H80000005
            End With
            fraTCode.Enabled = True ' TコードをTrueにする
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
            .intNCType = TH ' クリックされる前はTHだったので
            typNC(TH) = mudtCurrentNC ' 現在の設定をTHとして保存
            mudtCurrentNC = typNC(NT) ' THの設定を現在の設定にする
            SetMnuFormat ' メニューを再設定する
            lblMoji.Enabled = False ' 花文字をFalseにする
            With txtMoji
                .Enabled = False
                .BackColor = &H8000000F
            End With
            ' TコードをFalseにする
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

        ' ファイル名がセットされていたらタイトルバーに表示する
        If .strInFile <> "" Then
            Caption = strTitleText & " - " & .strInFile
        Else
            Caption = strTitleText
        End If

        ' ファイル名をTextBoxにセットする
        txtFileName.Text = .strOutFile

        ' 試し穴の移動量を表示する
        If .lngTestHole <> 0 Then
            txtTestHole.Text = Format(.lngTestHole / 100, "##0.00")
        Else
            txtTestHole.Text = ""
        End If
    End With

End Sub

'*********************************************************
' 用  途: ツールバー(Toolbar1)のButtonClickイベント
' 引  数: Button: Buttonオブジェクトへの参照
' 戻り値: 無し
'*********************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    ' ファイルを開くダイアログを表示する
    GetInputFile

End Sub

'*********************************************************
' 用  途: 出力ファイル名用テキストボックスのChangeイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtFileName_Change()

    ' メニューの再設定をする
    ' Call SetMnuFormat
    If Mid(frmMain.Caption, 8) <> "" And txtFileName <> "" Then
        cmdExec.Enabled = True
    Else
        cmdExec.Enabled = False
    End If

End Sub

'*********************************************************
' 用  途: 出力ファイル名用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtFileName_GotFocus()

    ' テキストボックスを選択状態にする
    With txtFileName
        .SelStart = 0
        .SelLength = Len(txtFileName)
    End With

End Sub

'*********************************************************
' 用  途: 出力ファイル名用テキストボックスのLostFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtFileName_LostFocus()

    ' フォーカスを移動する時,ファイル名を大文字に変換する
    txtFileName.Text = UCase(txtFileName.Text)

End Sub

'*********************************************************
' 用  途: 移動量入力用テキストボックスのGotFocusイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtIdou_GotFocus(Index As Integer)

    ' テキストボックスを選択状態にする
    With txtIdou(Index)
        .SelStart = 0
        .SelLength = Len(txtIdou(Index))
    End With

End Sub

'*********************************************************
' 用  途: 移動量入力用テキストボックスのValidateイベント
' 引  数: Index: コントロール配列のIndexプロパティ
'         Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtIdou_Validate(Index As Integer, Cancel As Boolean)

    With txtIdou(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 各種Xの値入力用テキストボックスのGotFocusイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInputX_GotFocus(Index As Integer)

    ' テキストボックスを選択状態にする
    With txtInputX(Index)
        .SelStart = 0
        .SelLength = Len(txtInputX(Index))
    End With

End Sub

'*********************************************************
' 用  途: 各種Xの値入力用テキストボックスのChangeイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInputX_Change(Index As Integer)

    Call sTotalSize ' 製品の全長を計算
    If Val(txtSosu) <= 2 Then
        Call sIdou ' 両面板の移動量の計算
    End If

End Sub

'*********************************************************
' 用  途: 各種Xの値入力用テキストボックスのValidateイベント
' 引  数: Index: コントロール配列のIndexプロパティ
'         Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtInputX_Validate(Index As Integer, Cancel As Boolean)

    With txtInputX(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 各種Yの値入力用テキストボックスのGotFocusイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInputY_GotFocus(Index As Integer)

    ' テキストボックスを選択状態にする
    With txtInputY(Index)
        .SelStart = 0
        .SelLength = Len(txtInputY(Index))
    End With

End Sub

'*********************************************************
' 用  途: 各種Yの値入力用テキストボックスのChangeイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInputY_Change(Index As Integer)

    Call sTotalSize ' 製品の全長を計算
    If Val(txtSosu) <= 2 Then
        Call sIdou ' 両面板の移動量の計算
        Call sStack ' スタック位置の設定
    End If

End Sub

'*********************************************************
' 用  途: 各種Yの値入力用テキストボックスのValidateイベント
' 引  数: Index: コントロール配列のIndexプロパティ
'         Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtInputY_Validate(Index As Integer, Cancel As Boolean)

    With txtInputY(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 花文字入力用テキストボックスのChangeイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtMoji_Change()

    ' メニューの再設定をする
    Call SetMnuFormat

    ' スタック位置の再設定をする
    Call sStack

End Sub

'*********************************************************
' 用  途: 花文字入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtMoji_GotFocus()

    ' テキストボックスを選択状態にする
    With txtMoji
        .SelStart = 0
        .SelLength = Len(txtMoji)
    End With

End Sub

'*********************************************************
' 用  途: 花文字入力用テキストボックスのLostFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtMoji_LostFocus()

    ' フォーカスを移動する時,花文字を大文字に変換する
    txtMoji.Text = UCase(txtMoji.Text)

End Sub

'*********************************************************
' 用  途: 花文字のTコード入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtMojiTCode_GotFocus()

    ' テキストボックスを選択状態にする
    With txtMojiTCode
        .SelStart = 0
        .SelLength = Len(txtMojiTCode)
    End With

End Sub

'*********************************************************
' 用  途: 花文字のTコード入力用テキストボックスのValidateイベント
' 引  数: Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtMojiTCode_Validate(Cancel As Boolean)

    With txtMojiTCode
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: SG/AGのTコード入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtSGAG_GotFocus()

    ' テキストボックスを選択状態にする
    With txtSGAG
        .SelStart = 0
        .SelLength = Len(txtSGAG)
    End With

End Sub

'*********************************************************
' 用  途: SG/AGのTコード入力用テキストボックスのValidateイベント
' 引  数: Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtSGAG_Validate(Cancel As Boolean)

    With txtSGAG
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのChangeイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_Change()

    Dim intSosu As Integer ' 層数

    ' 変数の設定
    intSosu = Abs(Val(txtSosu.Text))

    ' メニューの再設定をする
    SetMnuFormat

    ' スタック位置の設定
    Call sStack

    ' 両面板の時の設定
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
        txtTestHole = "80.00" ' 試し穴の値
    Else ' 多層板の時の設定
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
        txtTestHole = "-20.00" ' 試し穴の値
    End If

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_GotFocus()

    ' テキストボックスを選択状態にする
    With txtSosu
        .SelStart = 0
        .SelLength = Len(txtSosu)
    End With

End Sub

'*********************************************************
' 用  途: 層数入力用テキストボックスのValidateイベント
' 引  数: Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtSosu_Validate(Cancel As Boolean)

    ' 数字でない場合再入力させる
    If Not IsNumeric(txtSosu.Text) Then
        Cancel = True
        MsgBox "数字を入力して下さい。", vbCritical, "層数"
    End If

End Sub

'*********************************************************
' 用  途: 試し穴入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtTestHole_GotFocus()

    ' テキストボックスを選択状態にする
    With txtTestHole
        .SelStart = 0
        .SelLength = Len(txtTestHole)
    End With

End Sub

'*********************************************************
' 用  途: 試し穴入力用テキストボックスのValidateイベント
' 引  数: Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtTestHole_Validate(Cancel As Boolean)

    With txtTestHole
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub

'*********************************************************
' 用  途: 全長入力用テキストボックスのGotFocusイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub txtTotal_GotFocus(Index As Integer)

    ' テキストボックスを選択状態にする
    With txtTotal(Index)
        .SelStart = 0
        .SelLength = Len(txtTotal(Index))
    End With

End Sub

'*********************************************************
' 用  途: 全長入力用テキストボックスのValidateイベント
' 引  数: Cancel: コントロールがフォーカスを失うかどうかを
'                 決定する値
' 戻り値: 無し
'*********************************************************

Private Sub txtTotal_Validate(Index As Integer, Cancel As Boolean)

    With txtTotal(Index)
        If Not IsNumeric(.Text) And .Text <> "" Then
            Cancel = True
            MsgBox "数字を入力して下さい", , .ToolTipText
        End If
    End With

End Sub
