VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "NCデータ面付システムのﾊﾞｰｼﾞｮﾝ情報"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdAbout 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblAbout3 
      Caption         =   "WATABE Eiji"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmAbout.frx":000C
      Top             =   720
      Width           =   480
   End
   Begin VB.Line linAbout2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   3120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line linAbout1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   240
      X2              =   3120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblAbout2 
      Caption         =   "Copyright (C) 2000"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblAbout1 
      Caption         =   "NCデータ面付システム Ver.0.0.0"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'フォームの表示位置をfrmMainの中央に設定する
'    Me.Left = _
'        frmMain.Left + (frmMain.Width - Me.Width) / 2
'    Me.Top = _
'        frmMain.Top + (frmMain.Height - Me.Height) / 2

    lblAbout1.Caption = _
        "NCデータ面付システム Ver." & _
        App.Major & "." & _
        App.Minor & "." & _
        App.Revision

    With lblAbout2
        .AutoSize = True
        .Caption = "Copyright (C) 1998-2001"
    End With

End Sub

