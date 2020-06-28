NCArray.exe
  環境変数homeのディレクトリにNCArray.defを置くとEditorを変更する事が
  出来ます。これは、ツール → エディタの設定より優先します。
  --- NCArray.def ---
  Editor=c:\usr\local\xyzzy\xyzzy.exe
  --- End of File ---

  エディタを自動で起動したくない場合は、エディタの設定を空にして下さい。

  完全にアンインストールするにはレジストリキー、
  HKEY_CURRENT_USER\Software\VB and VBA Program Settings\NCArray
  を削除して下さい。

NCView.exe
  環境変数TEMPを設定して下さい。

  プロットするには、ツール → オプションメニューでNC2HPGL.EXEをフルパス
  で設定して下さい。

  完全にアンインストールするにはレジストリキー、
  HKEY_CURRENT_USER\Software\VB and VBA Program Settings\NCView
  を削除して下さい。
