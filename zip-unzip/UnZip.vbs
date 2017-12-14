Option Explicit
'==================================================================================
'   簡易ZIP解凍スクリプト
'----------------------------------------------------------------------------------
'   第１引数【必須】：ソースファイル名(絶対パス)
'   　　　　　　　　：必ずzipファイルである事
'   第２引数【任意】：宛先ファイル名(絶対パス)
'   　　　　　　　　：　省略した場合はzipファイルと同じフォルダ名となる
'==================================================================================

'==================================================================================
'CopyHereのオプション + でつなげる。
Const FOF_NOCONFIRMATION    = &H10    '上書き確認ダイアログを表示しない（[すべて上書き]と同じ）。


'==================================================================================
'ソースパス
Dim SrcPath
'宛先パス
Dim DestPath


Dim objFso
Dim objShell


'==================================================================================

'=======================================================
'  Zipファイルであるか
'=======================================================
Function IsZipFile(Src)

	IsZipFile = (Lcase(objFso.GetExtensionName(Src)) = "zip" )

End Function

'=======================================================
'  宛先パスの解決
'=======================================================
Sub ResolveDestPath()
	
	If WScript.Arguments.Count > 1 Then
		'第2引数があれば適用
		DestPath = WScript.Arguments.Item(1)
		Exit Sub
	End If
	
	Dim SrcParentDir
	Dim SrcBaseName
	
	SrcParentDir = objFso.GetParentFolderName(SrcPath)
	SrcBaseName = objFso.GetBaseName(SrcPath)
	
	DestPath = objFso.BuildPath(SrcParentDir, SrcBaseName)

End Sub

'=======================================================
'  Zip解凍をします。
'=======================================================
Sub UnCompressZip(Src, Dest)
	
	Call objShell.NameSpace(Dest).CopyHere(objShell.NameSpace(Src).Items, FOF_NOCONFIRMATION)

	
End Sub

'=======================================================
'  フォルダを再帰的に作成します
'=======================================================
Sub Mkdirs(ByVal strPath)
	Dim strParent   ' 親フォルダ

	strParent = objFso.GetParentFolderName(strPath)
	If objFso.FolderExists(strParent) = True Then
		If objFso.FolderExists(strPath) <> True Then
			objFso.CreateFolder strPath
		End If
	Else
 		Mkdirs strParent
		objFso.CreateFolder strPath
	End If
End Sub


'==================================================================================
'   メイン処理
'==================================================================================
Sub Main()

	'ソースファイルがない場合はエラー
	If WScript.Arguments.Count = 0 Then
		
		WScript.Echo "ファイルを指定して下さい"
		WScript.Quit -1
		Exit Sub
	End If

	SrcPath = WScript.Arguments.Item(0)

	Set objFso   = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("Shell.Application")
	
	'ソースファイルがない場合はエラー
	If objFso.FileExists(SrcPath) = False Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "ファイルがありません。:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'拡張子がzipであるか
	If IsZipFile(SrcPath) = False Then
		
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "zipファイルではありません。:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'宛先パスの解決
	Call ResolveDestPath()

	'宛先がファイルとしてすでに存在する場合はエラー
	If objFso.FileExists(DestPath) Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "コピー先が既にファイルとして存在しています。:" & DestPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'宛先ファイルのフォルダを作成
	Call Mkdirs(DestPath)
	
	'宛先ファイルを作成してコピー
	Call UnCompressZip(SrcPath, DestPath)
	
	'廃棄
	Set objFso   = Nothing
	Set objShell = Nothing

End Sub

Call Main()
