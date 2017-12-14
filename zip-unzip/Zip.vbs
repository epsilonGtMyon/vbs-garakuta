Option Explicit
'==================================================================================
'   簡易ZIP圧縮スクリプト
'----------------------------------------------------------------------------------
'   第１引数【必須】：ソースファイル名(絶対パス)
'   第２引数【任意】：宛先ファイル名(絶対パス)
'   　　　　　　　　：　省略した場合はソースファイル名.zipになる。
'==================================================================================

'==================================================================================
'ソースパス
Dim SrcPath
'宛先パス
Dim DestPath


Dim objFso
Dim objShell


'==================================================================================

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
	
	DestPath = objFso.BuildPath(SrcParentDir, SrcBaseName & ".zip")

End Sub

'=======================================================
'  Zip圧縮をします。
'=======================================================
Sub CompressZip(Src, Dest)

	Dim ItemCount
	ItemCount = 0

	'新しいzipファイルを作成します
	With objFso.CreateTextFile(Dest)
		.Write "PK" & Chr(5) & Chr(6) & String(18,0)
		.Close
	End With
	
	'zipファイルに引数で受け取ったファイルを入れます
	With objShell.NameSpace(Dest)
	
		.CopyHere Src
	
		ItemCount = ItemCount + 1
		Do Until .Items.Count = ItemCount
			'ファイルのコピーが終わるまで待つ必要があるようだ。
			WScript.sleep 1000
		Loop
	
	End With
	
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
	If objFso.FileExists(SrcPath) = False And objFso.FolderExists(SrcPath) = False Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "ファイルがありません。:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'宛先パスの解決
	Call ResolveDestPath()
	
	'宛先ファイルのフォルダを作成
	Call Mkdirs(objFso.GetParentFolderName(DestPath))
	
	'宛先ファイルを作成してコピー
	Call CompressZip(SrcPath, DestPath)
	
	'廃棄
	Set objFso   = Nothing
	Set objShell = Nothing

End Sub

Call Main()
