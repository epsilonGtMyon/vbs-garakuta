Option Explicit
'==================================================================================
'   SJISのテキストファイルをBOMなしUTF-8に変換する。
'----------------------------------------------------------------------------------
'   第１引数【必須】：ファイルパス
'==================================================================================


'=======================================================
'  定数定義
'=======================================================
Const adTypeBinary = 1
Const adTypeText   = 2

Const adSaveCreateNotExist  = 1
Const adSaveCreateOverWrite = 2


'=======================================================
'  SJISのファイルを読み込みます。
'=======================================================
Function ReadTextAsSjis(filePath)
	Dim fso
	Dim textFile
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set textFile = fso.OpenTextFile(filePath)
	
	ReadTextAsSjis = textFile.ReadAll()

	Call textFile.Close
	Set textFile = Nothing
	Set fso = Nothing

End Function

'=======================================================
'  文字列をBOMなしUTF-8で書き込みます。
'=======================================================
Sub WriteTextAsUtf8WithoutBom(filePath, content)
	Dim adoSt
	Dim textContentAsBytesAfterBom
	Set adoSt = CreateObject("ADODB.Stream")
	
	With adoSt
		.Type = adTypeText
		.Charset = "UTF-8"
		.Open
		
		'いったんテキストとして書き出す
		.WriteText content
		
		'先頭に戻してバイナリにしてからBOMの分をスキップ
		.Position = 0
		.Type = adTypeBinary
		.Position = 3
		
		'BOM以降の内容をバイナリで読み出す
		textContentAsBytesAfterBom = .Read
		
		'いったんリセットするためにストリームを開きなおす
		.Close
		.Open
		
		'退避しておいたBOM以降のバイナリを改めて書き出す。
		.Write textContentAsBytesAfterBom
		
		'保存
		.SaveToFile filePath, adSaveCreateOverWrite
		.Close
	End With

	Set adoSt = Nothing
End Sub


'==================================================================================
'   メイン処理
'==================================================================================
Sub Main()

	Dim srcFilePath
	Dim destFilePath
	If WScript.Arguments.Count = 0 Then
		Call MsgBox ( "第１引数にファイルを指定してください", vbOKOnly + vbCritical)
		Exit Sub
	End If
	
	'コマンドライン引数からパラメータ取得
	srcFilePath = WScript.Arguments(0)
	If WScript.Arguments.Count > 1 Then
		destFilePath = WScript.Arguments(1)
	Else
		destFilePath = srcFilePath
	End If
	
	'SJIS形式でのファイル読み取り
	Dim textContent
	textContent = ReadTextAsSjis(srcFilePath)
	
	'UTF-8で書き込み
	Call WriteTextAsUtf8WithoutBom(destFilePath, textContent)
	

End Sub

Call Main()

