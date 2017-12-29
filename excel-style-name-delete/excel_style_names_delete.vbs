'==================================================================================
'   スタイル・名前削除スクリプト
'----------------------------------------------------------------------------------
'   Excelブックのスタイル・名前を削除するスクリプト
'   このスクリプトファイルにドラッグ&ドロップでファイルをおいてください
'   削除処理を実行した後上書き保存を行います。
'
'   結構時間がかかります。
'   Excelからマクロで実行したほうが早いので、可能な方はそちらをお勧めします。
'   あくまでこちらは利用のしやすさを重視しています。
'
'   【注意】
'   上書き保存を行うので念のためブックのバックアップを残しておいてください。
'==================================================================================


'==================================================================================
'   ファンクション・プロシージャ定義
'==================================================================================
'=======================================================
'  ファイルパスから小文字の拡張子を取得します。
'=======================================================
Function GetFileExtension(filePath)
	With CreateObject("Scripting.FileSystemObject")
		GetFileExtension = LCase(.GetExtensionName(filePath)) 
    End With
End Function

'==================================================================================
'   メイン処理
'==================================================================================
Sub Main()

	Dim bookPath
	Dim fileExtension
	Dim objExcel
	Dim objDoc

	If WScript.Arguments.Count <> 1 Then
		Call MsgBox ( "ファイルをドラッグ&ドロップでこのスクリプトファイルに配置してください。", vbOKOnly + vbCritical)
		Exit Sub
	End If
	
	bookPath = WScript.Arguments(0)
	fileExtension = GetFileExtension(bookPath)
	If fileExtension <> "xls" And _ 
	   fileExtension <> "xlsx" And _ 
	   fileExtension <> "xlsm" Then
	   
		Call MsgBox ( "対応していないファイルです。", vbOKOnly + vbCritical)
		Exit Sub
	End If
	

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True

	Set objDoc = objExcel.Workbooks.Open(bookPath)
	
	'名前の削除	
	For Each n In objDoc.Names
		If (Instr(n.Name, "!Print_Area") > 0) Then
		ElseIf (Instr(n.Name, "!Print_Titles") > 0) Then
		Else
			n.Delete
		End If
	Next
	
	For Each s In objDoc.Styles
		If Not s.BuiltIn Then
			s.Delete
		End If
	Next
	
	objDoc.Save
	objDoc.Close
	
	objExcel.Quit
	Set objDoc = Nothing
	Set objExcel = Nothing

	Call MsgBox("書式・スタイルを削除しました", vbOKOnly + vbInformation)
End Sub

Call Main()
