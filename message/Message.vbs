Option Explicit
'==================================================================================
'   メッセージ表示スクリプト
'     タスクスケジュールから登録することでリマインダー代わりにできる
'----------------------------------------------------------------------------------
'   第１引数【必須】：表示する内容
'   第２引数【任意】：タイトル
'==================================================================================


Sub Main()

	Dim content
	Dim title

	If WScript.Arguments.Count < 1 Then
		WScript.Echo "メッセージを指定してください。"
		WScript.Quit -1
		Exit Sub
	End If
	
	'本文
	content = WScript.Arguments(0)
	
	'タイトル
	If WScript.Arguments.Count > 1 Then
		title = WScript.Arguments(1)
	Else
		title = ""
	End If
	
	'リマインダー代わりにしたいので vbSystemModal で前面にでるように
	Call MsgBox(content, vbOKOnly + vbInformation + vbSystemModal, title)

End Sub

Call Main()
