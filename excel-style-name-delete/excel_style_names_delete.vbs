'==================================================================================
'   �X�^�C���E���O�폜�X�N���v�g
'----------------------------------------------------------------------------------
'   Excel�u�b�N�̃X�^�C���E���O���폜����X�N���v�g
'   ���̃X�N���v�g�t�@�C���Ƀh���b�O&�h���b�v�Ńt�@�C���������Ă�������
'   �폜���������s������㏑���ۑ����s���܂��B
'
'   ���\���Ԃ�������܂��B
'   Excel����}�N���Ŏ��s�����ق��������̂ŁA�\�ȕ��͂�����������߂��܂��B
'   �����܂ł�����͗��p�̂��₷�����d�����Ă��܂��B
'
'   �y���Ӂz
'   �㏑���ۑ����s���̂ŔO�̂��߃u�b�N�̃o�b�N�A�b�v���c���Ă����Ă��������B
'==================================================================================


'==================================================================================
'   �t�@���N�V�����E�v���V�[�W����`
'==================================================================================
'=======================================================
'  �t�@�C���p�X���珬�����̊g���q���擾���܂��B
'=======================================================
Function GetFileExtension(filePath)
	With CreateObject("Scripting.FileSystemObject")
		GetFileExtension = LCase(.GetExtensionName(filePath)) 
    End With
End Function

'==================================================================================
'   ���C������
'==================================================================================
Sub Main()

	Dim bookPath
	Dim fileExtension
	Dim objExcel
	Dim objDoc

	If WScript.Arguments.Count <> 1 Then
		Call MsgBox ( "�t�@�C�����h���b�O&�h���b�v�ł��̃X�N���v�g�t�@�C���ɔz�u���Ă��������B", vbOKOnly + vbCritical)
		Exit Sub
	End If
	
	bookPath = WScript.Arguments(0)
	fileExtension = GetFileExtension(bookPath)
	If fileExtension <> "xls" And _ 
	   fileExtension <> "xlsx" And _ 
	   fileExtension <> "xlsm" Then
	   
		Call MsgBox ( "�Ή����Ă��Ȃ��t�@�C���ł��B", vbOKOnly + vbCritical)
		Exit Sub
	End If
	

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True

	Set objDoc = objExcel.Workbooks.Open(bookPath)
	
	'���O�̍폜	
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

	Call MsgBox("�����E�X�^�C�����폜���܂���", vbOKOnly + vbInformation)
End Sub

Call Main()
