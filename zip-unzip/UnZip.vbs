Option Explicit
'==================================================================================
'   �Ȉ�ZIP�𓀃X�N���v�g
'----------------------------------------------------------------------------------
'   ��P�����y�K�{�z�F�\�[�X�t�@�C����(��΃p�X)
'   �@�@�@�@�@�@�@�@�F�K��zip�t�@�C���ł��鎖
'   ��Q�����y�C�Ӂz�F����t�@�C����(��΃p�X)
'   �@�@�@�@�@�@�@�@�F�@�ȗ������ꍇ��zip�t�@�C���Ɠ����t�H���_���ƂȂ�
'==================================================================================

'==================================================================================
'CopyHere�̃I�v�V���� + �łȂ���B
Const FOF_NOCONFIRMATION    = &H10    '�㏑���m�F�_�C�A���O��\�����Ȃ��i[���ׂď㏑��]�Ɠ����j�B


'==================================================================================
'�\�[�X�p�X
Dim SrcPath
'����p�X
Dim DestPath


Dim objFso
Dim objShell


'==================================================================================

'=======================================================
'  Zip�t�@�C���ł��邩
'=======================================================
Function IsZipFile(Src)

	IsZipFile = (Lcase(objFso.GetExtensionName(Src)) = "zip" )

End Function

'=======================================================
'  ����p�X�̉���
'=======================================================
Sub ResolveDestPath()
	
	If WScript.Arguments.Count > 1 Then
		'��2����������ΓK�p
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
'  Zip�𓀂����܂��B
'=======================================================
Sub UnCompressZip(Src, Dest)
	
	Call objShell.NameSpace(Dest).CopyHere(objShell.NameSpace(Src).Items, FOF_NOCONFIRMATION)

	
End Sub

'=======================================================
'  �t�H���_���ċA�I�ɍ쐬���܂�
'=======================================================
Sub Mkdirs(ByVal strPath)
	Dim strParent   ' �e�t�H���_

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
'   ���C������
'==================================================================================
Sub Main()

	'�\�[�X�t�@�C�����Ȃ��ꍇ�̓G���[
	If WScript.Arguments.Count = 0 Then
		
		WScript.Echo "�t�@�C�����w�肵�ĉ�����"
		WScript.Quit -1
		Exit Sub
	End If

	SrcPath = WScript.Arguments.Item(0)

	Set objFso   = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("Shell.Application")
	
	'�\�[�X�t�@�C�����Ȃ��ꍇ�̓G���[
	If objFso.FileExists(SrcPath) = False Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "�t�@�C��������܂���B:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'�g���q��zip�ł��邩
	If IsZipFile(SrcPath) = False Then
		
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "zip�t�@�C���ł͂���܂���B:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'����p�X�̉���
	Call ResolveDestPath()

	'���悪�t�@�C���Ƃ��Ă��łɑ��݂���ꍇ�̓G���[
	If objFso.FileExists(DestPath) Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "�R�s�[�悪���Ƀt�@�C���Ƃ��đ��݂��Ă��܂��B:" & DestPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'����t�@�C���̃t�H���_���쐬
	Call Mkdirs(DestPath)
	
	'����t�@�C�����쐬���ăR�s�[
	Call UnCompressZip(SrcPath, DestPath)
	
	'�p��
	Set objFso   = Nothing
	Set objShell = Nothing

End Sub

Call Main()
