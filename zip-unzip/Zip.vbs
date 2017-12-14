Option Explicit
'==================================================================================
'   �Ȉ�ZIP���k�X�N���v�g
'----------------------------------------------------------------------------------
'   ��P�����y�K�{�z�F�\�[�X�t�@�C����(��΃p�X)
'   ��Q�����y�C�Ӂz�F����t�@�C����(��΃p�X)
'   �@�@�@�@�@�@�@�@�F�@�ȗ������ꍇ�̓\�[�X�t�@�C����.zip�ɂȂ�B
'==================================================================================

'==================================================================================
'�\�[�X�p�X
Dim SrcPath
'����p�X
Dim DestPath


Dim objFso
Dim objShell


'==================================================================================

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
	
	DestPath = objFso.BuildPath(SrcParentDir, SrcBaseName & ".zip")

End Sub

'=======================================================
'  Zip���k�����܂��B
'=======================================================
Sub CompressZip(Src, Dest)

	Dim ItemCount
	ItemCount = 0

	'�V����zip�t�@�C�����쐬���܂�
	With objFso.CreateTextFile(Dest)
		.Write "PK" & Chr(5) & Chr(6) & String(18,0)
		.Close
	End With
	
	'zip�t�@�C���Ɉ����Ŏ󂯎�����t�@�C�������܂�
	With objShell.NameSpace(Dest)
	
		.CopyHere Src
	
		ItemCount = ItemCount + 1
		Do Until .Items.Count = ItemCount
			'�t�@�C���̃R�s�[���I���܂ő҂K�v������悤���B
			WScript.sleep 1000
		Loop
	
	End With
	
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
	If objFso.FileExists(SrcPath) = False And objFso.FolderExists(SrcPath) = False Then
		Set objFso   = Nothing
		Set objShell = Nothing
		
		WScript.Echo "�t�@�C��������܂���B:" & SrcPath
		WScript.Quit -1
		Exit Sub
	End If
	
	'����p�X�̉���
	Call ResolveDestPath()
	
	'����t�@�C���̃t�H���_���쐬
	Call Mkdirs(objFso.GetParentFolderName(DestPath))
	
	'����t�@�C�����쐬���ăR�s�[
	Call CompressZip(SrcPath, DestPath)
	
	'�p��
	Set objFso   = Nothing
	Set objShell = Nothing

End Sub

Call Main()
