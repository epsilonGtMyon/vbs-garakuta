Option Explicit
'==================================================================================
'   ���b�Z�[�W�\���X�N���v�g
'     �^�X�N�X�P�W���[������o�^���邱�ƂŃ��}�C���_�[����ɂł���
'----------------------------------------------------------------------------------
'   ��P�����y�K�{�z�F�\��������e
'   ��Q�����y�C�Ӂz�F�^�C�g��
'==================================================================================


Sub Main()

	Dim content
	Dim title

	If WScript.Arguments.Count < 1 Then
		WScript.Echo "���b�Z�[�W���w�肵�Ă��������B"
		WScript.Quit -1
		Exit Sub
	End If
	
	'�{��
	content = WScript.Arguments(0)
	
	'�^�C�g��
	If WScript.Arguments.Count > 1 Then
		title = WScript.Arguments(1)
	Else
		title = ""
	End If
	
	'���}�C���_�[����ɂ������̂� vbSystemModal �őO�ʂɂł�悤��
	Call MsgBox(content, vbOKOnly + vbInformation + vbSystemModal, title)

End Sub

Call Main()
