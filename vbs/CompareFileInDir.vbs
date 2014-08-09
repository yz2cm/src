Option Explicit
'/**********************************************************************************************/
'
' �f�B���N�g��A��B�������킹�ŁA�����t�@�C�����m�̓��e���r���A���ʂ����O�t�@�C���ɏo�͂���B
'
'  - ��r�������f�B���N�g�����u�ݒ荀�ځv�Z�N�V�����ɒ�`���邱�ƁB
'  - ���O�t�@�C���͖{�X�N���v�g�Ɠ��K�w�ɏo�͂����i<�{�X�N���v�g�t�@�C����>.txt�j
'  - 1��2�̃f�B���N�g���\���͔C�ӁB
'  - �����t�@�C�����A����̃f�B���N�g���ɕ������݂���P�[�X�ɂ͔�Ή��B
'
'/**********************************************************************************************/

Dim objFSO
Dim targetDirPath_1
Dim targetDirPath_2 
Dim objWshShell
Dim logFilePath
Dim isExistDiff

'/**************************************************/
'/*              �ݒ荀��                          */
'/**************************************************/

targetDirPath_1 = "G:\test1"
targetDirPath_2 = "G:\test2"

'/**************************************************/

WScript.Echo "�ȉ��̃f�B���N�g�����r���܂�" & vbCrLf & _
			 "    " & targetDirPath_1 & vbCrLf & _
			 "    " & targetDirPath_2

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objWshShell = WScript.CreateObject("WScript.Shell")

logFilePath = objFSO.BuildPath( _
	objFSO.GetParentFolderName(Wscript.ScriptFullName), _
	objFSO.GetBaseName(Wscript.ScriptName) & ".txt")

'// ���ʃt�@�C���̍폜
If objFSO.FileExists(logFilePath) Then
	Call objFSO.DeleteFile(logFilePath, True)
End If

isExistDiff = False
Call CompareFileContentInBuddyDirs(targetDirPath_1, targetDirPath_2, "", logFilePath)

'// ���ʂ̕\��
If isExistDiff Then
	WScript.Echo "����_�̂���t�@�C�������o����܂���" & vbCrLf & _
				 "    " & logFilePath
	Call objWshShell.Run("cmd /c notepad " & logFilePath, 0, False)
Else
	WScript.Echo "����_�̂���t�@�C���͌�����܂���ł���"
End If

Set objFSO = Nothing
Set objWshShell = Nothing
'/**************************************************/
'
' CompareFileContentInBuddyDirs
'
' �f�B���N�g��1��2�������킹�ŁA�����t�@�C�����m����e��r����B
' [����]
'   dirPath_1   : �f�B���N�g��1
'   dirPath_2   : �f�B���N�g��2
'   filePath    : �󕶎���("")�Œ�
'   logFilePath : ��r���ʂ̏o�͐惍�O�t�@�C���i�ǋL�^�j
'
'/**************************************************/
Sub CompareFileContentInBuddyDirs(dirPath_1, dirPath_2, filePath, logFilePath)

	Dim objFolder
	Dim objSubFolder
	Dim objFile
	Dim cmdLine

	'// �w��t�H���_�̃t�H���_�I�u�W�F�N�g�𐶐�
	Set objFolder = objFSO.GetFolder(dirPath_1)

	'//
	'// �w��t�H���_�̃T�u�t�H���_�Q���ċA�T��
	'//
	For Each objSubFolder In objFolder.SubFolders
		Call CompareFileContentInBuddyDirs(objSubFolder.Path, dirPath_2, filePath, logFilePath)
	Next

	'//
	'// �w��t�H���_�����̃t�@�C���Q
	'//
	For Each objFile In objFolder.Files
		If Len(filePath) < 1 Then
			Call CompareFileContentInBuddyDirs(dirPath_2, dirPath_2, objFile.Path, logFilePath)
		ElseIf objFile.Name = objFSO.GetFileName(filePath) Then
			'// �t�@�C��������v �� �t�@�C�����e���r
			cmdLine = "cmd /c fc " & filePath & " " & objFile.Path & " >>" & logFilePath
			If objWshShell.Run(cmdLine, 0, True) Then
				isExistDiff = True
			End If
		End If
	Next

	Set objFolder = Nothing
	Set objSubFolder = Nothing
	Set objFile = Nothing

End Sub
