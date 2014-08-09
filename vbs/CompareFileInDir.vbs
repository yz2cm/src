Option Explicit
'/**********************************************************************************************/
'
' ディレクトリAとBを鏡合わせで、同名ファイル同士の内容を比較し、結果をログファイルに出力する。
'
'  - 比較したいディレクトリを「設定項目」セクションに定義すること。
'  - ログファイルは本スクリプトと同階層に出力される（<本スクリプトファイル名>.txt）
'  - 1と2のディレクトリ構成は任意。
'  - 同名ファイルが、一方のディレクトリに複数存在するケースには非対応。
'
'/**********************************************************************************************/

Dim objFSO
Dim targetDirPath_1
Dim targetDirPath_2 
Dim objWshShell
Dim logFilePath
Dim isExistDiff

'/**************************************************/
'/*              設定項目                          */
'/**************************************************/

targetDirPath_1 = "G:\test1"
targetDirPath_2 = "G:\test2"

'/**************************************************/

WScript.Echo "以下のディレクトリを比較します" & vbCrLf & _
			 "    " & targetDirPath_1 & vbCrLf & _
			 "    " & targetDirPath_2

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objWshShell = WScript.CreateObject("WScript.Shell")

logFilePath = objFSO.BuildPath( _
	objFSO.GetParentFolderName(Wscript.ScriptFullName), _
	objFSO.GetBaseName(Wscript.ScriptName) & ".txt")

'// 結果ファイルの削除
If objFSO.FileExists(logFilePath) Then
	Call objFSO.DeleteFile(logFilePath, True)
End If

isExistDiff = False
Call CompareFileContentInBuddyDirs(targetDirPath_1, targetDirPath_2, "", logFilePath)

'// 結果の表示
If isExistDiff Then
	WScript.Echo "相違点のあるファイルが検出されました" & vbCrLf & _
				 "    " & logFilePath
	Call objWshShell.Run("cmd /c notepad " & logFilePath, 0, False)
Else
	WScript.Echo "相違点のあるファイルは見つかりませんでした"
End If

Set objFSO = Nothing
Set objWshShell = Nothing
'/**************************************************/
'
' CompareFileContentInBuddyDirs
'
' ディレクトリ1と2を鏡合わせで、同名ファイル同士を内容比較する。
' [引数]
'   dirPath_1   : ディレクトリ1
'   dirPath_2   : ディレクトリ2
'   filePath    : 空文字列("")固定
'   logFilePath : 比較結果の出力先ログファイル（追記型）
'
'/**************************************************/
Sub CompareFileContentInBuddyDirs(dirPath_1, dirPath_2, filePath, logFilePath)

	Dim objFolder
	Dim objSubFolder
	Dim objFile
	Dim cmdLine

	'// 指定フォルダのフォルダオブジェクトを生成
	Set objFolder = objFSO.GetFolder(dirPath_1)

	'//
	'// 指定フォルダのサブフォルダ群を再帰探索
	'//
	For Each objSubFolder In objFolder.SubFolders
		Call CompareFileContentInBuddyDirs(objSubFolder.Path, dirPath_2, filePath, logFilePath)
	Next

	'//
	'// 指定フォルダ直下のファイル群
	'//
	For Each objFile In objFolder.Files
		If Len(filePath) < 1 Then
			Call CompareFileContentInBuddyDirs(dirPath_2, dirPath_2, objFile.Path, logFilePath)
		ElseIf objFile.Name = objFSO.GetFileName(filePath) Then
			'// ファイル名が一致 → ファイル内容を比較
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
