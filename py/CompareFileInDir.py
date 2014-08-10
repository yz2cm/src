# coding: utf-8

import os
import subprocess
import ctypes

targetdir_1 = 'G:\\test1'
targetdir_2 = 'G:\\test2'

user32 = ctypes.windll.user32

#
# メッセージボックス関数
#
def msg_box(message):
	user32.MessageBoxW(0, message.decode('utf-8'), os.path.basename(__file__).decode('utf-8'), 0x00000040)

#
# 指定ディレクトリを再帰探索し、ファイルのディレクトリパス名とベース名をジェネレートする
# - target_dir : 探索対象ディレクトリ
# + dest       : ファイルのディレクトリパス名
# + file       : ファイルのベース名
#
def list_files_recursive(target_dir):
	for dest, dirs, files in os.walk(target_dir):
		for file in files:
			yield dest, file

#
# Main
#
logFilePath = os.path.join(os.path.abspath(os.path.dirname(__file__)), os.path.splitext(os.path.basename(__file__))[0] + '.txt')

msg_box('以下のディレクトリを比較します\r\n' + targetdir_1 + '\r\n' + targetdir_2)

if os.path.exists(logFilePath):
	os.remove(logFilePath)
	
is_exist_diff = False
for dir_1, file_1 in list_files_recursive(targetdir_1):
	for dir_2, file_2 in list_files_recursive(targetdir_2):
		if file_1 == file_2:
			fpath_1 = os.path.join(dir_1, file_1)
			fpath_2 = os.path.join(dir_2, file_2)
			try:
				subprocess.check_call(["cmd", "/c fc " + fpath_1 + " " + fpath_2 + " >>" + logFilePath])
			except:
				is_exist_diff = True

if not is_exist_diff:
	msg_box("相違点はありませんでした")
	exit()

msg_box("相違点を検出しました")
subprocess.check_call(["notepad", logFilePath])
