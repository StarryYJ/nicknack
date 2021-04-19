from PyQt5 import QtCore, QtWidgets
import win32gui, win32api, win32con
import time
import win32clipboard as clipboard
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.date import DateTrigger


class WeChatTrick:

	def __init__(self):
		pass

	def __paste_file(self):
		"""
		Mimic click with 'win32api', send we-chat message by 'pressing' shortcut 'Ctrl + v'.
		"""

		win32api.keybd_event(17, 0, 0, 0)  # ctrl - Key code - 17
		win32api.keybd_event(86, 0, 0, 0)  # v - Key code - 86

		win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # release key
		win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

	def __wc_send(self):
		"""
		Mimic click with 'win32api', send we-chat message by 'pressing' shortcut 'Ctrl + Enter'.
		Keystroke can be adjusted according to user's WeChat setting.
		"""

		win32api.keybd_event(17, 0, 0, 0)  # alt - Key code - 17
		win32api.keybd_event(13, 0, 0, 0)  # Enter - Key code - 13

		win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
		win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

	def wc_send_file(self, name, file_path):
		"""
		Mimic click with 'win32gui', get window handle and activates the window. Then send file to it

		:param file_path: absolute path of the file you're going to send
		:param name: name of a separate WeChat window that you're going to send message/files to
		"""

		window = win32gui.FindWindow('ChatWnd', name)  # get window handle
		win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)
		window = win32gui.FindWindow('ChatWnd', name)

		# prepare material to send
		app = QtWidgets.QApplication([])
		data = QtCore.QMimeData()
		url = QtCore.QUrl.fromLocalFile(file_path)
		data.setUrls([url])
		app.clipboard().setMimeData(data)

		# get control, paste and send file
		win32gui.SetForegroundWindow(window)
		self.__paste_file()
		self.__wc_send()

	def __copy_text(self, txt_str):
		"""
		Cache information to the clipboard

		:param txt_str: a string format message
		"""

		clipboard.OpenClipboard()
		clipboard.EmptyClipboard()
		clipboard.SetClipboardData(win32con.CF_UNICODETEXT, txt_str)
		clipboard.CloseClipboard()

	def __sendTaskLog(self, name, file):
		"""
		Send message on we-chat.
		"""

		file = open(file, mode='r', encoding='UTF-8')
		s = file.read()

		window = win32gui.FindWindow('ChatWnd', name)  # get window handle
		win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)
		window = win32gui.FindWindow('ChatWnd', name)

		# get control, paste stuff and send
		win32gui.SetForegroundWindow(window)
		self.__copy_text(s)
		self.__paste_file()
		self.__wc_send()

	def wc_send_message(self, name, file):
		"""
		Send message stored in a text file to a friend or group on we-chat

		:param name: Who/Which group you're gonna send message to
		:param file: the txt file that the message comes from
		"""

		# set task and trigger, then execute and shutdown
		scheduler = BackgroundScheduler()
		trigger = DateTrigger(run_date=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time() + 5)))

		def fun():
			self.__sendTaskLog(name, file)

		scheduler.add_job(fun, trigger)
		try:
			scheduler.start()
			time.sleep(8)
			scheduler.shutdown(wait=False)
		except (KeyboardInterrupt, SystemExit):
			pass


if __name__ == '__main__':
	bot = WeChatTrick()

	# Send files
	bot.wc_send_file('42', r'E:/wcbot/test.xls')

	# Send message from a txt file
	bot.wc_send_message('42', 'send.txt')
