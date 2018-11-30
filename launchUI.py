
'''
08/04/2018

Author - Rishikesh Shendkar (rishikesh.shendkar@gmail.com)

Purpose - Used as an adapter to migrate outlook tasks to Jira.

Features Implemented:
1. Populate Jira with Microsoft outlook tasks
2. Check if the task is already present on Kanban board. If not, add.
3. Save the configuration

'''

# In-built Python module for sys.exit()
import sys
import PyJiraOut
import jira_rc

from PyQt5.QtWidgets import (QAction, QApplication, QDialog, QCheckBox,
		QDialog, QFormLayout, QGroupBox, QHBoxLayout,
		QLabel, QLineEdit, QMenu, QMenuBar, QPushButton, QTextEdit, QRadioButton,
		QVBoxLayout, QSystemTrayIcon, QErrorMessage )
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import (pyqtSlot, QSize, QTimer, QEventLoop)
from configparser import (SafeConfigParser, NoOptionError, NoSectionError)
from pathlib import Path



class Window(QDialog):
	def __init__(self):
		super(Window, self).__init__()

		self.runButton = QPushButton("Run")
		self.runButton.clicked.connect(self.confirm_btn)
		self.runButton.setEnabled(False)
		self.saveButton = QPushButton("Save Config")
		self.saveButton.clicked.connect(self.saveConfig)

		self.createFormGroupBox()
		self.getConfig()
		self.createSysTrayEntry()
		self.tray.activated.connect(self.iconActivated)
		self.setWindowTitle("JIRA-Outlook Adapter")
		self.tray.show()		
		self.showIconCheckBox.toggled.connect(self.tray.setVisible)
		self.showPwCheckBox.stateChanged.connect(lambda:self.pwToggle(self.showPwCheckBox))
		
		mainLayout = QVBoxLayout()
		mainLayout.addWidget(self.formGroupBox)
		mainLayout.addWidget(self.runButton)
		mainLayout.addStretch()
		mainLayout.addWidget(self.saveButton)
		self.setLayout(mainLayout)
		
	def formCheck(self):
		if len(self.jiraID.text()) > 0 and len(self.jiraUsername.text()) > 0 and len(self.jiraPassword.text()) > 0 and len(self.jiraPassword.text()) > 0 and len(self.boardName.text()) > 0 and len(self.boardID.text()) and len(self.jiraLink.text())  > 0:
			self.runButton.setEnabled(True)
		else:
			self.runButton.setEnabled(False)
			
	def getConfig(self):
		configFile = Path('adapter_config.ini')
		if configFile.exists():
			config = SafeConfigParser()
			config.read('adapter_config.ini')
			try:
				self.jiraID.setText(config.get('jiraout', 'jiraid'))
				self.jiraUsername.setText(config.get('jiraout', 'jirausername'))
				self.jiraPassword.setText(config.get('jiraout', 'jirapassword'))
				self.boardName.setText(config.get('jiraout', 'boardname'))
				self.boardID.setText(config.get('jiraout', 'boardid'))
				self.jiraLink.setText(config.get('jiraout', 'jiralink'))
			except NoSectionError:
				print("Section Error in the config file")
			except  NoOptionError:
				print("Option Error in the config file")
			
			
	def pwToggle(self,showPwCheckBox):
		if showPwCheckBox.isChecked() == True:
			self.jiraPassword.setEchoMode(QLineEdit.Normal)
		else:	
			self.jiraPassword.setEchoMode(QLineEdit.Password)
	
	@pyqtSlot()
	def saveConfig(self):
		if len(self.jiraID.text()) == 0 or len(self.jiraUsername.text()) == 0 or len(self.jiraPassword.text()) == 0 or len(self.jiraPassword.text()) == 0 or len(self.boardName.text()) == 0 or len(self.boardID.text()) == 0 or len(self.jiraLink.text()) == 0:
			error_dialog = QErrorMessage(self)
			error_dialog.showMessage('Please enter all the textfields.')
		else:
			configFile = open("adapter_config.ini", "w")
			configFile.truncate()
			configFile.close()
			config = SafeConfigParser()
			config.read('adapter_config.ini')
			config.add_section('jiraout')
			config.set('jiraout', 'jiraid', self.jiraID.text())
			config.set('jiraout', 'jirausername', self.jiraUsername.text())
			config.set('jiraout', 'jirapassword', self.jiraPassword.text())
			config.set('jiraout', 'boardname', self.boardName.text())
			config.set('jiraout', 'boardid', self.boardID.text())
			config.set('jiraout', 'jiralink', self.jiraLink.text())
			with open('adapter_config.ini', 'w') as f:
				config.write(f)

	def setVisible(self, visible):
		self.minimizeAction.setEnabled(visible)
		self.maximizeAction.setEnabled(not self.isMaximized())
		self.restoreAction.setEnabled(self.isMaximized() or not visible)
		super(Window, self).setVisible(visible)
		
	def iconActivated(self, reason):
		if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
			self.show()
	
	def createSysTrayEntry(self):
		# Create the menu
		self.minimizeAction = QAction("Minimize", self, triggered=self.hide)
		self.maximizeAction = QAction("Maximize", self, triggered=self.showMaximized)
		self.restoreAction = QAction("Restore", self, triggered=self.showNormal)
		self.quitAction = QAction("Quit", self, triggered=QApplication.instance().quit)
		self.menu = QMenu()
		self.menu.triggered[QAction].connect(self.processtrigger)
		self.menu.addAction(self.minimizeAction)
		self.menu.addAction(self.maximizeAction)
		self.menu.addAction(self.restoreAction)
		self.menu.addSeparator()
		self.menu.addAction(self.quitAction)
				
		# Create the tray
		self.tray = QSystemTrayIcon(self)
		self.tray.setContextMenu(self.menu)
		self.icon = QIcon(':/images/jira.png')
		self.tray.setIcon(self.icon)
		self.setWindowIcon(self.icon)
		self.tray.setVisible(True)

	def createFormGroupBox(self):
		self.formGroupBox = QGroupBox("JIRA Details")
		layout = QFormLayout()
		self.jiraID = QLineEdit()
		self.jiraUsername = QLineEdit()
		self.jiraPassword = QLineEdit()
		self.boardName = QLineEdit()
		self.boardID = QLineEdit()
		self.jiraLink = QLineEdit()
		self.jiraID.textChanged.connect(self.formCheck)
		self.jiraUsername.textChanged.connect(self.formCheck)
		self.jiraPassword.textChanged.connect(self.formCheck)
		self.boardName.textChanged.connect(self.formCheck)
		self.boardID.textChanged.connect(self.formCheck)
		self.jiraLink.textChanged.connect(self.formCheck)
		
		hbox1 = QHBoxLayout()
		self.oneTime = QRadioButton("One-time")
		self.scheduled = QRadioButton("Scheduled")
		self.scheduled.setChecked(True)
		hbox1.addWidget(self.oneTime)
		hbox1.addStretch()
		hbox1.addWidget(self.scheduled)

		hbox2 = QHBoxLayout()
		self.showPopups = QCheckBox("JIRA-Card pop-ups")
		self.showPopups.setChecked(True)
		self.showIconCheckBox = QCheckBox("Show icon")
		self.showIconCheckBox.setChecked(True)
		self.showPwCheckBox = QCheckBox("Show password")
		self.jiraPassword.setEchoMode(QLineEdit.Password)
		hbox2.addWidget(self.showPopups)
		hbox2.addStretch()
		hbox2.addWidget(self.showIconCheckBox)
		
		layout.addRow(QLabel("JIRA-ID(LAN-ID):"),self.jiraID)
		layout.addRow(QLabel("JIRA-Username:"),self.jiraUsername)
		layout.addRow(QLabel("JIRA-Password:"),self.jiraPassword)
		layout.addWidget(self.showPwCheckBox)
		layout.addRow(QLabel("Board name:"),self.boardName)
		layout.addRow(QLabel("Board ID:"),self.boardID)
		layout.addRow(QLabel("Jira Link:"),self.jiraLink)
		layout.addRow(hbox1)
		layout.addRow(hbox2)
		self.formGroupBox.setLayout(layout)

	def processtrigger(self,q):
		print (q.text()+" is triggered")
		if q.text() == "Quit":
			self.tray.hide()
			QApplication.instance().quit
			sys.exit()


	@pyqtSlot()
	def confirm_btn(self):
		if self.oneTime.isChecked() == True:
			print("One-Time")
			PyJiraOut.syncTasksToJira(self.jiraID.text(), self.jiraUsername.text(), self.jiraPassword.text(), self.boardName.text(), self.boardID.text(), self.jiraLink.text())
		if self.scheduled.isChecked() == True:
			print("Scheduled")
			while True:
				PyJiraOut.syncTasksToJira(self.jiraID.text(), self.jiraUsername.text(), self.jiraPassword.text(), self.boardName.text(), self.boardID.text(), self.jiraLink.text())
				loop = QEventLoop()
				QTimer.singleShot(9000000, loop.quit)
				loop.exec_()
		print("Success")

if __name__ == '__main__':
	app = QApplication(sys.argv)

	app_icon = QIcon()
	app_icon.addFile('/images/256x256.png', QSize(256,256))
	app.setWindowIcon(app_icon)
	
	if not QSystemTrayIcon.isSystemTrayAvailable():
		QMessageBox.critical(None, "Systray",
				"I couldn't detect any system tray on this system.")
		sys.exit(1)
	QApplication.setQuitOnLastWindowClosed(False)

	window = Window()
	window.show()
	sys.exit(app.exec_())
