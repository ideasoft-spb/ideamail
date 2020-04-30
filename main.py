import imapclient
import pyzmail
import sys
import smtplib
import imaplib
import pickle
import configparser
import base64
import os
import pyAesCrypt
import secrets
import email
from string import ascii_letters, digits, punctuation
from PyQt5 import QtWidgets, QtGui, QtCore


# Needed to increase IMAP requests amount
imaplib._MAXLINE = 1000000000


class LoginWindow(QtWidgets.QDialog):
    def __init__(self, *args, **kwargs):
        super(LoginWindow, self).__init__(*args, **kwargs)
        self.login_input = QtWidgets.QLineEdit()
        self.password_input = QtWidgets.QLineEdit()
        self.password_input.setEchoMode(QtWidgets.QLineEdit.Password)
        self.box = QtWidgets.QComboBox()
        self.box.addItems(["Yandex", "Gmail"])
        self.close_btn = QtWidgets.QPushButton("Login")
        self.close_btn.clicked.connect(self.close)
        self.form = QtWidgets.QFormLayout()
        self.form.addRow(self.box)
        self.form.addRow("Login:", self.login_input)
        self.form.addRow("Password:", self.password_input)
        self.form.addRow(self.close_btn)
        self.setLayout(self.form)
        self.setWindowTitle("Login")
        self.setWindowIcon(QtGui.QIcon(os.path.join('icons', 'user.png')))

    def exec_(self):
        super(LoginWindow, self).exec_()
        return (self.login_input.text(), self.password_input.text(), self.box.currentText())


class ChecklistWidget(QtWidgets.QWidget):
    def __init__(self, stringlist=[], checked=False, parent=None):
        super(ChecklistWidget, self).__init__(parent)
        self.setup()
        self.data = {}
        self.addNewElements(stringlist)

    def setup(self):
        self.model = QtGui.QStandardItemModel()
        self.listView = QtWidgets.QListView()
        self.listView.setModel(self.model)
        self.selectButton = QtWidgets.QPushButton('Select all')
        self.unselectButton = QtWidgets.QPushButton('Unselect all')
        hbox = QtWidgets.QHBoxLayout()
        hbox.addWidget(self.selectButton)
        hbox.addWidget(self.unselectButton)
        vbox = QtWidgets.QVBoxLayout(self)
        vbox.addWidget(self.listView)
        vbox.addLayout(hbox)
        self.selectButton.clicked.connect(self.select)
        self.unselectButton.clicked.connect(self.unselect)

    def getSelectedIndexes(self)-> list:
        self.refresh_data()
        items = list(self.data.items())
        return [i for i in range(len(items)) if items[i][1]]

    def getSelected(self)-> list:
        self.refresh_data()
        return [key for key in self.data if self.data[key]]

    def select(self):
        for i in self.data:
            self.data[i] = True
        self.refresh()

    def unselect(self):
        for i in self.data:
            self.data[i] = False
        self.refresh()

    def printData(self):
        self.refresh_data()
        print(self.data)

    def addNewElements(self, iterable, checked=False):
        for i in iterable:
            self.addNewElement(i, checked=checked)

    def addNewElement(self, el: str, checked=False):
        if el not in self.data:
            item = QtGui.QStandardItem(el)
            item.setCheckable(True)
            item.setCheckState((QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked))
            self.model.appendRow(item)
            self.data[el] = checked
        self.refresh()

    def refresh(self):
        self.model.clear()
        for i in self.data:
            item = QtGui.QStandardItem(i)
            item.setCheckable(True)
            item.setCheckState((QtCore.Qt.Checked if self.data[i] else QtCore.Qt.Unchecked))
            self.model.appendRow(item)

    def refresh_data(self):
        self.data = {self.model.item(i).text(): (True if self.model.item(i).checkState() == QtCore.Qt.Checked else False)
        for i in range(self.model.rowCount())}

    def deleteSelected(self, function=None):
        indexes = self.getSelectedIndexes()
        while indexes:
            i = indexes.pop(0)
            if function: function(self.model.item(i))
            self.model.removeRow(i)
            indexes = list(map(lambda x: x - 1 if (x > i) else x, indexes))
        self.refresh_data()


class NewMessageWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.setWindowIcon(QtGui.QIcon("icons/logo.png"))
        self.setWindowTitle("New message")
        self.setup()
        self.parent = parent
        self.files = {}
        self.initContacts()

    def setup(self):
        # central -> hbox -> vbox  -> QLabel, self.reciever, contactsbtn, QLabel, self.subject, QLabel, self.body_type, QLabel, self.body, sendbtn
        #                    vbox2 -> QLabel, self.fileChooseWidget, (hbox2 -> addbtn, deletebtn)

        central = QtWidgets.QWidget()
        vbox = QtWidgets.QVBoxLayout()
        vbox2 = QtWidgets.QVBoxLayout()
        hbox = QtWidgets.QHBoxLayout()
        sendbtn = QtWidgets.QPushButton("Send")
        self.reciever = QtWidgets.QComboBox()
        contactsbtn = QtWidgets.QPushButton("Refresh contacts")
        self.subject = QtWidgets.QLineEdit()
        self.body_type = QtWidgets.QComboBox()
        self.body = QtWidgets.QTextEdit()
        addbtn = QtWidgets.QPushButton("Add file")
        deletebtn = QtWidgets.QPushButton("Remove selected")
        hbox2 = QtWidgets.QHBoxLayout()
        self.fileChooseWidget = ChecklistWidget()

        hbox2.addWidget(addbtn)
        hbox2.addWidget(deletebtn)
        vbox2.addWidget(QtWidgets.QLabel("Attachments"))
        vbox2.addWidget(self.fileChooseWidget)
        vbox2.addLayout(hbox2)
        vbox.addWidget(QtWidgets.QLabel("Reciever"))
        vbox.addWidget(self.reciever)
        vbox.addWidget(contactsbtn)
        vbox.addWidget(QtWidgets.QLabel("Subject"))
        vbox.addWidget(self.subject)
        vbox.addWidget(QtWidgets.QLabel("Text type"))
        vbox.addWidget(self.body_type)
        vbox.addWidget(QtWidgets.QLabel("Email text"))
        vbox.addWidget(self.body)
        vbox.addWidget(sendbtn)
        hbox.addLayout(vbox)
        hbox.addLayout(vbox2)
        central.setLayout(hbox)
        self.setCentralWidget(central)

        self.body_type.addItems(["Plain", "HTML"])

        contactsbtn.clicked.connect(self.refreshContacts)
        deletebtn.clicked.connect(self.deleteSelected)
        addbtn.clicked.connect(self.selectFile)
        sendbtn.clicked.connect(self.sendEmail)

    def refreshContacts(self):
        self.reciever.clear()
        self.reciever.addItems(self.parent.contacts)

    def initContacts(self):
        self.reciever.addItems(list(self.parent.contacts.keys()))

    def selectFile(self):
        file = QtWidgets.QFileDialog.getOpenFileName()
        self.fileChooseWidget.addNewElement(os.path.basename(file[0]))
        self.files[file[0]] = os.path.basename(file[0])

    def deleteSelected(self):
        self.fileChooseWidget.deleteSelected(function=self.deleteFilePath)

    def deleteFilePath(self, item):
        for i in self.files:
            if self.files[i] == item.text():
                del self.files[i]
                break

    def sendEmail(self):
        body_type = self.body_type.currentText()
        if body_type == "HTML":
            body = self.body.toHtml()
        else:
            body = self.body.toPlainText()
        subject = self.subject.text()
        files = self.files
        reciever = self.reciever.currentText()
        message = email.mime.multipart.MIMEMultipart("alternative")
        message["Subject"] = subject
        message["To"] = self.parent.contacts[reciever]
        message["From"] = self.parent.USERNAME
        if body_type == "HTML":
            body = email.mime.text.MIMEText(body, "html")
        else:
            body = email.mime.text.MIMEText(body, "plain")
        message.attach(body)
        self.parent.smtpSession.sendmail(self.parent.USERNAME, self.parent.contacts[reciever], message.as_string())


class AddContactDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.setup()
        self.result = False
        self.setWindowTitle("New contact")

    def setup(self):
        # vbox -> QLabel, self.name, QLabel, self.email, hbox -> okbtn, cancelbtn
        vbox = QtWidgets.QVBoxLayout()
        self.name = QtWidgets.QLineEdit()
        self.email = QtWidgets.QLineEdit()
        hbox = QtWidgets.QHBoxLayout()
        okbtn = QtWidgets.QPushButton("OK")
        cancelbtn = QtWidgets.QPushButton("Cancel")

        hbox.addWidget(okbtn)
        hbox.addWidget(cancelbtn)
        vbox.addWidget(QtWidgets.QLabel("Name:"))
        vbox.addWidget(self.name)
        vbox.addWidget(QtWidgets.QLabel("Email:"))
        vbox.addWidget(self.email)
        vbox.addLayout(hbox)
        self.setLayout(vbox)

        okbtn.clicked.connect(self.ok)
        cancelbtn.clicked.connect(self.cancel)

    def ok(self):
        if self.email.text(): self.result = True
        self.close()

    def cancel(self):
        self.result = False
        self.close()

    def exec_(self):
        super(AddContactDialog, self).exec_()
        return self.result, self.name.text(), self.email.text()


class ContactsWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.parent = parent
        self.contacts = self.parent.contacts
        self.setup()
        self.setWindowTitle("My contacts")
        self.contacts_list.addNewElements(self.contacts)

    def setup(self):
        # central -> vbox -> ChecklistWidget, grid -> deletebtn, addbtn, importbtn, exportbtn
        central = QtWidgets.QWidget()
        vbox = QtWidgets.QVBoxLayout()
        self.contacts_list = ChecklistWidget()
        grid = QtWidgets.QGridLayout()
        deletebtn = QtWidgets.QPushButton("Delete")
        addbtn = QtWidgets.QPushButton("Add")
        importbtn = QtWidgets.QPushButton("Import")
        exportbtn = QtWidgets.QPushButton("Export")

        grid.addWidget(addbtn, 0, 0)
        grid.addWidget(deletebtn, 0, 1)
        grid.addWidget(importbtn, 1, 0)
        grid.addWidget(exportbtn, 1, 1)
        vbox.addWidget(self.contacts_list)
        vbox.addLayout(grid)
        central.setLayout(vbox)
        self.setCentralWidget(central)

        addbtn.clicked.connect(self.addContact)
        deletebtn.clicked.connect(self.deleteSelected)

    def addContact(self):
        d = AddContactDialog(parent=self)
        d.show()
        out = d.exec_()
        if out[0]:
            if out[1]:
                self.contacts_list.addNewElement(out[1])
                self.contacts[out[1]] = out[2]
            elif out[2]:
                self.contacts_list.addNewElement(out[2])
                self.contacts[out[2]] = out[2]

    def deleteSelected(self):
        self.contacts_list.deleteSelected(function=self.deleteContact)

    def deleteContact(self, item):
        name = item.text()
        if name in self.contacts:
            del self.contacts[name]
            return
        for i in self.contacts:
            if self.contacts[i] == name:
                del self.contacts[i]
                break


class FoldersWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.parent = parent
        self.setup()
        self.setWindowTitle("Email folders")

    def setup(self):
        central = QtWidgets.QWidget()

        self.setCentralWidget(central)

class MessagesWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.setWindowTitle("Login")
        self.data = {}
        self.contacts = {}
        if os.path.isfile('.ideamail.ini') and os.path.isfile(".ideamail.data.encr"):
            try:
                config = configparser.RawConfigParser()
                config.read('.ideamail.ini')
                pyAesCrypt.decryptFile('.ideamail.data.encr', '.ideamail.data', config['IdeaMail']['Password'], 1024*64)
                os.remove('.ideamail.data.encr')
                with open('.ideamail.data', 'br') as f:
                    data = pickle.load(f)
                self.USERNAME = data.get("USERNAME")
                self.PASSWORD = data.get("PASSWORD")
                self.MAILHOST = data.get("HOST")
                self.data = data.get("DATA")
                self.contacts = data.get("CONTACTS")
                self.serverLogin(self.USERNAME, self.PASSWORD, self.MAILHOST)
            except:
                self.showLoginDialog()
                self.createConfig()
        else:
            self.showLoginDialog()
            self.createConfig()

        self.createMainLayout()
        self.setWindowTitle("My messages")
        self.loadMessagesFromData()

    def addSubjectToModel(self, subject):
        if subject:
            subject = email.header.decode_header(subject.decode())[0][0]
            if isinstance(subject, str):
                self.model.appendRow(QtGui.QStandardItem(subject))
            else:
                self.model.appendRow(QtGui.QStandardItem(subject.decode()))
        else:
            self.model.appendRow(QtGui.QStandardItem("[No subject]"))

    def loadMessagesFromData(self):
        self.model.clear()
        for i in self.data:
            subject = self.data[i].subject
            self.addSubjectToModel(subject)

    def serverLogin(self, username, password, mailhost):
        self.smtpSession = smtplib.SMTP('smtp.{}.com'.format(self.MAILHOST.lower()), 587)
        self.smtpSession.starttls()
        self.imapSession = imapclient.IMAPClient('imap.{}.com'.format(self.MAILHOST.lower()), ssl=True)
        try:
            self.smtpSession.login(self.USERNAME, self.PASSWORD)
            self.imapSession.login(self.USERNAME, self.PASSWORD)
        except:
            errorWindow = QtWidgets.QMessageBox.critical(self, "Login error", "Check that login and password are valid.")
            self.showLoginDialog()

    def showLoginDialog(self):
        loginDialog = LoginWindow()
        loginDialog.show()
        out = loginDialog.exec_()
        if all(out[0:1]):
            self.USERNAME, self.PASSWORD, self.MAILHOST = out
            self.serverLogin(self.USERNAME, self.PASSWORD, self.MAILHOST)
        else:
            self.deleteLater()

    def createConfig(self):
        password = ''.join(secrets.choice(ascii_letters + punctuation + digits) for i in range(20))
        data = {'USERNAME': self.USERNAME,
                'PASSWORD': self.PASSWORD,
                "HOST": self.MAILHOST,
                'DATA': self.data,
                "CONTACTS": self.contacts}
        config = configparser.RawConfigParser()
        config['IdeaMail'] = {"Password":password}
        with open('.ideamail.ini', 'w') as f:
            config.write(f)
        with open(".ideamail.data", "bw") as f:
            pickle.dump(data, f)
        pyAesCrypt.encryptFile('.ideamail.data', '.ideamail.data.encr', password, 1024 * 64)
        os.remove('.ideamail.data')

    def createMainLayout(self):
        # Window
        # centralwidget -> vbox -> listview, hbox
        # menubar -> fileMenu -> logoutAction, exitAction
        # toolbar -> folderSelectAction, newMessageAction, refreshAction, contactsAction
        # self.statusbar
        # TODO: replace QListView with my checkable ListView widget
        hbox = QtWidgets.QHBoxLayout()
        listview = QtWidgets.QListView()
        vbox = QtWidgets.QVBoxLayout()
        centralwidget = QtWidgets.QWidget()
        self.model = QtGui.QStandardItemModel() # ItemModel for QListView
        toolbar = self.addToolBar("Main")
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('Почта')
        logoutAction = QtWidgets.QAction(QtGui.QIcon("icons/logout.png"), 'Log out', self)
        exitAction = QtWidgets.QAction(QtGui.QIcon("icons/close.png"), 'Выход', self)
        folderSelectAction = QtWidgets.QAction(QtGui.QIcon("icons/folder.png"), "Folder selecting", self)
        newMessageAction = QtWidgets.QAction(QtGui.QIcon("icons/new_message.png"), "New email", self)
        refreshAction = QtWidgets.QAction(QtGui.QIcon("icons/refresh.png"), "Refresh inbox", self)
        contactsAction = QtWidgets.QAction(QtGui.QIcon("icons/contacts.png"), "Open contacts", self)

        listview.setModel(self.model)
        listview.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers) # Text in listview can't be edited
        vbox.addWidget(listview)
        vbox.addLayout(hbox)
        self.setCentralWidget(centralwidget)
        centralwidget.setLayout(vbox)
        logoutAction.setShortcut("Ctrl+Shift+L")
        exitAction.setShortcut('Ctrl+Q')
        fileMenu.addAction(exitAction)
        fileMenu.addSeparator()
        fileMenu.addAction(logoutAction)
        folderSelectAction.setStatusTip("Choose email forder or label")
        newMessageAction.setStatusTip("Create a new email and send it")
        refreshAction.setStatusTip("Download all messages from selected folder")
        contactsAction.setStatusTip("Manage your contacts")
        toolbar.addAction(refreshAction)
        toolbar.addAction(folderSelectAction)
        toolbar.addAction(newMessageAction)
        toolbar.addAction(contactsAction)
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        self.statusbar = self.statusBar()
        self.statusbar.show()

        logoutAction.triggered.connect(self.logout)
        exitAction.triggered.connect(self.close)
        newMessageAction.triggered.connect(self.newMessage)
        refreshAction.triggered.connect(self.getMessages)
        contactsAction.triggered.connect(self.openContacts)

    def openContacts(self):
        self.contactsWindow = ContactsWindow(parent=self)
        self.contactsWindow.show()

    def newMessage(self):
        self.window = NewMessageWindow(parent=self)
        self.window.show()

    def logout(self):
        self.smtpSession.close()
        self.imapSession.logout()
        try:
            os.remove('.ideamail.data')
            os.remove('.ideamail.ini')
        except: pass
        self.deleteLater()

    def getMessages(self):
        self.model.clear()
        self.data = {}
        self.imapSession.select_folder('INBOX', readonly=0)
        UIDS = self.imapSession.search(['ALL'])
        UIDS.reverse()
        length = len(UIDS)
        indicator = QtWidgets.QProgressDialog("Downloading messages...", "Отмена", 0, length, parent=self)
        indicator.setWindowModality(QtCore.Qt.WindowModal)
        indicator.setWindowTitle("Downloading...")
        indicator.show()

        for i in UIDS:
            indicator.setValue(indicator.value() + 1)
            loop = QtCore.QEventLoop()
            QtCore.QTimer.singleShot(10, loop.quit)
            loop.exec_()

            if indicator.wasCanceled():
                indicator.setValue(length)
                break

            msgid, data = list(self.imapSession.fetch([i], ['ENVELOPE']).items())[0]
            envelope = data[b'ENVELOPE']
            self.data[i] = envelope
            subject = envelope.subject
            self.addSubjectToModel(subject)

    def closeEvent(self, event):
        reply = QtWidgets.QMessageBox.question(self, "Exit", "Are you sure you want to exit?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            try:
                self.createConfig()
                self.imapSession.logout()
                self.smtpSession.close()
            except: pass
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MessagesWindow()
    window.resize(400, 400)
    window.setWindowIcon(QtGui.QIcon("icons/logo.png"))
    window.show()
    sys.exit(app.exec_())
