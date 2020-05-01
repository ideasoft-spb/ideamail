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
from email.mime.application import MIMEApplication

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
        return self.login_input.text(), self.password_input.text(), self.box.currentText()


class ChecklistWidget(QtWidgets.QWidget):
    def __init__(self, stringlist=None, checked=False, parent=None, editable=False):
        super(ChecklistWidget, self).__init__(parent)
        if stringlist is None:
            stringlist = []
        self.setup()
        self.data = {}
        self.addNewElements(stringlist)
        if not editable:
            self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

    def setup(self):
        self.model = QtGui.QStandardItemModel()
        self.listView = QtWidgets.QListView()
        self.listView.setModel(self.model)
        hbox = QtWidgets.QHBoxLayout()
        vbox = QtWidgets.QVBoxLayout(self)
        vbox.addWidget(self.listView)
        vbox.addLayout(hbox)

    def get_selected_indexes(self) -> list:
        self.refresh_data()
        items = list(self.data.items())
        return [i for i in range(len(items)) if items[i][1]]

    def get_selected(self) -> list:
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
        self.data = {
            self.model.item(i).text(): (True if self.model.item(i).checkState() == QtCore.Qt.Checked else False)
            for i in range(self.model.rowCount())}

    def deleteSelected(self, function=None):
        indexes = self.get_selected_indexes()
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
        """Sets widgets and layouts."""
        central = QtWidgets.QWidget()
        vbox = QtWidgets.QVBoxLayout()
        vbox2 = QtWidgets.QVBoxLayout()
        hbox = QtWidgets.QHBoxLayout()
        send_button = QtWidgets.QPushButton("Send")
        self.receiver = QtWidgets.QComboBox()
        contacts_button = QtWidgets.QPushButton("Refresh contacts")
        self.subject = QtWidgets.QLineEdit()
        self.body_type = QtWidgets.QComboBox()
        self.body = QtWidgets.QTextEdit()
        add_button = QtWidgets.QPushButton("Add file")
        delete_button = QtWidgets.QPushButton("Remove")
        rename_button = QtWidgets.QPushButton("Rename")
        select_button = QtWidgets.QPushButton("Select all")
        unselect_button = QtWidgets.QPushButton("Unselect all")
        grid = QtWidgets.QGridLayout()
        self.fileChooseWidget = ChecklistWidget()

        grid.addWidget(add_button, 0, 0)
        grid.addWidget(rename_button, 0, 1)
        grid.addWidget(delete_button, 0, 2)
        grid.addWidget(select_button, 1, 0)
        grid.addWidget(unselect_button, 1, 1)
        vbox2.addWidget(QtWidgets.QLabel("Attachments"))
        vbox2.addWidget(self.fileChooseWidget)
        vbox2.addLayout(grid)
        vbox.addWidget(QtWidgets.QLabel("Receiver"))
        vbox.addWidget(self.receiver)
        vbox.addWidget(contacts_button)
        vbox.addWidget(QtWidgets.QLabel("Subject"))
        vbox.addWidget(self.subject)
        vbox.addWidget(QtWidgets.QLabel("Text type"))
        vbox.addWidget(self.body_type)
        vbox.addWidget(QtWidgets.QLabel("Email text"))
        vbox.addWidget(self.body)
        vbox.addWidget(send_button)
        hbox.addLayout(vbox)
        hbox.addLayout(vbox2)
        central.setLayout(hbox)
        self.setCentralWidget(central)

        self.body_type.addItems(["Plain", "HTML"])

        contacts_button.clicked.connect(self.refreshContacts)
        delete_button.clicked.connect(self.deleteSelected)
        add_button.clicked.connect(self.selectFile)
        rename_button.clicked.connect(self.rename)
        send_button.clicked.connect(self.sendEmail)
        select_button.clicked.connect(self.fileChooseWidget.select)
        unselect_button.clicked.connect(self.fileChooseWidget.unselect)

    def rename(self):
        default = self.fileChooseWidget.get_selected()
        if default: default = default[0]
        text, ok = QtWidgets.QInputDialog.getText(self, "Rename file", "Enter a new filename with format", text=default)
        if ok:
            del self.fileChooseWidget.data[default]
            key = self.search_key(default)
            self.files[key] = text
            self.fileChooseWidget.data[text] = True
            self.fileChooseWidget.refresh()

    def refreshContacts(self):
        self.receiver.clear()
        self.receiver.addItems(self.parent.contacts)

    def initContacts(self):
        self.receiver.addItems(list(self.parent.contacts.keys()))

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

    def search_key(self, text):
        for i in self.files:
            if self.files[i] == text:
                return i
        return None

    def sendEmail(self):
        body_type = self.body_type.currentText()
        if body_type == "HTML":
            body = self.body.toHtml()
        else:
            body = self.body.toPlainText()
        subject = self.subject.text()
        files = self.files
        receiver = self.receiver.currentText()
        message = email.mime.multipart.MIMEMultipart("alternative")
        message["Subject"] = subject
        message["To"] = self.parent.contacts[receiver]
        message["From"] = self.parent.USERNAME
        if body_type == "HTML":
            body = email.mime.text.MIMEText(body, "html")
        else:
            body = email.mime.text.MIMEText(body, "plain")
        message.attach(body)
        for file in self.files:
            try:
                with open(file, "rb") as f:
                    content = f.read()
                attach_file = MIMEApplication(content)
                attach_file.add_header('Content-Disposition', 'attachment', filename=self.files[file])
                message.attach(attach_file)
            except Exception as e:
                print(e)
        dialog = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Question, "Confirm sending", "Are you sure you want to send an email?",
            buttons=QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, parent=self)
        dialog.setDefaultButton(QtWidgets.QMessageBox.Yes)
        if dialog.exec_() == QtWidgets.QMessageBox.Yes:
            self.parent.smtpSession.sendmail(self.parent.USERNAME, self.parent.contacts[receiver], message.as_string())


class AddContactDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.setup()
        self.result = False
        self.setWindowTitle("New contact")

    def setup(self):
        """Sets widgets and layouts."""
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
        central = QtWidgets.QWidget()
        vbox = QtWidgets.QVBoxLayout()
        self.contacts_list = ChecklistWidget()
        grid = QtWidgets.QGridLayout()
        delete_button = QtWidgets.QPushButton("Delete")
        add_button = QtWidgets.QPushButton("Add")
        import_button = QtWidgets.QPushButton("Import")
        export_button = QtWidgets.QPushButton("Export")

        grid.addWidget(add_button, 0, 0)
        grid.addWidget(delete_button, 0, 1)
        grid.addWidget(import_button, 1, 0)
        grid.addWidget(export_button, 1, 1)
        vbox.addWidget(self.contacts_list)
        vbox.addLayout(grid)
        central.setLayout(vbox)
        self.setCentralWidget(central)

        add_button.clicked.connect(self.addContact)
        delete_button.clicked.connect(self.deleteSelected)

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
                pyAesCrypt.decryptFile('.ideamail.data.encr', '.ideamail.data', config['IdeaMail']['Password'],
                                       1024 * 64)
                os.remove('.ideamail.data.encr')
                with open('.ideamail.data', 'br') as f:
                    data = pickle.load(f)
                self.USERNAME = data.get("USERNAME")
                self.PASSWORD = data.get("PASSWORD")
                self.MAIL_HOST = data.get("HOST")
                self.data = data.get("DATA")
                self.contacts = data.get("CONTACTS")
                self.serverLogin(self.USERNAME, self.PASSWORD, self.MAIL_HOST)
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

    def serverLogin(self, username, password, mail_host):
        self.smtpSession = smtplib.SMTP('smtp.{}.com'.format(mail_host.lower()), 587)
        self.smtpSession.starttls()
        self.imapSession = imapclient.IMAPClient('imap.{}.com'.format(mail_host.lower()), ssl=True)
        try:
            self.smtpSession.login(self.USERNAME, self.PASSWORD)
            self.imapSession.login(self.USERNAME, self.PASSWORD)
        except:
            error_window = QtWidgets.QMessageBox.critical(self, "Login error",
                                                         "Check that login and password are valid.")
            self.showLoginDialog()

    def showLoginDialog(self):
        loginDialog = LoginWindow()
        loginDialog.show()
        out = loginDialog.exec_()
        if all(out[0:1]):
            self.USERNAME, self.PASSWORD, self.MAIL_HOST = out
            self.serverLogin(self.USERNAME, self.PASSWORD, self.MAIL_HOST)
        else:
            self.deleteLater()

    def createConfig(self):
        password = ''.join(secrets.choice(ascii_letters + punctuation + digits) for i in range(20))
        data = {'USERNAME': self.USERNAME,
                'PASSWORD': self.PASSWORD,
                "HOST": self.MAIL_HOST,
                'DATA': self.data,
                "CONTACTS": self.contacts}
        config = configparser.RawConfigParser()
        config['IdeaMail'] = {"Password": password}
        with open('.ideamail.ini', 'w') as f:
            config.write(f)
        with open(".ideamail.data", "bw") as f:
            pickle.dump(data, f)
        pyAesCrypt.encryptFile('.ideamail.data', '.ideamail.data.encr', password, 1024 * 64)
        os.remove('.ideamail.data')

    def createMainLayout(self):
        # TODO: replace QListView with my checkable ListView widget
        hbox = QtWidgets.QHBoxLayout()
        listview = QtWidgets.QListView()
        vbox = QtWidgets.QVBoxLayout()
        centralwidget = QtWidgets.QWidget()
        self.model = QtGui.QStandardItemModel()  # ItemModel for QListView
        toolbar = self.addToolBar("Main")
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('Mail')
        logoutAction = QtWidgets.QAction(QtGui.QIcon("icons/logout.png"), 'Log out', self)
        exitAction = QtWidgets.QAction(QtGui.QIcon("icons/close.png"), 'Exit', self)
        folderSelectAction = QtWidgets.QAction(QtGui.QIcon("icons/folder.png"), "Folder selecting", self)
        newMessageAction = QtWidgets.QAction(QtGui.QIcon("icons/new_message.png"), "New email", self)
        refreshAction = QtWidgets.QAction(QtGui.QIcon("icons/refresh.png"), "Refresh inbox", self)
        contactsAction = QtWidgets.QAction(QtGui.QIcon("icons/contacts.png"), "Open contacts", self)

        listview.setModel(self.model)
        listview.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)  # Text in listview can't be edited
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
        except:
            pass
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

            message_uid, data = list(self.imapSession.fetch([i], ['ENVELOPE']).items())[0]
            envelope = data[b'ENVELOPE']
            self.data[i] = envelope
            subject = envelope.subject
            self.addSubjectToModel(subject)

    def closeEvent(self, event):
        reply = QtWidgets.QMessageBox.question(self, "Exit", "Are you sure you want to exit?",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            try:
                self.createConfig()
                self.imapSession.logout()
                self.smtpSession.close()
            except:
                pass
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
