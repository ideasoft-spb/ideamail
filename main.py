import sys
import os
import webbrowser
import json
from PyQt5 import QtWidgets, QtGui, QtCore, QtWebEngineWidgets
import easygmail
import utils
from typing import Dict
import checklistwidget


SERVICE = None
PER_QUERY: int = 35


class AttachmentsWidget(QtWidgets.QWidget):
    def __init__(self, message: easygmail.Message = None, parent=None):
        super().__init__(parent=parent)
        self.setup()
        self.files = message.get_files()
        rows = [[i.filename, i.size] for i in message.get_files()]
        self.setFields(["Filename", "Size"])
        self.addNewItems(rows)

    def setup(self):
        self.tableview = QtWidgets.QTableView()
        self.model = QtGui.QStandardItemModel()
        self.tableview.setModel(self.model)
        self.selection_model = QtCore.QItemSelectionModel(self.model)
        self.tableview.setSelectionModel(self.selection_model)
        self.tableview.verticalHeader().hide()
        self.tableview.setSelectionBehavior(self.tableview.SelectRows)
        self.tableview.setSelectionMode(self.tableview.MultiSelection)
        self.tableview.setShowGrid(False)
        self.tableview.setEditTriggers(self.tableview.NoEditTriggers)
        self.tableview.doubleClicked.connect(self.rename)
        buttons = QtWidgets.QGridLayout()
        download_btn = QtWidgets.QPushButton("Download")
        select_btn = QtWidgets.QPushButton("Select all")
        unselect_btn = QtWidgets.QPushButton("Unselect all")
        select_btn.clicked.connect(self.tableview.selectAll)
        unselect_btn.clicked.connect(self.tableview.clearSelection)
        download_btn.clicked.connect(self.download)
        buttons.addWidget(download_btn, 0, 0, 1, 2)
        buttons.addWidget(select_btn, 1, 0)
        buttons.addWidget(unselect_btn, 1, 1)
        vbox = QtWidgets.QVBoxLayout()
        vbox.addWidget(self.tableview)
        vbox.addWidget(QtWidgets.QLabel("<i>Click to select several attachments</i>"))
        vbox.addLayout(buttons)
        self.setLayout(vbox)

    def rename(self, index):
        old = self.files[index.row()]
        new = QtWidgets.QInputDialog.getText(self, "Rename an attachment", "Enter a new filename")
        if new[1]:
            self.model.setItem(index.row(), 0, QtGui.QStandardItem(new[0]))
            old.filename = new[0]

    def download(self):
        """
        Downloads selected attachments and shows a messagebox with list of
        files that have been downloaded.
        """
        to_download = [self.files[i.row()] for i in self.selection_model.selectedRows()]
        if not to_download:
            return
        downloaded = []
        folder = QtWidgets.QFileDialog.getExistingDirectory(parent=self)
        if folder:
            for f in to_download:
                status = f.download(folder=folder)
                filename = os.path.join(folder, f.filename)
                downloaded.append(filename)
                if status == 3:
                    QtWidgets.QMessageBox.information(
                        self, "File already exists",
                        "File %s already exists. Rename the attachment file." % filename,
                        buttons=QtWidgets.QMessageBox.Close)
                    break
            QtWidgets.QMessageBox.information(
                self, "Downloaded files",
                "The following files have been downloaded:\n{}".format('\n'.join(downloaded)),
                buttons=QtWidgets.QMessageBox.Ok
            )

    def addNewItem(self, item: list, checked=False):
        if len(item) != len(self.fields):
            raise ValueError("Item's columns amount is not equal table fields amount.")
        fields = [QtGui.QStandardItem(i) for i in item]
        self.model.appendRow(fields)
        self.tableview.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.tableview.resizeColumnsToContents()

    def addNewItems(self, items: list, checked=False):
        for i in items:
            self.addNewItem(i, checked=checked)

    def setFields(self, fields: list):
        self.model.setHorizontalHeaderLabels(fields)
        self.fields = fields


class SearchWindow(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setup()

    def setup(self):
        layout = QtWidgets.QFormLayout()
        self.textField = QtWidgets.QLineEdit()
        self.labelBox = checklistwidget.CheckListWidget()
        layout.addRow("Query text", self.textField)
        layout.addRow("Labels", self.labelBox)
        self.labelBox.addNewElements(self.parent().gmail.labels)
        buttons = QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        buttonbox = QtWidgets.QDialogButtonBox(buttons)
        buttonbox.accepted.connect(self.accept)
        buttonbox.rejected.connect(self.reject)
        layout.addWidget(buttonbox)
        self.setLayout(layout)
        self.setWindowTitle("Search")

    def accept(self):
        self.parent().search_kwargs = {
            "query": self.textField.text(),
            "labels": self.labelBox.get_selected()
        }
        super().accept()


class AttachmentsWindow(QtWidgets.QMainWindow):
    def __init__(self, message, parent=None):
        super().__init__(parent=parent)
        self.message = message
        self.setup()

    def setup(self):
        central_widget = AttachmentsWidget(parent=self, message=self.message)
        self.setCentralWidget(central_widget)


class MessageWidget(QtWidgets.QWidget):
    def __init__(self, uid, parent=None):
        super().__init__(parent=parent)
        self.message = easygmail.Message(id=uid)
        self.setup()

    def setup(self):
        vbox = QtWidgets.QVBoxLayout()
        info = QtWidgets.QFormLayout()
        info.addRow("Subject:", QtWidgets.QLabel(self.message.subject))
        sender = QtWidgets.QLabel(self.message.sender)
        if self.message.sender_email:
            sender.setText(self.message.sender + " (<i>" + self.message.sender_email + "</i>)")
        info.addRow("From:", sender)
        if self.message.attachments:
            attachments = QtWidgets.QListWidget()
            attachments.addItems(self.message.filenames)
            open_btn = QtWidgets.QPushButton("Open attachments")
            open_btn.clicked.connect(self.open_files)
            info.addRow("Attachments:", attachments)
            info.addWidget(open_btn)
        if self.message.html:
            self.wv = QtWebEngineWidgets.QWebEngineView(self)
            self.wv.setHtml(self.message.html)
            self.wv.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
            QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Q"), self, self.close)
            QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+-"), self, self.zoom_minus)
            QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+="), self, self.zoom_plus)
        if self.message.body:
            self.bodyView = QtWidgets.QTextEdit()
            self.bodyView.setText(self.message.body)
            self.bodyView.setReadOnly(True)
        vbox.addLayout(info)
        if self.message.html:
            vbox.addWidget(QtWidgets.QLabel("<i>Email HTML content:</i>"))
            vbox.addWidget(self.wv)
        if self.message.body:
            vbox.addWidget(QtWidgets.QLabel("Email body:"))
            vbox.addWidget(self.bodyView)
        self.setLayout(vbox)

    def open_files(self):
        """Open a new attachments window."""
        window = AttachmentsWindow(self.message, parent=self)
        window.show()

    @QtCore.pyqtSlot()
    def zoom_minus(self):
        """Zoom out ``self.wv``"""
        self.wv.setZoomFactor(self.wv.zoomFactor() - 0.1)

    @QtCore.pyqtSlot()
    def zoom_plus(self):
        """Zoom in ``self.wv``"""
        self.wv.setZoomFactor(self.wv.zoomFactor() + 0.1)


class MessageWindow(QtWidgets.QMainWindow):
    def __init__(self, uid, parent=None):
        super().__init__(parent=parent)
        self.id = uid
        self.setup()

    def setup(self):
        self.msg_widget = MessageWidget(self.id, parent=self)
        self.msg_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        scrollarea = QtWidgets.QScrollArea()
        scrollarea.setWidget(self.msg_widget)
        self.setCentralWidget(scrollarea)
        scrollarea.setWidgetResizable(True)
        self.setWindowTitle("Message " + self.id)
        self.setGeometry(200, 200, 800, 600)


class MainWindow(QtWidgets.QMainWindow):
    """
    Main windows of this application.
    """
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.search = None
        self.login()
        self.data: Dict[str, dict] = {}
        self.base: Dict[str, dict] = {}
        if os.path.isfile('data.json'):
            self.load_data()
        else:
            self.save_data()
        self.setup()

    def load_data(self):
        with open("data.json", "r") as f:
            self.base = json.load(f)

    def save_data(self):
        with open("data.json", "w") as f:
            json.dump(self.base, f, indent=4)

    def add_message_data(self, uid: int):
        message = easygmail.Message(id=uid)
        keys = ['subject', 'sender', 'sender_email', 'date', 'labels']
        d = {i: getattr(message, i) for i in keys}
        self.data[uid] = d
        self.base[uid] = d

    def setup(self):
        """
        Initializes layouts & widgets for this window.
        """
        self.setWindowTitle("IdeaMail v2.0")
        exit_action = QtWidgets.QAction('Exit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.setStatusTip('Exit the application')
        exit_action.triggered.connect(self.close)
        login_action = QtWidgets.QAction('Login', self)
        login_action.setShortcut('Ctrl+L')
        login_action.setStatusTip('Login into the account')
        login_action.triggered.connect(self.login)
        logout_action = QtWidgets.QAction('Logout', self)
        logout_action.setShortcut('Ctrl+Shift+L')
        logout_action.setStatusTip('Logout from the account')
        logout_action.triggered.connect(self.logout)
        refresh_action = QtWidgets.QAction(QtGui.QIcon('icons/refresh.png'), 'Refresh', self)
        refresh_action.setShortcut('Ctrl+R')
        refresh_action.triggered.connect(self.refresh)
        vk_action = QtWidgets.QAction('VK page', self)
        vk_action.setStatusTip('Open our VK page')
        vk_action.triggered.connect(lambda x: self.open_page("vk.com/ideasoft_spb"))
        self.clear_data_action = QtWidgets.QAction("Clear data ({})".format(
            utils.get_size(os.path.getsize('data.json'))
        ), self)
        self.clear_data_action.triggered.connect(self.clear_data)

        self.statusbar: QtWidgets.QStatusBar = self.statusBar()

        appmenu = QtWidgets.QMenu("App", self)
        appmenu.addActions([
            login_action, exit_action,
            refresh_action, self.clear_data_action,
            logout_action
        ])
        helpmenu = QtWidgets.QMenu("Help", self)
        helpmenu.addAction(vk_action)
        menubar = self.menuBar()
        menubar.addMenu(appmenu)
        menubar.addMenu(helpmenu)
        toolbar: QtWidgets.QToolBar = self.addToolBar("Exit")
        toolbar.addAction(refresh_action)
        toolbar.setFloatable(False)
        toolbar.setAllowedAreas(QtCore.Qt.LeftToolBarArea | QtCore.Qt.TopToolBarArea)

        vbox = QtWidgets.QVBoxLayout()
        buttons = QtWidgets.QGridLayout()
        delete_btn = QtWidgets.QPushButton("Move to trash")
        load_more_btn = QtWidgets.QPushButton("More results")
        mark_as_read_btn = QtWidgets.QPushButton("Marks as read")
        star_btn = QtWidgets.QPushButton("Star")
        load_more_btn.clicked.connect(self.load_more)
        mark_as_read_btn.clicked.connect(lambda: self.remove_label("UNREAD"))
        star_btn.clicked.connect(lambda: self.add_label('STARRED'))

        # Search widget
        self.search_widget = QtWidgets.QPushButton("Open search")
        self.search_widget.clicked.connect(self.open_search)
        self.search_widget.setShortcut("Ctrl+F")

        # QTableView
        self.tableview = QtWidgets.QTableView()
        self.model = QtGui.QStandardItemModel()
        self.tableview.setModel(self.model)
        self.tableview.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)
        self.tableview.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.tableview.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeToContents)
        self.tableview.verticalHeader().hide()
        self.tableview.doubleClicked.connect(self.row_clicked)

        self.main_widget = QtWidgets.QWidget()

        vbox.addWidget(self.search_widget)
        vbox.addWidget(self.tableview)
        vbox.addLayout(buttons)
        buttons.addWidget(delete_btn, 0, 0)
        buttons.addWidget(load_more_btn, 0, 1)
        buttons.addWidget(mark_as_read_btn, 0, 2)
        buttons.addWidget(star_btn, 1, 0)
        self.main_widget.setLayout(vbox)
        self.setCentralWidget(self.main_widget)

    def logout(self):
        if os.path.isfile("token.pickle"):
            os.remove('token.pickle')
        exit()

    def add_label(self, label: str) -> None:
        rows = self.tableview.selectionModel().selectedRows()
        indexes = sorted([i.row() for i in rows])
        keys = list(self.data.keys())
        ids = [keys[i] for i in indexes]
        self.gmail.add_labels(ids, [label])
        for i in ids:
            if label in self.data[i]['labels']:
                self.data[i]['labels'].append(label)
        self.refresh_table()

    def remove_label(self, label: str) -> None:
        rows = self.tableview.selectionModel().selectedRows()
        indexes = sorted([i.row() for i in rows])
        keys = list(self.data.keys())
        ids = [keys[i] for i in indexes]
        self.gmail.remove_labels(ids, [label])
        for i in ids:
            if label in self.data[i]['labels']:
                self.data[i]['labels'].remove(label)
        self.refresh_table()

    def refresh(self):
        self.model.clear()
        self.add_mail()

    def open_search(self):
        """Open a search window."""
        window = SearchWindow(parent=self)
        window.exec_()
        if window.result() == 1:
            self.search = easygmail.Search(max_results=PER_QUERY, **self.search_kwargs)
            self.model.clear()
            self.add_search_data()

    def load_more(self):
        if self.search.has_next_page:
            self.search.next_page()
            self.add_search_data(cleardata=False)
        else:
            QtWidgets.QMessageBox.information(
                self, "Search results", "No search results were found.")

    def add_mail(self, **kwargs):
        self.search = easygmail.Search(max_results=PER_QUERY, **kwargs)
        self.add_search_data()

    def add_search_data(self, cleardata=True):
        if cleardata:
            self.data = {}
        r = self.search.results
        if not r:
            QtWidgets.QMessageBox.information(
                self, "Search results", "No search results were found.")
            return
        indicator = QtWidgets.QProgressDialog(
            "<i>Getting your messages...</b>", "Stop", 0, len(r)-1, parent=self)
        indicator.setWindowModality(QtCore.Qt.WindowModal)
        indicator.setWindowTitle("Getting messages...")
        indicator.show()
        self.model.setHorizontalHeaderLabels([
            "Subject", "Sender", "Email", "Date", "Starred"
        ])
        for i in r:
            uid = i['id']
            if uid not in self.data:
                if uid in self.base:
                    self.data[uid] = self.base[uid]
                else:
                    self.add_message_data(uid)
            self.add_message_to_table(self.data[uid])
            indicator.setValue(indicator.value() + 1)
            loop = QtCore.QEventLoop()
            QtCore.QTimer.singleShot(10, loop.quit)
            loop.exec_()
            if indicator.wasCanceled():
                indicator.setValue(len(r)-1)
                break
        self.save_data()
        self.refresh_data_size()

    def refresh_table(self, clear=True, data=None) -> None:
        """
        Refresh self.tableview.
        If ``clear`` is ``True``, clear a model.
        """
        if data is None:
            data = self.data
        if clear:
            self.model.clear()
        self.model.setHorizontalHeaderLabels(["Subject", "Sender", "Email", "Date", "Starred"])
        for i in data:
            self.add_message_to_table(data[i])

    def add_message_to_table(self, message: dict) -> None:
        keys = ['subject', 'sender', 'sender_email', 'date']
        values = [message.get(key) for key in keys]
        row = [QtGui.QStandardItem(value) for value in values]
        if 'UNREAD' in message['labels']:
            for i in row:
                i.setBackground(QtGui.QBrush(QtGui.QColor("#EBDCCB")))
        if 'STARRED' in message['labels']:
            row.append(QtGui.QStandardItem("\U0001F31F"))
            row[-1].setTextAlignment(QtCore.Qt.AlignCenter)
        self.model.appendRow(row)

    def row_clicked(self, index: QtCore.QModelIndex):
        window = MessageWindow(list(self.data.keys())[index.row()], parent=self)
        window.show()

    def clear_data(self):
        os.remove('data.json')
        self.base = {}
        self.save_data()
        self.refresh_data_size()

    def refresh_data_size(self):
        self.clear_data_action.setText(
            "Clear data ({})".format(
                utils.get_size(os.path.getsize('data.json')))
            )

    def closeEvent(self, event):
        self.save_data()
        event.accept()

    def open_page(self, url):
        webbrowser.open("https://" + url)

    def login(self):
        self.gmail = easygmail.Gmail()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.resize(400, 400)
    window.setWindowIcon(QtGui.QIcon("icons/logo.png"))
    window.showMaximized()
    window.show()
    sys.exit(app.exec_())
