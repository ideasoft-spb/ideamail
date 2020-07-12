import os
import pickle
import datetime
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from . import utils
from . import parcer
from .exceptions import GmailException


SCOPES = ['https://mail.google.com/']
SERVICE = None
USER = None


class Gmail(object):
    def __init__(self):
        self.build_service()
        self.labels: dict = self.get_all_labels()

    def __get_credentials(self):
        creds = None
        # If token.pickle exists, load the token.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # Otherwise, log in the user.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return creds

    def build_service(self) -> None:
        global SERVICE, USER
        SERVICE = build('gmail', 'v1', credentials=self.__get_credentials())
        USER = SERVICE.users().getProfile(userId='me')

    def get_all_labels(self) -> dict:
        """Return dictionary of all labels and their ids."""
        response = SERVICE.users().labels().list(userId='me').execute()
        return {i['name']: i['id'] for i in response['labels']}

    def add_labels(self, ids: list, labels: list, target: str = "messages") -> None:
        """
        Add labels to a list of messages.
        """
        data = {"removeLabelIds": [], "addLabelIds": labels}
        for i in ids:
            if target == "threads":
                SERVICE.users().threads().modify(userId='me', id=i, body=data).execute()
            elif target == "messages":
                SERVICE.users().messages().modify(userId='me', id=i, body=data).execute()

    def remove_labels(self, ids: list, labels: list, target: str = "messages") -> None:
        """
        Add labels to a list of messages.
        """
        data = {"removeLabelIds": labels, "addLabelIds": []}
        for i in ids:
            if target == "threads":
                SERVICE.users().threads().modify(userId='me', id=i, body=data).execute()
            elif target == "messages":
                SERVICE.users().messages().modify(userId='me', id=i, body=data).execute()


class Message(object):
    """
    A Gmail message object returned by get() API call.

    Attributes:
        service - gmail service object
        message - message object
        parcer - parcer class

        id - id of the message
        thead_id - id of a thread the message belongs to
        labels - list of labels set to the message

        subject - subject of the message
        snippet - snippet of the message
        timestamp - a time when the message was sent
        sender - a sender of the message
        recipient - a recipient of the message

        body - body of a message

        attachments - list of attachments of the message

    """
    def __init__(self, id: str = None, message: dict = None):
        """
        Initialize a new message. In order to do it you can pass
        id or message itself. If id is passed, API get() call is
        performed. Otherwise, message is parsed immediately.
        """
        if id:
            self.message: dict = SERVICE.users().messages().get(userId='me', id=id).execute()
        elif message:
            self.message: dict = message
        else:
            raise GmailException("Provide at least one argument: id or message")
        self.id = self.message['id']
        self.thread = self.message['threadId']
        self.snippet = self.message["snippet"]
        self.historyId = self.message["historyId"]
        self.labels = self.message['labelIds']
        self.datetime = datetime.datetime.fromtimestamp(int(self.message["internalDate"]) // 1000)
        self.size = utils.get_message_size(self.message["sizeEstimate"])
        self.body = None
        self.html = None
        self.subject = None
        self.sender = None
        self.sender_email = None
        self.recipient = None
        self.attachments = []
        self.parcer = parcer.Parcer(self)

    def __repr__(self):
        return "Message from=%r to=%r timestamp=%r subject=%r snippet=%r" % (
            self.sender,
            self.recipient,
            self.timestamp,
            self.subject,
            self.snippet,
        )

    @property
    def date(self, fmt=None):
        return str(self.datetime)

    def get_date(self, fmt):
        return self.datetime.strftime(fmt)

    def __str__(self):
        return self.__repr__()

    @property
    def filenames(self):
        """Return list of attachments' filenames."""
        return [i['filename'] for i in self.attachments]

    def get_files(self):
        return [
            Attachment(i, self) for i in self.attachments
        ]


class Attachment(object):
    def __init__(self, info, message):
        self.message_id = message.id
        self.id = info['id']
        self.size = utils.get_message_size(info['size'])
        self.filename = info['filename']

    def download(self, folder: str, overwrite: bool = False) -> int:
        """
        Downloads the attachment.

        Args:
            folder: Download folder
            overwrite: Shows if file with the same name needs to be overwritten.
        Returns:

            0 - operation was completed successfully
            1 - specified folder is a file
            2 - file with that name already exists
        """
        if not overwrite and os.path.isfile(os.path.join(folder, self.filename)):
            return 2
        if os.path.isfile(folder):
            return 1
        if not os.path.exists(folder):
            os.mkdir(folder)
        attachment = SERVICE.users().messages().attachments().get(
            id=self.id, messageId=self.message_id, userId='me').execute()
        attachment = base64.urlsafe_b64decode(attachment["data"])
        with open(os.path.join(folder, self.filename), "wb") as f:
            f.write(attachment)
        return 0

    def __str__(self):
        return self.filename


class Search(object):
    """
    Class for searching messages in mailbox.

    Attributes:

    """
    page = None  # Number of current page
    results = None  # Results of the search
    __kwargs = {}  # API call arguments
    __page_tokens = {}  # Dictionary with tokens

    def __init__(self, target: str = "messages", max_results=25, labels=None, query=None):
        if target not in ["messages", "threads"]:
            raise GmailException("Specify a valid \"target\" attribute.")
        self.__target = target
        self.__kwargs["maxResults"] = max_results
        if labels:
            self.__kwargs["labelIds"] = labels
        if query:
            self.__kwargs["q"] = query
        self.__search()

    def __search(self) -> None:
        """
        Performes the initial search (the first page of search results).
        """
        self.page = 1
        response = self._get_response()
        self.__page_tokens.update({
            self.page + 1: response.get("nextPageToken")
        })
        self.results = response.get(self.__target)

    def _get_response(self) -> dict:
        """Perform API list() call"""
        if self.__target == "messages":
            return SERVICE.users().messages().list(
                userId='me', **self.__kwargs).execute()
        return SERVICE.users().threads().list(
            userId='me', **self.__kwargs).execute()

    @property
    def has_next_page(self):
        return self.page + 1 in self.__page_tokens and self.__page_tokens[self.page + 1]

    def next_page(self) -> None:
        """
        Gets the next page of search and saves it in ``self.results``.

        Raises:
            GmailException - next page doen't exist
        """
        if self.page + 1 not in self.__page_tokens or not self.__page_tokens[self.page + 1]:
            raise GmailException("Next page of search results doesn't exist.")
        self.__kwargs["pageToken"] = self.__page_tokens[self.page + 1]
        response = self._get_response()
        self.page += 1
        self.__page_tokens.update({
            self.page + 1: response.get("nextPageToken")
        })
        self.results = response[self.__target]

    def prev_page(self) -> None:
        if self.page == 1:
            raise GmailException("Previous page of search doesn't exist")
        if self.page != 2:
            self.__kwargs["pageToken"] = self.__page_tokens.get(self.page - 1)
        else:
            if "pageToken" in self.__kwargs:
                del self.__kwargs["pageToken"]
        response = self._get_response()
        self.page -= 1
        self.results = response[self.__target]

    def full_results(self) -> list:
        """
        Gets list of messages/threads on the current page.
        """
        if self.__target == "messages":
            return [
                Message(self.service, id=i['id']) for i in self.results
            ]
        return [
            threads.Thread(self.service, i) for i in self.results
        ]


class Thread(object):
    """
    Represents a thread of Gmail messages. These objects are returned by
    the threads().get() API call. They contain references to a list
    of GmailMessage objects.
    """

    def __init__(self, service, data):
        self.service = service
        self.id = data['id']
        self.snippet = data['snippet']
        self._messages = None

    def get_messages(self):
        self._messages = [
            Message(
                self.service,
                message=i
            ) for i in self.service.users().threads().get(
                userId='me', id=self.id).execute()["messages"]
        ]
        return self._messages
