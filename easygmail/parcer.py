from . import utils
import re
import base64


sender_regex = re.compile(r"(.*)\<(.*)\>.*")


class Parcer(object):
    """
    Class for parcing information from ``get()`` API call
    and processing it to ``Message`` class.

    Attributes:
        message - Message class
        msg - get() API call result
    """
    def __init__(self, msg, parse=True):
        self.message = msg
        self.msg = msg.message
        if parse:
            self.headers()
            self.parts()

    def __set(self, attr, value):
        """Perform setattr() on Message class."""
        setattr(self.message, attr, value)

    def headers(self):
        """Get headers of an email."""
        for header in self.msg["payload"]["headers"]:
            if header["name"].upper() == "FROM":
                sender = header["value"]
                result = re.findall(sender_regex, sender)
                if result:
                    self.message.sender = result[0][0] if result[0][0] else sender
                    self.message.sender_email = result[0][1] if result[0][1] else None
                else:
                    self.message.sender = sender
                    self.message.sender_email = None
            if header["name"].upper() == "TO":
                self.__set("recipient", header["value"])
            if header["name"].upper() == "SUBJECT":
                self.__set("subject", header["value"])
            if header["name"].upper() == "CONTENT-TYPE":
                self.encoding = utils.parse_for_encoding(header["value"])

    def get_charset_from_part_headers(self, part) -> str:
        if part.get("headers"):
            for header in part["headers"]:
                if header["name"].upper() == "CONTENT-TYPE":
                    return utils.parse_for_encoding(header["value"])
        return "UTF-8"

    def parts(self):
        if "parts" in self.msg["payload"].keys():
            for part in self.msg["payload"]["parts"]:

                if part["mimeType"].upper() == "TEXT/PLAIN" and "data" in part["body"]:
                    self.encoding = self.get_charset_from_part_headers(part)
                    self.message.body = base64.urlsafe_b64decode(
                        part["body"]["data"]).decode(self.encoding)

                if part["mimeType"].upper() == "TEXT/HTML" and "data" in part["body"]:
                    encoding = self.get_charset_from_part_headers(part)
                    self.message.html = base64.urlsafe_b64decode(
                        part["body"]["data"]).decode(encoding)

                if part["mimeType"].upper() == "MULTIPART/ALTERNATIVE":
                    for multipartPart in part["parts"]:
                        if multipartPart["mimeType"].upper() == "TEXT/PLAIN" and "data" in multipartPart["body"]:
                            for header in multipartPart["headers"]:
                                if header["name"].upper() == "CONTENT-TYPE":
                                    self.encoding = utils.parse_for_encoding(header["value"])
                            self.__set("body", base64.urlsafe_b64decode(
                                multipartPart["body"]["data"]).decode(self.encoding))

                if "filename" in part.keys() and part["filename"] != "":
                    self.message.attachments.append(
                        {
                            "filename": part["filename"],
                            "id": part["body"]["attachmentId"],
                            "size": part["body"]["size"]
                        }
                    )

        # No parts in payload but payload.body exists
        elif "body" in self.msg["payload"].keys():
            content = base64.urlsafe_b64decode(
                self.msg["payload"]["body"]["data"]).decode(self.encoding)
            if self.msg["payload"]['mimeType'].upper() == "TEXT/PLAIN":
                self.__set("body", content)
            elif self.msg["payload"]['mimeType'].upper() == "TEXT/HTML":
                self.__set("html", content)
