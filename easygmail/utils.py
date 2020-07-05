import re


def removeQuotedParts(emailText):
    """Returns the text in ``emailText`` up to the quoted "reply" text that begins with
    "On Sun, Jan 1, 2018 at 12:00 PM al@inventwithpython.com wrote:" part."""
    replyPattern = re.compile(
        r"On (Sun|Mon|Tue|Wed|Thu|Fri|Sat), (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d+, \d\d\d\d at \d+:\d+ (AM|PM) (.*?) wrote:"
    )

    mo = replyPattern.search(emailText)
    if mo is None:
        return emailText
    else:
        return emailText[: mo.start()]


def parse_for_encoding(value):
    """
    Helper function called by Message's __init__
    Try to get an encoding in headers. If it's
    not found, set encoding to UTF-8 and hope for
    the best.
    """
    mo = re.search('charset="(.*?)"', value)
    if mo is None:
        emailEncoding = "UTF-8"
    else:
        emailEncoding = mo.group(1)
    return emailEncoding


def get_message_size(b: int) -> str:
    """
    Converts message size from bytes to
    kilobytes, megabytes, etc.

    Args:
        b - number of bytes

    Returns:
        (str) - converted size of a message
    """
    sizes = ["bytes", "Kb", "Mb", "Gb"]
    counter = 0
    while b > 1024 and counter < 3:
        b /= 1024
        counter += 1

    return str(round(b, 1)) + ' ' + sizes[counter]
