import os
import json


def get_size(b: int) -> str:
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


def copy_data_into_file(message):
    data = {}
    for attr in ["subject", "timestamp", "sender", "snippet"]:
        data.update({attr: getattr(message, attr)})
    with open(os.path.join('data', message.id), "w") as f:
        json.dump(data, f)


def get_data_from_file(uid: str):
    with open(os.path.join('data', uid), 'r') as f:
        return json.load(f)
