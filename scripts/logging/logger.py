"""Importing logging"""
import logging


def getLogger():
    __logger__ = logging.getLogger("")
    __logger__.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s - %(message)s')

    file_handler = logging.FileHandler('logs/pythonProject9_22052023.log')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    __logger__.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)
    __logger__.addHandler(console_handler)

    return __logger__


logger = getLogger()
