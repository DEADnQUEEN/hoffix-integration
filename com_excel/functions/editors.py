import datetime
from typing import Callable, Optional


def convert_to_date(date_format: str) -> Callable[[str], Optional[datetime.date]]:
    def func(value: any) -> Optional[datetime.date]:
        if not isinstance(value, str) or not value:
            return None
        return datetime.datetime.strptime(value, date_format).date()

    return func


def convert_date_to_format(date_format: str) -> Callable[[datetime.date], str]:
    def func(value: any) -> str:
        if value is None:
            return ""
        if not isinstance(value, datetime.date):
            return ""
        return value.strftime(date_format)

    return func

def to_lower(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    return text.lower()

def to_upper(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    return text.upper()


def strip(letters: str = None) -> Callable[[str], str]:
    if letters is None:
        letters = " "

    def func(text: str) -> str:
        if not isinstance(text, str):
            return ""

        return text.strip(letters)

    return func
