from typing import Callable
import datetime


def is_exact_equal(to: any) -> Callable[[any], bool]:
    def inside_is_equal(cell_value: any):
        return cell_value == to

    return inside_is_equal


def is_lowered_string_equal(to: str) -> Callable[[any], bool]:
    to = to.lower()
    def inside_is_lowered(cell_value: str):
        return cell_value.lower() == to
    return inside_is_lowered


def is_lowered_string_not_equal(to: str) -> Callable[[any], bool]:
    to = to.lower()
    def inside_is_lowered(cell_value: str):
        return cell_value.lower() != to
    return inside_is_lowered


def is_not_none(cell_value: any):
    return cell_value is not None


def from_date(date_from: datetime.date, convert_to_datetime: Callable[[str], datetime.datetime]) -> Callable[[any], bool]:
    def filter_date(cell_value: any) -> bool:
        if isinstance(cell_value, datetime.datetime):
            return cell_value.date() >= date_from
        if isinstance(cell_value, datetime.date):
            return cell_value >= date_from
        if isinstance(cell_value, str):
            try:
                return convert_to_datetime(cell_value).date() >= date_from
            except Exception:
                return False
        raise Exception
    return filter_date


def to_date(date_to: datetime.date, convert_to_datetime: Callable[[str], datetime.datetime]) -> Callable[[any], bool]:
    def filter_date(cell_value: any) -> bool:
        if isinstance(cell_value, datetime.datetime):
            return cell_value.date() <= date_to
        if isinstance(cell_value, datetime.date):
            return cell_value <= date_to
        if isinstance(cell_value, str):
            try:
                return convert_to_datetime(cell_value).date() <= date_to
            except Exception:
                return False
        raise Exception
    return filter_date

