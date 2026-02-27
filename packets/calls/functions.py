import datetime
from typing import Callable
from packets.calls import constants

def false_func(_: datetime.datetime) -> bool:
    return False

def is_same_date(order: dict[str, any], field: str) -> Callable[[any], bool]:
    date = order.get(field)
    if date is None:
        return false_func
    try:
        date = datetime.datetime.strptime(date, "%Y-%m-%d").date()
    except ValueError:
        return false_func

    def compare(value: any) -> bool:
        if isinstance(value, datetime.datetime):
            value = value.date()
        elif isinstance(value, datetime.date):
            pass
        elif isinstance(value, str):
            value = datetime.datetime.strptime(value, constants.DATETIME_FORMAT).date()
        else:
            raise Exception(f"Unknown date type: {type(value)}")
        return date == value

    return compare


def is_not_same_date(order: dict[str, any], field: str) -> Callable[[any], bool]:
    func = is_same_date(order, field)
    def compare(value: any) -> bool:
        return not func(value)
    return compare


FUNCTIONS: dict[str, tuple[Callable[[dict[str, any], Callable[[any], bool]]], str, str]] = {
    "callPrevdayCol": (is_same_date, "workDate", constants.CALL_START_FIELD),
    "callSamedayCol": (is_not_same_date, "workDate", constants.CALL_START_FIELD),
}

