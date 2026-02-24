import datetime
import json
from typing import Optional, Callable
import win32com.client

import utils.constants
import utils.com_excel.wrap
from utils import hoffix, functions

from packets.calls import constants


def send_request(api: hoffix.HoffixAPI, data: dict[str, any]) -> Optional[dict[str, any]]:
    for skip_field, skip_func in utils.constants.SKIP_FIELDS.items():
        if skip_field not in data:
            continue
        if skip_func(data[skip_field]):
            return None

    hoffix_order = api.find_order(data["orderServiceNumber"], data["workDate"])
    if hoffix_order is None:
        return None

    order_id = hoffix_order['visitOrderId']

    return api.get_full_order_data(order_id)


def write_to_excel(sourceWorkbook: str, sourceSheet: str, columns: dict[str, int], data: list[dict[str, any]]) -> None:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(sourceWorkbook)

    sheet = wb.Worksheets(sourceSheet)
    sh = utils.com_excel.wrap.Sheet(
        [
            utils.com_excel.wrap.Column(column=functions.convert_index_to_column(index), rename=column_name)
            for column_name, index in columns.items()
        ],
        sheet,
    )

    for row in data:
        sh.write(
            row['row'],
            [row]
        )


def collect_calls(order_data: dict[str, any]) -> list[dict[str, any]]:
    return order_data.get('callInfo', [])


def group_objects(object_array: list[dict[str, any]], field: str, function_list: list[Callable[[any], bool]]) -> list[list[dict[str, any]]]:
    groups = [[] for _ in range(len(function_list) + 1)]

    for item in object_array:
        value = item.get(field)

        if value is None:
            continue

        for index, function in enumerate(function_list):
            f = function(value)
            print(f, item)
            if f:
                groups[index].append(item)
                break
        else:
            groups[-1].append(item)

    return groups


def is_same_date(compare_to_date: datetime.datetime) -> Callable[[datetime.datetime], bool]:
    date = compare_to_date.date()

    def compare(value: any) -> bool:
        if isinstance(value, datetime.datetime):
            value = value.date()
        elif isinstance(value, datetime.date):
            pass
        elif isinstance(value, str):
            value = datetime.datetime.strptime(value, constants.DATETIME_FORMAT)
        else:
            raise Exception(f"Unknown date type: {type(value)}")
        return date == value

    return compare


def is_not_same_date(compare_to_date: datetime.datetime) -> Callable[[datetime.datetime], bool]:
    func = is_same_date(compare_to_date)
    def compare(value: any) -> bool:
        return not func(value)
    return compare


def string_format_call(call: dict[str, any]) -> str:
    start_time = call.get('callStartDT')
    if start_time is None:
        raise Exception("start time is None")

    end_time = call.get('callEndDT')
    if end_time is None:
        raise Exception("end time is None")

    start_time = datetime.datetime.strptime(start_time, constants.DATETIME_FORMAT)
    end_time = datetime.datetime.strptime(end_time, constants.DATETIME_FORMAT)

    delta = end_time - start_time
    state = 'X' if delta < constants.DELTA else 'V'

    start_end_str = f"{start_time.strftime(constants.TIME_FORMAT)} - {end_time.strftime(constants.TIME_FORMAT)}"
    hours, minutes, second = utils.functions.convert_seconds_to_time(int(delta.total_seconds()))
    delta_format = "{hours:2d}:{minutes:2d}:{seconds:2d}"
    delta_str = f"Длительность: {delta_format.format(hours=hours, minutes=minutes, seconds=second).replace(' ', '0')}"
    return f"{state} | {start_time.strftime(constants.DATE_FORMAT)} {start_end_str} {delta_str}"


def string_format_calls(calls: list[dict[str, any]]) -> str:
    call_strings = []
    for index, call in enumerate(calls):
        call_strings.append(f"{index + 1}) {string_format_call(call)}")

    return "\n".join(call_strings)


def parse_row(api: hoffix.HoffixAPI, row: dict[str, any]) -> dict[str, any]:
    order = send_request(api, row)
    if order is None:
        return {}
    print(json.dumps(order))
    calls = collect_calls(order)

    date = order.get("workDate")
    if date is None:
        raise Exception
    date = datetime.datetime.strptime(date, "%Y-%m-%d")

    try:
        same_date_call_group, ungrouped = group_objects(calls, "callStartDT", [is_same_date(date)])
    except Exception as e:
        print(order)
        raise e

    print("same date:")
    print(string_format_calls(same_date_call_group))
    print('---\nother:')
    print(string_format_calls(ungrouped))

    return {}


def main(excel_json: dict[str, any]):
    api = hoffix.HoffixAPI(excel_json['auth']['login'], excel_json['auth']['password'])

    for excel_row in excel_json["orders"]:
        parse_row(api, excel_row)

