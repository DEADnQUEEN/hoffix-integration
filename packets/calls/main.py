import datetime
from typing import Optional, Callable
import win32com.client

import utils.constants
import utils.com_excel.wrap
from utils import hoffix, functions

from packets.calls import constants
from packets.calls import functions as call_functions


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


def open_excel(sourceWorkbook: str, sourceSheet: str, columns: dict[str, int]) -> utils.com_excel.wrap.Sheet:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(sourceWorkbook)

    sheet = wb.Worksheets(sourceSheet)
    return utils.com_excel.wrap.Sheet(
        [
            utils.com_excel.wrap.Column(column=functions.convert_index_to_column(index), rename=column_name)
            for column_name, index in columns.items()
        ],
        sheet,
    )


def collect_calls(order_data: dict[str, any]) -> list[dict[str, any]]:
    return order_data.get('callInfo', [])


def group_objects(object_array: list[dict[str, any]], functions_data: dict[str, tuple[Callable, str]]) -> dict[str, list[any]]:
    groups = {
        key: []
        for key in functions_data.keys()
    }

    for item in object_array:
        for key, func_data in functions_data.items():
            function, field = func_data
            value = item.get(field)

            if value is None:
                break

            if function(value):
                groups[key].append(item)

    return groups


def string_format_call(call: dict[str, any]) -> str:
    start_time = call.get(constants.CALL_START_FIELD)
    if start_time is None:
        raise Exception("start time is None")

    end_time = call.get(constants.CALL_END_FIELD)
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


def unite_group(group: list[dict[str, any]]) -> str:
    return "\n".join([f"{index + 1}) {string_format_call(item)}" for index, item in enumerate(group)])


def parse_excel_row(api: hoffix.HoffixAPI, row: dict[str, any]) -> Optional[dict[str, any]]:
    order = send_request(api, row)
    if order is None:
        return None

    calls = collect_calls(order)

    grouped_calls = group_objects(calls, {key: (function[0](order, function[1]), function[2]) for key, function in call_functions.FUNCTIONS.items()})
    grouped_calls_string: dict[str, str] = {}
    for group, calls in grouped_calls.items():
        grouped_calls_string[group] = unite_group(calls)

    return grouped_calls_string


def main(excel_json: dict[str, any]):
    api = hoffix.HoffixAPI(excel_json['auth']['login'], excel_json['auth']['password'])

    sheet = open_excel(excel_json['ExcelData']['sourceWorkbook'], excel_json['ExcelData']['sourceSheet'], excel_json['ExcelData']['columns'])

    for excel_row in excel_json["orders"]:
        parsed_row = parse_excel_row(api, excel_row)
        if parsed_row is None:
            continue

        sheet.write(excel_row['excelRow'], parsed_row)

