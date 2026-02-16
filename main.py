import datetime
import json
import sys
from typing import Optional

import constants
import hoffix
import com_excel.wrap
import win32com.client


def edit_values(api: hoffix.HoffixAPI, base_order: dict[str, any], set_values: dict[str, any]) -> dict[str, any]:
    order: dict[str, any] = {**base_order}
    for key, value in set_values.items():
        if key not in order.keys():
            continue

        dataset = constants.FROM_DB_COLLECT.get(key)
        if dataset is not None:
            value = api.get_dataset_from_db(*dataset)[value]

        order[key] = value

    return order


def remap_put_json(order_data: dict[str, any]):
    edit_order_data = {}

    for key, value in constants.MAPPING_JSON.items():
        if isinstance(value, str):
            if not value:
                edit_order_data[key] = value
                continue
            path = iter(value.split("."))
            field = order_data.get(next(path), None)

            for item in path:
                if field is None:
                    break
                if not isinstance(field, dict):
                    raise TypeError
                field = field.get(item)

            if field is None:
                continue
        else:
            field = value

        edit_order_data[key] = field

    return edit_order_data


def send_request(api: hoffix.HoffixAPI, data: dict[str, any]) -> Optional[dict[str, any]]:
    response = {
        "datetime": datetime.datetime.now().strftime(constants.DATE_FORMAT),
        "order_id": data["orderServiceNumber"],
        "worker_rename": data["workerId"],
        "state": "Не выполнено",
        "comment": "",
    }

    for skip_field, skip_func in constants.SKIP_FIELDS.items():
        if skip_field not in data:
            continue
        if skip_func(data[skip_field]):
            return None

    hoffix_order = api.find_order(data["orderServiceNumber"], data["workDate"])
    if hoffix_order is None:
        response["comment"] = "Не найден заказ в Hoffix"
        return response

    order_id = hoffix_order['visitOrderId']

    hoffix_put_order = remap_put_json(api.get_full_order_data(order_id))
    updated_data = edit_values(api, hoffix_put_order, data)

    api.set_order_data(order_id, updated_data)

    response['state'] = "Выполнено"
    return response


def write_to_excel(filepath: str, data: list[dict[str, any]]) -> None:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(filepath)

    sheet = wb.Worksheets(constants.OUTPUT_LIST)
    sh = com_excel.wrap.Sheet(
        [
            com_excel.wrap.Column("A", rename="datetime"),
            com_excel.wrap.Column("B", rename="order_id"),
            com_excel.wrap.Column("D", rename="worker_rename", stop_if_null=False),
            com_excel.wrap.Column("E", rename="state"),
            com_excel.wrap.Column("F", rename="comment", stop_if_null=False),
        ],
        sheet,
    )

    sh.write(
        sh.count() + 1,
        data
    )


def main():
    if len(sys.argv) < 2:
        raise TypeError

    excel_json = json.loads(sys.argv[1].replace("\\", "\\\\"))
    api = hoffix.HoffixAPI(excel_json['auth']['login'], excel_json['auth']['password'])

    output = []

    for excel_row in excel_json["orders"]:
        request = send_request(api, excel_row)
        if request is not None:
            output.append(request)

    write_to_excel(excel_json['ExcelData']['sourceWorkbook'], output)

if __name__ == '__main__':
    main()
