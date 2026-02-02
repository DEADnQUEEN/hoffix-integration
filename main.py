from argparse import ArgumentTypeError

from playwright.sync_api import sync_playwright, Page
from win32com.client import CDispatch
import win32com.client
import datetime
import sys
import traceback

import com_excel.wrap
import constants, utils

from com_excel.functions import filters, editors


def get_password(wb) -> tuple[str, str]:
    column, row = constants.LOGIN_PASSWORD_FIELD
    row = int(row)
    sh = com_excel.wrap.Sheet(
        wb.Worksheets(constants.SETTING_LIST),
        [
            com_excel.wrap.Column(column, rename="settings"),
        ]
    )

    login_pair: str = sh.get_value(row, 0)
    login, password = login_pair.split(":", 1)

    return login, password


def auth_hoffix(page: Page, login: str, password: str) -> Page:
    return utils.auth(
        page,
        constants.HOFFIX_LOGIN_URL,
        constants.LOGIN_PASSWORD_FIELDS_SELECTOR,
        constants.CONFIRM_LOGIN_SELECTOR,
        login,
        password,
        constants.HOFFIX_MAIN_URL,
    )


def get_workers(workbook) -> dict[str, str]:
    worker_sheet = com_excel.wrap.Sheet(
        workbook.Worksheets(constants.WORKER_SHEET),
        [
            com_excel.wrap.Column("A", rename="renamed_for_hoffix", edit_value=[
                lambda value: str(value),
                editors.strip()
            ]),
            com_excel.wrap.Column("C", rename="excel_worker", edit_value=[lambda value: str(value)]),
        ],
    )

    return {
        row["excel_worker"]: row["renamed_for_hoffix"]
        for row in worker_sheet
    }


def get_workbook():
    if len(sys.argv) < 2 and not isinstance(sys.argv[1], str):
        raise ArgumentTypeError("No workbook specified")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    return excel.Workbooks.Open(sys.argv[1])


def get_services(wb, worker_mapping: dict[str, str]):
    def worker_mapping_function(worker: str) -> tuple[str, str]:
        return worker, worker_mapping.get(worker, "")

    sheet = wb.Worksheets(constants.SERVICES_SHEET)

    sh = com_excel.wrap.Sheet(
        sheet,
        [
            com_excel.wrap.Column(
                "B",
                start=11,
                stop_if_null=False,
                rename="date",
                edit_value=[
                    editors.convert_to_date(constants.EXCEL_DATE_FORMAT),
                    editors.convert_date_to_format(constants.HOFFIX_DATE_FORMAT),
                ]
            ),
            com_excel.wrap.Column("E", start=11, stop_if_null=False, rename="order_id"),
            com_excel.wrap.Column("O", start=11, rename="values_filter", hidden=True),
        ],
    )

    sh.add_column(
        com_excel.wrap.Column(
            sh.find_header_column("Мастер", 10),
            start=11,
            stop_if_null=False,
            skip_filters=[
                filters.is_not_none,
                filters.is_lowered_string_not_equal('отмена')
            ],
            rename="worker",
            edit_value=[
                worker_mapping_function,
            ]
        )
    )

    for row in sh:
        yield row


def format_url(order_id: str, date: str) -> str:
    return f"https://hoffix.hoff.ru/orders?search={order_id}&workDateFrom={date}&workDateTo={date}"


def parse_row(row: dict[str, str]) -> dict[str, str]:
    return {
        "order_id": row["order_id"],
        "datetime": datetime.datetime.now().strftime(constants.OUTPUT_DATE_FORMAT),
        "worker_name": row["worker"][0],
        "worker_rename": row["worker"][1],
        "comment": "",
        "state": "",
        "date": row["date"],
    }


def fill_row(page: Page, data) -> dict[str, str]:
    page.goto(format_url(data["order_id"], data["date"]))
    page.wait_for_selector(".el-table__header-wrapper")

    page.wait_for_timeout(50)

    try:
        page.locator(constants.TABLE_ELEMENTS).click(timeout=3_000)
    except Exception:
        data['comment'] = "Заказ не найден в Hoffix"
        data['state'] = "Невыполнено"
        return data

    if data["worker_rename"] == "":
        data['comment'] = "Не найден исполнитель"
        data['state'] = "Невыполнено"
        return data

    utils.get_safe_locator(page, constants.EDIT_BUTTON).click()
    utils.get_safe_locator(page, constants.WORKER_SELECT).click()

    while not page.evaluate('() => {return document.querySelectorAll("body > div > p.el-select-dropdown__empty").length === 0}'):
        page.wait_for_timeout(10)
    page.wait_for_timeout(50)

    utils.get_safe_locator(page, constants.WORKER_SELECT+"> input")

    if not page.evaluate(
        '(name) => {var found = false; var selector = document.querySelectorAll(".el-select-dropdown__item");for (let i = 0; i < selector.length;i++) {if (selector[i].querySelector("span").textContent.toLowerCase().includes(name.toLowerCase())){selector[i].click(); found = true;}} return found;}',
        data["worker_rename"]
    ):
        data['comment'] = "Исполнитель не найден"
        data['state'] = "Невыполнено"
        return data

    utils.get_safe_locator(page, constants.SAVE_BUTTON).click()
    while not page.evaluate(constants.WAIT_SCRIPT):
        page.wait_for_timeout(10)

    data['state'] = "Выполнено"

    return data


def write_to_excel(wb, data: list[dict[str, str]]) -> None:
    sheet = wb.Worksheets(constants.OUTPUT_LIST)
    sh = com_excel.wrap.Sheet(
        sheet,
        [
            com_excel.wrap.Column("A", rename="datetime"),
            com_excel.wrap.Column("B", rename="order_id"),
            com_excel.wrap.Column("C", rename="worker_name", stop_if_null=False),
            com_excel.wrap.Column("D", rename="worker_rename", stop_if_null=False),
            com_excel.wrap.Column("E", rename="state"),
            com_excel.wrap.Column("F", stop_if_null=False, rename="comment"),
        ],
    )

    sh.write(
        sh.count() + 1,
        data
    )


def main():
    wb = get_workbook()
    workers = get_workers(wb)
    login, password = get_password(wb)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        for _ in range(constants.ATTEMPTS):
            try:
                auth_hoffix(page, login, password)
                break
            except Exception:
                print("timeout to login")
        else:
            print("couldn't login. check internet connection or login, password")
            input("press enter to continue... ")
            return

        data = []
        for i, row in enumerate(get_services(wb, workers)):
            row = parse_row(row)

            try:
                for _ in range(constants.ATTEMPTS):
                    content = fill_row(page, row)
                    break
                else:
                    row["state"] = "Невыполнено"
                    row['comment'] = "Ошибка при подключении к Hoffix"
            except Exception:
                row["state"] = "Невыполнено"
                row['comment'] = "Ошибка скрипта"

            data.append(content)

    write_to_excel(wb, data)


if __name__ == '__main__':
    print(sys.argv)

    if len(sys.argv) < 2:
        raise ArgumentTypeError("Path not specified")

    try:
        main()
    except Exception as e:
        print(traceback.format_exc())
        input("Ошибка\npress enter to continue...")

        raise e
