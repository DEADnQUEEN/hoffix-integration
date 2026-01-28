import string
from typing import Callable

import constants
from com_excel.functions import filters


class Column:
    @property
    def rename(self):
        if self.__rename is None:
            return self.__column
        return self.__rename

    @rename.setter
    def rename(self, value: str):
        self.__rename = value


    @property
    def start(self):
        return self.__start

    def __init__(
            self,
            column: str,
            stop_filters: list[Callable[[any], bool]] = None,
            skip_filters: list[Callable[[any], bool]] = None,
            stop_if_null: bool = True,
            skip_if_null: bool = True,
            edit_value: list[Callable[[any], any]] = None,
            start: int = 1,
            group: int = constants.STEP,
            rename: str = None,
            hidden: bool = False,
    ):
        self.__stop_filters = stop_filters or []
        if stop_if_null and filters.is_not_none not in self.__stop_filters:
            self.__stop_filters.append(filters.is_not_none)

        self.__skip_filters = skip_filters or []
        if skip_if_null and filters.is_not_none not in self.__skip_filters:
            self.__skip_filters.append(filters.is_not_none)

        self.__edit_value = edit_value or []

        for letter in column:
            if letter not in string.ascii_uppercase:
                raise ValueError()

        self.__column = column
        self.__range = None
        self.__position = None

        self.__start = start if start > 1 else 1
        self.__start_from = self.__start
        self.__group = group if group > 1 else 1

        self.__ws = None
        self.__rename = rename

        self.hidden = hidden

    def set_sheet(self, ws):
        self.__ws = ws

    def __get_range(self):
        return self.__ws.Range(f"{self.__column}{self.__start}:{self.__column}{self.__group + self.__start}").Cells()

    def __iter__(self):
        self.__start = self.__start_from
        if self.__ws is None:
            raise AttributeError

        self.__range = iter(self.__get_range())

        return self

    def __get_values(self):
        try:
            value = next(self.__range)
        except StopIteration:
            self.__start += self.__group
            self.__range = iter(self.__get_range())
            value = next(self.__range)

        return value[0]

    def __next__(self) -> tuple[any, bool]:
        value = self.__get_values()

        for column_filter in self.__stop_filters:
            if not column_filter(value):
                raise StopIteration

        for skip_filter in self.__skip_filters:
            if not skip_filter(value):
                return None, True

        for edit_function in self.__edit_value:
            value = edit_function(value)
        return value, False

    def get_value(self, row: int) -> any:
        """
        get column value of exact row
        :param row: int row number
        :return: value in cell without filters and skips
        """

        if self.__ws is None:
            raise AttributeError

        return self.__ws.Range(f"{self.__column}{row}").Cells()

    def set_value(self, row: int, value: any):
        self.__ws.Range(f"{self.__column}{row}").Value = value

    def write(self, data: list[any], start_from: int = 1):
        if start_from <= 1:
            start_from = 1
        self.__ws.Range(
            f"{self.__column}{start_from}:{self.__column}{start_from + len(data)-1}"
        ).Value = [
            [value]
            for value in data
        ]


class Sheet:
    def __init__(
            self,
            columns: list[Column],
            sheet
    ):
        self.__columns = columns
        self.__sheet = sheet

        for column in self.__columns:
            column.set_sheet(self.__sheet)

    def __iter__(self):
        for column in self.__columns:
            column.set_sheet(self.__sheet)
            iter(column)
        return self

    def count(self):
        count = 0

        iter(self)
        while True:
            for column in self.__columns:
                try:
                    next(column)
                except StopIteration:
                    return count

            count += 1

    def write(
            self,
            start_from: int,
            data: list[dict[str, any]],
            order: list[str] = None
    ):
        order = order if order is not None else [column.rename for column in self.__columns]

        columns = {
            column.rename: column
            for column in self.__columns
        }

        values = {
            column: ["" for _ in range(len(data))]
            for column in order
            if column in columns
        }
        for row_index, row in enumerate(data):
            for column_index, column in enumerate(row.items()):
                key, value = column
                if key not in values.keys():
                    continue
                values[key][row_index] = value

        for order_column in order:
            column = columns.get(order_column)
            if column is None:
                continue

            column.write(values.get(order_column, []), start_from=start_from)

    def overwrite(
            self,
            data: list[dict[str, any]],
            from_row: int = 1
    ):
        if from_row <= 1:
            from_row = 1
        self.write(
            from_row,
            data
        )

    def __next__(self) -> dict[str, any]:
        while True:
            skip = False
            values = {}

            for column in self.__columns:
                try:
                    value, skip_value = next(column)
                    skip |= skip_value
                except StopIteration as stop:
                    raise stop

                if not column.hidden:
                    values[column.rename] = value

            if not skip:
                return values

    def get_value(self, row: int, column: int) -> any:
        return self.__columns[column].get_value(row)

