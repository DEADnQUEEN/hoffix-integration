from typing import Callable


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
