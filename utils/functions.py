import string

__alphabet = string.ascii_uppercase

def convert_index_to_column(index: int) -> str:
    letters = ""

    while index > 0:
        index -= 1
        letters = __alphabet[index % len(__alphabet)] + letters
        index //= len(__alphabet)

    return letters


def convert_seconds_to_time(seconds: int) -> tuple[int, int, int]:
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return hours, minutes, seconds

