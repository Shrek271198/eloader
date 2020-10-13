import math
from operator import __eq__, __ge__


def vlookup(key, table, column, approximate_match=True):
    """ This function emulates vlookup functionality found in Excel."""

    compare = __ge__ if approximate_match else __eq__
    try:
        return max(row for row in table if compare(key, row[0]))[column]
    except ValueError:
        return None


def round_up(n, decimals=0):
    """ This function emulates the round functionality found in Excel."""

    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier
