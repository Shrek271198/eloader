from operator import __ge__, __eq__
import math

def vlookup(key, table, column, approximate_match=True):
    compare = __ge__ if approximate_match else __eq__
    try:
        return max(row for row in table if compare(key, row[0]))[column]
    except ValueError:
        return None

def round_up(n, decimals=0):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier

