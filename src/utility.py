from operator import __ge__, __eq__

def vlookup(key, table, column, approximate_match=True):
    compare = __ge__ if approximate_match else __eq__
    try:
        return max(row for row in table if compare(key, row[0]))[column]
    except ValueError:
        return None

