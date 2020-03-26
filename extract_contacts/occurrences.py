import pandas as pd
import xlrd
import re
import os


def get_occurrences(pattern, data):
    regex = re.compile(pattern)
    all_occurrences = []

    for index, row in data.iterrows():
        row_occurrences = []

        for col in data.columns:
            cell_occurrences = regex.findall(str(row[col]))
            if cell_occurrences:
                # Normalize
                cell_occurrences = [e.replace(' ', '').lower()
                                    for e in cell_occurrences]
                row_occurrences += cell_occurrences

        if (len(row_occurrences) > 0):
            all_occurrences += row_occurrences

    # Clean duplicates
    all_occurrences = list(set(all_occurrences))

    return all_occurrences


def load_data(filename):
    if not os.path.isfile(filename):
        raise InvalidFile(f'"{filename}" is not a valid file.')

    try:
        data = pd.read_excel(filename)
    except xlrd.biffh.XLRDError:
        raise InvalidFileType(
            f'Error parsing the "{filename}" file. '
            'Perhaps this file is not an Excel file.'
        )

    return data


class InvalidFile(Exception):
    pass


class InvalidFileType(Exception):
    pass
