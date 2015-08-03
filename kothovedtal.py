import argparse

import xlrd


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('filenames', nargs='+')
    args = parser.parse_args()

    for filename in args.filenames:
        process(filename)


def remove_suffix(s, suf):
    if s.endswith(suf):
        return s[:-len(suf)]
    else:
        return s


def process(filename):
    if filename.endswith('xlsx'):
        year = remove_suffix(filename, '.xlsx')[-4:]
    else:
        year = remove_suffix(filename, '_excel.xls')[-4:]
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)

    ks = 'Matematik Matematik-økonomi Nanoscience IT Fysik Datalogi'.split()
    data = {
        k: ()
        for k in ks
    }

    try:
        row, col = get_start(sheet)
    except TypeError:
        sheet.dump()
        print(year)
        raise
    for i in range(row + 1, sheet.nrows):
        name = sheet.cell_value(i, col).strip()
        name = remove_suffix(name, ', Aarhus C, Studiestart: Sommerstart')
        if name == 'It':
            name = 'IT'
        if name in ('Nanoteknologi', 'Nanosceience'):
            name = 'Nanoscience'
        if data.get(name) == ():
            data[name] = sheet.row_values(i, col + 1)

    if not any(data.values()):
        sheet.dump()
    res = []
    for k in ks:
        v = data[k]
        if v:
            res.append(int(v[0]))
        else:
            res.append('-')
    print(repr((int(year), res)) + ',')


def get_start(sheet):
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            v = str(sheet.cell_value(i, j)).strip()
            if v == 'Aarhus Universitet':
                return i, j
    return None



if __name__ == "__main__":
    main()
