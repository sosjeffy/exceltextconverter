import pandas as pd


def _return_df(excel_filename: str, sheet_name: str):
    """Opens excel for usage"""
    return pd.read_excel(excel_filename, sheet_name)


def _write_file(filename:str, data:str):
    """Writes data to file"""
    with open(filename, 'a') as outfile:
        outfile.write(data)


def _collect_excel_name():
    """Collects filename of excel document to parse through"""
    excel_filename = input("Enter your excel file name (Format must be 'FILE_NAME.xlsx' w/o quotes): ")
    while excel_filename[-5:] != '.xlsx':
        print('Oops, that did not follow the format!')
        excel_filename = input("Enter your excel file name (Format must be 'FILE_NAME.xlsx' w/o quotes): ")
    return excel_filename


def _collect_sheet_name():
    """Collects name of the excel sheet to parse through. This assumes user can input correct sheet name"""
    sheet_name = input("Enter the name of the sheet that you wish to use: ")
    return sheet_name


def _collect_filename():
    """Collects desired .txt file name"""
    filename = input("Enter your desired output file name (Format must be 'FILE_NAME.txt' w/o quotes): ")
    while filename[-4:] != '.txt':
        print('Oops, that did not follow the format!')
        filename = input("Enter your desired output file name (Format must be 'FILE_NAME.txt' w/o quotes): ")
    return filename


def _collect_column_count() -> int:
    """Collects the # of desired columns to look at; number is assumed to be in range"""
    while True:
        try:
            count = int(input("Enter the number of columns you'd like to combine: "))
            break
        except ValueError:
            print("Oops, that was not a valid number.")
            continue
    return count


def _collect_column_names(column_count:int) -> list:
    """Collects name of every column; each column's name is assumed to be correct"""
    result = []
    for i in range(column_count):
        result.append(f'df["{input("Enter name of desired column: ")}"]')
    return result


def _collect_desired_first_col() -> str:
    """Collects user's desired first col"""
    col = input("Enter the column you'd like to appear first in your file (options w/o quotes: 'email'/'name'): ")
    while col != 'email' and col != 'name':
        print("Oops, that was not a valid entry. You may only enter 'email' or 'name'.")
        col = input("Enter the column you'd like to appear first in your file (options w/o quotes: 'email'/'name'): ")
    return col


def _build_zip(column_count: int) -> str:
    """Create function that will separate each entry's data and keep them together respectively"""
    return f'zip({", ".join(_collect_column_names(column_count))})'


def _fix_name(data: tuple, formatted_name: str, desired_first_col: str) -> tuple:
    """Replaces unformatted name w/ formatted name in tuple;name index initial assumed to be 0;puts email to 0 index"""
    data = list(data)
    if desired_first_col == 'email':
        emails = data[1]
        data[0] = emails
        data[1] = formatted_name
    else:
        data[0] = formatted_name
    return tuple(data)


def run():
    df = _return_df(_collect_excel_name(), _collect_sheet_name())  # For use in build_zip
    filename = _collect_filename()
    col = _collect_desired_first_col()
    l = 0
    for data in eval(_build_zip(_collect_column_count())):
        # Below corrects name formatting
        data = _fix_name(data, ' 	'.join([name.strip().lower().capitalize() for name in data[0].split()]), col)
        # Writes each line to file
        _write_file(filename, ' 	'.join([dat_element for dat_element in data]) + '\n')


if __name__ == "__main__":
    run()
    print('--------------Operation completed.--------------')






