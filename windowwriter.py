import win32gui
import csv


# Get handles of all visible windows and store in dictionary with
# their names as the keys.
def get_windows():

    def callback(hwnd, hwnd_dict):

        if win32gui.IsWindowVisible(hwnd):
            hwnd_dict[win32gui.GetWindowText(hwnd)] = hwnd
        return True

    hwnd_dict = {}

    win32gui.EnumWindows(callback, hwnd_dict)

    return hwnd_dict


# Reads csv file line by line using it to create a dictionary of macros
def macro_dict(filePath):

    with open(filePath) as csvFile:

        # Reads with dialect='excel' by default
        csvreader = csv.reader(csvFile)
        csvdict = {}

        for row in csvreader:

            if len(row) < 2:
                raise ValueError("Row with less than 2 columns found.")
            elif len(row) > 2:
                raise ValueError("Row with more than 2 columns found.")

            csvdict[row[0]] = row[1]

    return csvdict


if __name__ == '__main__':

    macros = macro_dict('./macros.csv')
    windows = get_windows()
