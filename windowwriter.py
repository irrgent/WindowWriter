import win32gui
import win32api
import win32con
import csv


class WindowNotFoundError(Exception):
    pass


# Return dict with titles and handles of visible windows.
def get_windows():

    def callback(hwnd, lst):

        if win32gui.IsWindowVisible(hwnd):
            windows[win32gui.GetWindowText(hwnd)] = hwnd
        return True

    windows = {}

    win32gui.EnumWindows(callback, windows)

    return windows


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


def send_input(wsh, win_title, hwnd, string):

    replace = {'\n': '{ENTER}', '\t': '{TAB}', '+': '{+}'}

    # If window is minimized restore it.
    if win32gui.IsIconic(hwnd):
        if not win32gui.ShowWindow(hwnd, win32con.SW_RESTORE):
            raise WindowNotFoundError(
                "ShowWindow returned 0 for window {} with hwnd {}.".format(
                    win_title, hwnd))

    if not wsh.AppActivate(win_title):
        raise WindowNotFoundError(
            "Could not activate window {} with hwnd {}.".format(
                win_title, hwnd))

    for key in list(string):
        if key in replace:
            wsh.SendKeys(replace[key])
        else:
            wsh.SendKeys(key)


if __name__ == '__main__':

    macros = macro_dict('./macros.csv')
    windows = get_windows()
