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


# Return value from dictionary selected by user
def select_from_dict(selections, print_selections=True):

    if all([isinstance(k, str) for k in selections.keys()]):
        key_type = "str"
    elif all([isinstance(k, int) for k in selections.keys()]):
        key_type = "int"
    else:
        raise TypeError(
            "Dictionary must be solely composed of strings xor integers.")

    if print_selections:

        for k in selections.keys():
            print("{:20}: {}".format(k, selections[k]))

    ret_val = None

    while ret_val is None:

        choice = input("> ")

        if key_type == "int":

            try:
                choice = int(choice)
            except ValueError:
                print("Please enter an integer.")
                continue

        try:
            ret_val = selections[choice]
        except KeyError:
            print("Invalid selection.")
            ret_val = None
            continue

    return ret_val


if __name__ == '__main__':

    macros = macro_dict('./macros.csv')
    windows = get_windows()
