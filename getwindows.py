import pywinauto
from pywinauto import application
import win32api
import win32process
import win32con
import win32gui


def list_windows():

    handles = pywinauto.findwindows.enum_windows()
    print(handles)

    processes = []

    for hwnd in handles:

        pid = win32process.GetWindowThreadProcessId(hwnd)

        try:
            handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION |
                                          win32con.PROCESS_VM_READ,
                                          False, pid[1])
            proc_name = win32process.GetModuleFileNameEx(handle, 0)
            processes.append(proc_name)
        except:
            print("Could not open process.")

    return processes


def get_windows():

    def callback(hwnd, hwnd_dict):

        if win32gui.IsWindowVisible(hwnd):
            hwnd_dict[win32gui.GetWindowText(hwnd)] = hwnd
        return True

    hwnd_dict = {}

    win32gui.EnumWindows(callback, hwnd_dict)

    return hwnd_dict


if __name__ == "__main__":

    wd = get_windows()

    for k in wd.keys():

        if 'notepad' in k.lower():

            app = application.Application().connect(handle=wd[k])

            dlg = app.top_window()

            dlg.type_keys("some text", with_spaces=True)
