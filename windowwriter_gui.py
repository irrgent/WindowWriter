import tkinter as tk
import windowwriter
import windowwriter_cli
from pywinauto import application


class MacroListbox(tk.Frame):

    def __init__(self, parent, macro_dict, *args, **kwargs):

        tk.Frame.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict
        self.win_hwnd = None
        self._listbox = tk.Listbox(self, selectmode=tk.SINGLE)
        self._listbox.bind("<<ListboxSelect>>", self.macro_select)
        self._listbox.pack()

        for _ in self._macro_dict.keys():
            self._listbox.insert(tk.END, _)

    def update_macros(self, new_dict):
        raise NotImplementedError

    def connect_window(self, hwnd):
        self.win_hwnd = hwnd
        self.app = application.Application().connect(handle=self.win_hwnd)

    def macro_select(self, event):

        if self.win_hwnd is None:
            print("No window selected.")
        else:
            wdg = event.widget
            idx = int(wdg.curselection()[0])
            value = wdg.get(idx)

            dlg = self.app.top_window()
            dlg.type_keys(self._macro_dict[value], with_spaces=True)

            print("Selected {}.".format(value))


def main():

    macro_dict = windowwriter.macro_dict("./macros.csv")
    windows = windowwriter.get_windows()
    hwnd = windows[
        windowwriter_cli.select_from_dict(
            windowwriter_cli.numbered_dict_keys(windows))]
    root = tk.Tk()
    lb = MacroListbox(root, macro_dict)
    lb.connect_window(hwnd)
    lb.pack()
    root.mainloop()


if __name__ == "__main__":
    main()
