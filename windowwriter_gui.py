import tkinter as tk
import windowwriter
import windowwriter_cli
import win32com.client as comclt


class MacroListbox(tk.Frame):

    def __init__(self, parent, macro_dict, *args, **kwargs):

        tk.Frame.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict
        self.win_title = None
        self._listbox = tk.Listbox(self, selectmode=tk.SINGLE)
        self._listbox.bind("<<ListboxSelect>>", self.macro_select)
        self._listbox.pack()

        for _ in self._macro_dict.keys():
            self._listbox.insert(tk.END, _)

    def update_macros(self, new_dict):
        raise NotImplementedError

    def connect_window(self, title):
        self.win_title = title
        self._wsh = comclt.Dispatch("WScript.Shell")

    def macro_select(self, event):

        if self.win_title is None:
            print("No window selected.")
        else:
            wdg = event.widget
            idx = int(wdg.curselection()[0])
            value = wdg.get(idx)

            windowwriter.send_input(self._wsh, self.win_title,
                                    self._macro_dict[value])

            print("Selected {}.".format(value))


def main():

    macro_dict = windowwriter.macro_dict("./macros.csv")
    windows = windowwriter.get_windows()

    numbered = windowwriter_cli.numbered_dict_keys(windows)
    win_title = windowwriter_cli.select_from_dict(numbered)

    root = tk.Tk()
    lb = MacroListbox(root, macro_dict)
    lb.connect_window(win_title)
    lb.pack()
    root.mainloop()


if __name__ == "__main__":
    main()
