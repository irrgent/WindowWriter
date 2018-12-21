import tkinter as tk
import windowwriter
import windowwriter_cli
import win32com.client as comclt


class MenuBar(tk.Menu):

    def __init__(self, parent, *args, **kwargs):

        tk.Menu.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.file_menu = tk.Menu(self)
        self.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Exit", command=parent.quit)

        self.edit_menu = tk.Menu(self)
        self.add_cascade(label="Edit", menu=self.edit_menu)


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


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)

        # TODO - allow different paths for macros.csv
        self.list_box = MacroListbox(
            self, windowwriter.macro_dict("./macros.csv"))

        self.list_box.pack(side=tk.TOP)
        self.cli_add_windows()

    def cli_add_windows(self):

        windows = windowwriter.get_windows()
        numbered = windowwriter_cli.numbered_dict_keys(windows)
        win_title = windowwriter_cli.select_from_dict(numbered)

        self.list_box.connect_window(win_title)


def main():

    root = tk.Tk()
    root.wm_attributes("-topmost", 1)
    main_app = MainApplication(root)
    app_menu = MenuBar(root)
    root.config(menu=app_menu)
    main_app.pack()
    root.mainloop()


if __name__ == "__main__":
    main()
