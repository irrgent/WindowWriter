import tkinter as tk
import windowwriter
import win32com.client as comclt


class MenuBar(tk.Menu):

    def __init__(self, parent, *args, **kwargs):

        tk.Menu.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.file_menu = tk.Menu(self)
        self.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Exit", command=parent.quit)
        self.file_menu.add_command(
            label="Refresh windows", command=parent.refresh_windows)

        self.edit_menu = tk.Menu(self)
        self.add_cascade(label="Edit", menu=self.edit_menu)


class MacroListbox(tk.Listbox):

    def __init__(self, parent, macro_dict, *args, **kwargs):

        tk.Listbox.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict
        self.win_title = None
        self.bind("<<ListboxSelect>>", self.macro_select)

        for _ in self._macro_dict.keys():
            self.insert(tk.END, _)

    def update_macros(self, new_dict):
        raise NotImplementedError

    def connect_window(self, title):
        self.win_title = title
        self._wsh = comclt.Dispatch("WScript.Shell")

    def disconnect_window(self):
        self.win_title = None

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


class MainApplication(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.top_menu = MenuBar(self)
        self.config(menu=self.top_menu)

        self.frame = tk.Frame(self)
        # TODO - allow different paths for macros.csv
        self.list_box = MacroListbox(
            self.frame,
            windowwriter.macro_dict("./macros.csv"), selectmode=tk.SINGLE)
        self.create_win_menu()

        self.frame.pack()
        self.list_box.pack(side=tk.TOP)
        self.win_menu.pack(side=tk.TOP)

    def create_win_menu(self):

        self.options = windowwriter.get_window_names()
        self.selected_var = tk.StringVar()
        self.selected_var.set("Select window:")
        self.selected_var.trace_id = self.selected_var.trace("w", self.select_win)
        self.win_menu = tk.OptionMenu(
            self.frame, self.selected_var, *self.options)

    def select_win(self, *args):
        self.list_box.connect_window(self.selected_var.get())

    def refresh_windows(self):

        self.win_menu.pack_forget()
        self.list_box.disconnect_window()
        self.create_win_menu()
        self.win_menu.pack(side=tk.TOP)


def main():

    root = MainApplication()
    root.wm_attributes("-topmost", 1)

    root.mainloop()


if __name__ == "__main__":
    main()
