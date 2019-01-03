import tkinter as tk
import windowwriter
import win32com.client as comclt


class MenuBar(tk.Menu):
    """
    Instance of a tkinter Menu to be used as
    the top drop down menu in an instance of
    MainApplication.
    """

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
    """
    tkinter Listbox that is populated with macros
    loaded by the user. Additional functions allow
    for sending keyboard input to a window.
    """

    def __init__(self, parent, macro_dict, *args, **kwargs):
        """
        parent - The parent tk.Frame instance

        macro_dict - Dictionary containing macros.
        Keys = macro and values = expansion of the
        macro.
        """

        tk.Listbox.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict
        self.win_title = None
        self.bind("<<ListboxSelect>>", self.macro_select)

        # Insert each of the macro keys into the listbox.
        for _ in self._macro_dict.keys():
            self.insert(tk.END, _)

    # Will allow for reloading the macros (not yet implemented)
    def update_macros(self, new_dict):
        raise NotImplementedError

    # 'Connect' to a window meaning set the title of the target
    # window and create a WScript.Shell instance.
    def connect_window(self, title):
        self.win_title = title
        self._wsh = comclt.Dispatch("WScript.Shell")

    def disconnect_window(self):
        self.win_title = None

    def macro_select(self, event):

        # Prompt user to select window if none selected.
        if self.win_title is None:

            win = tk.Toplevel()
            win.wm_attributes("-topmost", 1)
            win.wm_title("No window selected.")
            win_text = "Please select a window using the drop down menu."
            win_label = tk.Label(win, text=win_text)
            win_button = tk.Button(win, text='OK', command=win.destroy)

            win_label.pack(side=tk.TOP)
            win_button.pack(side=tk.BOTTOM)

        else:

            # Get selected Listbox item.
            wdg = event.widget
            idx = int(wdg.curselection()[0])
            value = wdg.get(idx)

            # Send selected macros associated text to the
            # selected window.
            windowwriter.send_input(self._wsh, self.win_title,
                                    self._macro_dict[value])

            print("Selected {}.".format(value))


class MainApplication(tk.Tk):
    """
    Manages all frames, widgets, menus etc for the application. A
    single MainApplication instance should be created first and
    all other widgets added within.
    """

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

    # Create a tk.OptionMenu instance containing a list of currently open
    # windows that the user can then select and send text to.
    def create_win_menu(self):

        self.options = windowwriter.get_window_names()
        self.selected_var = tk.StringVar()

        # Default selection, will dissapear once another option is selected.
        self.selected_var.set("Select window:")

        # Making a selection modifies selected_var and calls select_win.
        self.selected_var.trace_id = self.selected_var.trace("w", self.select_win)

        self.win_menu = tk.OptionMenu(
            self.frame, self.selected_var, *self.options)

    def select_win(self, *args):
        self.list_box.connect_window(self.selected_var.get())

    # Refresh window list by removing old win_menu and creating a new one
    # that will contain any newly opened windows.
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
