import tkinter as tk
from tkinter import filedialog
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

        self.file_menu.add_command(label="Open", command=parent.open_macros)

        self.edit_menu = tk.Menu(self)
        self.add_cascade(label="Edit", menu=self.edit_menu)


class MacroListbox(tk.Listbox):
    """
    tkinter Listbox that is populated with macros
    loaded by the user. Additional functions allow
    for sending keyboard input to a window.
    """

    def __init__(self, parent, macro_dict=None, *args, **kwargs):
        """
        parent - The parent tk.Frame instance

        macro_dict - Dictionary containing macros.
        Keys = macro and values = expansion of the
        macro.
        """

        tk.Listbox.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict

        # win_info will be updated by connect_window() which is called
        # when the user selects a window from the menu created in MainApplication
        self.win_info = None
        self.bind("<<ListboxSelect>>", self.macro_select)

        # Insert each of the macro keys into the listbox if a dictionary
        # was provided.
        if self._macro_dict is not None:
            for _ in self._macro_dict:
                self.insert(tk.END, _)
        else:
            self.insert(tk.END, "To load macros use")
            self.insert(tk.END, "file -> open")

    # Repopulate listbox with new macros.
    def update_macros(self, new_dict):

        self.delete(0, self.size())
        self._macro_dict = new_dict

        for _ in new_dict:
            self.insert(tk.END, _)

    # 'Connect' to a window meaning set the title of the target
    # window and create a WScript.Shell
    def connect_window(self, title, hwnd):
        self.win_info = (title, hwnd)
        self._wsh = comclt.Dispatch("WScript.Shell")

    def disconnect_window(self):
        self.win_info = None

    def macro_select(self, event):

        if self._macro_dict is None:
            ErrorPopup("No macros loaded.", "Error")

        elif self.win_info is None:
            win_text = "Please select a window using the drop down menu."
            ErrorPopup(win_text, "No window selected.")

        else:

            # Get selected Listbox item.
            wdg = event.widget
            idx = int(wdg.curselection()[0])
            value = wdg.get(idx)

            # Send selected macros associated text to the
            # selected window.
            windowwriter.send_input(self._wsh, self.win_info[0],
                                    self.win_info[1],
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

        self.list_scroll = tk.Scrollbar(self.frame)
        self.list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.list_box = MacroListbox(
            self.frame,
            selectmode=tk.SINGLE,
            yscrollcommand=self.list_scroll.set)

        self.list_scroll.config(command=self.list_box.yview)

        self.create_win_menu()

        self.frame.pack()
        self.list_box.pack(side=tk.TOP)
        self.win_menu.pack(side=tk.TOP)

    # Create a tk.OptionMenu instance containing a list of currently open
    # windows that the user can then select and send text to.
    def create_win_menu(self):

        self.options = windowwriter.get_windows()
        self.selected_var = tk.StringVar()

        # Default selection, will dissapear once another option is selected.
        self.selected_var.set("Select window:")

        # Making a selection modifies selected_var and calls select_win.
        self.selected_var.trace_id = self.selected_var.trace("w", self.select_win)

        self.win_menu = tk.OptionMenu(
            self.frame, self.selected_var, *list(self.options.keys()))

    def select_win(self, *args):
        title = self.selected_var.get()
        self.list_box.connect_window(title, self.options[title])

    # Refresh window list by removing old win_menu and creating a new one
    # that will contain any newly opened windows.
    def refresh_windows(self):

        self.win_menu.pack_forget()
        self.list_box.disconnect_window()
        self.create_win_menu()
        self.win_menu.pack(side=tk.TOP)

    # Open a new csv file containing macros, called from menu.
    def open_macros(self):
        name = filedialog.askopenfilename(
            title="Open file",
            filetypes=(("csv files", "*.csv"), ("all files", "*.*")))

        try:
            new_macros = windowwriter.macro_dict(name)
        except ValueError as e:

            ErrorPopup(e, "Error opening file.")
            return

        self.list_box.update_macros(new_macros)


class ErrorPopup(tk.Toplevel):

    def __init__(self, message, title, *args, **kwargs):
        tk.Toplevel.__init__(self, *args, **kwargs)

        self.wm_attributes("-topmost", 1)
        self.wm_title(title)
        self.lab = tk.Label(self, text=message)
        self.button = tk.Button(self, text='OK', command=self.destroy)

        self.lab.pack(side=tk.TOP)
        self.button.pack(side=tk.BOTTOM)


def main():

    root = MainApplication()
    root.wm_attributes("-topmost", 1)
    root.mainloop()


if __name__ == "__main__":
    main()
