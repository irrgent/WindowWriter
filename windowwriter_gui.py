import tkinter as tk
import windowwriter


class MacroListbox(tk.Frame):

    def __init__(self, parent, macro_dict, *args, **kwargs):

        tk.Frame.__init__(self, parent, *args, **kwargs)

        self._macro_dict = macro_dict
        self._listbox = tk.Listbox(self, selectmode=tk.SINGLE)
        self._listbox.bind("<<ListboxSelect>>", self.macro_select)
        self._listbox.pack()

        for _ in self._macro_dict.keys():
            self._listbox.insert(tk.END, _)

    def update_macros(self, new_dict):
        raise NotImplementedError

    def macro_select(self, event):
        wdg = event.widget

        idx = int(wdg.curselection()[0])
        value = wdg.get(idx)

        print("Selected {}.".format(value))


def main():

    macro_dict = windowwriter.macro_dict("./macros.csv")

    root = tk.Tk()
    MacroListbox(root, macro_dict).pack()
    root.mainloop()


if __name__ == "__main__":
    main()
