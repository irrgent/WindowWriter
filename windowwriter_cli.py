import windowwriter
import sys


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


def show_dict(show, sorted=False):

    dict_keys = list(show.keys())

    if sorted:
        dict_keys = sorted(dict_keys)

    for k in dict_keys:
        print("{:15}:\t{}".format(k, show[k]))


def main():

    help_msgs = {"help": "Display this message.",
                 "window": "Select a target window.",
                 "macros": "List all macros.",
                 "refresh": "Update open window list.",
                 "quit": "Exit the program."}

    help_str = """
    Usage: Enter a macro to send it to the currently selected window,
    or use one of the following commands.
    """

    quitted = False

    macros = windowwriter.macro_dict(sys.argv[1])
    macro_keys = list(macros.keys())

    windows = windowwriter.get_windows()
    current_window = None

    while not quitted:

        cmd = input("> ")

        if cmd in macro_keys:
            raise NotImplementedError

        elif cmd == "help":
            print(help_str)
            show_dict(help_msgs)

        elif cmd == "window":
            raise NotImplementedError

        elif cmd == "macros":
            show_dict(macros)

        elif cmd == "refresh":
            windows = windowwriter.get_windows()
            print("Windows list updated.")

        elif cmd == "quit":
            quitted = True

        else:
            print("Invalid command. Try 'help' for a list of commands.")


if __name__ == "__main__":

    main()
