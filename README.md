# WindowWriter

A simple GUI for Windows that allows you to send text associated with a pre-defined macro to an open window.

## Prerequisites:

Make sure you have Python 3 installed on your system.

Install dependencies with:

```
pip install pywin32
```

## Usage:

Run using:

```
python windowwriter_gui.py
```

The program works by loading a csv file with two columns.

The first column of this file should contain the shorthand for each of your macros that you want displayed in the GUI, while the second should be the expanded text you want sent to whatever window you are working on.

Once you have added your macros to a csv file they can be loaded by clicking file -> open, and the window that you would like to send text to can be selected using the drop down menu at the bottom of the screen.





