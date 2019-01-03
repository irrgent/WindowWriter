# WindowWriter

A simple GUI for Windows that allows you to send text associated with a pre-defined macro to an open window.

## Prerequisites:

Make sure you have Python 3 installed on your system.

Install dependencies with:

```
pip install pywin32
```

## Usage:

Create a CSV file in the same directory as the program called `macros.csv`. 

The first column of this file should contain the shorthand for each of your macros that you want displayed in the GUI, while the second should be the expanded text you want sent to whatever window you are working on.

Afterwords, simply run with: 

```
python windowwriter_gui.py
```




