source bin/activate
pyinstaller index.py -F \
--name "classrecorder-mac"
--add-binary='/System/Library/Frameworks/Tk.framework/Tk':'tk' \
--add-binary='/System/Library/Frameworks/Tcl.framework/Tcl':'tcl' \
--hidden-import xlsxwriter \
--hidden-import pynput \
--hidden-import xlrd \
--hidden-import datetime \
--hidden-import pyautogui \
--clean