.\Scripts\activate
pyinstaller index.py -F `
--name "classrecorder-windows"
--hidden-import xlsxwriter `
--hidden-import pynput `
--hidden-import xlrd `
--hidden-import datetime `
--hidden-import pyautogui `
--clean