# Class Recorder
A cross-platform (OSX & Windows) console-line application that performs user open class, start record, stop record, and close class functions automatically and records input Zoom classes automatically.

## What I Learned
- Use of Python to read and write to and from an excel file
- Usage of Python to run automated tasks with pyautogui
- Usage of Python to log mouse coordinates with pynput
- Manual implementation of mergesort algorithm
    - Object sorting with mergesort algorithm
- Usage of os listdir and path functions to list files within a relative directory
- Windows Python code bundling to an executable file with pyinstaller
- OSX Python code bundling to an executable file with pyinstaller
- Usage of .gitattributes file

## Dependencies
- Python v3.8.1
- pynput v1.6.8
- pyautogui
- xlsxwriter
- xlrd
- pyinstaller

The recommended screen recording software for using with this project is [screenrec](https://screenrec.com).

## Installation and Usage
Guide for installation and usage.

### Mac OS
1. Clone the repository
2. Run the `classrecorder-mac` file under `./dist` directory

If `classrecorder-mac` does not run because of a dependency issue or a version issue, the dependencies can be manually installed through pip, and then bundled and compiled with pyinstaller with the following command:

```sh
pyinstaller -F --name classrecorder-mac index.py
```

Alternatively, the program can be run through terminal with the usage of the following command (after installing the dependencies):

```sh
python index.py
```

### Windows
1. Clone the repository
2. Run the `classrecorder-windows.exe` file under `./dist` directory

If `classrecorder-windows.exe` does not run because of a dependency issue or a version issue, the dependencies can manually be installed through pip, and then bundled and compiled with pyinstaller with the following command:

```sh
pyinstaller -F --name classrecorder-windows index.py
```

Alternatively, the program can be run through terminal with the usage of the following command (after installing the dependencies):

```sh
python index.py
```

### Windows Automation Through Task Scheduler
To set up Windows Automation through Task Scheduler, edit the `runscript.bat` file under the `./winautomation` directory and add a path to both the .exe file within the `./dist` directory and index.py. After this, open Windows Task Scheduler, create a new task, schedule it to run whenever you need it to run, and target the `./winautomation/runscript.bat` file and run the task in the project's directory. 

Additionally, you can schedule another task to run the `./winautomation/sleep.bat` file to sleep your computer after automation has ended. 

## Making Changes to the Class Recorder Driver File
If you want to make changes to the class recorder driver file, you can simply follow one of these options:

### Option A
Delete or move the .xlsx file in the `./classrec` directory, and re-run the program to create a new class recorder driver.

### Option B
Edit the .xlsx file in the `./classrec` directory following the same format as the header-row specifies in each of the excel sheets within the .xlsx file. 

---

**Note:** This program will only read 1 .xlsx file within the `./classrec` directory. If more than one file exists within that directory, it will pick one and run it. Make sure to only have 0 or 1 .xlsx file within the `./classrec` directory at a time to ensure the program works the way you would expect it to work.

**Note:** Installing Python with **pyenv** for usage with pyinstaller requires the following command in the installation:

```sh
env PYTHON_CONFIGURE_OPTS="--enable-framework" pyenv install 3.7.6
```

And then installing dependencies with pip.