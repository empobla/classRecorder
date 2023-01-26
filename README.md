<!-- Improved compatibility of back to top link: See: https://github.com/othneildrew/Best-README-Template/pull/73 -->
<a name="readme-top"></a>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->
[![Contributors][contributors-shield]][contributors-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]



<!-- PROJECT LOGO -->
<br />
<div align="center">
<h3 align="center">classRecorder</h3>

  <p align="center">
    An OSX & Windows console-line application that allows users to automatically record their input Zoom classes.
    <br />
    <a href="https://github.com/empobla/classRecorder/issues">Report Bug</a>
    ·
    <a href="https://github.com/empobla/classRecorder/issues">Request Feature</a>
  </p>
</div>


[](#making-changes-to-the-class-recorder-driver-file)
<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li><a href="#abilities-mastered">Abilities Mastered</a></li>
    <li><a href="#dependency-list">Dependency List</a></li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li>
          <a href="#installation">Installation</a>
          <ul>
            <li><a href="#macos">MacOs</a></li>
            <li><a href="#windows">Windows</a></li>
          </ul>
        </li>
      </ul>
    </li>
    <li><a href="#making-changes-to-the-class-recorder-driver-file">Making Changes to the Class Recorder Driver File</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

A cross-platform (OSX & Windows) console-line application that performs user open class, start record, stop record, and close class functions automatically and records input Zoom classes automatically.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



### Built With

[![Python][Python]][Python-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ABILITIES MASTERED -->
## Abilities Mastered
* Use of Python to read and write to and from an excel file
* Usage of Python to run automated tasks with pyautogui
* Usage of Python to log mouse coordinates with pynput
* Manual implementation of mergesort algorithm
    * Object sorting with mergesort algorithm
* Usage of os listdir and path functions to list files within a relative directory
* Windows Python code bundling to an executable file with pyinstaller
* OSX Python code bundling to an executable file with pyinstaller
* Usage of .gitattributes file

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- DEPENDENCY LIST -->
## Dependency List
* Python v3.8.1
* pynput v1.6.8
* pyautogui
* xlsxwriter
* xlrd
* pyinstaller

The recommended screen recording software for using with this project is [screenrec](https://screenrec.com).

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running follow these simple example steps.

### Prerequisites

In order to run this project, all the dependencies must be installed properly beforehand. Installing Python with `pyenv` for usage with `pyinstaller` requires the following command in the installation:

```sh
env PYTHON_CONFIGURE_OPTS="--enable-framework" pyenv install 3.8.1
```

And then installing dependencies with `pip`.

### Installation

Guide for installation.

#### MacOS

1. Clone the repository
   ```sh
   git clone https://github.com/empobla/classRecorder.git
   ```
2. Run the `classrecorder-mac` file under `./dist` directory

If `classrecorder-mac` does not run because of a dependency issue or a version issue, the dependencies can be manually installed through pip, and then bundled and compiled with pyinstaller with the following command:

```sh
pyinstaller -F --name classrecorder-mac index.py
```

Alternatively, the program can be run through terminal with the usage of the following command (after installing the dependencies):

```sh
python index.py
```

<p align="right">(<a href="#readme-top">back to top</a>)</p>



#### Windows
1. Clone the repository
   ```sh
   git clone https://github.com/empobla/classRecorder.git
   ```
2. Run the `classrecorder-windows.exe` file under `./dist` directory

If `classrecorder-windows.exe` does not run because of a dependency issue or a version issue, the dependencies can manually be installed through pip, and then bundled and compiled with pyinstaller with the following command:

```sh
pyinstaller -F --name classrecorder-windows index.py
```

Alternatively, the program can be run through terminal with the usage of the following command (after installing the dependencies):

```sh
python index.py
```

#### Windows Automation Through Task Scheduler
To set up Windows Automation through Task Scheduler, edit the `runscript.bat` file under the `./winautomation` directory and add a path to both the `.exe` file within the `./dist` directory and `index.py`. After this, open Windows Task Scheduler, create a new task, schedule it to run whenever you need it to run, and target the `./winautomation/runscript.bat` file and run the task in the project's directory. 

Additionally, you can schedule another task to run the `./winautomation/sleep.bat` file to sleep your computer after automation has ended. 

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CHANGES TO CLASS RECORDER DRIVER FILE -->
## Making Changes to the Class Recorder Driver File
If you want to make changes to the class recorder driver file, you can simply follow one of these options:

### Option A
Delete or move the `.xlsx` file in the `./classrec` directory, and re-run the program to create a new class recorder driver.

### Option B
Edit the `.xlsx` file in the `./classrec` directory following the same format as the header-row specifies in each of the excel sheets within the `.xlsx` file. 

_Note: This program will only read 1 `.xlsx` file within the `./classrec` directory. If more than one file exists within that directory, it will pick one and run it. Make sure to only have 0 or 1 `.xlsx` file within the `./classrec` directory at a time to ensure the program works the way you would expect it to work._

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Emilio Popovits Blake - [Contact](https://emilioppv.com/contact)

Project Link: [https://github.com/empobla/classRecorder](https://github.com/empobla/classRecorder)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS -->
## Acknowledgments

* [Screenrec](https://screenrec.com)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/empobla/classRecorder.svg?style=for-the-badge
[contributors-url]: https://github.com/empobla/classRecorder/graphs/contributors
[issues-shield]: https://img.shields.io/github/issues/empobla/classRecorder.svg?style=for-the-badge
[issues-url]: https://github.com/empobla/classRecorder/issues
[license-shield]: https://img.shields.io/github/license/empobla/classRecorder.svg?style=for-the-badge
[license-url]: https://github.com/empobla/classRecorder/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/emilio-popovits


[Python]: https://img.shields.io/badge/python-3776ab?style=for-the-badge&logo=python&logoColor=ffdc52
[Python-url]: https://www.python.org