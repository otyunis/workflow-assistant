## Workflow Assistant
*Workflow Assistant* contains functionality for automating repetitive workflow tasks. At the moment, this inclcudes populating folder directories, unzipping password-protected *.zip* files (using 7-Zip), and editing *.docx* files (using a custom class, EditDocx, based on the *python-docx* library) -- editing functionality currently includes parsing tables/paragraphs, inserting/replacing text, inserting images, and finding/toggling checkboxes. *Workflow Assistant* contains a custom standalone GUI application using python's native tkinter library to automate an example of a specifc workflow task (*application.py*) -- the application can be easily extended to include additional functionality and accomodate different workflow tasks. The application's dependencies are included in *requirements.txt.*

## Requirements
- Python 3.11.1 ([Download Here](https://www.python.org/downloads/))
- 7-Zip ([Download Here](https://www.7-zip.org/download.html))
- OS: Windows 10

## Installation
1. Download this repo
2. Edit *config.json* *(update username, paths, and keys to reflect your project install)*
3. Edit python path in application.bat *(update path to reflect your python 3.11.1 install)*
### (optional)
4. Create .vbs shortcut *(right-click on application.vbs > create shortcut)*
5. Add icon to .vbs shortcut *(right-click on application.vbs - shortcut > properties > shortcut tab > change icon > browse > select icon.ico from application folder)*
6. change shortcut target *(rightl-click on application.vbs - shortcut > properties > shortcut tab > prepend wscript to target  -- see [Stack Overflow Answer](https://stackoverflow.com/questions/19318416/pin-a-shortcut-of-a-vbs-script-to-the-taskbar-w2008-server))*
7. pin shortcut to taskbar *(drag and drop application.vbs to taskbar)*

## Running the application
- Run the application by running *application.vbs* *(double-click on application.vbs or application.vbs  - shortcut)*
- *application.vbs* executes *application.bat* in a windowless shell -- *application.bat* automates the process of setting up the virtual environment needed to run *application.py* and then runs it with the appropriate python version.
- The first time running the application may take longer due to installation of project dependencies.
