# Hoffix COM parser
Dataparser based on COM lib to Excel files
# run
## Python3
```bash
python main.py "path_to_file" 
```
## EXE
```bash
filename "path_to_file"
```
# Pyinstaller bundle
!!! venv must be activated !!!
```bash
pyinstaller -y --onefile --collect-all "pywin32" --hidden-import "pywintypes" --hidden-import "playwright" main.py
```