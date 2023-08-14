python -m PyInstaller ./main.py -n CC --windowed --onefile --icon ./Setup/logo-mission-2020.ico
mkdir .\dist\Setup\
copy .\Setup\logo-mission-2020.ico .\dist\Setup\logo-mission-2020.ico /Y
timeout 3600