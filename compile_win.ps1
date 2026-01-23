
pip install pyinstaller
pip install pillow

pyinstaller `
    --add-data "icon.png:." `
    --icon=icon.png `
    --onefile `
    --name PnP2PDF `
    --windowed main.py
