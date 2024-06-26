# PnP2PDF
What is it? 

PnP2PDF (PnP to PDF) is a software that can help you to prepate your PnP cards for printing. The program helps you to assemble a PDF file from images of individual cards, double-sided cards are also supported.

# Download and Setup
[MacOS Version](https://drive.google.com/drive/folders/1fADjKJbXosjtHfRhAyEQ-Opt4FbnRdQ7?usp=sharing)

[Windows Version](https://drive.google.com/drive/folders/1BOEETxM4q-EuxWY9ZJXSHYg7T8a1HAd-?usp=sharing)

OR

1. [Install Python](https://www.python.org/downloads/)
2. [Install PIP](https://pip.pypa.io/en/stable/installation/)
3. Run `pip install PySimpleGUI` in terminal/command line
4. Run `pip install pillow` in terminal/command line
5. Open `PnP2PDF.py` file
   
# Program Interface:
<img width="584" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/76cb986b-e322-4fae-ba88-3925bf03d61b">

# Result:
<img width="741" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/5b78db15-d9a8-49b6-8c18-ecdcf37a425f">

# How to Prepare Double Sided Cards
1. Create folder.
2. Add all card backs you will use to this folder (If you have back template with size bigger than card you may name it like that `Card Back Name_64,5_91,3.png`, i.e. add `_{back width in mm}_{back height in mm}` in name of file, use a comma to separate the integer and fractional parts of the width/height).
3. Create folders with the same names like card backs from previous step (except you dont need add `_{back width in mm}_{back height in mm}` at the end of folder name) and place all cards that have corresponding back.
4. You can also create folder with name `Double Sided` for double sided cards that have no ordinary back, where you should place all sides together. Each pair of files will be printed as one double sided card. Be sure if files are sorted in alphabetical order.
5. Finally you can add folder named `Single Sided` with cards that don't have back and place your cards here.
6. That is all! Now choose the folder from step 1 as `Cards folder` in `PnP2PDF` window and enjoy your final result!

The final structure is shown below:

<img width="334" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/11636a16-23a6-4027-b687-1028d582da59">

P.S. You can add as much card backs and their folders as you want, not only two pair as in example.

# Known Issues
* Windows version shows error after closing a program (doesn't affect work).
* Windows version may cause Microsoft Defender alerts.
* Windows version may not draw cutting lines.





