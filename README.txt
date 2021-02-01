# img2xlsx
Converts an image file to an Excel spreadsheet of colored cells that correspond to the pixels of the image

## Instructions
Place an image in the img folder... If there is no image in the img folder, the function will not work... If there is more than one image, the function will use the first image
After the function is done, output.xlsx can be found in the xlsx folder

## Setup
Install pip (Tested with python 3.8.2 and pip 20.0.2 on Ubuntu 20.04)
Using pip, install openpyxl

## Misc
This function reduces the colors of the original image from 24bit to 15bit, this is done to avoid Excel's 64000 unique styles limitation
