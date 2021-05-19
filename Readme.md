# PaintExcel

PaintExcel allow to transform any image into a "painted" excel workbook: Each cell of the excel workbook will be filled with a color as the pixels of the image.

 <img src="/GIT Images/demo1.jpg" width="49%" > <img src="/GIT Images/demo2.jpg" width="49%" >

 This tiny project is written in python

 ## Dependencies

 This python script use [Python Image Library (Pillow originally PIL)](https://pillow.readthedocs.io/en/stable/) to work with images and [openpyxel](https://openpyxl.readthedocs.io/en/stable/) to work with excel workbook.

 ## Limitations

 Excel is throwing a style formatting error at the opening of the workbook if the painted cells exceed certain dimension (between 300×300 and 400×400). Modify the variable `max_size` at the beginning of the file to change the number of painted cells.

 ## How to used

 Simply add all the `.jpg` of `.png` files you want to convert in the same directory as the `PaintExcel.py` file and run the script. An excel workbook will be created for each image.

 ## Licence

 This software is distributed under the MIT license. Enjoy!
