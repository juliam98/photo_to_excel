# photo_to_excel
This code allows you to convert your photo to excel format. 

## The Photo-to-Excel converter works in the following way:
1. The photo you specify is first pixelated into 64x64 pixel format. (or any other size you specify)
2. Colours are extracted from each of the pixels.
3. Excel workbook is created, and the extracted colours are used to colour a grid of 64x64 cells in excel. 

<p align="center">
  <img src="tests/test_photos/output_example.png"  width="600">
</p>

## Instructions:
Create a `PhotoToExcel` object with the image path as the inpput. Then call the `rgb_to_excel` method to convert the photo to excel. The xlsx file will be saved in the same direcory and have the same name as the original image provided.

### Example: 
If you wanted to comemorate the Bush shoeing incident by converting the photo to excel, you would do the following:
```python
from photo_to_excel import PhotoToExcel

test_img_path = "tests/test_photos/bush_shoeing_incident.jpg"

test_img = PhotoToExcel(test_img_path, save_pixelated=True)
test_img.rgb_to_excel()
```

By selecting `save_pixelated=True`, you can save the pixelated image that is used in the process of creating the excel file.

And here is your finished result:

<p align="center">
  <img src="tests/test_photos/output_example2.png"  width="600">
</p>

Future releases will include adjustment for image dimention ratio, for now, you either have to calculate if yourself or use the default 64x64 pixel format and your photos will be a bit distorted. 

## Installation
You can install the package using pip:
```bash
pip install photo-to-excel
```

## Issues
If you encounter any issues, please let me know by creating an issue on this repository. Sugestions are welcome too.

## The purpose of the project:
You might ask 'why would I convert my photos to Excel?'. My response is: 'why not?'.
