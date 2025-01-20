# from photo_to_excel.photo_to_excel import PhotoToExcel
from photo_to_excel import PhotoToExcel

photos = [
    "tests/test_photos/neverforget.png",
    "tests/test_photos/bush_shoeing_incident.jpg",
    "tests/test_photos/baby Julia.png",
]

test_img = PhotoToExcel(photos[1], save_pixelated=True)
test_img.rgb_to_excel()