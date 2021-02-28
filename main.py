import xlsxwriter
import os
from PIL import Image
# all you need to change, cell size and photo directory
directory = 'res/image/'
cellWidth = 60
cellHeight = 250

cellWidthToPixelConversionRatio = 355 / 50
cellHeightToPixelConversionRatio = 400 / 300
cellWidthInPixels = cellWidth * cellWidthToPixelConversionRatio
cellHeightInPixels = cellHeight * cellHeightToPixelConversionRatio
defaultImageDpi = 96

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('images.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_default_row(cellHeight)
worksheet.set_column(0, 5, cellWidth)


# Insert an image.

index = 1
for filename in os.listdir(directory):
    imagePath = directory + '/' + filename
    print("Inserting image " + imagePath)
    image = Image.open(imagePath)
    image.save(imagePath, dpi=(defaultImageDpi,defaultImageDpi))
    imageWidth, imageHeight = image.size
    widthScale = (cellWidthInPixels / imageWidth)
    heightScale = (cellHeightInPixels / imageHeight)
    xOffset = 0
    yOffset = 0
    scale = min(widthScale, heightScale)
    if widthScale <= heightScale:
        xOffset = 0
        yOffset = (cellHeightInPixels - imageHeight * scale) / 2
    else:
        xOffset = (cellWidthInPixels - imageWidth * scale) / 2
        yOffset = 0
    worksheet.insert_image('B' + str(index), imagePath,
                           {'object_position': 1, 'x_offset': xOffset, 'y_offset': yOffset, 'x_scale': scale,
                            'y_scale': scale})
    index += 1

workbook.close()
print("done")
