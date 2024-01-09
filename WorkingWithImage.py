from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active
ws['A1']= "You should see tree logo below"

# create an image
img = Image('logo.png')

# add to worksheet and anchor next to cells
ws.add_image(img, 'A1')
wb.save('logo.xlsx')

