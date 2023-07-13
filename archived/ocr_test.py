from utility import *
from pdf2image import convert_from_path
import ocrmypdf

ocrmypdf.ocr("TFS.pdf", r"")
image = convert_from_path("TFS.pdf")
res= ocr_from_image(image)
print(res)

