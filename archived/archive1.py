import ezdxf
from ezdxf.addons import odafc
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import cv2
import pytesseract
import matplotlib.pyplot as plt
from PIL import Image
import aspose.cad as cad

def ocr_from_image(image_path, line_items_coordinates = None):
    image = cv2.imread(image_path)

    img = image
    if line_items_coordinates:
        # get co-ordinates to crop the image
        c = line_items_coordinates[1]

        # cropping image img = image[y0:y1, x0:x1]
        img = image[c[0][1]:c[1][1], c[0][0]:c[1][0]]    

    plt.figure(figsize=(10,10))
    plt.imshow(img)

    # convert the image to black and white for better OCR
    ret,thresh1 = cv2.threshold(img,120,255,cv2.THRESH_BINARY)

    # pytesseract image to string to get results
    text = str(pytesseract.image_to_string(thresh1, config='--psm 6'))
    return text

def dwg_to_pdf(filepath):
    # Load an existing DWG file
    image = cad.Image.load(filepath)

    # Initialize and specify CAD options
    rasterizationOptions = cad.imageoptions.CadRasterizationOptions()
    rasterizationOptions.page_width = 1200
    rasterizationOptions.page_height = 1200
    rasterizationOptions.layouts = ["Layout1"]

    # Specify PDF Options
    pdfOptions = cad.imageoptions.PdfOptions()
    pdfOptions.vector_rasterization_options = rasterizationOptions
    imagepath = filepath.replace(".dwg", ".pdf")
    # Save as PDF
    image.save(imagepath, pdfOptions)
    
    
def dwg_to_dxf(file_path):
    doc_path = file_path.replace("dwg", "dxf")
    if not doc_path:
        odafc.convert(file_path, doc_path)
    else:
        print("file already exists")
    return doc_path

def cad_to_image(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()
    fig,ax = plt.subplots()
    context = RenderContext(doc)
    out = MatplotlibBackend(ax)
    Frontend(context, out).draw_layout(msp, finalize=True)
    image_path = file_path.replace(".dxf", ".png")
    fig.savefig(image_path, dpi=300)
    return image_path


def convert_dxf_to_jpeg(dxf_file, jpeg_file):
    return

def mark_region(image_path):
    im = cv2.imread(image_path)

    gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (9,9), 0)
    thresh = cv2.adaptiveThreshold(blur,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV,11,30)

    # Dilate to combine adjacent text contours
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9,9))
    dilate = cv2.dilate(thresh, kernel, iterations=4)

    # Find contours, highlight text areas, and extract ROIs
    cnts = cv2.findContours(dilate, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    line_items_coordinates = []
    for c in cnts:
        area = cv2.contourArea(c)
        x,y,w,h = cv2.boundingRect(c)
        if y >= 600 and x <= 1000:
            if area > 10000:
                image = cv2.rectangle(im, (x,y), (2200, y+h), color=(255,0,255), thickness=3)
                line_items_coordinates.append([(x,y), (2200, y+h)])
        elif y >= 2400 and x<= 2000:
            image = cv2.rectangle(im, (x,y), (2200, y+h), color=(255,0,255), thickness=3)
            line_items_coordinates.append([(x,y), (2200, y+h)])
        else:
            image = im
    return image, line_items_coordinates
