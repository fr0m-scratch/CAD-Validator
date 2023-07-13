import ezdxf
from ezdxf.addons import odafc
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import cv2
import pytesseract
import matplotlib.pyplot as plt
from PIL import Image
import aspose.cad as cad


def generate_graph_rectangles(datapoints, width, height, w_correction, h_correction):
    rectangles = [((r1[0][0]-width+w_correction, r1[0][0]+w_correction, r1[0][1]+h_correction, r1[0][1]+height+h_correction), r1[1]) for r1 in datapoints]
    return rectangles

def set_graph_id(coordinates):
    rectangles =  [((14516.836040607368, 14936.836040607368, -2067.517680071807, -1770.5176800718068), 'HLY17TFS016E05045DD'), ((15955.756413856956, 16375.756413856956, -2067.517680071807, -1770.5176800718068), 'HLY17TFS016E05045DD'), ((14964.386723518339, 15384.386723518339, -2067.517680071807, -1770.5176800718068), 'HLY17TFS016E05045DD'), ((16439.979985211463, 16859.979985211463, -2067.517680071807, -1770.5176800718068), 'HLY17TFS016E05045DD'), ((15014.858999979917, 15434.858999979917, -3044.2277334884234, -2747.2277334884234), 'HLY17TFS026E05045DD'), ((15503.850813910469, 15923.850813910469, -3044.2277334884234, -2747.2277334884234), 'HLY17TFS026E05045DD'), ((15976.898566443171, 16396.89856644317, -3042.790969061651, -2745.790969061651), 'HLY17TFS026E05045DD'), ((15009.207440717302, 15429.207440717302, -3371.007519350775, -3074.007519350775), 'HLY17TFS026E05045DD'), ((15506.759262703286, 15926.759262703286, -3368.880558505289, -3071.880558505289), 'HLY17TFS026E05045DD'), ((15979.394427480038, 16399.394427480038, -3368.6279581031654, -3071.6279581031654), 'HLY17TFS026E05045DD'), ((15455.755832135936, 15875.755832135936, -2069.233752935705, -1772.233752935705), 'HLY17TFS016E05045DD'), ((16467.12678671578, 16887.12678671578, -3042.790969061651, -2745.790969061651), 'HLY17TFS026E05045DD'), ((16963.229449017614, 17383.229449017614, -3042.790969061651, -2745.790969061651), 'HLY17TFS026E05045DD'), ((17472.91146350187, 17892.91146350187, -3042.790969061651, -2745.790969061651), 'HLY17TFS026E05045DD'), ((14515.631717613047, 14935.631717613047, -3366.2343205941693, -3069.2343205941693), 'HLY17TFS026E05045DD'), ((14515.631717613047, 14935.631717613047, -2454.4379753074304, -2157.4379753074304), 'HLY17TFS016E05045DD'), ((14964.92238563453, 15384.92238563453, -2454.4379753074304, -2157.4379753074304), 'HLY17TFS016E05045DD'), ((15448.827003419403, 15868.827003419403, -2454.4379753074304, -2157.4379753074304), 'HLY17TFS016E05045DD'), ((18393.203455605595, 18813.203455605595, -3149.8260518229513, -2852.8260518229513), 'HLY17TFS026E05045DD'), ((18392.844932670738, 18812.844932670738, -2073.1582053069965, -1776.1582053069965), 'HLY17TFS016E05045DD'), ((14516.154647915657, 14936.154647915657, -3051.697802392868, -2754.697802392868), 'HLY17TFS026E05045DD')]
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1]
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return i, graph_id
    return None, None


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
    



def find_all_rectangle(file_path):
    """input file_path is a dxf file, return coordinates of all rectangles in the file"""
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()
    all_rectangles = []
    for entity in msp:
        if entity.dxftype() == "POLYLINE":
            if entity.is_closed:
                all_rectangles.append(entity)
    return all_rectangles