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

def find_vital_coordinates(entity_list):
        vital_coordinates = []
        for entity in entity_list:
            if entity.type == "AcDbAttributeDefinition" and entity.tag.startswith("H"):
                vital_coordinates.append((entity.pos, entity.tag))
        return vital_coordinates
    
def set_graph_id(coordinates, rectangles):
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1]
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return i, graph_id
    return None, None

def find_graph_id(coordinates, rectangles):
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1]
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return graph_id, rectangles[i]
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

    
def create_regions(coordinates, up, down, left, right):
        x, y = coordinates[0], coordinates[1]
        bounds = (x-left, x+right, y-down, y+up)
        return bounds
    
def generate_graph_rectangles(datapoints, width, height, w_correction, h_correction):
        rectangles = [((r1[0][0]-width+w_correction, r1[0][0]+w_correction, r1[0][1]+h_correction, r1[0][1]+height+h_correction), r1[1]) for r1 in datapoints]
        return rectangles
    
def get_center(coordinates):
        x, y = (coordinates[0][0]+coordinates[1][0]) /2 , (coordinates[0][1]+ coordinates[1][1]) /2
        return (x, y)
    
def generate_diagram_numbers(rectangles, tags):
    regions_with_diagrams = []
    for tag in tags:
        coordinates, tag = tag
        prefix, region = find_graph_id(coordinates, rectangles)
        if prefix[9] == "1":
            diagram_number = "FD-" + tag
        elif prefix[9] == "2":
            diagram_number = "SAMA-" + tag
        regions_with_diagrams.append((region[0], diagram_number))
    return regions_with_diagrams
            
def check_within(coordinates, rectangles, indice = None ):
    if indice == None:
        indice = [i for i in range(len(rectangles))]
    x,y = coordinates
    for i in range(len(rectangles)):
        rec = rectangles[i]
        x1, x2, y1, y2 = rec
        if x1 <= x <= x2 and y1 <= y <= y2:
            try:
                return indice[i], coordinates
            except:
                return indice, coordinates
    return None

def parameter_retriving(entity):
    """entity is a dxf entity, parameters is a list of parameters to be retrived"""
    relevant_attributes = ['ObjectName', 'InsertionPoint', 'coordinates', 'Area', 'Length', 'TextString', 'Handle', 'height', 'Layer', 'TagString', 'color']
    relevant_member_function = ['Coordinate', 'GetBoundingBox']
    res_dict = {}
    for function in relevant_member_function:
        try: 
            res = getattr(entity, function)()
            res_dict[function] = res
        except:
            res_dict[function] = None
    for attribute in relevant_attributes:
        try:
            attr = getattr(entity, attribute)
            res_dict[attribute] = attr
        except:
            res_dict[attribute] = None
    if res_dict['coordinates'] is not None:
        coordinates = []
        for i in range(len(res_dict['coordinates'])//2):
            coordinates.append(entity.Coordinate(i))
        res_dict['coordinates'] = coordinates
    return res_dict

