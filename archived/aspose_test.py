import aspose.cad as cad 

cadImage = cad.Image.load("TFS.dwg")

rasterizationOptions = cad.imageoptions.CadRasterizationOptions()

rasterizationOptions.page_width = 1600
rasterizationOptions.page_height = 1600

options = cad.imageoptions.JpegOptions()

options.vector_rasterization_options = rasterizationOptions

cadImage.save("TFS.jpeg", options)
print("Done")


