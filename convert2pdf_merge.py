from PIL import Image
import PyPDF2
import os
import glob

##USER INPUT
#Set working folder
working_folder = rf'G:\AsBuilts\Red\Water\2016 WSI\Pages'
output_folder = working_folder.strip('\Pages')
base = os.path.basename(os.path.normpath(working_folder))
output_file = output_folder + "/" + base + ".pdf"
pdf_set = os.path.basename(output_folder) + ".pdf"
print(pdf_set)

#Convert all files to pdfs

for filename in os.listdir(working_folder):
    if filename.endswith('.jpg'):
        img_path = os.path.join(working_folder, filename)
        img = Image.open(img_path)
        pdf_path = os.path.join(working_folder, filename[:-3] + 'pdf')
        img.save(pdf_path, 'PDF', resolution=100.0)
    else:
        pass
for filename in os.listdir(working_folder):
    if filename.endswith('.tiff'):
        img_path = os.path.join(working_folder, filename)
        img = Image.open(img_path)
        pdf_path = os.path.join(working_folder, filename[:-3] + 'pdf')
        img.save(pdf_path, 'PDF', resolution=100.0)
    else:
        pass
print("Conversion complete.")

# Merge pdfs
pdf_merger = PyPDF2.PdfFileMerger()
for filename in os.listdir(working_folder):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(working_folder, filename)
        pdf_merger.append(pdf_path)
pdf_merger.write(os.path.join(output_folder, pdf_set))
pdf_merger.close()

#Delete single pdfs
for filename in os.listdir(working_folder):
        if filename.endswith('.pdf'):
                pdf2delete = os.path.join(working_folder, filename)
                os.remove(pdf2delete)
        else:
                pass
print(pdf_set + " merge complete.")