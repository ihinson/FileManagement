# Import system modules
import arcpy
import pandas as pd
import os
import glob
from pathlib import Path
import shutil
from PIL import Image
import PyPDF2

#Part I: arcgis functions
#Workspace
grid = 'G:\Geodatabase\QAQC.gdb\QAQC_Progress'
asBuilts = 'G:\Geodatabase\Imagery.gdb\AsBulits_noAttachments'
drawings = 'G:\Geodatabase\Imagery.gdb\Drawings_noAttachments'

###USER INPUT###
#DEFINE THE PAGES TO SELECT 
workingPages = "pg7_8"
pages_sel = arcpy.management.SelectLayerByAttribute(
                in_layer_or_view=grid,
                selection_type="NEW_SELECTION",
                where_clause="Page = ' 7' Or Page = ' 8'",
                invert_where_clause=None
)

#Copy selection to new fc
grid = arcpy.management.CopyFeatures(pages_sel, 'G:/Geodatabase/QAQC.gdb/Grid')

#Select asBuilts and drawings from grid
asBuilt_sel = arcpy.management.SelectLayerByLocation(asBuilts, "HAVE_THEIR_CENTER_IN", grid)
drawings_sel = arcpy.management.SelectLayerByLocation(drawings, "HAVE_THEIR_CENTER_IN", grid)

#Copy the selected features
arcpy.management.CopyFeatures(asBuilt_sel, 'G:/Geodatabase/QAQC.gdb/AsBuilts')
arcpy.management.CopyFeatures(drawings_sel, 'G:/Geodatabase/QAQC.gdb/Drawings')

#Table to excel
in_table1 = 'G:/Geodatabase/QAQC.gdb/AsBuilts'
out_xls1 = "G:/QAQC/" + workingPages + "asBuilts.xlsx"
arcpy.conversion.TableToExcel(in_table1, out_xls1)

in_table2 = 'G:/Geodatabase/QAQC.gdb/Drawings'
out_xls2 = 'G:/QAQC/' + workingPages + "Drawings.xlsx"
arcpy.conversion.TableToExcel(in_table2, out_xls2)

#Part II: Merging Excels
#Combine excels into one sheet
df1 = pd.read_excel(out_xls1)
df2 = pd.read_excel(out_xls2)

values1 = df1[['Link']]
values2 = df2[['Link']]

dataframes = [values1, values2]
join = pd.concat(dataframes)

mergedFile = 'G:/QAQC/' + workingPages + ".xlsx"
join.to_excel(mergedFile)

#Cleanup
arcpy.env.workspace = r"G:\Geodatabase\QAQC.gdb"
cws = arcpy.env.workspace

fc_Delete = ["Grid","Asbuilts","Drawings"]

for fc in fc_Delete:

  fc_path = os.path.join(cws, fc)
  if arcpy.Exists(fc_path):
    arcpy.Delete_management(fc_path)

os.remove(out_xls1 )
os.remove(out_xls2)

print("Cleanup complete")

#Part III: copy Files to QAQC Folder

#DEFINE THE FILE LIST LOCATION
sheet = 'Sheet1'

#Create dataframe
df = pd.read_excel(mergedFile, sheet_name= sheet)

#Create list from dataframe
listFiles = df['Link'].tolist()

#Verify file list
#print(listFiles)

#Create folder to move files to
directory = rf'G:/QAQC/'
basename = os.path.basename(mergedFile)
fileBase = basename.strip(".xlsx")
working_folder = directory + fileBase
#print(working_folder)

Path(working_folder).mkdir(parents=True, exist_ok=True)

#Copy files
##if file is already in folder skip
for fileName in listFiles:
    try:
        shutil.copy(fileName, working_folder)
    except shutil.Error as err:
        pass
    except PermissionError:
        pass
    except FileExistsError:
        pass
    except OSError:
        pass
    except TypeError:
        pass

print("Copying complete")

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
    if filename.endswith(('.jpeg') OR ('.tiff')):
        img_path = os.path.join(working_folder, filename)
        img = Image.open(img_path)
        pdf_path = os.path.join(working_folder, filename[:-4] + 'pdf')
        img.save(pdf_path, 'PDF', resolution=100.0)
    else:
        pass