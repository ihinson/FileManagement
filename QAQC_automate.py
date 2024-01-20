# Import system modules
import arcpy
import pandas as pd
import os
import glob

#Workspace
grid = 'G:\Geodatabase\QAQC.gdb\QAQC_Progress'
asBuilts = 'G:\Geodatabase\Imagery.gdb\AsBulits_noAttachments'
drawings = 'G:\Geodatabase\Imagery.gdb\Drawings_noAttachments'

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

#Combine excels into one sheet
df1 = pd.read_excel(out_xls1)
df2 = pd.read_excel(out_xls2)

values1 = df1[['Link']]
values2 = df2[['Link']]

dataframes = [values1, values2]
join = pd.concat(dataframes)

mergedFile = 'G:/QAQC/' + workingPages + ".xlsx"
join.to_excel(mergedFile)

#TO DO
#bring in excel move script
#if file is already in folder skip