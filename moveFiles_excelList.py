import os.path
import pandas as pd
from pathlib import Path
import shutil


#DEFINE THE FILE LIST LOCATION
excel_file = 'G:/QAQC/QAQC_2_6.xlsx'
sheet = 'Use'

#Create dataframe
df = pd.read_excel(excel_file, sheet_name= sheet)

#Create list from dataframe
listFiles = df['Link'].tolist()

#Verify file list
#print(listFiles)

#Create folder to move files to
directory = rf'G:/QAQC/'
basename = os.path.basename(excel_file)
fileBase = basename.strip(".xlsx")
working_folder = directory + fileBase
#print(working_folder)

Path(working_folder).mkdir(parents=True, exist_ok=True)

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