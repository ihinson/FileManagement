import os
from glob import glob
from PIL import Image  
from pathlib import Path
from collections import defaultdict

# Define the folder containing the JPEG images
folder_path = "\\MUKGIS\GIS Files\scans\Reservoir 4 (Paine Field)\Paine-Field_Water-Main"
file_dict = {}
