import zipfile
import shutil
import os
import sys

folderDwnld = "E:\\win\\Downloads"
folderDstno = "C:\\orozco\\"
folderFinal = f"{folderDstno}\\blender-4.5"

os.chdir(folderDwnld)

def unzip_and_move(src_zip):
	if os.path.exists(folderFinal):
		shutil.rmtree(folderFinal)		# borrar lo viejo

	with zipfile.ZipFile(src_zip, 'r') as zip_ref:
		zip_ref.extractall(folderDstno)

	# temporal = src_zip.split('\\')[-1].replace('.zip','')
	temporal = f"{folderDstno}\\{src_zip}"		# E:\\apps\\art3d\\blender-4.4.0-beta+v44.b39c637d63ec-windows.amd64-release.zip
	temporal = temporal.replace('.zip','')

	os.rename(temporal, folderFinal)
	print(f"Unzipped files fueron movidos a {folderDstno}")

##########################################################################################################
files = [f for f in os.listdir() if (f.startswith('blender') and f.endswith('.zip'))]
files.sort(key=os.path.getmtime, reverse=False)

# src_zip = input("path of new Blender: ")	#'E:\\win\\Downloads\\blender-4.1.0-beta+v41.a677518cb50a-windows.amd64-release.zip'
src_zip = files[-1]		#descomprimiendo el ultimo (mas nuevo) de la lista

if os.path.exists(src_zip):
	print(f"desComprimiendo: {src_zip}")

unzip_and_move(src_zip)

for zip in files:
	if zip != src_zip:
		os.remove(zip)		#borrar los otros zips