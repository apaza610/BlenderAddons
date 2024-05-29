#aqui actualizamos la version de blender a la ultima descargada..
import zipfile
import shutil
import os
import sys

def unzip_and_move(src_zip, dest_dir):
	if not os.path.exists(dest_dir):
		os.makedirs(dest_dir)

	shutil.rmtree(f"{dest_dir}\\blender")

	with zipfile.ZipFile(src_zip, 'r') as zip_ref:
		zip_ref.extractall(dest_dir)

	temporal = src_zip.split('\\')[-1].replace('.zip','')

	os.rename(f"{dest_dir}\\{temporal}", f"{dest_dir}\\blender")
	print(f"Unzipped files were moved to... {dest_dir}")

os.chdir(r"E:\win\Downloads")
files = [f for f in os.listdir() if (f.startswith('blender') and f.endswith('.zip'))]
files.sort(key=os.path.getmtime)

# src_zip = input("path of new Blender: ")	#'E:\\win\\Downloads\\blender-4.1.0-beta+v41.a677518cb50a-windows.amd64-release.zip'
src_zip = files[-1]		#descomprimiendo el ultimo de los zip (mas nuevo) de la lista

if os.path.exists(src_zip):
	print(f"desComprimiendo: {src_zip}")

unzip_and_move(src_zip ,'E:\\apps\\art3d')