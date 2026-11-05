"""Reproducible Windows packaging. Run from this directory with a Python venv."""
import subprocess
import sys
from pathlib import Path
from PIL import Image, ImageDraw

ROOT=Path(__file__).parent.resolve()
OUT=ROOT/'dist'
OUT.mkdir(exist_ok=True)
icon=Image.new('RGBA',(256,256),'#15283F');draw=ImageDraw.Draw(icon)
draw.rounded_rectangle((6,6,250,250),radius=46,fill='#15283F')
for x,width in [(43,9),(61,17),(88,8),(106,5),(120,17),(147,8),(165,13),(188,7),(204,7)]:
    draw.rectangle((x,66,x+width,182),fill='#B9D0FF')
draw.rectangle((42,195,212,204),fill='#FFFFFF')
icon.save(ROOT/'corrugados.ico',sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
version="""VSVersionInfo(ffi=FixedFileInfo(filevers=(1,0,0,0),prodvers=(1,0,0,0),mask=0x3f,flags=0x0,OS=0x40004,fileType=0x1,subtype=0x0,date=(0,0)),kids=[StringFileInfo([StringTable('080904B0',[StringStruct('CompanyName',''),StringStruct('FileDescription','Corrugados - Consulta de productos'),StringStruct('FileVersion','1.0.0'),StringStruct('InternalName','Corrugados'),StringStruct('OriginalFilename','Corrugados.exe'),StringStruct('ProductName','Corrugados'),StringStruct('ProductVersion','1.0.0')])]),VarFileInfo([VarStruct('Translation',[2057,1200])])])"""
(ROOT/'version.txt').write_text(version,encoding='utf-8')
cmd=[sys.executable,'-m','PyInstaller','--noconfirm','--clean','--onefile','--windowed','--name','Corrugados','--icon',str(ROOT/'corrugados.ico'),'--version-file',str(ROOT/'version.txt'),'--distpath',str(OUT),'--workpath',str(ROOT/'build'),'--specpath',str(ROOT),'--collect-all','customtkinter','--add-data',str(ROOT/'corrugados.ico')+';.']
for module in ['numpy','pandas','matplotlib','scipy','pytest','pywinauto','win32com','IPython','notebook']:
    cmd+=['--exclude-module',module]
cmd += [str(ROOT/'app.py')]
subprocess.run(cmd,cwd=ROOT,check=True)
