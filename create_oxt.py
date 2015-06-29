# -*- coding: utf-8 -*-
import os
import shutil
import zipfile
from os import walk


office = 'libre'

if office == 'libre':
    path_to_office = "C:\\Program Files (x86)\\LibreOffice 4\\program"
else:
    path_to_office = "C:\\Program Files (x86)\\OpenOffice 4\\program"

path_to_source = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source'

files = [
    'factory/WindowContentFactories.xcu',
    'factory/WriterWindowState.xcu',
    'Addons.xcu',


    'META-INF/manifest.xml',
    'factory/Dialog1.xdl',
    'description.xml',

    'languages/lang_de.py',
    'languages/lang_en.py',

    'py/factory.py',
    'py/log_organon.py',
    'py/menu_bar.py',
    'py/menu_start.py',
    'py/bereiche.py',
    'py/projects.py',
    'py/funktionen.py',
    'py/baum.py',
    'py/xml_m.py',
    'py/export.py',
    'py/importX.py',
    'py/konstanten.py',
    'py/sidebar.py',
    'py/tabs.py',
    'py/version.py',
    'py/latex_export.py',
    'py/export2html.py',
    'py/zitate.py',
    'py/mausrad.py',
    'py/rawinputdata.py',
    'py/einstellungen.py',
    'py/werkzeug_wListe.py',
    'py/index.py',
    'py/design.py',

    'py/schalter.py',

    'libs/libGetWheel.so',
    'organon_settings.json',

    # Sidebar
    'factory/Dialog_Sidebar.xdl',
    'factory/Factory_Sidebar.xcu',
    'factory/Sidebar.xcu',
    'factory/ProtocolHandler.xcu',



    'description/desc_de.txt',
    'description/desc_en.txt']

# ICONS
mypath2 = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\img'
for (dirpath, dirnames, filenames) in walk(mypath2):
    for f in filenames:
        files.append(dirpath.split('source')[1].replace('\\', '', 1)+'/'+f)


# REGISTRATION FILES
mypath1 = os.path.join(path_to_source, 'registration')
for (dirpath, dirnames, filenames) in walk(mypath1):
    f_names = filenames
for i in f_names:
    files.append('registration/'+i)


def get_folder(path):
    files2 = []
    f = []
    for (dirpath, dirnames, filenames) in walk(path):
        f.append((dirpath, dirnames, filenames))
    for i in f:
        dirpath, dirnames, filenames = i
        path = dirpath.split('source\\')[1]
        for j in dirnames:
            files2.append(path+'\\'+j)
        for k in filenames:
            files2.append(path+'\\'+k)
    return files2

mypath = os.path.join(path_to_source, 'description', 'Handbuecher')
for f in get_folder(mypath):
    files.append(f)

mypath = os.path.join(path_to_source, 'personas')
for f in get_folder(mypath):
    files.append(f)

zip = zipfile.ZipFile('organon.oxt', 'w')

for file in files:
    zip.write(file)

zip.close()

# # .oxt in den Programmordner von LibreOffice kopieren
# shutil.copy("organon.oxt", path_to_office)

# # in den programmordner von libreoffice wechseln und unopkg starten
# os.chdir(path_to_office)
# os.system(r"unopkg add -f --shared organon.oxt")
# # swriter nach Installation erneut starten
# os.system('swriter.exe')
