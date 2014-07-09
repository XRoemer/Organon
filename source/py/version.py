# -*- coding: utf-8 -*-

import unohelper


class Version():
    
    def __init__(self,mb,pdk):
        self.mb = mb        
        
        global pd
        pd = pdk
        
    def pruefe_version(self):
        if 'Programmversion' in self.mb.xml_tree.getroot().attrib:
            version = self.mb.xml_tree.getroot().attrib['Programmversion']
        else:
            version = 0

        if version == 0:
            self.an_080b_anpassen()
        
        
    def an_080b_anpassen(self):
        if self.mb.debug: print(self.mb.debug_time(),'version an_080b_anpassen')
        # Programmversion in settings.xml einfuegen
        xml_root = self.mb.xml_tree.getroot()
        xml_root.attrib['Programmversion'] = self.mb.programm_version
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.xml_tree.write(Path)
        
        # Ordner Images in Programmordner/Files einfuegen
        if not os.path.exists(self.mb.pfade['images']):
            os.makedirs(self.mb.pfade['images'])  
            
        # Die fehlende sidebar_content.pkl wird beim Projektstart automatisch erzeugt.
        
        # Das dict wurde faelschlich als SystemPath erzeugt
        self.mb.settings_exp['speicherort'] = ''
        
        
        
        