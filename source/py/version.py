# -*- coding: utf-8 -*-

import unohelper


class Version():
    
    def __init__(self,mb,pdk):
        self.mb = mb        
        
        global pd
        pd = pdk
        
        self.version = 0
        
    def pruefe_version(self):
        if self.mb.debug: log(eval(insp))
        if 'Programmversion' in self.mb.props[T.AB].xml_tree.getroot().attrib:
            self.version = self.mb.props[T.AB].xml_tree.getroot().attrib['Programmversion']
        

        if self.version == 0:
            self.an_080b_anpassen()
        if self.version == '0.8.0b':
            self.an_089b_anpassen()
        
    def an_080b_anpassen(self):
        if self.mb.debug: log(eval(insp))
        
        # Programmversion in settings.xml einfuegen
        xml_root = self.mb.props[T.AB].xml_tree.getroot()
        xml_root.attrib['Programmversion'] = self.mb.programm_version
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.props['Projekt'].xml_tree.write(Path)
        
        # Ordner Images in Programmordner/Files einfuegen
        if not os.path.exists(self.mb.pfade['images']):
            os.makedirs(self.mb.pfade['images'])  
            
        # Die fehlende sidebar_content.pkl wird beim Projektstart automatisch erzeugt.
        # Das dict wurde faelschlich als SystemPath erzeugt
        self.mb.settings_exp['speicherort'] = ''
        self.version = '0.8.0b'
        
    def an_089b_anpassen(self):
        if self.mb.debug: log(eval(insp))
        
        if not os.path.exists(pfade['tabs']):
            os.makedirs(pfade['tabs'])
        self.mb.class_Bereiche.erzeuge_leere_datei()
        self.version = '0.8.9b'
        
        
        
        