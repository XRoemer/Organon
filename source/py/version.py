# -*- coding: utf-8 -*-

import unohelper


class Version():
    
    def __init__(self,mb,pdk):
        if mb.debug: log(inspect.stack)
        self.mb = mb        
        
        global pd
        pd = pdk
        
        self.version = 0
        
    def pruefe_version(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if 'Programmversion' in self.mb.props[T.AB].xml_tree.getroot().attrib:
                self.version = self.mb.props[T.AB].xml_tree.getroot().attrib['Programmversion']
    
    
            if self.version == 0:
                self.an_080b_anpassen()
            if self.version in ('0.8.0b','0.8.1b'):
                self.an_090b_anpassen()
            if self.version in ('0.9.0b'):
                self.an_091b_anpassen()
            if self.version in ('0.9.1b'):
                self.an_092b_anpassen()
            # an 093, 094, 095, 096  muss nichts angepasst werden
            self.neue_programmversion_eintragen()
        
        except:
            if self.mb.debug: log(inspect.stack,tb())
        
    def neue_programmversion_eintragen(self):
        if self.mb.debug: log(inspect.stack)
        
        # Programmversion in settings.xml einfuegen
        xml_root = self.mb.props[T.AB].xml_tree.getroot()
        xml_root.attrib['Programmversion'] = self.mb.programm_version
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.tree_write(self.mb.props['Projekt'].xml_tree,Path)
        
    def an_080b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        # Programmversion in settings.xml einfuegen
        xml_root = self.mb.props[T.AB].xml_tree.getroot()
        xml_root.attrib['Programmversion'] = self.mb.programm_version
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.tree_write(self.mb.props['Projekt'].xml_tree,Path)
        
        # Ordner Images in Programmordner/Files einfuegen
        if not os.path.exists(self.mb.pfade['images']):
            os.makedirs(self.mb.pfade['images'])  
            
        # Die fehlende sidebar_content.pkl wird beim Projektstart automatisch erzeugt.
        # Das dict wurde faelschlich als SystemPath erzeugt
        self.mb.settings_exp['speicherort'] = ''
        self.version = '0.8.0b'
        
    def an_090b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        if not os.path.exists(self.mb.pfade['tabs']):
            os.makedirs(self.mb.pfade['tabs'])
        self.mb.class_Bereiche.erzeuge_leere_datei()
        self.version = '0.9.0b'
    
    def an_091b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.settings_exp.update({'neues_proj':0})
        self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp)  
        
    def an_092b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        # Tags_time in den Sidebar dict eintragen
        try:
            from pickle import load as pickle_load
            pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
            with open(pfad, 'rb') as f:
                self.mb.dict_sb_content =  pickle_load(f)
            
            dict_sb_content = self.mb.dict_sb_content
            
            dict = {}
            dict.update({'zeit':None})
            dict.update({'datum':None})
            
            for ordn in dict_sb_content['ordinal']:
                if dict_sb_content['ordinal'][ordn]['Tags_time'] == '':
                    self.mb.dict_sb_content['ordinal'][ordn]['Tags_time'] = dict
            
            self.mb.class_Sidebar.speicher_sidebar_dict()
        except:
            if self.mb.debug: log(inspect.stack,tb())
        
        
        
        
        