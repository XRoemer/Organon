# -*- coding: utf-8 -*-

import unohelper


class Version():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb        
                
        self.vers_dict = {
        
        '0.9.0':self.an_091b_anpassen,
        '0.9.1':self.an_092b_anpassen,
        '0.9.2':None,
        '0.9.3':None,
        '0.9.4':None,
        '0.9.5':None,
        '0.9.6':None,
        '0.9.7':self.an_098b_anpassen,
        '0.9.8':self.an_0981b_anpassen,
        '0.9.8.1':None,
        '0.9.8.2':None,
        '0.9.8.3':None,
        '0.9.8.4':None,
        '0.9.8.5':None,
        '0.9.8.6':self.an_0987b_anpassen,
        '0.9.8.7':self.an_0988b_anpassen
        
        } 
        
        
    def pruefe_version(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.version = self.mb.props[T.AB].xml_tree.getroot().attrib['Programmversion']

            v = self.version
            v = v.replace('b','')
            
            loslegen = False
            
            if self.version != self.mb.programm_version:
            
                for vers in sorted(self.vers_dict):
                    if v == vers:
                        loslegen = True
                    if loslegen:
                        if self.vers_dict[vers] != None:
                            self.vers_dict[vers]()
                
                
                
                self.neue_programmversion_eintragen()

        except:
            log(inspect.stack,tb())

        
    def neue_programmversion_eintragen(self):
        if self.mb.debug: log(inspect.stack)
        
        # Programmversion in settings.xml einfuegen
        xml_root = self.mb.props[T.AB].xml_tree.getroot()
        xml_root.attrib['Programmversion'] = self.mb.programm_version
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.tree_write(self.mb.props['Projekt'].xml_tree,Path)
        
    
    def an_091b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.settings_exp.update({'neues_proj':0})
        self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp)  
        self.version = '0.9.1b'
        
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
            self.version = '0.9.2b'
        except:
            log(inspect.stack,tb())
        
        
    def an_098b_anpassen(self):
        if self.mb.debug: log(inspect.stack)

        pfade = self.mb.pfade
        # Organon/<Projekt Name>/Settings/Tags
        if not os.path.exists(pfade['icons']):
            os.makedirs(pfade['icons'])
        self.version = '0.9.8b'

    def an_0981b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        try:
            self.mb.settings_proj.update({'trenner': 'farbe'})
            self.mb.settings_proj.update({'trenner_farbe_hintergrund': KONST.FARBE_TRENNER_HINTERGRUND,})
            self.mb.settings_proj.update({'trenner_farbe_schrift': KONST.FARBE_TRENNER_SCHRIFT})
            self.mb.settings_proj.update({'trenner_user_url':''})
            
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)
        except:
            log(inspect.stack,tb())
        self.version = '0.9.8.1b'
        
        
    def an_0987b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        try:
            self.mb.settings_proj.update({'nutze_mausrad': False})
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)
        except:
            log(inspect.stack,tb())
        self.version = '0.9.8.7b'
    
    def an_0988b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        try:
            self.mb.settings_exp.update({'html_export' : {
                                                        'FETT' : 1,
                                                        'KURSIV' : 1,
                                                        'UEBERSCHRIFT' : 1,
                                                        'FUSSNOTE' : 1,
                                                        'FARBEN' : 1,
                                                        'AUSRICHTUNG' : 1,
                                                        'LINKS' : 1,
                                                        'ZITATE' : 0,
                                                        'SCHRIFTGROESSE' : 0,
                                                        'SCHRIFTART' : 0,
                                                        'CSS' : 0
                                                        },
                                         })
            
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp)
        except:
            log(inspect.stack,tb())
        self.version = '0.9.8.8b'
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        