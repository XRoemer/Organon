# -*- coding: utf-8 -*-

import unohelper


class Version():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb        
                
        self.vers_dict = {
        
        '0.9.7':self.an_098b_anpassen,
        '0.9.8':self.an_0981b_anpassen,
        '0.9.8.1':None,
        '0.9.8.2':None,
        '0.9.8.3':None,
        '0.9.8.4':None,
        '0.9.8.5':None,
        '0.9.8.6':self.an_0987b_anpassen,
        '0.9.8.7':self.an_0988b_anpassen,
        '0.9.8.8':None,
        '0.9.9.0':None,
        '0.9.9.1':None,
        '0.9.9.2':None,
        '0.9.9.3':None,
        '0.9.9.4':None,
        '0.9.9.5':None,
        '0.9.9.6':None,
        '0.9.9.7':None,
        '0.9.9.8':self.an_09981b_anpassen,
        '0.9.9.8.1':None,
        '0.9.9.8.2':None,
        '0.9.9.8.3':self.an_09984b_anpassen,
        
        } 
        
        
    def pruefe_version(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.version = self.mb.props[T.AB].xml_tree.getroot().attrib['Programmversion']

            v = self.version
            v = v.replace('b','')
            
            loslegen = False
            ver = sorted(self.vers_dict)
            if self.version != self.mb.programm_version:
                ver = sorted(self.vers_dict)
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
        self.mb.tree_write(self.mb.props['ORGANON'].xml_tree,Path)
        
      
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
        
    
        
    def an_09981b_anpassen(self):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['ORGANON']
        
        tree = props.xml_tree
        root = tree.getroot()
        root.tag = 'ORGANON'
        
        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.tree_write(self.mb.props['ORGANON'].xml_tree,Path)
        
        
    def an_09984b_anpassen(self):
        if self.mb.debug: log(inspect.stack) 
        
        try:
            # aendern :
            # Datumsformat fuer englische Version muss evtl angepasst werden
            
            # neue Datum Formatierung
            
            if self.mb.language == 'de':
                datum_format = ['dd','mm','yyyy']
            else:
                datum_format = ['mm','dd','yyyy']
                
            self.mb.settings_proj.update({'datum_trenner' : '.',
                                          'datum_format' : datum_format })
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
            
            message = u'''This project was created by an older version of Organon.
The settings of the tagging category date / time has changed.

If your project uses date tags, check if the formatting of dates is in the right order.
Standard Organon date formatting is day/month/year for the german version of Organon and
month/day/year for any other language version of Organon.

The formatting can be set under: Organon menu / File / Settings / Tags / Date Format

A backup of your project with the old settings will be created in the backup folder of your project. 

            '''
            
            self.mb.nachricht(message,'infobox')
            self.mb.erzeuge_Backup()
            
            # dict sidebar_content in dict tag ueberfuehren
            pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
            pfad3 = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl.Backup')
     
            from pickle import load as pickle_load     
            from pickle import dump as pickle_dump
            with open(pfad, 'rb') as f:
                dict_sb_content =  pickle_load(f)
                
            sb_panels = {
                'Synopsis':LANG.SYNOPSIS,
                'Notes':LANG.NOTIZEN,
                'Images':LANG.BILDER,
                'Tags_general':LANG.ALLGEMEIN,
                'Tags_characters':LANG.CHARAKTERE,
                'Tags_locations':LANG.ORTE,
                'Tags_objects':LANG.OBJEKTE,
                'Tags_time':LANG.ZEIT,
                'Tags_user1':LANG.BENUTZER1,
                'Tags_user2':LANG.BENUTZER2,
                'Tags_user3':LANG.BENUTZER3
                }
            
            sb_panels_tup = (
                'Synopsis',
                'Notes',
                'Images',
                'Tags_general',
                'Tags_characters',
                'Tags_locations',
                'Tags_objects',
                'Tags_time',
                'Tags_user1',
                'Tags_user2',
                'Tags_user3'
                )
            
            tags = {
                    'nr_name' : {
                        0 : [u'SYNOPSIS','txt'],
                        1 : [u'NOTIZEN','txt'],
                        2 : [u'BILDER','img'],
                        3 : [u'ALLGEMEIN','tag'],
                        4 : [u'CHARAKTERE','tag'],
                        5 : [u'ORTE','tag'],
                        6 : [u'OBJEKTE','tag'],
                        7 : [u'DATUM','date'],
                        8 : [u'ZEIT','time'],
                        9 : [u'BENUTZER1','tag'],
                        10 : [u'BENUTZER2','tag'],
                        11 : [u'BENUTZER3','tag']
                       },}
            
            
            alte_kats = list(dict_sb_content['ordinal'][ list(dict_sb_content['ordinal'])[0] ])
            
            name_index = { k : sb_panels_tup.index(k) for k in alte_kats if k in sb_panels_tup }
            index_name = { sb_panels_tup.index(k) : k for k in alte_kats if k in sb_panels_tup }
            
            tags['ordinale'] = { ord: {name_index[k] : i2
                    for k,i2 in i.items()}
                   for ord,i in dict_sb_content['ordinal'].items() if isinstance(i, dict)}
                  
            tags['sichtbare'] = [name_index[k] for k in dict_sb_content['sichtbare'] ]
            tags['sammlung'] = {name_index[k]:i for k,i in dict_sb_content['tags'].items() }
            tags['nr_name'] = { i : [ getattr(LANG,k[0]) , k[1] ] for i,k in tags['nr_name'].items()}
            tags['name_nr'] = {  k[0] : i for i,k in tags['nr_name'].items()}
            tags['abfolge'] = list(range(len(tags['nr_name'])))
            
            tags['nr_breite'] = {i:2 for i in range(12)}
            tags['nr_breite'].update({
                                        0 : 5,
                                        1 : 5,
                                        2 : 3
                                      })
            
            # Tags in Tags_Allgemein loeschen, die in anderen Tags vorhanden sind
            from itertools import chain
            alle_tags_in_anderen_panels = list(chain.from_iterable(
                                            [v for i,v in tags['sammlung'].items() if i != 3 ]
                                            ))
            
            for ordi in tags['ordinale']:
                for t in alle_tags_in_anderen_panels:
                    if t in tags['ordinale'][ordi][3]:
                        tags['ordinale'][ordi][3].remove(t)
                    if t in tags['sammlung'][3]:
                        tags['sammlung'][3].remove(t)
            
            
            # Zeit und Datum trennen
            for ordi in tags['ordinale']:
                
                for i in range(11,8,-1):
                    tags['ordinale'][ordi].update({i:tags['ordinale'][ordi][i-1]})
                
                if 'zeit' in tags['ordinale'][ordi][7]:  
                    tags['ordinale'][ordi][8] = tags['ordinale'][ordi][7]['zeit']
                else:
                    tags['ordinale'][ordi][8] = None
                
                if 'datum' not in tags['ordinale'][ordi][7]: 
                    tags['ordinale'][ordi][7] = None
                elif tags['ordinale'][ordi][7]['datum'] == None:
                    tags['ordinale'][ordi][7] = None
                else:
                    dat_split = tags['ordinale'][ordi][7]['datum'].split('.')
                    tags['ordinale'][ordi][7] = {
                                                 datum_format[0] : dat_split[0],
                                                 datum_format[1] : dat_split[1],
                                                 datum_format[2] : dat_split[2],
                                                 }
             
            for i in range(11,8,-1):
                tags['sammlung'].update({i:tags['sammlung'][i-1]})
                
            del tags['sammlung'][8]
            
            
            # neue zeit formatierung
            panel_nr = [i for i,v in tags['nr_name'].items() if v[1] == 'time'][0]
            
            for ordi in tags['ordinale']:
                
                zeit = tags['ordinale'][ordi][panel_nr]
                if zeit == None:
                    continue
                    
                zeit_str = str(zeit)
        
                if len(zeit_str) == 7:
                    zeit_str = '0' + zeit_str
        
                std = int(zeit_str[0:2])
                minu = int(zeit_str[2:4])
                
                tags['ordinale'][ordi][panel_nr] = '{0}:{1}'.format(std,minu)    
                
                
            # tags speichern
            pfad2 = os.path.join(self.mb.pfade['settings'],'tags.pkl')
            with open(pfad2, 'wb') as f:
                pickle_dump(tags, f,2)
            
            try:
                os.remove(pfad)
            except:
                pass
            try:
                os.remove(pfad3)
            except:
                pass
        except:
            log(inspect.stack,tb())
            
        
        
        
        
        
        
        
        