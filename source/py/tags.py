# -*- coding: utf-8 -*-

import unohelper
import pickle
from pickle import load as pickle_load
from pickle import dump as pickle_dump

class Tags(): 
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb     
        
        LANG.BILD = u'Bild'
        
        self.tags_template = {
                              
            'ordinale' : {},
            
            'sichtbare' : [
                           0,1,2,3,4,5,6,7
                           ],
                              
            'sammlung' : {
                          3 : [],
                          4 : [], 
                          5 : [], 
                          },
                              
            'abfolge' : [
                         0,1,2,3,4,5,6,7
                         ],
                              
            'nr_name' : {
                        0 : [LANG.SYNOPSIS,'txt'],
                        1 : [LANG.NOTIZEN,'txt'],
                        2 : [LANG.BILDER,'img'],
                        3 : [LANG.CHARAKTERE,'tag'],
                        4 : [LANG.ORTE,'tag'],
                        5 : [LANG.OBJEKTE,'tag'],
                        6 : [LANG.DATUM,'date'],
                        7 : [LANG.ZEIT,'time']
                       },
                              
            'nr_breite' : {
                           0 : 1.5,
                           1 : 2.5,
                           2 : 2.5,
                           3 : 1,
                           4 : 1,
                           5 : 1,
                           6 : 1,
                           7 : 1,
                           },
            'name_nr' : {}
            }
         
        
  
    def lege_tags_an(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.tags = self.tags_template
            tags = self.mb.tags
            tags['name_nr'] = { k[0] : i for i,k in tags['nr_name'].items()}

            # Anlegen der Verzeichnisstruktur 'ordinal'
            for ordinal in list(self.mb.props[T.AB].dict_bereiche['ordinal']):
                self.erzeuge_tags_ordinal_eintrag(ordinal)
        except:
            log(inspect.stack,tb())
            
    
    def erzeuge_tags_ordinal_eintrag(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        eintrag = {}
        
        for nr,artx in tags['nr_name'].items():
            art = artx[1]
            if art == 'txt':
                eintrag.update({nr:''})
            elif art == 'tag':
                eintrag.update({nr:[]})
            elif art == 'date':
                eintrag.update({nr:None})
            elif art == 'time':
                eintrag.update({nr:None})
            elif art == 'img':
                eintrag.update({nr:''})
        
        tags['ordinale'].update({ordinal:eintrag})
        
    
    def loesche_tag_eintrag(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        urls = [tags['ordinale'][ordinal][i] for i,v in tags['nr_name'].items() if v[1] == 'img']
        bild_panels = [i for i,v in tags['nr_name'].items() if v[1] == 'img']
                            
        del(self.mb.tags['ordinale'][ordinal])
        
        # Wenn das Bild im Dok nicht mehr vorkommt: loeschen
        for url in urls:
            if url != '':
                vorkommen = False
                for ordi in tags['ordinale']:
                    for panel in bild_panels:
                        if url in tags['ordinale'][ordi][panel]:
                            vorkommen = True
                if not vorkommen:
                    try:
                        os.remove(uno.fileUrlToSystemPath(url))
                    except:
                        log(inspect.stack,tb())
    
    def speicher_tags(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad = os.path.join(self.mb.pfade['settings'],'tags.pkl')
            with open(pfad, 'wb') as f:
                pickle_dump(self.mb.tags, f,2)
        except:
            log(inspect.stack,tb())


    def lade_tags(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad = os.path.join(self.mb.pfade['settings'],'tags.pkl')
            
            dict_exists, backup_exists = False, False
    
            if os.path.exists(pfad):
                dict_exists = True
                
            #if dict_exists:
            with open(pfad, 'rb') as f:
                self.mb.tags =  pickle_load(f)
                
    #         else:
    #             self.uebertrage_sb_content_nach_tags()
    #             with open(pfad, 'wb') as f:
    #                 pickle_dump(self.tags, f,2)
                
            self.mb.class_Sidebar.offen = {nr:1 for nr in self.mb.tags['abfolge']} 
        except:
            log(inspect.stack,tb())
            self.lege_tags_an()
            self.mb.class_Sidebar.offen = {nr:1 for nr in self.mb.tags['abfolge']} 
            
    def formatiere_datum(self,datum):
        if self.mb.debug: log(inspect.stack)
        
        
        datum_trenner = self.mb.settings_proj['datum_trenner']
        datum_format = self.mb.settings_proj['datum_format']
        
        gesplittet = datum.split(datum_trenner)
        
        if len(gesplittet) != 3:
            return None,None
        
        t = datum_format.index('dd')
        m = datum_format.index('mm')
        j = datum_format.index('yyyy')
        
        tag,monat,jahr = gesplittet[t],gesplittet[m],gesplittet[j]
        
        if 0 in (len(jahr),len(tag),len(monat)):
            return None,None
        if len(tag)>2 or int(tag) < 1 or int(tag) > 31:
            return None,None
        if len(monat)>2 or int(monat) < 1 or int(monat) > 12:
            return None,None
        
        if len(tag) == 1:
            tag = '0'+tag
        if len(monat) == 1:
            monat = '0'+monat
        
        datum = {
                 t:tag,
                 m:monat,
                 j:jahr
                 }
        text = '{0}{3}{1}{3}{2}'.format(datum[0],datum[1],datum[2],datum_trenner)
        
        odict = {
                 'dd':tag,
                 'mm':monat,
                 'yyyy':jahr
                 }
        
        return text,odict
    
    def formatiere_datumdict_nach_text(self,d_dict):
        if self.mb.debug: log(inspect.stack)
        
        trenner = self.mb.settings_proj['datum_trenner']
        form = self.mb.settings_proj['datum_format']
        datum_txt = '{1}{0}{2}{0}{3}'.format(trenner, d_dict[form[0]], d_dict[form[1]], d_dict[form[2]])
        
        return datum_txt

    def formatiere_zeit(self,zeit):
        if self.mb.debug: log(inspect.stack)
        
        if ':' in zeit and zeit.count(':') == 1:
            zeit_trenner = ':'
        elif '.' in zeit and zeit.count('.') == 1:
            zeit_trenner = '.'
        else:
            return None
            
        gesplittet = zeit.split(zeit_trenner)
        
        if len(gesplittet) != 2:
            return None

        std,minu = gesplittet[0],gesplittet[1]
        
        if 0 in (len(std),len(minu)):
            return None
        if len(std)>2 or int(std) < 0 or int(std) > 23:
            return None
        if len(minu)>2 or int(minu) < 0 or int(minu) > 59:
            return None
        
        if len(std) == 1:
            std = '0' + std
        if len(minu) == 1:
            minu = '0' + minu
        
        text = '{0}:{1}'.format(std,minu)
        
        return text

            
#         if os.path.exists(pfad+'.Backup'):
#             backup_exists = True
#         
#         
#         try:  
#             if dict_exists:          
#                 with open(pfad, 'rb') as f:
#                     self.mb.tags =  pickle_load(f)
#                 #self.ueberpruefe_dict_sb_content(not backup_exists)
#                 
#             if not backup_exists:
#                 self.erzeuge_tags_Backup()
#                 self.lade_Backup()
#         except:
#             log(inspect.stack,tb())
#             
#             self.lege_tags_an()
#             if not backup_exists:
#                 self.lade_Backup()
#             
#         if not dict_exists:
#             self.speicher_tags()
# 
#         self.erzeuge_tags_Backup()
        
    

        
#     def ueberpruefe_tags(self,backup_exists):
#         if self.mb.debug: log(inspect.stack)
#         
#         fehlende = []
# 
#         for ordinal in list(self.mb.props['ORGANON'].dict_bereiche['ordinal']):
#             if ordinal not in self.mb.dict_sb_content['ordinal']:
#                 fehlende.append(ordinal)
# 
#         if len(fehlende) > 0:
#             if backup_exists:
#                 self.lade_Backup(fehlende)
#             else:
#                 for f in fehlende:
#                     self.erzeuge_tags_ordinal_eintrag(f)
                      
    
#     def lade_Backup(self,fehlende = 'all'): 
#         if self.mb.debug: log(inspect.stack,None,'### Attention ###, sidebar_content.pkl.backup loaded!')
#         
#         pfad_Backup = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl.Backup')
#         
#         if fehlende == 'all':
#             fehlende = list(self.mb.props['ORGANON'].dict_bereiche['ordinal'])
#         
#         try:
#             with open(pfad_Backup, 'rb') as f:
#                 backup = pickle_load(f)
#         except:
#             log(inspect.stack,tb())
# 
#             for f in fehlende:
#                 self.erzeuge_tags_ordinal_eintrag(f)
#             return
# 
#         helfer = fehlende[:]  
#         
#         if backup != None:  
#             for f in fehlende:
#                 if f in backup['ordinal']:
#                     self.mb.dict_sb_content['ordinal'].update(backup['ordinal'][f])
#                     helfer.remove(f)
#         
#         fehlende = helfer      
#         for f in fehlende:
#             self.erzeuge_tags_ordinal_eintrag(f)
#          
#        
#     def erzeuge_tags_Backup(self):
#         if self.mb.debug: log(inspect.stack)
#         
#         pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
#         pfad_Backup = pfad + '.Backup'
#         from shutil import copy2
#         copy2(pfad, pfad_Backup)
        

    
    
    
        
        
        
        
        
        
        
        
        
        
        
        
        
    
    