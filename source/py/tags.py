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
                           0 : 2.5,
                           1 : 2.5,
                           2 : 2.5,
                           3 : 2,
                           4 : 2,
                           5 : 2,
                           6 : 2.3,
                           7 : 1.4,
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
        
    
    def loesche_ordinal_aus_tags(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        try:
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
        except:
            log(inspect.stack,tb())

    
    def loesche_tag_in_allen_dateien(self,tag):
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        panel_nr = [nr for nr in tags['sammlung'] if tag in tags['sammlung'][nr]][0]
        
        tags['sammlung'][panel_nr].remove(tag)
        
        for panels in tags['ordinale'].values():
            if tag in panels[panel_nr]:
                panels[panel_nr].remove(tag)
    
    
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
  
            with open(pfad, 'rb') as f:
                self.mb.tags =  pickle_load(f)
                
            self.mb.class_Sidebar.offen = {nr:1 for nr in self.mb.tags['abfolge']} 
        except:
            log(inspect.stack,tb())
            p = Popup(self.mb, zeit=3)
            p.text = LANG.TAGS_NICHT_GEFUNDEN
            self.lege_tags_an()
            self.mb.class_Sidebar.offen = {nr:1 for nr in self.mb.tags['abfolge']} 
        
        
        # Tags ueberpruefen
        try:
            
            ordinale = list(self.mb.props['ORGANON'].dict_bereiche['ordinal'])
            ordinale_tags = list(self.mb.tags['ordinale'])
            
            o_nicht_in_tags = [o for o in ordinale if o not in ordinale_tags]
            ueberfluessige_tags = [o for o in ordinale_tags if o not in ordinale]
            
            
            # 1) fehlende Tags neu anlegen
            if o_nicht_in_tags:
                
                root = self.mb.props['ORGANON'].xml_tree.getroot()
                namen = [root.find('.//' + o).attrib['Name'] for o in o_nicht_in_tags ]
                
                p = Popup(self.mb, zeit=3)
                p.text = LANG.LEGE_TAGS_FUER_DATEI_AN + '\n'.join(namen)
        
                for o in o_nicht_in_tags:
                    self.erzeuge_tags_ordinal_eintrag(o)
                    
                    
            # 2) ueberfluessige Tags loeschen
            if ueberfluessige_tags:
                                
                pfade = { o : os.path.join(self.mb.pfade['odts'], o + '.odt') for o in ueberfluessige_tags}
                existierende = { o : pf for o,pf in pfade.items() if os.path.exists(pf)}
                nicht_existierende = [o for o in ueberfluessige_tags if o not in existierende]
                
                # nicht existierende werden ohne Rueckmeldung geloescht
                for o in nicht_existierende:
                    del self.mb.tags['ordinale'][o]
                
                if existierende:

                    text = []
                    for i,e in enumerate(existierende):
                        if i < 20:
                            text.append(e)
                        else:
                            zaehler = i % 20
                            text[zaehler] = text[zaehler] + '  ' + e
                        
                    
                    nachricht = LANG.DATEIEN_NICHT_IM_PROJEKT_ABER_AUF_FP.format(self.mb.pfade['odts']) + '\n'.join(text)
                    
                    entscheidung = self.mb.entscheidung(nachricht,"warningbox",16777216)
                    # 3 = Nein oder Cancel, 2 = Ja
                    if entscheidung == 3:
                        return
                    elif entscheidung == 2:
                        for o,pfad in existierende.items():
                            del self.mb.tags['ordinale'][o]
                            os.remove(pfad)
                            self.mb.class_Bereiche.plain_txt_loeschen(o)

        
        except:
            log(inspect.stack,tb())
            
        
        
        
        
        
        
        
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
    
    
    def erstelle_tags_loeschfenster(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Tags_loeschen_Listener(self.mb)
            
            sammlung = self.mb.tags['sammlung']
            tag_panels = [[i,v[0]] for i,v in self.mb.tags['nr_name'].items() if v[1] == 'tag']            
            
            x = 150
            width = 100
            ctrls = {}
            y_max = 0
            
            for nr,name in tag_panels:

                prop_names = ('Label','Align')
                prop_values = (name,1)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", x, 15, width, 20, prop_names, prop_values)  
                
                ctrls.update({name + '###':control})
                
                y = 0
                for t in sammlung[nr]:
                    prop_names = ('Label','MultiLine')
                    prop_values = (t,True)
                    control, model = self.mb.createControl(self.mb.ctx, "Button", x + 10, y + 55, width - 20, 20, prop_names, prop_values)  
                    control.setActionCommand(t)
                    control.addActionListener(listener)
                    ctrls.update({t:control})
                    
                    y += 25
                    
                    if y > y_max:
                        y_max = y
                
                x += (width + 10)
            
            
            x1,y1 = 100,100
            posSize = (x1,y1,x,y_max + 60)
            
            win,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            
            for c,control in ctrls.items():
                cont.addControl(c, control)
                

            prop_names = ('Label',)
            prop_values = (LANG.LOESCHEN,)
            control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 10, 100, 25, prop_names, prop_values)  
            control.setActionCommand('organOn_tAg_loEschen')
            control.addActionListener(listener)
            cont.addControl('organOn_tAg_loEschen', control)
            ctrls.update({'organOn_tAg_loEschen':control})
            
            prop_names = ('Label',)
            prop_values = (LANG.AUSGEWAEHLTE,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 55, 100, 20, prop_names, prop_values) 
            model.FontWeight = 150.0 
            cont.addControl('ausgewaehlte_XXX', control)
            
            prop_names = ('Orientation',)
            prop_values = (1,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedLine", 130, 0, 2, 1000, prop_names, prop_values) 
            cont.addControl('fixed_line_XXX', control)
            
            control, model = self.mb.createControl(self.mb.ctx, "FixedLine", 0, 45, 1300, 2, (), ()) 
            cont.addControl('fixed_line_XXX', control)
            
            listener.ctrls = ctrls
            listener.hoehe = y_max
            listener.win = win
            
        except:
            log(inspect.stack,tb())
    
            
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
        

    
    
    
        
from com.sun.star.awt import XActionListener
class Tags_loeschen_Listener(unohelper.Base,XActionListener): 
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = None
        self.ausgewaehlte = {}
        self.hoehe = 0
        self.win = None
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:            
            cmd = ev.ActionCommand  
            
            if cmd == 'organOn_tAg_loEschen':
                for tag in self.ausgewaehlte:
                    self.mb.class_Tags.loesche_tag_in_allen_dateien(tag)
                self.win.dispose()
                self.mb.class_Sidebar.erzeuge_sb_layout()
                Popup(self.mb, zeit=1, parent=self.mb.topWindow).text = LANG.TAGS_GELOESCHT
                
            else:
                ctrl = self.ctrls[cmd]
                
                # Ausgewaehlte Tags
                if ctrl.PosSize.X == 10:
                    self.aus_ausgewaehlten_entfernen(cmd)
                # Aufgelistete Tags
                else:
                    self.zu_ausgewaehlten_hinzufuegen(cmd)
 
        except:
            log(inspect.stack,tb())
            
    
    def zu_ausgewaehlten_hinzufuegen(self,cmd):    
        if self.mb.debug: log(inspect.stack)
        
        tag_panels = [[i,v[0]] for i,v in self.mb.tags['nr_name'].items() if v[1] == 'tag']
        ctrl = self.ctrls[cmd]
        pos = ctrl.PosSize.X, ctrl.PosSize.Y
        
        
        
        y = 95 + 25 * len(self.ausgewaehlte)
        
        ctrl.setPosSize(10,y,0,0,3)
        
        self.ausgewaehlte.update({ cmd : [ctrl,pos] })
        
        if y > self.hoehe:
            self.win.setPosSize(0,0,0,y + 40,8)
            
        
    def aus_ausgewaehlten_entfernen(self,cmd): 
        if self.mb.debug: log(inspect.stack)
        
        ctrl = self.ctrls[cmd]
        
        x,y = self.ausgewaehlte[cmd][1]
        
        ctrl.setPosSize(x,y,0,0,3)
        
        del self.ausgewaehlte[cmd]
        
        zaehler = 0
        
        for [control,pos] in self.ausgewaehlte.values():
            y = 95 + 25 * zaehler
            control.setPosSize(0,y,0,0,2)
            zaehler += 1
            
        if y > self.hoehe:
            self.win.setPosSize(0,0,0,y + 40,8)
            
    def disposing(self,ev):
        return False
 
      
        
        
        
        
        
        
        
        
        
        
        
    
    