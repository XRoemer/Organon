# -*- coding: utf-8 -*-

import unohelper
import copy

class Einstellungen():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.container = None

        # enthaelt die LB Eintraege und entsprechende Organon props
        self.lb_dict = None
        # vice versa
        self.lb_dict2 = None
        
        self.datum_items = {
                       ('dd', 'mm', 'yyyy') : 'dd/mm/yyyy',
                       ('mm', 'dd', 'yyyy') : 'mm/dd/yyyy',
                       ('yyyy', 'mm', 'dd') : 'yyyy/mm/dd',
                       ('yyyy', 'dd', 'mm') : 'yyyy/dd/mm'
                       }
        self.datum_items2 = {v:i for i,v in self.datum_items.items()}
        
        
    def start(self):
        if self.mb.debug: log(inspect.stack)
        
        try:   
            self.erzeuge_einstellungsfenster()
        except Exception as e:
            log(inspect.stack,tb())


    def erzeuge_einstellungsfenster(self): 
        if self.mb.debug: log(inspect.stack)
        try:
            breite = 680
            hoehe = 490
            breite_listbox = 150
            
            sett = self.mb.settings_exp
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = 20
            Y = self.mb.win.AccessibleContext.LocationOnScreen.Y - posSize_main.Y +40
            
            # Listener erzeugen 
            listener = {}           
            listener.update( {'auswahl_listener': Auswahl_Item_Listener(self.mb)} )
            
            controls = self.dialog_einstellungen(listener, breite_listbox,breite,hoehe)
            ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)                
            
            # Hauptfenster erzeugen
            posSize = X,Y,breite,pos_y + 40
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            
            self.haupt_fenster = fenster
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                fenster_cont.addControl(c,ctrls[c])
            
            self.container = ctrls['control_Container']
            listener['auswahl_listener'].container = ctrls['control_Container']
            listener['auswahl_listener'].fenster = fenster
            
        except:
            log(inspect.stack,tb())
                
       
    
    def dialog_einstellungen(self,listener,breite_listbox,breite,hoehe):
        if self.mb.debug: log(inspect.stack)
        
        if self.mb.programm == 'LibreOffice':
            lb_items = (
                        LANG.DESIGN_TRENNER,
                        LANG.DESIGN_ORGANON,
                        LANG.DESIGN_PERSONA,
                        LANG.TAGS,
                        LANG.TEMPLATES_ORGANON,
                        LANG.SHORTCUTS,
                        LANG.UEBERSETZUNGEN,
                        LANG.MAUSRAD,
                        LANG.HTML_EXPORT,
                        LANG.LOG
                        )
        else:
            # OPEN OFFICE
            lb_items = (
                        LANG.DESIGN_TRENNER,
                        LANG.DESIGN_ORGANON,
                        LANG.TAGS,
                        LANG.TEMPLATES_ORGANON,
                        LANG.SHORTCUTS,
                        LANG.UEBERSETZUNGEN,
                        LANG.MAUSRAD,
                        LANG.HTML_EXPORT,
                        LANG.LOG
                        )
            
           
        controls = (
            10,
            ('controlE_calc',"FixedText",1,        
                                    20,0,250,20,    
                                    ('Label','FontWeight'),
                                    (LANG.EINSTELLUNGEN ,150),                  
                                    {} 
                                    ), 
            20,                                                  
            ('control_Liste',"ListBox",0,      
                                    20,0,breite_listbox,hoehe,    
                                    ('Border',),
                                    ( 2,),       
                                    {'addItems':(lb_items,0),'addItemListener':(listener['auswahl_listener'])} 
                                    ),  
            0,
            ('control_Container',"Container",0,      
                                    breite_listbox + 40,0,breite-60-breite_listbox ,hoehe ,    
                                    ('Border',),
                                    (1,),              
                                    {} 
                                    ),  
            hoehe - 20,
            )
        return controls
    
    

            
     

    def container_anpassen(self,container,max_breite=None,max_hoehe=None,fenster=None):
        if self.mb.debug: log(inspect.stack)
        
                  
        posSize = container.PosSize
        
        # Breite 
        if max_breite:
            if max_breite > posSize.Width:
                container.setPosSize(0,0,max_breite,0,4)
                if fenster:
                    unterschied = max_breite- posSize.Width
                    fenster.setPosSize(0,0,fenster.PosSize.Width + unterschied,0,4)
            
        # Hoehe
        if max_hoehe:            
            if max_hoehe > posSize.Height:
                container.setPosSize(0,0,0,max_hoehe,8)
                if fenster:
                    unterschied = max_hoehe - posSize.Height
                    fenster.setPosSize(0,0,0,fenster.PosSize.Height + unterschied,8)
  

from com.sun.star.awt import XItemListener   
class Auswahl_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.container = None
        self.fenster = None
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
        
        try:     
            sel = ev.value.Source.Items[ev.value.Selected] 

            for c in self.container.getControls():
                c.dispose()
            
            if sel == LANG.LOG:
                self.dialog_logging()
                    
            elif sel == LANG.MAUSRAD:
                self.dialog_mausrad()
                                        
            elif sel == LANG.DESIGN_TRENNER:
                self.dialog_trenner()
                
            elif sel == LANG.HTML_EXPORT:
                self.dialog_html_export()
                
            elif sel == LANG.DESIGN_ORGANON:
                self.mb.class_Organon_Design.container = self.container
                self.mb.class_Organon_Design.fenster = self.fenster
                self.mb.class_Organon_Design.dialog_organon_farben()
                
            elif sel == LANG.DESIGN_PERSONA:
                self.mb.class_Organon_Design.container = self.container
                self.mb.class_Organon_Design.dialog_persona()
                
            elif sel == LANG.SHORTCUTS:
                self.dialog_shortcuts()
                
            elif sel == LANG.UEBERSETZUNGEN:
                u = Uebersetzungen(self.mb).run()
                
            elif sel == LANG.TEMPLATES_ORGANON:
                self.dialog_templates()
                
            elif sel == LANG.TAGS:
                self.dialog_tags()

        except:
            log(inspect.stack,tb())
            

    def disposing(self,ev):
        return False
  
        
    def dialog_logging(self):
        if self.mb.debug: log(inspect.stack)

        try:
            ctx = self.mb.ctx
            mb = self.mb
            
            tab = 20
            tab1 = 40

            fenster_cont = self.container
            
            y = 10
            
            prop_names = ('Label','FontWeight',)
            prop_values = (LANG.EINSTELLUNGEN_LOGDATEI,200,)
            control, model = mb.createControl(ctx, "FixedText", tab, y, 200, 20, prop_names, prop_values)
            fenster_cont.addControl('Titel', control) 

            y += 50
            
            prop_names = ('Label','State')
            prop_values = (LANG.KONSOLENAUSGABE,self.mb.class_Log.output_console)
            controlCB, model = mb.createControl(ctx, "CheckBox", tab, y, 200, 20, prop_names, prop_values)
            controlCB.setActionCommand('Konsole')
             
            y += 30
            
            prop_names = ('Label','State',)
            prop_values = (LANG.ARGUMENTE_LOGGEN,self.mb.class_Log.log_args,)
            control_arg, model = mb.createControl(ctx, "CheckBox", tab1, y, 200, 20, prop_names, prop_values)
            control_arg.setActionCommand('Argumente')
            control_arg.Enable = (self.mb.class_Log.output_console == 1)
            
            y += 40
            
            prop_names = ('Label','State',)
            prop_values = (LANG.LOGDATEI_ERZEUGEN,self.mb.class_Log.write_debug_file,)
            control_log, model = mb.createControl(ctx, "CheckBox", tab1, y, 200, 20, prop_names, prop_values)
            control_log.setActionCommand('Logdatei')
            control_log.Enable = (self.mb.class_Log.output_console == 1)
            
            y += 20
            
            prop_names = ('Label',)
            prop_values = (LANG.SPEICHERORT,)
            control_filepath, model = mb.createControl(ctx, "FixedText", tab1, y, 200, 20, prop_names, prop_values)
            fenster_cont.addControl('Titel', control_filepath) 
            
            y += 20
            
            prop_names = ('Label',)
            prop_values = (mb.class_Log.location_debug_file,)
            control_path, model = mb.createControl(ctx, "FixedText", tab1, y, 600, 20, prop_names, prop_values)
            fenster_cont.addControl('Titel', control_path)
            
            # Breite des Log-Fensters setzen
            prefSize = control_path.getPreferredSize()
            Hoehe = prefSize.Height 
            Breite = prefSize.Width
            control_path.setPosSize(0,0,Breite+10,0,4)
            
            y += 20
            
            prop_names = ('Label',)
            prop_values = (LANG.AUSWAHL,)
            control_but, model = mb.createControl(ctx, "Button", tab1, y, 80, 20, prop_names, prop_values)
            control_but.setActionCommand('File')
            fenster_cont.addControl('Titel', control)
            
            
            # Listener setzen
            log_listener = Listener_Logging_Einstellungen(self.mb,control_log,control_arg,control_path)
            
            controlCB.addActionListener(log_listener)
            fenster_cont.addControl('Konsole', controlCB)
            
            control_log.addActionListener(log_listener)
            fenster_cont.addControl('Log', control_log) 
            
            control_but.addActionListener(log_listener)
            fenster_cont.addControl('but', control_but)
            
            control_arg.addActionListener(log_listener)
            fenster_cont.addControl('arg', control_arg)
            
        except:
            log(inspect.stack,tb())
            
    
    
    def dialog_html_export_elemente(self,listener,html_exp_settings):
        if self.mb.debug: log(inspect.stack)
        
        controls = [
            10,
            ('controlE_calc',"FixedText",1,        
                                    20,0,600,20,    
                                    ('Label','FontWeight'),
                                    (LANG.HTML_AUSWAHL ,150),                  
                                    {} 
                                    ), 
            50, ]
        
        elemente = 'FETT','KURSIV','AUSRICHTUNG','UEBERSCHRIFT','FUSSNOTE','FARBEN','LINKS'
                # ZITATE,SCHRIFTART,SCHRIFTGROESSE,CSS
                
        for el in elemente:
            controls.extend([
            ('control_{}'.format(el),"CheckBox",1,      
                                    20,0,200,20,    
                                    ('Label','State'),
                                    (getattr(LANG, el),html_exp_settings[el]),       
                                    {'setActionCommand':el,'addActionListener':(listener)} 
                                    ),  
            25])
            
        return controls

 
    def dialog_html_export(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            if self.mb.settings_exp == None:
                # Noch kein Projekt geladen
                return
            
            sett = self.mb.settings_exp['html_export']

            # Listener erzeugen 
            listener = Listener_HTML_Export_Einstellungen(self.mb)         
            
            controls = self.dialog_html_export_elemente(listener,sett)
            ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)                
              
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                         
        except:
            log(inspect.stack,tb())
            
                    
    def dialog_mausrad_elemente(self,listener,nutze_mausrad):
        if self.mb.debug: log(inspect.stack)
        
        controls = (
            10,
            ('controlE_calc',"FixedText",1,        
                                    20,0,400,20,    
                                    ('Label','FontWeight'),
                                    (LANG.NUTZE_MAUSRAD ,150),                  
                                    {} 
                                    ), 
            50,                                                  
            ('control_CB_calc',"CheckBox",1,      
                                    20,0,200,30,    
                                    ('Label','State'),
                                    (LANG.NUTZE_MAUSRAD,nutze_mausrad),       
                                    {'addActionListener':(listener)} 
                                    ),  
            50,
            ('control_Container_calc',"FixedText",1,     
                                    20,0,350,200,    
                                    ('MultiLine','Label'),
                                    (True,LANG.MAUSRAD_HINWEIS),              
                                    {} 
                                    ),  
            200,
            )
        return controls
    
    
    def dialog_mausrad(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            sett = self.mb.settings_orga
            
            try:
                if sett['mausrad']:
                    nutze_mausrad = 1
                else:
                    nutze_mausrad = 0
            except:
                sett['mausrad'] = False
                nutze_mausrad = 0
                       

            # Listener erzeugen 
            listener = Listener_Mausrad_Einstellungen(self.mb)         
            
            controls = self.dialog_mausrad_elemente(listener,nutze_mausrad)
            ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)                
             
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                         
        except:
            log(inspect.stack,tb())
            
    
    def dialog_trenner_elemente(self,trenner_dict,listener_CB,listener_URL):
        if self.mb.debug: log(inspect.stack)
        
        sett_trenner = self.mb.settings_orga['trenner']

        controls = [
            10,
            ('control',"FixedText",1,         
                                    'tab0',0,50,20,  
                                    ('Label','FontWeight'),
                                    (LANG.TRENNER_FORMATIERUNG ,150),                                             
                                    {} 
                                    ),
            50,
            ('control1',"CheckBox",1,        
                                    'tab0',0,50,20,   
                                    ('Label','State'),
                                    (LANG.LINIE,trenner_dict['strich']),                                                  
                                    {'addItemListener':(listener_CB)} 
                                    ) ,
            0,
            ('control11',"CheckBox",1,        
                                    'tab4x',0,50,20,   
                                    ('Label','State'),
                                    (LANG.KEIN_TRENNER,trenner_dict['keiner']),                                           
                                    {'addItemListener':(listener_CB)} 
                                    ),
            50,
            ('control2',"CheckBox",1,         
                                    'tab0',0,50,20,   
                                    ('Label','State'),
                                    (LANG.FARBE,trenner_dict['farbe']),                                                   
                                    {'addItemListener':(listener_CB)} 
                                    ), 
            20,
            ('control3',"FixedText",1,        
                                    'tab0x+15',0,400,60,   
                                    ('Label',),
                                    (LANG.TRENNER_ANMERKUNG,),                                                                            
                                    {} 
                                    ), 
             
            60,
            ('control7',"CheckBox",1,         
                                    'tab0',0,80,20,   
                                    ('Label','State'),
                                    (LANG.BENUTZER,trenner_dict['user']),                                                  
                                    {'addItemListener':(listener_CB)} 
                                    ),              
            20,
            ('control8',"FixedText",1,        
                                    'tab0+15',0,20,20,   
                                    ('Label',),
                                    ('URL: ',),                                                                            
                                    {'Enable':trenner_dict['user']==1} 
                                    ), 
            0,
            ('control10',"Button",1,          
                                    'tab1',0,60,20,    
                                    ('Label',),
                                    (LANG.AUSWAHL,),                                                                         
                                    {'Enable':trenner_dict['user']==1,'addActionListener':(listener_URL)} 
                                    ), 
            30,
            ('control9',"FixedText",1,        
                                    'tab1',-4,600,20,   
                                    ('Label',),
                                    (sett_trenner['trenner_user_url'],),                                            
                                    {'Enable':trenner_dict['user']==1} 
                                    ), 
            20,
            ]
        
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 15),
                 1 : (None, 1),
                 2 : (None, 5),
                 3 : (None, 5),
                 4 : (None, 5),
                 }
        
        abstand_links = 10
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                
        return controls2,max_breite
            
    def dialog_trenner(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            listener_CB = Listener_Trenner(self.mb)
            listener_URL = Listener_Trenner(self.mb)
            
            trenner = 'keiner', 'strich', 'farbe', 'user'
            trenner_dict = {}
            for t in trenner:
                if self.mb.settings_orga['trenner']['trenner'] == t:
                    trenner_dict.update({t: 1})
                else:
                    trenner_dict.update({t: 0})
             

            controls,max_breite = self.dialog_trenner_elemente(trenner_dict,listener_CB,listener_URL)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls) 


            # UEBERGABE AN LISTENER
            listener_CB.conts = {'controls':{'strich': ctrls['control1'],
                                             'farbe': ctrls['control2'],
                                             'user': ctrls['control7'],
                                             'keiner': ctrls['control11']
                                             },

                                 'user':{'url': ctrls['control8'],
                                         'auswahl': ctrls['control10']
                                         },
                                 'keiner':{},
                                 'strich':{}
                                 }
            
            listener_URL.url_textfeld = ctrls['control9']
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
            
        except:
            log(inspect.stack,tb())
            

    def dialog_shortcuts_elemente(self,listener):#,trenner_dict,listener_CB,listener_URL):
        if self.mb.debug: log(inspect.stack)
        
        controls = [10,]
                
        from collections import OrderedDict
        
        shorts = [
                 ['INSERT_DOC' , LANG.INSERT_DOC],
                 ['INSERT_DIR' , LANG.INSERT_DIR],
                 ['TRENNE_TEXT' , LANG.TRENNE_TEXT],
                 ['SUCHE' , LANG.SUCHE],
                 ['DATEIEN_VEREINEN' , LANG.DATEIEN_VEREINEN],
                 ['BENENNE_DATEI_UM' , LANG.BENENNE_DATEI_UM],
                 ['IN_PAPIERKORB_VERSCHIEBEN' , LANG.IN_PAPIERKORB_VERSCHIEBEN],
                 ['CLEAR_RECYCLE_BIN' , LANG.CLEAR_RECYCLE_BIN],
                 ['FORMATIERUNG_SPEICHERN2' , LANG.FORMATIERUNG_SPEICHERN2],
                 ['NEUER_TAB' , LANG.NEUER_TAB],
                 ['SCHLIESSE_TAB' , LANG.SCHLIESSE_TAB],
                 ['BACKUP' , LANG.BACKUP],
                 ['OEFFNE_ORGANIZER' , LANG.OEFFNE_ORGANIZER],
                 ['SHOW_TAG1' , LANG.SICHTBARKEIT + ' ' + LANG.SHOW_TAG1],
                 ['SHOW_TAG2' , LANG.SICHTBARKEIT + ' ' + LANG.SHOW_TAG2],
                 ['GLIEDERUNG' , LANG.SICHTBARKEIT + ' ' + LANG.GLIEDERUNG],
                 ['BAUMANSICHT_HOCH' , LANG.BAUMANSICHT_HOCH],
                 ['BAUMANSICHT_RUNTER' , LANG.BAUMANSICHT_RUNTER],
                 ['KONSOLENAUSGABE' , LANG.KONSOLENAUSGABE],
                 ]
        
        shortcuts = OrderedDict()
        
        for s in shorts:
            shortcuts.update({s[0]:s[1]})
        
        
        sett = self.mb.settings_orga['shortcuts']
        
        # 0 = keine Modifikation
        # 1 = Shift
        # 2 = Strg
        # 3 = Shift + Strg
        # 4 = Alt
        # 5 = Shift + Alt
        # 6 = Strg + Alt
        # 7 = Shift + Strg + Alt
        
        def get_settings(command):
            for m in sett:
                for n in sett[m]:
                    for l in sett[m]:
                        if sett[m][l] == command:
                            if   m == '2': a,b,c = 0,1,0
                            elif m == '3': a,b,c = 1,1,0
                            elif m == '4': a,b,c = 0,0,1
                            elif m == '5': a,b,c = 1,0,1
                            elif m == '6': a,b,c = 0,1,1
                            elif m == '7': a,b,c = 1,1,1
                            return a,b,c,l
                        
            return 0,0,0,'-'   
                        
                                
        for s in shortcuts:
            
            shift,ctrl,alt,key = get_settings(s)
            mods = shift*1 + ctrl*2 + alt*4
            
            try:
                if mods > 1:
                    items = self.mb.class_Shortcuts.get_moegliche_shortcuts(mods)
                    i2 = list(items)
                    i2.insert(1, key)
                    
                    items = tuple(i2)
                    sel = 1
                else:
                    items = ('-',)
                    sel = 0
            except:
                items = ('-',)
                sel = 0

            controls.extend([
            ('control_{}'.format(s.strip()),"FixedText",1,      
                                    'tab0',0,100,20,  
                                    ('Label',),
                                    (shortcuts[s] ,),                                             
                                    {} 
                                    ),  
            0,
            ('control_Shift{}'.format(s.strip()),"CheckBox",1,     
                                    'tab1',0,20,20,    
                                    ('Label','State'),
                                    ('Shift',shift),       
                                    {'setActionCommand':'shift+'+s.strip(),'addActionListener':(listener)} 
                                    ),  
            0,
            ('control_Ctrl{}'.format(s.strip()),"CheckBox",1,      
                                    'tab2',0,20,20,    
                                    ('Label','State'),
                                    ('Ctrl',ctrl),       
                                    {'setActionCommand':'ctrl+'+s.strip(),'addActionListener':(listener)} 
                                    ), 
            0,
            ('control_Alt{}'.format(s.strip()),"CheckBox",1,      
                                    'tab3',0,20,20,    
                                    ('Label','State'),
                                    ('Alt',alt),       
                                    {'setActionCommand':'alt+'+s.strip(),'addActionListener':(listener)} 
                                    ), 
            0,
            ('control_List{}'.format(s.strip()),"ListBox",0,      
                                    'tab4',0,50,18,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':(items,0),'SelectedItems':(sel,),'addItemListener':listener}
                                    ),
            21, 
            ('controlFL{0}'.format(s.strip()),"FixedLine",0,         
                                'tab0-max',0,200,1,   
                                (),
                                (),                                                  
                                {} 
                                ) ,  
        
            6,
            
            ])

        tabs = {
                 0 : (None, 10),
                 1 : (None, 5),
                 2 : (None, 5),
                 3 : (None, 5),
                 4 : (None, 0),
                 }
         
        abstand_links = 10
         
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
         
        listener.tabs = tabs3
        
        return controls2,max_breite
            
            
    def dialog_shortcuts(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Listener_Shortcuts(self.mb)             

            controls,max_breite = self.dialog_shortcuts_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls) 
            self.mb.class_Einstellungen.container_anpassen(self.container,max_breite,max_hoehe,self.fenster)            
            
            listboxen = []
            conts = {}
            
            # Controls in Hauptfenster eintragen
            # Listboxen und controls an listener uebergeben
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                if 'control_List' in c:
                    name = c.split('control_List')[1]
                    listboxen.append([ctrls[c],name])
                    
                conts.update({c.split('control')[1]:ctrls[c]})
                    
            listener.listboxen = listboxen
            listener.ctrls = conts
            
        except:
            log(inspect.stack,tb())
    
    
    def dialog_templates_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        templ = self.mb.settings_orga['templates_organon']
        pfad = templ['pfad']
        items = tuple(templ['templates'])
        
        controls = (
            10,
            ('control',"FixedText",1,         
                                    'tab0',0,250,20,  
                                    ('Label','FontWeight'),
                                    (LANG.TEMPLATES_ORGANON ,150),                                             
                                    {} 
                                    ),
            50,
            
            ('control1',"FixedText",1,         
                                    'tab0',0,250,20,  
                                    ('Label',),
                                    (LANG.VORLAGENORDNER ,),                                             
                                    {} 
                                    ),
            20,
            ('control2',"Button",1,        
                                    'tab0',0,150,25,   
                                    ('Label',),
                                    (LANG.ORDNER_AUSSUCHEN,),                                           
                                    {'setActionCommand':'ordner','addActionListener':(listener)} 
                                    ),
            30,
            ('controlpfad',"FixedText",1,        
                                    'tab0',0,500,20,  
                                    ('Label',),
                                    (LANG.PFAD+': ' + pfad,),                                             
                                    {} 
                                    ),
            60,
            
            ('control3',"FixedText",1,        
                                    'tab0',0,400,20,   
                                    ('Label',),
                                    (LANG.AKT_PRJ_ALS_TEMPL,),                                                                            
                                    {} 
                                    ), 
            20,
            ('controlspeicherntxt',"Edit",0,        
                                    'tab0',0,150,20,   
                                    (),
                                    (),                                                                            
                                    {} 
                                    ), 
            25,
            ('control7',"Button",1,        
                                    'tab0',0,150,25,   
                                    ('Label',),
                                    (LANG.SPEICHERN,),                                           
                                    {'setActionCommand':'speichern','addActionListener':(listener)} 
                                    ),
            60,
            
            ('control5',"FixedText",1,        
                                    'tab0',0,400,20,   
                                    ('Label',),
                                    (LANG.TEMPLATE_LOESCHEN,),                                                                            
                                    {} 
                                    ),              
            20,
            ('control8',"ListBox",0,       
                                    'tab0',0,150,20,   
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':(items,0),'addItemListener':listener}
                                    ), 
            0,
            ('control10',"Button",1,          
                                    'tab4',0,80,25,    
                                    ('Label',),
                                    (LANG.LOESCHEN,),                                                                         
                                    {'setActionCommand':'loeschen','addActionListener':(listener)} 
                                    ), 
            
            )
        
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 1),
                 1 : (None, 5),
                 2 : (None, 5),
                 3 : (None, 5),
                 4 : (None, 5),
                 }
        
        abstand_links = 10
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                
        return controls2,max_breite
            
            
    def dialog_templates(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            listener = Listener_Templates(self.mb)

            controls,max_breite = self.dialog_templates_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)


            # UEBERGABE AN LISTENER
            listener.ctrls = {
                              'pfad': ctrls['controlpfad'],
                              'templates': ctrls['control8'],
                              'speichern' : ctrls['controlspeicherntxt']
                              }
             
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
            
        except:
            log(inspect.stack,tb())
            

    def dialog_tags_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        
        type = {
             'txt':'TEXT',
             'img':'BILD',
             'tag':'TAG',
             'date':'DATUM',
             'time':'ZEIT'
             }
        
        items = (
                 LANG.TEXT,
                 LANG.TAG,
                 LANG.BILD,
                 LANG.DATUM,
                 LANG.ZEIT
                 )
        
        datum_items = self.mb.class_Einstellungen.datum_items
        akt_datum_format = datum_items[tuple(self.mb.settings_proj['datum_format'])]
        d_items = tuple(self.mb.class_Einstellungen.datum_items2)
        d_sel = d_items.index(akt_datum_format)
        
        url_pfeil_hoch = 'private:graphicrepository/res/commandimagelist/sc_arrowshapes.up-arrow.png'
        if self.mb.programm == 'LibreOffice':
            url_pfeil_hoch = 'private:graphicrepository/cmd/sc_arrowshapes.up-arrow.png'
        url_pfeil_runter = 'private:graphicrepository/res/commandimagelist/sc_arrowshapes.down-arrow.png'
        if self.mb.programm == 'LibreOffice':
            url_pfeil_runter = 'private:graphicrepository/cmd/sc_arrowshapes.down-arrow.png'
        
        btn_breite = 110
        btn_breite2 = 60
        btn_breite3 = 100


        controls = [
            10,
             
            ('controla',"FixedText",1,         
                                    'tab1',0,80,20,  
                                    ('Label','FontWeight'),
                                    (LANG.KATEGORIE ,150),                                             
                                    {} 
                                    ),
            0, 
            ('controlb',"FixedText",1,         
                                    'tab2',0,40,20,  
                                    ('Label','FontWeight'),
                                    (LANG.TYP ,150),                                             
                                    {} 
                                    ),
            0, 
            ('controlc',"FixedText",1,         
                                    'tab3',0,60,20,  
                                    ('Label','FontWeight'),
                                    (LANG.BREITE ,150),                                             
                                    {} 
                                    ),
            25, 
            ]
                
        
        tags = self.mb.tags
        
        y = 80
        
        for abf in tags['abfolge']:
            
            name,typ = tags['nr_name'][abf]
            breite = str(tags['nr_breite'][abf])
            controls.extend(
            
                [
                 ('control_radio{}'.format(abf),"RadioButton",1,     
                                        'tab0',0,10,20,  
                                        ('VerticalAlign',),
                                        (1,),                                             
                                        {} 
                                        ),
                 0,
                 ('control_name{}'.format(abf),"Edit",0,      
                                        'tab1',0,100,20,  
                                        ('Text','MaxTextLen'),
                                        (name ,100+abf),                                             
                                        {'addKeyListener':listener} 
                                        ),
                 0,
                 ('control_typ{}'.format(abf),"FixedText",1,      
                                        'tab2',0,30,20,  
                                        ('Label',),
                                        (getattr(LANG, type[typ]) ,),                                             
                                        {} 
                                        ),
                 0,
                 ('control_breite{}'.format(abf),"Edit",0,      
                                        'tab3',0,30,20,  
                                        ('Text','MaxTextLen'),
                                        (breite ,200+abf),                                             
                                        {'addKeyListener':listener} 
                                        ),
                 22,
                 
                 ])
            
            y += 30
        

        controls.extend([
       'Y=10',
        ('controld',"FixedText",1,          
                                    'tab5',0,150,20,  
                                    ('Label','FontWeight'),
                                    (LANG.NEUE_KATEGORIE ,150),                                             
                                    {} 
                                    ),
        30,
        ('controle',"FixedText",1,         
                                    'tab4',0,40,20,  
                                    ('Label',),
                                    (LANG.NAME + ':' ,),                                             
                                    {} 
                                    ),
        0,
        ('control_txt',"Edit",0,      
                                'tab5-max',0,140,20,  
                                (),
                                (),                                             
                                {} 
                                ), 
        30, 
        ('controlf',"FixedText",1,         
                                'tab4',0,30,20,  
                                ('Label',),
                                (LANG.TYP + ':' ,),                                             
                                {} 
                                ),
        0,
        ('control_typ',"ListBox",0,     
                                'tab5',0,btn_breite2,20,    
                                ('Border','Dropdown','LineCount'),
                                (2,True,15),       
                                {'addItems':(items,0),'SelectedItems':(0,)}
                                ),  
        30,
        ('controlgh',"FixedText",1,         
                                'tab4',0,30,20,  
                                ('Label',),
                                (LANG.BREITE + ':' ,),                                             
                                {} 
                                ),
        0,
        ('control_breite',"Edit",0,      
                                'tab5',0,btn_breite2,20,  
                                (),
                                (),                                             
                                {} 
                                ), 
        30,
        ('control_neue_kategorie',"Button",1,      
                                'tab5-max',0,btn_breite,25,    
                                ('Label',),
                                (LANG.NEUE_KATEGORIE,),       
                                {'setActionCommand':'neue_kategorie','addActionListener':(listener)} 
                                ),
        30,
        ('controlFa',"FixedLine",0,        
                                'tab4x-max',0,190,1,   
                                (),
                                (),                                                  
                                {} 
                                ) ,  
        10,
        #################################
        ('controlh',"FixedText",1,         
                                    'tab5',0,250,20,  
                                    ('Label','FontWeight'),
                                    (LANG.KATEGORIE_VERSCHIEBEN ,150),                                             
                                    {} 
                                    ),
        30,
        ('control_hochxxx',"Button",0,     
                                'tab5+30',0,25,25,    
                                ('ImageURL',),
                                ( url_pfeil_hoch,),   
                                {'setActionCommand':'hoch','addActionListener':(listener)} 
                                ), 
        0,
        ('control_runterxxx',"Button",0,     
                                'tab5',0,25,25,    
                                ('ImageURL',),
                                ( url_pfeil_runter,),       
                                {'setActionCommand':'runter','addActionListener':(listener)} 
                                ), 
        30,
        ('controlFb',"FixedLine",0,        
                                'tab4x-max',0,190,1,   
                                (),
                                (),                                                  
                                {} 
                                ) ,
        
        ################ Kategorie Loeschen #############
        10,
        ('control_loeschen',"Button",1,      
                                'tab5',0,btn_breite,25,    
                                ('Label',),
                                (LANG.KATEGORIE_LOESCHEN,),       
                                {'setActionCommand':'loeschen','addActionListener':(listener)} 
                                ),  
        30,
        ('controlFc',"FixedLine",0,         
                                'tab4x-max',0,190,1,   
                                (),
                                (),                                                  
                                {} 
                                ) ,
        ############# Datum Format ##################
        10, 
        ('controlg',"FixedText",1,        
                                'tab4x',0,80,20,  
                                ('Label',),
                                (LANG.DATUMSFORMAT + ':',),                                             
                                {} 
                                ),
        0,
        ('control_datum_format',"ListBox",1,      
                                'tab5+50',0,btn_breite3,20,    
                                ('Border','Dropdown','LineCount'),
                                (2,True,15),       
                                {'addItems':(d_items,0),'SelectedItems':(d_sel,0)}
                                ),
        30,
        ('controlFd',"FixedLine",0,        
                                'tab4x-max',0,190,1,   
                                (),
                                (),                                                  
                                {} 
                                ) ,
        ############### UEBERNEHMEN ###################################
        'Y=360',
        ('control_uebernehmen',"Button",1,     
                                'tab5-max',0,130,25,    
                                ('Label',),
                                (LANG.UEBERNEHMEN,),       
                                {'setActionCommand':'uebernehmen','addActionListener':(listener)} 
                                ), 
        
        ])
        
    
        

        
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 3),
                 1 : (100,  10),
                 2 : (None, 10),
                 3 : (35,   55),
                 4 : (None, 5),
                 5 : (None, 0),
                 }
        
        abstand_links = 10
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
        
        listener.tabs = tabs3        
        
        return controls2,max_breite
    
                
            
    def dialog_tags(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if self.mb.settings_exp == None:
                # Noch kein Projekt geladen
                return
            
            tags = self.mb.tags
            
            listener = Listener_Tags(self.mb,self.container,self)             

            controls,max_breite = self.dialog_tags_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)             
            self.mb.class_Einstellungen.container_anpassen(self.container,max_breite,max_hoehe,self.fenster)
 
            rows = {}
            
            # Controls in Hauptfenster eintragen
            # Listboxen und controls an listener uebergeben
            for c,i in ctrls.items():
                if 'radio' not in c:
                    self.container.addControl(c,ctrls[c])
                    
                    number = re.findall(r'\d+', c)
                    if number == []:
                        continue
                    else:
                        number = number[-1]
                    
                    name = c.replace('control_','').replace(str(number),'')
                    if number in rows:
                        rows[number].update({name:i})
                    else:
                        rows[number] = {name:i}
                        
                    
            # Radiobuttons separat einfuegen, damit sie
            # tatsaechlich als Radiobutton funktionieren
            for c,i in ctrls.items():
                if 'radio' in c:
                    self.container.addControl(c,ctrls[c])
                    
                    number = re.findall(r'\d+', c)
                    if number == []:
                        continue
                    else:
                        number = number[-1]
                    
                    name = c.replace('control_','').replace(str(number),'')
                    if number in rows:
                        rows[number].update({name:i})
                    else:
                        rows[number] = {name:i}
                        
                    rows[number]['original'] = [ int(number), tags['nr_name'][int(number)][0], tags['nr_breite'][int(number)] ]
                    rows[number]['y'] = i.PosSize.Y
            
            
            ctrls_tag_neu = {
                             'name' : ctrls['control_txt'],
                             'typ' : ctrls['control_typ'],
                             'breite' : ctrls['control_breite'],
                             }
            
            rows['0']['radio'].setState(True)  
            
            listener.rows = rows      
            listener.ctrls_tag_neu = ctrls_tag_neu
            
        except:
            log(inspect.stack,tb())
    
    

    

from com.sun.star.awt import XActionListener,XKeyListener,XFocusListener
class Listener_Tags(unohelper.Base, XActionListener,XKeyListener,XFocusListener):
    def __init__(self,mb,container,auswahl_item_listener):
        self.mb = mb
        self.container = container
        self.rows = None
        self.ctrls_typen = 'radio','name','typ','breite'
        self.ctrls_tag_neu = None
        self.auswahl_item_listener = auswahl_item_listener
        self.tabs = None
        
        self.types = {
             LANG.TEXT : 'txt',
             LANG.BILD : 'img',
             LANG.TAG : 'tag',
             LANG.DATUM : 'date',
             LANG.ZEIT : 'time'
             }
        
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        cmd = ev.ActionCommand
        
        try:
            try:
                sel_row = int([a for a in self.rows if self.rows[a]['radio'].State][0])
            except:
                # kein Eintrag selektiert
                return
            
            if cmd in ('runter','hoch'):
                self.eintrag_verschieben(sel_row, cmd)
            elif cmd == 'loeschen':
                self.kategorie_loeschen(sel_row)
            elif cmd == 'neue_kategorie':
                self.neue_kategorie(sel_row)
            elif cmd == 'uebernehmen':
                self.kategorien_uebernehmen(sel_row)
            
        except:
            log(inspect.stack,tb())        
   
   
    def eintrag_verschieben(self,sel_row,cmd):
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        
        if cmd == 'runter':
            richtung = 1
        else:
            richtung = -1
            
        if str(sel_row + richtung) not in self.rows:
            return
        
        sel_row1 = self.rows[str(sel_row)]
        sel_row2 = self.rows[str(sel_row + richtung)]
        
        ctrls1 = [sel_row1[a] for a in self.ctrls_typen]
        hoehe1 = sel_row1['y']
        ctrls2 = [sel_row2[a] for a in self.ctrls_typen]
        hoehe2 = sel_row2['y']
        
        for c in ctrls1:
            c.setPosSize(0,hoehe2,0,0,2)
        for c in ctrls2:
            c.setPosSize(0,hoehe1,0,0,2)
        
        sel_row1['y'] = hoehe2   
        sel_row2['y'] = hoehe1 
        
        a = self.rows[str(sel_row)]
        b = self.rows[str(sel_row + richtung)]
        
        self.rows[str(sel_row)] = b
        self.rows[str(sel_row + richtung)] = a
        
        
    def breite_validieren(self,breite):    
        if self.mb.debug: log(inspect.stack)
        
        try:
            return float(breite)
        except:
            pass
        
        try:
            return float(breite.replace(',','.'))
        except:
            pass
        
        return False
        
    
    def neue_kategorie(self,sel_row):
        if self.mb.debug: log(inspect.stack)
        
        name = self.ctrls_tag_neu['name'].Text
        typ = self.types[self.ctrls_tag_neu['typ'].SelectedItem]
        breite2 = self.ctrls_tag_neu['breite'].Text
        
        breite = self.breite_validieren(breite2)
        namen = [n['name'].Text for i,n in self.rows.items()]
        
        if not breite:
            self.mb.nachricht(LANG.KEINE_GUELTIGE_ZAHL.format(breite2),'warningbox')
            return
        
        if name in namen:
            self.mb.nachricht(LANG.KATEGORIE_EXISTIERT.format(name),'warningbox')
            return
        if name == '':
            self.mb.nachricht(LANG.KATEGORIE_NAMEN_EINGEBEN,'warningbox')
            return
        
        self.ctrls_neue_kategorie(sel_row,name,typ,breite)
        self.rows[str(sel_row)]['radio'].setState(True)  
        
        
    def neue_kategorie_elemente(self,name,typ,breite,row):   
        
        listener = self
        
        otype = {
             'txt':'TEXT',
             'img':'BILD',
             'tag':'TAG',
             'date':'DATUM',
             'time':'ZEIT'
             }
        
        controls = (
            
                [
                 ('control_radio{}'.format(row),"RadioButton",0,      
                                        'tab0',0,20,20,  
                                        ('VerticalAlign',),
                                        (1,),                                              
                                        {} 
                                        ),
                 0,
                 ('control_name{}'.format(row),"Edit",0,      
                                        'tab1',0,100,20,  
                                        ('Text','MaxTextLen'),
                                        (name ,100+row),                                             
                                        {'addKeyListener':listener} 
                                        ),
                 0,
                 ('control_typ{}'.format(row),"FixedText",1,      
                                        'tab2',0,30,20,  
                                        ('Label',),
                                        (getattr(LANG, otype[typ]) ,),                                             
                                        {} 
                                        ),
                 0,
                 ('control_breite{}'.format(row),"Edit",0,      
                                        'tab3',0,30,20,  
                                        ('Text','MaxTextLen'),
                                        (str(breite) ,200+row),                                             
                                        {'addKeyListener':listener} 
                                        ),
                 ])

        controls2 = []
        
        for c in controls:
            if not isinstance(c, int) and 'Y=' not in c:
                
                c2 = list(c)
                taba,tabb = c2[3],0
                nr = int(re.findall(r'\d+', taba )[0]) 
                c2[3] = self.tabs[nr] + int(tabb)
                    
                controls2.append(c2)
            else:
                controls2.append(c)
        
        return controls2
        
        
    def ctrls_neue_kategorie(self,sel_row,name,typ,breite): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            number = str(len(self.rows)) 
            
            controls = self.neue_kategorie_elemente(name,typ,breite,int(number))
            ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)             
            
            y = self.rows[ str( len(self.rows)-1 ) ]['y'] + 22
            
            # Controls in Hauptfenster eintragen
            for c,ct in ctrls.items():
                ct.setPosSize(0,y,0,0,2)
                self.container.addControl(c,ct)
                
            
            self.rows[number] = {}           
            self.rows[number]['original'] = 'neu'
            self.rows[number]['y'] = y
            self.rows[number]['breite'] = ctrls['control_breite{}'.format(number)]
            self.rows[number]['radio'] = ctrls['control_radio{}'.format(number)]
            self.rows[number]['typ'] = ctrls['control_typ{}'.format(number)]
            self.rows[number]['name'] = ctrls['control_name{}'.format(number)]
            
            
            # Radiobuttons muessen entfernt und neu erzeugt werden, damit
            # sie im Container hintereinander gesetzt werden und sie 
            # auch wie Radiobuttons funktionieren
            
            # alle Radiobuttons löschen
            radios = [self.rows[nr]['radio'] for nr in self.rows]
            for r in radios:
                r.dispose()
            
            # Radiobuttons wieder adden   
            y = 35
            controls = []
            
            abf = sorted([int(r) for r in self.rows])
            
            for r in abf:
                
                controls.extend([
                ('control_radio{}'.format(r),"RadioButton",0,      
                                10,y,20,20,  
                                ('VerticalAlign',),
                                (1,),                                             
                                {} 
                                ),
                
                ])
                y += 22
                
            ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls) 
            
            # Controls in Hauptfenster eintragen
            for c,ct in ctrls.items():
                self.container.addControl(c,ct)
                number = re.findall(r'\d+', c)[-1]
                self.rows[str(number)]['radio'] = ct
            
        except:
            log(inspect.stack,tb())
        
           
    def kategorie_loeschen(self,sel_row):
        if self.mb.debug: log(inspect.stack)
        
        
        sel_ctrls = [self.rows[str(sel_row)][a] for a in self.ctrls_typen]
        
        for c in sel_ctrls:
            c.dispose()
            
        nachfolger = sorted([int(a) for a in self.rows if int(a) > sel_row]) 
        
        if nachfolger == []:
            del self.rows[str(sel_row)] 
            
            # neue Selektion
            if str(sel_row-1) in self.rows:
                self.rows[str(sel_row-1)]['radio'].setState(True)  
            elif '0' in self.rows:
                self.rows['0']['radio'].setState(True)
            
            return 
               
        neue_hoehen = {a:self.rows[str(a-1)]['y'] for a in nachfolger}
        
        # y der ctrls und im dict neu setzen   
        for nr,y in neue_hoehen.items():
            self.rows[str(nr)]['y'] = y
            ctrls = [self.rows[str(nr)][a] for a in self.ctrls_typen]
            for c in ctrls:
                c.setPosSize(0,y,0,0,2)
        
        # self.rows neu setzen
        for n in nachfolger:
            self.rows[str(n-1)] = self.rows[str(n)]
        
        del self.rows[str(nachfolger[-1])]  
        
        # neue Selektion
        if str(sel_row) in self.rows:
            self.rows[str(sel_row)]['radio'].setState(True)  
        elif '0' in self.rows:
            self.rows['0']['radio'].setState(True)  
    
          
    def kategorien_uebernehmen(self,sel_row):
        if self.mb.debug: log(inspect.stack)
        
        reihen = {}
        
        for nr,ctrls in self.rows.items():
            
            reihen.update( {nr: {
            'name' : ctrls['name'].Text,
            'breite' : ctrls['breite'].Text,
            'typ' : ctrls['typ'].Text,
            'original' : ctrls['original']
            
            }})
        
        tags = self.mb.tags
        
        ord_alt_neu = {val['original'][0]:int(nr) for nr,val in reihen.items() if val['original'] != 'neu' }
        
        neu = [int(nr) for nr in reihen if reihen[nr]['original'] == 'neu']
        
        # alte Kategorien uebertragen
        tags2 = {}
        for kat in tags:
            
            if kat == 'abfolge':
                tags2[kat] = list(range(len(reihen)))
            if kat == 'name_nr':
                tags2[kat] = {eintr['name']:int(nr) for nr,eintr in reihen.items() }
            if kat == 'nr_name':
                tags2[kat] = {int(nr) : [ eintr['name'], self.types[ eintr['typ'] ] ]  for nr,eintr in reihen.items() }
            if kat == 'nr_breite':
                tags2[kat] = {int(nr):eintr['breite'] for nr,eintr in reihen.items() }
            
            if kat == 'sichtbare':
                tags2[kat] = [ord_alt_neu[s] for s in tags['sichtbare'] if s in ord_alt_neu ]
                
            if kat == 'sammlung':
                tags2[kat] = {ord_alt_neu[nr]:v for nr,v in tags['sammlung'].items() if nr in ord_alt_neu }
                geloescht = { tags['nr_name'][nr][0] : (tags['sammlung'][nr] if nr in tags['sammlung'] else '') 
                             for nr in tags['nr_name'] if nr not in ord_alt_neu }
            
            if kat == 'ordinale':
                tags2[kat] = {}
                
                for ordi,eintr in tags['ordinale'].items():
                    #print(ordi,eintr)
                    tags2[kat][ordi] = {ord_alt_neu[nr]:v for nr,v in eintr.items() if nr in ord_alt_neu }
                    
        
        # neue Kategorien einfuegen
        for nr in neu:
            
            art = self.types[ reihen[str(nr)]['typ'] ]
            
            if art == 'tag':
                tags2['sammlung'][nr] = []
                        
            if art == 'txt':
                eintrag = ''
            elif art == 'tag':
                eintrag = []
            elif art == 'date':
                eintrag = None
            elif art == 'time':
                eintrag = None
            elif art == 'img':
                eintrag = ''
                
            for ordi in tags2['ordinale']:
                tags2['ordinale'][ordi][nr] = eintrag
        
        
        # auf gueltige Eintraege ueberpruefen
        for nr,b in tags2['nr_breite'].items():
            breite = self.breite_validieren(b)
            if not breite:
                self.mb.nachricht(LANG.KEINE_GUELTIGE_ZAHL.format(b),'warningbox')
                return
            tags2['nr_breite'][nr] = breite
            
        for name,nr in tags2['name_nr'].items():
            if name == '':
                self.mb.nachricht(LANG.KATEGORIE_UNGUELTIG.format(nr+1),'warningbox')   
                return
        
 
        if geloescht != {}:
            
            txt = []
            
            for name,value in geloescht.items():
                txt.append(name)
                if value == '': 
                    txt.append('\r\n')
                    continue
                txt.append('\r\n     (')
                txt.append(LANG.VORHANDENE_TAGS + ':  ')
                for v in value:
                    txt.append(v)
                    txt.append(', ')
                txt.append(')\r\n')
            
            txt2 = ''.join(txt)
            
            # Info und Nachfrage
            entscheidung = self.mb.nachricht(LANG.KATEGORIE_UEBERNEHMEN.format(txt2),"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return

        # Tags uebernehmen
        self.mb.tags = tags2
        self.mb.class_Tags.speicher_tags()
        
        # neue Kategorien auf sichtbar setzen
        for n in neu:
            if n not in self.mb.tags['sichtbare']:
                self.mb.tags['sichtbare'].append(n)
        
        # Neues Datumsformat speichern
        datum_format = self.container.getControl('control_datum_format').SelectedItem
        dat_format_neu = self.mb.class_Einstellungen.datum_items2[datum_format]
        if dat_format_neu != tuple(self.mb.settings_proj['datum_format']):
            self.mb.settings_proj['datum_format'] = list(dat_format_neu)
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        
        # Ansicht erneuern
        for c in self.container.getControls():
            c.dispose()   
        self.auswahl_item_listener.dialog_tags()
        
        self.mb.class_Sidebar.offen = {ord_alt_neu[o]:v for o,v in self.mb.class_Sidebar.offen.items() if o in ord_alt_neu}
        for n in neu:
            self.mb.class_Sidebar.offen.update({n:1})
        self.mb.class_Sidebar.erzeuge_sb_layout()
        
        # Popup Info
        win = self.mb.class_Einstellungen.haupt_fenster
        self.mb.popup(LANG.AENDERUNGEN_UEBERNOMMEN,1,win)
        
         
    def disposing(self,ev):
        return False


class Listener_Logging_Einstellungen(unohelper.Base, XActionListener):
    def __init__(self,mb,control_log,control_arg,control_filepath):
        self.mb = mb
        self.control_log = control_log
        self.control_arg = control_arg
        self.control_filepath = control_filepath
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            if ev.ActionCommand == 'Konsole':
                
                self.control_log.Enable = (ev.Source.State == 1)
                self.control_arg.Enable = (ev.Source.State == 1)
                self.mb.class_Log.output_console = ev.Source.State
                self.mb.settings_orga['log_config']['output_console'] = ev.Source.State
                self.mb.schreibe_settings_orga()
                self.mb.debug = ev.Source.State
                
            elif ev.ActionCommand == 'Logdatei':
                self.mb.class_Log.write_debug_file = ev.Source.State
                self.mb.settings_orga['log_config']['write_debug_file'] = ev.Source.State
                self.mb.schreibe_settings_orga()
                
            elif ev.ActionCommand == 'Argumente':
                self.mb.class_Log.log_args = ev.Source.State
                self.mb.settings_orga['log_config']['log_args'] = ev.Source.State
                self.mb.schreibe_settings_orga()
                
            elif ev.ActionCommand == 'File':
                Folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
                Folderpicker.execute()
                
                if Folderpicker.Directory == '':
                    return
                filepath = uno.fileUrlToSystemPath(Folderpicker.getDirectory())
                
                self.mb.class_Log.location_debug_file = filepath
                self.control_filepath.Model.Label = filepath
                
                self.mb.settings_orga['log_config']['location_debug_file'] = filepath
                self.mb.schreibe_settings_orga()
   
        except:
            log(inspect.stack,tb())
    
    def disposing(self,ev):
        return False     
    
    
class Listener_Mausrad_Einstellungen(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            self.mb.settings_orga['mausrad'] = ev.Source.State == 1
            self.mb.nutze_mausrad = ev.Source.State == 1
            self.mb.schreibe_settings_orga()
        except:
            log(inspect.stack,tb())
    
    def disposing(self,ev):
        return False    
    
    
class Listener_HTML_Export_Einstellungen(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            if ev.Source.State == 1:
                state = 1
            else:
                state = 0
            self.mb.settings_exp['html_export'][ev.ActionCommand] = state
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
        except:
            log(inspect.stack,tb())
    
    def disposing(self,ev):
        return False    
    
    
class Listener_Trenner(unohelper.Base,XItemListener,XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.conts = {}
        self.url_textfeld = None
        
        
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)

        try:
            for name in self.conts['controls']:
                if ev.Source == self.conts['controls'][name]:
                    self.conts['controls'][name].State = 1
                    self.mb.settings_orga['trenner']['trenner'] = name
                    #self.aktiviere_controls(self.conts[name])
                else:
                    self.conts['controls'][name].State = 0
                    #self.deaktiviere_controls(self.conts[name])
            
            self.mb.schreibe_settings_orga() 
        except:
            log(inspect.stack,tb())

    def disposing(self,ev):
        return False
    
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        ofilter = ('Image','*.jpg;*.JPG;*.png;*.PNG;*.gif;*.GIF')
        filepath,ok = self.mb.class_Funktionen.filepicker2(ofilter=ofilter,url_to_sys=True)
        
        if not ok:
            return
            
        self.url_textfeld.Model.Label = filepath
        
        self.mb.settings_orga['trenner']['trenner_user_url'] = Filepicker.Files[0]
        self.mb.schreibe_settings_orga()   
                
                
                
class Listener_Shortcuts(unohelper.Base,XItemListener,XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.listboxen = None
        self.ctrls = None
        
        
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)

        try:
            for l in self.listboxen:
                if l[0] == ev.Source:
                    cmd = l[1]
                    break
            
            self.shortcut_loeschen(cmd)
            
            selektiert = ev.Source.getItem(ev.Selected)
            mods = self.mb.class_Shortcuts.get_mods(cmd,self.ctrls)
            
            if selektiert != '-':
                self.mb.settings_orga['shortcuts'][str(mods)].update({selektiert:cmd})
            
            self.mb.schreibe_settings_orga() 
        except:
            log(inspect.stack,tb())

    def disposing(self,ev):
        return False
    
    def shortcut_loeschen(self,cmd):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['shortcuts']
        for m in sett:
            for n in sett[m]:
                if sett[m][n] == cmd:
                    del(sett[m][n])
                    self.mb.schreibe_settings_orga() 
                    return
             
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            cmd = ev.ActionCommand.split('+')[1]
            mods = self.mb.class_Shortcuts.get_mods(cmd,self.ctrls)
            
            self.shortcut_loeschen(cmd)

            uebrige = self.mb.class_Shortcuts.get_moegliche_shortcuts(mods)
            
            listbox = self.ctrls['_List'+cmd]
            
            listbox.removeItems(0,listbox.ItemCount)
            listbox.addItems(uebrige,0)
            listbox.selectItemPos(0,True)                      
        except:
            log(inspect.stack,tb())


class Uebersetzungen():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb


    def run(self):
        if self.mb.debug: log(inspect.stack)
        
        organon_lang_files = self.get_organon_lang_files()
        lang_akt = self.langpy_auslesen()
        
        self.cont,self.container,self.fenster = self.container_erstellen()
        ctrls_ueber,ctrls,ctrls_konst = self.erstelle_Uebersetzungsfenster(lang_akt)
        self.dialog_uebersetzung(ctrls_ueber,lang_akt,organon_lang_files,ctrls_konst,self.fenster)
         
        self.gefaerbte = list(ctrls_ueber.values())

    def langpy_auslesen(self):
        if self.mb.debug: log(inspect.stack)
        
        def formatiere_txt(txt):
            
            try:
                
                t = txt.strip()
                
                if "'''" in t:
                    if  t.count("'''") == 1:
                        if "u'''" in t:
                            t = t + "'''"  
                        else:
                            t = "u'''" + t
                elif ' u"' in t:
                    pass
                elif "u'" in t:
                    t = re.sub("u'", "u'''", t)
                    t = t + "''"
                else:
                    t = "u'''" + t + "'''"

                txt2 = eval(t)
                
                return txt2
            except:
                log(inspect.stack,tb())
                return 'ERROR'
        
        
        
        pfad_imp = os.path.join(self.mb.path_to_extension,'languages','lang_en.py')
        with codecs_open(pfad_imp , "r",'utf-8') as f:
            lines = f.readlines()
        
        lines = [ l.strip() for l in lines[1:] ]                
        kinder = {'aufpasser':0,'ist_kind' : False}
        
        def aufpasser(zaehler,zeile):
            ind = zaehler - kinder['aufpasser']
                    
            if ind not in kinder:
                kinder.update({ind:[formatiere_txt(zeile),zaehler-1]})
                #print(zaehler-1,formatiere_txt(zeile))
                #time.sleep(.1)
            else:
                kinder[ind].append(formatiere_txt(zeile))
            
            kinder['aufpasser'] += 1
            
        
        def eintragen(value,zaehler):
            
            try:
                line = value.split('=')
                odic = {}

                if len(line) == 1:
                    zeile = line[0]
                    
                    if zeile.strip() == '':
                        zeile_danach = lines[zaehler+1]
                        zeile_d = '#' in zeile_danach or '=' in zeile_danach
                        
                        if kinder['ist_kind'] and not zeile_d:
                            aufpasser(zaehler,u' ')
                        else:
                            odic.update({'art':'leer','txt':''})
                            
                    elif zeile[0] == '#':
                        odic.update({'art':'kommentar','txt':zeile.replace('#','').strip()})
                    else:
                        aufpasser(zaehler,zeile)
                        
                    if zeile[-3:] == "'''":
                        kinder['ist_kind'] = False    
                    else:
                        kinder['ist_kind'] = True  
                else:
                    odic.update( {'art':'uebersetzung',
                                 'txt':formatiere_txt(line[1]),
                                 'konst':line[0].strip(),
                                 'kinder':[]
                                 })  
                    
                    if line[1].strip()[-3:] == "'''":
                        kinder['ist_kind'] = False    
                    else:
                        kinder['ist_kind'] = True  
                
                return odic 
            except:
                log(inspect.stack,tb())
                  
        
        lines_dic = { i:eintragen(l,i) for i,l in enumerate(lines)}
        lines_dic = {i:lines_dic[i] for i in lines_dic if lines_dic[i] != {} }
        
        del kinder['aufpasser']
        del kinder['ist_kind']
        
        for k in kinder:
            z = kinder[k][1]
            kinder[k].pop(1)

            lines_dic[z]['kinder'].extend(kinder[k])
        
        return lines_dic
    
    
    
    
    
    def dialog_uebersetzung_elemente(self,button_listener,organon_lang_files):
        if self.mb.debug: log(inspect.stack)            
        
        os_path = os.path.join(self.mb.path_to_extension,'languages')
        path_orga_lang = LANG.ORGANON_LANG_PATH + '\r\n\r\n' + os_path
        
        fensterbreite = 800
        
        items = 'lang_en','lang_de'
        sel = 0
        
        tab2 = 180
        breite = 160
        
        
        controls = [
            10,
            ('control_ref',"FixedText",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.REFERENZ,),                  
                                    {} 
                                    ), 
            20,
            ('control_ref_lb',"ListBox",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':(items,0),'SelectedItems':(sel,),
                                     'addItemListener':button_listener}
                                    ), 
            50,
            ('control_odl',"FixedText",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.ORG_SPRACHDATEI_LADEN,),                  
                                    {} 
                                    ), 
            20,
            ('control_orga',"ListBox",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':(tuple(organon_lang_files),0),
                                     'SelectedItems':(0,),
                                     'addItemListener':button_listener}
                                    ), 
            50,
            ('control_bdl',"Button",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.BENUTZER_DATEI_LADEN,),                  
                                    {'setActionCommand':'nutzer_laden',
                                     'addActionListener':(button_listener)}
                                    ), 
            60,
            ('control_txt',"Edit",0,        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('HelpText',),
                                    (LANG.EXPORTNAMEN_EINGEBEN,),                  
                                    {}
                                    ), 
            
            30,
            ('control_speichern',"Button",0,       
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.UEBERSETZUNG_SPEICHERN,),                  
                                    {'setActionCommand':'speichern',
                                     'addActionListener':(button_listener)}
                                    ),  
            100,
            ('control_path',"Edit", 0,       
                                    fensterbreite - tab2,0,breite,25,     
                                    ('Text','TextColor','MultiLine',
                                    'Border','BackgroundColor'), 
                                    (path_orga_lang,KONST.FARBE_SCHRIFT_DATEI, True,
                                     0,KONST.FARBE_HF_HINTERGRUND),             
                                    {} 
                                    ),  
            ]

            
        return controls

 
    def dialog_uebersetzung(self,ctrls_ueber,lang_akt,organon_lang_files,ctrls_konst,fenster_cont): 
        if self.mb.debug: log(inspect.stack)
        
        # Listener erzeugen 
        button_listener = Uebersetzung_Button_Listener(self.mb,ctrls_ueber,lang_akt,self,fenster_cont)        
        
        controls = self.dialog_uebersetzung_elemente(button_listener,organon_lang_files)
        ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)                
          
        # Controls in Hauptfenster eintragen
        for c in ctrls:
            ctrls[c].Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            ct = ['control_' + f for f in ('ref','ref_lb','odl','txt','orga')]
            if c in ct:
                ctrls[c].Model.TextColor = KONST.FARBE_SCHRIFT_DATEI
            self.cont.addControl(c,ctrls[c])
        
        ctrls['control_txt'].Model.BorderColor = KONST.FARBE_GEZOGENE_ZEILE
        
        button_listener.ctrls = ctrls_konst
        button_listener.titel_feld = ctrls['control_txt']
        button_listener.listbox_ref = [ctrls['control_ref_lb'],
                                       ctrls['control_orga'] ]
        
        mS = ctrls['control_path'].getMinimumSize()
        H = mS.Height
        ctrls['control_path'].setPosSize(0,0,0,H,8)
        
        
        # Scrollbar hinzufuegen
        top_h = self.mb.topWindow.PosSize.Height
        prozent = top_h - (top_h /10*9)

        fensterhoehe = top_h - prozent
        
        PosSize = 0,0,20,fensterhoehe
        scrollb = self.mb.class_Fenster.erzeuge_Scrollbar(self.cont,PosSize,self.container)
        self.mb.class_Mausrad.registriere_Maus_Focus_Listener(self.cont)

                    
    def container_erstellen(self):
        if self.mb.debug: log(inspect.stack)
        
        top_h = self.mb.topWindow.PosSize.Height
        prozent = top_h - (top_h /10*9)

        fensterhoehe = top_h - prozent
        fensterbreite = 800
        posSize = (self.mb.win.Size.Width,prozent / 2,fensterbreite,fensterhoehe)

        win, cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
        cont.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        
        container, model = self.mb.createControl(self.mb.ctx, "Container", 20,0,fensterbreite - 220, 1000, (), ())
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        cont.addControl('Container_innen',container)
        
        return cont,container,win
    
    
    def erstelle_Uebersetzungsfenster(self,odic):
        if self.mb.debug: log(inspect.stack)
                        
        ctrls = {}
        ctrls_konst = {}
        ctrls_ueber = {}
        
        def pref_size(ctrl):
            mS = ctrl.getMinimumSize()
            H = mS.Height
            # Effizienz:
            # nur bei der rueckgabe setzen
            # vorher mit if h > 20 abfragen
            ctrl.setPosSize(0,0,0,H,8)
            return H
        
        
        def get_txt_neu(konst):
            faerben = False
            try:
                txt = getattr(lang_user, konst)
            except:
                txt = '*******'
            if txt == '*******':
                faerben = True
            return faerben,txt
                           
        
        y = 10   
        try:
            for o in sorted(odic):
                
                art = odic[o]['art']
                
                if art == 'kommentar':
                    txt = odic[o]['txt']
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", 20,y+10,200, 20, 
                                                           ('Label','TextColor','Weight'), 
                                                           (txt,KONST.FARBE_SCHRIFT_ORDNER,150))
                    self.container.addControl(o, control)
                    ctrls.update({o:control})
                    
                    y += 35
                else:

                    if art == 'uebersetzung':
                        
                        konst = odic[o]['konst']
                        txt = [odic[o]['txt']]
                        
                        for k in odic[o]['kinder']:
                            txt.append(k)
                        
                        txt = '\r\n'.join(txt)
                         
                        control, model = self.mb.createControl(self.mb.ctx, "Edit", 20,y,500, 20, 
                                                               ('Text','TextColor','MultiLine',
                                                                'Border','BackgroundColor'), 
                                                               ( txt,KONST.FARBE_SCHRIFT_DATEI, True,
                                                                 0,KONST.FARBE_HF_HINTERGRUND) )
                        
                        self.container.addControl(str(o)+'txt', control)
                        ctrls.update({o:control})
                        ctrls_konst.update({konst:control})
                        
                        hoehe = pref_size(control)
                         
                        y += hoehe + 1
                                                        
                        control, model = self.mb.createControl(self.mb.ctx, "Edit", 40,y,500, hoehe +5, 
                                                                       ('Text','TextColor','Border',
                                                                        'MultiLine','BorderColor'), 
                                                                       ('*******',KONST.FARBE_SCHRIFT_DATEI,1,
                                                                        True,KONST.FARBE_GEZOGENE_ZEILE) )
                        model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE
                        self.container.addControl(str(o)+'ueber', control)

                        ctrls_ueber.update({konst:control})
                        
                        y += hoehe + 1 + 5
                        

            self.container.setPosSize(0,0,0,y,8)
            
            return ctrls_ueber,ctrls,ctrls_konst
        
        except:
            log(inspect.stack,tb())
    
    
    def get_organon_lang_files(self):
        if self.mb.debug: log(inspect.stack)
        
        datei_pfad = os.path.join(self.mb.path_to_extension,'languages')
        
        files = []
        for (dirpath, dirnames, filenames) in os.walk(datei_pfad):
            break
        
        for f in filenames:
            endung = f.split('.')[1]

            if 'lang_' in f and endung == 'py' and len(f) == 10:
                files.append(f.split('.')[0])
        return files
            
        
    def lade_uebersetzung(self,ctrls_ueber,lang_user):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            for c in self.gefaerbte:
                c.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            self.gefaerbte = []
            
            for konst in ctrls_ueber:

                faerben = False
                try:
                    txt = getattr(lang_user, konst)
                except:
                    txt = '*******'
                    
                if txt == '*******':
                    faerben = True
                
                
                ctrls_ueber[konst].Model.Text = txt
                if faerben:
                    ctrls_ueber[konst].Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE
                    self.gefaerbte.append(ctrls_ueber[konst])
                
        except:
            log(inspect.stack,tb())     
        
    def lade_modul_aus_datei(self,pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            global_env1 = {}
            helfer = {}
            code = "class Helfer():pass"
            code = compile(code, "new_lang", "exec")
            exec(code, global_env1, helfer)
            
            with codecs_open(pfad , "r",'utf-8') as fi:
                f = fi.readlines()
                f = ''.join(f[1:])
            
            global_env2 = {}
            local_env = {}
            code2 = compile(f, "new_lang", "exec")
            exec(code2, global_env2, local_env)
            
            lang_user = helfer['Helfer']
            for k,v in local_env.items():
                setattr(lang_user, k, v)
                
            return lang_user,True
        
        except Exception as e:
            try:
                self.mb.nachricht(LANG.IMPORT_GESCHEITERT.format(str(e)))
            except:
                pass
            return None,False
    
    
    def lang_eintrag_ueberpruefen(self,neuer_eintrag,konst):
        
        if neuer_eintrag == '*******':
            return True
        if neuer_eintrag == '':
            return True
        
        eintrag = getattr(LANG, konst)
        
        verbotene = ['=','#','\\']
        
        for v in verbotene:
            if v in neuer_eintrag:
                if konst != 'UNGUELTIGE_ZEICHEN':
                    self.mb.nachricht(LANG.EINTRAG_MIT_UNZULAESSIGEN_ZEICHEN.format(v,neuer_eintrag))
                    return False

        if eintrag.count('%s') != neuer_eintrag.count('%s'):
            self.mb.nachricht(LANG.UNGUELTIGE_ANZAHL1.format(neuer_eintrag))

        expr = '\{\d?\}' 
         
        if len(re.findall(expr,eintrag)) != len(re.findall(expr,neuer_eintrag)):
            self.mb.nachricht(LANG.UNGUELTIGE_ANZAHL2.format(neuer_eintrag))

        return True
        
    
        
    def format_lang_user_an_lang_akt_anpassen(self,ctrls_ueber,lang_en):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # auslesen
            
            new_dict = {}
            
            for c in ctrls_ueber:
                txt = ctrls_ueber[c].Text
                new_dict.update( {c: {'txt':txt,'kinder':[] } })
            
            # eintragen
            
            new_lang = copy.deepcopy(lang_en)
            
            for l in new_lang:
                if 'konst' in new_lang[l]:
                    konst = new_lang[l]['konst'] 
                    ok = self.lang_eintrag_ueberpruefen(new_dict[konst]['txt'],konst)
                    if not ok:
                        return None, False
                    new_lang[l]['txt'] = new_dict[konst]['txt']                    
            
            return new_lang, True
            
        except:
            log(inspect.stack,tb())                
        



from com.sun.star.awt import XAdjustmentListener
class Scroll_Listener (unohelper.Base,XAdjustmentListener):
    def __init__(self,mb,fenster_cont):   
        self.fenster_cont = fenster_cont
    def adjustmentValueChanged(self,ev):
        try:
            self.fenster_cont.setPosSize(0, -ev.Value,0,0,2)
        except Exception as e:
            print(e)
    def disposing(self,ev):
        return False
    
 
    
class Uebersetzung_Button_Listener(unohelper.Base, XActionListener,XItemListener):
    def __init__(self,mb,ctrls_ueber,lang_akt,class_Uebersetzung,fenster_cont):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb  
        self.ctrls = None
        self.ctrls_ueber = ctrls_ueber
        self.lang_akt = lang_akt
        self.titel_feld = None
        self.listbox_ref = None
        
        self.class_Uebersetzung = class_Uebersetzung
        
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        cmd = ev.ActionCommand
        try:
            if cmd == 'speichern':
                new_lang,ok = self.class_Uebersetzung.format_lang_user_an_lang_akt_anpassen(self.ctrls_ueber,self.lang_akt)
                if ok:
                    self.lang_speichern(new_lang)
                    
            if cmd == 'orga_laden':
                self.lang_user_laden()
            if cmd == 'nutzer_laden':
                self.lang_user_laden(False)
                
        except:
            log(inspect.stack,tb())
            
    def itemStateChanged(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        sel = ev.Source.SelectedItem

        if ev.Source == self.listbox_ref[0]:
            self.referenz_laden(sel)
        if ev.Source == self.listbox_ref[1]:
            datei_pfad = os.path.join(self.mb.path_to_extension,'languages',sel+'.py')
            lang_user,ok = self.class_Uebersetzung.lade_modul_aus_datei(datei_pfad)
            
            if not ok:
                return
            
            self.class_Uebersetzung.lade_uebersetzung(self.ctrls_ueber,lang_user)
        else:
            self.sprachdatei_setzen(sel)
        
    
#     def sprachdatei_setzen(self,sel):
#         return
#         if self.mb.debug: log(inspect.stack)
#         
#         try:
#             konsts = [l for l in dir(LANG) if l.isupper()]
#             
#             datei_pfad = os.path.join(self.mb.path_to_extension,'languages',sel+'.py')
#             lang_ref,ok = self.mb.class_Projekt.lade_modul_aus_datei(datei_pfad)
#             
#             if not ok:
#                 return
#             
#             omb = self.mb.props[T.AB].Hauptfeld.Context.Context.getControl('Organon_Menu_Bar')
#             
#             lang_file       = getattr(lang_ref, 'FILE')
#             lang_bearbeiten = getattr(lang_ref, 'BEARBEITEN_M')
#             lang_ansicht    = getattr(lang_ref, 'ANSICHT')
#             
#             #help_ordner     = getattr(lang_ref, 'INSERT_DIR')
#             #help_datei      = getattr(lang_ref, 'INSERT_DOC')
#             #help_speichern  = getattr(lang_ref, 'FORMATIERUNG_SPEICHERN')
#             
#             ctrl_file       = omb.getControl(LANG.FILE)
#             ctrl_bearbeiten = omb.getControl(LANG.BEARBEITEN_M)
#             ctrl_ansicht    = omb.getControl(LANG.ANSICHT)
#             #ctrl_ordner
#             #ctrl_datei
#             #ctrl_speichern
#             
#             
#             ctrl_file.Model.Label       = lang_file
#             ctrl_bearbeiten.Model.Label = lang_bearbeiten
#             ctrl_ansicht.Model.Label    = lang_ansicht
#             
#             
# #             Menueintraege = [
# #                 (LANG.FILE,'a'),
# #                 (LANG.BEARBEITEN_M,'a'),
# #                 (LANG.ANSICHT,'a'),            
# #                 ('Ordner','b',KONST.IMG_ORDNER_NEU_24,LANG.INSERT_DIR),
# #                 ('Datei','b',KONST.IMG_DATEI_NEU_24,LANG.INSERT_DOC),
# #                 ('Speichern','b','vnd.sun.star.extension://xaver.roemers.organon/img/lc_save.png',
# #                  LANG.FORMATIERUNG_SPEICHERN.format(LANG.KEINE))
# #                 ]
# 
#             for k in konsts:
#                 value = getattr(lang_ref, k)
#                 setattr(LANG, k, value)
#             
#      
#         except:
#             log(inspect.stack,tb())
       
        
    def referenz_laden(self,sel):
        if self.mb.debug: log(inspect.stack)
        
        try:
            datei_pfad = os.path.join(self.mb.path_to_extension,'languages',sel+'.py')
            lang_ref,ok = self.class_Uebersetzung.lade_modul_aus_datei(datei_pfad)
            
            if not ok:
                return
            
            for konst in self.ctrls:
                self.ctrls[konst].Model.Text = getattr(lang_ref, konst)
        except:
            log(inspect.stack,tb())  

            
    def lang_user_laden(self,use_path=True):
        if self.mb.debug: log(inspect.stack)
        
        if use_path:
            pfad = os.path.join(self.mb.path_to_extension,'languages')
            datei_pfad = self.mb.class_Funktionen.filepicker(filepath=pfad,ofilter='py')
        else:
            datei_pfad = self.mb.class_Funktionen.filepicker(ofilter='py')
        if datei_pfad:
            lang_user,ok = self.class_Uebersetzung.lade_modul_aus_datei(datei_pfad)
            
            if not ok:
                return
            
            self.class_Uebersetzung.lade_uebersetzung(self.ctrls_ueber,lang_user)
                
    def lang_speichern(self,new_lang):
        if self.mb.debug: log(inspect.stack)
        
        name = self.titel_feld.Model.Text
        
        if name.strip() == '':
            self.mb.nachricht(LANG.KEIN_NAME)
            return
        
        name = name + '.py'
        
        folder = self.mb.class_Funktionen.folderpicker()
        
        if not folder:
            return
        
        pfad = os.path.join(folder,name)
        
        txt_list = []
        txt_list.append('# -*- coding: utf-8 -*-\r\n\r\n')
        
        for l in sorted(new_lang):
            
            eintrag = new_lang[l]
            art = eintrag['art']
            txt = "'''\\\r\nu'''".join(eintrag['txt'].split('\r\n'))

            if art == 'kommentar':
                txt_list.append('# ' + txt + '\r\n')
            elif art == 'leer':
                txt_list.append('\r\n')
            elif art == 'uebersetzung':
                konst = new_lang[l]['konst']
                txt_list.append(konst + " = u'''" + txt  + "'''" + "\r\n")
        
        text = ''.join(txt_list)
        
        with codecs_open(pfad , "w",'utf-8') as f:
            f.write(text)  
              
    def disposing(self,ev):
        self.fenster_cont.dispose()
        pass
   
   
class Listener_Templates(unohelper.Base, XActionListener,XItemListener):
    def __init__(self,mb):
        self.mb = mb
        self.ctrls = None
        self.selektiertes_template = None
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            cmd = ev.ActionCommand
            
            if cmd == 'ordner':
                self.template_ordner_setzen()
            if cmd == 'speichern':
                self.template_speichern()
            if cmd == 'loeschen':
                self.template_loeschen()
        except:
            log(inspect.stack,tb())
    
    def itemStateChanged(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        sel = ev.Selected
        self.selektiertes_template = self.ctrls['templates'].Items[sel]
    
    def template_ordner_setzen(self):
        if self.mb.debug: log(inspect.stack)
        
        path = self.mb.class_Funktionen.folderpicker()
        if path != None:
            templs = self.mb.settings_orga['templates_organon']
            templs['pfad'] = path
            self.ctrls['pfad'].Model.Label = LANG.PFAD + ': ' + path
            
            self.update_templates()
            self.mb.schreibe_settings_orga()

    def template_speichern(self):
        if self.mb.debug: log(inspect.stack)

        txt = self.ctrls['speichern'].Model.Text
        templs = self.mb.settings_orga['templates_organon']
        pfad = templs['pfad']
        
        self.mb.class_Funktionen.vorlage_speichern(pfad,txt)
        self.update_templates()
        
    def update_templates(self):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.class_Funktionen.update_organon_templates()
        
        templs = self.mb.settings_orga['templates_organon']
        templates = tuple(templs['templates'])
        
        self.ctrls['templates'].Model.removeAllItems()
        self.ctrls['templates'].addItems(templates,0)
            
    def template_loeschen(self):
        if self.mb.debug: log(inspect.stack)
        
        if self.selektiertes_template == None:
            return
        
        templs = self.mb.settings_orga['templates_organon']
        pfad = templs['pfad']
        templ_pfad = os.path.join(pfad,self.selektiertes_template + '.organon')
        
        entscheidung = self.mb.nachricht(LANG.TEMPLATE_WIRKLICH_LOESCHEN.format(self.selektiertes_template),"warningbox",16777216)
        # 3 = Nein oder Cancel, 2 = Ja
        if entscheidung == 3:
            return
        
        import shutil
        try:
            shutil.rmtree(templ_pfad)
        except:
            log(inspect.stack,tb())
            
        self.update_templates()
        
    def disposing(self,ev):pass
    
    
    
    
    
    
    
    
    
                