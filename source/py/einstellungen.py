# -*- coding: utf-8 -*-

import unohelper
from uno import fileUrlToSystemPath
import json
import copy

class Einstellungen():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
        # prop, um den Listener fuer die Writer Design LBs
        # beim anfaenglichen Setzen nicht anzusprechen
        self.setze_listboxen = False
        # enthaelt die LB Eintraege und entsprechende Organon props
        self.lb_dict = None
        # vice versa
        self.lb_dict2 = None

        
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
            hoehe = 400
            breite_listbox = 150
            
            
            sett = self.mb.settings_exp
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = self.mb.dialog.Size.Width
            Y = posSize_main.Y            

            # Listener erzeugen 
            listener = {}           
            listener.update( {'auswahl_listener': Auswahl_Item_Listener(self.mb)} )
            
            controls = self.dialog_einstellungen(listener, breite_listbox,breite,hoehe)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)                
            
            # Hauptfenster erzeugen
            posSize = X,Y,breite,pos_y + 40
            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
            #fenster_cont.Model.Text = LANG.EXPORT
            
            #fenster_cont.addEventListener(listenerDis)
            self.haupt_fenster = fenster
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                fenster_cont.addControl(c,ctrls[c])
            
            listener['auswahl_listener'].container = ctrls['control_Container']
            
        except:
            log(inspect.stack,tb())
                
       
    
    def dialog_einstellungen(self,listener,breite_listbox,breite,hoehe):
        if self.mb.debug: log(inspect.stack)
        
        LANG.TEMPLATES = u'Templates'
        
        if self.mb.programm == 'LibreOffice':
            lb_items = (
                        LANG.DESIGN_TRENNER,
                        LANG.DESIGN_ORGANON,
                        LANG.DESIGN_PERSONA,
                        LANG.MAUSRAD,
                        #LANG.TEMPLATES,
                        LANG.SHORTCUTS,
                        LANG.HTML_EXPORT,
                        LANG.LOG
                        )
        else:
            lb_items = (
                        LANG.DESIGN_TRENNER,
                        LANG.DESIGN_ORGANON,
                        LANG.MAUSRAD,
                        #LANG.TEMPLATES,
                        LANG.SHORTCUTS,
                        LANG.HTML_EXPORT,
                        LANG.LOG
                        )
            
           
        controls = (
            10,
            ('controlE_calc',"FixedText",        
                                    20,0,50,20,    
                                    ('Label','FontWeight'),
                                    (LANG.EINSTELLUNGEN ,150),                  
                                    {} 
                                    ), 
            20,                                                  
            ('control_Liste',"ListBox",      
                                    20,0,breite_listbox,hoehe,    
                                    ('Border',),
                                    ( 2,),       
                                    {'addItems':lb_items,'addItemListener':(listener['auswahl_listener'])} 
                                    ),  
            0,
            ('control_Container',"Container",      
                                    breite_listbox + 40,0,breite-60-breite_listbox ,hoehe ,    
                                    ('BackgroundColor','Border'),
                                    (KONST.FARBE_ORGANON_FENSTER,1),              
                                    {} 
                                    ),  
            hoehe - 20,
            )
        return controls
    
    
              

from com.sun.star.awt import XItemListener   
class Auswahl_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.container = None
        
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
                self.mb.class_Organon_Design.dialog_organon_farben()
                
            elif sel == LANG.DESIGN_PERSONA:
                self.mb.class_Organon_Design.container = self.container
                self.mb.class_Organon_Design.dialog_persona()
                
            elif sel == LANG.SHORTCUTS:
                self.dialog_shortcuts()

        except:
            log(inspect.stack,tb())
            
    def disposing(self,ev):
        return False
  
        
    def dialog_logging(self):

        try:
            ctx = self.mb.ctx
            mb = self.mb
            breite = 650
            hoehe = 190
            
            tab = 20
            tab1 = 40
            
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = posSize_main.X +20
            Y = posSize_main.Y +20
            Width = breite
            Height = hoehe
            
            posSize = X,Y,Width,Height
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
            #fenster.setPosSize(0,0,Breite+10+tab1,0,4)
            
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
            ('controlE_calc',"FixedText",        
                                    20,0,50,20,    
                                    ('Label','FontWeight'),
                                    (LANG.HTML_AUSWAHL ,150),                  
                                    {} 
                                    ), 
            50, ]
        
        elemente = 'FETT','KURSIV','AUSRICHTUNG','UEBERSCHRIFT','FUSSNOTE','FARBEN','LINKS'
                # ZITATE,SCHRIFTART,SCHRIFTGROESSE,CSS
                
        for el in elemente:
            controls.extend([
            ('control_{}'.format(el),"CheckBox",      
                                    20,0,200,20,    
                                    ('Label','State'),
                                    (getattr(LANG, el),html_exp_settings[el]),       
                                    {'setActionCommand':el,'addActionListener':(listener,)} 
                                    ),  
            25])
            
        return controls

 
    def dialog_html_export(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:

            sett = self.mb.settings_exp['html_export']
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = posSize_main.X +20
            Y = posSize_main.Y +20            

            # Listener erzeugen 
            listener = Listener_HTML_Export_Einstellungen(self.mb)         
            
            controls = self.dialog_html_export_elemente(listener,sett)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)                
              
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                         
        except:
            log(inspect.stack,tb())
            
                    
    def dialog_mausrad_elemente(self,listener,nutze_mausrad):
        if self.mb.debug: log(inspect.stack)
        
        controls = (
            10,
            ('controlE_calc',"FixedText",        
                                    20,0,50,20,    
                                    ('Label','FontWeight'),
                                    (LANG.NUTZE_MAUSRAD ,150),                  
                                    {} 
                                    ), 
            50,                                                  
            ('control_CB_calc',"CheckBox",      
                                    20,0,200,220,    
                                    ('Label','State'),
                                    (LANG.NUTZE_MAUSRAD,nutze_mausrad),       
                                    {'addActionListener':(listener,)} 
                                    ),  
            30,
            ('control_Container_calc',"FixedText",      
                                    20,0,200,200,    
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
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = posSize_main.X +20
            Y = posSize_main.Y +20            

            # Listener erzeugen 
            listener = Listener_Mausrad_Einstellungen(self.mb)         
            
            controls = self.dialog_mausrad_elemente(listener,nutze_mausrad)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)                
             
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                         
        except:
            log(inspect.stack,tb())
            
    
    def dialog_trenner_elemente(self,trenner_dict,listener_CB,listener_URL):
        if self.mb.debug: log(inspect.stack)
        
        sett_trenner = self.mb.settings_orga['trenner']

        controls = (
            10,
            ('control',"FixedText",         
                                    'tab0',0,250,20,  
                                    ('Label','FontWeight'),
                                    (LANG.TRENNER_FORMATIERUNG ,150),                                             
                                    {} 
                                    ),
            50,
            ('control1',"CheckBox",         
                                    'tab0',0,200,20,   
                                    ('Label','State'),
                                    (LANG.LINIE,trenner_dict['strich']),                                                  
                                    {'addItemListener':(listener_CB)} 
                                    ) ,
            0,
            ('control11',"CheckBox",        
                                    'tab4x',0,150,20,   
                                    ('Label','State'),
                                    (LANG.KEIN_TRENNER,trenner_dict['keiner']),                                           
                                    {'addItemListener':(listener_CB)} 
                                    ),
            50,
            ('control2',"CheckBox",         
                                    'tab0',0,360,20,   
                                    ('Label','State'),
                                    (LANG.FARBE,trenner_dict['farbe']),                                                   
                                    {'addItemListener':(listener_CB)} 
                                    ), 
            20,
            ('control3',"FixedText",        
                                    'tab1',0,400,60,   
                                    ('Label',),
                                    (LANG.TRENNER_ANMERKUNG,),                                                                            
                                    {} 
                                    ), 
             
            60,
            ('control7',"CheckBox",         
                                    'tab0',0,80,20,   
                                    ('Label','State'),
                                    (LANG.BENUTZER,trenner_dict['user']),                                                  
                                    {'addItemListener':(listener_CB)} 
                                    ),              
            20,
            ('control8',"FixedText",        
                                    'tab1',0,20,20,   
                                    ('Label',),
                                    ('URL: ',),                                                                            
                                    {'Enable':trenner_dict['user']==1} 
                                    ), 
            0,
            ('control10',"Button",          
                                    'tab2',0,60,20,    
                                    ('Label',),
                                    (LANG.AUSWAHL,),                                                                         
                                    {'Enable':trenner_dict['user']==1,'addActionListener':(listener_URL,)} 
                                    ), 
            30,
            ('control9',"FixedText",        
                                    'tab1',-4,600,20,   
                                    ('Label',),
                                    (sett_trenner['trenner_user_url'],),                                            
                                    {'Enable':trenner_dict['user']==1} 
                                    ), 
            20,
            )
        
        # Tabs waren urspruenglich gesetzt, um sie in der Klasse Design richtig anzupassen.
        # Das fehlt. 'tab...' wird jetzt nur in Zahlen uebersetzt. Beim naechsten 
        # groesseren Fenster, das ich schreibe und das nachtraegliche Berechnungen benoetigt,
        # sollte der Code generalisiert werden. Vielleicht grundsaetzlich ein Modul fenster.py erstellen?
        tab0 = tab0x = 20
        tab1 = tab1x = 42
        tab2 = tab2x = 78
        tab3 = tab3x = 100
        tab4 = tab4x = 230
        
        tabs = ['tab0','tab1','tab2','tab3','tab4','tab0x','tab1x','tab2x','tab3x','tab4x']
        
        tabs_dict = {}
        for a in tabs:
            tabs_dict.update({a:locals()[a]  })

        controls2 = []
        
        for c in controls:
            if not isinstance(c, int):
                c2 = list(c)
                c2[2] = tabs_dict[c2[2]]
                controls2.append(c2)
            else:
                controls2.append(c)
        
        return controls2
            
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
             

            controls = self.dialog_trenner_elemente(trenner_dict,listener_CB,listener_URL)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)


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
        
        sett_trenner = self.mb.settings_orga['trenner']

        controls = [
            10,
            ('control',"FixedText",         
                                    'tab0',0,250,20,  
                                    ('Label','FontWeight'),
                                    (LANG.SHORTCUTS ,150),                                             
                                    {} 
                                    ),
            50, 
            ]
        
        
        
        from collections import OrderedDict
        
        shorts = [
                 ['INSERT_DOC' , LANG.INSERT_DOC],
                 ['INSERT_DIR' , LANG.INSERT_DIR],
                 ['TRENNE_TEXT' , LANG.TRENNE_TEXT],
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
            ('control_{}'.format(s.strip()),"FixedText",      
                                    'tab0',0,220,20,  
                                    ('Label',),
                                    (shortcuts[s] ,),                                             
                                    {} 
                                    ),  
            0,
            ('control_Shift{}'.format(s.strip()),"CheckBox",      
                                    'tab1',0,50,20,    
                                    ('Label','State'),
                                    ('Shift',shift),       
                                    {'setActionCommand':'shift+'+s.strip(),'addActionListener':(listener,)} 
                                    ),  
            0,
            ('control_Ctrl{}'.format(s.strip()),"CheckBox",      
                                    'tab2',0,50,20,    
                                    ('Label','State'),
                                    ('Ctrl',ctrl),       
                                    {'setActionCommand':'ctrl+'+s.strip(),'addActionListener':(listener,)} 
                                    ), 
            0,
            ('control_Alt{}'.format(s.strip()),"CheckBox",      
                                    'tab3',0,40,20,    
                                    ('Label','State'),
                                    ('Alt',alt),       
                                    {'setActionCommand':'alt+'+s.strip(),'addActionListener':(listener,)} 
                                    ), 
            -3,
            ('control_List{}'.format(s.strip()),"ListBox",      
                                    'tab4',0,50,18,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':items,'SelectedItems':(sel,),'addItemListener':listener}
                                    ), 
            22,
            
            
            
            ])
        
    
        
        
        # Tabs waren urspruenglich gesetzt, um sie in der Klasse Design richtig anzupassen.
        # Das fehlt. 'tab...' wird jetzt nur in Zahlen uebersetzt. Beim naechsten 
        # groesseren Fenster, das ich schreibe und das nachtraegliche Berechnungen benoetigt,
        # sollte der Code generalisiert werden. Vielleicht grundsaetzlich ein Modul fenster.py erstellen?
        tab0 = tab0x = 20
        tab1 = tab1x = 260
        tab2 = tab2x = 310
        tab3 = tab3x = 360
        tab4 = tab4x = 400
        
        tabs = ['tab0','tab1','tab2','tab3','tab4','tab0x','tab1x','tab2x','tab3x','tab4x']
        
        tabs_dict = {}
        for a in tabs:
            tabs_dict.update({a:locals()[a]  })

        controls2 = []
        
        for c in controls:
            if not isinstance(c, int):
                c2 = list(c)
                c2[2] = tabs_dict[c2[2]]
                controls2.append(c2)
            else:
                controls2.append(c)
        
        return controls2
            
    def dialog_shortcuts(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Listener_Shortcuts(self.mb)             

            controls = self.dialog_shortcuts_elemente(listener)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)             
            
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


from com.sun.star.awt import XActionListener
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
                self.mb.class_Funktionen.schreibe_settings_orga()
                self.mb.debug = ev.Source.State
                
            elif ev.ActionCommand == 'Logdatei':
                self.mb.class_Log.write_debug_file = ev.Source.State
                self.mb.settings_orga['log_config']['write_debug_file'] = ev.Source.State
                self.mb.class_Funktionen.schreibe_settings_orga()
                
            elif ev.ActionCommand == 'Argumente':
                self.mb.class_Log.log_args = ev.Source.State
                self.mb.settings_orga['log_config']['log_args'] = ev.Source.State
                self.mb.class_Funktionen.schreibe_settings_orga()
                
            elif ev.ActionCommand == 'File':
                Folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
                Folderpicker.execute()
                
                if Folderpicker.Directory == '':
                    return
                filepath = fileUrlToSystemPath(Folderpicker.getDirectory())
                
                self.mb.class_Log.location_debug_file = filepath
                self.control_filepath.Model.Label = filepath
                
                self.mb.settings_orga['log_config']['location_debug_file'] = filepath
                self.mb.class_Funktionen.schreibe_settings_orga()
   
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
            self.mb.class_Funktionen.schreibe_settings_orga()
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
            
            self.mb.class_Funktionen.schreibe_settings_orga() 
        except:
            log(inspect.stack,tb())

    def disposing(self,ev):
        return False
    
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        Filepicker.appendFilter('Image','*.jpg;*.JPG;*.png;*.PNG;*.gif;*.GIF')
        Filepicker.execute()
         
        if Filepicker.Files == '':
            return
    
        filepath =  uno.fileUrlToSystemPath(Filepicker.Files[0])
        
        self.url_textfeld.Model.Label = filepath
        
        self.mb.settings_orga['trenner']['trenner_user_url'] = Filepicker.Files[0]
        self.mb.class_Funktionen.schreibe_settings_orga()   
                
                
                
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
            
            self.mb.class_Funktionen.schreibe_settings_orga() 
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
                    self.mb.class_Funktionen.schreibe_settings_orga() 
                    return
             
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            cmd = ev.ActionCommand.split('+')[1]
            mods = self.mb.class_Shortcuts.get_mods(cmd,self.ctrls)
            
            self.shortcut_loeschen(cmd)

            uebrige = self.mb.class_Shortcuts.get_moegliche_shortcuts(mods)
            
            listbox = self.ctrls['_List'+cmd]
            item = listbox.getSelectedItem()
            
            listbox.removeItems(0,listbox.ItemCount)
            listbox.addItems(uebrige,0)
            listbox.selectItemPos(0,True)                      
        except:
            log(inspect.stack,tb())
            
# import sys 
# platform = sys.platform  
from traceback import format_exc as tb     
# def pydevBrk():  
#     # adjust your path 
#     if platform == 'linux':
#         sys.path.append('/home/xgr/.eclipse/org.eclipse.platform_4.4.1_1473617060_linux_gtk_x86_64/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
#     else:
#         sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
#     from pydevd import settrace
#     settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
# pd = pydevBrk               
    
import copy
from os import path as PATH, listdir
from codecs import open as codecs_open
class Take_Over_Old_Settings():
    '''
    Settings have to be loaded while the old extension is still available
    on the harddrive. During installation the old extension will be deleted.
    So getting both needs to be done at the very beginning.
    When an old extension is found, the newly created settings will
    be extended by the old ones.
    '''

    def get_settings_of_previous_installation(self,package_folder, extension_folder):
        try:
            dirs = [name for name in listdir(package_folder) if PATH.isdir(PATH.join(package_folder, name))]
            dirs.remove(extension_folder)
            
            files = None   
            organon_in_files = False    
             
            for d in dirs:
                files = listdir(PATH.join(package_folder,d))
                if 'organon.oxt' in files:
                    organon_in_files = True
                    break
            
            if files == None or organon_in_files == False :
                return None

            json_pfad_alt = PATH.join(package_folder,d,'organon.oxt','organon_settings.json')
            
            with codecs_open(json_pfad_alt) as data:  
                content = data.read().decode()  
                settings_orga_prev = json.loads(content)

            return settings_orga_prev
    
        except Exception as e:
            return None
    
    designs = []
    fehlende = []

    def _update_designs(self,a,b):
        fehlende = set(b['designs']).difference( set(a['designs']) )
        standard = copy.deepcopy(a['designs']['Standard'])
        for f in fehlende:
            a['designs'].update({f:standard})
        
        return list(b['designs']), fehlende
        
    
    def _compare_design(self,a1,b1): 
        for k in b1:
            if k in a1:
                if b1[k] != a1[k]:
                    return True  
        return False
    
    def _treat_design(self,a,b,key,path):
        
        if key in self.fehlende:
            # Wenn Design nur im alten Dict vorhanden war,
            # wird es direkt uebernommen
            self.merge(a[key], b[key], path + [str(key)])
        else:
            # Wenn Designs gleichen Namens sich unterscheiden,
            # wird eine neue Version "_old" eingefuegt
            ungleich = self._compare_design(a[key],b[key])
            if ungleich:
                
                k = key
                while k in a:
                    k = k + '_old'
                    
                standard = copy.deepcopy(a['Standard'])
                a[k] = standard
                self.merge(a[k], b[key], path + [str(key)])
            else:
                pass
        
        
    
    def merge(self,a, b, path=None):
        '''
        This method is an adjusted version from:
        http://stackoverflow.com/questions/7204805/python-dictionaries-of-dictionaries-merge
        merges b into a
        '''
        if path is None: path = []
        
        try:
            for key in b:
                if key in a:
                    if key == 'zuletzt_geladene_Projekte':
                        a[key] = b[key]
                    elif isinstance(a[key], dict) and isinstance(b[key], dict):
                        if key in self.designs:
                            self._treat_design(a, b, key, path)
                        elif key == 'designs':
                            self.designs,self.fehlende = self._update_designs(a,b)
                            self.merge(a[key], b[key], path + [str(key)])
                        else:
                            self.merge(a[key], b[key], path + [str(key)])
                    elif a[key] == b[key]:
                        pass # same leaf value
                    elif isinstance(a[key], list) and isinstance(b[key], list):
                        for idx, val in enumerate(b[key]):
                            a[key][idx] = self.merge(a[key][idx], b[key][idx], path + [str(key), str(idx)])
                    else:
                        # ueberschreiben der defaults mit alten Werten
                        a[key] = b[key]
                else:
                    # hier werden nur in b vorhandene keys gesetzt
                    # daher werden auch alte designs mit eigenem Namen ignoriert
                    pass
            
            return a
        except Exception as e:
            print(tb())
            return None
    
    # wird nicht verwendet
    def dict_to_list(self,odict,olist,predecessor=[]):
    
        for k in odict:
            value = odict[k]
            pre = predecessor[:]
                            
            if isinstance(value, dict):
                pre.append(k)
                self.dict_to_list(value,olist,pre)
            else:
                olist.append(predecessor+[k])
                
    # wird nicht verwendet
    def exchange_values(self,old_dict,standard,olist):

        # Set a given data in a dictionary with position provided as a list
        def setInDict(dataDict, mapList, value): 
            for k in mapList[:-1]: dataDict = dataDict[k]
            dataDict[mapList[-1]] = value
        
        # Get a given data from a dictionary with position provided as a list
        def getFromDict(dataDict, mapList):    
            for k in mapList: dataDict = dataDict[k]
            return dataDict

        value = getFromDict(old_dict,olist)
        try:
            # A value which is not member of the dict is ignored
            setInDict(standard,olist,value)
        except:
            pass



   
    
    
    
                