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
                        LANG.UEBERSETZUNGEN,
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
                        LANG.UEBERSETZUNGEN,
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
                
            elif sel == LANG.UEBERSETZUNGEN:
                u = Uebersetzungen(self.mb).run()

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
        self.dialog_uebersetzung(ctrls_ueber,lang_akt,organon_lang_files,ctrls_konst)
         
        self.gefaerbte = list(ctrls_ueber.values())

    def langpy_auslesen(self):
        if self.mb.debug: log(inspect.stack)
        
        def formatiere_txt(txt):
            
            try:
                
                t = txt.strip()
                
                if "'''" in t:
                    count = t.count("'''")
                    if  t.count("'''") == 1:
                        if "u'''" in t:
                            t = t + "'''"  
                        else:
                            t = "u'''" + t
                elif 'u"' in t:
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
        with codecs_open(pfad_imp , "r",'utf-8') as file:
            lines = file.readlines()
        
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
    
    
    
    
    
    def dialog_uebersetzung_elemente(self,win_dispose_listener,button_listener,organon_lang_files):
        if self.mb.debug: log(inspect.stack)            
        
        os_path = datei_pfad = os.path.join(self.mb.path_to_extension,'languages')
        path_orga_lang = LANG.ORGANON_LANG_PATH + '\r\n\r\n' + os_path
        
        fensterhoehe = 600
        fensterbreite = 800
        
        items = 'lang_en','lang_de'
        sel = 0
        
        tab1 = 140
        tab2 = 180
        
        breite = 160
        
        
        controls = [
            10,
            ('control_dispose',"Button",        
                                    fensterbreite - tab1,0,120,40,    
                                    ('Label',),
                                    (LANG.FENSTER_SCHLIESSEN,),                  
                                    {'addActionListener':(win_dispose_listener,)} 
                                    ), 
            70,
            ('control_ref',"FixedText",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.REFERENZ,),                  
                                    {} 
                                    ), 
            20,
            ('control_ref_lb',"ListBox",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':items,'SelectedItems':(sel,),
                                     'addItemListener':button_listener}
                                    ), 
            50,
            ('control_odl',"FixedText",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.ORG_SPRACHDATEI_LADEN,),                  
                                    {} 
                                    ), 
            20,
            ('control_orga',"ListBox",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Border','Dropdown','LineCount'),
                                    (2,True,15),       
                                    {'addItems':tuple(organon_lang_files),
                                     'SelectedItems':(0,),
                                     'addItemListener':button_listener}
                                    ), 
            50,
            ('control_bdl',"Button",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.BENUTZER_DATEI_LADEN,),                  
                                    {'setActionCommand':'nutzer_laden',
                                     'addActionListener':(button_listener,)}
                                    ), 
            60,
            ('control_txt',"Edit",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('HelpText',),
                                    (LANG.EXPORTNAMEN_EINGEBEN,),                  
                                    {}
                                    ), 
            
            30,
            ('control_speichern',"Button",        
                                    fensterbreite - tab2,0,breite,25,    
                                    ('Label',),
                                    (LANG.UEBERSETZUNG_SPEICHERN,),                  
                                    {'setActionCommand':'speichern',
                                     'addActionListener':(button_listener,)}
                                    ),  
            100,
            ('control_path',"Edit",        
                                    fensterbreite - tab2,0,breite,25,     
                                    ('Text','TextColor','MultiLine',
                                    'Border','BackgroundColor'), 
                                    (path_orga_lang,KONST.FARBE_SCHRIFT_DATEI, True,
                                     0,KONST.FARBE_HF_HINTERGRUND),             
                                    {} 
                                    ),  
            ]

            
        return controls

 
    def dialog_uebersetzung(self,ctrls_ueber,lang_akt,organon_lang_files,ctrls_konst): 
        if self.mb.debug: log(inspect.stack)
        
        # Listener erzeugen 
        win_dispose_listener = Window_Disposer(self.fenster,self.mb) 
        button_listener = Uebersetzung_Button_Listener(self.mb,ctrls_ueber,lang_akt,self)        
        
        controls = self.dialog_uebersetzung_elemente(win_dispose_listener,button_listener,organon_lang_files)
        ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)                
          
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
        scrollb = self.mb.erzeuge_Scrollbar(self.cont,PosSize,self.container)
        self.mb.class_Mausrad.registriere_Maus_Focus_Listener(self.cont)

                    
    def container_erstellen(self):
        if self.mb.debug: log(inspect.stack)
        
        top_h = self.mb.topWindow.PosSize.Height
        prozent = top_h - (top_h /10*9)

        fensterhoehe = top_h - prozent
        fensterbreite = 800
        posSize = (self.mb.win.Size.Width,prozent / 2,fensterbreite,fensterhoehe)
        

        if self.mb.platform == 'linux':
            self.mb.nachricht(LANG.USE_CLOSE_BUTTON,'warningbox')

        win, cont = self.mb.erzeuge_Dialog_Container(posSize,Flags=1+16+32+64+512)
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
            
            with codecs_open(pfad , "r",'utf-8') as file:
                f = file.readlines()
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
    
    
from com.sun.star.awt import XActionListener,XItemListener
class Window_Disposer(unohelper.Base, XActionListener):
    def __init__(self,win,mb):
        if mb.debug: log(inspect.stack)
        self.win = win 
        self.mb = mb     
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        self.win.dispose()
    def disposing(self,ev):
        pass  
    
class Uebersetzung_Button_Listener(unohelper.Base, XActionListener,XItemListener):
    def __init__(self,mb,ctrls_ueber,lang_akt,class_Uebersetzung):
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
            datei_pfad = self.mb.class_Funktionen.filepicker(filepath=pfad,filter='py')
        else:
            datei_pfad = self.mb.class_Funktionen.filepicker(filter='py')
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
        
        with codecs_open(pfad , "w",'utf-8') as file:
            file.write(text)  
              
    
    
        
    def disposing(self,ev):
        pass
   


   
    
    
    
                