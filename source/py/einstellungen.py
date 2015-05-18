# -*- coding: utf-8 -*-

import unohelper
from uno import fileUrlToSystemPath
import json
import copy

class Einstellungen():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb

        
        
    def start(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pass
                
            self.erzeuge_einstellungsfenster()
                
        except Exception as e:
            self.mb.nachricht('Export.export '+ str(e),"warningbox")
            log(inspect.stack,tb())


    def erzeuge_einstellungsfenster(self): 
        if self.mb.debug: log(inspect.stack)
        try:
            breite = 680
            hoehe = 400
            breite_listbox = 150
            
            
            sett = self.mb.settings_exp
            
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = posSize_main.X +20
            Y = posSize_main.Y +20            

            # Listener erzeugen 
            listener = {}           
            listener.update( {'auswahl_listener': Auswahl_Item_Listener(self.mb)} )
            
            controls = self.dialog_einstellungen(listener, breite_listbox,breite,hoehe)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)                
            
            # Hauptfenster erzeugen
            posSize = X,Y,breite,pos_y + 40
            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
            fenster_cont.Model.Text = LANG.EXPORT
            
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
        
        lb_items = LANG.TRENNER,LANG.ORGANON_DESIGN,LANG.MAUSRAD,LANG.HTML_EXPORT,LANG.LOG

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
                                        
            elif sel == LANG.TRENNER:
                self.dialog_trenner()
                
            elif sel == LANG.HTML_EXPORT:
                self.dialog_html_export()
                
            elif sel == LANG.ORGANON_DESIGN:
                self.dialog_organon_farben()
                
            elif sel == LANG.EINSTELLUNGEN_EXPORTIEREN:
                self.dialog_einstellungen_exp()

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
                                    (getattr(LANG, el)[0],html_exp_settings[el]),       
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
#             design = self.mb.class_Design
#             design.set_default(tabs)
            
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
            


    def dialog_organon_farben_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        design_items = 'standard','olive','maritim','zander'
        
        controls = [
            10,
            ('control1',"FixedText",         
                                    'tab0',0,168,20,  
                                    ('Label','FontWeight'),
                                    (LANG.ORGANON_DESIGN ,150),                                             
                                    {} 
                                    ),
            30,
            ('controlTit1',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label','FontWeight'),
                                    (LANG.MENULEISTE,150),                                                                    
                                    {} 
                                    ),  
            20,
            ('control2',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control3',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.HINTERGRUND,),                                                                    
                                    {} 
                                    ),  
                    
            20,
            ('control4',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_SCHRIFT,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control5',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.SCHRIFT,),                                                                    
                                    {} 
                                    ),  
            0,
            ###############################################################
            ('controlF1',"FixedLine",         
                                    'tab0',20,168,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
            24,
            ('controlTit2',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label','FontWeight'),
                                    (LANG.BAUMANSICHT,150),                                                                    
                                    {} 
                                    ),  
            20,
            ('control6',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_HF_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control7',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.HINTERGRUND,),                                                                    
                                    {} 
                                    ),  
            20,
            ('control8',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_SCHRIFT_ORDNER,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control9',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.ORDNER,),                                                                    
                                    {} 
                                    ),  
            20,
            ('control10',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_SCHRIFT_DATEI,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control11',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.DATEI,),                                                                    
                                    {} 
                                    ),  
            0,
            ###############################################################
            ('controlF2',"FixedLine",         
                                    'tab0',20,168,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
            24,
            ('controlTit3',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label','FontWeight'),
                                    (LANG.ZEILE,150),                                                                    
                                    {} 
                                    ),  
            20,
            ('control12',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_AUSGEWAEHLTE_ZEILE,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control13',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.AUSGEWAEHLTE_ZEILE,),                                                                    
                                    {} 
                                    ),  
            20,
            ('control14',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_EDITIERTE_ZEILE,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control15',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.EDITIERTE_ZEILE,),                                                                    
                                    {} 
                                    ),  
                    
            20,
            ('control16',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_GEZOGENE_ZEILE,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control17',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.GEZOGENE_ZEILE,),                                                                    
                                    {} 
                                    ),  
            0,
            ###############################################################
            ('controlF3',"FixedLine",         
                                    'tab0',20,168,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
            24,
            ('control18',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_GLIEDERUNG,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control19',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.GLIEDERUNG,),                                                                    
                                    {} 
                                    ), 
            ###############################################################
            ('controlF4',"FixedLine",         
                                    'tab0',20,168,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
            24,
            ('controlTit4',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label','FontWeight'),
                                    (LANG.TRENNER,150),                                                                    
                                    {} 
                                    ),  
            20,
            ('control20',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_TRENNER_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control21',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE_HINTERGRUND,),                                                                    
                                    {} 
                                    ),  
            20,
            ('control22',"FixedText",        
                                    'tab0',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_TRENNER_SCHRIFT,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control23',"FixedText",        
                                    'tab2',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE_SCHRIFT,),                                                                    
                                    {} 
                                    ), 

                  
            -296,
            ###############################################################
            ###############################################################
            ######################  DESIGN ################################
            ###############################################################
            ('controlD1',"FixedText",        
                                    'tab3',0,168,20,  
                                    ('Label','FontWeight'),
                                    (LANG.DESIGNS,150),                                                                    
                                    {} 
                                    ), 
            25, 
            ###############################################################
            ###############################################################
            ######################## SPEICHERN ############################
            ###############################################################
            ('controlE1',"Edit",          
                                    'tab4',0,100,20,    
                                    ('HelpText',),
                                    (LANG.AUSWAHL,),                                                                         
                                    {} 
                                    ), 
            30,
            ('controlB1',"Button",          
                                    'tab4',0,100,23,    
                                    ('Label',),
                                    (LANG.NEUES_DESIGN,),                                                                         
                                    {'setActionCommand':'neues_design','addActionListener':(listener,)} 
                                    ), 
            30,
            ('controlB3',"Button",          
                                    'tab4',0,100,23,    
                                    ('Label',),
                                    (LANG.LOESCHEN,),                                                                         
                                    {'setActionCommand':'loeschen','addActionListener':(listener,)} 
                                    ), 
            30,
            ('controlB4',"Button",          
                                    'tab4',0,100,23,    
                                    ('Label',),
                                    (LANG.EXPORT_2,),                                                                         
                                    {'setActionCommand':'export','addActionListener':(listener,)} 
                                    ), 
            30,
            ('controlB5',"Button",          
                                    'tab4',0,100,23,    
                                    ('Label',),
                                    (LANG.IMPORT_2,),                                                                         
                                    {'setActionCommand':'import','addActionListener':(listener,)} 
                                    ), 
            30,
            
            
            ]
        

        
        
        # Tabs waren urspruenglich gesetzt, um sie in der Klasse Design richtig anzupassen.
        # Das fehlt. 'tab...' wird jetzt nur in Zahlen uebersetzt. Beim naechsten 
        # groesseren Fenster, das ich schreibe und das nachtraegliche Berechnungen benoetigt,
        # sollte der Code generalisiert werden. Vielleicht grundsaetzlich ein Modul fenster.py erstellen?
        tab0 = tab0x = 20
        tab1 = tab1x = 42
        tab2 = tab2x = 78
        tab3 = tab3x = 220
        tab4 = tab4x = 340
        
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
    
    
    def dialog_organon_design_elemente_RB(self,listener):
        if self.mb.debug: log(inspect.stack)
                
        sett = self.mb.settings_orga
        design_items = list(sett['designs'])
        
        
        
        
        controls = []
        aktiv = sett['organon_farben']['aktiv']
            
        for d in design_items:
            state = int(aktiv == d)
            design_control = [
                                 
            ('control%s'%d,"RadioButton",      
                                    'tab3',70,100,20,    
                                    ('Label','State'),
                                    (d,state), 
                                    {'setActionCommand':d,'addActionListener':(listener,)}      
                                    ), 
            20,]  

            controls.extend(design_control)
            
        

        # Tabs waren urspruenglich gesetzt, um sie in der Klasse Design richtig anzupassen.
        # Das fehlt. 'tab...' wird jetzt nur in Zahlen uebersetzt. Beim naechsten 
        # groesseren Fenster, das ich schreibe und das nachtraegliche Berechnungen benoetigt,
        # sollte der Code generalisiert werden. Vielleicht grundsaetzlich ein Modul fenster.py erstellen?
        tab0 = tab0x = 20
        tab1 = tab1x = 42
        tab2 = tab2x = 78
        tab3 = tab3x = 220
        tab4 = tab4x = 200
        
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
    
            
    def dialog_organon_farben(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            listener = Listener_Organon_Farben(self.mb,self)

            controls = self.dialog_organon_farben_elemente(listener)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)

            
#             # UEBERGABE AN LISTENER
            listener.ctrls = {
                            'menu_hintergrund' : ctrls['control2'],
                            'menu_schrift' : ctrls['control4'],
                            'hf_hintergrund' : ctrls['control6'],
                            'schrift_ordner' : ctrls['control8'],
                            'schrift_datei' : ctrls['control10'],
                            'ausgewaehlte_zeile' : ctrls['control12'],
                            'editierte_zeile' : ctrls['control14'],
                            'gezogene_zeile' : ctrls['control16'],
                            'gliederung' : ctrls['control18'],
                            'textfeld' : ctrls['controlE1'],
                            'trenner_farbe_hintergrund' : ctrls['control20'],
                            'trenner_farbe_schrift' : ctrls['control22'],
                            }

            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                
                
            
            
            controls = self.dialog_organon_design_elemente_RB(listener)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)
            
            ctrls_RB = {}
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                ctrls_RB.update({ ctrls[c].Model.Label : ctrls[c] })
            
            listener.RBs = ctrls_RB
    
        except:
            log(inspect.stack,tb())
            
            
    def dialog_einstellungen_exp_elemente(self,listener,nutze_mausrad):
        if self.mb.debug: log(inspect.stack)
        
        controls = (
            10,
            ('controlE_calc',"FixedText",        
                                    20,0,50,20,    
                                    ('Label','FontWeight'),
                                    (LANG.NUTZE_MAUSRAD ,150),                  
                                    {} 
                                    ), 
            20,                                                  
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
    
    
    def dialog_einstellungen_exp(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:

            sett = self.mb.settings_proj
            
            try:
                if sett['nutze_mausrad']:
                    nutze_mausrad = 1
                else:
                    nutze_mausrad = 0
            except:
                self.mb.settings_proj['nutze_mausrad'] = False
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
    
    
from com.sun.star.awt import XMouseListener   
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
                
                
class Listener_Organon_Farben(unohelper.Base,XMouseListener,XActionListener):
    
    def __init__(self,mb,auswahl_item_listener):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = {}
        self.RBs = {}
        self.auswahl_item_listener = auswahl_item_listener
              
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack) 

        try:

            for c in self.ctrls:
                if self.ctrls[c] == ev.Source:
                    #print(c)
                    farbe = self.waehle_farbe(ev,c)
                    
                    if c =='hf_hintergrund':
                        self.setze_farbe_hintergrund(farbe)
                        
                    elif c =='menu_hintergrund':
                        self.setze_farbe_menuleiste_hintergrund(farbe)
                    elif c =='menu_schrift':
                        self.setze_farbe_menuleiste_schrift(farbe)
                        
                    elif c =='schrift_datei':
                        self.setze_farbe_schrift_dateien(farbe)
                    elif c =='schrift_ordner':
                        self.setze_farbe_schrift_ordner(farbe)
                        
                        
                    elif c == 'ausgewaehlte_zeile':
                        KONST.FARBE_AUSGEWAEHLTE_ZEILE = farbe
                    elif c == 'editierte_zeile':
                        KONST.FARBE_EDITIERTE_ZEILE = farbe
                    elif c == 'gezogene_zeile':
                        KONST.FARBE_GEZOGENE_ZEILE = farbe
                    
                    elif c == 'gliederung':
                        self.setze_farbe_gliederung(farbe)
                    
                    elif c == 'trenner_farbe_hintergrund':
                        self.setze_farbe_trenner(farbe,c)
                    elif c == 'trenner_farbe_schrift':
                        self.setze_farbe_trenner(farbe,c)
                    
                    ev.Source.Model.BackgroundColor = farbe  
                    self.mb.class_Funktionen.schreibe_settings_orga()
                    break
        except:
            log(inspect.stack,tb())   
    
    def waehle_farbe(self,ev,art):  
        if self.mb.debug: log(inspect.stack)

        farbe = self.mb.class_Funktionen.waehle_farbe(self.mb.settings_orga['organon_farben'][art])
        
        self.mb.settings_orga['organon_farben'][art] = farbe
        self.mb.class_Funktionen.schreibe_settings_orga()
        return farbe  
      
            
    def mouseExited(self, ev):
        return False
    def mouseEntered(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
    
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            cmd = ev.ActionCommand
            
            if cmd == 'neues_design':
                self.erzeuge_neues_design()
                
            elif cmd == 'loeschen':
                self.loesche_design()
            
            elif cmd == 'export':
                self.export_design()
                
            elif cmd == 'import':
                self.import_design()
                
            else:
                self.setze_design(cmd)
        except:
            log(inspect.stack,tb()) 
    
    def import_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs'] 
        
        pfad = self.mb.class_Funktionen.filepicker()
        if pfad == None: return
        
        odict = self.mb.class_Funktionen.oeffne_json(pfad)
        if odict == None:
            self.mb.nachricht(LANG.KEINE_JSON_DATEI,"warningbox")
            return
        
        
        for k in odict:
            neu = copy.deepcopy(k)
            while k in sett:
                k = k+'x'
            sett.update({ k : odict[neu] })
        
        self.mb.class_Funktionen.schreibe_settings_orga()
        self.ansicht_erneuern()
    
    def export_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs'] 
        
        name = self.ctrls['textfeld'].Model.Text
        if name == '':
            self.mb.nachricht(LANG.EXPORTNAMEN_EINGEBEN,"infobox")
            return
        
        pfad = self.mb.class_Funktionen.folderpicker()
        
        if pfad == None:
            return
        
        pfad = os.path.join(pfad,name+'.json')
        
        if os.path.exists(pfad):
            # 16777216 Flag fuer YES_NO
            entscheidung = self.mb.nachricht(LANG.DATEI_EXISTIERT,"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return
            
        with open(pfad, 'w') as outfile:
            json.dump(sett, outfile,indent=4, separators=(',', ': '))
        

    def loesche_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs']  
        active,act_name = None, None
        
        for c in self.RBs:
            if self.RBs[c].State:
                active,act_name = self.RBs[c], self.RBs[c].Model.Label
        print(act_name)  

        self.RBs.pop(act_name, None)
        self.mb.settings_orga['designs'].pop(act_name,None)
        
        self.ansicht_erneuern()
             
    
    def erzeuge_neues_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs']
        
        name = self.ctrls['textfeld'].Model.Text
        if name == '':
            self.mb.nachricht(LANG.DESIGNNAMEN_EINGEBEN,"infobox")
            return
        if name in sett:
            self.mb.nachricht(LANG.DESIGN_EXISTIERT,"infobox")
            return
        
        new_set = copy.deepcopy(self.mb.settings_orga['organon_farben'])
        new_set.pop('aktiv',None)
        
        sett.update({ name : new_set })
        self.mb.class_Funktionen.schreibe_settings_orga()
        
        self.ansicht_erneuern()
        
            
    def setze_design(self,cmd):
        if self.mb.debug: log(inspect.stack)

        try:

            sett = self.mb.settings_orga['designs'][cmd]
            
            self.setze_farbe_hintergrund(sett['hf_hintergrund'])
            self.ctrls['hf_hintergrund'].Model.BackgroundColor = sett['hf_hintergrund']
            
            self.setze_farbe_menuleiste_hintergrund(sett['menu_hintergrund'])
            self.ctrls['menu_hintergrund'].Model.BackgroundColor = sett['menu_hintergrund']
            
            self.setze_farbe_menuleiste_schrift(sett['menu_schrift'])
            self.ctrls['menu_schrift'].Model.BackgroundColor = sett['menu_schrift']
            
            self.setze_farbe_schrift_dateien(sett['schrift_datei'])
            self.ctrls['schrift_datei'].Model.BackgroundColor = sett['schrift_datei']
            
            self.setze_farbe_schrift_ordner(sett['schrift_ordner'])
            self.ctrls['schrift_ordner'].Model.BackgroundColor = sett['schrift_ordner']
            
            self.setze_farbe_gliederung(sett['gliederung'])
            self.ctrls['gliederung'].Model.BackgroundColor = sett['gliederung']
            
            KONST.FARBE_AUSGEWAEHLTE_ZEILE = sett['ausgewaehlte_zeile']
            self.ctrls['ausgewaehlte_zeile'].Model.BackgroundColor = sett['ausgewaehlte_zeile']
            
            KONST.FARBE_EDITIERTE_ZEILE = sett['editierte_zeile']
            self.ctrls['editierte_zeile'].Model.BackgroundColor = sett['editierte_zeile']
            
            KONST.FARBE_GEZOGENE_ZEILE = sett['gezogene_zeile']
            self.ctrls['gezogene_zeile'].Model.BackgroundColor = sett['gezogene_zeile']
            
            KONST.FARBE_TRENNER_HINTERGRUND   = sett['trenner_farbe_hintergrund']
            KONST.FARBE_TRENNER_SCHRIFT       = sett['trenner_farbe_schrift']
            self.ctrls['trenner_farbe_hintergrund'].Model.BackgroundColor = sett['trenner_farbe_hintergrund']
            self.ctrls['trenner_farbe_schrift'].Model.BackgroundColor = sett['trenner_farbe_schrift']
            self.setze_farbe_trenner(sett['trenner_farbe_hintergrund'],'trenner_farbe_hintergrund')
            
            self.mb.settings_orga.update({'organon_farben': copy.deepcopy(sett) })
            self.mb.settings_orga['organon_farben'].update({'aktiv': cmd})
            
            self.mb.class_Funktionen.schreibe_settings_orga()
            
        except:
            log(inspect.stack,tb())  
         
    
    def ansicht_erneuern(self):
        if self.mb.debug: log(inspect.stack)
        
        for c in self.auswahl_item_listener.container.getControls():
                c.dispose()
        self.auswahl_item_listener.dialog_organon_farben()
        
    
    def setze_farbe_trenner(self,farbe,cmd):
        if self.mb.debug: log(inspect.stack)
        
        if cmd == 'trenner_farbe_hintergrund':
            KONST.FARBE_TRENNER_HINTERGRUND = farbe
        else:
            KONST.FARBE_TRENNER_SCHRIFT = farbe
        
        secs = self.mb.doc.TextSections
        names = secs.ElementNames
        
        trenner = []
        
        for n in names:
            if 'trenner' in n:
                trenner.append(secs.getByName(n))
        for t in trenner[:-1]:
            t.BackColor = KONST.FARBE_TRENNER_HINTERGRUND
            t.Anchor.CharColor = KONST.FARBE_TRENNER_SCHRIFT 

               
    def setze_farbe_menuleiste_hintergrund(self,farbe):
        if self.mb.debug: log(inspect.stack)

        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            menuleiste = hf.Context.Context.getControl('Organon_Menu_Bar')
            menuleiste.Model.BackgroundColor = farbe
        
            try:
                for c in menuleiste.Controls:
                    c.Model.BackgroundColor = farbe
            except Exception as e:
                print(e)
                
        KONST.FARBE_MENU_HINTERGRUND = farbe
        
            
    def setze_farbe_menuleiste_schrift(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            menuleiste = hf.Context.Context.getControl('Organon_Menu_Bar')
        
            try:
                for c in menuleiste.Controls:
                    try:
                        c.Model.TextColor = farbe
                    except:
                        pass
            except Exception as e:
                print(e)
                
        KONST.FARBE_MENU_SCHRIFT = farbe
        
       
    def setze_farbe_hintergrund(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        # Dialog Fenster
        self.mb.dialog.Model.BackgroundColor = farbe
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            hf.Model.BackgroundColor = farbe
            hf.Context.Context.Model.BackgroundColor = farbe
            
            # Zeilen im Hauptfeld
            zeilen = hf.Controls
                
            for z in zeilen:
                z.Model.BackgroundColor = farbe
                #Controls in Zeilen
                for el in z.Controls:
                    el.Model.BackgroundColor = farbe
                    
        KONST.FARBE_HF_HINTERGRUND = farbe


    def setze_farbe_schrift_ordner(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            papierkorb = self.mb.props[tab].Papierkorb
            
            xml = self.mb.props[tab].xml_tree
            root = xml.getroot()
            ordner = root.findall(".//*[@Art='dir']")
            
            #Projektordner anhaengen
            ordner.append(root.find('.//nr0'))
            #Papierkorb anhaengen
            ordner.append(root.find('.//{}'.format(papierkorb)))

            for o in ordner:
                zeile = hf.getControl(o.tag)
                textfeld = zeile.getControl('textfeld')
                textfeld.Model.TextColor = farbe

        KONST.FARBE_SCHRIFT_ORDNER = farbe
        
                
    def setze_farbe_schrift_dateien(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            
            xml = self.mb.props[tab].xml_tree
            root = xml.getroot()
            ordner = root.findall(".//*[@Art='pg']")
                    
            for o in ordner:
                zeile = hf.getControl(o.tag)
                textfeld = zeile.getControl('textfeld')
                textfeld.Model.TextColor = farbe
                    
        KONST.FARBE_SCHRIFT_DATEI = farbe

                    
    def setze_farbe_gliederung(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            
            # Zeilen im Hauptfeld
            zeilen = hf.Controls
            
            for z in zeilen:   
                gliederung = z.getControl('tag3')
                if gliederung != None:
                    gliederung.Model.TextColor = farbe
        
        KONST.FARBE_GLIEDERUNG = farbe

                
                
        


                

                
                
                