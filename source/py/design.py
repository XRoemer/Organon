# -*- coding: utf-8 -*-

import unohelper
import json
import copy

class Organon_Design():
    
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
        
        # wird bei Aufruf von dialog_organon_farben oder dialog_persona
        # durch den Auswahl_Item_Listener gesetzt 
        self.container = None
        self.fenster = None
        
        self.personas_kopiert = False


    def get_personas(self):
        if self.mb.debug: log(inspect.stack)
        
        oPathSettings = self.mb.createUnoService("com.sun.star.util.PathSettings")
            
        gallery_paths = oPathSettings.Gallery
        gal_user_path = uno.fileUrlToSystemPath( [p for p in gallery_paths.split(';') if 'user' in p][0] )
        
        files = os.listdir(gal_user_path)
        
        personas = []
        personas_dict = {}
        
        if 'personas' in files:
            
            files_pers = os.listdir(os.path.join(gal_user_path,'personas'))
            
            for f in files_pers:
                pfad = os.path.join(gal_user_path,'personas',f)
                entries = os.listdir(pfad)
                
                for e in entries:
                    pfad_e = os.path.join(pfad,e)
                     
         
                    control, model = self.mb.createControl(self.mb.ctx, "ImageControl", 0, 0, 10, 10, ('ImageURL',), 
                                                            (uno.systemPathToFileUrl(pfad_e),))
                    pref_X = control.getPreferredSize().Width
                    pref_Y = control.getPreferredSize().Height
                    model.dispose()
                    control.dispose()
                    
                    if pref_X >= 2500 and pref_Y >= 200:
                        personas.append(f)
                        personas_dict.update({f:pfad_e})
                        break
                    
        return personas_dict,personas,gal_user_path
    
    
    def setze_persona(self,persona_url,nutze_persona='own'):
        if self.mb.debug: log(inspect.stack)
        
        ctx = self.mb.ctx
        smgr = self.mb.ctx.ServiceManager
           
        config_provider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider",ctx)
  
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = "nodepath"
        prop.Value = "/org.openoffice.Office.Common/Misc"
               
        config_access = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
             
        config_access.Persona = nutze_persona
        config_access.PersonaSettings = persona_url
                
        config_access.commitChanges()
        
        
    def setze_dok_farben(self,dok_farbe,hintergrund):
        if self.mb.debug: log(inspect.stack)
        
        try:
            smgr = uno.getComponentContext().ServiceManager
            config_provider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider",self.mb.ctx)
            
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = "nodepath"
            prop.Value = "/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice"\
            ".Office.UI:ColorScheme['{0}']/DocColor".format(self.mb.programm)
            config_access_doc = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
            prop.Value = "/org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice"\
            ".Office.UI:ColorScheme['{0}']/AppBackground".format(self.mb.programm)
            config_access_bg = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
            
            
            if dok_farbe != None:
                config_access_doc.Color = dok_farbe
                config_access_bg.Color = hintergrund
            else:
                config_access_doc.setPropertyValue('Color',None)
                config_access_bg.setPropertyValue('Color',None)
    
            config_access_doc.commitChanges()
            config_access_bg.commitChanges()
        except:
            log(inspect.stack,tb())
                     

    def dialog_organon_farben_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)

        
        ctrls = [
                 ('Titel',LANG.MENULEISTE),
                 ('Farbe','menu_hintergrund',KONST.FARBE_MENU_HINTERGRUND),
                 ('Beschreibung',LANG.HINTERGRUND),
                 ('Farbe','menu_schrift',KONST.FARBE_MENU_SCHRIFT),
                 ('Beschreibung',LANG.SCHRIFT),
                 ('Sep'),
                 
                 ('Titel',LANG.BAUMANSICHT),
                 ('Farbe','hf_hintergrund',KONST.FARBE_HF_HINTERGRUND),
                 ('Beschreibung',LANG.HINTERGRUND),
                 ('Farbe','schrift_ordner',KONST.FARBE_SCHRIFT_ORDNER),
                 ('Beschreibung',LANG.ORDNER),
                 ('Farbe','schrift_datei',KONST.FARBE_SCHRIFT_DATEI),
                 ('Beschreibung',LANG.DATEI),
                 ('Sep'),
                 
                 ('Titel',LANG.ZEILE),
                 ('Farbe','ausgewaehlte_zeile',KONST.FARBE_AUSGEWAEHLTE_ZEILE),
                 ('Beschreibung',LANG.AUSGEWAEHLTE_ZEILE),
                 ('Farbe','editierte_zeile',KONST.FARBE_EDITIERTE_ZEILE),
                 ('Beschreibung',LANG.EDITIERTE_ZEILE),
                 ('Farbe','gezogene_zeile',KONST.FARBE_GEZOGENE_ZEILE),
                 ('Beschreibung',LANG.GEZOGENE_ZEILE),
                 ('Sep'),
                 
                 ('Farbe','gliederung',KONST.FARBE_GLIEDERUNG),
                 ('Beschreibung',LANG.GLIEDERUNG),
                 ('Sep'),
                 
                 ('Titel',LANG.TRENNER),
                 ('Farbe','trenner_farbe_hintergrund',KONST.FARBE_TRENNER_HINTERGRUND),
                 ('Beschreibung',LANG.HINTERGRUND),
                 ('Farbe','trenner_farbe_schrift',KONST.FARBE_TRENNER_SCHRIFT),
                 ('Beschreibung',LANG.SCHRIFT),
                 ('Sep'),
                 
                 ('Titel',LANG.TABS),
                 ('Farbe','tabs_hintergrund',KONST.FARBE_TABS_HINTERGRUND),
                 ('Beschreibung',LANG.HINTERGRUND),
                 ('Farbe','tabs_schrift',KONST.FARBE_TABS_SCHRIFT),
                 ('Beschreibung',LANG.SCHRIFT),
                 ('Farbe','tabs_sel_hintergrund',KONST.FARBE_TABS_SEL_HINTERGRUND),
                 ('Beschreibung',u'{0} {1}'.format(LANG.AUSGEWAEHLTE_ZEILE,LANG.HINTERGRUND_ABK)),
                 ('Farbe','tabs_sel_schrift',KONST.FARBE_TABS_SEL_SCHRIFT),
                 ('Beschreibung',u'{0} {1}'.format(LANG.AUSGEWAEHLTE_ZEILE,LANG.SCHRIFT)),
                 ('Farbe','tabs_trenner',KONST.FARBE_TABS_TRENNER),
                 ('Beschreibung',LANG.TRENNER),
                 ('Sep'),
                 
                 ('Farbe','linien',KONST.FARBE_LINIEN),
                 ('Beschreibung',LANG.LINIEN),
                 ('Sep'),
                 
                 ('Titel',LANG.DEAKTIVIERTE_BUTTONS),
                 ('Farbe','deaktiviert',KONST.FARBE_DEAKTIVIERT),
                 ('Beschreibung',LANG.SCHRIFT),
                 
                 ]
        

        controls = [
            10,
            ('control_Writer_Design',"CheckBox",1,        
                                    'tab2',3,100,20,  
                                    ('Label','FontWeight','State'),
                                    (LANG.WRITER_DESIGN,150,self.mb.settings_orga['organon_farben']['design_office']),                                                                    
                                    {'setActionCommand':'writer_design','addActionListener':(listener)} 
                                    ),
            0,
            ('control_wDesign_bearbeiten',"Button",1,          
                                    'tab3-max',-2,100,23,    
                                    ('Label',),
                                    (LANG.BEARBEITEN,),                                                                         
                                    {'setActionCommand':'writer_design_bearbeiten','addActionListener':(listener)} 
                                    ), 
            30,
            ('control_Fenster_Design',"CheckBox",1,        
                                    'tab2',3,100,20,  
                                    ('Label','FontWeight','State'),
                                    (LANG.DESIGN_ORGANON_FENSTER,150,self.mb.settings_orga['organon_farben']['design_organon_fenster']),                                                                    
                                    {'setActionCommand':'fenster_design','addActionListener':(listener)} 
                                    ),
            ]
        
        
        controls.extend(['Y=-10'])
        
        zaehler = 0
        
        for c in ctrls:

            if c[0] == 'Titel':
            
                controls.extend([
                                 20,
                                [u'Tit_{}'.format(c[1]),"FixedText",1,        
                                        'tab1',0,100,20,  
                                        ('Label','FontWeight'),
                                        (c[1],150),                                                                    
                                        {} 
                                        ],  
                                ]
                                )
            elif c[0] == 'Farbe':
            
                controls.extend([
                                 20,
                                [c[1],"FixedText",0,        
                                        'tab0',0,32,16,  
                                        ('BackgroundColor','Label','Border'),
                                        (c[2],'    ',1),       
                                        {'addMouseListener':(listener)} 
                                        ],
                                ]  
                                )
            elif c[0] == 'Beschreibung':
                
                controls.extend([
                                [u'Beschr_{}'.format(zaehler),"FixedText", 1,       
                                        'tab1',0,100,20,  
                                        ('Label',),
                                        (c[1],),                                                                    
                                        {} 
                                        ], 
                                ]
                                )
                
            elif c == 'Sep':
                
                controls.extend([
                                [u'fixed_L{}'.format(zaehler),"FixedLine",0,         
                                        'tab0x-tab1-E',20,168,1,   
                                        (),
                                        (),                                                  
                                        {} 
                                        ],
                               5 ]
                                )
            zaehler += 1
                

        controls.extend([
            

                  
            'Y=80',
            ######################  DESIGN ################################
            ('controlD1',"FixedText",1,        
                                    'tab2',0,168,20,  
                                    ('Label','FontWeight'),
                                    (LANG.DESIGNS,150),                                                                    
                                    {} 
                                    ), 
            25, 
            ######################## SPEICHERN ############################
            ('textfeld',"Edit",1,          
                                    'tab3-max',0,100,20,    
                                    ('HelpText',),
                                    (LANG.AUSWAHL,),                                                                         
                                    {} 
                                    ), 
            30,
            ('controlB1',"Button",1,          
                                    'tab3',0,100,23,    
                                    ('Label',),
                                    (LANG.NEUES_DESIGN,),                                                                         
                                    {'setActionCommand':'neues_design','addActionListener':(listener)} 
                                    ), 
            30,
            ('controlB3',"Button",1,          
                                    'tab3-max',0,100,23,    
                                    ('Label',),
                                    (LANG.LOESCHEN,),                                                                         
                                    {'setActionCommand':'loeschen','addActionListener':(listener)} 
                                    ), 
            30,
            ('controlB4',"Button",1,          
                                    'tab3-max',0,100,23,    
                                    ('Label',),
                                    (LANG.EXPORT_2,),                                                                         
                                    {'setActionCommand':'export','addActionListener':(listener)} 
                                    ), 
            30,
            ('controlB5',"Button",1,          
                                    'tab3-max',0,100,23,    
                                    ('Label',),
                                    (LANG.IMPORT_2,),                                                                         
                                    {'setActionCommand':'import','addActionListener':(listener)} 
                                    ), 
            'Y=40',
            ]

            )
            
            
            
        sett = self.mb.settings_orga
        design_items = list(sett['designs'])

        
        aktiv = sett['organon_farben']['aktiv']
            
        for d in design_items:
            state = int(aktiv == d)
            design_control = [
                                 
            ('control%s_radio'%d,"RadioButton",1,      
                                    'tab2',70,100,20,    
                                    ('Label','State'),
                                    (d,state), 
                                    {'setActionCommand':d,'addActionListener':(listener)}      
                                    ), 
            20,]  

            controls.extend(design_control)
            
            
            
        tabs = {
                 0 : (None, 10),
                 1 : (None, 20),
                 2 : (None, 25),
                 3 : (None, 0),
                 }
        
        abstand_links = 10
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
        
        listener.tabs = tabs3        
        
        return controls2,max_breite
    
 
            
    def dialog_organon_farben(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Listener_Organon_Farben(self.mb,self)

            controls,max_breite = self.dialog_organon_farben_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)
            self.mb.class_Einstellungen.container_anpassen(self.container,max_breite,max_hoehe+10,self.fenster)
            
            # UEBERGABE AN LISTENER
            listener.ctrls = ctrls
            
            radios = {}
            # Controls in Hauptfenster eintragen
            for c,v in ctrls.items():
                if 'radio' in c: 
                    radios[c] = v
                else:
                    self.container.addControl(c,v)
                 
            
            ctrls_RB = {}
            
            # RadioButtons in Hauptfenster eintragen
            for c in radios:
                self.container.addControl(c,ctrls[c])
                ctrls_RB.update({ ctrls[c].Model.Label : ctrls[c] })
            
            listener.RBs = ctrls_RB
    
        except:
            log(inspect.stack,tb())
            
            
    def dialog_persona_elemente(self,listener,persona_items):
        if self.mb.debug: log(inspect.stack)
        
        controls = [
            1,
            ('control_theme',"ImageControl",0,         
                                    'tab0x-max',0,self.container.Size.Width-6,80,  
                                    ('Label','FontWeight','BackgroundColor','Border'),
                                    (LANG.ORGANON_DESIGN ,150,KONST.FARBE_MENU_HINTERGRUND,0),                                             
                                    {} 
                                    ),
            20,
            ('control_theme_tit',"Edit",0,        
                                    'tab4x+40',0,80,25,  
                                    ('Text','FontWeight','TextColor','PaintTransparent','Border'),
                                    (LANG.PERSONA,150,KONST.FARBE_MENU_SCHRIFT,True,0),                                                                    
                                    {} 
                                    ),  
            30, 
            ('control_theme_untertit',"Edit",1,        
                                    'tab2x',0,200,25,  
                                    ('Text','TextColor','PaintTransparent','Border'),
                                    (LANG.THEMA_UNTERTITEL,KONST.FARBE_MENU_SCHRIFT,True,0),                                                                    
                                    {} 
                                    ),  
            ############################ PERSONAS ######################################
            'Y=120',
            ('control_VP',"FixedText",1,       
                                    'tab1',0,150,20,  
                                    ('Label','FontWeight'),
                                    (LANG.VORHANDENE_PERSONAS,150),                                                                    
                                    {} 
                                    ),  
            30,
            ('control_personas_list',"ListBox",0,       
                                    'tab1',0,120,20,  
                                    ('Border','Dropdown','LineCount'),
                                    ( 2,True,10),       
                                    {'addItems':(persona_items,0),'addItemListener':listener}
                                    ), 
            ('control_Abstand',"FixedText",1,       
                                    'tab0',0,150,20,  
                                    (),
                                    (),                                                                    
                                    {} 
                                    ), 
             
                    
            ############################# EIGENES PERSONAS #####################################
            'Y=120',
            ('control_farbe_schrift',"FixedText",0,       
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_SCHRIFT,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control5',"FixedText",1,      
                                    'tab4+40',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE_SCHRIFT,),                                                                    
                                    {} 
                                    ),  
            25,
            ('control_solid',"RadioButton",1,        
                                    'tab3x',0,150,25,  
                                    ('Label','State'),
                                    (LANG.HINTERGRUND_EINFARBIG,1),                                                                    
                                    {'setActionCommand':'control_solid','addActionListener':(listener)}
                                    ), 
            25,
            ('control_farbe_hintergrund',"FixedText",0,        
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control3',"FixedText",1,        
                                    'tab4+40',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE,),                                                                    
                                    {} 
                                    ),  
                    
            25,
            ('control_gradient',"RadioButton",1,        
                                    'tab3x',0,150,25,  
                                    ('Label',),
                                    (LANG.HINTERGRUND_GRADIENT,),                                                                    
                                    {'setActionCommand':'control_gradient','addActionListener':(listener)}
                                    ), 
            25,
            ('control_farbe_hintergrund2',"FixedText",0,        
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),
            0,
            ('control_f2text',"FixedText",1,        
                                    'tab4+40',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE2,),                                                                    
                                    {} 
                                    ),
            25,
            ('control_url_nutzer',"RadioButton",1,       
                                    'tab3x',0,150,25,  
                                    ('Label',),
                                    (LANG.HINTERGRUND_NUTZER,),                                                                    
                                    {'setActionCommand':'control_url_nutzer','addActionListener':(listener)}
                                    ), 
              
                    
            20,
            ('control_url_aussuchen',"Button",1,        
                                    'tab4',0,80,20,  
                                    ('Label',),
                                    (LANG.AUSWAHL,),                                                                    
                                    {'setActionCommand':'control_url_aussuchen','addActionListener':(listener)}
                                    ), 
            30,
            ('control_url',"FixedText",0,         
                                    'tab4x',0,300,20,    
                                    ('Label',),
                                    (LANG.URL,),                                                                         
                                    {} 
                                    ), 
            20,
            ('control_url_info',"FixedText",1,          
                                    'tab4x',0,300,20,    
                                    ('Label',),
                                    (LANG.URL_INFO,),                                                                         
                                    {} 
                                    ), 
    
            'Y=120',
            ###############################################################
            ######################## SPEICHERN ############################
            ###############################################################
 
            
            ('control_personas_name',"Edit",0,          
                                    'tab6-tab6-E',0,100,20,    
                                    ('HelpText',),
                                    (LANG.PERSONANAMEN_EINGEBEN,),                                                                         
                                    {} 
                                    ), 
            30,
            ('control_neues_thema',"Button",1,          
                                    'tab6-tab6-E',0,100,23,    
                                    ('Label',),
                                    (LANG.NEUES_PERSONA_THEMA,),                                                                         
                                    {'setActionCommand':'neues_persona','addActionListener':(listener)} 
                                    ),
            60,
            ('control_thema_loeschen',"Button",1,          
                                    'tab6-tab6-E',0,50,23,    
                                    ('Label',),
                                    (LANG.PERSONA_LOESCHEN,),                                                                         
                                    {'setActionCommand':'loesche_persona','addActionListener':(listener)} 
                                    ), 
            120,
            ('control_kein_persona',"Button",1,         
                                    'tab6-tab6-E',0,50,23,    
                                    ('Label',),
                                    (LANG.KEIN_PERSONA,),                                                                         
                                    {'setActionCommand':'kein_persona','addActionListener':(listener)} 
                                    ), 
            60,
            ('control_anwenden',"Button",1,         
                                    'tab6-tab6-E',0,50,23,    
                                    ('Label',),
                                    (LANG.PERSONA_ANWENDEN,),                                                                         
                                    {'setActionCommand':'anwenden','addActionListener':(listener)} 
                                    ),
            ]


    
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 5),
                 1 : (None, 30),
                 2 : (None, 0),
                 3 : (None, 0),
                 4 : (None, 80),
                 5 : (None, 0),
                 6 : (None, 10),
                 7 : (None, 0),
                 }
        
        abstand_links = 2
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
        
        listener.tabs = tabs3        
        
        return controls2,max_breite
    
              
    def dialog_persona(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.kopiere_personas()
            
            # vorhandene personas auslesen
            personas_dict,personas_list,gallery_user_path = self.mb.class_Organon_Design.get_personas()        
            listener = Listener_Persona(self.mb,personas_dict,personas_list,gallery_user_path)
            
            # controls erzeugen
            controls,max_breite = self.dialog_persona_elemente(listener,tuple(personas_list))
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)
            
            container = self.mb.class_Einstellungen.container
            fenster = self.mb.class_Einstellungen.haupt_fenster
            self.mb.class_Einstellungen.container_anpassen(container,max_breite,max_hoehe,fenster)
            
            # Uebergabe an listener
            listener.ctrls = ctrls
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                self.container.addControl(c,ctrls[c])
                
            listener.ctrl_container = self.container
            listener.theme_ansicht_erneuern()

        except:
            log(inspect.stack,tb())
            
            
    def kopiere_personas(self):
        if self.personas_kopiert:
            return
        self.personas_kopiert = True
        
        if self.mb.debug: log(inspect.stack)
            
        try:
            personas_dict,personas_list,gallery_user_path = self.mb.class_Organon_Design.get_personas()
            personas_list = [f.lower() for f in personas_list]
            sett = self.mb.settings_orga['designs']
            
            personas_url = os.path.join(self.mb.path_to_extension,'personas')
            files = os.listdir(personas_url)
            
            import shutil
            
            for design in files:
                if design in sett:
                    if design.lower() not in personas_list:
                        try:
                            src = os.path.join(personas_url,design)
                            dst = os.path.join(gallery_user_path,'personas',design)
                            shutil.copytree(src, dst)
                        except:
                            log(inspect.stack,tb())
                            
        except:
            log(inspect.stack,tb())
           
           
    def set_app_style(self,win,settings_orga):
        try:
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            frame = desktop.Frames.getByIndex(0)
            
            rot = 16275544
    
            hf = KONST.FARBE_HF_HINTERGRUND
            #menu = KONST.FARBE_MENU_HINTERGRUND
            schrift = KONST.FARBE_SCHRIFT_DATEI
            #menu_schrift = KONST.FARBE_MENU_SCHRIFT
            #selected = KONST.FARBE_AUSGEWAEHLTE_ZEILE
            #ordner = KONST.FARBE_SCHRIFT_ORDNER
            linien = KONST.FARBE_LINIEN
            deaktiviert = KONST.FARBE_DEAKTIVIERT
            
            sett = settings_orga['organon_farben']['office']
            
            def get_farbe(value):
                if isinstance(value, int):
                    return value
                else:
                    return settings_orga['organon_farben'][value]
            
            # Kann button_schrift evt. herausgenommen werden?
            button_schrift = get_farbe(sett['button_schrift'])
            
            #statusleiste_schrift = get_farbe(sett['statusleiste_schrift'])
            #statusleiste_hintergrund = get_farbe(sett['statusleiste_hintergrund'])
            
            felder_hintergrund = get_farbe(sett['felder_hintergrund'])
            felder_schrift = get_farbe(sett['felder_schrift'])
            
            # Sidebar
            sidebar_eigene_fenster_hintergrund = get_farbe(sett['sidebar']['eigene_fenster_hintergrund'])
            sidebar_selected_hintergrund = get_farbe(sett['sidebar']['selected_hintergrund'])
            sidebar_selected_schrift = get_farbe(sett['sidebar']['selected_schrift'])
            sidebar_schrift = get_farbe(sett['sidebar']['schrift'])
            
            #trenner_licht = get_farbe(sett['trenner_licht'])
            #trenner_schatten = get_farbe(sett['trenner_schatten'])
            
            # Lineal
            OO_anfasser_trenner = get_farbe(sett['OO_anfasser_trenner'])
            OO_lineal_tab_zwischenraum = get_farbe(sett['OO_lineal_tab_zwischenraum'])
            OO_schrift_lineal_sb_liste = get_farbe(sett['OO_schrift_lineal_sb_liste'])
            
            LO_anfasser_text = get_farbe(sett['LO_anfasser_text'])
            #LO_tabsumrandung = get_farbe(sett['LO_tabsumrandung'])
            #LO_lineal_bg_innen = get_farbe(sett['LO_lineal_bg_innen'])
            #LO_tab_fuellung = get_farbe(sett['LO_tab_fuellung'])
            #LO_tab_trenner = get_farbe(sett['LO_tab_trenner'])
            
            
            LO = ('LibreOffice' in frame.Title)
            
            STYLES = {  
                      # Allgemein
                        'ButtonRolloverTextColor' : button_schrift, # button rollover
                        
                        'FieldColor' : felder_hintergrund, # Hintergrund Eingabefelder
                        'FieldTextColor' : felder_schrift,# Schrift Eingabefelder
                        
                        # Trenner
                        'LightColor' : linien, # Fenster Trenner
                        'ShadowColor' : linien, # Fenster Trenner
                        
                        # OO Lineal + Trenner
                         
                        'DarkShadowColor' : (LO_anfasser_text if LO    # LO Anfasser + Lineal Text
                                            else OO_anfasser_trenner), # OO Anfasser +  Document Fenster Trenner 
                        'WindowTextColor' : (schrift if LO      # Felder (Navi) Schriftfarbe Sidebar 
                                             else OO_schrift_lineal_sb_liste),     # Felder (Navi) Schriftfarbe Sidebar + OO Lineal Schriftfarbe   
                            
                        # Sidebar
                        'LabelTextColor' : sidebar_schrift, # Schriftfarbe Sidebar + allg Dialog
                        'DialogColor' : sidebar_eigene_fenster_hintergrund, # Hintergrund Sidebar Dialog
                        'FaceColor' : (schrift if LO        # LO Formatvorlagen Treeview Verbinder
                                        else hf),           # OO Hintergrund Organon + Lineal + Dropdowns  
                        'WindowColor' : (hf if LO                           # LO Dialog Hintergrund
                                        else OO_lineal_tab_zwischenraum),   # OO Lineal Tabzwischenraum
                        'HighlightColor' : sidebar_selected_hintergrund, # Sidebar selected Hintergrund
                        'HighlightTextColor' : sidebar_selected_schrift, # Sidebar selected Schrift
                        
                        
    #                     'ActiveBorderColor' : rot,#k.A.
    #                     'ActiveColor' : rot,#k.A.
    #                     'ActiveTabColor' : rot,#k.A.
    #                     'ActiveTextColor' : rot,#k.A.
    #                     'ButtonTextColor' : rot,# button Textfarbe / LO Statuszeile Textfarbe
    #                     'CheckedColor' : rot,#k.A.
    #                     'DeactiveBorderColor' : rot,#k.A.
    #                     'DeactiveColor' : rot,#k.A.
    #                     'DeactiveTextColor' : rot,#k.A.
                        'DialogTextColor' : rot,#k.A.
                        'DisableColor' : deaktiviert,
    #                     'FieldRolloverTextColor' : rot,#k.A.
    #                     'GroupTextColor' : rot,#k.A.
    #                     'HelpColor' : rot,#k.A.
    #                     'HelpTextColor' : rot,#k.A.
    #                     'InactiveTabColor' : rot,#k.A.
    #                     'InfoTextColor' : rot,#k.A.
    #                     'MenuBarColor' : rot,#k.A.
    #                     'MenuBarTextColor' : rot,#k.A.
    #                     'MenuBorderColor' : rot,#k.A.
                        'MenuColor' : rot,#k.A.
                        'WindowColor' : hf,#k.A.
    
    #                     'MenuHighlightColor' : rot,#k.A.
    #                     'MenuHighlightTextColor' : rot,#k.A.
                        'MenuTextColor' : schrift,#k.A.
    #                     'MonoColor' : rot, #k.A.
                        'RadioCheckTextColor' : schrift,#k.A.
    #                     'WorkspaceColor' : rot, #k.A.
    #                     erzeugen Fehler:
    #                     'FaceGradientColor' : 502,
                        'SeparatorColor' : 502,                    
                        }
            
     
            def stilaenderung(win,ignore=[]):
    
                for s in STYLES:
                    if s in ignore: 
                        pass
                    else:
                        try:
                            val = STYLES[s]
                            setattr(win.StyleSettings, s, val)
                        except Exception as e:
                            pass
                        
                    win.Model.BackgroundColor = hf 
                    #win.setForeground(statusleiste_schrift)     # Schrift Statuszeile
            
    
    
            
            # folgende Properties wuerden die Eigenschaften
            # der Office Menubar und aller Buttons setzen
            ignore = ['ButtonTextColor',
                     'LightColor',
                     'MenuBarTextColor',
                     'MenuBorderColor',
                     'ShadowColor'
                     ]
    
    
            
            stilaenderung(win)
            
        except Exception as e:
            log(inspect.stack,tb())
      
                
from com.sun.star.awt import XMouseListener, XActionListener, XItemListener
class Listener_Organon_Farben(unohelper.Base,XMouseListener,XActionListener):
    
    def __init__(self,mb,auswahl_item_listener):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = {}
        self.RBs = {}
        self.auswahl_item_listener = auswahl_item_listener
        self.lb_dict = None
        self.verwendetes_persona = None
        self.bearbeitungs_fenster = None
        
              
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack) 

        try:

            for c in self.ctrls:
                if self.ctrls[c] == ev.Source:

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
                        
                    elif c == 'tabs_sel_hintergrund':
                        self.setze_farbe_sel_tab_hintergrund(farbe)
                    elif c == 'tabs_sel_schrift':
                        self.setze_farbe_sel_tab_schrift(farbe)
                    elif c == 'tabs_hintergrund':
                        self.setze_farbe_tab_hintergrund(farbe)
                    elif c == 'tabs_schrift':
                        self.setze_farbe_tab_schrift(farbe)
                    elif c == 'tabs_trenner':
                        self.setze_farbe_tab_trenner(farbe)
                        
                    elif c == 'linien':
                        self.setze_farbe_linien(farbe)
                        
                    elif c == 'deaktiviert':
                        self.setze_farbe_deaktiviert(farbe)
                    
                    ev.Source.Model.BackgroundColor = farbe  
                    self.mb.schreibe_settings_orga()
                    break
        except:
            log(inspect.stack,tb())   
            
            
            
            setze_farbe(self.setze_farbe_sel_tab_hintergrund,'tabs_sel_hintergrund')
            setze_farbe(self.setze_farbe_sel_tab_schrift,'tabs_sel_schrift')
            setze_farbe(self.setze_farbe_tab_hintergrund,'tabs_hintergrund')
            setze_farbe(self.setze_farbe_tab_schrift,'tabs_schrift')
            setze_farbe(self.setze_farbe_sel_tab_trenner,'tabs_trenner')
    
    def waehle_farbe(self,ev,art):  
        if self.mb.debug: log(inspect.stack)
        print(art)
        farbe = self.mb.class_Funktionen.waehle_farbe(self.mb.settings_orga['organon_farben'][art])
        
        self.mb.settings_orga['organon_farben'][art] = farbe
        self.mb.schreibe_settings_orga()
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
                
            elif cmd == 'writer_design_bearbeiten':
                loc = ev.Source.Context.Context.AccessibleContext.AccessibleParent.PosSize
                self.writer_design_bearbeiten(loc)
                
            elif cmd == 'writer_design':
                self.nutze_writer_design(ev.Source)
            
            elif cmd == 'fenster_design':
                design = self.mb.settings_orga['organon_farben']['design_organon_fenster']
                design = 0 if design else 1
                self.mb.settings_orga['organon_farben']['design_organon_fenster'] = design
                
                
            else:
                self.setze_design(cmd,ev)
        except:
            log(inspect.stack,tb()) 
            
            
    def writer_design_bearbeiten(self,loc):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.class_Organon_Design.setze_listboxen = True

            sett = self.mb.settings_orga
            
            ctrls,max_hoehe,max_breite = self.dialog_writer_design()

            x,y = loc.X + loc.Width + 20,loc.Y

               
            # Hauptfenster erzeugen
            posSize = x,y,max_breite,max_hoehe + 20
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            self.bearbeitungs_fenster = fenster
            
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                fenster_cont.addControl(c,ctrls[c])

            
            # Listboxeintraege auswaehlen            
            def lb_auswaehlen(d):                
                for k in d:
                    ct = 'control_{}LB'.format(k)
                    if ct in ctrls:
                        if d[k] in self.lb_dict:
                            ctrls[ct].selectItem(self.lb_dict[d[k]],True)
                                            
                    
            office = sett['organon_farben']['office']
            sb = office['sidebar']  
            
            lb_auswaehlen(office)
            lb_auswaehlen(sb)
            
            if self.mb.programm == 'LibreOffice':
                # Persona auswaehlen
                persona_url = sett['organon_farben']['office']['persona_url']
                persona = persona_url.split('/')[0]
    
                ctrl_persona = ctrls['control_personas_list']
                if persona in ctrl_persona.Items:
                    ctrl_persona.selectItem(persona,True)
            
            self.mb.class_Organon_Design.setze_listboxen = False
            
        except:
            log(inspect.stack,tb()) 
            

    def dialog_writer_design(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Listener_Writer_Design(self.mb)

            controls,max_breite = self.dialog_writer_design_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)  
                
            listener.ctrls = ctrls

            return ctrls, max_hoehe ,max_breite
  
        except:
            log(inspect.stack,tb())
            # bei einer Exception scheitert die weitere Anzeige.
            # Es fehlt ein Fallback
            return [], 0 


    def dialog_writer_design_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        from collections import OrderedDict
                
        controls = [
            ]
        
        
        LO = (self.mb.programm == 'LibreOffice')
        
        sett = self.mb.settings_orga['organon_farben']['office']

        cts1 = OrderedDict()
        cts1['1tit'] = ['tit',LANG.STATUSLEISTE]
        cts1['statusleiste_schrift'] = [LANG.SCHRIFT,sett['statusleiste_schrift']]
        cts1['statusleiste_hintergrund'] = [LANG.HINTERGRUND,sett['statusleiste_hintergrund']]
        
        cts1['2sep'] = ['sep',]
        
        cts1['3tit'] = ['tit',LANG.FELDER]
        cts1['felder_schrift'] = [LANG.SCHRIFT,sett['felder_schrift']]
        cts1['felder_hintergrund'] = [LANG.HINTERGRUND,sett['felder_hintergrund']]
        
        cts1['4sep'] = ['sep',]
        
        cts1['5tit'] = ['tit',LANG.TRENNER]
        cts1['trenner_licht'] = [LANG.TRENNER_LICHT,sett['trenner_licht']]
        cts1['trenner_schatten'] = [LANG.TRENNER_SCHATTEN,sett['trenner_schatten']]
        
        cts1['6sep'] = ['sep',]
        
        cts1['7tit'] = ['tit',LANG.LINEAL]
        cts1['OO_anfasser_trenner'] = [LANG.ANFASSER,sett['OO_anfasser_trenner']]                     if not LO else [] 
        cts1['OO_lineal_tab_zwischenraum'] = [LANG.ZWISCHENRAUM,sett['OO_lineal_tab_zwischenraum']]   if not LO else [] 
        cts1['OO_schrift_lineal_sb_liste'] = [LANG.SCHRIFT,sett['OO_schrift_lineal_sb_liste']]        if not LO else [] 
        cts1['LO_anfasser_text'] = [LANG.ANFASSER_TEXT,sett['LO_anfasser_text']]        if LO else [] 
        cts1['LO_tabsumrandung'] = [LANG.TABSUMRANDUNG,sett['LO_tabsumrandung']]        if LO else [] 
        cts1['LO_lineal_bg_innen'] = [LANG.LINEAL_BG_INNEN,sett['LO_lineal_bg_innen']]  if LO else [] 
        cts1['LO_tab_fuellung'] = [LANG.TAB_FUELLUNG,sett['LO_tab_fuellung']]           if LO else [] 
        cts1['LO_tab_trenner'] = [LANG.TAB_TRENNER,sett['LO_tab_trenner']]              if LO else [] 
        

        def get_farbe(value):
            try:
                if isinstance(value, int):
                    return value,None
                else:
                    return self.mb.settings_orga['organon_farben'][value],value
            except:
                log(inspect.stack,tb())
                return 0,None
                
        
        lb_items = (
                    LANG.MENULEISTE  + ' ' + LANG.HINTERGRUND,
                    LANG.MENULEISTE  + ' ' + LANG.SCHRIFT,
                    LANG.BAUMANSICHT + ' ' + LANG.HINTERGRUND,
                    LANG.BAUMANSICHT + ' ' + LANG.ORDNER,
                    LANG.BAUMANSICHT + ' ' + LANG.DATEI,
                    LANG.ZEILE       + ' ' + LANG.AUSGEWAEHLTE_ZEILE,
                    LANG.ZEILE       + ' ' + LANG.EDITIERTE_ZEILE,
                    LANG.ZEILE       + ' ' + LANG.GEZOGENE_ZEILE,
                    LANG.GLIEDERUNG,
                    LANG.TRENNER     + ' ' + LANG.HINTERGRUND,
                    LANG.TRENNER     + ' ' + LANG.SCHRIFT,
                    
                    LANG.TABS     + ' ' + LANG.HINTERGRUND,
                    LANG.TABS     + ' ' + LANG.SCHRIFT,
                    LANG.TABS     + ' ' + LANG.AUSGEWAEHLTE_ZEILE + ' ' + LANG.HINTERGRUND_ABK,
                    LANG.TABS     + ' ' + LANG.AUSGEWAEHLTE_ZEILE + ' ' + LANG.SCHRIFT,
                    LANG.TABS     + ' ' + LANG.AUSGEWAEHLTE_ZEILE + ' ' + LANG.TRENNER
                    )
        
        lb_items_syn = (
                        "menu_hintergrund",
                        "menu_schrift",
                        "hf_hintergrund",
                        "schrift_ordner",
                        "schrift_datei",
                        "ausgewaehlte_zeile",
                        "editierte_zeile",
                        "gezogene_zeile",
                        "gliederung",
                        "trenner_farbe_hintergrund",
                        "trenner_farbe_schrift",
                        
                        "tabs_hintergrund",
                        "tabs_schrift",
                        "tabs_sel_hintergrund",
                        "tabs_sel_schrift",
                        "linien",
                        )
        
        self.lb_dict = {lb_items_syn[i] : lb_items[i] for i in range(len(lb_items))}
        self.lb_dict2 = {lb_items[i] : lb_items_syn[i]  for i in range(len(lb_items))}
        
        
        def erzeuge_conts(cts,tabs,abstand_nach_sep=0):
            
            controls = []
            
            for c in cts:
                
                tab_line = '{0}-{1}-E'.format(tabs[0],tabs[2])
                
                if cts[c] == []: continue
                
                elif 'sep' in cts[c][0]:
                    cont = [
                            ('control_{}'.format(c),"FixedLine",0,         
                                    tab_line,20,20,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
                            abstand_nach_sep,
                            ]
                    controls.extend(cont)
                    
                elif 'tit' in cts[c][0]:
                    cont = [
                            30,
                            ('control_{}'.format(c),"FixedText",1,        
                                        tabs[2],0,50,20,  
                                        ('Label','FontWeight'),
                                        (cts[c][1],150),                                                                    
                                        {} 
                                        ), 
                            ]
                    controls.extend(cont)
                    
                else:
                    
                    farbe,sel = get_farbe(cts[c][1])

                    cont = [
                            20,
                            ('control_{}'.format(c),"FixedText",0,        
                                        tabs[0],0,32,16,  
                                        ('BackgroundColor','Label','Border'),
                                        (farbe,'    ',1),       
                                        {'addMouseListener':(listener)} 
                                        ), 
                            0,                                                  
                            ('control_{}LB'.format(c),"ListBox",0,      
                                        tabs[1],0,40,15,    
                                        ('Border','Dropdown','LineCount'),
                                        ( 2,True,16),       
                                        {'addItems':(lb_items,0), 'addItemListener':listener} 
                                        ), 
                            0,
                            ('control_{}L'.format(c),"FixedText",1,        
                                        tabs[2],0,20,20,  
                                        ('Label',),
                                        (cts[c][0],),                                                                    
                                        {} 
                                        ),
                            ]
                    
                    
                    controls.extend(cont)
            
            return controls
        
        conts = erzeuge_conts(cts1,['tab0','tab1','tab2'])
        controls.extend(conts)

        
        # SEITENLEISTE
        
        cts3 = OrderedDict()
        cts3[''] = ['tit',LANG.SIDEBAR]
        cts3['hintergrund'] = [LANG.HINTERGRUND,sett['sidebar']['hintergrund']]
        cts3['schrift'] = [LANG.SCHRIFT,sett['sidebar']['schrift']]
        cts3['1'] = ['sep',]
        cts3['titel_hintergrund'] = [LANG.TITEL_HINTERGRUND,sett['sidebar']['titel_hintergrund']]
        cts3['titel_schrift'] = [LANG.TITEL_SCHRIFT,sett['sidebar']['titel_schrift']]
        cts3['2'] = ['sep',]
        cts3['panel_titel_hintergrund'] = [LANG.PANEL_TITEL_HINTERGRUND,sett['sidebar']['panel_titel_hintergrund']]
        cts3['panel_titel_schrift'] = [LANG.PANEL_TITEL_SCHRIFT,sett['sidebar']['panel_titel_schrift']]
        cts3['3'] = ['sep',]
        cts3['leiste_hintergrund'] = [LANG.LEISTE_HINTERGRUND,sett['sidebar']['leiste_hintergrund']]
        cts3['leiste_selektiert_hintergrund'] = [LANG.LEISTE_SELEKTIERT_HINTERGRUND,sett['sidebar']['leiste_selektiert_hintergrund']]
        cts3['leiste_icon_umrandung'] = [LANG.LEISTE_ICON_UMRANDUNG,sett['sidebar']['leiste_icon_umrandung']]
        cts3['4'] = ['sep',]
        cts3['border_horizontal'] = [LANG.BORDER_HORIZONTAL,sett['sidebar']['border_horizontal']]
        cts3['border_vertical'] = [LANG.BORDER_VERTICAL,sett['sidebar']['border_vertical']]
        cts3['5'] = ['sep',]
        cts3['eigene_fenster_hintergrund'] = [LANG.EIGENE_FENSTER_HINTERGRUND,sett['sidebar']['eigene_fenster_hintergrund']]
        cts3['6'] = ['sep',]
        cts3['selected_hintergrund'] = [LANG.SELECTED_HINTERGRUND,sett['sidebar']['selected_hintergrund']]
        cts3['selected_schrift'] = [LANG.SELECTED_SCHRIFT,sett['sidebar']['selected_schrift']]
         
        
        conts = erzeuge_conts(cts3,['tab3','tab4','tab5'],10)
        conts[0] = 'Y=30'
        controls.extend(conts)


        # PERSONAS
        
        if LO:
            personas_dict,personas_list,gallery_user_path = self.mb.class_Organon_Design.get_personas() 
            
            cts = [
                   'Y=30',
                    ('control_VP',"FixedText",1,        
                                            'tab6x',0,50,20,  
                                            ('Label','FontWeight'),
                                            (LANG.PERSONA,150),                                                                    
                                            {} 
                                            ), 
                   0,
                   ('control_use_persona',"CheckBox",1,       
                                            'tab8',0,50,20,  
                                            ('Label','State'),
                                            ( LANG.NUTZE_PERSONA,sett['nutze_personas']),       
                                            {'addActionListener':(listener)}
                                            ), 
                    20,
                    ('control_personas_list',"ListBox",0,        
                                            'tab6x-tab7-E',0,102,20,  
                                            ('Border','Dropdown','LineCount'),
                                            ( 2,True,10),       
                                            {'addItems':(tuple(personas_list),0),'addItemListener':listener}
                                            ), 
                   
                   30,
                   ('controlFix2',"FixedLine",0,         
                                    'tab6-max',20,20,1,   
                                    (),
                                    (),                                                  
                                    {} 
                                    ) ,
                   -24,
                   ]
            
            controls.extend(cts)
            
            cts2 = {'persona_schrift' : [LANG.SCHRIFT,sett['persona_schrift']], }
            conts = erzeuge_conts(cts2,['tab6','tab7','tab8'])
            controls.extend(conts)
        
        
        # DOCUMENT/DOCUMENTBACKGROUND
        if LO:
            Y = 30
        else:
            Y = 'Y=30'
            
        cts = [
               Y,
                ('control_DB',"FixedText",1,        
                                        'tab6x',0,50,20,  
                                        ('Label','FontWeight'),
                                        (LANG.DOCUMENT,150),                                                                    
                                        {} 
                                        ), 
               0,
               ('control_nutze_dok_farbe',"CheckBox",1,        
                                        'tab8',0,50,20,  
                                        ('Label','State'),
                                        ( LANG.NUTZE_DOK_FARBE,sett['nutze_dok_farbe']),       
                                        {'addActionListener':(listener)}
                                        ), 
                0,
                ]
        
        controls.extend(cts)
                
        cts4 = OrderedDict()
        cts4['dok_hintergrund'] = [LANG.HINTERGRUND,sett['dok_hintergrund']]
        cts4['dok_farbe'] = [LANG.DOCUMENT,sett['dok_farbe']]
                
        conts = erzeuge_conts(cts4,['tab6','tab7','tab8'])
        controls.extend(conts)      
        
        
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 5),
                 1 : (None, 5),
                 2 : (None, 30),
                 3 : (None, 5),
                 4 : (None, 5),
                 5 : (None, 30),
                 6 : (None, 5),
                 7 : (None, 5),
                 8 : (None, 0),
                 }
        
        abstand_links = 10
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
        
        listener.tabs = tabs3        
        
        return controls2,max_breite
    
    
    def nutze_writer_design(self,src):
        if self.mb.debug: log(inspect.stack)

        if src.State:
            Popup(self.mb, 'info').text = LANG.WRITER_DESIGN_INFO
            if self.mb.programm == 'LibreOffice':
                self.mb.class_Organon_Design.kopiere_personas()
            

        self.mb.settings_orga['organon_farben']['design_office'] = src.State
        self.mb.schreibe_settings_orga()

    
    def import_design(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
        
            sett = self.mb.settings_orga['designs'] 
            
            pfad = self.mb.class_Funktionen.filepicker()
            if pfad == None: return
            
            odict = self.mb.class_Funktionen.oeffne_json(pfad)
            if odict == None:
                Popup(self.mb, 'error').text = LANG.KEINE_JSON_DATEI
                return
            
            
            vorbild = sett[list(sett)[0]]
            
            
            for k in odict:
                neu = copy.deepcopy(k)
                while k in sett:
                    k = k+'x'
                
    
                # ueberpruefen, ob ein Wert fehlt
                for key,value in vorbild['office']['sidebar'].items():
                    if key not in odict[neu]['office']['sidebar']:
                        odict[neu]['office']['sidebar'][key] = value
                        
                for key,value in vorbild['office'].items():
                    if key == 'sidebar': continue
                    if key not in odict[neu]['office']:
                        odict[neu]['office'][key] = value
                        
                for key,value in vorbild.items():
                    if key == 'office': continue
                    if key not in odict[neu]:
                        odict[neu][key] = value
                
                
                sett.update({ k : odict[neu] })
            
            self.mb.schreibe_settings_orga()
            self.ansicht_erneuern()
            
        except:
            log(inspect.stack,tb())
            
    
    def export_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs'] 
        
        name = self.ctrls['textfeld'].Model.Text
        if name == '':
            Popup(self.mb, 'info').text = LANG.EXPORTNAMEN_EINGEBEN
            return
        
        pfad = self.mb.class_Funktionen.folderpicker()
        
        if pfad == None:
            return
        
        pfad = os.path.join(pfad,name+'.json')
        
        if os.path.exists(pfad):
            # 16777216 Flag fuer YES_NO
            entscheidung = self.mb.entscheidung(LANG.DATEI_EXISTIERT,"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return
            
        with open(pfad, 'w') as outfile:
            json.dump(sett, outfile,indent=4, separators=(',', ': '))
        

    def loesche_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs']  
        
        if len(sett) < 2:
            Popup(self.mb, 'warning').text = LANG.NICHT_ALLE_DESIGNS_LOESCHEN
            return
        
        active,act_name = None, None
        
        for c in self.RBs:
            if self.RBs[c].State:
                active,act_name = self.RBs[c], self.RBs[c].Model.Label

        self.RBs.pop(act_name, None)
        self.mb.settings_orga['designs'].pop(act_name,None)
        
        self.ansicht_erneuern()
             
    
    def erzeuge_neues_design(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['designs']
        
        name = self.ctrls['textfeld'].Model.Text
        if name == '':
            Popup(self.mb, 'info').text = LANG.DESIGNNAMEN_EINGEBEN
            return
        if name in sett:
            Popup(self.mb, 'info').text = LANG.DESIGN_EXISTIERT
            return
        
        new_set = copy.deepcopy(self.mb.settings_orga['organon_farben'])
        new_set.pop('aktiv',None)
        
        sett.update({ name : new_set })
        self.mb.schreibe_settings_orga()
        
        self.ansicht_erneuern()
        
            
    def setze_design(self,cmd,ev):
        if self.mb.debug: log(inspect.stack)

        try:

            sett = self.mb.settings_orga['designs'][cmd]
            
            def setze_farbe(fkt,value):
                fkt(sett[value])
                self.ctrls[value].Model.BackgroundColor = sett[value]
                
            def setze_farbe2(konst,value):
                setattr(KONST, konst, sett[value])
                self.ctrls[value].Model.BackgroundColor = sett[value]
            
            setze_farbe(self.setze_farbe_hintergrund,'hf_hintergrund')            
            setze_farbe(self.setze_farbe_menuleiste_hintergrund,'menu_hintergrund')
            setze_farbe(self.setze_farbe_menuleiste_schrift,'menu_schrift')
            setze_farbe(self.setze_farbe_schrift_dateien,'schrift_datei')
            setze_farbe(self.setze_farbe_schrift_ordner,'schrift_ordner')
            setze_farbe(self.setze_farbe_gliederung,'gliederung')
            
            setze_farbe(self.setze_farbe_sel_tab_hintergrund,'tabs_sel_hintergrund')
            setze_farbe(self.setze_farbe_sel_tab_schrift,'tabs_sel_schrift')
            setze_farbe(self.setze_farbe_tab_hintergrund,'tabs_hintergrund')
            setze_farbe(self.setze_farbe_tab_schrift,'tabs_schrift')
            setze_farbe(self.setze_farbe_tab_trenner,'tabs_trenner')
            setze_farbe(self.setze_farbe_linien,'linien')
            setze_farbe(self.setze_farbe_deaktiviert,'deaktiviert')
                        
            setze_farbe2('FARBE_AUSGEWAEHLTE_ZEILE','ausgewaehlte_zeile')            
            setze_farbe2('FARBE_EDITIERTE_ZEILE','editierte_zeile')
            setze_farbe2('FARBE_GEZOGENE_ZEILE','gezogene_zeile')
            setze_farbe2('FARBE_TRENNER_HINTERGRUND','trenner_farbe_hintergrund')
            setze_farbe2('FARBE_TRENNER_SCHRIFT','trenner_farbe_schrift')
            
            self.setze_farbe_trenner(sett['trenner_farbe_hintergrund'],'trenner_farbe_hintergrund')
            
            use_design = self.mb.settings_orga['organon_farben']['design_office']
                     
            self.mb.settings_orga.update({'organon_farben': copy.deepcopy(sett) })
            self.mb.settings_orga['organon_farben'].update({'aktiv': cmd})
            
            self.mb.settings_orga['organon_farben']['design_office'] = use_design
            sett['design_office'] = use_design

            if sett['design_office']:
                
                if self.mb.programm == 'LibreOffice':
                    if sett['office']['nutze_personas']:
                        self.mb.class_Organon_Design.setze_persona(sett['office']['persona_url'])
                    else:
                        self.mb.class_Organon_Design.setze_persona('','')
                        
                if sett['office']['nutze_dok_farbe']:
                    dok_farbe = sett['office']['dok_farbe']
                    dok_hintergrund = sett['office']['dok_hintergrund']
                    self.mb.class_Organon_Design.setze_dok_farben(dok_farbe,dok_hintergrund)
                else:
                    self.mb.class_Organon_Design.setze_dok_farben(None,None)
                    
            
            if self.bearbeitungs_fenster != None:
                self.bearbeitungs_fenster.dispose()
                loc = ev.Source.Context.Context.AccessibleContext.AccessibleParent.PosSize
                self.writer_design_bearbeiten(loc)
            
            self.mb.schreibe_settings_orga()
            
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
            if hf == None: continue
            menuleiste = hf.Context.Context.getControl('Organon_Menu_Bar')
            menuleiste.Model.BackgroundColor = farbe
        
            try:
                for c in menuleiste.Controls:
                    c.Model.BackgroundColor = farbe
            except Exception as e:
                pass
                #print(e)
                
        KONST.FARBE_MENU_HINTERGRUND = farbe
        
            
    def setze_farbe_menuleiste_schrift(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            if hf == None: continue
            menuleiste = hf.Context.Context.getControl('Organon_Menu_Bar')
        
            try:
                for c in menuleiste.Controls:
                    try:
                        c.Model.TextColor = farbe
                    except:
                        pass
            except Exception as e:
                pass
                #print(e)
                
        KONST.FARBE_MENU_SCHRIFT = farbe
        
       
    def setze_farbe_hintergrund(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        # Dialog Fenster
        self.mb.prj_tab.Model.BackgroundColor = farbe
        
        for tab in self.mb.props:
            # Hauptfeld
            hf = self.mb.props[tab].Hauptfeld
            if hf == None: continue
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
            if hf == None: continue
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
            if hf == None: continue
            
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
            if hf == None: continue
            
            # Zeilen im Hauptfeld
            zeilen = hf.Controls
            
            for z in zeilen:   
                gliederung = z.getControl('tag3')
                if gliederung != None:
                    gliederung.Model.TextColor = farbe
        
        KONST.FARBE_GLIEDERUNG = farbe
        
    def setze_farbe_sel_tab_hintergrund(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        tab_ctrl = self.mb.tabsX.tableiste.getControl(T.AB)
        tab_ctrl.Model.BackgroundColor = farbe
        
        KONST.FARBE_TABS_SEL_HINTERGRUND = farbe
        
    def setze_farbe_sel_tab_schrift(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        tab_ctrl = self.mb.tabsX.tableiste.getControl(T.AB)
        tab_ctrl.Model.TextColor = farbe
        
        KONST.FARBE_TABS_SEL_SCHRIFT = farbe
        
    def setze_farbe_tab_hintergrund(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            if tab != T.AB:
                tab_ctrl = self.mb.tabsX.tableiste.getControl(tab)
                tab_ctrl.Model.BackgroundColor = farbe
        
        KONST.FARBE_TABS_HINTERGRUND = farbe
        
    def setze_farbe_tab_schrift(self,farbe):
        if self.mb.debug: log(inspect.stack)
        
        for tab in self.mb.props:
            if tab != T.AB:
                tab_ctrl = self.mb.tabsX.tableiste.getControl(tab)
                tab_ctrl.Model.TextColor = farbe
        
        KONST.FARBE_TABS_SCHRIFT = farbe
        
    def setze_farbe_tab_trenner(self,farbe):
        if self.mb.debug: log(inspect.stack)

        self.mb.tabsX.tableiste.Model.BackgroundColor = farbe
        
        KONST.FARBE_TABS_TRENNER = farbe

    def setze_farbe_linien(self,farbe):
        if self.mb.debug: log(inspect.stack)

        KONST.FARBE_LINIEN = farbe
        
    def setze_farbe_deaktiviert(self,farbe):
        if self.mb.debug: log(inspect.stack)

        KONST.FARBE_DEAKTIVIERT = farbe



class Listener_Writer_Design(unohelper.Base,XItemListener,XMouseListener,XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb  
        self.ctrls = {}

        
    def itemStateChanged(self, ev): 
        # Wenn die LBs nicht vom user gesetzt werden, sondern
        # von Organon, keine Aktion ausfuehren 
        if self.mb.class_Organon_Design.setze_listboxen == True: return
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['organon_farben']['office']
        LO = (self.mb.programm == 'LibreOffice')
        
        if LO:
            if ev.Source == self.ctrls['control_personas_list']:
                self.setze_persona()
                return

        # Farbe des Farbfeldes anhand der Listenauswahl neu setzen
        for c in self.ctrls:
            if self.ctrls[c] == ev.Source:
                selektiert = ev.Source.SelectedItem
                prop = c.replace('LB','')
                farbfeld = self.ctrls[prop]
                farb_prop = self.mb.class_Organon_Design.lb_dict2[selektiert]
                farbe = self.mb.settings_orga['organon_farben'][farb_prop]
                farbfeld.Model.BackgroundColor = farbe
                
                # prop in organon_farben eintragen
                prop2 = c.replace('control_','').replace('LB','')
                
                if prop2 in sett:
                    sett[prop2] = farb_prop
                elif prop2 in sett['sidebar']:
                    sett['sidebar'][prop2] = farb_prop
                
                if LO:    
                    if prop2 == 'persona_schrift':
                        if sett['nutze_personas']:
                            self.setze_persona()
                            return

                self.mb.schreibe_settings_orga()
                return
                    
    
    def setze_persona(self):
        if self.mb.debug: log(inspect.stack) 
        
        personas_dict,personas,gal_user_path = self.mb.class_Organon_Design.get_personas() 
        
        neues_pers = self.ctrls['control_personas_list'].SelectedItem
        png = os.path.basename(personas_dict[neues_pers])
        farbe = self.ctrls['control_persona_schrift'].Model.BackgroundColor
        hex_farbe = self.mb.class_Funktionen.dezimal_to_hex(farbe)
        pers_url = '{0}/{1};{0}/{1};#{2};#636363'.format(neues_pers,png,hex_farbe)
        
        sett = self.mb.settings_orga['organon_farben']['office']
        sett['persona_schrift'] = farbe
        sett['persona_url'] = pers_url
        if sett['nutze_personas']:
            self.mb.class_Organon_Design.setze_persona(pers_url)

        self.mb.schreibe_settings_orga()
        
                    
    def disposing(self,ev):
        return False    
  
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack) 
        
        farbe = self.mb.class_Funktionen.waehle_farbe(ev.Source.Model.BackgroundColor)
        ev.Source.Model.BackgroundColor = farbe
        
        LO = (self.mb.programm == 'LibreOffice')
        
        for c in self.ctrls:
            if self.ctrls[c] == ev.Source:
                listbox = self.ctrls[c+'LB']
                self.mb.class_Organon_Design.setze_listboxen = True
                listbox.selectItem(listbox.SelectedItem,False)
                self.mb.class_Organon_Design.setze_listboxen = False
                
                prop = c.replace('control_','')
                sett = self.mb.settings_orga['organon_farben']['office']

                if prop in sett:
                    sett[prop] = farbe
                elif prop in sett['sidebar']:
                    sett['sidebar'][prop] = farbe
                
                if LO:    
                    if prop == 'persona_schrift':
                        if sett['nutze_personas']:
                            self.setze_persona()
                            return

                self.mb.schreibe_settings_orga()

                return

    
    def mouseExited(self, ev):
        return False
    def mouseEntered(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        label = ev.Source.Model.Label
        sett = self.mb.settings_orga['organon_farben']['office']
        
        if label == LANG.NUTZE_PERSONA:
            sett['nutze_personas'] = ev.Source.State
            if ev.Source.State:
                self.mb.class_Organon_Design.setze_persona(sett['persona_url'])
            else:
                self.mb.class_Organon_Design.setze_persona('','')
        elif label == LANG.NUTZE_DOK_FARBE:
            sett['nutze_dok_farbe'] = ev.Source.State
            
            dok_farbe = self.ctrls['control_dok_farbe'].Model.BackgroundColor
            dok_hintergrund = self.ctrls['control_dok_hintergrund'].Model.BackgroundColor
            
            if ev.Source.State:
                self.mb.class_Organon_Design.setze_dok_farben(dok_farbe,dok_hintergrund)
            else:
                self.mb.class_Organon_Design.setze_dok_farben(None,None)
            
        self.mb.schreibe_settings_orga()
                
        
class Listener_Persona(unohelper.Base,XItemListener,XActionListener,XMouseListener):
    
    def __init__(self,mb,personas_dict,personas_list,gallery_user_path):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = {}
        self.personas_dict = personas_dict
        self.personas_list = personas_list
        self.gallery_user_path = gallery_user_path
        self.ctrl_container = None
        self.first_dispose = True
        
        
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)

        try:
            auswahl = self.personas_list[ev.Selected]
            url = uno.systemPathToFileUrl(self.personas_dict[auswahl])
            
            self.theme_ansicht_erneuern(url)

            self.ctrls['control_solid'].State = 0
            self.ctrls['control_gradient'].State = 0
            self.ctrls['control_url_nutzer'].State = 0
            
        except:
            log(inspect.stack,tb())
    
    
    def theme_ansicht_erneuern(self,url=''):
        if self.mb.debug: log(inspect.stack)
        
        if url == 'unveraendert':
            url = self.ctrls['control_theme'].Model.ImageURL    
        
        
        X = self.ctrls['control_theme'].PosSize.X
        breite = self.ctrl_container.Size.Width-6
        
        self.ctrls['control_theme'].dispose()
        
        bg_color = self.ctrls['control_farbe_hintergrund'].Model.BackgroundColor
        
        controls = [
                1,
                ('control_theme',"ImageControl",0,        
                    X,0,breite,80,  
                    ('Label','FontWeight','ImageURL','Border','BackgroundColor'),
                    (LANG.ORGANON_DESIGN ,150,url,0,bg_color),                                             
                    {} 
                    ),
                ]

        ctrls,pos_y = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)
        
        self.ctrl_container.addControl('control_theme',ctrls['control_theme'])
        self.ctrls.update({'control_theme': ctrls['control_theme']})
        
         
        for c in self.ctrls['control_theme_tit'],self.ctrls['control_theme_untertit']:
            c.Model.TextColor = self.ctrls['control_farbe_schrift'].Model.BackgroundColor

            
    def disposing(self,ev):
        
        if self.first_dispose:
            self.first_dispose = False
            self.loesche_pngs_in_tmp_ordner
        return False
    
    
    def loesche_pngs_in_tmp_ordner(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ps = self.mb.createUnoService("com.sun.star.util.PathSettings")
            tempDir = uno.fileUrlToSystemPath(ps.Temp)
            temp_files = os.listdir(tempDir)

            ausdruck = r'\((.*),(.*),(.*)\)\((.*),(.*),(.*)\)\.png'
            
            zu_loeschende = []
            
            for t in temp_files:
                match_ob = re.search(ausdruck,t)
                if match_ob:
                    zu_loeschende.append(match_ob.string)
            
            for z in zu_loeschende:
                try:
                    pfad = os.path.join(tempDir,z)
                    os.remove(pfad)
                except:
                    pass
        except:
            log(inspect.stack,tb())
            
            
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:          
            if ev.ActionCommand == 'control_url_aussuchen':
                url = self.file_aussuchen()
                if url == None:
                    return
                else:
                    self.ctrls['control_url'].Model.Label = url
                    if self.ctrls['control_url_nutzer'].State == 1:
                        self.theme_ansicht_erneuern(url)
                return
            
            elif ev.ActionCommand == 'neues_persona':
                self.erstelle_neues_persona()
                return
            elif ev.ActionCommand == 'kein_persona':
                self.mb.class_Organon_Design.setze_persona('','')
                return
            elif ev.ActionCommand == 'loesche_persona':
                self.loesche_persona()
                return
            elif ev.ActionCommand == 'anwenden':
                self.persona_setzen()
                return
            
            elif ev.ActionCommand == 'control_solid':
                self.theme_ansicht_erneuern()
                
            elif ev.ActionCommand == 'control_gradient':
                url = self.erzeuge_personas(True)
                self.theme_ansicht_erneuern(url)
                
            elif ev.ActionCommand == 'control_url_nutzer':
                
                if self.ctrls['control_url'].Model.Label == LANG.URL:
                    url = self.file_aussuchen()
                    if url == None:
                        return
                    else:
                        self.ctrls['control_url'].Model.Label = url
                        self.theme_ansicht_erneuern(url)
                else:
                    url = self.ctrls['control_url'].Model.Label
                    self.theme_ansicht_erneuern(url)
                
            
            for c in 'control_solid','control_gradient','control_url_nutzer':
                self.ctrls[c].State = 0
            self.ctrls[ev.ActionCommand].State = 1
            
            # Listbox leere Ansicht
            items = self.ctrls['control_personas_list'].getItems()
            self.ctrls['control_personas_list'].removeItems(0,len(items))
            self.ctrls['control_personas_list'].addItems(items,0)
            
        except:
            log(inspect.stack,tb())
            
        
    def erstelle_neues_persona(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            name = self.ctrls['control_personas_name'].Text
            
            if name == '':
                Popup(self.mb, 'info').text = LANG.PERSONANAMEN_EINGEBEN
                return
            elif name in self.personas_dict:
                Popup(self.mb, 'info').text = LANG.PERSONANAMEN_EXISTIERT
                return
                
            url = self.ctrls['control_theme'].Model.ImageURL
            
            if url == '':
                url = self.erzeuge_personas()
                
            url2 = uno.fileUrlToSystemPath(url)
                        
            pers_path = os.path.join(self.gallery_user_path,'personas',name)
            try:
                os.mkdir(pers_path)
            except:
                Popup(self.mb, 'error').text = LANG.ORDNER_NICHT_ERSTELLT.format(pers_path)
                return
            
            pfad1 = os.path.join(pers_path,'header.png')
            pfad2 = os.path.join(pers_path,'footer.png')
            
            import shutil
            shutil.copy( url2 , pfad1 )
            shutil.copy( url2 , pfad2 )
            
            self.personas_dict.update({name:url2})
            self.personas_list.insert(0,name)
            self.ctrls['control_personas_list'].addItem(name,0)
            
        except:
            log(inspect.stack,tb())

    
    def persona_setzen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            thema = self.ctrls['control_personas_list'].SelectedItem
            
            if thema == '':
                Popup(self.mb, 'info').text = LANG.KEIN_PERSONA_SELEKTIERT
                return
            
            url = self.personas_dict[thema]
            dateiname = os.path.basename(url)
            farbe_schrift = self.ctrls['control_farbe_schrift'].Model.BackgroundColor
            
            v1 = thema + '/' + dateiname
            v2 = thema + '/' + dateiname
            v3 = '#' + self.mb.class_Funktionen.dezimal_to_hex(farbe_schrift)
            v4 = '#636363' # wird in LO noch nicht genutzt
            
            # Bsp: # value = 'neu/vordergrund.png;Dark Blue/Footer.png;#ffffff;#636363' 
            value = ';'.join([v1,v2,v3,v4])
            
            self.mb.class_Organon_Design.setze_persona(value)
            
        except:
            log(inspect.stack,tb())
    
    
    def loesche_persona(self):
        if self.mb.debug: log(inspect.stack)
        
        import shutil
        
        selected = self.ctrls['control_personas_list'].SelectedItem
        pos = self.ctrls['control_personas_list'].SelectedItemPos
        
        if selected == '':
            Popup(self.mb, 'info').text = LANG.KEIN_PERSONA_SELEKTIERT
            return
        else:
            entscheidung = self.mb.entscheidung(LANG.LOESCHEN_BESTAETIGEN.format(selected),"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return

        try:
            pfad = os.path.dirname(self.personas_dict[selected])
            shutil.rmtree(pfad)
        except:
            Popup(self.mb, 'warning').text = LANG.ORDNER_NICHT_ENTFERNT.format(pfad)
        
        del self.personas_dict[selected] 
        self.personas_list.remove(selected)
        self.ctrls['control_personas_list'].removeItems(pos,1)
        
        self.ctrls['control_solid'].State = 1
        self.theme_ansicht_erneuern()
        
        
    def file_aussuchen(self): 
        if self.mb.debug: log(inspect.stack)
        
        ofilter = ('Image','*.jpg;*.JPG;*.png;*.PNG;*.gif;*.GIF')
        filepath,ok = self.mb.class_Funktionen.filepicker2(ofilter=ofilter)
        
        if not ok:
            return None
        else:
            return filepath
         
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack) 
        
        try:

            for c in self.ctrls:
                if self.ctrls[c] == ev.Source:
                    cmd = c
                    break

            if cmd == 'control_farbe_schrift':
                farbe,rgb = self.get_farbe(cmd)
                self.ctrls['control_farbe_schrift'].Model.BackgroundColor = farbe
                self.theme_ansicht_erneuern(url='unveraendert')
                    
            elif cmd == 'control_farbe_hintergrund':
                farbe,rgb = self.get_farbe(cmd)
                self.ctrls['control_farbe_hintergrund'].Model.BackgroundColor = farbe
                
                if self.ctrls['control_solid'].State == 1:
                    self.theme_ansicht_erneuern()

                elif self.ctrls['control_gradient'].State == 1:
                    url = self.erzeuge_personas(gradient=True)
                    self.theme_ansicht_erneuern(url)
                
                    
            elif cmd == 'control_farbe_hintergrund2':
                farbe,rgb = self.get_farbe(cmd)
                self.ctrls['control_farbe_hintergrund2'].Model.BackgroundColor = farbe
                if self.ctrls['control_gradient'].State == 1:
                    url = self.erzeuge_personas(gradient=True)
                    self.theme_ansicht_erneuern(url)
      
        except:
            log(inspect.stack,tb())
            
            
    def get_farbe(self,cmd):
        if self.mb.debug: log(inspect.stack) 
        
        ctrl = self.ctrls[cmd]
        urspr_farbe = ctrl.Model.BackgroundColor
        
        farbe = self.mb.class_Funktionen.waehle_farbe(initial_value=urspr_farbe)

        return farbe,self.mb.class_Funktionen.dezimal_to_rgb(farbe) 
             
    def mouseExited(self, ev):
        return False
    def mouseEntered(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    
    def erzeuge_png(self,data,height,width):
        '''
        based on code from Guido Draheim, see:
        http://stackoverflow.com/questions/8554282/creating-a-png-file-in-python
        '''
        if self.mb.debug: log(inspect.stack) 
        
        import struct
        import zlib
        
        def I1(value):
            return struct.pack("!B", value & (2**8-1))
        
        def I1_2(value):
                return struct.pack("!B", value & (2**8-1))
        
        def I4(value):
            return struct.pack("!I", value & (2**32-1))
        
        # generate these chunks depending on image type
        makeIHDR = True
        makeIDAT = True
        makeIEND = True
        png = b"\x89" + "PNG\r\n\x1A\n".encode('ascii')
        
        if makeIHDR:
            
            colortype = 2 # rgb color image, no alpha (no palette)
            bitdepth = 8 # with one byte per pixel (0..255)
            compression = 0 # zlib (no choice here)
            filtertype = 0 # adaptive (each scanline seperately)
            interlaced = 0 # no
            
            IHDR = I4(width) + I4(height) + I1(bitdepth)
            IHDR += I1(colortype) + I1(compression)
            IHDR += I1(filtertype) + I1(interlaced)
            block = "IHDR".encode('ascii') + IHDR
            
            png += I4(len(IHDR)) + block + I4(zlib.crc32(block))
        
        if makeIDAT:
            
            raw = []
            
            for y in range(height):
                raw.append(b"\0") # no filter for this scanline
                for x in range(width):
                    for i in range(3):
                        c = I1( data[y*width + x ][i])
                        raw.append(c)
            
            raw = b''.join(raw)

            compressor = zlib.compressobj()
            compressed = compressor.compress(raw)
            compressed += compressor.flush() #!!
            block = "IDAT".encode('ascii') + compressed
            png += I4(len(compressed)) + block + I4(zlib.crc32(block))
            
        if makeIEND:
            block = "IEND".encode('ascii')
            png += I4(0) + block + I4(zlib.crc32(block))
            
        return png
                

    def erzeuge_personas(self,gradient=False,url=''):         
        if self.mb.debug: log(inspect.stack) 
                
        start = self.mb.class_Funktionen.dezimal_to_rgb(self.ctrls['control_farbe_hintergrund'].Model.BackgroundColor)
        end = self.mb.class_Funktionen.dezimal_to_rgb(self.ctrls['control_farbe_hintergrund2'].Model.BackgroundColor)
        
        ps = self.mb.createUnoService("com.sun.star.util.PathSettings")
        tempDir = uno.fileUrlToSystemPath(ps.Temp)
        
        temp_png = str(start)+str(end)+'.png'
        
        # pruefen, ob png bereits im temp Ordner existiert
        temp_files = os.listdir(tempDir)
        if temp_png in temp_files:
            pfad = uno.systemPathToFileUrl(os.path.join(tempDir,temp_png))
            self.ctrls['control_theme'].Model.ImageURL = pfad
            return pfad
            
        
        height = 200
        width = 2500
        RANGE = height*width
        
        diff0 = end[0] - start[0] 
        diff1 = end[1] - start[1]
        diff2 = end[2] - start[2]
        
        d0 = diff0/RANGE
        d1 = diff1/RANGE
        d2 = diff2/RANGE
        
        r,g,b = start
        
        if url == '':
            pfad = os.path.join(tempDir,str(start)+str(end)+'.png')

        if gradient:
            StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
            StatusIndicator.start(LANG.ERZEUGE_GRADIENT,2)
            StatusIndicator.setValue(2)
            colors = [ [int(r + (x+1)*d0), int(g + (x+1)*d1), int(b + (x+1)*d2) ] for x in range(RANGE)] 
        else:        
            colors = [ start for x in range(RANGE)] 
        
        png = self.erzeuge_png( colors ,height,width)
        
        
        with open(pfad,"wb") as f: 
            f.write(png)  
             
        self.ctrls['control_theme'].Model.ImageURL = pfad   
        
        if gradient:
            StatusIndicator.end()
        
        return uno.systemPathToFileUrl(pfad)    

   
    
 
    
    
    
                