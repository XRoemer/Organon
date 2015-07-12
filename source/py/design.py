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
        
        #self.mb.nachricht(LANG.AENDERUNG_NACH_NEUSTART,"infobox")
        
        
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
                        
        controls = [
            10,
            ('control1',"FixedText",         
                                    'tab0',0,168,20,  
                                    ('Label','FontWeight'),
                                    (LANG.ORGANON_DESIGN ,150),                                             
                                    {} 
                                    ),
            0,
            ('control_Writer_Design',"CheckBox",        
                                    'tab3',3,100,20,  
                                    ('Label','FontWeight','State'),
                                    (LANG.WRITER_DESIGN,150,self.mb.settings_orga['organon_farben']['design_office']),                                                                    
                                    {'setActionCommand':'writer_design','addActionListener':(listener,)} 
                                    ),
            0,
            ('control_wDesign_bearbeiten',"Button",          
                                    'tab4',-2,100,23,    
                                    ('Label',),
                                    (LANG.BEARBEITEN,),                                                                         
                                    {'setActionCommand':'writer_design_bearbeiten','addActionListener':(listener,)} 
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
                                    (LANG.HINTERGRUND,),                                                                    
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
                                    (LANG.SCHRIFT,),                                                                    
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

            # UEBERGABE AN LISTENER
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
            
            
    def dialog_persona_elemente(self,listener,persona_items):
        if self.mb.debug: log(inspect.stack)
        
        controls = [
            1,
            ('control_theme',"ImageControl",         
                                    'tab0',0,self.container.Size.Width-6,80,  
                                    ('Label','FontWeight','BackgroundColor','Border'),
                                    (LANG.ORGANON_DESIGN ,150,KONST.FARBE_MENU_HINTERGRUND,0),                                             
                                    {} 
                                    ),
            20,
            ('control_theme_tit',"FixedText",        
                                    'tab_theme',0,100,30,  
                                    ('Label','FontWeight','TextColor'),
                                    (LANG.THEMA_TITEL,150,KONST.FARBE_MENU_SCHRIFT),                                                                    
                                    {} 
                                    ),  
            30,
            ('control_theme_untertit',"FixedText",        
                                    'tab_theme2',0,200,30,  
                                    ('Label','TextColor'),
                                    (LANG.THEMA_UNTERTITEL,KONST.FARBE_MENU_SCHRIFT),                                                                    
                                    {} 
                                    ),  
            ############################ PERSONAS ######################################
            'Y=120',
            ('control_VP',"FixedText",        
                                    'tab1',0,150,20,  
                                    ('Label','FontWeight'),
                                    (LANG.VORHANDENE_PERSONAS,150),                                                                    
                                    {} 
                                    ),  
            30,
            ('control_personas_list',"ListBox",        
                                    'tab1',0,120,20,  
                                    ('Border','Dropdown','LineCount'),
                                    ( 2,True,10),       
                                    {'addItems':persona_items,'addItemListener':listener}
                                    ), 
             
                    
            ############################# EIGENES PERSONAS #####################################
            'Y=120',
            ('control_farbe_schrift',"FixedText",        
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_SCHRIFT,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control5',"FixedText",        
                                    'tab5',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE_SCHRIFT,),                                                                    
                                    {} 
                                    ),  
            25,
            ('control_solid',"RadioButton",        
                                    'tab3',0,150,25,  
                                    ('Label','State'),
                                    (LANG.HINTERGRUND_EINFARBIG,1),                                                                    
                                    {'setActionCommand':'control_solid','addActionListener':(listener,)}
                                    ), 
            25,
            ('control_farbe_hintergrund',"FixedText",        
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),  
            0,
            ('control3',"FixedText",        
                                    'tab5',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE,),                                                                    
                                    {} 
                                    ),  
                    
            25,
            ('control_gradient',"RadioButton",        
                                    'tab3',0,150,25,  
                                    ('Label',),
                                    (LANG.HINTERGRUND_GRADIENT,),                                                                    
                                    {'setActionCommand':'control_gradient','addActionListener':(listener,)}
                                    ), 
            25,
            ('control_farbe_hintergrund2',"FixedText",        
                                    'tab4',0,32,16,  
                                    ('BackgroundColor','Label','Border'),
                                    (KONST.FARBE_MENU_HINTERGRUND,'    ',1),       
                                    {'addMouseListener':(listener)} 
                                    ),
            0,
            ('control_f2text',"FixedText",        
                                    'tab5',0,100,20,  
                                    ('Label',),
                                    (LANG.FARBE2,),                                                                    
                                    {} 
                                    ),
            25,
            ('control_url_nutzer',"RadioButton",        
                                    'tab3',0,150,25,  
                                    ('Label',),
                                    (LANG.HINTERGRUND_NUTZER,),                                                                    
                                    {'setActionCommand':'control_url_nutzer','addActionListener':(listener,)}
                                    ), 
              
                    
            20,
            ('control_url_aussuchen',"Button",        
                                    'tab4',0,80,20,  
                                    ('Label',),
                                    (LANG.AUSWAHL,),                                                                    
                                    {'setActionCommand':'control_url_aussuchen','addActionListener':(listener,)}
                                    ), 
            30,
            ('control_url',"FixedText",          
                                    'tab4',0,300,20,    
                                    ('Label',),
                                    (LANG.URL,),                                                                         
                                    {} 
                                    ), 
            20,
            ('control_url_info',"FixedText",          
                                    'tab4',0,300,20,    
                                    ('Label',),
                                    (LANG.URL_INFO,),                                                                         
                                    {} 
                                    ), 
    
            'Y=120',
            ###############################################################
            ######################## SPEICHERN ############################
            ###############################################################
 
            
            ('control_personas_name',"Edit",          
                                    'tab7',0,100,20,    
                                    ('HelpText',),
                                    (LANG.PERSONANAMEN_EINGEBEN,),                                                                         
                                    {} 
                                    ), 
            30,
            ('control_neues_thema',"Button",          
                                    'tab7',0,100,23,    
                                    ('Label',),
                                    (LANG.NEUES_PERSONA_THEMA,),                                                                         
                                    {'setActionCommand':'neues_persona','addActionListener':(listener,)} 
                                    ),
            60,
            ('control_thema_loeschen',"Button",          
                                    'tab7',0,100,23,    
                                    ('Label',),
                                    (LANG.PERSONA_LOESCHEN,),                                                                         
                                    {'setActionCommand':'loesche_persona','addActionListener':(listener,)} 
                                    ), 
            150,
            ('control_anwenden',"Button",          
                                    'tab6',0,150,23,    
                                    ('Label',),
                                    (LANG.PERSONA_ANWENDEN,),                                                                         
                                    {'setActionCommand':'anwenden','addActionListener':(listener,)} 
                                    ), 
            ]

        # Tabs waren urspruenglich gesetzt, um sie in der Klasse Design richtig anzupassen.
        # Das fehlt. 'tab...' wird jetzt nur in Zahlen uebersetzt. Beim naechsten 
        # groesseren Fenster, das ich schreibe und das nachtraegliche Berechnungen benoetigt,
        # sollte der Code generalisiert werden. Vielleicht grundsaetzlich ein Modul fenster.py erstellen?
        tab0_0 = 1
        tab_theme = self.container.Size.Width/2 - 30
        tab_theme2 = self.container.Size.Width/2 - 100
        tab0 = 0
        tab1 = 20
        tab2 = 35
        tab3 = self.container.Size.Width/2 - 47
        tab4 = self.container.Size.Width/2 - 30
        tab5 = self.container.Size.Width/2 + 10
        tab6 = 310
        tab7 = 360
        
        tabs = ['tab_theme','tab_theme2','tab0_0','tab0','tab1','tab2','tab3','tab4','tab5','tab6','tab7']
        
        tabs_dict = {}
        for a in tabs:
            tabs_dict.update({a:locals()[a]  })

        controls2 = []
        for c in controls:
            if not isinstance(c, int) and 'Y=' not in c:
                c2 = list(c)
                c2[2] = tabs_dict[c2[2]]
                controls2.append(c2)
            else:
                controls2.append(c)
         
        return controls2
    
              
    def dialog_persona(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.kopiere_personas()
            
            # vorhandene personas auslesen
            personas_dict,personas_list,gallery_user_path = self.mb.class_Organon_Design.get_personas()        
            listener = Listener_Persona(self.mb,personas_dict,personas_list,gallery_user_path)
            
            # controls erzeugen
            controls = self.dialog_persona_elemente(listener,tuple(personas_list))
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)
            
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
                
            elif cmd == 'writer_design_bearbeiten':
                loc = ev.Source.Context.Context.AccessibleContext.AccessibleParent.PosSize
                self.writer_design_bearbeiten(loc)
                
            elif cmd == 'writer_design':
                self.nutze_writer_design(ev.Source)
                
                
            else:
                self.setze_design(cmd,ev)
        except:
            log(inspect.stack,tb()) 
            
            
    def writer_design_bearbeiten(self,loc):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.class_Design.setze_listboxen = True
            
            breite = 930
            hoehe = 400

            sett = self.mb.settings_orga
            
            ctrls,pos_y = self.dialog_writer_design()

            x,y = loc.X + loc.Width + 20,loc.Y

               
            # Hauptfenster erzeugen
            posSize = x,y,breite,pos_y + 40
            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
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
                aktiv = sett['organon_farben']['aktiv']
                persona_url = sett['organon_farben']['office']['persona_url']
                persona = persona_url.split('/')[0]
    
                ctrl_persona = ctrls['control_personas_list']
                if persona in ctrl_persona.Items:
                    ctrl_persona.selectItem(persona,True)
            
            self.mb.class_Design.setze_listboxen = False
            
        except:
            log(inspect.stack,tb()) 
            

    def dialog_writer_design(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Listener_Writer_Design(self.mb)

            controls = self.dialog_writer_design_elemente(listener)
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)
                
            listener.ctrls = ctrls

            return ctrls, pos_y 
  
        except:
            log(inspect.stack,tb())
            # bei einer Exception scheitert die weitere Anzeige.
            # Es fehlt ein Fallback
            return [], 0 


    def dialog_writer_design_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        from collections import OrderedDict
        fixed_line_length = 250
                
        controls = [
            10,
            ('control1',"FixedText",         
                                    'tab0',0,168,20,  
                                    ('Label','FontWeight'),
                                    (LANG.WRITER_DESIGN ,150),                                             
                                    {} 
                                    ),
            10,
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
                    LANG.TRENNER     + ' ' + LANG.SCHRIFT
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
                        )
        
        self.lb_dict = {lb_items_syn[i] : lb_items[i] for i in range(len(lb_items))}
        lb_dict2 = {lb_items[i] : lb_items_syn[i]  for i in range(len(lb_items))}
        self.mb.class_Design.lb_dict = self.lb_dict
        self.mb.class_Design.lb_dict2 = lb_dict2
        
        
        def erzeuge_conts(cts,tabs,abstand_nach_sep=0):
            
            controls = []
            
            for c in cts:
                if cts[c] == []: continue
                
                elif 'sep' in cts[c][0]:
                    cont = [
                            ('control_{}'.format(c),"FixedLine",         
                                    tabs[0],20,fixed_line_length,1,   
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
                            ('control_{}'.format(c),"FixedText",        
                                        tabs[0],0,100,20,  
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
                            ('control_{}'.format(c),"FixedText",        
                                        tabs[0],0,32,16,  
                                        ('BackgroundColor','Label','Border'),
                                        (farbe,'    ',1),       
                                        {'addMouseListener':(listener)} 
                                        ), 
                            0,                                                  
                            ('control_{}LB'.format(c),"ListBox",      
                                        tabs[1],0,60,15,    
                                        ('Border','Dropdown','LineCount'),
                                        ( 2,True,15),       
                                        {'addItems':lb_items, 'addItemListener':listener} 
                                        ), 
                            0,
                            ('control_{}L'.format(c),"FixedText",        
                                        tabs[2],0,170,20,  
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
        conts[0] = 'Y=50'
        controls.extend(conts)


        # PERSONAS
        
        if LO:
            personas_dict,personas_list,gallery_user_path = self.mb.class_Organon_Design.get_personas() 
            
            cts = [
                   'Y=50',
                    ('control_VP',"FixedText",        
                                            'tab6',0,100,20,  
                                            ('Label','FontWeight'),
                                            (LANG.PERSONA,150),                                                                    
                                            {} 
                                            ), 
                   0,
                   ('control_use_persona',"CheckBox",        
                                            'tab8',0,120,20,  
                                            ('Label','State'),
                                            ( LANG.NUTZE_PERSONA,sett['nutze_personas']),       
                                            {'addActionListener':(listener,)}
                                            ), 
                    20,
                    ('control_personas_list',"ListBox",        
                                            'tab6',0,120,20,  
                                            ('Border','Dropdown','LineCount'),
                                            ( 2,True,10),       
                                            {'addItems':tuple(personas_list),'addItemListener':listener}
                                            ), 
                   
                   30,
                   ('controlFix2',"FixedLine",         
                                    'tab6',20,fixed_line_length,1,   
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
            Y = 40
        else:
            Y = 'Y=50'
            
        cts = [
               Y,
                ('control_DB',"FixedText",        
                                        'tab6',0,100,20,  
                                        ('Label','FontWeight'),
                                        (LANG.DOCUMENT,150),                                                                    
                                        {} 
                                        ), 
               0,
               ('control_nutze_dok_farbe',"CheckBox",        
                                        'tab8',0,120,20,  
                                        ('Label','State'),
                                        ( LANG.NUTZE_DOK_FARBE,sett['nutze_dok_farbe']),       
                                        {'addActionListener':(listener,)}
                                        ), 
                0,
                ]
        
        controls.extend(cts)
                
        cts4 = OrderedDict()
        cts4['dok_hintergrund'] = [LANG.HINTERGRUND,sett['dok_hintergrund']]
        cts4['dok_farbe'] = [LANG.DOCUMENT,sett['dok_farbe']]
                
        conts = erzeuge_conts(cts4,['tab6','tab7','tab8'])
        controls.extend(conts)      
        
        
        tab0 = 20
        tab1 = 80
        tab2 = 150
        tab3 = 320
        tab4 = 380
        tab5 = 450
        tab6 = 660
        tab7 = 720
        tab8 = 790
        
        tabs = ['tab0','tab1','tab2','tab3','tab4','tab5','tab6','tab7','tab8']
        
        tabs_dict = {}
        for a in tabs:
            tabs_dict.update({a:locals()[a]  })

        controls2 = []
        for c in controls:
            if not isinstance(c, int) and 'Y=' not in c:
                c2 = list(c)
                c2[2] = tabs_dict[c2[2]]
                controls2.append(c2)
            else:
                controls2.append(c)

        return controls2
    
    
    def nutze_writer_design(self,src):
        if self.mb.debug: log(inspect.stack)

        if src.State:
            self.mb.nachricht(LANG.WRITER_DESIGN_INFO,"infobox")
            if self.mb.programm == 'LibreOffice':
                self.mb.class_Organon_Design.kopiere_personas()
            

        self.mb.settings_orga['organon_farben']['design_office'] = src.State
        self.mb.class_Funktionen.schreibe_settings_orga()

    
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
        
            
    def setze_design(self,cmd,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            # Fuer den Tabhintergrund
            # self.mb.tabsX.Window.setBackground(KONST.FARBE_HF_HINTERGRUND)
            
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
                pass
                #print(e)
                
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
                pass
                #print(e)
                
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


class Listener_Writer_Design(unohelper.Base,XItemListener,XMouseListener,XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb  
        self.ctrls = {}

        
    def itemStateChanged(self, ev): 
        # Wenn die LBs nicht vom user gesetzt werden, sondern
        # von Organon, keine Aktion ausfuehren 
        if self.mb.class_Design.setze_listboxen == True: return
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
                farb_prop = self.mb.class_Design.lb_dict2[selektiert]
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

                self.mb.class_Funktionen.schreibe_settings_orga()
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

        self.mb.class_Funktionen.schreibe_settings_orga()
        
                    
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

                self.mb.class_Funktionen.schreibe_settings_orga()

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
            
        self.mb.class_Funktionen.schreibe_settings_orga()
                
        
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
            
            tab0_0 = 1
            tab_theme = self.ctrl_container.Size.Width/2 - 30
            tab_theme2 = self.ctrl_container.Size.Width/2 - 100
            
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
            
        self.ctrls['control_theme'].dispose()
        
        bg_color = self.ctrls['control_farbe_hintergrund'].Model.BackgroundColor
        
        controls = [
                1,
                ('control_theme',"ImageControl",         
                    1,0,self.ctrl_container.Size.Width-6,80,  
                    ('Label','FontWeight','ImageURL','Border','BackgroundColor'),
                    (LANG.ORGANON_DESIGN ,150,url,0,bg_color),                                             
                    {} 
                    ),
                ]

        ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)
        
        self.ctrl_container.addControl('control_theme',ctrls['control_theme'])
        self.ctrls.update({'control_theme': ctrls['control_theme']})
        self.ctrls['control_theme'].draw(1,1)
         
        for c in self.ctrls['control_theme_tit'],self.ctrls['control_theme_untertit']:
            x,y = c.PosSize.X,c.PosSize.Y
            c.Model.TextColor = self.ctrls['control_farbe_schrift'].Model.BackgroundColor
            c.draw(x,y+1)
    
            
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
                self.mb.nachricht(LANG.PERSONANAMEN_EINGEBEN,"infobox")
                return
            elif name in self.personas_dict:
                self.mb.nachricht(LANG.PERSONANAMEN_EXISTIERT,"infobox")
                return
                
            url = self.ctrls['control_theme'].Model.ImageURL
            
            if url == '':
                url = self.erzeuge_personas()
                
            url2 = uno.fileUrlToSystemPath(url)
                        
            pers_path = os.path.join(self.gallery_user_path,'personas',name)
            try:
                os.mkdir(pers_path)
            except:
                self.mb.nachricht(LANG.ORDNER_NICHT_ERSTELLT.format(pers_path),"warningbox")
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
                self.mb.nachricht(LANG.KEIN_PERSONA_SELEKTIERT,"infobox")
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
            self.mb.nachricht(LANG.KEIN_PERSONA_SELEKTIERT,"infobox")
            return
        else:
            entscheidung = self.mb.nachricht(LANG.LOESCHEN_BESTAETIGEN.format(selected),"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return

        try:
            pfad = os.path.dirname(self.personas_dict[selected])
            shutil.rmtree(pfad)
        except:
            self.mb.nachricht(LANG.ORDNER_NICHT_ENTFERNT.format(pfad),"warningbox")
        
        del self.personas_dict[selected] 
        self.personas_list.remove(selected)
        self.ctrls['control_personas_list'].removeItems(pos,1)
        
        self.ctrls['control_solid'].State = 1
        self.theme_ansicht_erneuern()
        
        
    def file_aussuchen(self): 
        if self.mb.debug: log(inspect.stack)
          
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        Filepicker.appendFilter('Image','*.jpg;*.JPG;*.png;*.PNG;*.gif;*.GIF')
        Filepicker.execute()
        
        if Filepicker.Files == '':
            return None
        
        return Filepicker.Files[0]
        
         
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack) 
        
        try:

            for c in self.ctrls:
                if self.ctrls[c] == ev.Source:
                    cmd = c
                    break

            if cmd == 'control_farbe_schrift':
                farbe,rgb = self.get_farbe()
                self.ctrls['control_farbe_schrift'].Model.BackgroundColor = farbe
                self.theme_ansicht_erneuern(url='unveraendert')
                    
            elif cmd == 'control_farbe_hintergrund':
                farbe,rgb = self.get_farbe()
                self.ctrls['control_farbe_hintergrund'].Model.BackgroundColor = farbe
                
                if self.ctrls['control_solid'].State == 1:
                    self.theme_ansicht_erneuern()

                elif self.ctrls['control_gradient'].State == 1:
                    url = self.erzeuge_personas(gradient=True)
                    self.theme_ansicht_erneuern(url)
                
                    
            elif cmd == 'control_farbe_hintergrund2':
                farbe,rgb = self.get_farbe()
                self.ctrls['control_farbe_hintergrund2'].Model.BackgroundColor = farbe
                if self.ctrls['control_gradient'].State == 1:
                    url = self.erzeuge_personas(gradient=True)
                    self.theme_ansicht_erneuern(url)
      
        except:
            log(inspect.stack,tb())
            
            
    def get_farbe(self):
        if self.mb.debug: log(inspect.stack) 
                  
        cp = self.mb.ctx.ServiceManager.createInstanceWithContext("com.sun.star.cui.ColorPicker",self.mb.ctx)
        cp.execute()
        cp.dispose()
        
        farbe = cp.PropertyValues[0].Value
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

   
    
 
    
    
    
                