# -*- coding: utf-8 -*-
import uno
import unohelper
from traceback import print_exc as tb
import sys
import os
import xml.etree.ElementTree as ElementTree
import time
from codecs import open as codecs_open
from math import floor as math_floor
import re
import konstanten as KONST
import copy
import inspect

platform = sys.platform



class Menu_Bar():
    
    def __init__(self,args,tab = 'Projekt'):
        
        pdk,dialog,ctx,tabsX,path_to_extension,win,dict_sb,debugX,factory,menu_start = args
        
        global debug
        debug = debugX
        if debug:
            self.get_pyPath()
        
        self.win = win
        self.pd = pdk
        global pd,IMPORTS
        pd = pdk
        
        global T,log
        T = Tab()
        log = Log(self).log
        
        IMPORTS = ('uno','unohelper','sys','os','ElementTree','time',
                   'codecs_open','math_floor','re','tb','platform','KONST',
                   'pd','copy','Props','T','log','inspect')
        
        if 'LibreOffice' in sys.executable:
            self.programm = 'LibreOffice'
        elif 'OpenOffice' in sys.executable:
            self.programm = 'OpenOffice'
        else:
            # Fuer Linux / OSX fehlt
            self.programm = 'LibreOffice'
        
        # Konstanten
        self.factory = factory
        self.dialog = dialog
        self.ctx = ctx
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
        self.doc = self.get_doc()        
        self.current_Contr = self.doc.CurrentController 
        self.undo_mgr = self.doc.UndoManager
        self.viewcursor = self.current_Contr.ViewCursor
        self.tabsX = tabsX
        self.platform = sys.platform
        self.language = None
        self.lang = self.lade_Modul_Language()
        self.path_to_extension = path_to_extension
        self.programm_version = self.get_programm_version()
        self.filters_import = None
        self.filters_export = None
        self.BEREICH_EINFUEGEN = self.get_BEREICH_EINFUEGEN()
        self.anleitung_geladen = False
        self.speicherort_last_proj = self.get_speicherort()
        self.projekt_name = None
        self.menu_start = menu_start
        self.sec_helfer = None

        
        # Properties
        self.props = {}
        self.props.update({T.AB :Props()})
        self.dict_sb = dict_sb              # drei Unterdicts: sichtbare, eintraege, controls
        self.dict_sb_content = None
        
        self.tabs = {tabsX.ActiveTabID:(dialog,'Projekt')}
        self.active_tab_id = tabsX.ActiveTabID
        self.tab_id_old = self.active_tab_id
        self.tab_umgeschaltet = False
        self.bereich_wurde_bearbeitet = False
        
        
        # Settings
        self.settings_exp = None
        self.settings_imp = None
        self.settings_proj = {}
        self.user_styles = ()

        # Pfade
        self.pfade = {}

        # Klassen   
        self.Key_Handler = Key_Handler(self)
        self.ET = ElementTree  
        self.Mitteilungen = Mitteilungen(self.ctx,self)
         
        self.class_Hauptfeld,self.class_Zeilen_Listener = self.get_Klasse_Hauptfeld()
        self.class_Projekt =    self.lade_modul('projects','.Projekt(self, pd)')   
        self.class_XML =        self.lade_modul('xml_m','.XML_Methoden(self,pd)')
        self.class_Funktionen = self.lade_modul('funktionen','.Funktionen(self, pd)')     
        self.class_Export =     self.lade_modul('export','.Export(self,pd)')
        self.class_Import =     self.lade_modul('importX','.ImportX(self,pd)') 
        self.class_Sidebar =    self.lade_modul('sidebar','.Sidebar(self,pd)') 
        self.class_Bereiche =   self.lade_modul('bereiche','.Bereiche(self)')
        self.class_Version =    self.lade_modul('version','.Version(self,pd)') 
        self.class_Tabs =       self.lade_modul('tabs','.Tabs(self)') 
        
        
        # Listener  
        self.VC_selection_listener  = ViewCursor_Selection_Listener(self)          
        self.w_listener             = Dialog_Window_Listener(self)    
        self.undo_mgr_listener      = Undo_Manager_Listener(self)
        self.tab_listener           = Tab_Listener(self)
        
        self.Listener = {}
        self.Listener.update({'Menu_Kopf_Listener':Menu_Kopf_Listener(self)})
        self.Listener.update({'Menu_Kopf_Listener2':Menu_Kopf_Listener2(self)})
        
        self.undo_mgr.addUndoManagerListener(self.undo_mgr_listener)
        self.dialog.addWindowListener(self.w_listener)
        self.tabsX.addTabListener(self.tab_listener)

        self.use_UM_Listener = False
        
        # fuers debugging
        self.debug = debug
        if self.debug: 
            print('Debug = True')
            self.time = time
            self.timer_start = self.time.clock()
            
            


#         UD_properties = self.doc.DocumentProperties.UserDefinedProperties


           
       
    def get_doc(self):
        
        enum = self.desktop.Components.createEnumeration()
        comps = []
        
        while enum.hasMoreElements():
            comps.append(enum.nextElement())
            
        # Wenn ein neues Dokument geoeffnet wird, gibt es bei der Initialisierung
        # noch kein Fenster, aber die Komponente wird schon aufgefuehrt.
        # Hat die zuletzt erzeugte Komponente comps[0] kein ViewData,
        # dann wurde sie neu geoeffnet.
        if comps[0].ViewData == None:
            doc = comps[0]
        else:
            doc = self.desktop.getCurrentComponent() 
            
        return doc
    
    def get_BEREICH_EINFUEGEN(self):
        UM = self.doc.UndoManager
        newSection = self.doc.createInstance("com.sun.star.text.TextSection")
        cur = self.doc.Text.createTextCursor()  
        cur.gotoEnd(False)
        self.doc.Text.insertTextContent(cur, newSection, False)
        BEREICH_EINFUEGEN = UM.getCurrentUndoActionTitle()
        cur.gotoRange(newSection.Anchor,True)
        newSection.dispose()
        cur.setString('')
        return BEREICH_EINFUEGEN
    
    def get_programm_version(self):
        pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        for ext in pip.ExtensionList:
            if ext[0] == 'xaver.roemers.organon':
                version = ext[1]
        return version
    
    def erzeuge_Menu(self,win):
        try:             
            listener = Menu_Kopf_Listener(self) 
            listener2 = Menu_Kopf_Listener2(self) 
            self.erzeuge_MenuBar_Container(win)
            
            self.erzeuge_Menu_Kopf_Datei(listener,win)
            self.erzeuge_Menu_Kopf_Bearbeiten(listener,win)
            self.erzeuge_Menu_Kopf_Optionen(listener,win)
            
            if debug:
                self.erzeuge_Menu_Kopf_Test(listener,win)
            if T.AB == 'Projekt':
                self.erzeuge_Menu_neuer_Ordner(listener2,win)
                self.erzeuge_Menu_Kopf_neues_Dokument(listener2,win)
            self.erzeuge_Menu_Kopf_Papierkorb_leeren(listener2,win)
            
        except Exception as e:
                self.Mitteilungen.nachricht('erzeuge_Menu ' + str(e),"warningbox")
                tb()

    
    def erzeuge_MenuBar_Container(self,win):
        menuB_control, menuB_model = self.createControl(self.ctx, "Container", 2, 2, 1000, 20, (), ())          
        menuB_model.BackgroundColor = KONST.Color_MenuBar_Container
         
        win.addControl('Organon_Menu_Bar', menuB_control)


    def erzeuge_Menu_Kopf_Datei(self,listener,win):
        control, model = self.createControl(self.ctx, "FixedText", 0, 2, 35, 20, (), ())           
        model.Label = self.lang.FILE  
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar') 
        MenuBarCont.addControl('Datei', control)
        
        
    def erzeuge_Menu_Kopf_Bearbeiten(self,listener,win):
        control, model = self.createControl(self.ctx, "FixedText", 37, 2, 60, 20, (), ())           
        model.Label = self.lang.BEARBEITEN_M             
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar') 
        MenuBarCont.addControl('Bearbeiten', control)
    
    
    def erzeuge_Menu_Kopf_Optionen(self,listener,win):         
        control, model = self.createControl(self.ctx, "FixedText", 100, 2, 55, 20, (), ())           
        model.Label = self.lang.OPTIONS           
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')      
        MenuBarCont.addControl('Optionen', control)
        
        
    def erzeuge_Menu_Kopf_Test(self,listener,win):
        control, model = self.createControl(self.ctx, "FixedText", 300, 2, 50, 20, (), ())           
        model.Label = 'Test'                     
        control.addMouseListener(listener) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')   
        MenuBarCont.addControl('Projekt', control)
        
        
    def erzeuge_Menu_neuer_Ordner(self,listener2,win):
        control, model = self.createControl(self.ctx, "ImageControl", 170, 0, 20, 20, (), ())   
        model.ImageURL = KONST.IMG_ORDNER_NEU_24
        
        model.HelpText = self.lang.INSERT_DIR
        model.Border = 0                    
        control.addMouseListener(listener2) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')     
        MenuBarCont.addControl('Ordner', control)
        
        
    def erzeuge_Menu_Kopf_neues_Dokument(self,listener2,win):
        control, model = self.createControl(self.ctx, "ImageControl", 190, 0, 20, 20, (), ())           
        model.ImageURL = KONST.IMG_DATEI_NEU_24
                    
        model.HelpText = self.lang.INSERT_DOC
        model.Border = 0                    
        control.addMouseListener(listener2) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')   
        MenuBarCont.addControl('neues_Dokument', control)
  
        
    def erzeuge_Menu_Kopf_Papierkorb_leeren(self,listener2,win):
        control, model = self.createControl(self.ctx, "ImageControl", 240, 0, 20, 20, (), ())           
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leeren.png'
        model.HelpText = self.lang.CLEAR_RECYCLE_BIN
        model.Border = 0                       
        control.addMouseListener(listener2) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')     
        MenuBarCont.addControl('Papierkorb_leeren', control)


    def erzeuge_Menu_DropDown_Container(self,ev,BREITE = KONST.Breite_Menu_DropDown_Container, HOEHE = KONST.Hoehe_Menu_DropDown_Container):
        smgr = self.smgr
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)    
        oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", self.ctx)
        
        # Create Uno Struct
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
        oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
        # global oWindow
        oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
        oWindowDesc.WindowServiceName = ""
        oWindowDesc.Parent = toolkit.getDesktopWindow()
        oWindowDesc.ParentIndex = -1
        oWindowDesc.WindowAttributes = 32
        
        # Set Window Attributes
        gnDefaultWindowAttributes = 1   # + 16 + 32 + 64 + 128 
    #         com_sun_star_awt_WindowAttribute_SHOW + \
    #         com_sun_star_awt_WindowAttribute_BORDER + \
    #         com_sun_star_awt_WindowAttribute_MOVEABLE + \
    #         com_sun_star_awt_WindowAttribute_CLOSEABLE + \
    #         com_sun_star_awt_WindowAttribute_SIZEABLE
          
        X = ev.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.Source.AccessibleContext.LocationOnScreen.value.Y + ev.Source.Size.value.Height
    
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = BREITE
        oRect.Height = HOEHE
        
        oWindowDesc.Bounds = oRect
        #pd()
        # specify the window attributes.
        oWindowDesc.WindowAttributes = gnDefaultWindowAttributes
        # create window
        oWindow = toolkit.createWindow(oWindowDesc)
         
        # create frame for window
        oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", self.ctx)
        oFrame.initialize(oWindow)
        oFrame.setCreator(self.desktop)
        oFrame.activate()
        
        # create new control container
        cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", self.ctx)
        cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", self.ctx)
        cont_model.BackgroundColor = KONST.MENU_DIALOG_FARBE  
        cont.setModel(cont_model)
        # need createPeer just only the container
        cont.createPeer(toolkit, oWindow)
        cont.setPosSize(0, 0, 0, 0, 15)
       
        oFrame.setComponent(cont, None)
        
        # create Listener
        listener = DropDown_Container_Listener(self)
        cont.addMouseListener(listener) 
        listener.ob = oWindow
        Name = ev.value.Source.AccessibleContext.AccessibleName

        self.menu_fenster = oWindow

        if Name == self.lang.FILE:
            self.erzeuge_Menu_DropDown_Eintraege_Datei(oWindow, cont)
        if Name == self.lang.BEARBEITEN_M:
            self.erzeuge_Menu_DropDown_Eintraege_Bearbeiten(oWindow, cont)
        if Name == self.lang.OPTIONS:
            self.erzeuge_Menu_DropDown_Eintraege_Optionen(oWindow, cont)
        return Name
  
    
    def erzeuge_Menu_DropDown_Eintraege_Datei(self,window,cont):
        lang = self.lang
        control, model = self.createControl(self.ctx, "ListBox", 10 ,  10 , 
                                        KONST.Breite_Menu_DropDown_Eintraege-6, 
                                        KONST.Hoehe_Menu_DropDown_Eintraege-6, (), ())   
        control.setMultipleMode(False)
        
        items = (lang.NEW_PROJECT, 
                lang.OPEN_PROJECT ,
                '---------', 
                lang.NEW_DOC, 
                lang.NEW_DIR,
                '---------',
                lang.EXPORT_2, 
                lang.IMPORT_2,
                '---------',
                lang.BACKUP)
        
        if T.AB != 'Projekt':
            
            items = (
                lang.EXPORT_2, 
                '---------',
                lang.BACKUP)
            
        
        control.addItems(items, 0)
        model.BackgroundColor = KONST.MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Datei', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)
        
    
    def erzeuge_Menu_DropDown_Eintraege_Optionen(self,window,cont):
        try:
            if self.projekt_name != None:
                tag1 = self.settings_proj['tag1']
                tag2 = self.settings_proj['tag2']
                tag3 = self.settings_proj['tag3']
            else:
                tag1 = 0
                tag2 = 0
                tag3 = 0
                
            y = 10
            
            # Titel Baumansicht
            control, model = self.createControl(self.ctx, "FixedText", 10, y,KONST.BREITE_DROPDOWN_OPTIONEN-16, 30-6, (), ())   
            model.Label = self.lang.SICHTBARE_TAGS_BAUMANSICHT
            model.FontWeight = 200
            cont.addControl('Titel_Baumansicht',control)
            
            y += 20
            
            # Tag1
            control_tag1, model_tag1 = self.createControl(self.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag1.Label = self.lang.SHOW_TAG1
            model_tag1.State = tag1
            
            
            y += 16
            
            # Tag2
            control_tag2, model_tag2 = self.createControl(self.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag2.Label = self.lang.SHOW_TAG2
            control_tag2.Enable = False
            model_tag2.State = tag2
            
            y += 16
                
            # Tag3
            control_tag3, model_tag3 = self.createControl(self.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag3.Label = self.lang.SHOW_TAG3
            control_tag3.Enable = False
            model_tag3.State = tag3
                
                
            tag1_listener = Tag1_Item_Listener(self,model_tag1)
            control_tag1.addItemListener(tag1_listener)
            cont.addControl('Checkbox_Tag1', control_tag1)
            cont.addControl('Checkbox_Tag2', control_tag2)
            cont.addControl('Checkbox_Tag3', control_tag3)
            
            
            y += 24
            
            
            # Titel Sidebar
            control, model = self.createControl(self.ctx, "FixedText", 10, y,KONST.BREITE_DROPDOWN_OPTIONEN-20, 30-6, (), ())   
            model.Label = self.lang.SICHTBARE_TAGS_SEITENLEISTE
            model.FontWeight = 200
            cont.addControl('Titel_Baumansicht',control)
    
            y += 20
            
            
            sb_panels_tup = self.class_Sidebar.sb_panels_tup
            sb_panels1 = self.class_Sidebar.sb_panels1
            # Tags Sidebar
            
            listener_SB = Tag_SB_Item_Listener(self)
            
            for panel in sb_panels_tup:
            
                control, model = self.createControl(self.ctx, "CheckBox", 10, y, 
                                                          KONST.BREITE_DROPDOWN_OPTIONEN-20, 20, (), ())   
                model.Label = sb_panels1[panel]
                if panel in self.dict_sb['sichtbare']:
                    model.State = True
                else:
                    model.State = False
                cont.addControl(panel, control)
                control.addItemListener(listener_SB)
                y += 16
    
            y += 24
            
            HOEHE_LISTBOX = 70
            # ListBox
            control, model = self.createControl(self.ctx, "ListBox", 10, y, KONST.BREITE_DROPDOWN_OPTIONEN-20, 
                                            HOEHE_LISTBOX, (), ())   
            control.setMultipleMode(False)
            
            items = (  
                      self.lang.ZEIGE_TEXTBEREICHE,
                     '-------',
                     'Homepage',
                     'Feedback')
            
            control.addItems(items, 0)
            model.BackgroundColor = KONST.MENU_DIALOG_FARBE
            model.Border = False
            
            cont.addControl('Eintraege_Optionen', control)
            
            listener = DropDown_Item_Listener(self)  
            listener.window = window    
            control.addItemListener(listener)  
            
            y += HOEHE_LISTBOX + 10
            
            window.setPosSize(0,0,0,y,8)
        except:
            tb()
            #pd()
    
    def erzeuge_Menu_DropDown_Eintraege_Bearbeiten(self,window,cont):
            
        y = 10
                
        # ListBox
        control, model = self.createControl(self.ctx, "ListBox", 10, y, KONST.Breite_Menu_DropDown_Eintraege-6, 
                                        KONST.Hoehe_Menu_DropDown_Eintraege - 30, (), ())   
        control.setMultipleMode(False)
        
        items = (  
                  self.lang.NEUER_TAB,
                  '---------',
                  self.lang.TRENNE_TEXT,
                  self.lang.UNFOLD_PROJ_DIR
                  )
        if T.AB != 'Projekt':
            items = (  
                  self.lang.NEUER_TAB,
                  self.lang.SCHLIESSE_TAB,
                  self.lang.IMPORTIERE_IN_TAB,
                  '---------',
                  self.lang.UNFOLD_PROJ_DIR
                  )      
        
        control.addItems(items, 0)
        model.BackgroundColor = KONST.MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Bearbeiten', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)  
       
    
    def get_speicherort(self):
        pfad = os.path.join(self.path_to_extension,'pfade.txt')
        
        if os.path.exists(pfad):            
            with codecs_open( pfad, "r","utf-8") as file:
                filepath = file.read() 
            return filepath
        else:
            return None
            
    def get_Klasse_Hauptfeld(self):

        if debug:
            modul = 'h_feld'
            h_feld = load_reload_modul(modul,pyPath,self)  
            
            for imp in IMPORTS:
                exec('h_feld.%s=%s' %(imp,imp))
        else: 
            import h_feld   
            
            for imp in IMPORTS:
                exec('h_feld.%s=%s' %(imp,imp))
                   
        Klasse_Hauptfeld = h_feld.Main_Container(self)
        Klasse_Zeilen_Listener = h_feld.Zeilen_Listener(self.ctx,self)
        return Klasse_Hauptfeld,Klasse_Zeilen_Listener

    def lade_modul(self,modul,arg = None): 
        
        try: 
            if debug:
                load_reload_modul(modul,pyPath,self)
                exec('import '+modul)
                
                for imp in IMPORTS:
                    exec(modul+'.%s=%s' %(imp,imp))

                if arg == None:
                    return eval(modul)
                else:
                    func = modul+arg
                    return eval(func)
            else:
                exec('import '+modul) 
                
                for imp in IMPORTS:
                    exec(modul+'.%s=%s' %(imp,imp))
                
                if arg == None:
                    return eval(modul)
                else:
                    return eval(modul+arg)
        except:
            tb()
     
  
    def lade_Modul_Language(self):
        language = self.doc.CharLocale.Language
        
        if language not in ('de'):
            language = 'en'
            
        self.language = language
        
        import lang_en 
        try:
            exec('import lang_' + language)
        except:
            pass

        if 'lang_' + language in vars():
            lang = vars()['lang_' + language]
        else:
            lang = vars()[lang_en]   

        return lang
    
    
    def Test(self):
        try:
            self.class_Projekt.test()
        except:
            traceback.print_exc()
            
    def erzeuge_Zeile(self,ordner_oder_datei):
        try:
            self.class_Hauptfeld.erzeuge_neue_Zeile(ordner_oder_datei)    
        except:
            tb()      
            
    def leere_Papierkorb(self):
        self.class_Hauptfeld.leere_Papierkorb()   
        
    def erzeuge_Backup(self):
        
        try:
            pfad_zu_backup_ordner = os.path.join(self.pfade['projekt'],'Backups')
            if not os.path.exists(pfad_zu_backup_ordner):
                os.makedirs(pfad_zu_backup_ordner)
            
            lt = time.localtime()
            t = time.strftime(" %d.%m.%Y  %H.%M.%S", lt)
            
            neuer_projekt_name = self.projekt_name + t
            pfad_zu_neuem_ordner = os.path.join(pfad_zu_backup_ordner,neuer_projekt_name)
            
            tree = self.props['Projekt'].xml_tree
            root = tree.getroot()
            
            all_elements = root.findall('.//')
            ordinale = []
            
            for el in all_elements:
                ordinale.append(el.tag)        
        
            self.class_Export.kopiere_projekt(neuer_projekt_name,pfad_zu_neuem_ordner,
                                                 ordinale,tree,self.dict_sb_content,True)  
            os.rename(pfad_zu_neuem_ordner,pfad_zu_neuem_ordner+'.organon')  
            self.Mitteilungen.nachricht('Backup erzeugt unter: %s' %pfad_zu_neuem_ordner+'.organon', "infobox")           
        except:
            tb()
            
            
    def debug_time(self):
        zeit = "%0.2f" %(self.time.clock()-self.timer_start)
        return zeit

    def entferne_alle_listener(self):
        if debug: log(inspect.stack)
        
        #return
        self.current_Contr.removeSelectionChangeListener(self.VC_selection_listener) 
        self.current_Contr.removeKeyHandler(self.keyhandler)
        self.dialog.removeWindowListener(self.w_listener)
        self.undo_mgr.removeUndoManagerListener(self.undo_mgr_listener)
        
    def erzeuge_Dialog_Container(self,posSize,Flags=1+32+64+128):

        ctx = self.ctx
        smgr = self.smgr
        
        X,Y,Width,Height = posSize
    
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
        oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)
    
        # Create Uno Struct
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
        oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
        # global oWindow
        oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
        oWindowDesc.WindowServiceName = ""
        oWindowDesc.Parent = toolkit.getDesktopWindow()
        oWindowDesc.ParentIndex = -1
        oWindowDesc.WindowAttributes = Flags # Flags fuer com.sun.star.awt.WindowAttribute
    
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = Width 
        oRect.Height = Height 
        
        oWindowDesc.Bounds = oRect
    
        # create window
        oWindow = toolkit.createWindow(oWindowDesc)
         
        # create frame for window
        oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame",ctx)
        oFrame.initialize(oWindow)
        #oFrame.setCreator(self.desktop)
        oFrame.activate()
        oFrame.Name = 'Xaver'
        oFrame.Title = 'Xaver2'
        # create new control container
        cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
        cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
        cont_model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE  # 9225984
        cont.setModel(cont_model)
        # need createPeer just only the container
        cont.createPeer(toolkit, oWindow)
        cont.setPosSize(0, 0, 0, 0, 15)
        #pd()
        oFrame.setComponent(cont, None)
        cont.Model.Text = 'Gabriel'

        return oWindow,cont
     
    def loesche_undo_Aktionen(self):
        if debug: log(inspect.stack)
        
        undoMgr = self.doc.UndoManager
        undoMgr.reset()
        
    def speicher_settings(self,dateiname,eintraege):
        if debug: log(inspect.stack)
        
        path = os.path.join(self.pfade['settings'],dateiname)
        imp = str(eintraege).replace(',',',\n')
            
        with open(path , "w") as file:
            file.writelines(imp)
    
    def tree_write(self,tree,pfad):  
        if debug: log(inspect.stack) 
        # diese Methode existiert, um alle Schreibvorgaenge
        # des XML_trees kontrollieren zu koennen
        tree.write(pfad)
        if 'Element' in pfad:
            print(pfad)
        
             
    def oeffne_dokument_in_neuem_fenster(self,URL):
        self.new_doc = self.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,())
            
        contWin = self.new_doc.CurrentController.Frame.ContainerWindow               
        contWin.setPosSize(0,0,870,900,12)
        
        lmgr = self.new_doc.CurrentController.Frame.LayoutManager
        for elem in lmgr.Elements:
        
            if lmgr.isElementVisible(elem.ResourceURL):
                lmgr.hideElement(elem.ResourceURL)
                
        lmgr.HideCurrentUI = True  
        
        viewSettings = self.new_doc.CurrentController.ViewSettings
        viewSettings.ZoomType = 3
        viewSettings.ZoomValue = 100
        viewSettings.ShowRulers = False
    
    
    def get_pyPath(self):
        global pyPath
        pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
        if platform == 'linux':
            pyPath = '/home/xgr/workspace/organonEclipse/py'
            sys.path.append(pyPath)    
   
    # Handy function provided by hanya (from the OOo forums) to create a control, model.
    def createControl(self,ctx,type,x,y,width,height,names,values):
        smgr = ctx.getServiceManager()
        ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
        ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
        ctrl_model.setPropertyValues(names,values)
        ctrl.setModel(ctrl_model)
        ctrl.setPosSize(x,y,width,height,15)
        return (ctrl, ctrl_model)
    
    def createUnoService(self,serviceName):
        sm = uno.getComponentContext().ServiceManager
        return sm.createInstanceWithContext(serviceName, uno.getComponentContext())


class Props():
    def __init__(self):
        self.dict_zeilen_posY = {}
        self.dict_ordner = {}               # enthaelt alle Ordner und alle ihnen untergeordneten Zeilen
        self.dict_bereiche = {}             # drei Unterdicts: Bereichsname,ordinal,Bereichsname-ordinal
        
        self.Hauptfeld = None               # alle Zeilen, Controls
        self.sichtbare_bereiche = []        # Bereichsname ('OrganonSec'+ nr)
        self.kommender_Eintrag = 0
        self.selektierte_zeile = None       # control des Zeilencontainers, Name = ordinal
        self.selektierte_zeile_alt = None   # control 'textfeld' der Zeile
        self.Papierkorb = None              # ordinal des Papierkorbs - wird anfangs einmal gesetzt und bleibt konstant    
        self.Projektordner = None           # ordinal des Projektordners - wird anfangs einmal gesetzt und bleibt konstant 
        self.Papierkorb_geleert = False
        self.tastatureingabe = False
        self.zuletzt_gedrueckte_taste = None

        self.xml_tree = None
        
        self.tab_auswahl = Tab_Auswahl()


class Tab_Auswahl():
    def __init__(self):
        self.rb = None
        self.eigene_auswahl = None
        self.eigene_auswahl_use = None
        
        self.seitenleiste_use = None
        self.seitenleiste_log = None
        self.seitenleiste_log_tags = None
        self.seitenleiste_tags = None
        
        self.baumansicht_use = None
        self.baumansicht_log = None
        self.baumansicht_tags = None
        
        self.suche_use = None
        self.suche_log = None
        #self.seitenleiste_log_tags = None
        self.suche_term = None
        
        self.behalte_hierarchie_bei = None
        self.tab_name = None
        
        
        
class Log():
    
    def __init__(self,mb):  
        self.mb = mb     
        self.time = time
        self.timer_start = self.time.clock()
        
    def debug_time(self):
        zeit = "%0.2f" %(self.time.clock()-self.timer_start)
        return zeit
    
    def log(self,args):
        if debug:
            
            try:
                info = args()

                try:
                    caller = info[2][3]
                except:
                    caller = 'EventObject'
                    
                function = info[1][3]
                modul = info[1][0].f_locals['self'].__class__.__name__

                if modul in ('ViewCursor_Selection_Listener'):
                    return
                if function in ('verlinke_Sektion'):
                    return
                
                if len(modul) > 18:
                    modul = modul[0:18]
                
                string = '%-7s %-18s %-40s %s( caller: %s )' %(self.debug_time(),modul,function,'',caller)
                #time.sleep(0.05)
                print(string)
            
                #self.do(function)
            except:
                tb()

            
        else:
            return
        
    def do(self,*args):
        return
        sections = self.mb.doc.TextSections
        names = sections.ElementNames

        if args[0] == 'schalte_sichtbarkeit_der_Bereiche':
            pd()
        

class Tab ():
    def __init__(self):
        self.AB = 'Projekt'
        
from com.sun.star.awt import XTabListener
class Tab_Listener(unohelper.Base,XTabListener):
    
    def __init__(self,mb):
        self.mb = mb
        # Obwohl der Tab_Listener nur einmal gesetzt wird, wird activated immer 2x aufgerufen (Bug?)
        # Um den Code nicht doppellt auszufuehren, wird id_old verwendet
        self.id_old = None
        

    def inserted(self,id):return False
    def removed(self,id):return False
    def changed(self,id):return False
    def activated(self,id):
        # activated wird beim Erzeugen eines neuen tabs
        # gerufen, bevor self.mb.tabs gesetzt wurde. Daher hier try/except
        # T.AB wird stattdessen in erzeuge_neuen_tab() gesetzt
        
        #print('tab activated',id,self.id_old)
        
        try:
            if id != self.id_old:
                
                if self.mb.bereich_wurde_bearbeitet:
                    tab_name = self.mb.tabs[self.mb.tab_id_old][1]
                    
                    ordinal = self.mb.props[tab_name].selektierte_zeile.AccessibleName
                    bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                    # nachfolgende Zeile erzeugt bei neuem Tab Fehler - 
                    path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)  
                    self.mb.bereich_wurde_bearbeitet = False
                                
                self.mb.tab_umgeschaltet = True
                self.mb.active_tab_id = id
                sichtbare_bereiche = self.mb.props['Projekt'].sichtbare_bereiche
                try:
                    # Wenn neuer Tab erzeugt wird, wird hier ein Fehler erzeugt.
                    # Ist aber egal
                    T.AB = self.mb.tabs[id][1]
                except:
                    pass
                self.mb.props['Projekt'].sichtbare_bereiche = sichtbare_bereiche
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche()

                if self.mb.props[T.AB].selektierte_zeile_alt != None:
                    self.mb.class_Sidebar.passe_sb_an(self.mb.props[T.AB].selektierte_zeile_alt)
                    
             
            #pd()        
            self.id_old = id
        except:
            tb()
            pass
        
    def deactivated(self,id):
        self.mb.tab_id_old = id
        return False
    def disposing(self,arg):return False
    

from com.sun.star.awt import XMouseListener, XItemListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT 
    
class Menu_Kopf_Listener (unohelper.Base, XMouseListener):
    def __init__(self,mb):
        self.mb = mb
        self.menu_Kopf_Eintrag = 'None'
        self.mb.geoeffnetesMenu = None
        
    def mousePressed(self, ev):
        if ev.Buttons == MB_LEFT:
            if self.menu_Kopf_Eintrag == self.mb.lang.FILE:
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)
            elif self.menu_Kopf_Eintrag == 'Test':
                self.mb.Test()          
            elif self.menu_Kopf_Eintrag == self.mb.lang.OPTIONS:
                BREITE = KONST.BREITE_DROPDOWN_OPTIONEN
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev,BREITE)
            elif self.menu_Kopf_Eintrag == self.mb.lang.BEARBEITEN_M:
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)
            
                
            self.mb.loesche_undo_Aktionen()
            return False

    def mouseEntered(self, ev):
        ev.value.Source.Model.FontUnderline = 1     
        if self.menu_Kopf_Eintrag != ev.value.Source.Text:
            self.menu_Kopf_Eintrag = ev.value.Source.Text  
            if None not in (self.menu_Kopf_Eintrag,self.mb.geoeffnetesMenu):
                if self.menu_Kopf_Eintrag != self.mb.geoeffnetesMenu:
                    self.mb.menu_fenster.dispose()
 
        return False
   
    def mouseExited(self, ev):         
        ev.value.Source.Model.FontUnderline = 0
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False

class Menu_Kopf_Listener2 (unohelper.Base, XMouseListener):
    def __init__(self,Class_MenuBar):
        self.mb = Class_MenuBar
        self.geklickterMenupunkt = None
        
    def mousePressed(self, ev):
        #print('mousePressed, Menu_Kopf_Listener2')
        if ev.Buttons == 1:
            if self.mb.projekt_name != None:
                if ev.Source.Model.HelpText == self.mb.lang.INSERT_DOC:            
                    self.mb.erzeuge_Zeile('dokument')
                if ev.Source.Model.HelpText == self.mb.lang.INSERT_DIR:            
                    self.mb.erzeuge_Zeile('Ordner')
                if ev.Source.Model.HelpText == self.mb.lang.CLEAR_RECYCLE_BIN:  
                    self.mb.leere_Papierkorb()
                    
                self.mb.loesche_undo_Aktionen()
                return False
        
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False

class DropDown_Container_Listener (unohelper.Base, XMouseListener):
        def __init__(self,mb):
            self.ob = None
            self.mb = mb
           
        def mousePressed(self, ev):
            #print('mousePressed,DropDown_Container_Listener')  
            if ev.Buttons == MB_LEFT:
                return False
       
        def mouseExited(self, ev): 
            #print('mouseExited')                      
            if self.enthaelt_Punkt(ev):
                pass
            else:            
                self.ob.dispose() 
                self.mb.geoeffnetesMenu = None   
            return False
        
        def enthaelt_Punkt(self, ev):
            #print('enthaelt_Punkt') 
            X = ev.value.X
            Y = ev.value.Y
            
            XTrue = (0 <= X < ev.value.Source.Size.value.Width)
            YTrue = (0 <= Y < ev.value.Source.Size.value.Height)
            
            if XTrue and YTrue:           
                return True
            else:
                return False

        def mouseEntered(self,ev):
            return False
    
class DropDown_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,Class_MenuBar):
        self.mb = Class_MenuBar
        self.window = None
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)      
        sel = ev.value.Source.Items[ev.value.Selected]
        
        #print('self.mb.bereich_wurde_bearbeitet',self.mb.bereich_wurde_bearbeitet)
        # hier evt. Abfrage, ob Bereich bearbeitet -> speichern

        LANG = self.mb.lang
        self.do()
        
        if sel == LANG.NEW_PROJECT:
            self.mb.class_Projekt.erzeuge_neues_Projekt()
            
        elif sel == LANG.OPEN_PROJECT:
            self.mb.class_Projekt.lade_Projekt()
            
        elif sel == LANG.NEW_DOC:
            self.mb.erzeuge_Zeile('dokument')
            
        elif sel == LANG.NEW_DIR:
            self.mb.erzeuge_Zeile('Ordner')
            
        elif sel == LANG.EXPORT_2:
            self.mb.class_Export.export()
            
        elif sel == LANG.IMPORT_2:
            self.mb.class_Import.importX()
            
        elif sel == LANG.UNFOLD_PROJ_DIR:
            self.mb.class_Funktionen.projektordner_ausklappen()
            
        elif sel == LANG.NEUER_TAB:
            self.mb.class_Tabs.start(False)
            
        elif sel == LANG.SCHLIESSE_TAB:
            self.mb.class_Tabs.schliesse_Tab()
            
        elif sel == LANG.ZEIGE_TEXTBEREICHE:
            oBool = self.mb.current_Contr.ViewSettings.ShowTextBoundaries
            self.mb.current_Contr.ViewSettings.ShowTextBoundaries = not oBool  
             
        elif sel == 'Homepage':
            import webbrowser
            webbrowser.open('https://github.com/XRoemer/Organon')
            
        elif sel == 'Feedback':
            import webbrowser
            webbrowser.open('http://organon4office.wordpress.com/')
            
        elif sel == LANG.BACKUP:
            self.mb.erzeuge_Backup()
            
        elif sel == LANG.TRENNE_TEXT:
            self.mb.class_Funktionen.teile_text()
            
        elif sel == LANG.IMPORTIERE_IN_TAB:
            self.mb.class_Tabs.start(True)
            
            

        self.mb.bereich_wurde_bearbeitet = False
        self.mb.loesche_undo_Aktionen()
        
    def do(self): 
        self.window.dispose()
        self.mb.geoeffnetesMenu = None


class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,model):
        self.mb = mb
        self.model = model
        
    # XItemListener    
    def itemStateChanged(self, ev):        
        try:
            sett = self.mb.settings_proj
            
            if self.model.State == 1:
                sett['tag1'] = 1
            else:
                sett['tag1'] = 0
            
            if not sett['tag1']:
                sett['tag1'] = 0
                self.mache_tag1_sichtbar(False)
            else:
                sett['tag1'] = 1
                self.mache_tag1_sichtbar(True) 
            
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        except:
            tb()
    
    def mache_tag1_sichtbar(self,sichtbar):
    
        # alle Zeilen
        controls_zeilen = self.mb.props[T.AB].Hauptfeld.Controls
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        
        if not sichtbar:
            for contr_zeile in controls_zeilen:
                tag1_contr = contr_zeile.getControl('tag1')
                text_contr = contr_zeile.getControl('textfeld')
                posSizeX = text_contr.PosSize.X
                
                text_contr.setPosSize(posSizeX-16,0,0,0,1)

                if self.mb.settings_proj['tag2']:
                    tag2_contr = contr_zeile.getControl('tag2')
                if self.mb.settings_proj['tag3']:
                    tag3_contr = contr_zeile.getControl('tag3')
                    
                tag1_contr.dispose()
                
        if sichtbar:
            for contr_zeile in controls_zeilen:
                text_contr = contr_zeile.getControl('textfeld')
                text_posX = text_contr.PosSize.X
                text_contr.setPosSize(text_posX + 16 ,0,0,0,1)                  

                icon_contr = contr_zeile.getControl('icon')
                icon_posX_end = icon_contr.PosSize.X + icon_contr.PosSize.Width 

                Color__Container = 10202
                Attr = (icon_posX_end,2,16,16,'egal', Color__Container)    
                PosX,PosY,Width,Height,Name,Color = Attr

                ord_zeile = contr_zeile.AccessibleContext.AccessibleName

                zeile_xml = root.find('.//'+ord_zeile)
                tag1 = zeile_xml.attrib['Tag1']

                control_tag1, model_tag1 = self.mb.createControl(self.mb.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
                model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
                model_tag1.Border = 0
                control_tag1.addMouseListener(self.mb.class_Hauptfeld.tag1_listener)

                contr_zeile.addControl('tag1',control_tag1)

class Tag_SB_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb):
        self.mb = mb
        
    def itemStateChanged(self, ev):        
        name = ev.Source.AccessibleContext.AccessibleName
        state = ev.Source.State
        
        panels = self.mb.class_Sidebar.sb_panels2
        if state == 1:
            self.mb.dict_sb['sichtbare'].append(panels[name])
        else:
            self.mb.dict_sb['sichtbare'].remove(panels[name])
        
        try:
            first_element = list(self.mb.dict_sb['controls'])[0]
            self.mb.dict_sb['controls'][first_element][1].requestLayout()
        except:
            tb()


            
from com.sun.star.awt import Rectangle
from com.sun.star.awt import WindowDescriptor         
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK,YES_NO_CANCEL, DEF_NO

class Mitteilungen():
    def __init__(self,ctx,mb):
        self.ctx = ctx
        self.mb = mb     
    
    def nachricht(self, MsgText, MsgType="errorbox", MsgButtons=OK):                 

        smgr = self.ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
        doc = desktop.getCurrentComponent()                
        ParentWin = doc.CurrentController.Frame.ContainerWindow

        MsgTitle = "Mitteilung"
        MsgType = MsgType.lower()
        #available msg types
        MsgTypes = ("messbox", "infobox", "errorbox", "warningbox", "querybox")
     
        if MsgType not in MsgTypes:
            MsgType = "messbox"
        
        #describe window properties.
        aDescriptor = WindowDescriptor()
        aDescriptor.Type = MODALTOP
        aDescriptor.WindowServiceName = MsgType
        aDescriptor.ParentIndex = -1
        aDescriptor.Parent = ParentWin
        aDescriptor.Bounds = Rectangle()
        aDescriptor.WindowAttributes = MsgButtons
        
        tk = ParentWin.getToolkit()
        msgbox = tk.createWindow(aDescriptor)
        msgbox.MessageText = MsgText
        
        x = msgbox.execute()
        msgbox.dispose()
        return x

from com.sun.star.document import XUndoManagerListener 
class Undo_Manager_Listener(unohelper.Base,XUndoManagerListener): 
    
    def __init__(self,mb):
        self.mb = mb 
        self.textbereiche = ()
        
        
    def enteredContext(self,ev):
        if self.mb.use_UM_Listener == False:
            return
        if ev.UndoActionTitle == self.mb.BEREICH_EINFUEGEN:
            if self.mb.debug: log(inspect.stack)
            
            if self.mb.doc.TextSections.Count == 0:
                self.textbereiche = ()
            else:
                self.textbereiche = self.mb.doc.TextSections.ElementNames
    
    
    def leftContext(self,ev):
        if self.mb.use_UM_Listener == False:
            return
        if ev.UndoActionTitle == self.mb.BEREICH_EINFUEGEN:
            if self.mb.debug: log(inspect.stack)
            
            for tbe in self.mb.doc.TextSections.ElementNames:
                if 'trenner' not in tbe:
                    if tbe not in self.textbereiche:
                        self.bereich_in_OrganonSec_einfuegen(tbe)
    
    
    def undoActionAdded(self,ev):return False
    def actionUndone(self,ev):return False
    def actionRedone(self,ev):return False
    def allActionsCleared(self,ev):return False
    def redoActionsCleared(self,ev):return False
    def resetAll(self,ev):return False
    def enteredHiddenContext(self,ev):return False
    def leftHiddenContext(self,ev):return False
    def cancelledContext(self,ev):return False
    def disposing(self,ev):return False
    
    def bereich_in_OrganonSec_einfuegen(self,tbe):
        if self.mb.debug: log(inspect.stack)
        
        text = self.mb.doc.Text
        vc = self.mb.viewcursor
        TS = self.mb.doc.TextSections
        
        sec = TS.getByName(tbe)
        sec_name = sec.Name
        
        if sec.ParentSection == None:
            
            position_neuer_Bereich = None
            
            cur2 = self.mb.doc.Text.createTextCursorByRange(vc)
            cur2.collapseToStart()
            cur2.goLeft(1,False)
            
            if cur2.TextSection == sec:
                position_neuer_Bereich = 'davor'
            
            else:
                cur2.gotoRange(vc,False)
                cur2.collapseToEnd()
                cur2.goRight(1,False)
             
                if cur2.TextSection == sec:
                    position_neuer_Bereich = 'danach'
            
            cur = self.mb.doc.Text.createTextCursorByRange(vc)
            cur.collapseToEnd()
            
            if position_neuer_Bereich == 'davor':
                
                cur.gotoRange(sec.Anchor,True)
                cur.setString('')
                self.mb.doc.Text.insertString(vc,' ',False)
                cur.gotoRange(vc,False)
                goLeft = 1
                
            elif position_neuer_Bereich == 'danach':
        
                self.mb.doc.Text.insertString(vc,' ',False)
                
                cur.gotoRange(sec.Anchor,True)
                cur.setString('')
                cur.gotoRange(vc,False)
                cur.goLeft(1,False)
                goLeft = 2
                
            else:
                return
            
            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection.setName(sec_name)
            
            self.mb.undo_mgr.removeUndoManagerListener(self.mb.undo_mgr_listener)
            self.mb.doc.Text.insertTextContent(cur,newSection,False)
            self.mb.undo_mgr.addUndoManagerListener(self.mb.undo_mgr_listener)
            
            vc.goLeft(goLeft,False)

        # Wenn ein Bereich eingefuegt wurde, auf jeden Fall speichern
        section = vc.TextSection
        while section != None:
            bereichsname = section.Name
            section = section.ParentSection
        
        path = self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname]
        path = uno.systemPathToFileUrl(path)
        self.mb.props[T.AB].tastatureingabe = True
        self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)
        

from com.sun.star.awt import XKeyHandler
class Key_Handler(unohelper.Base, XKeyHandler):
    
    def __init__(self,mb):
        #if debug:print('init Keyhandler')
        self.mb = mb
        self.mb.keyhandler = self
        mb.current_Contr.addKeyHandler(self)
        
    def keyPressed(self,ev):
        #print(ev.KeyChar)
        self.mb.props[T.AB].tastatureingabe = True
        self.mb.props[T.AB].zuletzt_gedrueckte_taste = ev
        return False
        
    def keyReleased(self,ev):
        
        if self.mb.projekt_name != None:
            # Wenn eine OrganonSec durch den Benutzer geloescht wird, wird die Aktion rueckgaengig gemacht
            # KeyCodes: backspace, delete
            if ev.KeyCode in (1283,1286):
                anz_im_bereiche_dict = len(self.mb.props[T.AB].dict_bereiche['ordinal'])
                anz_im_dok = 0
                for sec in self.mb.doc.TextSections.ElementNames:
                    if 'OrganonSec' in sec:
                        anz_im_dok += 1
                if anz_im_dok < anz_im_bereiche_dict:
                    if self.mb.debug: log(inspect.stack)
                    self.mb.doc.UndoManager.undo()
        return False


from com.sun.star.view import XSelectionChangeListener
class ViewCursor_Selection_Listener(unohelper.Base, XSelectionChangeListener):
    
    def __init__(self,mb):
        if debug: log(inspect.stack)
        self.mb = mb
        self.ts_old = 'nicht vorhanden'
        self.mb.selbstruf = False
        
    def disposing(self,ev):
        if self.mb.debug: log(inspect.stack)
        return False
    
    def selectionChanged(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            if self.mb.selbstruf:
                if self.mb.debug: print('selection selbstruf')
                return False
    
            selected_ts = self.mb.current_Contr.ViewCursor.TextSection            
            if selected_ts == None:
                return False
            
            s_name = selected_ts.Name
            
            # stellt sicher, dass nur selbst erzeugte Bereiche angesprochen werden
            # und der Trenner uebersprungen wird
            if 'trenner'  in s_name:
                print('trenner')
                if self.mb.props[T.AB].zuletzt_gedrueckte_taste == None:
                    try:
                        self.mb.viewcursor.goDown(1,False)
                    except:
                        self.mb.viewcursor.goUp(1,False)
                    return False
                # 1024,1027 Pfeil runter,rechts
                elif self.mb.props[T.AB].zuletzt_gedrueckte_taste.KeyCode in (1024,1027):  
                    self.mb.viewcursor.goDown(1,False)       
                else:
                    self.mb.viewcursor.goUp(1,False)
                # sollte der viewcursor immer noch auf einem Trenner stehen,
                # befindet er sich im letzten Bereich -> goUp    
                if 'trenner' in self.mb.viewcursor.TextSection.Name:
                        self.mb.viewcursor.goUp(1,False)
                return False 
            
            # test ob ausgewaehlter Bereich ein Kind-Bereich ist -> Selektion wird auf Parent gesetzt
            elif 'trenner' not in s_name and 'OrganonSec' not in s_name:
                sec = []
                self.test_for_parent_section(selected_ts,sec)
                selected_ts = sec[0]
                s_name = selected_ts.Name
                
                # steht nach test_for... selcted_text... nicht auf einer OrganonSec, 
                # ist der Bereich außerhalb des Organon trees
                if 'OrganonSec' not in selected_ts.Name:
                    return False
                
            self.so_name =  None   
                 
            if self.mb.props[T.AB].selektierte_zeile_alt != None:
                ts_old_ordinal = self.mb.props[T.AB].selektierte_zeile_alt.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
                ts_old_bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ts_old_ordinal]
                self.ts_old = self.mb.doc.TextSections.getByName(ts_old_bereichsname)            
                self.so_name = self.mb.props[T.AB].dict_bereiche['ordinal'][ts_old_ordinal]
                
            if self.ts_old == 'nicht vorhanden':
                #print('selek gewechs, old nicht vorhanden')
                self.mb.bereich_wurde_bearbeitet = True
                self.ts_old = selected_ts 
                return False 
            elif self.mb.props[T.AB].Papierkorb_geleert == True:
                #print('selek gewechs, Papierkorb_geleert')
                # fehlt: nur speichern, wenn die Datei nicht im Papierkorb gelandet ist
                self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.ts_old.FileLink.FileURL,self.so_name)
                self.ts_old = selected_ts 
                self.mb.props[T.AB].Papierkorb_geleert = False 
                return False       
            else:
                if self.ts_old == selected_ts:
                    #print('selek nix gewechs',self.so_name , s_name)
                    self.mb.bereich_wurde_bearbeitet = True
                    return False                
                else:
                    #print('selek gewechs',self.so_name , s_name)
                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.ts_old.FileLink.FileURL,self.so_name)
                    self.ts_old = selected_ts 
                    self.farbe_der_selektion_aendern(selected_ts.Name)
                    return False 
        except:
            tb()

   
    def test_for_parent_section(self,selected_text_sectionX,sec):
        if selected_text_sectionX.ParentSection != None:
            selected_text_sectionX = selected_text_sectionX.ParentSection
            self.test_for_parent_section(selected_text_sectionX,sec)
        else:
            sec.append(selected_text_sectionX)
        
                  
    def farbe_der_selektion_aendern(self,bereichsname):      
        #pd()
        ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][bereichsname]
        zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)
        textfeld = zeile.getControl('textfeld')
        
        self.mb.props[T.AB].selektierte_zeile = zeile.AccessibleContext
        # selektierte Zeile einfaerben, ehem. sel. Zeile zuruecksetzen
        textfeld.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE 
        if self.mb.props[T.AB].selektierte_zeile_alt != None:  
            self.mb.props[T.AB].selektierte_zeile_alt.Model.BackgroundColor = KONST.FARBE_ZEILE_STANDARD
            self.mb.class_Sidebar.passe_sb_an(textfeld)
        self.mb.props[T.AB].selektierte_zeile_alt = textfeld
  
    


        

            
from com.sun.star.awt import XWindowListener
class Dialog_Window_Listener(unohelper.Base,XWindowListener):
    
    def __init__(self,mb):
        self.mb = mb
        
    def windowResized(self,ev):
        #print('windowResized')
        self.korrigiere_hoehe_des_scrollbalkens()
        self.mb.class_Hauptfeld.korrigiere_scrollbar()
    def windowMoved(self,ev):pass
        #print('windowMoved')
    def windowShown(self,ev):
        self.korrigiere_hoehe_des_scrollbalkens()
        #print('windowShown')
    
    def windowHidden(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:                
            if not self.mb.tab_umgeschaltet:  
                # speichern, wenn Organon beendet wird.
                # aenderungen nach tabwechsel werden in Tab_Listener.activated() gespeichert
                if self.mb.bereich_wurde_bearbeitet:
                    tab_name = self.mb.tabs[self.mb.tab_id_old][1]
                    ordinal = self.mb.props[tab_name].selektierte_zeile.AccessibleName
                    bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                    path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)   
                               
                if 'files' in self.mb.pfade: 
                    self.mb.class_Sidebar.speicher_sidebar_dict()       
                    self.mb.class_Sidebar.dict_sb_zuruecksetzen()
                
                self.mb.class_Sidebar.toggle_sicht_sidebar()
                self.mb.entferne_alle_listener() 
                self.mb = None
            else:
                self.mb.tab_umgeschaltet = False
        except:
            tb()
    
    def korrigiere_hoehe_des_scrollbalkens(self):
        try:
            active_tab = self.mb.active_tab_id
            win = self.mb.tabs[active_tab][0]
            nav_cont_aussen = win.getControl('Hauptfeld_aussen')

            # nav_cont_aussen ist None, wenn noch kein Projekt geoeffnet wurde
            if nav_cont_aussen != None:
                nav_cont = nav_cont_aussen.getControl('Hauptfeld')
                  
                MenuBar = win.getControl('Organon_Menu_Bar')
                MBHoehe = MenuBar.PosSize.value.Height + MenuBar.PosSize.value.Y
                NCHoehe = 0 #nav_cont.PosSize.value.Height
                NCPosY  = nav_cont.PosSize.value.Y
                Y =  NCHoehe + NCPosY + MBHoehe
                Height = win.PosSize.value.Height - Y -25
                
                scrll = win.getControl('ScrollBar')
                scrll.setPosSize(0,0,0,Height,8)
        except:
            tb()


    def disposing(self,arg):
        return False

    
################ TOOLS ################################################################



def load_reload_modul(modul,pyPath,mb):
    try:
        if pyPath not in sys.path:
            sys.path.append(pyPath)

        #print('lade:',modul)
        exec('import '+ modul)
        del(sys.modules[modul])
        try:
            if mb.programm == 'LibreOffice':
                import shutil
                shutil.rmtree(os.path.join(pyPath,'__pycache__'))

            elif mb.programm == 'OpenOffice':

                path_menu = __file__.split(__name__)
                path = path_menu[0] + modul + '.pyc'
                #print(path)
                try:
                    os.remove(path)
                except:
                    pass
        except:
            tb()
                            
        exec('import '+ modul)

        return eval(modul)
    except:
        tb()
        

    
    
