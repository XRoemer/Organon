# -*- coding: utf-8 -*-
import uno
import unohelper
import traceback
import sys
import os
import xml.etree.ElementTree as ElementTree
import time
import codecs

global oxt
oxt = False

tb = traceback.print_exc
platform = sys.platform

if oxt:
    pyPath = 'E:\\Eclipse_Workspace\\orga\\organon\\py'
    if platform == 'linux':
        pyPath = '/home/xgr/Arbeitsordner/organon/py'
        sys.path.append(pyPath)
    
Color_MenuBar_Container = 13027014 #16777215 #weiss
Color_MenuBar_MenuEintraege = 6279861
Color_Menu_Container = 16771990
MENU_DIALOG_FARBE = 14804725#13884141 #305099 # blau

Breite_Menu_DropDown_Eintraege = 150
Hoehe_Menu_DropDown_Eintraege = 180
Abstand = 10
Breite_Menu_DropDown_Container = Breite_Menu_DropDown_Eintraege + Abstand
Hoehe_Menu_DropDown_Container = Hoehe_Menu_DropDown_Eintraege + Abstand

IMG_ORDNER_NEU_24 =   'vnd.sun.star.extension://xaver.roemers.organon/img/OrdnerNeu_24.png'
IMG_DATEI_NEU_24 =    'vnd.sun.star.extension://xaver.roemers.organon/img/neueDatei_24.png'

from com.sun.star.awt import XMouseListener, XItemListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT


class Menu_Bar():
    
    def __init__(self,pdk,dialog,ctx,tabs,path_to_extension):
        
        self.pd = pdk
        global pd
        pd = pdk
        #pd()

        if 'LibreOffice' in sys.executable:
            self.programm = 'LibreOffice'
        elif 'OpenOffice' in sys.executable:
            self.programm = 'OpenOffice'
        else:
            # Fuer Linux / OSX fehlt
            self.programm = 'LibreOffice'
        
        # Konstanten
        self.dialog = dialog
        self.ctx = ctx
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
        self.doc = self.desktop.getCurrentComponent() 
        self.current_Contr = self.doc.CurrentController 
        self.viewcursor = self.current_Contr.ViewCursor
        self.tabs = tabs
        self.platform = sys.platform
        self.lang = self.lade_Modul_Language()
        self.path_to_extension = path_to_extension

        # Properties
        self.projekt_name = None
        self.speicherort_last_proj = self.get_speicherort()
        self.Hauptfeld = None               # alle Zeilen, Controls
        self.dict_zeilen_posY = {}
        self.dict_ordner = {}               # enthaelt alle Ordner und alle ihnen untergeordneten Zeilen
        self.dict_bereiche = {}             # besitzt drei Unterdicts: Bereichsname,ordinal,Bereichsname-ordinal
        self.sichtbare_bereiche = []        # Bereichsname ('OrganonSec'+ nr)
        self.xml_tree = None
        self.xml_tree_settings = None
        self.kommender_Eintrag = 0
        self.selektierte_zeile = None       # control des Zeilencontainers, Name = ordinal
        self.selektierte_Zeile_alt = None   # control 'textfeld' der Zeile
        self.Papierkorb = None              # ordinal des Papierkorbs - wird anfangs einmal gesetzt und bleibt konstant      
        self.Papierkorb_geleert = False
        self.bereich_wurde_bearbeitet = False
        self.tastatureingabe = False
        self.zuletzt_gedrueckte_taste = None
        self.exp_settings = None
        self.imp_settings_xml = None
        self.settings_org_xml = None
        
        #Settings
        self.tag1_visible = None
        self.tag2_visible = None
        self.tag3_visible = None
        # Pfade
        self.pfade = {}

        # Klassen    
        self.class_Hauptfeld,self.class_Zeilen_Listener = self.get_Klasse_Hauptfeld()
        self.class_Projekt = None # wird durch die folgende Zeile geladen
        self.lade_Modul_Projekte() # = import projects        
        self.class_XML = self.get_xml_class(dialog,ctx,self,pdk)
        self.class_Funktionen = self.lade_Modul_Funktionen()
        self.Key_Handler = Key_Handler(self)
        self.ET = ElementTree  
        self.VC_selection_listener = None # wird in get_Klasse_Bereiche gesetzt
        self.w_listener = None # wird in get_Klasse_Bereiche gesetzt
        self.class_Bereiche = self.get_Klasse_Bereiche()  
        self.class_Export = self.lade_Modul_Export()
        self.class_Import = self.lade_Modul_Import()
        self.Mitteilungen = Mitteilungen(self.dialog,self.ctx,self)
         

        # noch einzufuegende ???
        self.hf_controls = None
        

        # fuers debugging
        self.debug = True
        if self.debug: print('Debug = True ; Menu_Bar')
        self.time = time
        self.timer_start = self.time.clock()

        self.dialog.addWindowListener(self.w_listener)
        
        # prueft, ob eine Organon Datei geladen wurde
        UD_properties = self.doc.DocumentProperties.UserDefinedProperties
        has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
        
        
#         if has_prop:
#             self.entferne_alle_listener()
#             dialog_contr = self.dialog.Controls
#             for contr in dialog_contr:
#                 contr.dispose()
        if has_prop:    
            self.projekt_name = UD_properties.getPropertyValue('ProjektName')
            self.erzeuge_MenuBar_Container()
            self.erzeuge_Menu()
            #self.class_Projekt.lade_Projekt(False)
           
       
    
   
    def erzeuge_Menu(self):
        try:             
            listener = Menu_Kopf_Listener(self) 
            listener2 = Menu_Kopf_Listener2(self) 
            self.erzeuge_MenuBar_Container()
            
            self.erzeuge_Menu_Kopf_Datei(listener)
            self.erzeuge_Menu_Kopf_Optionen(listener)
            
            if oxt:
                self.erzeuge_Menu_Kopf_Test(listener)
            
            self.erzeuge_Menu_neuer_Ordner(listener2)
            self.erzeuge_Menu_Kopf_neues_Dokument(listener2)
            self.erzeuge_Menu_Kopf_Papierkorb_leeren(listener2)
            
        except Exception as e:
                self.Mitteilungen.nachricht(str(e),"warningbox")
                tb()

    
    def erzeuge_MenuBar_Container(self):
        
        Attr = (2, 2, 1000, 20, 'zweiter_cont', Color_MenuBar_Container)    
        PosX, PosY, Width, Height, Name, Color = Attr
        
        menuB_control, menuB_model = createControl3(self.ctx, "Container", PosX, PosY, Width, Height, (), ())          
        menuB_model.BackgroundColor = Color    
         
        self.dialog.addControl('Organon_Menu_Bar', menuB_control)
    
   
   
        
    def erzeuge_Menu_Kopf_Datei(self,listener):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        
        Attr = (0, 2, 35, 20, 'datei', Color_MenuBar_MenuEintraege)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "FixedText", PosX, PosY, Width, Height, (), ())           
        #model.BackgroundColor = Color
        model.Label = self.lang.FILE             
        control.addMouseListener(listener)
           
        MenuBarCont.addControl('Datei', control)
    
    
    def erzeuge_Menu_Kopf_Optionen(self,listener):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        
        Attr = (37, 2, 55, 20, 'optionen', Color_MenuBar_MenuEintraege)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "FixedText", PosX, PosY, Width, Height, (), ())           
        #model.BackgroundColor = Color
        model.Label = self.lang.OPTIONS           
        control.addMouseListener(listener)
               
        MenuBarCont.addControl('Optionen', control)
        
        
    def erzeuge_Menu_Kopf_Test(self,listener):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        
        Attr = (300, 2, 50, 20, 'projekt', Color_MenuBar_MenuEintraege)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "FixedText", PosX, PosY, Width, Height, (), ())           
        #model.BackgroundColor = Color
        model.Label = 'Test'                     
        control.addMouseListener(listener) 
              
        MenuBarCont.addControl('Projekt', control)
        
        
    def erzeuge_Menu_neuer_Ordner(self,listener2):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        Color = 102
        Attr = (120, 0, 20, 20, 'probe', Color)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "ImageControl", PosX, PosY, Width, Height, (), ())   
        model.ImageURL = IMG_ORDNER_NEU_24
        
            
        model.HelpText = self.lang.INSERT_DIR
        model.Border = 0                    
        control.addMouseListener(listener2) 
              
        MenuBarCont.addControl('Ordner', control)
        
        
    def erzeuge_Menu_Kopf_neues_Dokument(self,listener2):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        Color = 102
        Attr = (140, 0, 20, 20, 'probe', Color)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "ImageControl", PosX, PosY, Width, Height, (), ())           
        model.ImageURL = IMG_DATEI_NEU_24
                    
        model.HelpText = self.lang.INSERT_DOC
        model.Border = 0                    
        control.addMouseListener(listener2) 
            
        MenuBarCont.addControl('neues_Dokument', control)
        
    def erzeuge_Menu_Kopf_Papierkorb_leeren(self,listener2):
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        Color = 102
        Attr = (220, 0, 20, 20, 'probe', Color)    
        PosX, PosY, Width, Height, Name, Color = Attr
         
        control, model = createControl3(self.ctx, "ImageControl", PosX, PosY, Width, Height, (), ())           
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leeren.png'
        model.HelpText = self.lang.CLEAR_RECYCLE_BIN
        model.Border = 0                       
        control.addMouseListener(listener2) 
             
        MenuBarCont.addControl('Papierkorb_leeren', control)
        



    def erzeuge_Menu_DropDown_Container(self,ev):
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
        gnDefaultWindowAttributes = 1 + 16  # + 32 +64+128 
    #         com_sun_star_awt_WindowAttribute_SHOW + \
    #         com_sun_star_awt_WindowAttribute_BORDER + \
    #         com_sun_star_awt_WindowAttribute_MOVEABLE + \
    #         com_sun_star_awt_WindowAttribute_CLOSEABLE + \
    #         com_sun_star_awt_WindowAttribute_SIZEABLE
          
        X = ev.value.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.value.Source.AccessibleContext.LocationOnScreen.value.Y + ev.value.Source.Size.value.Height
    
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = Breite_Menu_DropDown_Container
        oRect.Height = Hoehe_Menu_DropDown_Container
        
        oWindowDesc.Bounds = oRect
      
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
        cont_model.BackgroundColor = MENU_DIALOG_FARBE  # 9225984
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
        if Name == self.lang.OPTIONS:
            self.erzeuge_Menu_DropDown_Eintraege_Optionen(oWindow, cont)
        return Name
    
    def erzeuge_Menu_DropDown_Eintraege_Datei(self,window,cont):
        lang = self.lang
        control, model = createControl3(self.ctx, "ListBox", 10 ,  10 , Breite_Menu_DropDown_Eintraege-6, Hoehe_Menu_DropDown_Eintraege-6, (), ())   
        control.setMultipleMode(False)
        item = (lang.NEW_PROJECT, lang.OPEN_PROJECT ,'---------', lang.NEW_DOC, lang.NEW_DIR,'---------',lang.EXPORT_2, lang.IMPORT_2)
        control.addItems(item, 0)
        model.BackgroundColor = MENU_DIALOG_FARBE
        model.Border = False
        #model.FontCharWidth = 20
        
        cont.addControl('Eintraege_Datei', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)
        
    
    def erzeuge_Menu_DropDown_Eintraege_Optionen(self,window,cont):
        
        # Tag1
        control_tag1, model_tag1 = createControl3(self.ctx, "CheckBox", 10, 10, Breite_Menu_DropDown_Eintraege-6, 30-6, (), ())   
        model_tag1.Label = self.lang.SHOW_TAG1
        
        if self.tag1_visible == True:
            model_tag1.State = 1
        else:
            model_tag1.State = 0
            
        tag1_listener = Tag1_Item_Listener(self,model_tag1)
        control_tag1.addItemListener(tag1_listener)
        cont.addControl('Checkbox_Tag1', control_tag1)
        

        # ListBox
        control, model = createControl3(self.ctx, "ListBox", 10, 34, Breite_Menu_DropDown_Eintraege-6, 
                                        Hoehe_Menu_DropDown_Eintraege - 30, (), ())   
        control.setMultipleMode(False)
        item = ( self.lang.UNFOLD_PROJ_DIR, '-------', '#Homepage', '#Updates', '#Etc.')
        control.addItems(item, 0)
        model.BackgroundColor = MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Optionen', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)   
    
    
    def get_speicherort(self):
        pfad = os.path.join(self.path_to_extension,'pfade.txt')
        
        if os.path.exists(pfad):            
            with codecs.open( pfad, "r","utf-8") as file:
                filepath = file.read() 
            return filepath
        else:
            return None
            
    def get_Klasse_Hauptfeld(self):

        if oxt:
            modul = 'h_feld'
            h_feld = load_reload_modul(modul,pyPath,self)  
        else: 
            import h_feld   
                   
        Klasse_Hauptfeld = h_feld.Main_Container(self)
        Klasse_Zeilen_Listener = h_feld.Zeilen_Listener(self.Hauptfeld,self.ctx,self)
        return Klasse_Hauptfeld,Klasse_Zeilen_Listener

            
            
    def get_Klasse_Bereiche(self):
        
        if oxt:
            modul = 'bereiche'
            bereiche = load_reload_modul(modul,pyPath,self) 
        else:
            import bereiche
            
        self.VC_selection_listener = bereiche.ViewCursor_Selection_Listener(self)          
        Klasse_Bereiche = bereiche.Bereiche(self)
        self.w_listener = bereiche.Dialog_Window_Listener(self)
        
        return Klasse_Bereiche
      
    
    def lade_Modul_Projekte(self):
        if oxt:
            modul = 'projects'
            projects = load_reload_modul(modul,pyPath,self)
        else:
            import projects
            
        np = projects.Projekt(self, pd)
        self.class_Projekt = np

    
    def lade_Modul_Funktionen(self):
        if oxt:
            modul = 'funktionen'
            funktionen = load_reload_modul(modul,pyPath,self)
        else:
            import funktionen
            
        ff = funktionen.Funktionen(self, pd)
        return ff
    
    def lade_Modul_Language(self):
        language = self.doc.CharLocale.Language
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
    
    
    def lade_Modul_Export(self):
        if oxt:
            modul = 'export'
            export = load_reload_modul(modul,pyPath,self)
        else:
            import export
          
        Export = export.Export(self,pd)  
        
        return Export
    
    def lade_Modul_Import(self):
        if oxt:
            modul = 'importX'
            importX = load_reload_modul(modul,pyPath,self)
        else:
            import importX
          
        ImportX = importX.ImportX(self,pd)  
        
        return ImportX
    
    def erzeuge_neue_Projekte(self):
        try:
            self.class_Projekt.test()
        except:
            traceback.print_exc()
            
    def erzeuge_Zeile(self,ordner_oder_datei):
        self.class_Hauptfeld.erzeuge_neue_Zeile(ordner_oder_datei)          

            
    def leere_Papierkorb(self):
        self.class_Hauptfeld.leere_Papierkorb()          

    def speicher_Projekt(self):
        self.class_Projekt.speicher_Projekt()          

    def debug_time(self):
        zeit = "%0.2f" %(self.time.clock()-self.timer_start)
        return zeit
    
    def get_xml_class(self,dialog,ctx,mb,pd):
    
        if oxt:
            modul = 'xml_m'
            xml_m = load_reload_modul(modul,pyPath,self) 
        else:
            import xml_m
            
        XML_M = xml_m.XML_Methoden(mb,pd)
        return XML_M


    def entferne_alle_listener(self):
        self.current_Contr.removeSelectionChangeListener(self.VC_selection_listener) 
        self.current_Contr.removeKeyHandler(self.keyhandler)
        self.dialog.removeWindowListener(self.w_listener)
    
    
class Menu_Kopf_Listener (unohelper.Base, XMouseListener):
    def __init__(self,mb):
        self.mb = mb
        self.menu_Kopf_Eintrag = 'None'
        self.mb.geoeffnetesMenu = None
        
    def mousePressed(self, ev):
        if ev.Buttons == MB_LEFT:
            #print('maus gepresst')
            if self.menu_Kopf_Eintrag == self.mb.lang.FILE:
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)            
            if self.menu_Kopf_Eintrag == self.mb.lang.OPTIONS:
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)
            if self.menu_Kopf_Eintrag == 'Test':
                self.mb.erzeuge_neue_Projekte()
            return False

    def mouseEntered(self, ev):
        #print('maus kommt')
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
            if ev.Source.Model.HelpText == self.mb.lang.INSERT_DOC:            
                self.mb.erzeuge_Zeile('dokument')
            if ev.Source.Model.HelpText == self.mb.lang.INSERT_DIR:            
                self.mb.erzeuge_Zeile('Ordner')
            if ev.Source.Model.HelpText == self.mb.lang.CLEAR_RECYCLE_BIN:            
                self.mb.leere_Papierkorb()
#             if ev.Source.Model.HelpText == 'Projekt speichern':            
#                 self.mb.speicher_Projekt()
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
        sel = ev.value.Source.Items[ev.value.Selected]

        if sel == self.mb.lang.NEW_PROJECT:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.class_Projekt.erzeuge_neues_Projekt()
        if sel == self.mb.lang.OPEN_PROJECT:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.class_Projekt.lade_Projekt()
        if sel == self.mb.lang.NEW_DOC:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.erzeuge_Zeile('dokument')
        if sel == self.mb.lang.NEW_DIR:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.erzeuge_Zeile('Ordner')
        if sel == self.mb.lang.EXPORT_2:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.class_Export.export()
        if sel == self.mb.lang.IMPORT_2:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.class_Import.importX()
        if sel == self.mb.lang.UNFOLD_PROJ_DIR:
            self.window.dispose()
            self.mb.geoeffnetesMenu = None
            self.mb.class_Funktionen.projektordner_ausklappen()
            


class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,model):
        self.mb = mb
        self.model = model
        
    # XItemListener    
    def itemStateChanged(self, ev):        

        if self.model.State == 1:
            self.mb.tag1_visible = True
        else:
            self.mb.tag1_visible = False
        
        tree = self.mb.xml_tree_settings
        root = tree.getroot()  
        xml_tag1 = root.find(".//tag1")

        if self.mb.tag1_visible == False:
            xml_tag1.attrib['sichtbar'] = 'nein'
            self.mache_tag1_sichtbar(False)
        else:
            xml_tag1.attrib['sichtbar'] = 'ja'
            self.mache_tag1_sichtbar(True)

        Path = os.path.join(self.mb.pfade['settings'] , 'settings.xml' )
        self.mb.xml_tree_settings.write(Path)  
    
    def mache_tag1_sichtbar(self,sichtbar):
    
        # alle Zeilen
        controls_zeilen = self.mb.Hauptfeld.Controls
        tree = self.mb.xml_tree
        root = tree.getroot()
        
        if not sichtbar:
            for contr_zeile in controls_zeilen:
                tag1_contr = contr_zeile.getControl('tag1')
                text_contr = contr_zeile.getControl('textfeld')
                posSizeX = text_contr.PosSize.X
                
                text_contr.setPosSize(posSizeX-16,0,0,0,1)

                if self.mb.tag2_visible:
                    tag2_contr = contr_zeile.getControl('tag2')
                if self.mb.tag3_visible:
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

                control_tag1, model_tag1 = createControl3(self.mb.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
                model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
                model_tag1.Border = 0
                control_tag1.addMouseListener(self.mb.class_Hauptfeld.tag1_listener)

                contr_zeile.addControl('tag1',control_tag1)

    
    

from com.sun.star.awt import XKeyHandler
class Key_Handler(unohelper.Base, XKeyHandler):
    
    def __init__(self,mb):
        if oxt:print('init Keyhandler')
        self.mb = mb
        self.mb.keyhandler = self
        mb.current_Contr.addKeyHandler(self)
        
    def keyPressed(self,ev):
        #print(ev.KeyChar)
        self.mb.tastatureingabe = True
        self.mb.zuletzt_gedrueckte_taste = ev
        return False
        
    def keyReleased(self,ev):
        return False
               
    
def get_work_path(ctx):        
    smgr = ctx.getServiceManager()        
    paths = smgr.createInstance( "com.sun.star.util.PathSettings" )
    doc_path = paths.Work_writable
    return doc_path+'/Organon'
                

            
from com.sun.star.awt import Rectangle
from com.sun.star.awt import WindowDescriptor         
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK,YES_NO_CANCEL, DEF_NO

class Mitteilungen():
    def __init__(self,dialog,ctx,mb):
        self.dialog = dialog  
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

        
   
################ TOOLS ################################################################

# Handy function provided by hanya (from the OOo forums) to create a control, model.
def createControl3(ctx, type, x, y, width, height, names, values):
   smgr = ctx.getServiceManager()
   ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type, ctx)
   ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type, ctx)
   ctrl_model.setPropertyValues(names, values)
   ctrl.setModel(ctrl_model)
   ctrl.setPosSize(x, y, width, height, 15)
   return (ctrl, ctrl_model)



def createUnoService(serviceName):
  sm = uno.getComponentContext().ServiceManager
  return sm.createInstanceWithContext(serviceName, uno.getComponentContext())


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
                if platform == 'linux':
                    shutil.rmtree(pyPath+'/__pycache__')
                else:
                    shutil.rmtree(pyPath+'\\__pycache__')
            elif mb.programm == 'OpenOffice':

                path_menu = __file__.split(__name__)
                path = path_menu[0] + modul + '.pyc'
                #print(path)
                try:
                    os.remove(path)
                except:
                    pass
        except:
            traceback.print_exc()
                            
        exec('import '+ modul)
        
        return eval(modul)
    except:
        traceback.print_exc()
        

    
    
