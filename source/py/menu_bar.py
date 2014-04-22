# -*- coding: utf-8 -*-
import uno
import unohelper
import traceback
import sys
import os
import xml.etree.ElementTree as ElementTree
import time
import codecs
import math
import re
import konstanten as KONST
import copy

tb = traceback.print_exc
platform = sys.platform


global oxt
oxt = False


if oxt:
    pyPath = 'E:\\Eclipse_Workspace\\orga\\organon\\py'
    if platform == 'linux':
        pyPath = '/home/xgr/Arbeitsordner/organon/py'
        sys.path.append(pyPath)
    


class Menu_Bar():
    
    def __init__(self,pdk,dialog,ctx,tabs,path_to_extension):
        
        self.pd = pdk
        global pd,IMPORTS
        pd = pdk

        IMPORTS = ('traceback','uno','unohelper','sys','os','ElementTree','time','codecs','math','re','tb','platform','KONST','pd','copy')
        
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
        self.doc = self.get_doc()        
        self.current_Contr = self.doc.CurrentController 
        self.viewcursor = self.current_Contr.ViewCursor
        self.tabs = tabs
        self.platform = sys.platform
        self.language = None
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
        self.kommender_Eintrag = 0
        self.selektierte_zeile = None       # control des Zeilencontainers, Name = ordinal
        self.selektierte_Zeile_alt = None   # control 'textfeld' der Zeile
        self.Papierkorb = None              # ordinal des Papierkorbs - wird anfangs einmal gesetzt und bleibt konstant    
        self.Projektordner = None  
        self.Papierkorb_geleert = False
        self.bereich_wurde_bearbeitet = False
        self.tastatureingabe = False
        self.zuletzt_gedrueckte_taste = None
        
        self.xml_tree = None
        #Settings
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
        bereiche =              self.lade_modul('bereiche')
        self.class_Bereiche =        bereiche.Bereiche(self)
        self.VC_selection_listener = bereiche.ViewCursor_Selection_Listener(self)          
        self.w_listener =            bereiche.Dialog_Window_Listener(self)
        
        #self.doc_listener = Doc_Listener(self)
        
        
        

        # fuers debugging
        self.debug = False
        if self.debug: print('Debug = True')
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
#         if has_prop:    
#             self.projekt_name = UD_properties.getPropertyValue('ProjektName')
#             self.erzeuge_MenuBar_Container()
#             self.erzeuge_Menu()
            #self.class_Projekt.lade_Projekt(False)
           
       
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
        menuB_control, menuB_model = createControl3(self.ctx, "Container", 2, 2, 1000, 20, (), ())          
        menuB_model.BackgroundColor = KONST.Color_MenuBar_Container
         
        self.dialog.addControl('Organon_Menu_Bar', menuB_control)


    def erzeuge_Menu_Kopf_Datei(self,listener):
        control, model = createControl3(self.ctx, "FixedText", 0, 2, 35, 20, (), ())           
        model.Label = self.lang.FILE             
        control.addMouseListener(listener)
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar') 
        MenuBarCont.addControl('Datei', control)
    
    
    def erzeuge_Menu_Kopf_Optionen(self,listener):         
        control, model = createControl3(self.ctx, "FixedText", 37, 2, 55, 20, (), ())           
        model.Label = self.lang.OPTIONS           
        control.addMouseListener(listener)
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar')      
        MenuBarCont.addControl('Optionen', control)
        
        
    def erzeuge_Menu_Kopf_Test(self,listener):
        control, model = createControl3(self.ctx, "FixedText", 300, 2, 50, 20, (), ())           
        model.Label = 'Test'                     
        control.addMouseListener(listener) 
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar')   
        MenuBarCont.addControl('Projekt', control)
        
        
    def erzeuge_Menu_neuer_Ordner(self,listener2):
        control, model = createControl3(self.ctx, "ImageControl", 120, 0, 20, 20, (), ())   
        model.ImageURL = KONST.IMG_ORDNER_NEU_24
        
        model.HelpText = self.lang.INSERT_DIR
        model.Border = 0                    
        control.addMouseListener(listener2) 
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar')     
        MenuBarCont.addControl('Ordner', control)
        
        
    def erzeuge_Menu_Kopf_neues_Dokument(self,listener2):
        control, model = createControl3(self.ctx, "ImageControl", 140, 0, 20, 20, (), ())           
        model.ImageURL = KONST.IMG_DATEI_NEU_24
                    
        model.HelpText = self.lang.INSERT_DOC
        model.Border = 0                    
        control.addMouseListener(listener2) 
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar')   
        MenuBarCont.addControl('neues_Dokument', control)
  
        
    def erzeuge_Menu_Kopf_Papierkorb_leeren(self,listener2):
        control, model = createControl3(self.ctx, "ImageControl", 220, 0, 20, 20, (), ())           
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leeren.png'
        model.HelpText = self.lang.CLEAR_RECYCLE_BIN
        model.Border = 0                       
        control.addMouseListener(listener2) 
        
        MenuBarCont = self.dialog.getControl('Organon_Menu_Bar')     
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
          
        X = ev.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.Source.AccessibleContext.LocationOnScreen.value.Y + ev.Source.Size.value.Height
    
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = KONST.Breite_Menu_DropDown_Container
        oRect.Height = KONST.Hoehe_Menu_DropDown_Container
        
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
        cont_model.BackgroundColor = KONST.MENU_DIALOG_FARBE  # 9225984
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
        control, model = createControl3(self.ctx, "ListBox", 10 ,  10 , 
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
                lang.IMPORT_2)
        
        control.addItems(items, 0)
        model.BackgroundColor = KONST.MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Datei', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)
        
    
    def erzeuge_Menu_DropDown_Eintraege_Optionen(self,window,cont):
        try:
            # Tag1
            control_tag1, model_tag1 = createControl3(self.ctx, "CheckBox", 10, 10, 
                                                      KONST.Breite_Menu_DropDown_Eintraege-6, 30-6, (), ())   
            model_tag1.Label = self.lang.SHOW_TAG1
            
            if self.settings_proj['tag1']:
                model_tag1.State = 1
            else:
                model_tag1.State = 0
                
            tag1_listener = Tag1_Item_Listener(self,model_tag1)
            control_tag1.addItemListener(tag1_listener)
            cont.addControl('Checkbox_Tag1', control_tag1)
            
    
            # ListBox
            control, model = createControl3(self.ctx, "ListBox", 10, 34, KONST.Breite_Menu_DropDown_Eintraege-6, 
                                            KONST.Hoehe_Menu_DropDown_Eintraege - 30, (), ())   
            control.setMultipleMode(False)
            
            items = ( self.lang.UNFOLD_PROJ_DIR, 
                     '-------', 
                     '#Homepage', 
                     '#Updates', 
                     '#Etc.')
            
            control.addItems(items, 0)
            model.BackgroundColor = KONST.MENU_DIALOG_FARBE
            model.Border = False
            
            cont.addControl('Eintraege_Optionen', control)
            
            listener = DropDown_Item_Listener(self)  
            listener.window = window    
            control.addItemListener(listener)  
        except:
            tb() 
    
    
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
            
            for imp in IMPORTS:
                exec('h_feld.%s=%s' %(imp,imp))
        else: 
            import h_feld   
            
            for imp in IMPORTS:
                exec('h_feld.%s=%s' %(imp,imp))
                   
        Klasse_Hauptfeld = h_feld.Main_Container(self)
        Klasse_Zeilen_Listener = h_feld.Zeilen_Listener(self.Hauptfeld,self.ctx,self)
        return Klasse_Hauptfeld,Klasse_Zeilen_Listener

    def lade_modul(self,modul,arg = None): 
        
        try: 
            if oxt:
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
    
    
    def erzeuge_neue_Projekte(self):
        try:
            self.class_Projekt.test()
        except:
            traceback.print_exc()
            
    def erzeuge_Zeile(self,ordner_oder_datei):
        self.class_Hauptfeld.erzeuge_neue_Zeile(ordner_oder_datei)          
            
    def leere_Papierkorb(self):
        self.class_Hauptfeld.leere_Papierkorb()                  

    def debug_time(self):
        zeit = "%0.2f" %(self.time.clock()-self.timer_start)
        return zeit

    def entferne_alle_listener(self):
        return
        self.current_Contr.removeSelectionChangeListener(self.VC_selection_listener) 
        self.current_Contr.removeKeyHandler(self.keyhandler)
        self.dialog.removeWindowListener(self.w_listener)
        
    def erzeuge_Dialog_Container(self,posSize):
        
        EXPORT_DIALOG_FARBE = 305099

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
        oWindowDesc.WindowAttributes = 1  +32 +64 + 128 # Flags fuer com.sun.star.awt.WindowAttribute
    
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
        #oFrame.setCreator(self.mb.desktop)
        oFrame.activate()
    
        # create new control container
        cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
        cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
        cont_model.BackgroundColor = EXPORT_DIALOG_FARBE  # 9225984
        cont.setModel(cont_model)
        # need createPeer just only the container
        cont.createPeer(toolkit, oWindow)
        cont.setPosSize(0, 0, 0, 0, 15)
    
        oFrame.setComponent(cont, None)
        return oWindow,cont
     
    def loesche_undo_Aktionen(self):
        undoMgr = self.doc.UndoManager
        undoMgr.reset()
        
    def speicher_settings(self,dateiname,eintraege):
        
        path = os.path.join(self.pfade['settings'],dateiname)
        imp = str(eintraege).replace(',',',\n')
            
        with open(path , "w") as file:
            file.writelines(imp)
   


# from com.sun.star.document import XDocumentEventListener
# class Doc_Listener(unohelper.Base,XDocumentEventListener):
#     def __init__(self,mb):
#         self.mb = mb
#     def documentEventOccured(self,ev):
#         
#         
#         print('documentEventOccured')
#         print(ev.EventName)
#         return
#         def pydevBrk():  
#             # adjust your path 
#             print('hier')
#             import sys
#             sys.path.append(r'C:\Users\Homer\Desktop\Programme\eclipse\plugins\org.python.pydev_3.1.0.201312121632\pysrc')  
#             from pydevd import settrace
#             settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True)
#         
#         #pydevBrk()
#         
#         if ev.EventName == 'OnUnfocus':
#             print(self.mb.Anzahl_Componenten)
#             import traceback
#             
#             try:
#                 enum = self.mb.desktop.Components.createEnumeration()
#                 comps = []
#                 while enum.hasMoreElements():
#                     comps.append(enum.nextElement())
#                 Anzahl_Componenten = len(comps)
#                 if Anzahl_Componenten > self.mb.Anzahl_Componenten:
#                     
#                     
#                     from com.sun.star.awt import Rectangle,WindowDescriptor 
#                     from com.sun.star.awt.WindowClass import MODALTOP
#                     from com.sun.star.awt.VclWindowPeerAttribute import OK,YES_NO_CANCEL, DEF_NO
#                     
#                     global Rectangle,WindowDescriptor,MODALTOP,OK,YES_NO_CANCEL, DEF_NO
#                     
#                     self.mb.Mitteilungen.nachricht(u'Sie haben eine weitere Instanz von OO/LO geöffnet. Organon funktioniert nun nicht mehr. Wenn Sie mehrere Instanzen von OO/LO nutzen möchten, öffnen Sie alle Instanzen und laden als letzte die Organon Datei.',"warningbox")
#                     pydevBrk()
#             except:
#                 
#                 
#                 
#                 
#                 print('fehler')
#                 traceback.print_exc()
#             
#     def disposing(self,ev):      
#         return False


from com.sun.star.awt import XMouseListener, XItemListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT 
    
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
            elif self.menu_Kopf_Eintrag == self.mb.lang.OPTIONS:
                self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)
            elif self.menu_Kopf_Eintrag == 'Test':
                self.mb.erzeuge_neue_Projekte()
                
            self.mb.loesche_undo_Aktionen()
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
        sel = ev.value.Source.Items[ev.value.Selected]

        if sel == self.mb.lang.NEW_PROJECT:
            self.do()
            self.mb.class_Projekt.erzeuge_neues_Projekt()
        elif sel == self.mb.lang.OPEN_PROJECT:
            self.do()
            self.mb.class_Projekt.lade_Projekt()
        elif sel == self.mb.lang.NEW_DOC:
            self.do()
            self.mb.erzeuge_Zeile('dokument')
        elif sel == self.mb.lang.NEW_DIR:
            self.do()
            self.mb.erzeuge_Zeile('Ordner')
        elif sel == self.mb.lang.EXPORT_2:
            self.do()
            self.mb.class_Export.export()
        elif sel == self.mb.lang.IMPORT_2:
            self.do()
            self.mb.class_Import.importX()
        elif sel == self.mb.lang.UNFOLD_PROJ_DIR:
            self.do()
            self.mb.class_Funktionen.projektordner_ausklappen()
        
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
            set = self.mb.settings_proj
            
            if self.model.State == 1:
                set['tag1'] = 1
            else:
                set['tag1'] = 0
            
            if not set['tag1']:
                set['tag1'] = 0
                self.mache_tag1_sichtbar(False)
            else:
                set['tag1'] = 1
                self.mache_tag1_sichtbar(True) 
            
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        except:
            tb()
    
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

                control_tag1, model_tag1 = createControl3(self.mb.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
                model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
                model_tag1.Border = 0
                control_tag1.addMouseListener(self.mb.class_Hauptfeld.tag1_listener)

                contr_zeile.addControl('tag1',control_tag1)

    
    

from com.sun.star.awt import XKeyHandler
class Key_Handler(unohelper.Base, XKeyHandler):
    
    def __init__(self,mb):
        #if oxt:print('init Keyhandler')
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
            traceback.print_exc()
                            
        exec('import '+ modul)

        return eval(modul)
    except:
        traceback.print_exc()
        

    
    
