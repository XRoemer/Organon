# -*- coding: utf-8 -*-
import uno
import unohelper

from traceback import format_exc as tb
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
from pprint import pformat
import webbrowser
import json

from tabs import TabsX
platform = sys.platform



class Menu_Bar():
    
    def __init__(self,args,tab = 'ORGANON'):
        
        try:
            
            (pdk,
             dialog,
             ctx,
             path_to_extension,
             win,
             dict_sb,
             debugX,
             factory,
             menu_start,
             logX,
             class_LogX,
             settings_organon,
             ) = args

            
            ###### DEBUGGING ########
            global debug,log,load_reload
            debug = debugX
            log = logX
            
            self.debug = debugX
            if self.debug: 
                self.time = time
                self.timer_start = self.time.clock()
            
            # Wird beim Debugging auf True gesetzt    
            load_reload = sys.dont_write_bytecode
    
            if self.debug: log(inspect.stack)
            
                
            ###### DEBUGGING END ########
            
            
            
            self.win = win
            self.pd = pdk
            global pd,IMPORTS,LANG
            pd = pdk
            
            global T,prj_tab
            T = Tab()
            
            # Konstanten
            self.factory = factory
            self.dialog = dialog
            self.ctx = ctx
            self.smgr = self.ctx.ServiceManager
            self.toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
            self.topWindow = self.toolkit.getActiveTopWindow() 
            self.desktop = self.smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
            self.doc = self.get_doc()  
            self.current_Contr = self.doc.CurrentController 
            self.programm = self.get_office_name() 
            self.undo_mgr = self.doc.UndoManager
            self.viewcursor = self.current_Contr.ViewCursor
            self.platform = sys.platform
            self.language = None
            LANG = self.lade_Modul_Language()
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
            self.mb_hoehe = 0
            
            
            
            # Properties
            self.props = {}
            self.props.update({T.AB :Props()})
            self.dict_sb = dict_sb              # drei Unterdicts: sichtbare, eintraege, controls
            self.tags = None
            self.registrierte_maus_listener = []
            self.maus_fenster = None
            self.mausrad_an = False
            self.texttools_geoeffnet = False
            self.writer_vorlagen = {}

            # Settings
            self.settings_orga = settings_organon
            self.settings_exp = None
            self.settings_imp = None
            self.settings_proj = {}
            
    
            # Pfade
            self.pfade = {}
            
            IMPORTS = {'uno':uno,
                       'unohelper':unohelper,
                       'sys':sys,
                       'os':os,
                       'ElementTree':ElementTree,
                       'time':time,
                       'codecs_open':codecs_open,
                       'math_floor':math_floor,
                       're':re,
                       'tb':tb,
                       'platform':platform,
                       'KONST':KONST,
                       'pd':pd,
                       'copy':copy,
                       'Props':Props,
                       'T':T,
                       'log':log,
                       'inspect':inspect,
                       'webbrowser':webbrowser,
                       'LANG':LANG}
            
            
            # Klassen   
            self.ET = ElementTree  
            
            self.nachricht = Mitteilungen(self.ctx,self).nachricht
            self.nachricht_si = Mitteilungen(self.ctx,self).nachricht_si
            self.popup = Mitteilungen(self.ctx,self).popup
            
            self.class_Baumansicht,self.class_Zeilen_Listener = self.get_Klasse_Baumansicht()
            self.class_Projekt =        self.lade_modul('projects','Projekt')   
            self.class_XML =            self.lade_modul('xml_m','XML_Methoden')
            self.class_Funktionen =     self.lade_modul('funktionen','Funktionen')     
            self.class_Export =         self.lade_modul('export','Export')
            self.class_Import =         self.lade_modul('importX','ImportX') 
            self.class_Sidebar =        self.lade_modul('sidebar','Sidebar') 
            self.class_Tags =           self.lade_modul('tags','Tags')
            self.class_Bereiche =       self.lade_modul('bereiche','Bereiche')
            self.class_Version =        self.lade_modul('version','Version') 
            self.class_Tabs =           self.lade_modul('tabs','Tabs') 
            self.class_Latex =          self.lade_modul('latex_export','ExportToLatex') 
            self.class_Html =           self.lade_modul('export2html','ExportToHtml') 
            self.class_Zitate =         self.lade_modul('zitate','Zitate') 
            self.class_werkzeug_wListe= self.lade_modul('werkzeug_wListe','WListe') 
            self.class_Index =          self.lade_modul('index','Index')
            self.class_Mausrad =        self.lade_modul('mausrad','Mausrad')
            self.class_Einstellungen =  self.lade_modul('einstellungen','Einstellungen')
            self.class_Organon_Design = self.lade_modul('design','Organon_Design')
            self.class_Organizer =      self.lade_modul('organizer','Organizer')
            self.class_Shortcuts =      self.lade_modul('shortcuts','Shortcuts')
            self.class_Fenster =        self.lade_modul('fenster','Fenster')
            
            
            # TABS
            self.tabsX = self.lade_modul('tabs','TabsX')
            self.tabsX.run(self.win)
            self.prj_tab = self.tabsX.erzeuge_neuen_tab('ORGANON')
            self.dialog.addControl('tableiste',self.tabsX.tableiste)
            
            
            # Plattformabhaengig
            if self.platform == 'win32':
                self.class_RawInputReader = self.lade_modul('rawinputdata','RawInputReader')
            
            self.class_Log = class_LogX
            self.class_Design = Design()
            self.class_Gliederung = Gliederung()
            
            #self.class_Greek =      self.lade_modul('greek2latex','.Greek(self,pd)')

            
            # Listener  
            self.use_UM_Listener = False

            self.Listener = Listener(self)  
  
            self.class_Funktionen.update_organon_templates()
            
        except:
            log(inspect.stack,tb())
            # fehlt:
            # Den User ueber den Fehler benachrichtigen
            
            
  

    def get_doc(self):
        if self.debug: log(inspect.stack)
        
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
    
    def get_office_name(self):
        if self.debug: log(inspect.stack)
        
        frame = self.current_Contr.Frame
        if 'LibreOffice' in frame.Title:
            programm = 'LibreOffice'
        elif 'OpenOffice' in frame.Title:
            programm = 'OpenOffice'
        else:
            # Fuer Linux / OSX fehlt
            programm = 'LibreOffice'
        
        return programm
    
    def get_BEREICH_EINFUEGEN(self):
        if self.debug: log(inspect.stack)
        
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
        if self.debug: log(inspect.stack)
        
        pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        for ext in pip.ExtensionList:
            if ext[0] == 'xaver.roemers.organon':
                version = ext[1]
                
        return version
    
    def erzeuge_Menu(self,win,tab=False):
        if self.debug: log(inspect.stack)
        
        try:             
            listener = Menu_Leiste_Listener(self) 
            listener2 = Menu_Leiste_Listener2(self) 
            
            # CONTAINER
            menuB_control, menuB_model = self.createControl(self.ctx, "Container", 2, 2, 1000, 20, (), ())          
            menuB_model.BackgroundColor = KONST.FARBE_MENU_HINTERGRUND
    
            win.addControl('Organon_Menu_Bar', menuB_control)
            win.setPosSize(0,self.tabsX.tableiste_hoehe,0,0,2)
            bereich = self.props[T.AB].selektierte_zeile_alt
            
            if tab:
                Menueintraege = [
                    (LANG.FILE,'a'),
                    (LANG.BEARBEITEN_M,'a'),
                    (LANG.ANSICHT,'a'),            
                    ('Speichern','b','vnd.sun.star.extension://xaver.roemers.organon/img/lc_save.png',
                     LANG.FORMATIERUNG_SPEICHERN.format(LANG.KEINE))
                    ]
            else:
                
                Menueintraege = [
                    (LANG.FILE,'a'),
                    (LANG.BEARBEITEN_M,'a'),
                    (LANG.ANSICHT,'a'),            
                    ('Ordner','b',KONST.IMG_ORDNER_NEU_24,LANG.INSERT_DIR),
                    ('Datei','b',KONST.IMG_DATEI_NEU_24,LANG.INSERT_DOC),
                    ('Speichern','b','vnd.sun.star.extension://xaver.roemers.organon/img/lc_save.png',
                     LANG.FORMATIERUNG_SPEICHERN.format(LANG.KEINE))
                    ]
            
            x = 0
            
            # SPRACHE
            for eintrag in Menueintraege:
                if eintrag[1] == 'a':
            
                    control, model = self.createControl(self.ctx, "FixedText", x, 2, 0, 20, (), ())           
                    model.Label = eintrag[0]  
                    model.TextColor = KONST.FARBE_MENU_SCHRIFT
                    control.addMouseListener(listener)
                    breite = control.getPreferredSize().Width
                    control.setPosSize(0,0,breite,0,4)
                    
                    menuB_control.addControl(eintrag[0], control)
                    
                    x += breite + 5
                 
            x += 15

            # ICONS
            for eintrag in Menueintraege:
                if eintrag[1] == 'b':  
                    
                    h = 0
                    
                    if T.AB != 'ORGANON':
                        if eintrag[0] in ('Ordner','Datei'):
                            x += 22
                            continue 
                        
                    if eintrag[0] == 'Speichern':
                        x += 20
                        h = -2
                        
                    control, model = self.createControl(self.ctx, "ImageControl", x, 0 - h * .5, 20 + h, 20 + h, (), ())           
                    model.ImageURL = eintrag[2]
                    model.HelpText = eintrag[3]
                    model.Border = 0                    
                    control.addMouseListener(listener2) 
                    
                    menuB_control.addControl(eintrag[0], control)  
                    
                    x += 22
            # TEST
            if load_reload:
                control, model = self.createControl(self.ctx, "FixedText", x+30, 2, 50, 20, (), ())           
                model.Label = 'Test'  
                model.TextColor = KONST.FARBE_MENU_SCHRIFT
                control.addMouseListener(listener)
                
                menuB_control.addControl('Test', control)

        except Exception as e:
                self.nachricht('erzeuge_Menu ' + str(e),"warningbox")
                log(inspect.stack,tb())
                
        
        
        
     
    def erzeuge_Menu_DropDown_Eintraege(self,items):
        if self.debug: log(inspect.stack)

        controls = []
        SEPs = []
        
        xBreite = 0
        y = 10
        
        listener = Auswahl_Menu_Eintrag_Listener(self)
        
        for item in items:
            
            if item != 'SEP':
                prop_names = ('Label',)
                prop_values = (item,)
                control, model = self.createControl(self.ctx, "FixedText", 30, y, 50,50, prop_names, prop_values)

                prefSize = control.getPreferredSize()
     
                Hoehe = prefSize.Height 
                Breite = prefSize.Width + 20
                control.setPosSize(0,0,Breite ,Hoehe,12)
                
                control.addMouseListener(listener)
                
                if item == LANG.TEXTTOOLS:
                    prop_names = ('ImageURL','Border')
                    prop_values = ('vnd.sun.star.extension://xaver.roemers.organon/img/pfeil2.png',0)
                    controlTexttools, modelT = self.createControl(self.ctx, "ImageControl", Breite , y+5, 4,6, prop_names, prop_values)
                    controls.append(controlTexttools)
                
                if xBreite < Breite:
                    xBreite = Breite
                    
                y += Hoehe
                
            else:
                
                # Waagerechter Trenner
                control, model = self.createControl(self.ctx, "FixedLine", 30, y, 50,10,(),())
                model.TextColor = 0
                model.TextLineColor = 0
                SEPs.append(control)
                y += 10

            controls.append(control)
 
        
        # Senkrechter Trenner
        control, model = self.createControl(self.ctx, "FixedLine", 20, 10, 5,y -10 ,(),())
        controls.append(control)
        model.Orientation = 1

        for sep in SEPs:
            sep.setPosSize(0,0,xBreite ,0,4)
        
        try:
            controlTexttools.setPosSize(xBreite + 30,0,0 ,0,1)
        except:
            pass
        
        return controls,listener,y - 5,xBreite +20


    def erzeuge_Menu_DropDown_Eintraege_Ansicht(self,items):
        if self.debug: log(inspect.stack)
        try:
            controls = []
            SEPs = []

            if self.projekt_name != None:
                tag1 = self.settings_proj['tag1']
                tag2 = self.settings_proj['tag2']
                tag3 = self.settings_proj['tag3']
            else:
                tag1 = 0
                tag2 = 0
                tag3 = 0
        
            xBreite = 0
            y = 10
            
            listener = Auswahl_Menu_Eintrag_Listener(self)
            tag_TV_listener = DropDown_Tags_TV_Listener(self)
            tag_SB_listener = DropDown_Tags_SB_Listener(self)
            

            for item in items:

                if item[0] != 'SEP':
                    prop_names = ('Label',)
                    prop_values = (item[0],)
                    control, model = self.createControl(self.ctx, "FixedText", 30, y, 50,50, prop_names, prop_values)
                    # Image
                    control_I, model_I = self.createControl(self.ctx, "ImageControl", 5, y, 16,16, (), ())
                    model_I.Border = 0
                    
                    
                    
                    if item[1] == 'Ueberschrift':
                        model.FontWeight = 150
                        
                        
                    elif item[1] == '' and item[0] != 'SEP':
                        control.addMouseListener(listener)
                        
                        
                    elif item[1] == 'Tag_TV':
                        
                        control.addMouseListener(tag_TV_listener)

                        if self.projekt_name != None:
                            sett = self.settings_proj
                            tag1,tag2,tag3 = sett['tag1'],sett['tag2'],sett['tag3']
                        else:
                            tag1,tag2,tag3 = 0,0,0

                        if item[0] == LANG.SHOW_TAG1 and tag1 == 1:
                            model_I.ImageURL = 'private:graphicrepository/svx/res/apply.png'
                        elif item[0] == LANG.SHOW_TAG2 and tag2 == 1:
                            model_I.ImageURL = 'private:graphicrepository/svx/res/apply.png'
                        elif item[0] == LANG.GLIEDERUNG and tag3 == 1:
                            model_I.ImageURL = 'private:graphicrepository/svx/res/apply.png'
                            
                            
                    elif item[1] == 'Tag_SB':
                        
                        control.addMouseListener(tag_SB_listener)
                        panel_nr = self.tags['name_nr'][item[0]]
                        
                        if panel_nr in self.tags['sichtbare']:
                            model_I.ImageURL = 'private:graphicrepository/svx/res/apply.png' 
                        
                    prefSize = control.getPreferredSize()
         
                    Hoehe = prefSize.Height 
                    Breite = prefSize.Width + 15
                    control.setPosSize(0,0,Breite ,Hoehe,12)

                    if xBreite < Breite:
                        xBreite = Breite
                        
                    y += Hoehe
                    
                else:
                    
                    # Waagerechter Trenner
                    control, model = self.createControl(self.ctx, "FixedLine", 30, y, 50,10,(),())
                    model.setPropertyValue('TextColor',102)
                    model.setPropertyValue('TextLineColor',102)
                    SEPs.append(control)
                    y += 10

                controls.append((control,item[0]))
                controls.append((control_I,item[0]+'_icon'))
            
            # Senkrechter Trenner
            control, model = self.createControl(self.ctx, "FixedLine", 20, 10, 5,y -10 ,(),())
            controls.append((control,''))
            model.Orientation = 1
    
            for sep in SEPs:
                sep.setPosSize(0,0,xBreite ,0,4)
    
            return controls,listener,y - 5,xBreite +20
        except:
            log(inspect.stack,tb())
     
    
    def get_speicherort(self):
        if self.debug: log(inspect.stack)
        
        pfad = os.path.join(self.path_to_extension,'pfade.txt')
        
        if os.path.exists(pfad):            
            with codecs_open( pfad, "r","utf-8") as file:
                filepath = file.read() 
            return filepath
        else:
            return None
            
    def get_Klasse_Baumansicht(self):
        if self.debug: log(inspect.stack)

        import baum   
        
        for imp in IMPORTS:
            setattr(baum,imp,IMPORTS[imp])
                   
        Klasse_Hauptfeld = baum.Baumansicht(self)
        Klasse_Zeilen_Listener = baum.Zeilen_Listener(self.ctx,self)
        return Klasse_Hauptfeld,Klasse_Zeilen_Listener

    def lade_modul(self,modul,arg = None): 
        if self.debug: log(inspect.stack)
        
        try: 
                
            mod = __import__(modul)
            
            for imp in IMPORTS:
                setattr(mod,imp,IMPORTS[imp])

            if arg == None:
                return mod
            else:                    
                oClass = getattr(mod, arg)
                return oClass(self)
        except:
            log(inspect.stack,tb())
     
  
    def lade_Modul_Language(self):
        if self.debug: log(inspect.stack)
   
        config_provider = self.smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider",self.ctx)
        
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = "nodepath"
        prop.Value = "org.openoffice.Setup/L10N"
               
        config_access = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
        locale = config_access.getByName('ooLocale')
        
        lang = __import__('lang_en')
        
        # Wenn die Uebersetzung nicht existiert und logging eingeschaltet ist,
        # wird der versuchte Import als Fehler geloggt.
        # aendern
        
        try:
            loc = locale[0:2]
            self.language = loc
            try:
                lang2 = __import__('lang_{}'.format(loc))
                attributes = [a for a in dir(lang2) if '__' not in a]
                for l in attributes:
                    try:
                        try:
                            attr = getattr(lang2, l)
                        except:
                            attr = ''
                        if attr.strip() not in ('','*******',None):
                            setattr(lang, l, attr)
                    except:
                        if self.debug:log(inspect.stack,tb())
                        
            except:
                if self.debug:log(inspect.stack,tb())

        except:
            if self.debug:log(inspect.stack,tb())
        
        return lang
    
    def lade_RawInputReader(self):
        if self.debug: log(inspect.stack)
        
        if self.platform != 'win32':
            return None
        
        import rawinputdata
        
        if load_reload:
            # reload laedt nur rawinputdata. Aenderungen in
            # RawInputReader werden nicht neu geladen.
            # 
            # Die Methode lade_modul funktionierte ebenfalls nicht,
            # da bei erneutem Oeffnen von Organon die globalen Variablen
            # alle auf None gesetzt sind.
            #
            # Keine Idee fuer eine Loesung bis jetzt
            if self.programm == 'OpenOffice':
                reload(rawinputdata)
            
        return rawinputdata.RawInputReader
    
    
    def Test(self):
        try:
            self.class_Projekt.test()
        except:
            log(inspect.stack,tb())
            
            
    def leere_Papierkorb(self):
        if self.debug: log(inspect.stack)
        
        self.class_Baumansicht.leere_Papierkorb()   
        
    def erzeuge_Backup(self):
        if self.debug: log(inspect.stack)
        
        try:
            pfad_zu_backup_ordner = os.path.join(self.pfade['projekt'],'Backups')
            if not os.path.exists(pfad_zu_backup_ordner):
                os.makedirs(pfad_zu_backup_ordner)
            
            lt = time.localtime()
            t = time.strftime(" %d.%m.%Y  %H.%M.%S", lt)
            
            neuer_projekt_name = self.projekt_name + t
            pfad_zu_neuem_ordner = os.path.join(pfad_zu_backup_ordner,neuer_projekt_name)
            
            tree = copy.deepcopy(self.props['ORGANON'].xml_tree)
            root = tree.getroot()
            
            all_elements = root.findall('.//')
            ordinale = []
            
            for el in all_elements:
                ordinale.append(el.tag)        
        
            self.class_Export.kopiere_projekt(neuer_projekt_name,pfad_zu_neuem_ordner,
                                                 ordinale,tree,self.tags,True)  
            os.rename(pfad_zu_neuem_ordner,pfad_zu_neuem_ordner+'.organon')  
            self.nachricht('Backup erzeugt unter: %s' %pfad_zu_neuem_ordner+'.organon', "infobox")           
        except:
            log(inspect.stack,tb())
            
          
    def debug_time(self):
        return "%0.2f" %(self.time.clock()-self.timer_start)

     
    def loesche_undo_Aktionen(self):
        if self.debug: log(inspect.stack)
        
        undoMgr = self.doc.UndoManager
        undoMgr.reset()
        
    def speicher_settings(self,dateiname,eintraege):
        if self.debug: log(inspect.stack)
        
        path = os.path.join(self.pfade['settings'],dateiname)
        imp = pformat(eintraege)

        with codecs_open(path , "w","utf-8") as file:
            file.write(imp)

    def schreibe_settings_orga(self):
        if self.debug: log(inspect.stack)
        
        path = os.path.join(self.path_to_extension,"organon_settings.json")
        
        with open(path, 'w') as outfile:
            json.dump(self.settings_orga, outfile,indent=4, separators=(',', ': '))
                
    def pruefe(self):
        
        tree = self.props[T.AB].xml_tree
        root = tree.getroot()
        all = root.findall('.//')
        pars = [a.attrib['Parent'] for a in all]
        if 'Projekt' in pars:
            print('Fehler')
    
    def tree_write(self,tree,pfad):  
        if self.debug: log(inspect.stack) 
        # diese Methode existiert, um alle Schreibvorgaenge
        # des XML_trees bei Bedarf kontrollieren zu koennen
        #self.pruefe()
        tree.write(pfad)
    
    def prettyprint(self,pfad,oObject,w=600):
        
        from pprint import pformat
        imp = pformat(oObject,width=w)
        with codecs_open(pfad , "w",'utf-8') as file:
            file.write(imp)
            
    
    def kalkuliere_und_setze_Control(self,ctrl,h_or_w = None):
        #if self.debug: log(inspect.stack)
        
        prefSize = ctrl.getPreferredSize()
        Hoehe = prefSize.Height 
        Breite = prefSize.Width #+10
        
        if h_or_w == None:
            ctrl.setPosSize(0,0,Breite,Hoehe,12)
        elif h_or_w == 'w_pruefen':
            if Breite > ctrl.Size.Width:
                ctrl.setPosSize(0,0,Breite,0,4)
        elif h_or_w == 'h':
            ctrl.setPosSize(0,0,0,Hoehe,8)
        elif h_or_w == 'w':
            ctrl.setPosSize(0,0,Breite,0,4)
            
        return Breite,Hoehe
        
   
    # Handy function provided by hanya (from the OOo forums) to create a control, model.
    def createControl(self,ctx,type,x,y,width,height,names,values):
        try:
            ctrl = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
            ctrl_model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
            ctrl_model.setPropertyValues(names,values)
            ctrl.setModel(ctrl_model)
            ctrl.setPosSize(x,y,width,height,15)
            #ctrl_model.BackgroundColor = 501
            return (ctrl, ctrl_model)
        except:
            log(inspect.stack,tb())
            
    
    def createUnoService(self,serviceName):
        sm = uno.getComponentContext().ServiceManager
        return sm.createInstanceWithContext(serviceName, uno.getComponentContext())
    
    def dispatch(self,frame,cmd,oprop=('',None),oprop2=('',None)):
        if self.debug: log(inspect.stack)
              
        sm = uno.getComponentContext().ServiceManager
        dispatcher = sm.createInstanceWithContext("com.sun.star.frame.DispatchHelper", uno.getComponentContext())
        
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = oprop[0]
        prop.Value = oprop[1]
        
        prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop2.Name = oprop2[0]
        prop2.Value = oprop2[1]
        
        res = dispatcher.executeDispatch(frame, ".uno:{}".format(cmd), "", 0, (prop,prop2))
        return res
    
    def erzeuge_texttools_fenster(self,ev,m_win):
        if self.debug: log(inspect.stack)
        
        loc_menu = ev.Source.Context.AccessibleContext.LocationOnScreen
        loc_cont = self.current_Contr.Frame.ContainerWindow.AccessibleContext.LocationOnScreen
        
        x3 = loc_menu.X - loc_cont.X    # Position des Dropdown Menus
        y3 = loc_menu.Y - loc_cont.Y    # Position des Dropdown Menus
        
        x3 += ev.Source.Context.PosSize.Width
        y3 += ev.Source.PosSize.Y
        
        items = self.menuEintraege(LANG,LANG.TEXTTOOLS)
        
        controls,listener,Hoehe,Breite = self.erzeuge_Menu_DropDown_Eintraege(items)
        controls = list((x,'') for x in controls)
        
        posSize = x3+2,y3,Breite +20,Hoehe +20
        win,cont = self.class_Fenster.erzeuge_Dialog_Container(posSize,Flags=1+512)
        
        # Listener fuers Dispose des Fensters
        listener2 = Schliesse_Menu_Listener(self,texttools=True)
        cont.addMouseListener(listener2) 
        listener2.ob = win
        
        listener.window = win
        listener.texttools = True
        
        listener.win2 = m_win
        
        for c in controls:
            cont.addControl(c[1],c[0])
            
            
    def menuEintraege(self,LANG,menu):
    
        if menu == LANG.FILE:
            
            items = (LANG.NEW_PROJECT, 
                    LANG.OPEN_PROJECT ,
                    'SEP', 
                    LANG.NEW_DOC, 
                    LANG.NEW_DIR,
                    'SEP',
                    LANG.EXPORT_2, 
                    LANG.IMPORT_2,
                    'SEP',
                    LANG.BACKUP,
                    LANG.EINSTELLUNGEN)
            
            if T.AB != 'ORGANON':
                items = (
                    LANG.EXPORT_2, 
                    'SEP',
                    LANG.BACKUP,
                    LANG.EINSTELLUNGEN)
                
        elif menu == LANG.BEARBEITEN_M:
            items = ( 
                LANG.ORGANIZER,
                LANG.NEUER_TAB,
                'SEP',
                LANG.TRENNE_TEXT,
                LANG.TEXT_BATCH_DEVIDE,
                LANG.DATEIEN_VEREINEN,
                'SEP',
                LANG.TEXTTOOLS,
                'SEP',
                LANG.UNFOLD_PROJ_DIR,
                LANG.CLEAR_RECYCLE_BIN
                )
            if T.AB != 'ORGANON':
                items = (  
                    LANG.ORGANIZER,
                    'SEP',
                    LANG.NEUER_TAB,
                    LANG.SCHLIESSE_TAB,
                    LANG.IMPORTIERE_IN_TAB,
                    'SEP',
                    #LANG.TEXTVERGLEICH,
                    #LANG.WOERTERLISTE,
                    LANG.TEXTTOOLS,
                    'SEP',
                    LANG.UNFOLD_PROJ_DIR,
                    LANG.CLEAR_RECYCLE_BIN
                    )  
                 
        elif menu == LANG.ANSICHT:
            
            tags = self.tags
            panels = [(tags['nr_name'][i][0],'Tag_SB') for i in tags['abfolge']]
            
            items = [
                (LANG.SICHTBARE_TAGS_BAUMANSICHT,'Ueberschrift'),
                (LANG.SHOW_TAG1,'Tag_TV'),
                (LANG.SHOW_TAG2,'Tag_TV'),
                (LANG.GLIEDERUNG,'Tag_TV'),
                ('SEP',''),
                (LANG.SICHTBARE_TAGS_SEITENLEISTE,'Ueberschrift')
                ] + panels + [
                ('SEP',''),
                (LANG.ZEIGE_TEXTBEREICHE,''),
                ('SEP',''),
                (LANG.HOMEPAGE,''),
                (LANG.FEEDBACK,''),
                 ]
            
        elif menu == LANG.TEXTTOOLS:
            items = (LANG.TEXTVERGLEICH,
                    LANG.WOERTERLISTE,
                    LANG.ERZEUGE_INDEX
                     )
          
        return items
        

       

class Props():
    def __init__(self):
        if debug: log(inspect.stack)
        
        self.dict_zeilen_posY = {}
        self.dict_posY_ctrl = {}
        self.dict_ordner = {}               # enthaelt alle Ordner und alle ihnen untergeordneten Zeilen
        self.dict_bereiche = {}             # drei Unterdicts: Bereichsname,ordinal,Bereichsname-ordinal
        
        self.Hauptfeld = None               # alle Zeilen, Controls
        self.sichtbare_bereiche = []        # Bereichsname ('OrganonSec'+ nr)
        self.kommender_Eintrag = 0
        self.selektierte_zeile = None       # ordinal des Zeilencontainers
        self.selektierte_zeile_alt = None   # control 'textfeld' der Zeile
        self.Papierkorb = None              # ordinal des Papierkorbs - wird anfangs einmal gesetzt und bleibt konstant    
        self.Projektordner = None           # ordinal des Projektordners - wird anfangs einmal gesetzt und bleibt konstant 
        self.Papierkorb_geleert = False
        self.tastatureingabe = False
        self.zuletzt_gedrueckte_taste = None

        self.xml_tree = None
        
        self.tab_auswahl = Tab_Auswahl()
        
class Listener():
    ''' Verwaltet alle Listener, die sich auf das Projekt und 
    das Hauptfenster beziehen
    '''
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
        # Listener  
        self.VC_selection_listener  = ViewCursor_Selection_Listener(self.mb)  
        self.w_listener             = Dialog_Window_Size_Listener(self.mb)    
        self.undo_mgr_listener      = Undo_Manager_Listener(self.mb)
        self.listener_doc_close     = Document_Close_Listener(self.mb)
        self.keyhandler             = Key_Handler(self.mb)  
        
        self.blocked = False

        self.states = {f.replace('remove_',''):False for f in dir(self) if 'remove' in f}
            
        
        
    def alle_listener_ausschalten(self,ausnahmen = []):
        if self.mb.debug: log(inspect.stack,extras=listener)
        
        
    def block(self):  
        ''' blockiert alle Listener''' 
        if self.mb.debug: log(inspect.stack)

        for key in self.states:
            if self.states[key] == True:
                fkt = getattr(self, 'remove_' + key)
                fkt()
                
        self.blocked = True
                
    def unblock(self):   
        ''' schaltet Blockade ab und die zuvor aktivierten Listener wieder an''' 
        if self.mb.debug: log(inspect.stack)
        
        self.blocked = False
        
        for key in self.states:
            if self.states[key] == True:
                fkt = getattr(self, 'add_' + key)
                fkt()
                
    def entferne_alle_Listener(self):
        if self.mb.debug: log(inspect.stack)
        for key in self.states:
            try:
                if self.states[key] == True:
                    fkt = getattr(self, 'remove_' + key)
                    fkt()
                    self.states[key] = False
                    
                
                self.mb.current_Contr.removeKeyHandler(self.keyhandler)
                
            except:
                pass
                
    def starte_alle_Listener(self):
        if self.mb.debug: log(inspect.stack)

        for key in self.states:
            if self.states[key] == False:
                fkt = getattr(self, 'add_' + key)
                fkt()
                self.states[key] = True
        
    def versuchter_start(self,listener):
        if self.mb.debug: log(inspect.stack,extras=listener)
          
    
    def add_VC_selection_listener(self):
        if self.blocked : return
        if not self.states['VC_selection_listener']:
            if self.mb.debug: log(inspect.stack)
            self.mb.current_Contr.addSelectionChangeListener(self.VC_selection_listener)
            self.states['VC_selection_listener'] = True
        else:
            self.versuchter_start('VC_selection_listener')
    
    def remove_VC_selection_listener(self):
        if self.blocked : return
        if self.mb.debug: log(inspect.stack)
        self.mb.current_Contr.removeSelectionChangeListener(self.VC_selection_listener)
        self.states['VC_selection_listener'] = False
        
    def add_Dialog_Window_Size_Listener(self):
        if self.blocked : return
        if not self.states['Dialog_Window_Size_Listener']:
            if self.mb.debug: log(inspect.stack)
            self.mb.win.addWindowListener(self.w_listener)
            self.states['Dialog_Window_Size_Listener'] = True
        else:
            self.versuchter_start('Dialog_Window_Size_Listener')
    
    def remove_Dialog_Window_Size_Listener(self):
        if self.blocked : return
        if self.mb.debug: log(inspect.stack)
        self.mb.win.removeWindowListener(self.w_listener)
        self.states['Dialog_Window_Size_Listener'] = False
        
    def add_Dialog_Event_Listener(self):
        if self.blocked : return
        if not self.states['Dialog_Event_Listener']:
            if self.mb.debug: log(inspect.stack)
            self.mb.prj_tab.AccessibleContext.AccessibleParent.addEventListener(self.w_listener)
            self.states['Dialog_Event_Listener'] = True
        else:
            self.versuchter_start('Dialog_Event_Listener')

            
    def remove_Dialog_Event_Listener(self):
        if self.blocked : return
        if self.mb.debug: log(inspect.stack)
        try:
            self.mb.prj_tab.AccessibleContext.AccessibleParent.removeEventListener(self.w_listener)
        except:
            # Fehler, wenn Organon geschlossen wird:
            # dialog hat bereits kein Fenster mehr
            pass
        self.states['Dialog_Event_Listener'] = False

            

    def add_Undo_Manager_Listener(self):
        if self.blocked : return
        if not self.states['Undo_Manager_Listener']:
            if self.mb.debug: log(inspect.stack)
            self.mb.undo_mgr.addUndoManagerListener(self.undo_mgr_listener)
            self.states['Undo_Manager_Listener'] = True
        else:
            self.versuchter_start('Undo_Manger_Listener')
    
    def remove_Undo_Manager_Listener(self):
        if self.blocked : return
        if self.mb.debug: log(inspect.stack)
        self.mb.undo_mgr.removeUndoManagerListener(self.undo_mgr_listener)
        self.states['Undo_Manager_Listener'] = False
        
    def add_Document_Close_Listener(self):
        if self.blocked : return
        if not self.states['Document_Close_Listener']:
            if self.mb.debug: log(inspect.stack)
            self.mb.doc.addDocumentEventListener(self.listener_doc_close)
            self.states['Document_Close_Listener'] = True
        else:
            self.versuchter_start('Document_Close_Listener')
    
    def remove_Document_Close_Listener(self):
        if self.blocked : return
        if self.mb.debug: log(inspect.stack)
        self.mb.doc.removeDocumentEventListener(self.listener_doc_close)
        self.states['Document_Close_Listener'] = False
        




class Tab_Auswahl():
    def __init__(self):
        if debug: log(inspect.stack)
        
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
        self.suche_term = None
        
        self.behalte_hierarchie_bei = None
        self.tab_name = None
        
        
       

class Tab ():
    def __init__(self):
        if debug: log(inspect.stack)
        self.AB = 'ORGANON'

    

class Design():
    '''
    Alte Klasse um die Tabs in Organon Fenster zu berechnen.
    Ueberfluessig, sobald alle Organon Fenster auf das fenster.py Modul umgestellt sind.
    '''
    def __init__(self):
        if debug: log(inspect.stack)
        
        self.default_tab = {}
        self.custom_tab = {}
        self.tabs = {}
        self.new_tabs = {}
        
    
    def set_default(self,tabs):
        if debug: log(inspect.stack)
        
        summe = 0
        i = 0
        for tab in tabs:
            self.default_tab.update({'tab%sx'%(i):tab})
            self.custom_tab.update({'tab%sx'%(i):0})
            
            summe = 0
            for j in range(i+1):
                summe += self.default_tab['tab%sx'%(j)]
                
            self.tabs.update({'tab%sx'%(i):summe})

            i += 1
    
            
    def setze_tab(self,tab_name,value):  
        
        for x in tab_name:
            if x.isdigit():
                break
            
        x = int(x)
        
        if self.custom_tab['tab%sx'%(x+1)] < value:
            if value > self.default_tab['tab%sx'%(x+1)]:
                self.custom_tab['tab%sx'%(x+1)] = value
          
            
    def kalkuliere_tabs(self):
        if debug: log(inspect.stack)
        
        self.new_tabs = {}
        
        for i in range(len(self.custom_tab)):
            summe = 0
            for j in range(i+1):
                if self.custom_tab['tab%sx'%(j)] != 0:
                    summe += self.custom_tab['tab%sx'%(j)]
                else:
                    summe += self.default_tab['tab%sx'%(j)]
            self.new_tabs.update( {'tab%sx'%(i) : summe} )
            self.new_tabs.update( {'tab%s'%(i) : summe} )

class Gliederung():
    def rechne(self,tree):  
        if debug: log(inspect.stack)
        
        root = tree.getroot()
        all_elem = root.findall('.//')
        
        self.lvls = {}
        for i in range(10):
            self.lvls.update({i: 0})

        gliederung = {}
        lvl = 1
        
        for el in all_elem:
            
            lvl_el = int(el.attrib['Lvl']) + 1
            
            if lvl_el == lvl:
                self.lvls[lvl] += 1
            elif lvl_el > lvl:
                self.lvls[lvl+1] += 1
                lvl += 1
            elif lvl_el < lvl:
                self.lvls[lvl_el] += 1
                lvl = lvl_el #self.lvls[lvl_el]
                for i in range(lvl,10):
                    self.lvls[i+1] = 0
            
            glied = ''
            for l in range(lvl_el):
                glied = glied + str(self.lvls[l+1]) + '.'
                
            gliederung.update({el.tag:glied})

        return gliederung 
     

    

from com.sun.star.awt import XMouseListener, XItemListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT 
    
class Menu_Leiste_Listener (unohelper.Base, XMouseListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.menu_Kopf_Eintrag = 'None'
        self.mb.geoeffnetesMenu = None
    
    def mousePressed(self, ev):
        if ev.Buttons == MB_LEFT:
            if self.mb.debug: log(inspect.stack)
            try:
                controls = []
                
                if self.menu_Kopf_Eintrag == 'Test':
                    self.mb.Test()   
                    return
                
                else:
                    items = self.mb.menuEintraege(LANG,self.menu_Kopf_Eintrag)
                    if self.menu_Kopf_Eintrag == LANG.ANSICHT:
                        controls,listener,Hoehe,Breite = self.mb.erzeuge_Menu_DropDown_Eintraege_Ansicht(items)
                    else:
                        controls,listener,Hoehe,Breite = self.mb.erzeuge_Menu_DropDown_Eintraege(items)
                        controls = list((x,'') for x in controls)

                    loc_cont = self.mb.current_Contr.Frame.ContainerWindow.AccessibleContext.LocationOnScreen
                    
                    x = self.mb.prj_tab.AccessibleContext.LocationOnScreen.X - loc_cont.X + ev.Source.PosSize.X
                    y = self.mb.prj_tab.AccessibleContext.LocationOnScreen.Y - loc_cont.Y + ev.Source.PosSize.Y + 20
                    posSize = x,y,Breite +20,Hoehe +20
                    
                    oWindow,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize,1+16+512,parent=self.mb.win)
                    
                    # Listener fuers Dispose des Fensters
                    listener2 = Schliesse_Menu_Listener(self.mb)
                    cont.addMouseListener(listener2) 
                    listener2.ob = oWindow
                    

                self.mb.geoeffnetesMenu = self.menu_Kopf_Eintrag
                listener.window = oWindow
                self.mb.menu_fenster = oWindow 
                
                for c in controls:
                    cont.addControl(c[1],c[0])
            except:
                log(inspect.stack,tb())
                

    def mouseEntered(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        ev.value.Source.Model.FontUnderline = 1     
        if self.menu_Kopf_Eintrag != ev.value.Source.Text:
            self.menu_Kopf_Eintrag = ev.value.Source.Text  
            if None not in (self.menu_Kopf_Eintrag,self.mb.geoeffnetesMenu):
                if self.menu_Kopf_Eintrag != self.mb.geoeffnetesMenu:
                    self.mb.menu_fenster.dispose()
 
        return False
   
    def mouseExited(self, ev):  
        if self.mb.debug: log(inspect.stack)
               
        ev.value.Source.Model.FontUnderline = 0
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False

class Menu_Leiste_Listener2 (unohelper.Base, XMouseListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.geklickterMenupunkt = None
        
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        if ev.Buttons == 1:
            if self.mb.debug: log(inspect.stack)
            
            if self.mb.projekt_name != None:
                
                if ev.Source.Model.HelpText == LANG.INSERT_DOC:            
                    self.mb.class_Baumansicht.erzeuge_neue_Zeile('dokument')
                    
                if ev.Source.Model.HelpText == LANG.INSERT_DIR:            
                    self.mb.class_Baumansicht.erzeuge_neue_Zeile('Ordner')
                    
                if ev.Source.Model.HelpText[:10] == LANG.FORMATIERUNG_SPEICHERN[:10]:

                    props = self.mb.props[T.AB]
                    zuletzt = props.selektierte_zeile_alt
                    bereichsname = props.dict_bereiche['ordinal'][zuletzt]
                    path = props.dict_bereiche['Bereichsname'][bereichsname]
                    self.mb.props[T.AB].tastatureingabe = True

                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(uno.systemPathToFileUrl(path),bereichsname)
                    
                    # Bestaetigung ausgeben
                    tree = self.mb.props[T.AB].xml_tree
                    root = tree.getroot()        
                    source = root.find('.//'+zuletzt)  
                    name = source.attrib['Name']
                    
                    self.mb.nachricht_si(LANG.FORMATIERUNG_SPEICHERN.format(name),1)

                self.mb.loesche_undo_Aktionen()
                return False
        
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        if ev.Source.Model.HelpText[:10] == LANG.FORMATIERUNG_SPEICHERN[:10]:
            ev.Source.Model.HelpText = LANG.FORMATIERUNG_SPEICHERN.format(self.get_zuletzt_benutzte_datei())
            
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
    
    def get_zuletzt_benutzte_datei(self):
        if self.mb.debug: log(inspect.stack)
        try:
            props = self.mb.props[T.AB]
            zuletzt = props.selektierte_zeile_alt
            xml = props.xml_tree
            root = xml.getroot()
            return root.find('.//' + zuletzt).attrib['Name']
        except:
            return 'None'

class Auswahl_Menu_Eintrag_Listener(unohelper.Base, XMouseListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.window = None
        # Wenn dieser Listener von den Texttools aus gerufen wird,
        # wird win2 auf das normale Menufenster gesetzt und disposiert
        self.texttools = False
        self.win2 = None
        
    def mouseReleased(self, ev):
        if self.mb.debug: log(inspect.stack)  
        try:
            sel = ev.Source.Text
            if sel not in [LANG.TEXTTOOLS,]:
                self.do()

            if sel == LANG.NEW_PROJECT:
                self.mb.class_Projekt.dialog_neues_projekt_anlegen()
                
            elif sel == LANG.OPEN_PROJECT:
                self.mb.class_Projekt.lade_Projekt()
                
            elif sel == LANG.NEW_DOC:
                self.mb.class_Baumansicht.erzeuge_neue_Zeile('dokument')
                
            elif sel == LANG.NEW_DIR:
                self.mb.class_Baumansicht.erzeuge_neue_Zeile('Ordner')
                
            elif sel == LANG.EXPORT_2:
                self.mb.class_Export.export()
                
            elif sel == LANG.IMPORT_2:
                self.mb.class_Import.importX()
                
            elif sel == LANG.UNFOLD_PROJ_DIR:
                self.mb.class_Funktionen.projektordner_ausklappen()
                
            elif sel == LANG.NEUER_TAB:
                self.mb.class_Tabs.start(False)
                
            elif sel == LANG.SCHLIESSE_TAB:
                self.mb.tabsX.schliesse_Tab()
                
            elif sel == LANG.ZEIGE_TEXTBEREICHE:
                oBool = self.mb.current_Contr.ViewSettings.ShowTextBoundaries
                self.mb.current_Contr.ViewSettings.ShowTextBoundaries = not oBool  
                 
            elif sel == LANG.HOMEPAGE:
                webbrowser.open('https://github.com/XRoemer/Organon')
                
            elif sel == LANG.FEEDBACK:
                webbrowser.open('http://organon4office.wordpress.com/')
                
            elif sel == LANG.BACKUP:
                self.mb.erzeuge_Backup()
                
            elif sel == LANG.TRENNE_TEXT:
                self.mb.class_Funktionen.teile_text()
                
            elif sel == LANG.IMPORTIERE_IN_TAB:
                self.mb.class_Tabs.start(True)
                
            elif sel == LANG.CLEAR_RECYCLE_BIN:  
                self.mb.leere_Papierkorb()
                
            elif sel == LANG.TRENNER:  
                self.mb.erzeuge_Trenner_Enstellungen()
                
            elif sel == LANG.TEXTVERGLEICH:  
                self.mb.class_Zitate.start()
                
            elif sel == LANG.TEXTTOOLS:  
                self.mb.texttools_geoeffnet = True
                self.mb.erzeuge_texttools_fenster(ev,self.window)
                
            elif sel == LANG.WOERTERLISTE:  
                self.mb.class_werkzeug_wListe.start()
                
            elif sel == LANG.ERZEUGE_INDEX:  
                self.mb.class_Index.start()
                
            elif sel == LANG.EINSTELLUNGEN:
                self.mb.class_Einstellungen.start()
                
            elif sel == LANG.ORGANIZER:
                self.mb.class_Organizer.run()
            
            elif sel == LANG.TEXT_BATCH_DEVIDE:
                self.mb.class_Funktionen.teile_text_batch()
                
            elif sel == LANG.DATEIEN_VEREINEN:
                self.mb.class_Funktionen.vereine_dateien()
                
    
            self.mb.loesche_undo_Aktionen()
        except:
            log(inspect.stack,tb())
        
    def do(self): 
        if self.mb.debug: log(inspect.stack)
        self.window.dispose()
        self.mb.geoeffnetesMenu = None
        
        if self.texttools:
            self.texttools = False
            self.win2.dispose()
            self.mb.texttools_geoeffnet = False
        

    def mouseExited(self,ev):
        ev.value.Source.Model.FontWeight = 100
        return False
    def mouseEntered(self,ev):
        ev.value.Source.Model.FontWeight = 150
        return False
    def mousePressed(self,ev):
        return False
    def disposing(self,ev):
        return False
    
    
class Schliesse_Menu_Listener (unohelper.Base, XMouseListener):
        def __init__(self,mb,texttools=False):
            if mb.debug: log(inspect.stack)
            
            self.ob = None
            self.mb = mb
            self.texttools = texttools
           
        def mousePressed(self, ev):
            return False
       
        def mouseExited(self, ev): 
            if self.mb.debug: log(inspect.stack)
            
            #print('texttools',self.mb.texttools_geoeffnet,self.texttools)
            
            if self.mb.texttools_geoeffnet and self.texttools == False:
                return

            point = uno.createUnoStruct('com.sun.star.awt.Point')
            point.X = ev.X
            point.Y = ev.Y

            enthaelt_Punkt = ev.Source.AccessibleContext.containsPoint(point)

            if not enthaelt_Punkt:        
                self.ob.dispose() 
                self.mb.geoeffnetesMenu = None   
                
                if self.texttools:
                    self.mb.texttools_geoeffnet = False
                    
            return False
        

        def mouseEntered(self,ev):
            return False
        def mouseReleased(self,ev):
            return False
        def disposing(self,ev):
            return False


class DropDown_Tags_TV_Listener(unohelper.Base, XMouseListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
    
    def mouseExited(self, ev):
        ev.value.Source.Model.FontWeight = 100
        return False
    def mouseEntered(self,ev):
        ev.value.Source.Model.FontWeight = 150
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack)
            
        try:
            text = ev.Source.Text
            sett = self.mb.settings_proj
            
            
            if text == LANG.SHOW_TAG1:
                nummer = 1
                lang_show_tag = LANG.SHOW_TAG1
                
            elif text == LANG.SHOW_TAG2:
                if not self.mb.class_Funktionen.pruefe_galerie_eintrag():
                    return
                nummer = 2
                lang_show_tag = LANG.SHOW_TAG2
            
            elif text == LANG.GLIEDERUNG:
                nummer = 3
                lang_show_tag = LANG.GLIEDERUNG
            
            tag = 'tag%s'%nummer
            sett[tag] = not sett[tag]
            
            self.mb.class_Funktionen.mache_tag_sichtbar(sett[tag],tag)
            ctrl = ev.Source.Context.getControl(lang_show_tag+'_icon')
            
            if sett[tag]:
                ctrl.Model.ImageURL = 'private:graphicrepository/svx/res/apply.png' 
            else:
                ctrl.Model.ImageURL = '' 
            
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        except:
            log(inspect.stack,tb())
    
    
       

class DropDown_Tags_SB_Listener(unohelper.Base, XMouseListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
    def mouseExited(self, ev):
        ev.value.Source.Model.FontWeight = 100
        return False
    def mouseEntered(self,ev):
        ev.value.Source.Model.FontWeight = 150
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
    def mousePressed(self, ev):   
        if self.mb.debug: log(inspect.stack)    
        try:
            name = ev.Source.Text            
            ctrl = ev.Source.Context.getControl(name+'_icon') 
            #state = ev.Source.State
            
            panels_nr = self.mb.tags['name_nr'][name]

            if ctrl.Model.ImageURL == 'private:graphicrepository/svx/res/apply.png':
                self.mb.tags['sichtbare'].remove(panels_nr)
                ctrl.Model.ImageURL = '' 
            else:
                self.mb.tags['sichtbare'].append(panels_nr)
                ctrl.Model.ImageURL = 'private:graphicrepository/svx/res/apply.png' 
                
            self.mb.class_Sidebar.erzeuge_sb_layout()

        except:
            log(inspect.stack,tb())   


            
from com.sun.star.awt import Rectangle
from com.sun.star.awt import WindowDescriptor         
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.awt.VclWindowPeerAttribute import OK,YES_NO_CANCEL, DEF_NO

class Mitteilungen():
    def __init__(self,ctx,mb):
        if mb.debug: log(inspect.stack)
        
        self.ctx = ctx
        self.mb = mb     
    
    def nachricht(self, MsgText, MsgType="errorbox", MsgButtons=OK):  
        if self.mb.debug: log(inspect.stack)           

        smgr = self.ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
        doc = desktop.getCurrentComponent()                
        ParentWin = doc.CurrentController.Frame.ContainerWindow

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
        
        farben = self.mb.settings_orga['organon_farben']
        
        if farben['design_organon_fenster']:
            msgbox.setBackground(farben['hf_hintergrund'])
            msgbox.StyleSettings.FieldTextColor = farben['schrift_datei']
        
        x = msgbox.execute()
        msgbox.dispose()
        
        return x
    
    def nachricht_si(self,nachricht,zeit):
        if self.mb.debug: log(inspect.stack)  
        
        if len(nachricht) < 200:
            x = 200 - len(nachricht)
            nachricht = int(x/2)*' ' + nachricht + int(x/2)*' '
        
        StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
        StatusIndicator.start(nachricht,2)
        StatusIndicator.setValue(2)
        time.sleep(zeit)
        StatusIndicator.end()
        
    def popup(self,nachricht,zeit=3,parent=None):
        if self.mb.debug: log(inspect.stack)  
        
        if parent == None:
            parent = self.mb.toolkit.getActiveTopWindow() 
            if parent == None:
                parent = self.mb.win
        
        posSize = parent.PosSize.Width/2-100, parent.PosSize.Height/2-50, 0, 0
        fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize,Flags=1+32+64+128,parent=parent)
        
        ctrl, model = self.mb.createControl(self.ctx, "FixedText", 10, 15, 0, 0, (), ())          
        model.Label = nachricht
        
        pref_Size = ctrl.getPreferredSize()
        Hoehe = pref_Size.Height + 30
        Breite = pref_Size.Width + 20
        ctrl.setPosSize(0,0,Breite,Hoehe,12)
        fenster.setPosSize(0,0,Breite,Hoehe,12)
    
        fenster_cont.addControl('Text', ctrl)
        
        if zeit != 'frei':
            fenster.draw(1,1)
            time.sleep(zeit)
            fenster.dispose()

        return fenster, model
        
        

from com.sun.star.document import XUndoManagerListener 
class Undo_Manager_Listener(unohelper.Base,XUndoManagerListener): 
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
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
    
    
    def undoActionAdded(self,ev):pass
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
            
            self.mb.Listener.remove_Undo_Manager_Listener()
            self.mb.doc.Text.insertTextContent(cur,newSection,False)
            self.mb.Listener.add_Undo_Manager_Listener()
            
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
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        mb.current_Contr.addKeyHandler(self)
        
    def keyPressed(self,ev):
        
        code = ev.KeyCode
        mods = ev.Modifiers
        
        if mods > 1:
            # 0 = keine Modifikation
            # 1 = Shift
            # 2 = Strg
            # 3 = Shift + Strg
            # 4 = Alt
            # 5 = Shift + Alt
            # 6 = Strg + Alt
            # 7 = Shift + Strg + Alt
            self.mb.class_Shortcuts.shortcut_ausfuehren(code,mods)
        
        else:
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.props[T.AB].zuletzt_gedrueckte_taste = ev
            return False
        
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
    
    def disposing(self,ev):
        return False


from com.sun.star.view import XSelectionChangeListener
class ViewCursor_Selection_Listener(unohelper.Base, XSelectionChangeListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ts_old = 'nicht vorhanden'
        self.mb.selbstruf = False
        
    def disposing(self,ev):
        if self.mb.debug: log(inspect.stack)
        return False
    
    def selectionChanged(self,ev):
        #if self.mb.debug: log(inspect.stack)

        try:
            if self.mb.selbstruf:
                return False

            selected_ts = self.mb.current_Contr.ViewCursor.TextSection   

            if selected_ts == None:
                try:
                    vc = ev.Source.ViewCursor
                    anchor = vc.Text.Anchor
                    inner_sec = anchor.TextSection
                    selected_ts = self.mb.class_Zeilen_Listener.finde_eltern_bereich(inner_sec)
                except:
                    log(inspect.stack,tb())
                    return False

            s_name = selected_ts.Name
            props = self.mb.props[T.AB]

            # stellt sicher, dass nur selbst erzeugte Bereiche angesprochen werden
            # und der Trenner uebersprungen wird
            if 'trenner'  in s_name:

                if props.zuletzt_gedrueckte_taste == None:
                    try:
                        self.mb.viewcursor.goDown(1,False)
                    except:
                        self.mb.viewcursor.goUp(1,False)
                    return False
                # 1024,1027 Pfeil runter,rechts
                elif props.zuletzt_gedrueckte_taste.KeyCode in (1024,1027):  
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
                # ist der Bereich ausserhalb des Organon trees
                if 'OrganonSec' not in selected_ts.Name:
                    return False
                
            props = self.mb.props[T.AB]  
                
            self.so_name =  None   

            if props.selektierte_zeile_alt != None:
                ts_old_bereichsname = props.dict_bereiche['ordinal'][props.selektierte_zeile_alt]
                self.ts_old = self.mb.doc.TextSections.getByName(ts_old_bereichsname)            
                self.so_name = props.dict_bereiche['ordinal'][props.selektierte_zeile_alt]
            
            if self.ts_old == 'nicht vorhanden':
                #print('selek gewechs, old nicht vorhanden')
                self.ts_old = selected_ts 
                ordinal = props.dict_bereiche['Bereichsname-ordinal'][s_name]
                props.selektierte_zeile = ordinal
                props.selektierte_zeile_alt = ordinal
                return False 
            
            elif props.Papierkorb_geleert == True:
                #print('selek gewechs, Papierkorb_geleert')
                self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.ts_old.FileLink.FileURL,self.so_name)
                self.ts_old = selected_ts 
                props.Papierkorb_geleert = False 
                return False       
            else:
                if self.ts_old == selected_ts:
                    #print('selek nix gewechs',self.so_name , s_name)
                    return False                
                else:
                    #print('selek gewechs',self.so_name , s_name)                    
                    self.farbe_der_selektion_aendern(selected_ts.Name)
                    if props.tastatureingabe:
                        self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.ts_old.FileLink.FileURL,self.so_name)
                        
                    self.ts_old = selected_ts  
        except:
            log(inspect.stack,tb())

   
    def test_for_parent_section(self,selected_text_sectionX,sec):
        try:
            if selected_text_sectionX.ParentSection != None:
                selected_text_sectionX = selected_text_sectionX.ParentSection
                self.test_for_parent_section(selected_text_sectionX,sec)
            else:
                sec.append(selected_text_sectionX)
        except:
            log(inspect.stack,tb())
            
    def farbe_der_selektion_aendern(self,bereichsname): 
        if self.mb.debug: log(inspect.stack)    
        
        props = self.mb.props[T.AB]
        ordinal = props.dict_bereiche['Bereichsname-ordinal'][bereichsname]
        zeile = props.Hauptfeld.getControl(ordinal)
        textfeld = zeile.getControl('textfeld')
        
        props.selektierte_zeile = zeile.AccessibleContext.AccessibleName
        # selektierte Zeile einfaerben, ehem. sel. Zeile zuruecksetzen
        textfeld.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE 
        
        if props.selektierte_zeile_alt != None: 
            ctrl = props.Hauptfeld.getControl(props.selektierte_zeile_alt).getControl('textfeld') 
            ctrl.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            self.mb.class_Sidebar.erzeuge_sb_layout()
            
        props.selektierte_zeile_alt = textfeld.Context.AccessibleContext.AccessibleName
     
            
            
from com.sun.star.awt import XWindowListener
from com.sun.star.lang import XEventListener
class Dialog_Window_Size_Listener(unohelper.Base,XWindowListener,XEventListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
    
    def windowResized(self,ev):
        #print('windowResized')
        self.mb.tabsX.layout_tab_zeilen(ev.Width)
        self.mb.class_Baumansicht.korrigiere_scrollbar()

    def windowMoved(self,ev):pass
        #print('windowMoved')
    def windowShown(self,ev):
        pass
        #self.mb.class_Baumansicht.korrigiere_scrollbar()
        #print('windowShown')
    def windowHidden(self,ev):pass
   
    def disposing(self,arg):
        if self.mb.debug: log(inspect.stack)
        
        try:                
            # speichern, wenn Organon beendet wird.
            # aenderungen nach tabwechsel werden in Tab_Listener.activated() gespeichert
            if self.mb.props[T.AB].tastatureingabe:
                ordinal = self.mb.props[T.AB].selektierte_zeile
                bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)   

            if 'files' in self.mb.pfade: 
                self.mb.class_Tags.speicher_tags()   
                
                self.mb.tags['sichtbare'] = []
                # Wenn Organon durch den Organon Schalter geschlossen wird,
                # existiert die Seitenleiste noch und wird hier wieder zurueckgesetzt
                desk = self.mb.desktop
                if hasattr(desk.CurrentComponent,'CurrentController'):
                    self.mb.dict_sb['erzeuge_sb_layout']()

            self.mb.Listener.entferne_alle_Listener()
            self.mb = None

        except:
            log(inspect.stack,tb())
            
        return False


from com.sun.star.document import XDocumentEventListener
class Document_Close_Listener(unohelper.Base,XDocumentEventListener):
    '''
    Lets Writer close without warning.
    As everything is saved by Organon, a warning isn't necessary.
    Even more: it might confuse the user.
    
    '''
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb

    def documentEventOccured(self,ev):
        if ev.EventName == 'OnPrepareViewClosing':
            if self.mb.debug: log(inspect.stack)
            # Um das Dokument ohne Speicherabfrage zu schliessen
            self.mb.doc.setModified(False)
            
    def disposing(self,ev):
        return False
    
