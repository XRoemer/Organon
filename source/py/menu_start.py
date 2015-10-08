# -*- coding: utf-8 -*-
import uno
import unohelper
from traceback import format_exc as tb
import sys
from os import walk,path,remove
from codecs import open as codecs_open
from inspect import stack as inspect_stack
from shutil import copyfile
import json


platform = sys.platform



class Menu_Start():
    
    def __init__(self,args):
        
        (pdk,
         dialog,
         ctx,
         path_to_extension,
         win,
         dict_sb,
         debugX,
         factory,
         logX,
         class_LogX,
         konst,
         settings_orga) = args
        
        global debug,log,class_Log,KONST
        debug = debugX
        log = logX
        class_Log = class_LogX
        KONST = konst
        
        if debug: log(inspect_stack)      
        
        self.win = win
        self.pd = pdk
        global pd
        pd = pdk
        
        try:
             
            # Konstanten
            self.factory = factory
            self.dialog = dialog
            self.ctx = ctx
            self.smgr = self.ctx.ServiceManager
            self.desktop = self.smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
            self.path_to_extension = path_to_extension
            self.programm = self.get_office_name()
            self.platform = sys.platform
            self.language = None
            self.LANG = self.lade_Modul_Language()
            self.settings_orga = settings_orga 
            self.zuletzt_geladene_Projekte = self.get_zuletzt_geladene_Projekte()
            self.dict_sb = dict_sb
            self.templates = {}
            
            try:
                self.templates.update({'standard_stil':self.get_stil(),
                                       'get_stil':self.get_stil})
            except:
                log(inspect_stack,tb())
                
            
            
        except Exception as e:
            log(inspect_stack,tb())
        
    
    
    def wurde_als_template_geoeffnet(self):
        if debug: log(inspect_stack)
        try:
            enum = self.desktop.Components.createEnumeration()
            comps = []

            while enum.hasMoreElements():
                comps.append(enum.nextElement())
                
            # Wenn ein neues Dokument geoeffnet wird, gibt es bei der Initialisierung
            # noch kein Fenster, aber die Komponente wird schon aufgefuehrt.
            doc = comps[0]
            
            # Pruefen, ob doc von Organon erzeugt wurde
            ok = False
            for a in doc.Args:
                if a.Name == 'DocumentTitle':
                    if a.Value.split(';')[0] == 'opened by Organon':
                        ok = True
                        projekt_pfad = a.Value.split(';')[1]
                        break
            if not ok:
                return False        
            
            #projekt_pfad = 'C:\\Users\\Homer\\Documents\\organon projekte\\test2.organon\\test2.organon'
            #projekt_pfad = 
            #self.erzeuge_Startmenu()
#         
#             # Das Editfeld ueberdeckt kurzzeitig das Startmenu fuer eine bessere Anzeige
#             control, model = self.createControl(self.ctx,"Edit",0,0,1500,1500,(),() )  
#             model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
#             self.cont.addControl('wer',control)
            
            self.cont.dispose()
            self.erzeuge_Menu()
            
            prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop2.Name = 'Overwrite'
            prop2.Value = True
            
            doc.storeAsURL(projekt_pfad,(prop2,)) 
            
            sys_pfad = uno.fileUrlToSystemPath(projekt_pfad)
            orga_name = path.basename(sys_pfad).split('.')[0] + '.organon'
            sys_pfad1 = sys_pfad.split(orga_name)[0]
            pfad = path.join(sys_pfad1,orga_name,orga_name)
            

            self.Menu_Bar.class_Projekt.lade_Projekt(False,pfad)
            
          
        except:
            log(inspect_stack,tb())

    def get_office_name(self):
        if debug: log(inspect_stack)
        frame = self.desktop.Frames.getByIndex(0)

        if 'LibreOffice' in frame.Title:
            programm = 'LibreOffice'
        elif 'OpenOffice' in frame.Title:
            programm = 'OpenOffice'
        else:
            # Fuer Linux / OSX fehlt
            programm = 'LibreOffice'
        
        return programm
       
    def erzeuge_Startmenu(self):
        if debug: log(inspect_stack)
            
        # Hauptfeld_Aussen
        Attr = (0,0,1000,1800,'Hauptfeld_aussen1', 0)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        self.cont, model1 = self.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        model1.BackgroundColor = KONST.FARBE_HF_HINTERGRUND

        self.dialog.addControl('Hauptfeld_aussen1',self.cont)  


        Attr = (150,60,120,153,'Hauptfeld_aussen1', 0)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control, model = self.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
        
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/organon icon_120.png' 
        model.Border = False   
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        self.cont.addControl('Hauptfeld_aussen1',control)  
        
       
        self.listener = Menu_Listener(self)
        
        
        PosX = 30
        PosY = 50
        Width = 100
        Height = 30
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('neues_projekt')
        control.addActionListener(self.listener)
        model.Label = self.LANG.NEW_PROJECT
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND

        
        self.cont.addControl('Hauptfeld_aussen1',control)  
        
        PosY += 50
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('projekt_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.OPEN_PROJECT
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        
        self.cont.addControl('Hauptfeld_aussen1',control) 
        
        
        PosY += 70
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('anleitung_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.LOAD_DESCRIPTION
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        
        self.cont.addControl('Hauptfeld_aussen1',control) 
        
        PosY += 150
        
        if self.zuletzt_geladene_Projekte == []:
            return

        for proj in self.zuletzt_geladene_Projekte:
            
            
            name,pfad = proj
            
            control, model = self.createControl(self.ctx,"FixedText",PosX,PosY,200,20,(),() )  
            control.addMouseListener(self.listener)
            model.Label = name
            model.TextColor = KONST.FARBE_SCHRIFT_DATEI
            model.HelpText = pfad
            self.cont.addControl('Hauptfeld_aussen1',control) 
            PosY += 25

        self.wurde_als_template_geoeffnet()
        
    
    def erzeuge_Menu(self):
        if debug: log(inspect_stack)
          
        try:   
            import menu_bar
            
            args = (pd,
                    self.dialog,
                    self.ctx,
                    self.path_to_extension,
                    self.win,
                    self.dict_sb,
                    debug,
                    self.factory,
                    self,
                    log,
                    class_Log,
                    self.settings_orga,
                    self.templates
                    )
            
            self.module_mb = menu_bar
            
            self.Menu_Bar = menu_bar.Menu_Bar(args)
            self.Menu_Bar.erzeuge_Menu(self.Menu_Bar.prj_tab)
        except:
            log(inspect.stack,tb())    
    
           
              
    def lade_Modul_Language(self):
        if debug: log(inspect_stack)
        try:  
            enum = self.desktop.Components.createEnumeration()
            comps = []
            
            while enum.hasMoreElements():
                comps.append(enum.nextElement())
    
            language = comps[0].CharLocale.Language
    
            if language not in ('de'):
                language = 'en'
                
            self.language = language
            
            import lang_en 
            try:
                exec('import lang_' + language)
            except Exception as e:
                log(inspect_stack,tb())    
    
            if 'lang_' + language in vars():
                lang = vars()['lang_' + language]
            else:
                lang = vars()[lang_en]  
            
            return lang
        except Exception as e:
            log(inspect_stack,tb())    
    
    
    def get_zuletzt_geladene_Projekte(self):
        if debug: log(inspect_stack)
        
        try:
            projekte = self.settings_orga['zuletzt_geladene_Projekte']
            
            # Fuer projekte erstellt vor v0.9.9.8b
            if isinstance(projekte, dict):
                list_proj = list(projekte)
                projekte = [[p,projekte[p]] for p in list_proj]
                
            
            inexistent = [p for p in projekte if not path.exists(p[1])]
            
            for i in inexistent:
                index = projekte.index(i)
                del(projekte[index])
                
            self.settings_orga['zuletzt_geladene_Projekte'] = projekte
            
            return projekte
        except Exception as e:
            print(e)
            try:
                if debug: log(inspect_stack,tb())
            except:
                pass
        return []
    
    def get_doc(self):
        if debug: log(inspect_stack)
        
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
    
    def get_stil(self):
        if debug: log(inspect_stack)

        try:
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            doc = desktop.getCurrentComponent() 
            
            if doc == None:
                doc = self.get_doc()                
            
            oStyleFamilies = doc.StyleFamilies 
            oPageStyles = oStyleFamilies.getByName("PageStyles")
            oDefaultStyle = oPageStyles.getByName("Standard")
            
            style = {}
            
            nicht_nutzbar_OO = ['DisplayName','FooterText','FooterTextLeft','FooterTextRight',
                            'HeaderText','HeaderTextLeft','HeaderTextRight',
                            'ImplementationName','IsPhysical','Name','ParentStyle',
                            'PropertiesToDefault','PropertyToDefault'
                            ]
            
            nicht_nutzbar_LO = ['DisplayName','FooterText','FooterTextLeft','FooterTextRight',
                            'HeaderText','HeaderTextLeft','HeaderTextRight',
                            'ImplementationName','IsPhysical','Name','ParentStyle',
                            'PropertiesToDefault','PropertyToDefault',
    
                            'FillBitmap','FooterTextFirst','HeaderTextFirst',
                            'FillBackground','FillBitmapLogicalSize','FillBitmapMode',
                            'FillBitmapName','FillBitmapOffsetX','FillBitmapOffsetY',
                            'FillBitmapPositionOffsetX','FillBitmapPositionOffsetY',
                            'FillBitmapRectanglePoint','FillBitmapSizeX','FillBitmapSizeY',
                            'FillBitmapStretch','FillBitmapTile','FillColor',
                            'FillGradient','FillGradientName','FillGradientStepCount',
                            'FillHatch','FillHatchName','FillStyle','FillTransparence',
                            'FillTransparenceGradient','FillTransparenceGradientName',
                            'FooterIsShared','HeaderIsShared'
                            ]
            
            for o in dir(oDefaultStyle):
                value = getattr(oDefaultStyle,o)
                if type(value) in [str,int,type(u''),bool,type(None)]:
                    style.update({o:value})
                    
            if self.programm == 'OpenOffice':
                nicht_nutzbar = nicht_nutzbar_OO
            else:
                nicht_nutzbar = nicht_nutzbar_LO
                    
            default_template_style = {s:style[s] for s in style}# if s[0] not in nicht_nutzbar}
            
            #log(inspect_stack,extras=str(len(default_template_style)))
            
        except Exception as e:
            log(inspect_stack,tb())
            return {}
        
        return default_template_style
   
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




from com.sun.star.awt import XActionListener,XMouseListener
    
class Menu_Listener (unohelper.Base, XActionListener,XMouseListener):
    def __init__(self,menu):
        self.menu = menu

    def actionPerformed(self, ev):
        if debug: log(inspect_stack)
        
        try:
            if ev.ActionCommand == 'neues_projekt':
                self.menu.cont.dispose()
                self.menu.erzeuge_Menu()
                self.menu.Menu_Bar.class_Projekt.erzeuge_neues_Projekt()
                
            elif ev.ActionCommand == 'projekt_laden':
                self.menu.cont.dispose()
                self.menu.erzeuge_Menu()
                self.menu.Menu_Bar.class_Projekt.lade_Projekt()
                
            elif ev.ActionCommand == 'anleitung_laden':
                pfad = self.get_Org_description_path()
                self.menu.cont.dispose()
                self.menu.erzeuge_Menu()
                self.menu.Menu_Bar.class_Projekt.lade_Projekt(False,pfad)
                self.menu.Menu_Bar.anleitung_geladen = True
        except:
            log(inspect.stack,tb())
            
    def get_Org_description_path(self):
        if debug: log(inspect_stack)
        
        path_HB = path.join(self.menu.path_to_extension,'description','Handbuecher')

        ordner = []
        for (dirpath, dirnames, filenames) in walk(path_HB):
            ordner.extend(dirnames)
            break
        
        if self.menu.language in ordner:
            path_HB = path.join(path_HB,self.menu.language)
        else:
            path_HB = path.join(path_HB,'en')
            
        projekt_name = []
        for (dirpath, dirnames, filenames) in walk(path_HB):
            projekt_name.extend(dirnames)
            break
        
        desc_path = path.join(path_HB,projekt_name[0],projekt_name[0])
        return desc_path
    
    def mouseReleased(self, ev):  
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self, ev):
        if debug: log(inspect_stack)
        
        try:
            projekt_pfad = ev.Source.Model.HelpText
            
            # Das Editfeld ueberdeckt kurzzeitig das Startmenu fuer eine bessere Anzeige
            control, model = self.menu.createControl(self.menu.ctx,"Edit",0,0,1500,1500,(),() )  
            model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            self.menu.cont.addControl('wer',control)
    
            self.menu.erzeuge_Menu()
            self.menu.Menu_Bar.class_Projekt.lade_Projekt(False,projekt_pfad)
            
            self.menu.cont.dispose()
        except:
            log(inspect.stack,tb())
        
    def disposing(self,ev):
        pass
           

