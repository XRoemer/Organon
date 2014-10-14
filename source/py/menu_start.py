# -*- coding: utf-8 -*-
import uno
import unohelper
import traceback
import sys
from os import walk,path,remove
from codecs import open as codecs_open
from inspect import stack as inspect_stack

import konstanten as KONST

tb = traceback.print_exc
platform = sys.platform

class Menu_Start():
    
    def __init__(self,args):
        
        (pdk,
         dialog,
         ctx,
         tabs,
         path_to_extension,
         win,
         dict_sb,
         debugX,
         load_reloadX,
         factory,
         logX,
         class_LogX) = args
        
        global debug,log,class_Log,load_reload
        debug = debugX
        log = logX
        class_Log = class_LogX
        load_reload = load_reloadX
        
        if debug: log(inspect_stack)
        
        if load_reload:
            self.get_pyPath()       
        
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
            self.programm = self.get_office_name()
            self.tabs = tabs
            self.platform = sys.platform
            self.language = None
            self.LANG = self.lade_Modul_Language()
            self.path_to_extension = path_to_extension
            self.zuletzt_geladene_Projekte = self.get_zuletzt_geladene_Projekte()
            self.win = win
            self.dict_sb = dict_sb
            
            dialog.Model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        except:
            log(inspect.stack,tb())


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
        model1.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD

        self.dialog.addControl('Hauptfeld_aussen1',self.cont)  


        Attr = (150,60,120,153,'Hauptfeld_aussen1', 0)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control, model = self.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
        
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/organon icon_120.png' 
        model.Border = False   
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
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
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD

        
        self.cont.addControl('Hauptfeld_aussen1',control)  
        
        PosY += 50
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('projekt_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.OPEN_PROJECT
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        
        self.cont.addControl('Hauptfeld_aussen1',control) 
        
        
        PosY += 70
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('anleitung_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.LOAD_DESCRIPTION
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        
        self.cont.addControl('Hauptfeld_aussen1',control) 
        
        PosY += 150
        
        if self.zuletzt_geladene_Projekte == None:
            return
        
        for proj in self.zuletzt_geladene_Projekte:
            name,pfad = proj
            
            control, model = self.createControl(self.ctx,"FixedText",PosX,PosY,200,20,(),() )  
            control.addMouseListener(self.listener)
            model.Label = name
            model.HelpText = pfad
            self.cont.addControl('Hauptfeld_aussen1',control) 
            PosY += 25


            
            
            
    
    def erzeuge_Menu(self):
        if debug: log(inspect_stack)
        
        try:   
            if load_reload:
                modul = 'menu_bar'
                menu_bar = load_reload_modul(modul,pyPath,self)  # gleichbedeutend mit: import menu_bar
            else:
                import menu_bar
            
            
            args = (pd,
                    self.dialog,
                    self.ctx,
                    self.tabs,
                    self.path_to_extension,
                    self.win,
                    self.dict_sb,
                    debug,
                    load_reload,
                    self.factory,
                    self,
                    log,
                    class_Log
                    )
            
            self.module_mb = menu_bar
            self.Menu_Bar = menu_bar.Menu_Bar(args)
            self.Menu_Bar.erzeuge_Menu(self.dialog)
        except:
            log(inspect.stack,tb())    
        
              
    def lade_Modul_Language(self):
        if debug: log(inspect_stack)
        
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
        except:
            pass

        if 'lang_' + language in vars():
            lang = vars()['lang_' + language]
        else:
            lang = vars()[lang_en]   

        return lang
    
    def get_pyPath(self):
        if debug: log(inspect_stack)
        
        global pyPath
        pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
        if platform == 'linux':
            pyPath = '/home/xgr/workspace/organon/Organon/source/py'
            sys.path.append(pyPath)
    
    def get_zuletzt_geladene_Projekte(self):
        if debug: log(inspect_stack)
        
        try:
            pfad = path.join(self.path_to_extension,'zuletzt_geladene_Projekte.txt')
            
            if not path.exists(pfad):
                with codecs_open(pfad, "w",encoding='utf8') as file:
                    file.write('') 
            else:
                with codecs_open(pfad, "r",encoding='utf8') as file:
                    lines = file.readlines() 
                
            x = list(a.split('++oo++') for a in lines)
            geladene_Projekte = list((a,b.replace('\n','')) for a,b in x if path.exists(b.replace('\n','')))
            return geladene_Projekte
        except:
            if debug: log(inspect_stack,tb())
            return None
    
   
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
        print('released')
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self, ev):
        if debug: log(inspect_stack)
        
        projekt_pfad = ev.Source.Model.HelpText
        
        # Das Editfeld ueberdeckt kurzzeitig das Startmenu fuer eine bessere Anzeige
        control, model = self.menu.createControl(self.menu.ctx,"Edit",0,0,1500,1500,(),() )  
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        self.menu.cont.addControl('wer',control)

        self.menu.erzeuge_Menu()
        self.menu.Menu_Bar.class_Projekt.lade_Projekt(False,projekt_pfad)
        
        self.menu.cont.dispose()
        
    def disposing(self,ev):
        pass
           
    
################ TOOLS ################################################################



def load_reload_modul(modul,pyPath,mb):
    try:
        if pyPath not in sys.path:
            sys.path.append(pyPath)

        exec('import '+ modul)
        del(sys.modules[modul])
        try:
            if mb.programm == 'LibreOffice':
                import shutil
                shutil.rmtree(path.join(pyPath,'__pycache__'))

            elif mb.programm == 'OpenOffice':

                path_menu = __file__.split(__name__)
                pfad = path_menu[0] + modul + '.pyc'

                try:
                    remove(pfad)
                except:
                    pass
        except:
            log(inspect.stack,tb())
                            
        exec('import '+ modul)

        return eval(modul)
    except:
        log(inspect.stack,tb())
        

    
    
