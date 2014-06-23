# -*- coding: utf-8 -*-
import uno
import unohelper
import traceback
import sys
from os import walk,path,remove

import konstanten as KONST

tb = traceback.print_exc
platform = sys.platform

class Menu_Start():
    
    def __init__(self,pdk,dialog,ctx,tabs,path_to_extension,win,debugX):
        
        global debug
        debug = debugX
        if debug:
            self.get_pyPath()
        
        self.win = win
        self.pd = pdk
        global pd
        pd = pdk
        
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
        self.tabs = tabs
        self.platform = sys.platform
        self.language = None
        self.LANG = self.lade_Modul_Language()
        self.path_to_extension = path_to_extension
        self.win = win
        
        dialog.Model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
      

    def erzeuge_Startmenu(self):
        try:
            self.erzeuge_Buttons()
        except:
            tb()
        
    
    def erzeuge_Buttons(self):
            
        # Hauptfeld_Aussen
        Attr = (0,0,1000,1800,'Hauptfeld_aussen', 0)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        self.cont, model1 = self.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        model1.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
                        
        self.dialog.addControl('Hauptfeld_aussen',self.cont)  


        Attr = (150,60,120,142,'Hauptfeld_aussen', 0)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control, model = self.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
        
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/organon icon.png' 
        model.Border = False   
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        self.cont.addControl('Hauptfeld_aussen',control)  
        
       
        self.listener = Menu_Listener(self)
        
        
        PosX = 30
        PosY = 50
        Width = 100
        Height = 30
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('neues_projekt')
        control.addActionListener(self.listener)
        model.Label = self.LANG.NEW_PROJECT
        
        self.cont.addControl('Hauptfeld_aussen',control)  
        
        PosY += 50
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('projekt_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.OPEN_PROJECT
          
        self.cont.addControl('Hauptfeld_aussen',control) 
        
        
        PosY += 70
        
        control, model = self.createControl(self.ctx,"Button",PosX,PosY,Width,Height,(),() )  
        control.setActionCommand('anleitung_laden')
        control.addActionListener(self.listener)
        model.Label = self.LANG.LOAD_DESCRIPTION
          
        self.cont.addControl('Hauptfeld_aussen',control) 

    
    def erzeuge_Menu(self):
        try:   
            if debug:
                modul = 'menu_bar'
                menu_bar = load_reload_modul(modul,pyPath,self)  # gleichbedeutend mit: import menu_bar
            else:
                import menu_bar
                
            self.Menu_Bar = menu_bar.Menu_Bar(self.pd,self.dialog,self.ctx,self.tabs,self.path_to_extension,self.win,debug)
            self.Menu_Bar.erzeuge_Menu()
        except:
            tb()    
        
              
    def lade_Modul_Language(self):

        language = self.dialog.AccessibleContext.Locale.Language
        
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
        global pyPath
        pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
        if platform == 'linux':
            pyPath = '/home/xgr/Arbeitsordner/organon/py'
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




from com.sun.star.awt import XActionListener
    
class Menu_Listener (unohelper.Base, XActionListener):
    def __init__(self,menu):
        self.menu = menu

    def actionPerformed(self, ev):
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
        except:
            tb()
            
    def get_Org_description_path(self):
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
                shutil.rmtree(path.join(pyPath,'__pycache__'))

            elif mb.programm == 'OpenOffice':

                path_menu = __file__.split(__name__)
                pfad = path_menu[0] + modul + '.pyc'
                #print(path)
                try:
                    remove(pfad)
                except:
                    pass
        except:
            traceback.print_exc()
                            
        exec('import '+ modul)

        return eval(modul)
    except:
        traceback.print_exc()
        

    
    
