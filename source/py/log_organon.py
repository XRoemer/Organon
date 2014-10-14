# -*- coding: utf-8 -*-

from time import clock,sleep
from os.path import join,exists
from unohelper import Base
from uno import fileUrlToSystemPath
from traceback import print_exc as tb

class Log():
    
    def __init__(self,path_to_extension,pdX,tbX):  
        
        global debug,pd,tb
        pd = pdX
        tb = tbX
        
        self.path_to_extension = path_to_extension
        self.debug = False  
        self.load_reload = False      
        self.timer_start = clock()
        self.path_to_project_settings = None
        
        # Default Debug Settings
        self.location_debug_file = path_to_extension
        self.output_console = 0
        self.write_debug_file = 0
        self.log_args = 0
        
        if not self.load_debug_config_file():
            self.write_debug_config_file()
            
        if self.write_debug_file:
            self.schreibe_logfile_kopfzeile()
        
        
    def write_debug_config_file(self):
        
        string = (
                  'location_debug_file=' + self.location_debug_file,
            'output_console=' + str(self.output_console),
            'write_debug_file=' + str(self.write_debug_file),
            'log_args=' + str(self.log_args)
            )
        
        print(string)
        
        path = path = join(self.path_to_extension,'log_config.txt')
        
        with open(path , "w") as file:
            for s in string:
                file.write(s +'\n')
   
    
    def load_debug_config_file(self):
        
        path = join(self.path_to_extension,'log_config.txt')
        
        if not exists(path):
            return False

        with open(path , "r") as file:
            
            line = file.readline()
            value = line.split('location_debug_file=')[1].replace('\n','')
            self.location_debug_file = value
            
            line = file.readline()
            value = line.split('output_console=')[1].replace('\n','')
            self.output_console = int(value)
            self.debug = int(value)

            line = file.readline()
            value = line.split('write_debug_file=')[1].replace('\n','')
            self.write_debug_file = int(value)
            
            line = file.readline()
            value = line.split('log_args=')[1].replace('\n','')
            self.log_args = int(value)

        return True
    
    def schreibe_logfile_kopfzeile(self):
        try:

            path = join(self.location_debug_file,'organon_log.txt')
            
            with open(path , "a") as file:
                file.write('################'+'\r\n')
                file.write(''+'\r\n')
                file.write('Organon opened'+'\r\n')
                file.write(''+'\r\n')
                file.write('################'+'\r\n')
        except:
            print(tb())
        
        
    def debug_time(self):
        zeit = "%0.2f" %(clock()-self.timer_start)
        return zeit
    
    def log(self,args,traceb = None,extras = None):

        try:
            info = args()

            try:
                caller = info[2][3]
                caller_class = info[2][0].f_locals['self'].__class__.__name__
            except:
                caller = 'EventObject'
                caller_class = ''
            
            call = caller_class + '.' + caller + ' )'
            
            if self.log_args:
                try:
                    argues = self.format_argues(info[1][0].f_locals)
                except:
                    argues = ''
                
            function = info[1][3]
            modul = info[1][0].f_locals['self'].__class__.__name__

            if modul in ('ViewCursor_Selection_Listener'):
                return
            
            if self.load_reload:
                if function in ('mouseEntered','mouseExited'):
                    return
                #sleep(0.04)

            if len(modul) > 18:
                modul = modul[0:18]
            
            if self.log_args:
                string = '%-7s %-18s %-40s %s( caller: %-60s args: %s ' %(self.debug_time(),modul,function,'',call,argues)
            else:
                string = '%-7s %-18s %-40s %s( caller: %s' %(self.debug_time(),modul,function,'',call)
            
            
            try:
                print(string)
            except:
                pass
            
            if self.write_debug_file:
                path = join(self.location_debug_file,'organon_log.txt')
                with open(path , "a") as file:
                    file.write(string+'\n')
                
                if traceb != None:
                    print(traceb)
                    with open(path , "a") as file:
                        file.write('### ERROR ### \r\n')
                        file.write(traceb+'\r\n')
                    
                    path2 = join(self.location_debug_file,'error_log.txt')
                    with open(path2 , "a") as file:
                        file.write('### ERROR ### \r\n')
                        file.write(traceb+'\r\n')
                    
                    
                if extras != None:
                    print(extras)
                    with open(path , "a") as file:
                        file.write(extras+'\r')
                        
            
            # Fehler werden auf jeden Fall geloggt        
            if traceb != None:
                
                path2 = join(self.location_debug_file,'error_log.txt')
                with open(path2 , "a") as file:
                    file.write('### ERROR ### \r\n')
                    file.write(traceb+'\r\n')
                
                
                if self.path_to_project_settings != None:
                    path3 = join(self.path_to_project_settings,'error_log.txt')
                    with open(path3 , "a") as file:
                        file.write('### ERROR ### \r\n')
                        file.write(traceb+'\r\n')
                
                
                try:
                    if not self.write_debug_file:
                        print(traceb)
                except:
                    pass
            
        except Exception as e:
            try:
                print(e)
                path = join(self.location_debug_file,'organon_log_error.txt')
                with open(path , "a") as file:
                    file.write(str(e) +'\r\n')
                    file.write(traceb +'\r\n')
            except:
                pass

    def format_argues(self,argues):
        try:
            a = []
            
            for arg in argues:
                if arg in ('self','ev'):
                    continue
                inhalt = str(argues[arg])
                if 'pyuno object' in inhalt:
                    inhalt = 'pyuno object'
                a.append((arg,inhalt))
            
            a = str(a)[1:-1]
            return a
        except:
            print(tb())
            return 'Fehler'
        
        
        
    def logging_optionsfenster(self,mb):

        try:
            lang = mb.lang
            ctx = mb.ctx
            
            breite = 650
            hoehe = 190
            
            tab = 10
            tab1 = 30
            
            
            posSize_main = mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = posSize_main.X +20
            Y = posSize_main.Y +20
            Width = breite
            Height = hoehe
            
            posSize = X,Y,Width,Height
            fenster,fenster_cont = mb.erzeuge_Dialog_Container(posSize)

            
            y = 10
            
            prop_names = ('Label','FontWeight',)
            prop_values = (lang.EINSTELLUNGEN_LOGDATEI,200,)
            control, model = mb.createControl(ctx, "FixedText", tab, y, 200, 20, prop_names, prop_values)
            fenster_cont.addControl('Titel', control) 


            y += 30
            
            prop_names = ('Label','State')
            prop_values = (lang.KONSOLENAUSGABE,self.output_console)
            controlCB, model = mb.createControl(ctx, "CheckBox", tab, y, 200, 20, prop_names, prop_values)
            controlCB.setActionCommand('Konsole')
             
            y += 20
            
            prop_names = ('Label','State',)
            prop_values = (lang.ARGUMENTE_LOGGEN,self.log_args,)
            control_arg, model = mb.createControl(ctx, "CheckBox", tab1, y, 200, 20, prop_names, prop_values)
            control_arg.setActionCommand('Argumente')
            control_arg.Enable = (self.output_console == 1)
            
            y += 30
            
            prop_names = ('Label','State',)
            prop_values = (lang.LOGDATEI_ERZEUGEN,self.write_debug_file,)
            control_log, model = mb.createControl(ctx, "CheckBox", tab1, y, 200, 20, prop_names, prop_values)
            control_log.setActionCommand('Logdatei')
            control_log.Enable = (self.output_console == 1)
            
            
            y += 20
            
            prop_names = ('Label',)
            prop_values = (lang.SPEICHERORT,)
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
            fenster.setPosSize(0,0,Breite+10+tab1,0,4)
            
            y += 20
            
            prop_names = ('Label',)
            prop_values = (lang.AUSWAHL,)
            control_but, model = mb.createControl(ctx, "Button", tab1, y, 80, 20, prop_names, prop_values)
            control_but.setActionCommand('File')
            fenster_cont.addControl('Titel', control)
            
            
            # Listener setzen
            log_listener = Listener_Log_Optionen(mb,control_log,control_arg,control_path)
            
            controlCB.addActionListener(log_listener)
            fenster_cont.addControl('Konsole', controlCB)
            
            control_log.addActionListener(log_listener)
            fenster_cont.addControl('Log', control_log) 
            
            control_but.addActionListener(log_listener)
            fenster_cont.addControl('but', control_but)
            
            control_arg.addActionListener(log_listener)
            fenster_cont.addControl('arg', control_arg)
            

        except:
            print(tb())
        
from com.sun.star.awt import XActionListener
class Listener_Log_Optionen(Base, XActionListener):
    def __init__(self,mb,control_log,control_arg,control_filepath):
        self.mb = mb
        self.control_log = control_log
        self.control_arg = control_arg
        self.control_filepath = control_filepath
                
    def actionPerformed(self,ev):
        # self.mb.class_Log ersetzen durch mb

        try:
            if ev.ActionCommand == 'Konsole':
                
                self.control_log.Enable = (ev.Source.State == 1)
                self.control_arg.Enable = (ev.Source.State == 1)
                self.mb.class_Log.output_console = ev.Source.State
                self.mb.class_Log.write_debug_config_file()
                self.mb.debug = ev.Source.State
                
            elif ev.ActionCommand == 'Logdatei':
                self.mb.class_Log.write_debug_file = ev.Source.State
                self.mb.class_Log.write_debug_config_file()
                
            elif ev.ActionCommand == 'Argumente':
                self.mb.class_Log.log_args = ev.Source.State
                self.mb.class_Log.write_debug_config_file()
                
            elif ev.ActionCommand == 'File':
                Folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
                Folderpicker.execute()
                
                if Folderpicker.Directory == '':
                    return
                filepath = fileUrlToSystemPath(Folderpicker.getDirectory())
                
                self.mb.class_Log.location_debug_file = filepath
                self.control_filepath.Model.Label = filepath
                
                self.mb.class_Log.write_debug_config_file()
            
                
        except:
            print(tb())
    

    
    def disposing(self,ev):
        return False     
        
        
        
             