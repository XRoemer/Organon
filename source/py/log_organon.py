# -*- coding: utf-8 -*-

from time import clock,sleep
from os.path import join,exists
from unohelper import Base
from uno import fileUrlToSystemPath
from traceback import print_exc as tb
import uno
import inspect


#####  DEBUGGING ##########
import sys
platform = sys.platform
         
def pydevBrk():  
    # adjust your path 
    if platform == 'linux':
        sys.path.append('/opt/eclipse/plugins/org.python.pydev_3.8.0.201409251235/pysrc')  
    else:
        sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_3.5.0.201405201709/pysrc')  
    from pydevd import settrace
    settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
pd = pydevBrk 
#####  DEBUGGING END ##########



class Log():
    
    def __init__(self,path_to_extension,pdX,tbX):  
        
        global debug,pd,tb
        pd = pdX
        tb = tbX
        
        self.path_to_extension = path_to_extension
        self.debug = False  
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
        
        #print(string)
        
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
            
            text = '################\r\n' \
                   '\r\n' \
                    'Organon opened\r\n' \
                    '\r\n'\
                    '################\r\n'
                    
            with open(path , "a") as file:
                file.write(text)
                
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
            try:
                modul = info[1][0].f_locals['self'].__class__.__name__
            except:
                # Wenn aus einer Methode ohne Klasse gerufen wird, existiert kein 'self'
                modul = str(info[1][0])
                

            if modul in ('ViewCursor_Selection_Listener'):
                return
            
            if function in ('mouseEntered','mouseExited','entferne_Trenner'):
                return

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
                        try:
                            file.write(traceb+'\r\n')
                        except:
                            print('ERROR ON WRITING ERROR TO FILE')
                            file.write(str(traceb)+'\r\n')
                            
                    
                    path2 = join(self.location_debug_file,'error_log.txt')
                    with open(path2 , "a") as file:
                        file.write('### ERROR ### \r\n')
                        file.write(traceb+'\r\n')
                    
                    
                if extras != None:
                    print(extras)
                    with open(path , "a") as file:
                        file.write(extras+'\r')
            
            #self.suche()        
#             # HELFER          
#             nachricht = self.suche()
#             if nachricht != None:
#                 with open(path , "a") as file:
#                     file.write(nachricht+'\r\n')
            
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
                print(tb())
                path = join(self.location_debug_file,'organon_log_error.txt')
                with open(path , "a") as file:
                    file.write(str(e) +'\r\n')
                    file.write(str(tb()) +'\r\n')
            except:
                print(tb())
                with open(path , "a") as file:
                    file.write(str(tb()) +'\r\n')
                    

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

         
    def suche(self):  
        # Fuer eine Fehlersuche, die bei jedem Methodenaufruf gestartet wird
        
        try:
            ctx = uno.getComponentContext()  
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            doc = desktop.getCurrentComponent()
            
            secs = doc.TextSections
            el_names = secs.ElementNames
            
            for n in el_names:
                if 'Bereich' in n:
                    pd()
            
        except Exception as e:
            print(tb())
            #pd()
            return e
        
        return None
        
        
             