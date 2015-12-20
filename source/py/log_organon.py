# -*- coding: utf-8 -*-

from time import clock,sleep
from os.path import join,exists,basename
from unohelper import Base
from uno import fileUrlToSystemPath
from traceback import print_exc as tb
import uno
import inspect
from codecs import open as codecs_open


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
    
    def __init__(self,path_to_extension,pdX,tbX,log_config):  
        
        global debug,pd,tb
        pd = pdX
        tb = tbX
        
        self.path_to_extension = path_to_extension
        self.timer_start = clock()
        self.path_to_project_settings = None
        
        # Default Debug Settings
        if log_config['location_debug_file'] == '':
            self.location_debug_file = path_to_extension
        else:
            # Fehlt: Pruefen, ob der Pfad existiert
            self.location_debug_file = log_config['location_debug_file']
        
        self.debug = log_config['output_console']
        self.output_console = log_config['output_console']
        self.write_debug_file = log_config['write_debug_file']
        self.log_args = log_config['log_args']
        
         
        if self.debug and self.write_debug_file:
            self.schreibe_logfile_kopfzeile()

    
    def schreibe_logfile_kopfzeile(self):
        try:

            path = join(self.location_debug_file,'organon_log.txt')
            
            text = '################\r\n' \
                   '\r\n' \
                    'Organon opened\r\n' \
                    '\r\n'\
                    '################\r\n'
                    
            with codecs_open( path, "a","utf-8") as file:
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
                xclass = info[1][0].f_locals['self'].__class__.__name__
            except:
                # Wenn aus einer Methode ohne Klasse gerufen wird, existiert kein 'self'
                xclass = str(info[1][0])
                
            try:
                modul = basename(info[1][1]).split('.')[0]
            except:
                modul = ''

#             if xclass in ('ViewCursor_Selection_Listener'):
#                 return
#             
#             if function in ('mouseEntered','mouseExited','entferne_Trenner'):
#                 return
            
            if function in ('mouseEntered','mouseExited'):
                return
            
            

            if self.log_args:
                string = '{0: <8.8}  {1: <12.12}  {2: <24.24}  {3: <34.34}  ( caller: {4: <44.44}  args:{5}'.format(self.debug_time(),modul,xclass,function,call,argues)
            else:
                string = '{0: <8.8}  {1: <12.12}  {2: <24.24}  {3: <34.34}  ( caller: {4: <44.44}'.format(self.debug_time(),modul,xclass,function,call)

            try:
                print(string)
            except:
                pass
            
            if self.write_debug_file:
                path = join(self.location_debug_file,'organon_log.txt')
                with codecs_open( path, "a","utf-8") as file:
                    file.write(string+'\n')
                
                if traceb != None:
                    print(traceb)
                    
                    with codecs_open( path, "a","utf-8") as file:
                        file.write('### ERROR ### \r\n')
                        try:
                            file.write(traceb+'\r\n')
                        except:
                            print('ERROR ON WRITING ERROR TO FILE')
                            file.write(str(traceb)+'\r\n')
    
                if extras != None:
                    print(extras)
                    with codecs_open( path, "a","utf-8") as file:
                        file.write(extras+'\r\n')
            
#             self.suche()        
#             # HELFER          
#             nachricht = self.suche()
#             if nachricht != None:
#                 with open(path , "a") as file:
#                     file.write(nachricht+'\r\n')
            
#             self.helfer()
            
            # Fehler werden auf jeden Fall geloggt        
            if traceb != None:
                
                path2 = join(self.location_debug_file,'error_log.txt')
                with codecs_open( path2, "a","utf-8") as file:
                    file.write('### ERROR ###1 \r\n')
                    file.write(traceb+'\r\n')
                
                try:
                    if not self.write_debug_file:
                        print(traceb)
                except:
                    pass
            
        except Exception as e:
            try:
                print(str(e))
                print(tb())
                path = join(self.location_debug_file,'organon_log_error.txt')
                with codecs_open( path, "a","utf-8") as file:
                    file.write(str(e) +'\r\n')
                    file.write(str(tb()) +'\r\n')
            except:
                print(tb())
                with codecs_open( path, "a","utf-8") as file:
                    file.write(str(tb()) +'\r\n')
                    

    def format_argues(self,argues):
        try:
            a = []
            
            for arg in argues:
                if arg in ('self','ev'):
                    continue
                try:
                    inhalt = str(argues[arg])
                except Exception as e:
                    inhalt = 'Fehler: ' + str(e) + '\r\n'
                    
                if 'pyuno object' in inhalt:
                    inhalt = 'pyuno object'
                a.append((arg,inhalt))
            
            # aendern
            #a = unicode(a)[1:-1]
            a = a[1:-1]
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
            return e
        
        return None
    
    def helfer(self):
        try:
            pass
        except Exception as e:
            print(tb())
            return e
        
        
        
        
        
             