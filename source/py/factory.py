
# -*- coding: utf-8 -*-
import sys
import traceback
import uno
import unohelper


debug = False

tb = traceback.print_exc
platform = sys.platform

if debug:
    pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
    if platform == 'linux':
        pyPath = '/home/xgr/Arbeitsordner/organon/py'
        sys.path.append(pyPath)


from com.sun.star.lang import (XSingleComponentFactory, 
    XServiceInfo)



def pydevBrk():  
    # adjust your path 
    if platform == 'linux':
        sys.path.append('/opt/eclipse/plugins/org.python.pydev_3.3.3.201401272249/pysrc')  
    else:
        sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_3.5.0.201405201709/pysrc')  
    from pydevd import settrace
    settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
#pydevBrk()

class Factory(unohelper.Base, XSingleComponentFactory):
    """ This factory instantiate new window content. 
    Registration of this class have to be there under 
    /org.openoffice.Office.UI/WindowContentFactories/Registered/ContentFactories.
    See its schema for detail. """
    
    # Implementation name should be match with name of 
    # the configuration node and FactoryImplementation value.
    IMPLE_NAME = "xaver.roemers.organon.factory"
    SERVICE_NAMES = IMPLE_NAME,
    #pydevBrk()
    @classmethod
    def get_imple(klass):
        #pydevBrk()
        return klass, klass.IMPLE_NAME, klass.SERVICE_NAMES
    
    def __init__(self, ctx, *args):
        self.ctx = ctx
        if debug:print("factory init")
        #pydevBrk()
    def do(self):
        return
    
    def createInstanceWithArgumentsAndContext(self, args, ctx):
        #if debug:print('createInstanceWithArgumentsAndContext in Factory')
        try:
            
            CWHandler = ContainerWindowHandler(ctx)
            self.CWHandler = CWHandler
            
            win,tabs = create_window(ctx,self)
            
            window = self.CWHandler.window2

            path_to_extension = __file__.decode("utf-8").split('organon.oxt')[0] + 'organon.oxt'
            
            start_main(pydevBrk,window,ctx,tabs,path_to_extension,win,debug)  
            
            return win 
        except Exception as e:
            print(e)
            traceback.print_exc()



g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(*Factory.get_imple())





from com.sun.star.beans import NamedValue

RESOURCE_URL = "private:resource/dockingwindow/9809"
EXT_ID = "xaver.roemers.organon"

def create_window(ctx,factory):
    #if debug: print('create_window')

    dialog1 = "vnd.sun.star.extension://xaver.roemers.organon/factory/Dialog1.xdl"

    tabs = ctx.getServiceManager().createInstanceWithContext("com.sun.star.comp.framework.TabWindowService",ctx)
    id = tabs.insertTab() # Create new tab, return value is tab id
    # Valid properties are: 
    # Title, ToolTip, PageURL, EventHdl, Image, Disabled.
    v1 = NamedValue("PageURL", dialog1)
    v2 = NamedValue("Title", "ORGANON")
    v3 = NamedValue("EventHdl", factory.CWHandler)
    tabs.setTabProps(id, (v1, v2, v3))
    tabs.activateTab(id) 
    
    tabs.Window.setProperty('Name','ProjektFenster')
    
    window = tabs.Window # real window
    
    return window,tabs


from com.sun.star.awt import XWindowListener,XActionListener,XContainerWindowEventHandler
class ContainerWindowHandler(unohelper.Base, XContainerWindowEventHandler):
    
    def __init__(self, ctx):
        #if debug:print('init ContainerWindowHandler')
        self.ctx = ctx
        self.window2 = None
        #pydevBrk()
    
    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, obj, name):
        #if debug:print('callHandlerMethod')
        #pydevBrk()
        if name == "external_event":
            if obj == "initialize":
                self.window2 = window
                self._initialize(window)
    
    def getSupportedMethodNames(self):
        return "external_event",
    
    def _initialize(self, window):
        #if debug:print('_initialize in ContainerWindowHandler')

        path_to_current = __file__.decode("utf-8")
        pyPath = path_to_current.split('factory.py')[0]
        sys.path.append(pyPath)
        
        pyPath_lang = pyPath.replace('py','languages')
        sys.path.append(pyPath_lang)

    def disposing(self, ev):
        pass
    

def start_main(pd,window,ctx,tabs,path_to_extension,win,debug):

    dialog = window
        
    if debug:
        modul = 'menu_start'
        menu_start = load_reload_modul(modul,pyPath)  # gleichbedeutend mit: import menu_bar
    else:
        import menu_start

    Menu_Start = menu_start.Menu_Start(pd,dialog,ctx,tabs,path_to_extension,win,debug)
    Menu_Start.erzeuge_Startmenu()

        

############################ TOOLS ###############################################################

def load_reload_modul(modul,pyPath):
    try:
        sys.path.append(pyPath)
        
        exec('import '+ modul)
        del(sys.modules[modul])
        try:
            import shutil

            if 'LibreOffice' in sys.executable:
                if platform == 'linux':
                    shutil.rmtree(pyPath+'/__pycache__')
                else:
                    shutil.rmtree(pyPath+'\\__pycache__')
                
            elif 'OpenOffice' in sys.executable:
                pass
        except:
            pass
                            
        exec('import '+ modul)
        
        return eval(modul)
    except:
        traceback.print_exc()