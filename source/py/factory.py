
# -*- coding: utf-8 -*-
import sys
import traceback
import uno
import unohelper

global oxt
oxt = False

tb = traceback.print_exc
platform = sys.platform

if oxt:
    pyPath = 'E:\\Eclipse_Workspace\\orga\\organon\\py'
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
        sys.path.append(r'C:\Users\Homer\Desktop\Programme\eclipse\plugins\org.python.pydev_3.1.0.201312121632\pysrc')  
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
        if oxt:print("factory init")
        #pydevBrk()
    def do(self):
        print('do')    
    
    def createInstanceWithArgumentsAndContext(self, args, ctx):
        if oxt:print('createInstanceWithArgumentsAndContext in Factory')
        try:
            
            CWHandler = ContainerWindowHandler(ctx)
            self.CWHandler = CWHandler
            
            win,tabs = create_window(ctx,self)
            
            window = self.CWHandler.window2

            path_to_extension = __file__.decode("utf-8").split('organon.oxt')[0] + 'organon.oxt'
            
            start_main(pydevBrk,window,ctx,tabs,path_to_extension)  
            
            return win 
        except Exception as e:
            print(e)



g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(*Factory.get_imple())





from com.sun.star.beans import NamedValue



RESOURCE_URL = "private:resource/dockingwindow/9809"
EXT_ID = "xaver.roemers.organon"

def create_window(ctx,factory):
    #print('create_window')

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
    
    #pydevBrk()
    return window,tabs


from com.sun.star.awt import XWindowListener,XActionListener,XContainerWindowEventHandler
class ContainerWindowHandler(unohelper.Base, XContainerWindowEventHandler):
    
    def __init__(self, ctx):
        if oxt:print('init ContainerWindowHandler')
        self.ctx = ctx
        self.window2 = None
        #pydevBrk()
    
    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, obj, name):
        if oxt:print('callHandlerMethod')
        #pydevBrk()
        if name == "external_event":
            if obj == "initialize":
                self.window2 = window
                self._initialize(window)
    
    def getSupportedMethodNames(self):
        return "external_event",
    
    def _initialize(self, window):
        if oxt:print('_initialize in ContainerWindowHandler')

        path_to_current = __file__.decode("utf-8")
        pyPath = path_to_current.split('factory.py')[0]
        sys.path.append(pyPath)
        
        pyPath_lang = pyPath.replace('py','languages')
        sys.path.append(pyPath_lang)

    def disposing(self, ev):
        pass
    

def start_main(pd,window,ctx,tabs,path_to_extension):

    dialog = window
    
    if oxt:
        modul = 'menu_bar'
        menu_bar = load_reload_modul(modul,pyPath)  # gleichbedeutend mit: import menu_bar
    else:
        import menu_bar
        
    Menu_Bar = menu_bar.Menu_Bar(pd,dialog,ctx,tabs,path_to_extension)
    Menu_Bar.erzeuge_Menu()

        

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
            pass#traceback.print_exc()
                            
        exec('import '+ modul)
        
        return eval(modul)
    except:
        traceback.print_exc()