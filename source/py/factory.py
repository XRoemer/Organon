
# -*- coding: utf-8 -*-
import sys
from traceback import print_exc as tb
import uno
import unohelper
from os import path as PATH


####################################################
                # DEBUGGING #
####################################################

load_reload = False

platform = sys.platform
        
if load_reload:
    pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
    if platform == 'linux':
        pyPath = '/home/xgr/workspace/organonEclipse/py'
        sys.path.append(pyPath)

def pydevBrk():  
    # adjust your path 
    if platform == 'linux':
        sys.path.append('/opt/eclipse/plugins/org.python.pydev_3.3.3.201401272249/pysrc')  
    else:
        sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_3.5.0.201405201709/pysrc')  
    from pydevd import settrace
    settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
pd = pydevBrk
#pydevBrk()

def load_logging(path_to_extension):
    try:
        if load_reload:
            modul = 'log_organon'
            log_organon = load_reload_modul(modul,pyPath)  
        else:
            import log_organon
        
        class_Log = log_organon.Log(path_to_extension,pd,tb)
        class_Log.load_reload = load_reload
        debug = class_Log.debug
        log1 = class_Log.log
        return log1,class_Log,debug
    except:
        tb()


####################################################
                # SIDEBAR #
####################################################


dict_sb = {'sichtbare':['empty_project'],
           'controls':{},
           'erzeuge_sb_layout':None,
           'optionsfenster':None,
           'sb_closed':None}


from com.sun.star.ui import XUIElementFactory
class ElementFactory( unohelper.Base, XUIElementFactory ):

    def __init__( self, ctx ):
        self.ctx = ctx   
        
    def createUIElement(self,url,args):    
        cmd = url.split('/')[-1]

        try:
            xParentWindow = None
            xFrame = None
            xUIElement = None

            for arg in args:
                if arg.Name == "Frame":
                    xFrame = arg.Value
                elif arg.Name == "ParentWindow":
                    xParentWindow = arg.Value
                elif arg.Name == "Sidebar":
                    sidebar = arg.Value
            
            xUIElement = XUIPanel(self.ctx, xFrame, xParentWindow, url)

            # getting the real panel window 
            # for setting the content       
            xUIElement.getRealInterface()
            panelWin = xUIElement.Window
            
            # panelWin has to be set visible
            panelWin.Visible = True
            panelWin.Model.BackgroundColor = 14804725            
            
            conts = dict_sb['controls']
            
            if cmd in dict_sb['sichtbare']:

                conts.update({cmd:(xUIElement,sidebar,xParentWindow)})
                dict_sb.update({'controls':conts})

                if dict_sb['erzeuge_sb_layout'] != None:
                    erzeuge_sb_layout = dict_sb['erzeuge_sb_layout']
                    dict_sb['sb_closed'] = False
                    erzeuge_sb_layout(cmd,'factory')
                    

                return xUIElement
            else:
                if cmd in conts:
                    del(conts[cmd])
                return None
            
        except Exception as e:
            print('createUIElement '+ str(e))
            #pd()
            tb()
       
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
        ElementFactory,
        "org.apache.openoffice.Organon.sidebar.OrganonSidebarFactory",
        ("com.sun.star.task.Job",),) 



####################################################
                # TREEVIEW #
####################################################


from com.sun.star.lang import XSingleComponentFactory
class Factory(unohelper.Base, XSingleComponentFactory):
    """ This factory instantiate new window content. 
    Registration of this class have to be there under 
    /org.openoffice.Office.UI/WindowContentFactories/Registered/ContentFactories.
    See its schema for detail. """
    
    # Implementation name should be match with name of 
    # the configuration node and FactoryImplementation value.
    IMPLE_NAME = "xaver.roemers.organon.factory"
    SERVICE_NAMES = IMPLE_NAME,

    @classmethod
    def get_imple(klass):
        return klass, klass.IMPLE_NAME, klass.SERVICE_NAMES
    
    def __init__(self, ctx, *args):
        self.ctx = ctx
        self.load_reload_modul = load_reload_modul
        
        global path_to_extension, log, debug, class_Log
        
        #path_to_extension = __file__.decode("utf-8").split('organon.oxt')[0] + 'organon.oxt'
        path_to_extension = get_path_to_extension()
        debug = False
        
        print("factory init")

    def do(self):
        return
    
    def createInstanceWithArgumentsAndContext(self, args, ctx):
        
        try:
            CWHandler = ContainerWindowHandler(ctx)
            self.CWHandler = CWHandler
            
            win,tabs = create_window(ctx,self)
            window = self.CWHandler.window2

            start_main(window,ctx,tabs,path_to_extension,win,self)  
            
            return win 
        
        except Exception as e:
            print('Factory '+e)
            tb()



#g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(*Factory.get_imple())





from com.sun.star.beans import NamedValue

RESOURCE_URL = "private:resource/dockingwindow/9809"
EXT_ID = "xaver.roemers.organon"

def create_window(ctx,factory):

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
        self.ctx = ctx
        self.window2 = None
    
    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, obj, name):

        if name == "external_event":
            if obj == "initialize":
                self.window2 = window
                self._initialize(window)
    
    def getSupportedMethodNames(self):
        return "external_event",
    
    def _initialize(self, window):

        path_to_current = __file__.decode("utf-8")
        pyPath = path_to_current.split('factory.py')[0]
        sys.path.append(pyPath)
        
        pyPath_lang = pyPath.replace('py','languages')
        sys.path.append(pyPath_lang)

    def disposing(self, ev):
        pass
    

dict_sb.update({'CWHandler':ContainerWindowHandler(uno.getComponentContext())})


def start_main(window,ctx,tabs,path_to_extension,win,factory):

    dialog = window
        
    if load_reload:
        modul = 'menu_start'
        menu_start = load_reload_modul(modul,pyPath)  # gleichbedeutend mit: import menu_bar
    else:
        import menu_start
        
    try:
        path_to_extension = get_path_to_extension()
        log,class_Log,debug = load_logging(path_to_extension)
    except:
        tb()
            
    args = (pd,
            dialog,
            ctx,
            tabs,
            path_to_extension,
            win,
            dict_sb,
            debug,
            load_reload,
            factory,
            log,
            class_Log)
    
    Menu_Start = menu_start.Menu_Start(args)
    Menu_Start.erzeuge_Startmenu()




####################################################
                # SIDEBAR #
####################################################


from com.sun.star.lang import XComponent
from com.sun.star.ui import XUIElement, XToolPanel,XSidebarPanel, LayoutSize
from com.sun.star.ui.UIElementType import TOOLPANEL as UET_TOOLPANEL

class XUIPanel( unohelper.Base,  XSidebarPanel, XUIElement, XToolPanel, XComponent ):

    def __init__ ( self, ctx, frame, xParentWindow, url ):

        self.ctx = ctx
        self.xParentWindow = xParentWindow
        self.window = None
        
        self.height = 100
        
        self.ResourceURL = url
        self.Frame = frame
        self.Type = UET_TOOLPANEL

        
    # XUIElement
    def getRealInterface( self ):
        
        if not self.window:
            dialogUrl = "vnd.sun.star.extension://xaver.roemers.organon/factory/Dialog1.xdl"
            smgr = self.ctx.ServiceManager
            
            provider = smgr.createInstanceWithContext("com.sun.star.awt.ContainerWindowProvider",self.ctx)  
            self.window = provider.createContainerWindow(dialogUrl,"",self.xParentWindow, None)
            
        return self
    
#     @property
#     def Frame(self):
#         self.frame = frame
#      
#     @property
#     def ResourceURL(self):
#         return self.URL
#      
#     @property
#     def Type(self):
#         return UET_TOOLPANEL
    
    # XComponent
    def dispose(self):
        dict_sb['sb_closed'] = True
        pass
     
    def addEventListener(self, ev): pass
     
    def removeEventListener(self, ev): pass
     
    # XToolPanel
    def createAccessible(self, parent):
        return self
     
    @property
    def Window(self):
        return self.window
     
    # XSidebarPanel
    def getHeightForWidth(self, width):
        #print("getHeightForWidth: %s" % width)
        #return LayoutSize(0, -1, 0) # full height
        return LayoutSize(self.height, self.height, self.height)
    
    def getMinimalWidth(self,*args):
        return 100


from com.sun.star.frame import XDispatch,XDispatchProvider
class Sidebar_Options_Dispatcher(unohelper.Base,XDispatch,XDispatchProvider):

    IMPLE_NAME = "org.apache.openoffice.Organon.sidebar.ProtocolHandler"
    SERVICE_NAMES = IMPLE_NAME,
     
    @classmethod
    def get_imple(klass):
        #pydevBrk()
        return klass, klass.IMPLE_NAME, klass.SERVICE_NAMES
    
    def __init__(self,*args):
        pass
    def queryDispatches(self,*args):
        return
    def queryDispatch(self,featureURL,frameName,searchFlag):
        return self
    def dispatch(self,featureURL,*args):
        optionsfenster = dict_sb['optionsfenster']
        optionsfenster(featureURL.Path)
    def addStatusListener(self,listener,featureURL):
        #print('addStatusListener', featureURL.Path)
        return
    def removeStatusListener(self,listener,featureURL):
        #print('removeStatusListener', featureURL.Path)
        return
    

        
        
g_ImplementationHelper.addImplementation(*Sidebar_Options_Dispatcher.get_imple())

    

############################ TOOLS ###############################################################

def get_path_to_extension():
    
    ctx = uno.getComponentContext()
    pip = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")   
    ploc = pip.getPackageLocation('xaver.roemers.organon')
    pfad = uno.fileUrlToSystemPath(ploc)
    
    return pfad

# def get_office_name():
#     
#     frame = self.current_Contr.Frame
#     if 'LibreOffice' in frame.Title:
#         programm = 'LibreOffice'
#     elif 'OpenOffice' in frame.Title:
#         programm = 'OpenOffice'
#     else:
#         # Fuer Linux / OSX fehlt
#         programm = 'LibreOffice'
#     
#     return programm

def load_reload_modul(modul,pyPath):
    try:
        if pyPath not in sys.path:
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
        tb()