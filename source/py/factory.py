# -*- coding: utf-8 -*-

'''
This modul is the entry point of the extension Organon.
It sets up the paths further needed, starts logging
and while installing for the first time, it
gets the settings of an old Organon installation,
if there was one.

The factories for the sidebar and the treeview are registered here.
The factory of the 'schalter' is registerd in schalter.py. That one creates
a button in OO's/LO's menubar for opening a docking window. 
On opening the docking window the class 'Factory' of this module is called.
(Thanks to Hanya from the OO Forum. I couldn't have done that without is help.)

The dictonary 'dict_sb' is used to keep the connection between the sidebar 
and the treeview.

The class Factory starts the modul menu_start, which creates the first 
window of Organon.

'''


import sys
from traceback import print_exc as tb
import uno
import unohelper
from os import path as PATH, listdir
import inspect
import json,copy
from codecs import open as codecs_open



####################################################
            # DEBUGGING / REMOTE CONTROL #
####################################################

load_reload = False

if load_reload:
    sys.dont_write_bytecode = True

platform = sys.platform
        
if load_reload:
    pyPath = 'H:\\Programmierung\\Eclipse_Workspace\\Organon\\source\\py'
    if platform == 'linux':
        pyPath = '/media/windows/Organon/source/py'
    sys.path.append(pyPath)

def pydevBrk():  
    # adjust your path 
    if platform == 'linux':
        sys.path.append('/home/xgr/.eclipse/org.eclipse.platform_4.4.1_1473617060_linux_gtk_x86_64/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
    else:
        sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
    from pydevd import settrace
    settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
pd = pydevBrk



def load_logging(path_to_extension,log_config):
    
    import log_organon
    
    class_Log = log_organon.Log(path_to_extension,pd,tb,log_config)
    debug = class_Log.debug
    log1 = class_Log.log
    return log1,class_Log,debug


####################################################
            # DEBUGGING END #
####################################################


def get_paths():
    # PackageInformationProvider doesn't work at this stage
    # of initialising the extension.
    # So this is a workaround
    path_to_current = inspect.stack()[0][1]
    pfad = inspect.stack()[0][1]
    
    def get_parent(path):
        return PATH.dirname(path)
    
    while 'organon.oxt' in pfad:
        pfad = get_parent(pfad)
    
    package_folder = get_parent(pfad)
    path_to_extension = PATH.join(pfad,'organon.oxt')
    extension_folder = PATH.basename(pfad)
    
    return package_folder, extension_folder, path_to_extension



def get_organon_settings(path_to_extension):
    
    json_pfad = PATH.join(path_to_extension,'organon_settings.json')
        
    with codecs_open(json_pfad) as data:  
        content = data.read().decode()  
        settings_orga = json.loads(content)
    
    return settings_orga



class Take_Over_Old_Settings():
    
    '''
    Settings have to be loaded while the old extension is still available
    on the harddrive. During installation the old extension will be deleted.
    So getting both needs to be done at the very beginning.
    When an old extension is found, the newly created settings will
    be extended by the old ones.
    '''
    
    def get_settings_of_previous_installation(self,package_folder, extension_folder):
        try:
            dirs = [name for name in listdir(package_folder) if PATH.isdir(PATH.join(package_folder, name))]
            dirs.remove(extension_folder)
            
            files = None   
            organon_in_files = False    
             
            for d in dirs:
                files = listdir(PATH.join(package_folder,d))
                if 'organon.oxt' in files:
                    organon_in_files = True
                    break
            
            if files == None or organon_in_files == False :
                return None

            json_pfad_alt = PATH.join(package_folder,d,'organon.oxt','organon_settings.json')
            
            with codecs_open(json_pfad_alt) as data:  
                content = data.read().decode()  
                settings_orga_prev = json.loads(content)
            
            return settings_orga_prev
    
        except Exception as e:
            return None
    
    
    def run(self,st_dict,old_dict):
        
        try:
            standard_dict = copy.deepcopy(st_dict)
            dict_as_list = []
            self.dict_to_list(old_dict,dict_as_list)
            
            # Update Settings
            for n in dict_as_list:
                if 'designs' not in n:
                    self.exchange_values(old_dict,standard_dict,n)

            # Update Designs
            for k in old_dict['designs']:
    
                if k not in standard_dict['designs']:
                    standard_dict['designs'].update( { k : old_dict['designs'][k] } )
                else:
                    create_new = self.compare_dicts(old_dict['designs'][k],standard_dict['designs'][k])
                    if create_new:
                        k2 = copy.deepcopy(k)
                        while k2 in standard_dict['designs']:
                            k2 = k2 + '_old'
                        standard_dict['designs'].update( { k2 : old_dict['designs'][k] } )
        
            return standard_dict   
        except:
            return None    
    
    
    def compare_dicts(self,dict1,dict2):
        try:
            ungleich = False
            for k in dict1:
                if dict1[k] != dict2[k]:
                    ungleich = True
            return ungleich
        except Exception as e:
            return True
    
    def dict_to_list(self,odict,olist,predecessor=[]):
    
        for k in odict:
            value = odict[k]
            pre = predecessor[:]
                            
            if isinstance(value, dict):
                pre.append(k)
                self.dict_to_list(value,olist,pre)
            else:
                olist.append(predecessor+[k])

    def exchange_values(self,old_dict,standard,olist):

        # Set a given data in a dictionary with position provided as a list
        def setInDict(dataDict, mapList, value): 
            for k in mapList[:-1]: dataDict = dataDict[k]
            dataDict[mapList[-1]] = value
        
        # Get a given data from a dictionary with position provided as a list
        def getFromDict(dataDict, mapList):    
            for k in mapList: dataDict = dataDict[k]
            return dataDict

        value = getFromDict(old_dict,olist)
        try:
            # A value which is not member of the dict is ignored
            setInDict(standard,olist,value)
        except:
            pass


###### GET PATHS ###### 
package_folder, extension_folder, path_to_extension = get_paths()

###### GET SETTINGS ######
settings_orga = get_organon_settings(path_to_extension)
settings_orga_prev = Take_Over_Old_Settings().get_settings_of_previous_installation(package_folder, extension_folder) 

if settings_orga_prev != None:
    
    # Die Einstellungen einzeln zu kopieren, ist an dieser Stelle eigentlich noch
    # zu aufwendig, man koennte sie auch einfach direkt kopieren. Wenn aber neue Eintraege
    # in den Settings hinzukommen, ist dies eine sichere Methode, bereits gesetzte Settings
    # zu uebernehmen, waehrend die neuen ihren default-Wert aus der Installations Datei behalten.
    neuer_dict = Take_Over_Old_Settings().run(settings_orga,settings_orga_prev)
    
    if neuer_dict != None:
        settings_orga = neuer_dict
     
        json_pfad = PATH.join(path_to_extension,'organon_settings.json')
        with open(json_pfad, 'w') as outfile:
            json.dump(neuer_dict, outfile,indent=4, separators=(',', ': '))



###### START LOGGING ######   
try:   
    log,class_Log,debug = load_logging(path_to_extension,settings_orga['log_config'])        
except:
    sys.path.append(PATH.join(path_to_extension,'py'))
    log,class_Log,debug = load_logging(path_to_extension,settings_orga['log_config'])  


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
                elif arg.Name == "Theme":
                    theme = arg.Value

            xUIElement = XUIPanel(self.ctx, xFrame, xParentWindow, url, theme)

            # getting the real panel window 
            # for setting the content       
            xUIElement.getRealInterface()
            panelWin = xUIElement.Window
            
            # panelWin has to be set visible
            panelWin.Visible = True
            #panelWin.Model.BackgroundColor = 14804725           
            
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
            #print('createUIElement '+ str(e))
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



def set_konst(dialog):
    
    try:
        
        sett_orgn = settings_orga
        KONST.FARBE_HF_HINTERGRUND = sett_orgn['organon_farben']['hf_hintergrund']
        KONST.FARBE_MENU_HINTERGRUND = sett_orgn['organon_farben']['menu_hintergrund']
        
        KONST.FARBE_MENU_SCHRIFT = sett_orgn['organon_farben']['menu_schrift']
        KONST.FARBE_SCHRIFT_ORDNER = sett_orgn['organon_farben']['schrift_ordner']
        KONST.FARBE_SCHRIFT_DATEI = sett_orgn['organon_farben']['schrift_datei']
        
        KONST.FARBE_AUSGEWAEHLTE_ZEILE = sett_orgn['organon_farben']['ausgewaehlte_zeile']
        KONST.FARBE_EDITIERTE_ZEILE = sett_orgn['organon_farben']['editierte_zeile']
        KONST.FARBE_GEZOGENE_ZEILE  = sett_orgn['organon_farben']['gezogene_zeile']
        
        KONST.FARBE_GLIEDERUNG  = sett_orgn['organon_farben']['gliederung']
        
        KONST.FARBE_TRENNER_HINTERGRUND   = sett_orgn['organon_farben']['trenner_farbe_hintergrund']
        KONST.FARBE_TRENNER_SCHRIFT       = sett_orgn['organon_farben']['trenner_farbe_schrift']
         
        dialog.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND

    except:
        tb()



import konstanten as KONST
def start_main(window,ctx,tabs,path_to_extension,win,factory):

    try:   
        dialog = window
        
        import menu_start

        set_konst(dialog)

        args = (pd,
                dialog,
                ctx,
                tabs,
                path_to_extension,
                win,
                dict_sb,
                debug,
                factory,
                log,
                class_Log,
                KONST,
                settings_orga)
        
        Menu_Start = menu_start.Menu_Start(args)
        Menu_Start.erzeuge_Startmenu()
    except Exception as e:
        print(e)
        log(inspect.stack,tb())




####################################################
                # SIDEBAR #
####################################################


from com.sun.star.lang import XComponent
from com.sun.star.ui import XUIElement, XToolPanel,XSidebarPanel, LayoutSize
from com.sun.star.ui.UIElementType import TOOLPANEL as UET_TOOLPANEL

class XUIPanel( unohelper.Base,  XSidebarPanel, XUIElement, XToolPanel, XComponent ):

    def __init__ ( self, ctx, frame, xParentWindow, url ,theme):

        self.ctx = ctx
        self.xParentWindow = xParentWindow
        self.window = None
        
        self.height = 100
        
        self.ResourceURL = url
        self.Frame = frame
        self.Type = UET_TOOLPANEL
        self.Theme = theme
        
    # XUIElement
    def getRealInterface( self ):
        
        if not self.window:
            dialogUrl = "vnd.sun.star.extension://xaver.roemers.organon/factory/Dialog1.xdl"
            smgr = self.ctx.ServiceManager
            
            provider = smgr.createInstanceWithContext("com.sun.star.awt.ContainerWindowProvider",self.ctx)  
            self.window = provider.createContainerWindow(dialogUrl,"",self.xParentWindow, None)
            
        return self
    
    
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

    
