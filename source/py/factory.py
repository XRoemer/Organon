# -*- coding: utf-8 -*-
'''
This modul is the entry point of the extension Organon.
It sets up the paths further needed, starts logging
and while installing for the first time, it
gets the settings of an old Organon installation,
if there is one.

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
from traceback import format_exc as tb
import uno
import unohelper
from os import path as PATH, listdir
import inspect
import json,copy
from codecs import open as codecs_open


####################################################
            # DEBUGGING / REMOTE CONTROL #
####################################################

load_reload = True

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
        # Ubuntu
        #sys.path.append('/home/xgr/.eclipse/org.eclipse.platform_4.4.1_1473617060_linux_gtk_x86_64/plugins/org.python.pydev_4.0.0.201504132356/pysrc') 
        # Fedora
        sys.path.append('/root/.p2/pool/plugins/org.python.pydev_4.4.0.201510052309/pysrc')     
    else:
        sys.path.append(r'C:/Users/Homer/.p2/pool/plugins/org.python.pydev_4.4.0.201510052309/pysrc')  
    from pydevd import settrace
    settrace('localhost', port=5678, stdoutToServer=True, stderrToServer=True) 
pd = pydevBrk

#pd()


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
        try:
            content = data.read()
            settings_orga = json.loads(content)  
        except Exception as e:
            with codecs_open(json_pfad) as data:  
                content = data.read().decode()  
                settings_orga = json.loads(content)
                
    return settings_orga





###### GET PATHS ###### 
package_folder, extension_folder, path_to_extension = get_paths()

###### GET SETTINGS ######
sys.path.append(PATH.join(path_to_extension,'py'))


class Take_Over_Old_Settings():
    '''
    Settings have to be loaded while the old extension is still available
    on the harddrive. During installation the old extension will be deleted.
    So getting both needs to be done at the very beginning.
    When an old extension is found, the newly created settings will
    be extended by the old ones.
    '''
    def __init__(self):
        pass
    
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
    
    designs = []
    fehlende = []

    def _update_designs(self,a,b):
        fehlende = set(b['designs']).difference( set(a['designs']) )
        standard = copy.deepcopy(a['designs']['Standard'])
        for f in fehlende:
            a['designs'].update({f:standard})
        
        return list(b['designs']), fehlende
        
    
    def _compare_design(self,a1,b1): 
        for k in b1:
            if k in a1:
                if b1[k] != a1[k]:
                    return True  
        return False
    
    def _treat_design(self,a,b,key,path):
        
        if key in self.fehlende:
            # Wenn Design nur im alten Dict vorhanden war,
            # wird es direkt uebernommen
            self.merge(a[key], b[key], path + [str(key)])
        else:
            # Wenn Designs gleichen Namens sich unterscheiden,
            # wird eine neue Version "_old" eingefuegt
            ungleich = self._compare_design(a[key],b[key])
            if ungleich:
                
                k = key
                while k in a:
                    k = k + '_old'
                    
                standard = copy.deepcopy(a['Standard'])
                a[k] = standard
                self.merge(a[k], b[key], path + [str(key)])
            else:
                pass
        
        
    
    def merge(self,a, b, path=None):
        '''
        This method is an adjusted version from:
        http://stackoverflow.com/questions/7204805/python-dictionaries-of-dictionaries-merge
        merges b into a
        '''
        if path is None: path = []
        
        try:
            for key in b:
                if key in a:
                    if key == 'zuletzt_geladene_Projekte':
                        a[key] = b[key]
                    elif isinstance(a[key], dict) and isinstance(b[key], dict):
                        if key in self.designs:
                            self._treat_design(a, b, key, path)
                        elif key == 'designs':
                            self.designs,self.fehlende = self._update_designs(a,b)
                            self.merge(a[key], b[key], path + [str(key)])
                        else:
                            self.merge(a[key], b[key], path + [str(key)])
                    elif a[key] == b[key]:
                        pass # same leaf value
                    elif isinstance(a[key], list) and isinstance(b[key], list):
                        for idx, val in enumerate(b[key]):
                            a[key][idx] = self.merge(a[key][idx], b[key][idx], path + [str(key), str(idx)])
                    else:
                        # ueberschreiben der defaults mit alten Werten
                        a[key] = b[key]
                else:
                    # hier werden nur in b vorhandene keys gesetzt
                    # daher werden auch alte designs mit eigenem Namen ignoriert
                    
                    # nur vorhandene shortcuts werden in den dict eingetragen. Daher
                    # existiert in a kein key von b und muss hier gesetzt werden.
                    if 'shortcuts' in path:
                        a[key] = b[key]
                    
            
            return a
        except Exception as e:
            print(tb())
            return None
    
    # wird nicht verwendet
    def dict_to_list(self,odict,olist,predecessor=[]):
    
        for k in odict:
            value = odict[k]
            pre = predecessor[:]
                            
            if isinstance(value, dict):
                pre.append(k)
                self.dict_to_list(value,olist,pre)
            else:
                olist.append(predecessor+[k])
                
    # wird nicht verwendet
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



TO = Take_Over_Old_Settings
settings_orga = get_organon_settings(path_to_extension)
settings_orga_prev = TO().get_settings_of_previous_installation(package_folder, extension_folder) 

if settings_orga_prev != None:
    
    # Die Einstellungen einzeln zu kopieren, ist an dieser Stelle eigentlich noch
    # zu aufwendig, man koennte sie auch einfach direkt kopieren. Wenn aber neue Eintraege
    # in den Settings hinzukommen, ist dies eine sichere Methode, bereits gesetzte Settings
    # zu uebernehmen, waehrend die neuen ihren default-Wert aus der Installations Datei behalten.
    # Nicht mehr benoetigte Settings werden entfernt.
    
    neuer_dict = TO().merge(settings_orga,settings_orga_prev)
    
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


###### SET CONSTANTS ######
import konstanten as KONST
def set_konst():
    
    try:
        sett = settings_orga
        
        # ORGANON DESIGN
        KONST.FARBE_HF_HINTERGRUND      = sett['organon_farben']['hf_hintergrund']
        KONST.FARBE_MENU_HINTERGRUND    = sett['organon_farben']['menu_hintergrund']
        
        KONST.FARBE_MENU_SCHRIFT        = sett['organon_farben']['menu_schrift']
        KONST.FARBE_SCHRIFT_ORDNER      = sett['organon_farben']['schrift_ordner']
        KONST.FARBE_SCHRIFT_DATEI       = sett['organon_farben']['schrift_datei']
        
        KONST.FARBE_AUSGEWAEHLTE_ZEILE  = sett['organon_farben']['ausgewaehlte_zeile']
        KONST.FARBE_EDITIERTE_ZEILE     = sett['organon_farben']['editierte_zeile']
        KONST.FARBE_GEZOGENE_ZEILE      = sett['organon_farben']['gezogene_zeile']
        
        KONST.FARBE_GLIEDERUNG          = sett['organon_farben']['gliederung']
        
        KONST.FARBE_TRENNER_HINTERGRUND = sett['organon_farben']['trenner_farbe_hintergrund']
        KONST.FARBE_TRENNER_SCHRIFT     = sett['organon_farben']['trenner_farbe_schrift']
        
        KONST.FARBE_TABS_HINTERGRUND    = sett['organon_farben']['tabs_hintergrund']
        KONST.FARBE_TABS_SCHRIFT        = sett['organon_farben']['tabs_schrift']
        KONST.FARBE_TABS_SEL_HINTERGRUND= sett['organon_farben']['tabs_sel_hintergrund']
        KONST.FARBE_TABS_SEL_SCHRIFT    = sett['organon_farben']['tabs_sel_schrift']
        KONST.FARBE_TABS_TRENNER        = sett['organon_farben']['tabs_trenner']
        KONST.FARBE_LINIEN              = sett['organon_farben']['linien']
        KONST.FARBE_DEAKTIVIERT         = sett['organon_farben']['deaktiviert']
        
    except Exception as e:
        log(inspect.stack,tb())

set_konst()

####################################################
                # SIDEBAR #
####################################################


dict_sb = {
           'controls' : {},
           'erzeuge_sb_layout' : None,
           'setze_sidebar_design'  : None,
           'design_gesetzt' : False,
           'sb_closed' : True,
           'seitenleiste' : None,
           'orga_sb' : None
           }


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
            
            conts = dict_sb['controls']
            
            if cmd not in conts:
                conts.update({cmd:(xUIElement,sidebar,xParentWindow,args)})
                dict_sb.update({'controls':conts})
            else:
                cont = dict_sb['controls'][cmd][0]
                cont.dispose()
                
            conts.update({cmd:(xUIElement,sidebar,xParentWindow,args)})
            dict_sb.update({'controls':conts})
            
            if dict_sb['erzeuge_sb_layout'] != None:
 
                dict_sb['sb_closed'] = False
                dict_sb['erzeuge_sb_layout']()
            else:
                pos_y = 10
                height = 50 
                width = 282 
                panelWin = xUIElement.Window

                text = u'No Project loaded'
                prop_names = ('Label',)
                prop_values = (text,)
                control, model = self.createControl(self.ctx, "FixedText", 10, pos_y, width, height, prop_names, prop_values)  
                panelWin.addControl('Synopsis', control)
        
                        
            return xUIElement
     
        except Exception as e:
            log(inspect.stack,tb())
            
            
    def createControl(self,ctx,type,x,y,width,height,names,values):
        try:
            smgr = ctx.getServiceManager()
            ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
            ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
            ctrl_model.setPropertyValues(names,values)
            ctrl.setModel(ctrl_model)
            ctrl.setPosSize(x,y,width,height,15)
            return (ctrl, ctrl_model)
        except:
            log(inspect.stack,tb())
       
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
        ElementFactory,
        "org.apache.openoffice.Organon.sidebar.OrganonSidebarFactory",
        ("com.sun.star.task.Job",),) 



####################################################
                # TREEVIEW #
####################################################
def get_parent():
    
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    
    if desktop.CurrentFrame:
        parent = desktop.CurrentFrame.ContainerWindow
    else:
        enum = desktop.Components.createEnumeration()
        comps = []
        
        while enum.hasMoreElements():
            comps.append(enum.nextElement())

        doc = comps[0]
        parent = doc.CurrentController.Frame.ContainerWindow
    
    return parent
    

def erzeuge_Dialog_Container(posSize,Flags=1+32+64+128):
                
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    
    X,Y,Width,Height = posSize
    
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
    oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)

    # Create Uno Struct
    oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
    oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
    # global oWindow
    oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "CONTAINER")
    oWindowDesc.WindowServiceName = ""
    oWindowDesc.Parent = get_parent()
    oWindowDesc.ParentIndex = -1
    oWindowDesc.WindowAttributes = Flags # Flags fuer com.sun.star.awt.WindowAttribute

    oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
    oReturnValue, oRect = oXIdlClass.createObject(None)
    oRect.X = X
    oRect.Y = Y
    oRect.Width = Width 
    oRect.Height = Height 
    
    oWindowDesc.Bounds = oRect

    # create window
    oWindow = toolkit.createWindow(oWindowDesc)
     
    # create frame for window
    oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame",ctx)
    oFrame.initialize(oWindow)
    oFrame.setCreator(desktop)
    oFrame.activate()

    # create new control container
    cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
    cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
    cont_model.BackgroundColor = KONST.FARBE_ORGANON_FENSTER  # 9225984
    
    #cont_model.ForegroundColor = KONST.FARBE_SCHRIFT_DATEI
    cont.setModel(cont_model)
    # need createPeer just only the container
    cont.createPeer(toolkit, oWindow)
    #cont.setPosSize(0, 0, 0, 0, 15)

    oFrame.setComponent(cont, None)
    
    # PosSize muss erneut gesetzt werden, um die Anzeige zu erneuern,
    # sonst bleibt ein Teil des Fensters schwarz
    oWindow.setPosSize(0,0,Width,Height,12)

    return oWindow,cont



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
        '''
        Initialising the factory which is responsible for creating
        Organon's dialog window.
        
        This is also the first available place to set the design.
        Before no services can be created, later on the dialog window's background
        isn't editable anymore.
        '''

        if debug: log(inspect.stack)
        self.debug = debug
        
        self.ctx = ctx
        
        print("factory init")
        
        if settings_orga['organon_farben']['design_office']:
            set_app_style()
            
    
    def createInstanceWithArgumentsAndContext(self, args, ctx):
        if self.debug: log(inspect.stack)
        
        try:
            self.pypath_erweitern()
            
            posSize = 0,0,0,0
            win,cont = erzeuge_Dialog_Container(posSize)
            self.start_main(cont,ctx,path_to_extension,win)
            
            return win
        
        except Exception as e:
            print(str(e))
            log(inspect.stack,tb())
    
    
    def pypath_erweitern(self):
        if self.debug: log(inspect.stack)
        
        try:
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            frame = desktop.Frames.getByIndex(0)
            
            path_to_current = __file__.decode("utf-8")
            pyPath = path_to_current.split('factory.py')[0]
            sys.path.append(pyPath)
            
            pyPath_lang = pyPath.replace('py','languages')
            sys.path.append(pyPath_lang)
            
        except Exception as e:
            log(inspect.stack,tb())
            
    
    def start_main(self,window,ctx,path_to_extension,win):
        if self.debug: log(inspect.stack)
        
        try:
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            frame = desktop.Frames.getByIndex(0)
            
            if desktop.ActiveFrame != None:
                frame = desktop.ActiveFrame
            
            try:
                if settings_orga['organon_farben']['design_office'] and g.geladen == 'neu':
                    
                    g.geladen = 'alt'
                    
                    listener = Listener_To_Restart_Component()
                    
                    if 'OpenOffice' in frame.Title : 
                        listener.win = window
                        window.addWindowListener(listener)    
                
                    if 'LibreOffice' in frame.Title:
                        eventb = ctx.getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
                        listener.eventb= eventb
                        eventb.addDocumentEventListener(listener)
                        
                    return
                    
            except Exception as e:
                log(inspect.stack,tb()) 
            
            dialog = window
            debug = settings_orga['log_config']['output_console']
    
            import menu_start
            
            args = (pd,
                    dialog,
                    ctx,
                    path_to_extension,
                    win,
                    dict_sb,
                    debug,
                    self,
                    log,
                    class_Log,
                    KONST,
                    settings_orga)
            
            Menu_Start = menu_start.Menu_Start(args)
            Menu_Start.erzeuge_Startmenu()
    
        except Exception as e:
            log(inspect.stack,tb())
    

g_ImplementationHelper.addImplementation(*Factory.get_imple())



from com.sun.star.beans import NamedValue

RESOURCE_URL = "private:resource/dockingwindow/9809"
EXT_ID = "xaver.roemers.organon"


from com.sun.star.document import  XDocumentEventListener
from com.sun.star.awt import  XWindowListener
class Listener_To_Restart_Component(unohelper.Base,XWindowListener,XDocumentEventListener):
    
    def __init__(self):
        if debug: log(inspect.stack)
        # LO
        self.eventb = None
        # OO
        self.win = None

    
    def windowShown(self,ev): 
        # OO
        if debug: log(inspect.stack)
        
        self.win.removeWindowListener(self) 
        self.lade_Komponente_neu()
    
    def documentEventOccured(self,ev):
        # LO
        if ev.EventName == 'OnLayoutFinished':
            if debug: log(inspect.stack)
            self.eventb.removeDocumentEventListener(self)
            self.lade_Komponente_neu()
    
    def lade_Komponente_neu(self):
        if debug: log(inspect.stack)
        
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

        try:
            from threading import Thread

            def sleeper(desktop1):   
                URL="private:factory/swriter"
                desktop1.ActiveFrame.loadComponentFromURL(URL,'_top','',())  

            t = Thread(target=sleeper,args=(desktop,))
            t.start()
                
        except Exception as e:
            log(inspect.stack,tb())
            
    def windowHidden(self,ev):pass
    def windowNormalized(self,ev):pass
    def windowDeactivated(self,ev):pass
    def windowOpened(self,ev):pass
    def windowResized(self,ev):pass
    def windowMoved(self,ev):pass
    def disposing(self,ev):pass
    







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



####################################################
                # WRITER DESIGN #
####################################################  
  

def set_app_style():
    if debug: log(inspect.stack)
    
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        frame = desktop.Frames.getByIndex(0)
        comp = frame.ComponentWindow
        
        rot = 16275544

        hf = KONST.FARBE_HF_HINTERGRUND
        menu = KONST.FARBE_MENU_HINTERGRUND
        schrift = KONST.FARBE_SCHRIFT_DATEI
        menu_schrift = KONST.FARBE_MENU_SCHRIFT
        selected = KONST.FARBE_AUSGEWAEHLTE_ZEILE
        ordner = KONST.FARBE_SCHRIFT_ORDNER
        
        sett = settings_orga['organon_farben']['office']
        
        def get_farbe(value):
            if isinstance(value, int):
                return value
            else:
                return settings_orga['organon_farben'][value]
        
        # Kann button_schrift evt. herausgenommen werden?
        button_schrift = get_farbe(sett['button_schrift'])
        
        statusleiste_schrift = get_farbe(sett['statusleiste_schrift'])
        statusleiste_hintergrund = get_farbe(sett['statusleiste_hintergrund'])
        
        felder_hintergrund = get_farbe(sett['felder_hintergrund'])
        felder_schrift = get_farbe(sett['felder_schrift'])
        
        # Sidebar
        sidebar_eigene_fenster_hintergrund = get_farbe(sett['sidebar']['eigene_fenster_hintergrund'])
        sidebar_selected_hintergrund = get_farbe(sett['sidebar']['selected_hintergrund'])
        sidebar_selected_schrift = get_farbe(sett['sidebar']['selected_schrift'])
        sidebar_schrift = get_farbe(sett['sidebar']['schrift'])
        
        trenner_licht = get_farbe(sett['trenner_licht'])
        trenner_schatten = get_farbe(sett['trenner_schatten'])
        
        # Lineal
        OO_anfasser_trenner = get_farbe(sett['OO_anfasser_trenner'])
        OO_lineal_tab_zwischenraum = get_farbe(sett['OO_lineal_tab_zwischenraum'])
        OO_schrift_lineal_sb_liste = get_farbe(sett['OO_schrift_lineal_sb_liste'])
        
        LO_anfasser_text = get_farbe(sett['LO_anfasser_text'])
        LO_tabsumrandung = get_farbe(sett['LO_tabsumrandung'])
        LO_lineal_bg_innen = get_farbe(sett['LO_lineal_bg_innen'])
        LO_tab_fuellung = get_farbe(sett['LO_tab_fuellung'])
        LO_tab_trenner = get_farbe(sett['LO_tab_trenner'])
        
        
        LO = ('LibreOffice' in frame.Title)
        
        STYLES = {  
                  # Allgemein
                    'ButtonRolloverTextColor' : button_schrift, # button rollover
                    
                    'FieldColor' : felder_hintergrund, # Hintergrund Eingabefelder
                    'FieldTextColor' : felder_schrift,# Schrift Eingabefelder
                    
                    # Trenner
                    'LightColor' : trenner_licht, # Fenster Trenner
                    'ShadowColor' : trenner_schatten, # Fenster Trenner
                    
                    # OO Lineal + Trenner
                     
                    'DarkShadowColor' : (LO_anfasser_text if LO    # LO Anfasser + Lineal Text
                                        else OO_anfasser_trenner), # OO Anfasser +  Document Fenster Trenner 
                    'WindowTextColor' : (schrift if LO      # Felder (Navi) Schriftfarbe Sidebar 
                                         else OO_schrift_lineal_sb_liste),     # Felder (Navi) Schriftfarbe Sidebar + OO Lineal Schriftfarbe   
                        
                    # Sidebar
                    'LabelTextColor' : sidebar_schrift, # Schriftfarbe Sidebar + allg Dialog
                    'DialogColor' : sidebar_eigene_fenster_hintergrund, # Hintergrund Sidebar Dialog
                    'FaceColor' : (hf if LO        # LO Formatvorlagen Treeview Verbinder
                                    else hf),           # OO Hintergrund Organon + Lineal + Dropdowns  
                    'WindowColor' : (hf if LO                           # LO Dialog Hintergrund
                                    else OO_lineal_tab_zwischenraum),   # OO Lineal Tabzwischenraum
                    'HighlightColor' : sidebar_selected_hintergrund, # Sidebar selected Hintergrund
                    'HighlightTextColor' : sidebar_selected_schrift, # Sidebar selected Schrift
                    
                    
#                     'ActiveBorderColor' : rot,#k.A.
#                     'ActiveColor' : rot,#k.A.
#                     'ActiveTabColor' : rot,#k.A.
#                     'ActiveTextColor' : rot,#k.A.
#                     'ButtonTextColor' : rot,# button Textfarbe / LO Statuszeile Textfarbe
#                     'CheckedColor' : rot,#k.A.
#                     'DeactiveBorderColor' : rot,#k.A.
#                     'DeactiveColor' : rot,#k.A.
#                     'DeactiveTextColor' : rot,#k.A.
#                     'DialogTextColor' : rot,#k.A.
#                     'DisableColor' : rot,
#                     'FieldRolloverTextColor' : rot,#k.A.
#                     'GroupTextColor' : rot,#k.A.
#                     'HelpColor' : rot,#k.A.
#                     'HelpTextColor' : rot,#k.A.
#                     'InactiveTabColor' : rot,#k.A.
#                     'InfoTextColor' : rot,#k.A.
#                     'MenuBarColor' : rot,#k.A.
#                     'MenuBarTextColor' : rot,#k.A.
#                     'MenuBorderColor' : rot,#k.A.
#                     'MenuColor' : rot,#k.A.
#                     'MenuHighlightColor' : rot,#k.A.
#                     'MenuHighlightTextColor' : rot,#k.A.
#                     'MenuTextColor' : rot,#k.A.
#                     'MonoColor' : rot, #k.A.
#                     'RadioCheckTextColor' : rot,#k.A.
#                     'WorkspaceColor' : rot, #k.A.
#                     erzeugen Fehler:
#                     'FaceGradientColor' : 502,
#                     'SeparatorColor' : 502,                    
                    }
        
 
        def stilaenderung(win,ignore=[]):

            for s in STYLES:
                if s in ignore: 
                    pass
                else:
                    try:
                        val = STYLES[s]
                        setattr(win.StyleSettings, s, val)
                    except Exception as e:
                        pass
                    
                win.setBackground(statusleiste_hintergrund) # Hintergrund Statuszeile
                win.setForeground(statusleiste_schrift)     # Schrift Statuszeile
        
        


        top_wins = []
        
        for i in range (toolkit.TopWindowCount):
            top_wins.append(toolkit.getTopWindow(i))
        
        
        # folgende Properties wuerden die Eigenschaften
        # der Office Menubar und aller Buttons setzen
        ignore = ['ButtonTextColor',
                 'LightColor',
                 'MenuBarTextColor',
                 'MenuBorderColor',
                 'ShadowColor'
                 ]

        for t in top_wins:
            stilaenderung(t,ignore)
            # Hier wird in OO auch das Lineal gesetzt
            # Es kann nicht einzeln angesprochen werden
        
        layoutmgr = frame.LayoutManager
        statusbar = layoutmgr.getElement("private:resource/statusbar/statusbar")
        try:
            stilaenderung(statusbar.RealInterface)
        except:
            pass
        
        
        STYLES_LINEAL = {
                        'ShadowColor' : LO_tabsumrandung, # Tabsumrandung
                        'WindowColor' : LO_lineal_bg_innen, #Lineal: Hintergrund innen
                        'WorkspaceColor' : LO_tab_fuellung, # Tabhalter Fuellung
                        'DialogColor' : LO_tab_trenner, #Lineal Seitenraender
                        }
        
         
        for t in comp.Windows:
            try:
                # Bordercolor
                stilaenderung(t)

                try:
                    for j in t.Windows:
                        # Lineal nur in LO
                        for s in STYLES_LINEAL:
                            val = STYLES_LINEAL[s]
                            try:
                                setattr(j.StyleSettings, s, val)
                            except:
                                pass
                        #j.setBackground(rot) # Lineal Leiste Hintergrund in LO                       
                except Exception as e:
                    pass
  
            except Exception as e:
                log(inspect.stack,tb())
        
    except Exception as e:
        log(inspect.stack,tb())





class geladen():
    geladen = 'neu'
g = geladen()
 
