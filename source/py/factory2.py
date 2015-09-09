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
        sys.path.append('/home/xgr/.eclipse/org.eclipse.platform_4.4.1_1473617060_linux_gtk_x86_64/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
    else:
        sys.path.append(r'H:/Programme/eclipse/plugins/org.python.pydev_4.0.0.201504132356/pysrc')  
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

from einstellungen import Take_Over_Old_Settings as TO

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
        
        # KONST.FARBE_ORGANON_FENSTER
    except Exception as e:
        print(tb())

set_konst()

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
        '''
        Initialising the factory which is responsible for creating
        Organon's dialog window.
        
        This is also the first available place to set the design.
        Before no services can be created, later on the dialog window's background
        isn't editable anymore.
        '''
        self.ctx = ctx
        print("factory init")
        
        if settings_orga['organon_farben']['design_office']:
            set_app_style()
            
    
    def createInstanceWithArgumentsAndContext(self, args, ctx):
        
        try:
            CWHandler = ContainerWindowHandler(ctx)
            self.CWHandler = CWHandler
            
            win,tabs = create_window(ctx,self)
            window = self.CWHandler.window2

            window.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            start_main(window,ctx,tabs,path_to_extension,win,self)  
            
            return win
        
        except Exception as e:
            print('Factory '+e)
            tb()


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



from com.sun.star.awt import  XTopWindowListener
class Top_Window_Listener(unohelper.Base,XTopWindowListener):
    
    def __init__(self,frame):
        self.frame = frame

    def windowOpened(self,ev):
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        try:

            # um den Listener nur einmal anzusprechen
            if g.geladen == 'neu':
                g.geladen = 'alt'    
                    
                URL="private:factory/swriter"
                self.frame.loadComponentFromURL(URL,'_top',0,())   
        
        except Exception as e:
            log(inspect.stack,tb())

    def windowActivated(self,ev):   pass
    def windowClosing(self,ev):     pass
    def windowClosed(self,ev):      pass
    def windowMinimized(self,ev):   pass
    def windowNormalized(self,ev):  pass
    def windowDeactivated(self,ev): pass
        

from com.sun.star.document import  XDocumentEventListener
class Doc_Event_Listener(unohelper.Base,XDocumentEventListener):
    
    def __init__(self,frame):
        self.frame = frame
        
    def documentEventOccured(self,ev):

        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        
        if ev.EventName == 'OnLayoutFinished':
            self.eventb.removeDocumentEventListener(self.listener)

            try:
                # um den Listener nur einmal anzusprechen
                if g.geladen == 'neu':
                    g.geladen = 'alt'    
                if desktop.Name == '':
                    desktop.Name = 'gestartet'
                    
                    from threading import Thread

                    def sleeper(desktop1):   
                        import time
                        #time.sleep(2)  
                        URL="private:factory/swriter"
                        desktop1.ActiveFrame.loadComponentFromURL(URL,'_top','',())  

                    t = Thread(target=sleeper,args=(desktop,))
                    t.start()
                    
            except Exception as e:
                log(inspect.stack,tb())



def start_main(window,ctx,tabs,path_to_extension,win,factory):

    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        frame = desktop.Frames.getByIndex(0)
        
        if desktop.ActiveFrame != None:
            frame = desktop.ActiveFrame
        
        if settings_orga['organon_farben']['design_office']:
            if 'OpenOffice' in frame.Title:
                # in OO muss die Componente neu gestartet werden, damit das
                # Lineal eingefaerbt wird.    
                listener = Top_Window_Listener(frame)
                frame.ContainerWindow.addTopWindowListener(listener)    

        
        dialog = window
        debug = settings_orga['log_config']['output_console']

        import menu_start
        
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



####################################################
                # WRITER DESIGN #
####################################################  
  

def set_app_style():
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
                    'FaceColor' : (schrift if LO        # LO Formatvorlagen Treeview Verbinder
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
        
        try:
            # set listener to restart LO
            if 'LibreOffice' in frame.Title:
                listener = Doc_Event_Listener(frame)  
                doc = desktop.getCurrentComponent() 
                eventb = ctx.getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
                listener.eventb= eventb
                listener.listener = listener
                eventb.addDocumentEventListener(listener)
             
        except Exception as e:
            log(inspect.stack,tb()) 


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
        stilaenderung(statusbar.RealInterface)
        
        
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
 
