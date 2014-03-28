# -*- coding: utf-8 -*-

#print('Projekte')
import traceback
import uno
import unohelper

tb = traceback.print_exc

BREITE_TAG1_CONTAINER = 100
HOEHE_TAG1_CONTAINER = 350
URL_IMGS = 'vnd.sun.star.extension://xaver.roemers.organon/img/'

IMG_ORDNER_GEOEFFNET_16 = 'vnd.sun.star.extension://xaver.roemers.organon/img/OrdnerGeoeffnet_16.png'



class Funktionen():
    
    def __init__(self,mb,pdk):
        self.mb = mb
        
        global pd
        pd = pdk
        
        
        
    def projektordner_ausklappen(self):
        if self.mb.debug: print(self.mb.debug_time(),'projektordner_ausklappen')
        
        tree = self.mb.xml_tree
        root = tree.getroot()
         
        xml_projekt = root.find(".//*[@Name='Projekt']")
        alle_elem = xml_projekt.findall('.//')
        
        projekt_zeile = self.mb.Hauptfeld.getControl(xml_projekt.tag)
        icon = projekt_zeile.getControl('icon')
        icon.Model.ImageURL = IMG_ORDNER_GEOEFFNET_16
        
        for zeile in alle_elem:
            zeile.attrib['Sicht'] = 'ja'
            if zeile.attrib['Art'] == 'dir':
                zeile.attrib['Zustand'] = 'auf'
                hf_zeile = self.mb.Hauptfeld.getControl(zeile.tag)
                icon = hf_zeile.getControl('icon')
                icon.Model.ImageURL = IMG_ORDNER_GEOEFFNET_16
                
        tag = xml_projekt.tag
        self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_hf(xml_projekt.tag,xml_projekt,'zu',True)
        self.mb.class_Projekt.erzeuge_dict_ordner() 
         
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        self.mb.xml_tree.write(Path)
    
    
    def erzeuge_Tag1_Container(self,ev):
    
        smgr = self.mb.smgr
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.mb.ctx)    
        oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", self.mb.ctx)

        # Create Uno Struct
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
        oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
        # global oWindow
        oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
        oWindowDesc.WindowServiceName = ""
        oWindowDesc.Parent = toolkit.getDesktopWindow()
        oWindowDesc.ParentIndex = -1
        oWindowDesc.WindowAttributes = 32

        # Set Window Attributes
        gnDefaultWindowAttributes = 1 + 16  # + 32 +64+128 
#         com_sun_star_awt_WindowAttribute_SHOW + \
#         com_sun_star_awt_WindowAttribute_BORDER + \
#         com_sun_star_awt_WindowAttribute_MOVEABLE + \
#         com_sun_star_awt_WindowAttribute_CLOSEABLE + \
#         com_sun_star_awt_WindowAttribute_SIZEABLE
          
        X = ev.value.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.value.Source.AccessibleContext.LocationOnScreen.value.Y + ev.value.Source.Size.value.Height

        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = BREITE_TAG1_CONTAINER
        oRect.Height = HOEHE_TAG1_CONTAINER
        
        oWindowDesc.Bounds = oRect

        # specify the window attributes.
        oWindowDesc.WindowAttributes = gnDefaultWindowAttributes
        # create window
        oWindow = toolkit.createWindow(oWindowDesc)
         
        # create frame for window
        oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame", self.mb.ctx)
        oFrame.initialize(oWindow)
        oFrame.setCreator(self.mb.desktop)
        oFrame.activate()

        # create new control container
        cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", self.mb.ctx)
        cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", self.mb.ctx)
        cont_model.BackgroundColor = 305099  # 9225984
        cont.setModel(cont_model)
        # need createPeer just only the container
        cont.createPeer(toolkit, oWindow)
        cont.setPosSize(0, 0, 0, 0, 15)
        
        oFrame.setComponent(cont, None)
        
        # create Listener
        listener = Tag1_Container_Listener()
        cont.addMouseListener(listener) 
        listener.ob = oWindow  
        
#    if ev.value.Source.AccessibleContext.AccessibleName == 'Datei':
        self.erzeuge_Menu_DropDown_Eintraege_Datei(oWindow, cont,ev.Source)
   
    
    def erzeuge_Menu_DropDown_Eintraege_Datei(self,window,cont,source):
        control, model = createControl(self.mb.ctx, "ListBox", 4 ,  4 , 
                                       BREITE_TAG1_CONTAINER - 10, HOEHE_TAG1_CONTAINER - 10, (), ())   
        control.setMultipleMode(False)

        items = ('leer',
                'blau',
                'braun',
                'creme',
                'gelb',
                'grau',
                'gruen',
                'hellblau',
                'hellgrau',
                'lila',
                'ocker',
                'orange',
                'pink',
                'rostrot',
                'rot',
                'schwarz',
                'tuerkis',
                'weiss')
                
        control.addItems(items, 0)           
        
        for item in items:
            pos = items.index(item)
            model.setItemImage(pos,URL_IMGS+'punkt_%s.png' %item)
        
        
        tag_item_listener = Tag1_Item_Listener(self.mb,window,source)
        control.addItemListener(tag_item_listener)
        
        cont.addControl('Eintraege_Tag1', control)


                

       
from com.sun.star.awt import XMouseListener,XItemListener
class Tag1_Container_Listener (unohelper.Base, XMouseListener):
        def __init__(self):
            pass
           
        def mousePressed(self, ev):
            #print('mousePressed,Tag1_Container_Listener')  
            if ev.Buttons == MB_LEFT:
                return False
       
        def mouseExited(self, ev): 
            #print('mouseExited')                      
            if self.enthaelt_Punkt(ev):
                pass
            else:            
                self.ob.dispose()    
            return False
        
        def enthaelt_Punkt(self, ev):
            #print('enthaelt_Punkt') 
            X = ev.value.X
            Y = ev.value.Y
            
            XTrue = (0 <= X < ev.value.Source.Size.value.Width)
            YTrue = (0 <= Y < ev.value.Source.Size.value.Height)
            
            if XTrue and YTrue:           
                return True
            else:
                return False   
             
class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,window,source):
        self.mb = mb
        self.window = window
        self.source = source
        
    # XItemListener    
    def itemStateChanged(self, ev):        
        sel = ev.value.Source.Items[ev.value.Selected]
        #print(sel)
        # image tag1 aendern
        self.source.Model.ImageURL = URL_IMGS+'punkt_%s.png' %sel
         # tag1 in xml datei einf�gen und speichern
        ord_source = self.source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
        tree = self.mb.xml_tree
        root = tree.getroot()        
        source_xml = root.find('.//'+ord_source)
        source_xml.attrib['Tag1'] = sel
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        tree.write(Path)
        
        self.window.dispose()
       
         

             
################ TOOLS ################################################################

# Handy function provided by hanya (from the OOo forums) to create a control, model.
def createControl(ctx,type,x,y,width,height,names,values):
   smgr = ctx.getServiceManager()
   ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
   ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
   ctrl_model.setPropertyValues(names,values)
   ctrl.setModel(ctrl_model)
   ctrl.setPosSize(x,y,width,height,15)
   return (ctrl, ctrl_model)
def createUnoService(serviceName):
  sm = uno.getComponentContext().ServiceManager
  return sm.createInstanceWithContext(serviceName, uno.getComponentContext())

