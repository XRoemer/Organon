# -*- coding: utf-8 -*-

import unohelper


class Funktionen():
    
    def __init__(self,mb,pdk):
        self.mb = mb        
        
        global pd
        pd = pdk
        
    def projektordner_ausklappen(self):
        if self.mb.debug: print(self.mb.debug_time(),'projektordner_ausklappen')
        
        tree = self.mb.xml_tree
        root = tree.getroot()

        xml_projekt = root.find(".//*[@Name='%s']" % self.mb.projekt_name)
        alle_elem = xml_projekt.findall('.//')

        projekt_zeile = self.mb.Hauptfeld.getControl(xml_projekt.tag)
        icon = projekt_zeile.getControl('icon')
        icon.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
        
        for zeile in alle_elem:
            zeile.attrib['Sicht'] = 'ja'
            if zeile.attrib['Art'] in ('dir','prj'):
                zeile.attrib['Zustand'] = 'auf'
                hf_zeile = self.mb.Hauptfeld.getControl(zeile.tag)
                icon = hf_zeile.getControl('icon')
                icon.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
                
        tag = xml_projekt.tag
        self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_hf(xml_projekt.tag,xml_projekt,'zu',True)
        self.mb.class_Projekt.erzeuge_dict_ordner() 
        self.mb.class_Hauptfeld.korrigiere_scrollbar()    
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        self.mb.xml_tree.write(Path)

       
    def erzeuge_Tag1_Container(self,ev):
 
        Width = KONST.BREITE_TAG1_CONTAINER
        Height = KONST.HOEHE_TAG1_CONTAINER
        X = ev.value.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.value.Source.AccessibleContext.LocationOnScreen.value.Y + ev.value.Source.Size.value.Height
        posSize = X,Y,Width,Height 
        flags = 1+16+32+128
        #flags=1+32+64+128
        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize,flags)
        
        # create Listener
        listener = Tag1_Container_Listener()
        fenster_cont.addMouseListener(listener) 
        listener.ob = fenster  
        
        self.erzeuge_Menu_DropDown_Eintraege_Datei(fenster, fenster_cont,ev.Source)

    
    def erzeuge_Menu_DropDown_Eintraege_Datei(self,window,cont,source):
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 4 ,  4 , 
                                       KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
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
            model.setItemImage(pos,KONST.URL_IMGS+'punkt_%s.png' %item)
        
        
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

        # image tag1 aendern
        self.source.Model.ImageURL = KONST.URL_IMGS+'punkt_%s.png' %sel
        # tag1 in xml datei einfuegen und speichern
        ord_source = self.source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
        tree = self.mb.xml_tree
        root = tree.getroot()        
        source_xml = root.find('.//'+ord_source)
        source_xml.attrib['Tag1'] = sel
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        tree.write(Path)
        
        self.window.dispose()

         

             