# -*- coding: utf-8 -*-

import unohelper


class Funktionen():
    
    def __init__(self,mb,pdk):
        self.mb = mb        
        
        global pd
        pd = pdk
        
    def projektordner_ausklappen(self):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()

        xml_projekt = root.find(".//*[@Name='%s']" % self.mb.projekt_name)
        alle_elem = xml_projekt.findall('.//')

        projekt_zeile = self.mb.props[T.AB].Hauptfeld.getControl(xml_projekt.tag)
        icon = projekt_zeile.getControl('icon')
        icon.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
        
        for zeile in alle_elem:
            zeile.attrib['Sicht'] = 'ja'
            if zeile.attrib['Art'] in ('dir','prj'):
                zeile.attrib['Zustand'] = 'auf'
                hf_zeile = self.mb.props[T.AB].Hauptfeld.getControl(zeile.tag)
                icon = hf_zeile.getControl('icon')
                icon.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
                
        tag = xml_projekt.tag
        self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_hf(xml_projekt.tag,xml_projekt,'zu',True)
        self.mb.class_Projekt.erzeuge_dict_ordner() 
        self.mb.class_Hauptfeld.korrigiere_scrollbar()    
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        self.mb.props[T.AB].xml_tree.write(Path)

       
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
        
        self.erzeuge_ListBox_Tag1(fenster, fenster_cont,ev.Source)

    
    def erzeuge_ListBox_Tag1(self,window,cont,source):
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
    
    
    def find_parent_section(self,sec):
        
        def find_parsection(section):
            
            if section == None:
                # Diese Bedingung wird nur bei einem Fehler durchlaufen, dann naemlich
                # wenn der Bereich 'OrgInnerSec' faelschlich umbenannt wurde.
                # Diese Bedingung soll sicherstellen, dass die Funktion auf jeden Fall funktioniert
                return self.parsection
            
            elif 'OrgInnerSec' not in section.Name:
                find_parsection(section.ParentSection)
            else:
                self.parsection = section
                
        find_parsection(sec)
        
        return self.parsection
        
    def teile_text(self):
        try:

            zeilenordinal =  self.mb.props[T.AB].selektierte_zeile.AccessibleName     
            vc = self.mb.viewcursor
            cur = self.mb.doc.Text.createTextCursor()
            sec = vc.TextSection
            
            # parent section finden   
            parsection = self.find_parent_section(sec)
            
            # Laenge des Textanfangs zaehlen
            cur.gotoRange(vc,False)
            cur.gotoRange(self.parsection.Anchor.Start,True)
            
            string = cur.String
            s = string.replace('\n','') # Zeilenumbruech ausschliessen
            textanfang_laenge = len(s)          

            # sicherheitshalber alte Datei speichern
            url_source = os.path.join(self.mb.pfade['odts'],zeilenordinal + '.odt')
            URL_source = uno.systemPathToFileUrl(url_source)
            
            orga_sec_name_alt = self.mb.props['Projekt'].dict_bereiche['ordinal'][zeilenordinal]
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(URL_source,orga_sec_name_alt)

            # Ende in der getrennten Datei loeschen
            cur.gotoRange(vc,False)
            cur.gotoRange(self.parsection.Anchor.End,True)
            cur.setString('')
            
            # erzeuge neue Zeile
            nr_neue_zeile = self.mb.class_Hauptfeld.erzeuge_neue_Zeile('dokument')
            ordinal_neue_zeile = 'nr'+ str(nr_neue_zeile)
       
            # neuen File erzeugen        
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True

            url_target = os.path.join(self.mb.pfade['odts'],ordinal_neue_zeile + '.odt')
            URL_target = uno.systemPathToFileUrl(url_target)
            
            new_doc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL_source,'_blank',0,(prop,))
            
            # Sichtbarkeit schalten, damit OrgInnerSec erzeugt wird
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche(ordinal_neue_zeile)
                        
            # OrgInnerSec der neuen Datei setzen
            sec = new_doc.TextSections.getByName(parsection.Name)
            orga_sec_name = self.mb.props['Projekt'].dict_bereiche['ordinal'][ordinal_neue_zeile]
            orga_sec = self.mb.doc.TextSections.getByName(orga_sec_name)
            orga_inner_sec_name = orga_sec.ChildSections[0].Name
            
            sec.setName(orga_inner_sec_name)
            # Speichern
            new_doc.storeToURL(URL_target,())
            new_doc.close(False)

            # File Link setzen, um Anzeige zu erneuern
            sec = self.mb.doc.TextSections.getByName(orga_sec_name)
            
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FileURL = ''
            
            SFLink2 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink2.FileURL = sec.FileLink.FileURL
            SFLink2.FilterName = 'writer8'
    
            sec.setPropertyValue('FileLink',SFLink)
            sec.setPropertyValue('FileLink',SFLink2)
            
            # Anfang in der neuen datei loeschen
            parsection = self.find_parent_section(sec)
            vc.gotoRange(parsection.Anchor.Start,True)
            vc.gotoStart(False)
            vc.goRight(textanfang_laenge,True)
            vc.setString('')
                  
            # Einstellungen, tags der alten Datei fuer neue uebernehmen
            self.mb.dict_sb_content['ordinal'][ordinal_neue_zeile] = self.mb.dict_sb_content['ordinal'][zeilenordinal]
            
            tree = self.mb.props['Projekt'].xml_tree
            root = tree.getroot()
            alt = root.find('.//'+zeilenordinal)
            neu = root.find('.//'+ordinal_neue_zeile)
            
            neu.attrib['Tag1'] = alt.attrib['Tag1']
            neu.attrib['Tag2'] = alt.attrib['Tag2']
            neu.attrib['Tag3'] = alt.attrib['Tag3']
            
            
            # nach Aenderung Speichern
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(URL_target,orga_sec_name)
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(URL_source,orga_sec_name_alt)
            for tag in self.mb.dict_sb_content['sichtbare']:
                self.mb.class_Sidebar.erzeuge_sb_layout(tag,'teile_text')

        except:
            tb()
        
        #pd()


     
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
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()        
        source_xml = root.find('.//'+ord_source)
        source_xml.attrib['Tag1'] = sel
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        tree.write(Path)
        
        self.window.dispose()

         

             