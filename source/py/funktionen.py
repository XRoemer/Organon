# -*- coding: utf-8 -*-

import unohelper
from math import sqrt
from shutil import copyfile


class Funktionen():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb        

        
    def projektordner_ausklappen(self,ordinal = None):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        
        if ordinal == None:
            xml_projekt = root.find(".//*[@Name='%s']" % self.mb.projekt_name)
            alle_elem = xml_projekt.findall('.//')
        else:
            xml_projekt = root.find(".//%s" % ordinal)
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
        self.mb.class_Baumansicht.korrigiere_scrollbar()    
        
        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
        self.mb.tree_write(self.mb.props[T.AB].xml_tree,Path)

       
    def erzeuge_Tag1_Container(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        Width = KONST.BREITE_TAG1_CONTAINER
        Height = KONST.HOEHE_TAG1_CONTAINER
        X = ev.value.Source.AccessibleContext.LocationOnScreen.value.X - self.mb.topWindow.AccessibleContext.LocationOnScreen.value.X +20
        Y = ev.value.Source.AccessibleContext.LocationOnScreen.value.Y - self.mb.topWindow.AccessibleContext.LocationOnScreen.value.Y
        posSize = X,Y,Width,Height 
        flags = 1+16+32+128
        #flags=1+32+64+128
        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize,flags)

        # create Listener
        listener = Tag_Container_Listener()
        fenster_cont.addMouseListener(listener) 
        listener.ob = fenster  
        
        self.erzeuge_ListBox_Tag1(fenster, fenster_cont,ev.Source)
        
    def erzeuge_ListBox_Tag1(self,window,cont,source):
        if self.mb.debug: log(inspect.stack)
        
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 4 ,  4 , 
                                       KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control.setMultipleMode(False)
        model.Border = 0
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        
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
        
        
            
    def erzeuge_Tag2_Container(self,ev,ordinal):
        if self.mb.debug: log(inspect.stack)
        try:
            
            controls = []
            icons_gallery,icons_prj_folder = self.get_icons()

            # create Listener
            listener = Tag_Container_Listener()
            listener2 = Tag2_Images_Listener(self.mb)
            listener2.ordinal = ordinal
            listener2.icons_dict = {}
            
            y = 10
            x = 0            
            
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10,y,200, 20, (), ())  
            model.Label = 'Kein Icon:'
            model.FontWeight = 150
            width,h = self.mb.kalkuliere_und_setze_Control(control,'w')
            controls.append(control)

            control, model = self.mb.createControl(self.mb.ctx, "ImageControl", 20 + width,y,18, 18, (), ())  
            model.ImageURL = ''
            model.setPropertyValue('HelpText','Kein Icon')
            model.setPropertyValue('Border',1)
            control.addMouseListener(listener2) 
                            
            controls.append(control)
            
            y += 25
            
            
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10,y,200, 20, (), ())  
            model.Label = 'Im Projekt verwendete Icons:'
            model.FontWeight = 150
            prefW,h = self.mb.kalkuliere_und_setze_Control(control,'w')
            controls.append(control)
            
            y += 25
            
            anzahl = sqrt(len(icons_prj_folder))

            for iGal in icons_prj_folder:
                control, model = self.mb.createControl(self.mb.ctx, "ImageControl", x +10,y,16, 16, (), ())  

                model.ImageURL = iGal[1]
                model.setPropertyValue('HelpText',iGal[0])
                model.setPropertyValue('Border',0)
                model.setPropertyValue('ScaleImage' ,True)
                control.addMouseListener(listener2) 
                listener2.icons_dict.update({iGal[0]:iGal[1]})
                
                controls.append(control)
                
                x += 25
                
                if x> anzahl*25:
                    y+=25
                    x = 0
            
            y += 25
            
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10,y,200, 20, (), ())  
            model.Label = 'In der Galerie vorhandene Icons:'
            model.FontWeight = 150
            prefW1,h = self.mb.kalkuliere_und_setze_Control(control,'w')
            controls.append(control)
            
            y += 25
            x = 0
            anzahl2 = sqrt(len(icons_gallery))
            
            
            for iGal in icons_gallery:
                control, model = self.mb.createControl(self.mb.ctx, "ImageControl", x +10,y,16, 16, (), ())  

                model.ImageURL = iGal[1]
                model.setPropertyValue('HelpText',iGal[0])
                model.setPropertyValue('Border',0)
                model.setPropertyValue('ScaleImage' ,True)
                control.addMouseListener(listener2) 
                                
                controls.append(control)
                
                x += 25
                
                if x> anzahl2*25:
                    y+=25
                    x = 0

            breite = sorted((prefW,prefW1,int(anzahl)*25,int(anzahl2)*25))[-1]


            X = ev.value.Source.AccessibleContext.LocationOnScreen.value.X - self.mb.topWindow.AccessibleContext.LocationOnScreen.value.X +40
            Y = ev.value.Source.AccessibleContext.LocationOnScreen.value.Y - self.mb.topWindow.AccessibleContext.LocationOnScreen.value.Y -60

            posSize = X,Y,breite + 20,y +25 
            flags = 1+16+32+128

            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize,flags)
            fenster_cont.addMouseListener(listener) 
            listener.ob = fenster  
            
            for control in controls:
                fenster_cont.addControl('wer',control)
            
        except:
            log(inspect.stack,tb())
            
            
            
    def get_icons(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            icons_gallery = []
            icons_prj_folder = []
            icons_prj_folder_names = []
            
                
            icons_folder = self.mb.pfade['icons']
                        
            for root, dirs, files in os.walk(icons_folder):
                for f in files:
                    name = f
                    path = os.path.join(root,f)
                    path2 = uno.systemPathToFileUrl(path)
                    icons_prj_folder.append((name,path2))
                    icons_prj_folder_names.append(name)
            
            gallery = self.mb.createUnoService("com.sun.star.gallery.GalleryThemeProvider")
            org = gallery.getByName('Organon Icons')

            for i in range(org.Count):
                url = org.getByIndex(i).URL
                
                if os.path.basename(url) not in icons_prj_folder_names:
                    url_os = uno.fileUrlToSystemPath(url)
                    name = os.path.basename(url_os).split('.')[0]
                    icons_gallery.append((name,url))
            
            
            return icons_gallery, icons_prj_folder 

        except:
            log(inspect.stack,tb())
        
    def find_parent_section(self,sec):
        if self.mb.debug: log(inspect.stack)
        
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
        if self.mb.debug: log(inspect.stack)
        
        try:

            zeilenordinal =  self.mb.props[T.AB].selektierte_zeile    
            kommender_eintrag = self.mb.props['Projekt'].kommender_Eintrag
            
            url_source = os.path.join(self.mb.pfade['odts'],zeilenordinal + '.odt')
            URL_source = uno.systemPathToFileUrl(url_source)
            helfer_url = URL_source+'helfer'
             
            vc = self.mb.viewcursor
            cur_old = self.mb.doc.Text.createTextCursor()
            sec = vc.TextSection
            text = self.mb.doc.Text
            
            # parent section finden   
            parsection = self.find_parent_section(sec)
            
            # Bookmark setzen
            bm = self.mb.doc.createInstance('com.sun.star.text.Bookmark')
            bm.Name = 'kompliziertkompliziert'            
            text.insertTextContent(vc,bm,False)
            
            # alte Datei in Helferdatei speichern
            orga_sec_name_alt = self.mb.props['Projekt'].dict_bereiche['ordinal'][zeilenordinal]
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(helfer_url,orga_sec_name_alt)
             
            # erzeuge neue Zeile
            nr_neue_zeile = self.mb.class_Baumansicht.erzeuge_neue_Zeile('dokument')
            ordinal_neue_zeile = 'nr'+ str(nr_neue_zeile)
       
            # aktuelle datei unsichtbar oeffnen        
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True

            url_target = os.path.join(self.mb.pfade['odts'],ordinal_neue_zeile + '.odt')
            URL_target = uno.systemPathToFileUrl(url_target)
               
            doc_new = self.mb.doc.CurrentController.Frame.loadComponentFromURL(helfer_url,'_blank',0,(prop,))  
            cur_new = doc_new.Text.createTextCursor()
            
            # OrgInnerSec umbenennen
            new_OrgInnerSec_name = 'OrgInnerSec' + str(kommender_eintrag)
            sec = doc_new.TextSections.getByName(parsection.Name)
            sec.setName(new_OrgInnerSec_name)
            
            # Textanfang und Bookmark in Datei loeschen 
            bms = doc_new.Bookmarks
            bm2 = bms.getByName('kompliziertkompliziert')
            new_OrgInnerSec = doc_new.TextSections.getByName(new_OrgInnerSec_name)
            
            cur_new.gotoRange(bm2.Anchor,False)
            cur_new.gotoRange(new_OrgInnerSec.Anchor.Start,True)
            cur_new.setString('')
            bm2.dispose()
            
            # alte datei ueber neue speichern
            doc_new.storeToURL(URL_target,())
            doc_new.close(False)
            
            # Helfer loeschen
            os.remove(uno.fileUrlToSystemPath(helfer_url))
            
            # Ende in der getrennten Datei loeschen
            cur_old.gotoRange(bm.Anchor,False)
            cur_old.gotoRange(self.parsection.Anchor.End,True)
            cur_old.setString('')
            
            # Bookmark wird von cursor geloescht
            
            # Sichtbarkeit schalten
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche(ordinal_neue_zeile)
            
            # alte Datei speichern
            orga_sec_name_alt = self.mb.props['Projekt'].dict_bereiche['ordinal'][zeilenordinal]
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(URL_source,orga_sec_name_alt)
            
            # File Link setzen, um Anzeige zu erneuern
            sec = self.mb.doc.TextSections.getByName(new_OrgInnerSec_name)
             
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FileURL = ''
             
            SFLink2 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink2.FileURL = sec.FileLink.FileURL
            SFLink2.FilterName = 'writer8'
     
            sec.setPropertyValue('FileLink',SFLink)
            sec.setPropertyValue('FileLink',SFLink2)
            
            vc.gotoStart(False)
            
            # Einstellungen, tags der alten Datei fuer neue uebernehmen
            self.mb.dict_sb_content['ordinal'][ordinal_neue_zeile] = copy.deepcopy(self.mb.dict_sb_content['ordinal'][zeilenordinal])
            
            tree = self.mb.props['Projekt'].xml_tree
            root = tree.getroot()
            alt = root.find('.//'+zeilenordinal)
            neu = root.find('.//'+ordinal_neue_zeile)
            
            neu.attrib['Tag1'] = alt.attrib['Tag1']
            neu.attrib['Tag2'] = alt.attrib['Tag2']
            neu.attrib['Tag3'] = alt.attrib['Tag3']
            
            for tag in self.mb.dict_sb_content['sichtbare']:
                self.mb.class_Sidebar.erzeuge_sb_layout(tag,'teile_text')
                
        except Exception as e:
            self.mb.nachricht('teile_text ' + str(e),"warningbox")
            log(inspect.stack,tb())

     
from com.sun.star.awt import XMouseListener,XItemListener
class Tag_Container_Listener (unohelper.Base, XMouseListener):
    def __init__(self):
        self.ob = None
       
    def mousePressed(self, ev):
        return False
   
    def mouseExited(self, ev): 
        
        point = uno.createUnoStruct('com.sun.star.awt.Point')
        point.X = ev.X
        point.Y = ev.Y

        enthaelt_Punkt = ev.Source.AccessibleContext.containsPoint(point)
        
        if enthaelt_Punkt:
            pass
        else:            
            self.ob.dispose()    
        return False
    
    def mouseEntered(self, ev):  
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
  
            
class Tag2_Images_Listener (unohelper.Base, XMouseListener):
    def __init__(self,mb):
        self.mb = mb
        self.ordinal = None
        self.icons_dict = None
       
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack) 

        url = ev.Source.Model.ImageURL

        if url != '':
            self.galerie_icon_im_prj_ordner_evt_loeschen()
            url = self.galerie_icon_im_prj_ordner_speichern(url) 
        else:
            self.galerie_icon_im_prj_ordner_evt_loeschen() 

        self.tag2_in_allen_tabs_xml_anpassen(self.ordinal,url)

    def tag2_in_allen_tabs_xml_anpassen(self,ord_source,url):
        if self.mb.debug: log(inspect.stack) 
        try:
            tabnamen = self.mb.props.keys()

            for name in tabnamen:
            
                tree = self.mb.props[name].xml_tree
                root = tree.getroot()        
                source_xml = root.find('.//'+ord_source)
                
                if source_xml != None:
                
                    source_xml.attrib['Tag2'] = url
                    
                    if name == 'Projekt':
                        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
                    else:
                        Path = os.path.join(self.mb.pfade['tabs'], name + '.xml')
                    
                    tag2_button = self.mb.props[name].Hauptfeld.getControl(ord_source).getControl('tag2')
                    tag2_button.Model.ImageURL = url
                        
                    self.mb.tree_write(tree,Path)
        except:
            log(inspect.stack,tb())

    
    def galerie_icon_im_prj_ordner_speichern(self,url):  
        if self.mb.debug: log(inspect.stack)
        
        try:            
            url = uno.fileUrlToSystemPath(url)
            pfad_icons_prj_ordner = self.mb.pfade['icons']
            name = os.path.basename(url)
            neuer_pfad = os.path.join(pfad_icons_prj_ordner,name)

            if not os.path.exists(neuer_pfad):
                copyfile(url, neuer_pfad)
                
            return uno.systemPathToFileUrl(neuer_pfad)

        except:
            log(inspect.stack,tb())
    
    def galerie_icon_im_prj_ordner_evt_loeschen(self): 
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props['Projekt'].xml_tree
        root = tree.getroot()
        
        ord_xml = root.find('.//' + self.ordinal)
        url = ord_xml.attrib['Tag2']
        
        if 'uno_packages' in url:
            return
        all_xml = root.findall('.//')
        
        for el in all_xml:
            if el.tag == self.ordinal:
                continue
            if el.attrib['Tag2'] == url:
                return

        # Wenn die url nicht mehr im Dokument vorhanden ist, wird sie geloescht
        os.remove(uno.fileUrlToSystemPath(url))

   
    def mouseExited(self, ev): 
        ev.value.Source.Model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        return False
    def mouseEntered(self, ev):    
        ev.value.Source.Model.BackgroundColor = 102
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
        
             
class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,window,source):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.window = window
        self.source = source
        
    def itemStateChanged(self, ev):   
        if self.mb.debug: log(inspect.stack) 

        sel = ev.value.Source.Items[ev.value.Selected]

        # image tag1 aendern
        self.source.Model.ImageURL = KONST.URL_IMGS+'punkt_%s.png' %sel
        
        # tag1 in xml datei einfuegen und speichern
        ord_source = self.source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
        
        self.tag1_in_allen_tabs_xml_anpassen(ord_source,sel)
        
        self.window.dispose()

    def tag1_in_allen_tabs_xml_anpassen(self,ord_source,sel):
        if self.mb.debug: log(inspect.stack) 
        
        try:
            tabnamen = self.mb.props.keys()
            
            for name in tabnamen:
            
                tree = self.mb.props[name].xml_tree
                root = tree.getroot()        
                source_xml = root.find('.//'+ord_source)
                
                if source_xml != None:
                
                    source_xml.attrib['Tag1'] = sel
                    
                    if name == 'Projekt':
                        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
                    else:
                        Path = os.path.join(self.mb.pfade['tabs'], name + '.xml')
                    
                    tag1_button = self.mb.props[name].Hauptfeld.getControl(ord_source).getControl('tag1')
                    tag1_button.Model.ImageURL = KONST.URL_IMGS+'punkt_%s.png' %sel
                    
                    self.mb.tree_write(tree,Path)
        except:
            log(inspect.stack,tb())

        
        
        
   
