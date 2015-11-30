# -*- coding: utf-8 -*-

import unohelper
from math import sqrt
from shutil import copyfile
import json

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

       
    def erzeuge_Tag1_Container(self,ord_source,X,Y,window_parent=None):
        if self.mb.debug: log(inspect.stack)

        Width = KONST.BREITE_TAG1_CONTAINER
        Height = KONST.HOEHE_TAG1_CONTAINER
        
        posSize = X,Y,Width,Height 
        flags = 1+16+32+128
        #flags=1+32+64+128

        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize,flags,window_parent)

        # create Listener
        listener = Tag_Container_Listener()
        fenster_cont.addMouseListener(listener) 
        listener.ob = fenster  

        self.erzeuge_ListBox_Tag1(fenster, fenster_cont,ord_source,window_parent)
            
    def erzeuge_ListBox_Tag1(self,window,cont,ord_source,window_parent):
        if self.mb.debug: log(inspect.stack)
        
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 4 ,  4 , 
                                       KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control.setMultipleMode(False)
        model.Border = 0
        model.BackgroundColor = KONST.FARBE_ORGANON_FENSTER
        
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
        
        
        tag_item_listener = Tag1_Item_Listener(self.mb,window,ord_source)
        tag_item_listener.window_parent = window_parent
            
        control.addItemListener(tag_item_listener)
        
        cont.addControl('Eintraege_Tag1', control)
        
        
            
    def erzeuge_Tag2_Container(self,ordinal,X,Y,window_parent=None):
        if self.mb.debug: log(inspect.stack)
        try:
            
            controls = []
            icons_gallery,icons_prj_folder = self.get_icons()

            # create Listener
            listener = Tag_Container_Listener()
            listener2 = Tag2_Images_Listener(self.mb)
            listener2.ordinal = ordinal
            listener2.icons_dict = {}
            listener2.window_parent = window_parent
            
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

            posSize = X,Y,breite + 20,y +25 
            flags = 1+16+32+128

            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize,flags,parent=window_parent)
            fenster_cont.addMouseListener(listener) 
            listener.ob = fenster  
            
            for control in controls:
                fenster_cont.addControl('wer',control)
            
            listener2.win = fenster
            
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
            kommender_eintrag = self.mb.props['ORGANON'].kommender_Eintrag
            
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
            
            # neuen dateinamen herausfinden
            cur_text = self.mb.doc.Text.createTextCursor()
            cur_text.gotoRange(bm.Anchor,False)
            cur_text.goRight(60,True)
            neuer_Name = cur_text.String.split('\n')[0]
            
            # alte Datei in Helferdatei speichern
            orga_sec_name_alt = self.mb.props['ORGANON'].dict_bereiche['ordinal'][zeilenordinal]
            self.mb.props[T.AB].tastatureingabe = True
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(helfer_url,orga_sec_name_alt)
             
            # erzeuge neue Zeile
            nr_neue_zeile = self.mb.class_Baumansicht.erzeuge_neue_Zeile('dokument',neuer_Name)
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
            orga_sec_name_alt = self.mb.props['ORGANON'].dict_bereiche['ordinal'][zeilenordinal]
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
            self.mb.tags['ordinale'][ordinal_neue_zeile] = copy.deepcopy(self.mb.tags['ordinale'][zeilenordinal])
            
            tree = self.mb.props['ORGANON'].xml_tree
            root = tree.getroot()
            alt = root.find('.//'+zeilenordinal)
            neu = root.find('.//'+ordinal_neue_zeile)
            
            neu.attrib['Tag1'] = alt.attrib['Tag1']
            neu.attrib['Tag2'] = alt.attrib['Tag2']
            neu.attrib['Tag3'] = alt.attrib['Tag3']
            
            self.mb.class_Sidebar.erzeuge_sb_layout()
                
        except Exception as e:
            self.mb.nachricht('teile_text ' + str(e),"warningbox")
            log(inspect.stack,tb())
            
    def teile_text_batch(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB != 'ORGANON':
            self.mb.popup(LANG.FUNKTIONIERT_NUR_IM_PROJEKT_TAB)
            
        ttb = Teile_Text_Batch(self.mb)
        ttb.erzeuge_fenster()
    
    
    def vereine_dateien(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            props = self.mb.props['ORGANON']
            selektiert = props.selektierte_zeile
            
            sichtbare = [props.dict_zeilen_posY[b][0] for b in sorted(props.dict_zeilen_posY)]
            index = sichtbare.index(selektiert)
            
            
            def pruefe_ob_kombi_moeglich():
                            
                if index > len(sichtbare) -2:
                    self.mb.popup(LANG.KEINE_KOMBINATION_MOEGLICH,1)
                    return False
                
                nachfolger = sichtbare[index + 1]
                
                tree = props.xml_tree
                root = tree.getroot()
                sel_xml = root.find('.//'+selektiert)
                nachfolger_xml = root.find('.//'+nachfolger)
                
                lvl_sel = sel_xml.attrib['Lvl']
                lvl_nach = nachfolger_xml.attrib['Lvl']
                
                if T.AB != 'ORGANON': 
                    self.mb.popup(LANG.FUNKTIONIERT_NUR_IM_PROJEKT_TAB)  
                    return False, None    
                  
                elif nachfolger in props.dict_ordner[props.Papierkorb]:
                    self.mb.popup(LANG.KEINE_KOMBINATION_MOEGLICH,1)
                    return False, None    
                
                elif lvl_sel > lvl_nach:
                    self.mb.popup(LANG.KEINE_KOMBINATION_MOEGLICH,1)
                    return False, None    
                
                elif nachfolger in props.dict_ordner:
                    elems = nachfolger_xml.findall('.//')
                    if len(elems) > 0:
                        self.mb.popup(LANG.KEINE_KOMBINATION_MOEGLICH,1)
                        return False, None    
                    
                return True, nachfolger
            
            ok, nachfolger = pruefe_ob_kombi_moeglich()
            
            if not ok:
                return
                    
            url1 = self.get_pfad(selektiert)
            sec_name1 = 'OrgInnerSec' + selektiert.replace('nr','')
             
             
            url2 = self.get_pfad(nachfolger)
            sec_name2 = 'OrgInnerSec' + nachfolger.replace('nr','')
             
            doc = self.lade_doc_kombi(url1,url2,sec_name1,sec_name2)
             
            self.mb.class_Baumansicht.selektiere_zeile(nachfolger)
             
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Overwrite'
            prop.Value = True
            doc.storeToURL(url1,(prop,))
            doc.close(False)
              
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FileURL = url1
            sections = self.mb.doc.TextSections
            sec = sections.getByName(sec_name1)
                
            par_sec = sec.ParentSection
            SFLink_helfer = self.mb.sec_helfer.FileLink
            
            # Das Setzen des FileLinks löst den VC Listener aus
            # Er wird via selektiere_zeile() und schalte_sichtbarkeit...() wieder gesetzt
            self.mb.Listener.remove_VC_selection_listener()
            par_sec.setPropertyValue('FileLink',SFLink_helfer)
            par_sec.setPropertyValue('FileLink',SFLink)
                
            papierkorb = self.mb.props[T.AB].Papierkorb
            ordinal = self.mb.props[T.AB].selektierte_zeile
            self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(nachfolger,papierkorb,'inPapierkorbEinfuegen')
            self.mb.class_Baumansicht.selektiere_zeile(selektiert)

        except:
            log(inspect.stack,tb())
    
    
    def get_pfad(self,ord):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['ORGANON']
        
        sec_name = props.dict_bereiche['ordinal'][ord]
        pfad = props.dict_bereiche['Bereichsname'][sec_name]    
        
        url = uno.systemPathToFileUrl(pfad)
        return url   
    
    
    def lade_doc_kombi(self,url1,url2,sec_name1,sec_name2):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
            
            URL="private:factory/swriter"
            doc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,))
            
            newSection0 = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection0.setName('parent')
            
            SFLink1 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink1.FileURL = url1
            newSection1 = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection1.setPropertyValue('FileLink',SFLink1)
            newSection1.setName('test1')
            
            SFLink2 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink2.FileURL = url2
            newSection2 = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection2.setPropertyValue('FileLink',SFLink2)
            newSection2.setName('test2')
             
            cur = doc.Text.createTextCursor()            
            vc = doc.CurrentController.ViewCursor

            doc.Text.insertTextContent(vc, newSection2, False)
            vc.gotoStart(False)
            doc.Text.insertTextContent(vc, newSection1, False)
            
            cur.gotoStart(False)
            cur.gotoEnd(True)
            doc.Text.insertTextContent(cur, newSection0, True)
            
            newSection1.dispose()
            newSection2.dispose()
            
            cur.gotoEnd(False)
            cur.goLeft(1,True)
            cur.setString('')
            
            sections = doc.TextSections
            sec = sections.getByName(sec_name1)
            sec.dispose()
            sec = sections.getByName(sec_name2)
            sec.dispose()
            
            newSection0.setName(sec_name1)
   
            return doc
        except:
            log(inspect.stack,tb())  
    
            
    def verbotene_buchstaben_austauschen(self,term):
                            
        verbotene = '<>:"/\\|?*'

        term =  ''.join(c for c in term if c not in verbotene).strip()
        if term != '':
            return term
        else:
            return 'invalid_name'
        
    def waehle_farbe(self,initial_value=0):
        if self.mb.debug: log(inspect.stack)
        
        cp = self.mb.createUnoService("com.sun.star.ui.dialogs.ColorPicker")
        
        values = cp.getPropertyValues()
        values[0].Value = initial_value
        cp.setPropertyValues(values)
        
        cp.execute()
        cp.dispose()
        
        farbe = cp.PropertyValues[0].Value

        return farbe
    
    def dezimal_to_rgb(self,farbe):
        import struct
             
        f1 = hex(farbe).lstrip('0x')
        if len(f1) < 6:
            f1 = (6-len(f1)) * '0' + f1
        return struct.unpack('BBB',bytes.fromhex(f1))   
    
    def dezimal_to_hex(self,farbe):
        import struct
             
        f1 = hex(farbe).lstrip('0x')
        if len(f1) < 6:
            f1 = (6-len(f1)) * '0' + f1
        return f1 
            
    def folderpicker(self,filepath=None,sys=True):
        if self.mb.debug: log(inspect.stack)
        
        folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        if filepath != None:
            folderpicker.setDisplayDirectory(filepath)
        folderpicker.execute()
        
        if folderpicker.Directory == '':
            return None
        
        if sys:
            return uno.fileUrlToSystemPath(folderpicker.getDirectory())
        else:
            return folderpicker.getDirectory()
    
    def filepicker(self,filepath=None,filter=None,sys=True):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        
        # Bug in Office: funktioniert nicht
        if filepath != None:
            Filepicker.setDisplayDirectory(filepath)
            
            
        if filter != None:
            Filepicker.appendFilter('lang_py_file','*.' + filter)
            
        Filepicker.execute()

        if Filepicker.Files == '':
            return None
        if sys:
            return uno.fileUrlToSystemPath(Filepicker.Files[0])
        else:
            return Filepicker.Files[0]
        
        
    def filepicker2(self,filter=None,sys=True):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        
        # Bug in Office: funktioniert nicht
#         if filepath != None:
#             Filepicker.setDisplayDirectory(filepath)
            
        if filter != None:
            Filepicker.appendFilter(*filter)
            
        Filepicker.execute()
        
        file_len = len(Filepicker.getSelectedFiles())
                   
        if file_len == 0:
            return None,False
        
        path = Filepicker.Files[0]

        if sys:
            return uno.fileUrlToSystemPath(path),True
        else:
            return path,True
        
    
    def oeffne_json(self,pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            with codecs_open(pfad) as data:  
                content = data.read().decode()  
                odict = json.loads(content)
                return odict
        except:
            log(inspect.stack,tb())  
            return None
        
    def get_zuletzt_benutzte_datei(self):
        if self.mb.debug: log(inspect.stack)
        try:
            props = self.mb.props[T.AB]
            zuletzt = props.selektierte_zeile_alt
            xml = props.xml_tree
            root = xml.getroot()
            return root.find('.//' + zuletzt).attrib['Name']
        except:
            return 'None'
        
    def mache_tag_sichtbar(self,sichtbar,tag_name):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_proj
        tags = sett['tag1'],sett['tag2'],sett['tag3']
        
        for tab_name in self.mb.props:
        
            # alle Zeilen
            controls_zeilen = self.mb.props[tab_name].Hauptfeld.Controls
            tree = self.mb.props[tab_name].xml_tree
            root = tree.getroot()
            
            gliederung  = None
            if sett['tag3']:
                gliederung = self.mb.class_Gliederung.rechne(tree)
            
            if not sichtbar:
                for contr_zeile in controls_zeilen:
                    ord_zeile = contr_zeile.AccessibleContext.AccessibleName
                    if ord_zeile == self.mb.props[T.AB].Papierkorb:
                        continue
                    
                    self.mb.class_Baumansicht.positioniere_icons_in_zeile(contr_zeile,tags,gliederung)
                    tag_contr = contr_zeile.getControl(tag_name)
                    tag_contr.dispose()
 
                    
            if sichtbar:
                for contr_zeile in controls_zeilen:                    

                    ord_zeile = contr_zeile.AccessibleContext.AccessibleName
                    if ord_zeile == self.mb.props[T.AB].Papierkorb:
                        continue
                    
                    zeile_xml = root.find('.//'+ord_zeile)
                    
                    if tag_name == 'tag1':
                        farbe = zeile_xml.attrib['Tag1']
                        url = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % farbe
                        listener = self.mb.class_Baumansicht.tag1_listener
                    elif tag_name == 'tag2':
                        url = zeile_xml.attrib['Tag2']
                        listener = self.mb.class_Baumansicht.tag2_listener
                    elif tag_name == 'tag3':
                        url = ''
                    
                    if tag_name in ('tag1','tag2'):
                        PosX,PosY,Width,Height = 0,2,16,16
                        control_tag1, model_tag1 = self.mb.createControl(self.mb.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
                        model_tag1.ImageURL = url
                        model_tag1.Border = 0
                        control_tag1.addMouseListener(listener)
                    else:
                        PosX,PosY,Width,Height = 0,2,16,16
                        control_tag1, model_tag1 = self.mb.createControl(self.mb.ctx,"FixedText",PosX,PosY,Width,Height,(),() )     
                        model_tag1.TextColor = KONST.FARBE_GLIEDERUNG
                        
                    contr_zeile.addControl(tag_name,control_tag1)
                    self.mb.class_Baumansicht.positioniere_icons_in_zeile(contr_zeile,tags,gliederung)
                    
                    
    def pruefe_galerie_eintrag(self):
        if self.mb.debug: log(inspect.stack)
        
        gallery = self.mb.createUnoService("com.sun.star.gallery.GalleryThemeProvider")
            
        if 'Organon Icons' not in gallery.ElementNames:
            
            paths = self.mb.smgr.createInstance( "com.sun.star.util.PathSettings" )
            gallery_pfad = uno.fileUrlToSystemPath(paths.Gallery_writable)
            gallery_ordner = os.path.join(gallery_pfad,'Organon Icons')
                    
            entscheidung = self.mb.nachricht(LANG.BENUTZERDEFINIERTE_SYMBOLE_NUTZEN %gallery_ordner,"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return False
            elif entscheidung == 2:
                try:
                    iGal = gallery.insertNewByName('Organon Icons')  
                    path_icons = os.path.join(self.mb.path_to_extension,'img','Organon Icons')
                    
                    from shutil import copy 
                    
                    # Galerie anlegen
                    if not os.path.exists(gallery_ordner):
                        os.makedirs(gallery_ordner)
                    
                    # Organon Icons einfuegen
                    for (dirpath,dirnames,filenames) in os.walk(path_icons):
                        for f in filenames:
                            url_source = os.path.join(dirpath,f)
                            url_dest   = os.path.join(gallery_ordner,f)
                            
                            copy(url_source,url_dest)
 
                            url = uno.systemPathToFileUrl(url_dest)
                            iGal.insertURLByIndex(url,0)
                    
                    return True
                
                except:
                    log(inspect.stack,tb())
        
        return True
    
    def get_writer_shortcuts(self):
        if self.mb.debug: log(inspect.stack)
        
        ctx = self.mb.ctx
        smgr = self.mb.ctx.ServiceManager
           
        config_provider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider",ctx)
  
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = "nodepath"
        prop.Value = "org.openoffice.Office.Accelerators"
               
        config_access = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
        
        glob = config_access.PrimaryKeys.Global 
        modules = config_access.PrimaryKeys.Modules
        textdoc = modules.getByName('com.sun.star.text.TextDocument')
        
        shortcuts = {}
        
        elements = textdoc.ElementNames
        
        for e in elements:
            try:
                if e in textdoc.ElementNames:
                    sc = textdoc.getByName(e)
                    shortcuts.update({e:sc.Command})
            except:
                shortcuts.update({e:'?'})
                
        elements = glob.ElementNames
        
        for e in elements:
            try:
                if e in textdoc.ElementNames:
                    sc = textdoc.getByName(e)
                    shortcuts.update({e:sc.Command})
            except:
                shortcuts.update({e:'?'})
        
        
        
        erweitert = ['F'+str(a) for a in range(1,13)]
        erweitert.extend(['DOWN','LEFT','RIGHT','UP'])
        
        mod1 = []
        mod2 = []
        shift = []
        
        mod1_mod2 = []
        shift_mod1 = []
        shift_mod2 = []
        shift_mod1_mod2 = []
        
        for s in shortcuts:
            cmd = s.split('_')
            
            if len(cmd) == 2:
                if len(cmd[0]) == 1 or cmd[0] in erweitert:
                    
                    if 'MOD1' in s:
                        mod1.append(cmd[0])
                    elif 'MOD2' in s:
                        mod2.append(cmd[0])
            elif len(cmd) == 3:
                if len(cmd[0]) == 1 or cmd[0] in erweitert:
                    if 'SHIFT_MOD1' in s:
                        shift_mod1.append(cmd[0])
                    elif 'SHIFT_MOD2' in s:
                        shift_mod2.append(cmd[0])
                    elif 'MOD1_MOD2' in s:
                        mod1_mod2.append(cmd[0])
            elif len(cmd) == 4:
                if len(cmd[0]) == 1 or cmd[0] in erweitert:
                    if 'SHIFT_MOD1_MOD2' in s:
                        shift_mod1_mod2.append(cmd[0])
            
            
        used = {
                2:sorted(mod1),
                3:sorted(shift_mod1),
                4:sorted(mod2),
                5:sorted(shift_mod2),
                6:sorted(mod1_mod2),
                7:sorted(shift_mod1_mod2)
                }
        return used
        
        
    #def find_differences(self,obj):
#         ctx = uno.getComponentContext()
#         smgr = ctx.ServiceManager
#         desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
#         doc = desktop.getCurrentComponent() 
#         current_Contr = doc.CurrentController
#         viewcursor = current_Contr.ViewCursor
#         
#         object = obj
#         max_lvl = 3
        
    def get_attribs(self,obj,max_lvl):
        
        results = {}
        def get_attr(obj,lvl):
            
            for key in dir(obj):
                
                try:
                    value = getattr(obj, key)
                    if 'callable' in str(type(value)):
                        continue
                except :
                    #print(key)
                    continue
        
                if key not in results:
                    if type(value) in (
                                       type(None),
                                       type(True),
                                       type(1),
                                       type(.1),
                                       type('string'),
                                       type(()),
                                       type([]),
                                       type(b''),
                                       type(r''),
                                       type(u'')
                                       ):
                        results.update({key: value})
                        
                    elif lvl < max_lvl:
                        try:
                            results.update({key: get_attr(value,lvl+1)})
                        except:
                            pass
        
        
        get_attr(obj, 0)
        return results
        
    def find_differences(self,dict1,dict2):
        diff = []
        
        def findDiff(d1, d2, path=""):
            for k in d1.keys():
                if not k in d2:
                    print (path, ":")
                    print (k + " as key not in d2", "\n")
                else:
                    if type(d1[k]) is dict:
                        if path == "":
                            path = k
                        else:
                            path = path + "->" + k
                        findDiff(d1[k],d2[k], path)
                    else:
                        if d1[k] != d2[k]:
                            diff.append((path,k,d1[k],d2[k]))
                            path = ''
        findDiff(dict1,dict2)
        return diff
    
    
    def leere_hf(self):
        if self.mb.debug: log(inspect.stack)
        
        contr = self.mb.prj_tab.getControl('Hauptfeld_aussen') 
        contr.dispose()
        contr = self.mb.prj_tab.getControl('ScrollBar')
        contr.dispose()


    def erzeuge_treeview_mit_checkbox(self,tab_name='ORGANON',listener_innen=None,pos=None,auswaehlen=None):
        if self.mb.debug: log(inspect.stack)
        
        control_innen, model = self.mb.createControl(self.mb.ctx,"Container",20,0,400,100,(),() )
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        
        if auswaehlen:
            listener_innen = Auswahl_CheckBox_Listener(self.mb)
        
        x,y,ctrls = self.erzeuge_treeview_mit_checkbox_eintraege(tab_name,
                                                                 control_innen,
                                                                 listener=listener_innen,
                                                                 auswaehlen=auswaehlen)
        control_innen.setPosSize(0, 0,x,y + 20,12)
        
        if not pos:
            X,Y = 0,0
        else:
            X,Y = pos
        
        x += 40
        y += 10
        
        
        erzeuge_scrollbar = False
        if y > 800:
            y = 800
            erzeuge_scrollbar = True
            
        posSize = X,Y,x,y
        
        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
        fenster_cont.Model.Text = LANG.AUSWAHL
        fenster_cont.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND

        fenster_cont.addControl('Container_innen', control_innen)
        
        if auswaehlen:
            listener_innen.ctrls = ctrls
        
        if erzeuge_scrollbar:
            self.mb.erzeuge_Scrollbar(fenster_cont,(0,0,0,y),control_innen)
            self.mb.class_Mausrad.registriere_Maus_Focus_Listener(fenster_cont)
        
        return y,fenster,fenster_cont,control_innen,ctrls
    
    
    def erzeuge_treeview_mit_checkbox_eintraege(self,tab_name,control_innen,listener=None,auswaehlen=None):
        if self.mb.debug: log(inspect.stack)
        try:
            sett = self.mb.settings_exp
            
            tree = self.mb.props[tab_name].xml_tree
            root = tree.getroot()
            
            baum = []
            self.mb.class_XML.get_tree_info(root,baum)
            
            y = 10
            x = 10
                
            # Titel AUSWAHL
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x,y ,300,20,(),() )  
            control.Text = LANG.AUSWAHL_TIT
            model.FontWeight = 150.0
            model.TextColor = KONST.FARBE_SCHRIFT_DATEI
            control_innen.addControl('Titel', control)
            
            y += 30
            
            # Untereintraege auswaehlen
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40,y ,300,20,(),() )  
            control.Text = LANG.ORDNER_CLICK
            model.FontWeight = 150.0
            model.TextColor = KONST.FARBE_SCHRIFT_DATEI
            control_innen.addControl('ausw', control)
            x_pref = control.getPreferredSize().Width + x + 40
            
            control, model = self.mb.createControl(self.mb.ctx,"CheckBox",x+20,y ,20,20,(),() )  
            control.State = sett['auswahl']
            control.ActionCommand = 'untereintraege_auswaehlen'

            if listener:
                control.addActionListener(listener)
                control.ActionCommand = 'untereintraege_auswaehlen'
            control_innen.addControl('Titel', control)
    
            y += 30
            
            ctrls = {}
            
            
            for eintrag in baum:
    
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
                
                if art == 'waste':
                    break
                
                control1, model1 = self.mb.createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,400,20,(),() )  
                control1.Text = name
                control_innen.addControl('Titel', control1)
                pref = control1.getPreferredSize().Width
                
    
                if x_pref < x + 40+20*int(lvl) + pref:
                    x_pref = x + 40+20*int(lvl) + pref
                
                control2, model2 = self.mb.createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
                model2.Border = False
                
                if art in ('dir','prj'):
                    model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
                    model1.TextColor = KONST.FARBE_SCHRIFT_ORDNER
                else:
                    model2.ImageURL = 'private:graphicrepository/res/sx03150.png' 
                    model1.TextColor = KONST.FARBE_SCHRIFT_DATEI
                control_innen.addControl('Titel', control2)   
                  
                    
                control3, model3 = self.mb.createControl(self.mb.ctx,"CheckBox",x+20*int(lvl),y ,20,20,(),() )  
                control_innen.addControl(ordinal, control3)
                if listener:
                    control3.addActionListener(listener)
                    control3.ActionCommand = ordinal+'xxx'+name
                    if auswaehlen:
                        if ordinal in sett['ausgewaehlte']:
                            model3.State = sett['ausgewaehlte'][ordinal]
                
                ctrls.update({ordinal:[control1,control2,control3]})
                
                y += 20 
                
            return x_pref,y,ctrls
        except:
            log(inspect.stack,tb())

            
    def update_organon_templates(self):  
        if self.mb.debug: log(inspect.stack)
        
        templ = self.mb.settings_orga['templates_organon']
        pfad = templ['pfad']
        
        if pfad == '':
            return
        
        templates = []
        
        for root, dirs, files in os.walk(pfad):
            break
        
        for d in dirs:
            name = d.split('.')
            if len(name) == 2:
                if name[1] == 'organon':
                    templates.append(name[0])
                    
        templ['templates'] = templates
        
    def vorlage_speichern(self,pfad,name):
        if self.mb.debug: log(inspect.stack)

        pfad_zu_neuem_ordner = os.path.join(pfad,name)
        
        tree = copy.deepcopy(self.mb.props['ORGANON'].xml_tree)
        root = tree.getroot()
        
        all_elements = root.findall('.//')
        ordinale = []
        
        for el in all_elements:
            ordinale.append(el.tag)        
    
        self.mb.class_Export.kopiere_projekt(name,pfad_zu_neuem_ordner,ordinale,tree,self.mb.tags,True)  
        os.rename(pfad_zu_neuem_ordner,pfad_zu_neuem_ordner+'.organon')
    
    def kopiere_ordner(self,src, dst):
        if self.mb.debug: log(inspect.stack)
        
        import shutil,errno
        try:
            shutil.copytree(src, dst)
        except OSError as exc: 
            if exc.errno == errno.ENOTDIR:
                shutil.copy(src, dst)
            else: 
                log(inspect.stack,tb())
                
    def projekt_umbenannt_speichern(self,alter_pfad,neuer_pfad,name):
        if self.mb.debug: log(inspect.stack)
        
        alter_name = os.path.basename(alter_pfad)

        self.kopiere_ordner(alter_pfad,neuer_pfad)
         
        alt = os.path.join(neuer_pfad,alter_name)
        neu = os.path.join(neuer_pfad,name + '.organon')

        os.rename(alt,neu)
         
        pfad_el_tree = os.path.join(neuer_pfad,'Settings','ElementTree.xml')
              
        xml_tree = self.mb.ET.parse(pfad_el_tree)
        root = xml_tree.getroot()
         
        prj_xml = root.find(".//*[@Art='prj']")
        prj_xml.attrib['Name'] = name
         
        xml_tree.write(pfad_el_tree)
        
        

class Teile_Text_Batch():
    def __init__(self,mb):
        self.mb = mb
        
    def erzeuge_fenster(self):
        if self.mb.debug: log(inspect.stack)
        
     
        try:
            self.dialog_batch_devide()
        except:
            log(inspect.stack,tb())
            
            
    def dialog_batch_devide_elemente(self):
        if self.mb.debug: log(inspect.stack)
        
        listener = self.listener
        
        controls = [
            10,
            ('control_Titel',"FixedText",        
                                    20,0,250,20,    
                                    ('Label','FontWeight'),
                                    (LANG.TEXT_BATCH_DEVIDE ,150),                  
                                    {}
                                    ), 
            35,
            ('control_Text',"Edit",        
                                    20,0,350,20,    
                                    (),
                                    (),                  
                                    {}
                                    ), 
                    
                    
            40, ]
        
        elemente = 'GANZES_WORT','REGEX','UEBERSCHRIFTEN','LEERZEILEN'
                
                
        for el in elemente:
            controls.extend([
            ('control_{}'.format(el),"CheckBox",      
                                    20,0,200,20,    
                                    ('Label','State'),
                                    (getattr(LANG, el),0),        
                                    {'setActionCommand':el,'addActionListener':(listener,)} 
                                    ),  
            25 if el != 'REGEX' else 45])
            
        controls.extend([
            -35,
            ('control_start',"Button",      
                                    290,0,80,30,    
                                    ('Label',),
                                    (LANG.START,),        
                                    {'setActionCommand':'start','addActionListener':(listener,)} 
                                    ),  
            0])
        
           
        return controls

 
    def dialog_batch_devide(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
            X = self.mb.dialog.Size.Width
            Y = posSize_main.Y         

            # Listener erzeugen 
            self.listener = Batch_Text_Devide_Listener(self.mb)         
            
            controls = self.dialog_batch_devide_elemente()
            ctrls,pos_y = self.mb.erzeuge_fensterinhalt(controls)   
            
            self.listener.ctrls = ctrls
            self.listener.ttb = self
            
            # Hauptfenster erzeugen
            posSize = X,Y,380,210
            fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)             
              
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                fenster_cont.addControl(c,ctrls[c])
           
            self.fenster = fenster
           
        except:
            log(inspect.stack,tb())
            
            
    def werte_controls_aus(self,ctrls):
        if self.mb.debug: log(inspect.stack)
        
        text = ctrls['control_Text'].Text
        ganzes_wort = ctrls['control_GANZES_WORT'].State
        regex = ctrls['control_REGEX'].State
        ueberschriften = ctrls['control_UEBERSCHRIFTEN'].State
        leerzeilen = ctrls['control_LEERZEILEN'].State
        
        if text == '' and not ueberschriften and not leerzeilen:
            self.mb.nachricht(LANG.NICHTS_AUSGEWAEHLT_BATCH,'infobox')
        else:
            args = text,ganzes_wort,regex,ueberschriften,leerzeilen
            self.fenster.dispose()
            self.run(args)


    def run(self,args):
        if self.mb.debug: log(inspect.stack)
        
        try:
            text,ganzes_wort,regex,ueberschriften,leerzeilen = args
            
            ergebnis1 = []
            ergebnis2 = []
            ergebnis3 = []
            
            url = self.get_pfad() 
            doc = self.lade_doc(url)  
            
            erster = self.get_ersten_paragraph(doc)    
                   
            if text != '':
                ergebnis3 = self.get_suchbegriff(doc,suchbegriff=text,ganzes_wort=ganzes_wort,regex=regex)
            if ueberschriften:
                ergebnis1 = self.get_ueberschriften(doc)
            if leerzeilen:
                ergebnis2 = self.get_leerzeilen(doc)
                
            ordnung = sorted(erster + ergebnis1 + ergebnis2 + ergebnis3, key = lambda x : (x[2])) 
    
            ausgesonderte,ordnung = self.erstelle_bereiche(doc,ordnung)              
            ordnung = [o for o in ordnung if o not in ausgesonderte] 
            
            if len(ordnung) > 10:
                entscheidung = self.mb.nachricht(LANG.ERSTELLT_WERDEN.format(len(ordnung)),"warningbox",16777216)
                # 3 = Nein oder Cancel, 2 = Ja
                if entscheidung == 3:
                    doc.close(False)
                    return
  
            tree = self.erstelle_tree(ordnung)  
            
            root = tree.getroot()
            anz = len(root.findall('.//'))
            
            if anz < 2:
                self.mb.nachricht(LANG.KEINE_TRENNUNG,'infobox')
                return
  
            
            speicherordner = self.mb.pfade['odts']
            pfad_helfer_system = os.path.join(speicherordner,'batchhelfer.odt')
            pfad_helfer = uno.systemPathToFileUrl(pfad_helfer_system)
            
            doc.storeToURL(pfad_helfer,())
            doc.close(False)
               
            tree_new = self.fuege_tree_in_xml_ein(tree)
            anz = self.neue_Dateien_erzeugen(pfad_helfer, tree)
            
            os.remove(pfad_helfer_system)
            self.schreibe_neuen_elementtree(tree_new,anz)
            self.lege_dict_sbs_an(tree)
            
            self.mb.Listener.remove_Undo_Manager_Listener()
            self.mb.Listener.remove_VC_selection_listener()
            
            self.mb.class_Projekt.lade_Projekt2()
            
        except:
            log(inspect.stack,tb())
            try:
                doc.close(False)
            except:
                pass
            
            
        
    def get_pfad(self):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['ORGANON']
        selektiert = props.selektierte_zeile
        
        sec_name = props.dict_bereiche['ordinal'][selektiert]
        pfad = props.dict_bereiche['Bereichsname'][sec_name]    
        
        url = uno.systemPathToFileUrl(pfad)
        return url
    
    
    def lade_doc(self,url):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
            
            URL="private:factory/swriter"
            doc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,))
                                                    
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FileURL = url
            SFLink.FilterName = 'writer8'
            
            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection.setPropertyValue('FileLink',SFLink)
            newSection.setName('test')
            
            cur = doc.Text.createTextCursor()
            doc.Text.insertTextContent(cur, newSection, True)

            sections = doc.TextSections
            secs = []
            for i in range(sections.Count):
                secs.append(sections.getByIndex(i))

            for s in secs:
                s.dispose()
                
            return doc
        except:
            log(inspect.stack,tb())
            
    
     
    def get_ueberschriften(self,doc):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sd = doc.createSearchDescriptor()
            sd.SearchStyles = True
            viewcursor = doc.CurrentController.ViewCursor
             
            StyleFamilies = doc.StyleFamilies
            ParagraphStyles = StyleFamilies.getByName("ParagraphStyles")
             
            headings = []
            display_names = []

            for p in ParagraphStyles.ElementNames:
                if 'Heading' in p[0:7]:
                    headings.append(p)
                    elem = ParagraphStyles.getByName(p)
                    display_names.append(elem.DisplayName)
             
            ergebnisse = []

            for dn in range(len(display_names)):
                sd.SearchString = display_names[dn]
                ergebnisse.append(doc.findAll(sd))
             
            ordnung = []
            x = 0

            for e in ergebnisse:
                for count in range(e.Count):
                    erg = e.getByIndex(count)
                    viewcursor.gotoRange(erg.Start,False)
                    vd = int(doc.CurrentController.ViewData.split(';')[1])
                    ordnung.append([erg.ParaStyleName,viewcursor.Page,vd,erg,erg.String])
                    x += 1

        except:
            log(inspect.stack,tb())
            ordnung = []
         
        return ordnung
    
    
    def get_leerzeilen(self,doc):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sd = doc.createSearchDescriptor()
            viewcursor = doc.CurrentController.ViewCursor
            
            regex_leerzeile = '^$'
            
            ergebnisse = []

            sd.SearchRegularExpression = True
            sd.SearchString = regex_leerzeile

            ergebnisse1 = doc.findAll(sd)
            
            sd.SearchString = '^\s*$'
            ergebnisse2 = doc.findAll(sd)
            
            ordnung = []
            zeilen = []
            x = 0
            
            try:
                for count in range(ergebnisse1.Count):
                    erg = ergebnisse1.getByIndex(count)
                    if erg.ParaStyleName == 'Footnote':
                        continue
                    viewcursor.gotoRange(erg.Start,False)
                    vd = int(doc.CurrentController.ViewData.split(';')[1])
                    zeilen.append(['leer',viewcursor.Page,vd,viewcursor.Start,erg.String])
                    x += 1
            except:
                log(inspect.stack,tb())
                
            try: 
                for count in range(ergebnisse2.Count):
                    erg = ergebnisse2.getByIndex(count)
                    if erg.ParaStyleName == 'Footnote':
                            continue
                    viewcursor.gotoRange(erg.Start,False)
                    vd = int(doc.CurrentController.ViewData.split(';')[1])
                    zeilen.append(['leer',viewcursor.Page,vd,viewcursor.Start,erg.String])
                    x += 1
            except:
                log(inspect.stack,tb())
                                
        except:
            log(inspect.stack,tb())
            
        
        return zeilen
    
    
    def get_suchbegriff(self,doc,suchbegriff,ganzes_wort,regex):
        if self.mb.debug: log(inspect.stack)
        
        try:
            zeilen = []
            
            sd = doc.createSearchDescriptor()
            viewcursor = doc.CurrentController.ViewCursor

            if regex:
                sd.SearchRegularExpression = True
            if ganzes_wort:
                sd.SearchWords = True
                
            sd.SearchString = suchbegriff
            ergebnisse = doc.findAll(sd)
            
            x = 0

            for count in range(ergebnisse.Count):
                
                erg = ergebnisse.getByIndex(count)
                #erg.CharBackColor = 502
                viewcursor.gotoRange(erg.Start,False)

                vd = int(doc.CurrentController.ViewData.split(';')[1])
                #print(viewcursor.CharStyleName + viewcursor.ParaStyleName)
                if 'footnote' not in viewcursor.ParaStyleName.lower():
                    zeilen.append(['suchbegriff',viewcursor.Page,vd,erg,erg.String])
                    x += 1
               
        except:
            log(inspect.stack,tb())
        
        return zeilen
               
     
    def erstelle_bereiche(self,doc,ordnung):
        if self.mb.debug: log(inspect.stack)
        
        cur = doc.Text.createTextCursor()
        cur2 = doc.Text.createTextCursor()
        
        x = self.mb.props['ORGANON'].kommender_Eintrag
        ausgesonderte = []
        
        try:                
            # pro Ueberschrift 1 Ordner
            for o in range(len(ordnung)):
                 
                fund = ordnung[o][3]
                if o + 1 < len(ordnung):
                    fund_ende = ordnung[o+1][3]
                else:
                    fund_ende = None
                
                # Cursor setzen
                if ordnung[o][0] == 'leer':
                    cur.gotoRange(fund.Start,False)
                    cur.goRight(1,False)
                    fund = cur.Start
                                                
                cur.gotoRange(fund.Start,False)
                
#                 cur.setString('ANFANG'+str(o))
#                 cur.CharBackColor = 502
                
                if fund_ende != None:
                    cur.gotoRange(fund_ende.Start,True)
                else:
                    cur.gotoEnd(True)
                    
                
                # leere Bereiche aussparen
                if cur.String.strip() == '':
                    ausgesonderte.append(ordnung[o])
                    continue
                
#                 cur2.gotoRange(cur,False)
#                 cur2.collapseToEnd()
#                 cur2.setString('ENDE'+str(o))
#                 cur2.CharBackColor = 12765426
                
                try:
                    newSection = doc.createInstance("com.sun.star.text.TextSection")
                    doc.Text.insertTextContent(cur, newSection, True)
                    newSection.setName('organon_nr{}'.format(x))
                    ordnung[o][4] = cur.String
                    x += 1
                except:
                    ausgesonderte.append(ordnung[o])
                    
        except:
            log(inspect.stack,tb())
        
        return ausgesonderte,ordnung
    
     
    def setze_attribute(self,element,name,art,level,parent):
        if self.mb.debug: log(inspect.stack)
        
        if art == 'dir':
            zustand = 'auf'
        else:
            zustand = '-'
        
        element.attrib['Name'] = name.split('\n')[0]
        element.attrib['Zustand'] = zustand
        element.attrib['Sicht'] = 'ja'
        element.attrib['Parent'] = parent.tag
        element.attrib['Lvl'] = str(level)
        element.attrib['Art'] = art
        
        element.attrib['Tag1'] = 'leer'
        element.attrib['Tag2'] = 'leer'
        element.attrib['Tag3'] = 'leer'
     
     
    def erstelle_tree(self,geordnete):
        if self.mb.debug: log(inspect.stack)
        
        # XML TREE
        et = self.mb.ET    
        root = et.Element('root')
        root.attrib['NameH'] = 'AAA'
        tree = et.ElementTree(root)
        
        props = self.mb.props['ORGANON']
        
        kE = props.kommender_Eintrag
        el = root
        
        selektiert = props.selektierte_zeile
        root_orig = props.xml_tree.getroot()
        sel_xml = root_orig.find('.//'+selektiert)
        lvl_sel = int(sel_xml.attrib['Lvl'])
        
        lvl = lvl_sel - 1
        erste_pg = True
        
        try:
            for o in range(len(geordnete)):
                
                heading = geordnete[o][0]
                name = geordnete[o][4]
                name = name.strip() if len(name) < 60 else name[0:60].strip()
                
                if heading in ['leer','suchbegriff','erster']:
                    if erste_pg: 
                        if el.tag == 'root':
                            par = el
                        else:
                            par = root.find('.//'+el.tag)
                        lvl += 1
                    else:
                        par = root.find('.//'+el.tag+'/..')
                    el = et.SubElement(par,'nr'+str(o+kE))
                    
                    el.attrib['NameH'] = heading
                
                    par = root.find('.//'+el.tag+'/..')
                    self.setze_attribute(el,name,'pg',lvl,par)
                    erste_pg = False
                    continue
                
                elif heading > el.attrib['NameH']:
                    el = et.SubElement(el,'nr'+str(o+kE))
                    lvl += 1
                    
                elif heading == el.attrib['NameH']:
                    par = root.find('.//'+el.tag+'/..')
                    el = et.SubElement(par,'nr'+str(o+kE))
                     
                else:
                    par = el
                    while heading <= par.attrib['NameH']:
                        par = root.find('.//'+par.tag+'/..')
                        lvl -= 1
                    lvl += 1

                    el = et.SubElement(par,'nr'+str(o+kE))
                   
                el.attrib['NameH'] = heading
                
                par = root.find('.//'+el.tag+'/..')
                self.setze_attribute(el,name,'dir',lvl,par)
                erste_pg = True
        except:
            log(inspect.stack,tb())
            tree = None
            
        all = root.findall('.//')
        for a in all:
            del a.attrib['NameH']
        
        return tree
    
    
    def get_ersten_paragraph(self,doc):
        if self.mb.debug: log(inspect.stack)
        
        vc = doc.CurrentController.ViewCursor
        vc.gotoStart(False)
        erg = vc.Start
        vd = 0
        return [['erster',vc.Page,vd,erg,erg.String]]
    
    
    def neue_Dateien_erzeugen(self,pfad_helfer,tree):
        if self.mb.debug: log(inspect.stack)
        
        root = tree.getroot()
        neue_dateien = root.findall('.//')
        ordinale = [n.tag for n in neue_dateien]

        StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
        StatusIndicator.start(LANG.ERZEUGE_DATEI %(1,len(ordinale)),len(ordinale))
        
        speicherordner = self.mb.pfade['odts']
        
        zaehler = 0
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
            
            URL="private:factory/swriter"
            
            for o in ordinale:
                StatusIndicator.setValue(zaehler)
                StatusIndicator.setText(LANG.ERZEUGE_DATEI %(zaehler+1,len(ordinale)))
                
                new_doc = self.mb.desktop.loadComponentFromURL(URL,'_blank',8+32,(prop,))
                cur = new_doc.Text.createTextCursor()
                
                SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                SFLink.FileURL = pfad_helfer
                SFLink.FilterName = 'writer8'
                
                newSection = new_doc.createInstance("com.sun.star.text.TextSection")
                newSection.Name = 'OrgInnerSec' + o.replace('nr','')
                new_doc.Text.insertTextContent(cur, newSection, True)
                
                cur.goLeft(1,False)
                cur.setString(' ')
                
                newSection2 = new_doc.createInstance("com.sun.star.text.TextSection")
                newSection2.setPropertyValue('FileLink',SFLink)
                newSection2.LinkRegion = 'organon_' + o
                  
                new_doc.Text.insertTextContent(cur, newSection2, True)
                newSection2.dispose()
                
                cur.gotoEnd(False)
                cur.goLeft(1,True)
                cur.setString('')
                
                pfad = os.path.join(speicherordner,o +'.odt')
                pfad2 = uno.systemPathToFileUrl(pfad)
                new_doc.storeToURL(pfad2,())
                new_doc.close(False)
                
                zaehler += 1
            
        except:
            log(inspect.stack,tb())
            StatusIndicator.end()
            return 0
            
        StatusIndicator.end()
        return zaehler
    
    
    def fuege_tree_in_xml_ein(self,new_tree):
        if self.mb.debug: log(inspect.stack)
        
        try:        
            tree_old = copy.deepcopy(self.mb.props[T.AB].xml_tree)
            root = tree_old.getroot()
            
            name_selek_zeile = self.mb.props['ORGANON'].selektierte_zeile
            xml_selekt_zeile = root.find('.//'+name_selek_zeile)
            
            parent = root.find('.//'+xml_selekt_zeile.tag+'/..')
            liste = list(parent)
            index_sel = liste.index(xml_selekt_zeile)
            
            children_new_tree = list(new_tree.getroot())
            
            for c in range(len(children_new_tree)):
                child = children_new_tree[c]
                child.attrib['Parent'] = parent.tag
                parent.insert(index_sel+1+c,child)  
            
            return tree_old
        
        except:
            log(inspect.stack,tb())
            
            
    def schreibe_neuen_elementtree(self,tree,zaehler):
        root = tree.getroot()
        kE = int(root.attrib['kommender_Eintrag'])  
        root.attrib['kommender_Eintrag'] = str(kE + zaehler)
        
        self.mb.props['ORGANON'].kommender_Eintrag += zaehler
        
        path = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
        self.mb.tree_write(tree,path)
        
    def lege_dict_sbs_an(self,tree):
        root = tree.getroot()
        all = root.findall('.//')
        for a in all:
            self.mb.class_Tags.erzeuge_tags_ordinal_eintrag(a.tag)
            


from com.sun.star.awt import XActionListener
class Batch_Text_Devide_Listener (unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
        
        self.ctrls = None
        self.ttb = None
       
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        cmd = ev.ActionCommand
        
        if cmd == 'GANZES_WORT':
            ctrl = self.ctrls['control_REGEX']
            ctrl.State = 0
        elif cmd == 'REGEX':
            ctrl = self.ctrls['control_GANZES_WORT']
            ctrl.State = 0
        elif cmd == 'start':
            self.ttb.werte_controls_aus(self.ctrls)
        
        
    def disposing(self,ev):
        return False

        
     
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
        self.window_parent = None
       
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack) 

        url = ev.Source.Model.ImageURL
        
        if url != '':
            self.galerie_icon_im_prj_ordner_evt_loeschen()
            url = self.galerie_icon_im_prj_ordner_speichern(url) 
        else:
            if self.window_parent == None:
                # Beim Aufruf aus Calc kann nicht geloescht werden,
                # geloescht wird daher beim Schliessen von Calc
                self.galerie_icon_im_prj_ordner_evt_loeschen()
            
        self.tag2_in_allen_tabs_xml_anpassen(self.ordinal,url)

        if self.window_parent != None:
            self.icon_in_calc_anpassen(self.ordinal,url)
        
        self.win.dispose()

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
                    
                    if name == 'ORGANON':
                        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
                    else:
                        Path = os.path.join(self.mb.pfade['tabs'], name + '.xml')
                    
                    tag2_button = self.mb.props[name].Hauptfeld.getControl(ord_source).getControl('tag2')
                    if tag2_button != None:
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
        
        try:
            tree = self.mb.props['ORGANON'].xml_tree
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
        except:
            log(inspect.stack,tb())
            
    
    def icon_in_calc_anpassen(self,ord_source,url):
        if self.mb.debug: log(inspect.stack) 
        try:
            # ctx von calc
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            
            draw_page = desktop.CurrentComponent.DrawPages.getByIndex(0)
            
            form = draw_page.Forms.getByName('IMGU_'+ ord_source)
            form.ControlModels[0].setPropertyValue('ImageURL',url)
            
            if url == '':
                form.ControlModels[0].setPropertyValue('Border',2)
                form.ControlModels[0].setPropertyValue('BorderColor',4147801)
            else:
                form.ControlModels[0].setPropertyValue('Border',0)
                hintergrund = self.mb.settings_orga['organon_farben']['office']['dok_hintergrund']
                form.ControlModels[0].setPropertyValue('BorderColor',hintergrund)
                
        except:
            log(inspect.stack,tb())
   
    def mouseExited(self, ev): 
        ev.value.Source.Model.BackgroundColor = KONST.FARBE_ORGANON_FENSTER
        return False
    def mouseEntered(self, ev):    
        ev.value.Source.Model.BackgroundColor = 102
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
        
             
class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,window,ord_source):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.window = window
        self.ord_source = ord_source
        self.window_parent = None
        
    def itemStateChanged(self, ev):   
        if self.mb.debug: log(inspect.stack) 
        
        try:
            sel = ev.value.Source.Items[ev.value.Selected]
    
            # image tag1 aendern
            src = self.mb.props[T.AB].Hauptfeld.getControl(self.ord_source).getControl('tag1')
            if src != None:
                src.Model.ImageURL = KONST.URL_IMGS+'punkt_%s.png' %sel
    
            url = self.tag1_in_allen_tabs_xml_anpassen(self.ord_source,sel)
            
            # Wenn von Calc gerufen
            if self.window_parent != None:
                self.icon_in_calc_anpassen(self.ord_source,url)
    
            self.window.dispose()
        except:
            log(inspect.stack,tb())

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
                    
                    if name == 'ORGANON':
                        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
                    else:
                        Path = os.path.join(self.mb.pfade['tabs'], name + '.xml')
                    
                    tag1_button = self.mb.props[name].Hauptfeld.getControl(ord_source).getControl('tag1')
                    if tag1_button != None:
                        tag1_button.Model.ImageURL = KONST.URL_IMGS+'punkt_%s.png' %sel
                    
                    self.mb.tree_write(tree,Path)
                    
            return KONST.URL_IMGS+'punkt_%s.png' %sel
        except:
            log(inspect.stack,tb())


    def icon_in_calc_anpassen(self,ord_source,url):
        if self.mb.debug: log(inspect.stack) 
        try:
            # ctx von calc
            ctx = uno.getComponentContext()
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
            
            draw_page = desktop.CurrentComponent.DrawPages.getByIndex(0)
            
            form = draw_page.Forms.getByName('IMG_'+ ord_source)
            form.ControlModels[0].setPropertyValue('ImageURL',url)
            
            if 'punkt_leer' in url:
                form.ControlModels[0].setPropertyValue('Border',2)
                form.ControlModels[0].setPropertyValue('BorderColor',4147801)
            else:
                form.ControlModels[0].setPropertyValue('Border',0)
                hintergrund = self.mb.settings_orga['organon_farben']['office']['dok_hintergrund']
                form.ControlModels[0].setPropertyValue('BorderColor',hintergrund)
            
        except:
            log(inspect.stack,tb())
            
            
            
from com.sun.star.awt import XActionListener   
class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.ctrls = None
    
    def disposing(self,ev):
        return False

    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        sett = self.mb.settings_exp

        if ev.ActionCommand == 'untereintraege_auswaehlen':
            sett['auswahl'] = self.toggle(sett['auswahl'])
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
        else:
            ordinal,titel = ev.ActionCommand.split('xxx')
            state = ev.Source.Model.State
            sett['ausgewaehlte'].update({ordinal:state})
            
            props = self.mb.props[T.AB]
            try:
                if sett['auswahl']:
                    if ordinal in props.dict_ordner:
                        
                        tree = props.xml_tree
                        root = tree.getroot()
                        C_XML = self.mb.class_XML
                        ord_xml = root.find('.//'+ordinal)
                        
                        eintraege = []
                        # selbstaufruf nur fuer den debug
                        C_XML.selbstaufruf = False
                        C_XML.get_tree_info(ord_xml,eintraege)
                        
                        ordinale = []
                        for eintr in eintraege:
                            ordinale.append(eintr[0])
                            

                        
                        for ordn in ordinale:
                            if ordn in self.ctrls:
                                control = self.ctrls[ordn][2]
                                control.Model.State = state
                                titel = self.ctrls[ordn][1]
                                sett['ausgewaehlte'].update({ordn:state}) 

            except:
                if self.mb.debug: log(inspect.stack,tb())
                

    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:              
            return 1  


from com.sun.star.lang import XEventListener
class Window_Dispose_Listener(unohelper.Base,XEventListener):
    '''
    Closing the dialog window holding 50+ controls might
    freeze Writer. This listener closes the window
    explicitly
    
    '''
    def __init__(self,fenster,mb,ctrls):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.fenster = fenster
        self.ctrls = ctrls
    
    def disposing(self,ev):
        if self.mb.programm == 'OpenOffice':
            for ct in self.ctrls:
                for c in ct:
                    c.dispose()
            self.fenster.dispose()
            self.fenster = None
            return
        if self.mb: log(inspect.stack)
        self.fenster.dispose()
        return False
     