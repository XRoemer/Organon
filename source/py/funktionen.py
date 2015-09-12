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
            
            # neuen dateinamen herausfinden
            cur_text = self.mb.doc.Text.createTextCursor()
            cur_text.gotoRange(bm.Anchor,False)
            cur_text.goRight(60,True)
            neuer_Name = cur_text.String.split('\n')[0]
            
            # alte Datei in Helferdatei speichern
            orga_sec_name_alt = self.mb.props['Projekt'].dict_bereiche['ordinal'][zeilenordinal]
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
            
            
    def verbotene_buchstaben_austauschen(self,term):
                            
        verbotene = '<>:"/\\|?*'

        term =  ''.join(c for c in term if c not in verbotene)
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
    
    def schreibe_settings_orga(self):
        if self.mb.debug: log(inspect.stack)
        
        path = os.path.join(self.mb.path_to_extension,"organon_settings.json")

        with open(path, 'w') as outfile:
            json.dump(self.mb.settings_orga, outfile,indent=4, separators=(',', ': '))
            
    def folderpicker(self,filepath=None):
        if self.mb.debug: log(inspect.stack)
        
        folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        if filepath != None:
            folderpicker.setDisplayDirectory(filepath)
        folderpicker.execute()
        
        if folderpicker.Directory == '':
            return None
        
        return uno.fileUrlToSystemPath(folderpicker.getDirectory())
    
    def filepicker(self,filepath=None,filter=None):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        #Filepicker.appendFilter('Organon Project','*.organon')
        #ofilter = Filepicker.getCurrentFilter()
        Filepicker.execute()

        if Filepicker.Files == '':
            return None

        return uno.fileUrlToSystemPath(Filepicker.Files[0])
    
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
                sc = textdoc.getByName(e)
                shortcuts.update({e:sc.Command})
            except:
                shortcuts.update({e:'?'})
                
        elements = glob.ElementNames
        
        for e in elements:
            try:
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
                if not d2.has_key(k):
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
        
        try:
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
                    
                    if name == 'Projekt':
                        Path = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')
                    else:
                        Path = os.path.join(self.mb.pfade['tabs'], name + '.xml')
                    
                    tag1_button = self.mb.props[name].Hauptfeld.getControl(ord_source).getControl('tag1')
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
        except:
            log(inspect.stack,tb())
            
            
            
