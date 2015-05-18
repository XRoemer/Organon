# -*- coding: utf-8 -*-

import unohelper


class Baumansicht():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.dialog = mb.dialog
        self.ctx = mb.ctx
        self.mb = mb
        self.listener_treeview_symbol = TreeView_Symbol_Listener(self.ctx,self.mb)
        self.tag1_listener = Tag1_Listener(self.mb)
        self.tag2_listener = Tag2_Listener(self.mb)
        
    def start(self):
        if self.mb.debug: log(inspect.stack)

        self.mb.props[T.AB].Hauptfeld = self.erzeuge_Feld_Baumansicht(self.mb.dialog)  
        self.erzeuge_Scrollbar()

        
    def erzeuge_Feld_Baumansicht(self,win):
        if self.mb.debug: log(inspect.stack)
        # Das aeussere Hauptfeld wird fuers Scrollen benoetigt. Das innere und eigentliche
        # Hauptfeld scrollt dann in diesem Hauptfeld_aussen

        # Hauptfeld_Aussen
        Attr = (22,KONST.ZEILENHOEHE+2,1000,1800,'Hauptfeld_aussen')    
        PosX,PosY,Width,Height,Name = Attr
        
        control1, model1 = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        
             
        win.addControl('Hauptfeld_aussen',control1)  
        
        # eigentliches Hauptfeld
        Attr = (0,0,1000,10000,)    
        PosX,PosY,Width,Height = Attr
         
        control2, model2 = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  

        model2.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        control1.addControl('Hauptfeld',control2)  
        self.mb.hauptfeld_aussen = control1
        return control2

  
    def erzeuge_Zeile_in_der_Baumansicht(self,eintrag,class_Zeilen_Listener,gliederung,index=0,tab_name = 'Projekt'):
        if self.mb.debug: log(inspect.stack)
        
        # wird in projects aufgerufen
        ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag        

        ##### Aeusserer Container #######
        
        Attr = (2,KONST.ZEILENHOEHE*index,600,20)    
        PosX,PosY,Width,Height = Attr
        
        control, model = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        model.Text = ordinal
        model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
            
        self.mb.props[tab_name].Hauptfeld.addControl(ordinal,control)
        
        sett = self.mb.settings_proj
        sicht_tag1 = int(sett['tag1'])
        sicht_tag2 = int(sett['tag2'])
        sicht_tag3 = int(sett['tag3'])
        
        ### einzelne Elemente #####

        # Icon
        x = 16+int(lvl)*16
        control2, model2 = self.mb.createControl(self.ctx,"ImageControl",x,2,16,16,(),() ) 
        model2.Border = 0 
        control2.addMouseListener(self.listener_treeview_symbol)
        
        if art == 'waste':
            self.mb.props[T.AB].Papierkorb = ordinal
        
        if art == 'prj':
            self.mb.props[T.AB].Projektordner = ordinal
             
        
        if art in ('dir','prj'):

            if zustand == 'auf':
                model2.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16 
            else:
                bild_ordner = KONST.IMG_ORDNER_16

                tree = self.mb.props[T.AB].xml_tree 
                root = tree.getroot()

                ordner_xml = root.find('.//'+ordinal)
                if ordner_xml != None:
                    childs = list(ordner_xml)
                    if len(childs) > 0:
                        bild_ordner = KONST.IMG_ORDNER_VOLL_16

                model2.ImageURL = bild_ordner
            schriftfarbe = KONST.FARBE_SCHRIFT_ORDNER
                
        elif art == 'pg':
            model2.ImageURL = 'private:graphicrepository/res/sx03150.png' 
            schriftfarbe = KONST.FARBE_SCHRIFT_DATEI
            
        elif art == 'waste':
            if zustand == 'zu':
                model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png' 
            else:
                model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_offen.png' 
            sicht_tag1 = sicht_tag2 = sicht_tag3 = False
            
            schriftfarbe = KONST.FARBE_SCHRIFT_ORDNER
        

        control.addControl('icon',control2)                       
        # return ist nur fuer neu angelegte Dokumente nutzbar 
        
                
        x += 18
        
        if sicht_tag1:
            # Tag1 Farbe
            control_tag1, model_tag1 = self.mb.createControl(self.ctx,"ImageControl",x,2,16,16,(),() )  
            model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
            model_tag1.Border = 0
            control_tag1.addMouseListener(self.tag1_listener)
            control.addControl('tag1',control_tag1)
            x += 16
            
        if sicht_tag2:
            # Tag2 BENUTZERDEFINIERT
            control_tag2, model_tag2 = self.mb.createControl(self.ctx,"ImageControl",x,2,16,16,(),() )  
            model_tag2.ImageURL = tag2
            model_tag2.Border = 0
            control_tag2.addMouseListener(self.tag2_listener)
            
            control.addControl('tag2',control_tag2)
            x += 18
        
        if sicht_tag3:
            PosX,PosY,Width,Height = 0,2,16,16
            control_tag3, model_tag3 = self.mb.createControl(self.mb.ctx,"FixedText",x,2,16,16,(),() )
            model_tag3.Label = gliederung[ordinal]
            model_tag3.TextColor = KONST.FARBE_GLIEDERUNG
            breite,hoehe = self.mb.kalkuliere_und_setze_Control(control_tag3,'w')
            control.addControl('tag3',control_tag3)
            x += breite + 4

        
        # Textfeld
        control1, model1 = self.mb.createControl(self.ctx,"Edit",x,0,400,20,(),() )  
        model1.Text = name
        model1.Border = False
        model1.BackgroundColor = KONST.FARBE_HF_HINTERGRUND
        model1.ReadOnly = True
        model1.TextColor = schriftfarbe

        control1.addMouseListener(class_Zeilen_Listener) 
        control1.addMouseMotionListener(class_Zeilen_Listener)
        control1.addFocusListener(class_Zeilen_Listener)

        control.addControl('textfeld',control1)
        
        
        if sicht == 'nein':
            control.Visible = False
        else:              
            index += 1
        return index  
    
    
    def positioniere_icons_in_zeile(self,contr_zeile,tags,gliederung):
        #if self.mb.debug: log(inspect.stack)
        
        try:
            tag1,tag2,tag3 = tags
            
            x = contr_zeile.getControl('icon').PosSize.X +18
            breite = 0
            
            if tag1:
                tag1_contr = contr_zeile.getControl('tag1')
                tag1_contr.setPosSize(x,0,0,0,1)
                x += 16
                breite = 16
                
            if tag2:
                tag2_contr = contr_zeile.getControl('tag2')
                tag2_contr.setPosSize(x,0,0,0,1)
                x += 18
                breite = 18
                
            if tag3:
                tag3_contr = contr_zeile.getControl('tag3')
                
                ordinal = tag3_contr.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
                if ordinal != self.mb.props[T.AB].Papierkorb:
                    tag3_contr.Model.Label = gliederung[ordinal]
                else:
                    tag3_contr.Model.Label = ''
                breite,hoehe = self.mb.kalkuliere_und_setze_Control(tag3_contr)  
                breite += 4              
                tag3_contr.setPosSize(x,0,0,0,1)
                x += breite
            

            text_contr = contr_zeile.getControl('textfeld')
            if text_contr.PosSize.X != x:
                text_contr.setPosSize(x,0,0,0,1)
            
        except:
            log(inspect.stack,tb())
            
     
    def erzeuge_Scrollbar(self,win = None):
        if self.mb.debug: log(inspect.stack)
        
        if win == None:
            win = self.dialog
        nav_cont_aussen = win.getControl('Hauptfeld_aussen')
        control_innen = nav_cont_aussen.getControl('Hauptfeld')
          
        MenuBar = win.getControl('Organon_Menu_Bar')
        MBHoehe = MenuBar.PosSize.value.Height + MenuBar.PosSize.value.Y

        ctrl_innen_posY  = control_innen.PosSize.value.Y
        Y =  ctrl_innen_posY + MBHoehe
        Height = win.PosSize.value.Height - Y - 25

        PosSize = 0,Y,0,Height
        control = self.mb.erzeuge_Scrollbar(win,PosSize,control_innen)
        
        self.mb.scrollbar = control



    def korrigiere_scrollbar(self):
        if self.mb.debug: log(inspect.stack)

        active_tab = self.mb.active_tab_id
        win = self.mb.tabs[active_tab][0]
        SB = win.getControl('ScrollBar')        
        
        hoehe = sorted(list(self.mb.props[T.AB].dict_zeilen_posY))

        sb_hoehe = SB.Size.Height
        baum_hoehe = hoehe[-1] + 20
        if hoehe != []:
            max =  baum_hoehe - (sb_hoehe) 
            if max > 0:
                SB.Visible = True
                SB.LineIncrement = baum_hoehe/sb_hoehe*50
                SB.BlockIncrement = 200
                SB.Maximum =  baum_hoehe  
                SB.VisibleSize = sb_hoehe - 40
                # Bei Mausradnutzung Mausradlistener einschalten
                if self.mb.settings_orga['mausrad']:
                    self.mb.mausrad_an = True
            else:
                SB.Maximum  = 1
                SB.Visible = False
                nav_cont_aussen = win.getControl('Hauptfeld_aussen')
                nav_cont = nav_cont_aussen.getControl('Hauptfeld')
                nav_cont.setPosSize(0, 0,0,0,2)
                # Mausradlistener ausschalten
                self.mb.mausrad_an = False
           
                
    # Nur fuers Debugging
    def finde_falschen_bereich(self):
        if self.mb.debug: log(inspect.stack)
        
        sections = self.mb.doc.TextSections
        names = sections.ElementNames
        if 'Bereich1' in names:
            print('Bereich ist drin')
            pd()

    
    def erzeuge_neue_Zeile(self,ordner_oder_datei):
        if self.mb.debug: log(inspect.stack)
        
        if self.mb.props[T.AB].selektierte_zeile == None:       
            self.mb.nachricht(LANG.ZEILE_AUSWAEHLEN,'infobox')
            return None
        else:
            try:
                StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
                StatusIndicator.start(LANG.ERZEUGE_NEUE_ZEILE,2)
                StatusIndicator.setValue(2)
                
                self.mb.doc.lockControllers()
                
                ord_sel_zeile = self.mb.props[T.AB].selektierte_zeile
                
                # XML TREE
                tree = self.mb.props[T.AB].xml_tree
                root = tree.getroot()
                xml_sel_zeile = root.find('.//'+ord_sel_zeile)
                
                # Props ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3
                ordinal = 'nr'+root.attrib['kommender_Eintrag']
                parent = xml_sel_zeile.attrib['Parent']
                name = ordinal
                lvl = xml_sel_zeile.attrib['Lvl']
                tag1 = 'leer' #xml_sel_zeile.attrib['Tag1']
                tag2 = xml_sel_zeile.attrib['Tag2']
                tag3 = xml_sel_zeile.attrib['Tag3']
                sicht = 'ja' 
                if ordner_oder_datei == 'Ordner':
                    art = 'dir'
                    zustand = 'zu'
                else:
                    art = 'pg'
                    zustand = '-'            
                eintrag = ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 
                
                # neue Zeile / neuer XML Eintrag
                self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)
                
                if self.mb.settings_proj['tag3']:
                    gliederung = self.mb.class_Gliederung.rechne(tree)
                else:
                    gliederung = None

                self.erzeuge_Zeile_in_der_Baumansicht(eintrag,self.mb.class_Zeilen_Listener,gliederung)
                self.mb.class_Sidebar.lege_dict_sb_content_ordinal_an(ordinal)
                            
                # neue Datei / neuen Bereich anlegen           
                # kommender Eintrag wurde in erzeuge_XML_Eintrag schon erhoeht
                nr = int(root.attrib['kommender_Eintrag']) - 1          
                inhalt = ordinal
                
                self.mb.class_Bereiche.erzeuge_neue_Datei2(nr,inhalt)
                self.mb.class_Bereiche.erzeuge_bereich2(nr,sicht)

                # Zeilen anordnen
                source = ordinal
                target = xml_sel_zeile.tag
                action = 'drunter'  
    
                # in zeilen_neu_ordnen wird auch die xml datei geschrieben
                self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(source,target,action)
                self.korrigiere_scrollbar()

                if self.mb.doc.hasControllersLocked(): 
                    self.mb.doc.unlockControllers()

            except:
                log(inspect.stack,tb())
            StatusIndicator.end()
            
            return nr
            
            
    def kontrolle(self):
        if self.mb.debug: log(inspect.stack)
        
        root2 = self.mb.props[T.AB].xml_tree.getroot()
        alle = root2.findall('.//')
        asd = []
        for i in alle:
            asd.append((i.tag,i.attrib['Lvl'],i.attrib['Sicht'],i))


    def leere_Papierkorb(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener)            
            
            if T.AB != 'Projekt':
                ordinal = sorted(list(self.mb.props[T.AB].dict_bereiche['ordinal']))[0]
                zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)        
                self.mb.props[T.AB].selektierte_zeile = zeile.AccessibleContext.AccessibleName
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche()
                
                self.mb.props[T.AB].selektierte_zeile_alt = None
                
            tree = self.mb.props[T.AB].xml_tree
            root = tree.getroot()
            C_XML = self.mb.class_XML
    
            papierkorb_xml = root.find(".//"+self.mb.props[T.AB].Papierkorb)  
            papierkorb_xml.attrib['Zustand'] = 'zu'
            papierkorb_inhalt1 = []
            C_XML.selbstaufruf = False
            C_XML.get_tree_info(papierkorb_xml,papierkorb_inhalt1)        
            selektierter_ist_im_papierkorb = False
    
            # Zeilen im Hauptfeld loeschen
            for eintrag in papierkorb_inhalt1:        
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
    
                if not selektierter_ist_im_papierkorb:
                    if self.mb.props[T.AB].selektierte_zeile == ordinal:
                        selektierter_ist_im_papierkorb = True
                
                if art != 'waste':
                    contr = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)               
                    contr.dispose()
    
            # Eintraege in XML Tree loeschen
            papierkorb_inhalt = list(papierkorb_xml)
            for verworfene in papierkorb_inhalt:
                papierkorb_xml.remove(verworfene)
            
            if T.AB != 'Projekt':
                Path = os.path.join(self.mb.pfade['tabs'] , T.AB +'.xml' )
            else:                    
                Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
            self.mb.tree_write(tree,Path)
            
#             self.mb.selbstruf = True
#             if selektierter_ist_im_papierkorb:
#                 self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche(self.mb.props[T.AB].Papierkorb)
                
            
            self.erneuere_selektierungen(selektierter_ist_im_papierkorb) 
    
            # loesche Bereich(e) und Datei(en)
            if T.AB == 'Projekt':
                self.loesche_Bereiche_und_Dateien(papierkorb_inhalt1,papierkorb_inhalt)
            #self.mb.selbstruf = False

           
            # XML,Ansicht und Dicts neu ordnen
            # vielleicht gibt es noch eine andere Moeglichkeit?
            # hier wird der gesamte Baum fuer jede Neuordnung nochmals durchlaufen
            # -> Performance?
            # Insgesamt wird der Baum 4x durchlaufen
            self.mb.class_Projekt.erzeuge_dict_ordner()
            self.mb.class_Zeilen_Listener.update_dict_zeilen_posY()    # <- dict wird durch tree aufgebaut. schon erneuert?    
            
            self.mb.props[T.AB].Papierkorb_geleert = True
            self.erneuere_dict_bereiche()
            self.korrigiere_scrollbar()
    
        except:
            log(inspect.stack,tb())
                
    def erneuere_selektierungen(self,selektierter_ist_im_papierkorb):
        if self.mb.debug: log(inspect.stack)
        
        zeile_Papierkorb = self.mb.props[T.AB].Hauptfeld.getControl(self.mb.props[T.AB].Papierkorb)
        textfeld_Papierkorb = zeile_Papierkorb.getControl('textfeld')
        icon_Papierkorb = zeile_Papierkorb.getControl('icon')
        icon_Papierkorb.Model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png'
        
        if selektierter_ist_im_papierkorb:
            self.mb.props[T.AB].selektierte_zeile = None
            textfeld_Papierkorb.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE
            self.mb.props[T.AB].selektierte_zeile_alt = textfeld_Papierkorb.Context.AccessibleContext.AccessibleName
            papierkorb_bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][self.mb.props[T.AB].Papierkorb]
            bereich_papierkorb = self.mb.doc.TextSections.getByName(papierkorb_bereichsname)
            
            bereich_papierkorb.IsVisible = True
            self.mb.props['Projekt'].sichtbare_bereiche = [papierkorb_bereichsname]
            
            

    def loesche_Bereiche_und_Dateien(self,papierkorb_inhalt1,papierkorb_inhalt):
        if self.mb.debug: log(inspect.stack)

        for inhalt in papierkorb_inhalt1:
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = inhalt
            
            try:
                if art != 'waste':
                    zahl = ordinal.split('nr')[1]
                    # loesche datei ordinal
                    Path = os.path.join(self.mb.pfade['odts'], 'nr%s.odt' % zahl)
                    os.remove(Path)
                    
                    bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                    
                    # loesche text der Datei im Dokument
                    sections = self.mb.doc.TextSections
                    sec = sections.getByName(bereichsname) 
                    
                    # loesche Sidebareintrag
                    self.mb.class_Sidebar.loesche_dict_sb_content_eintrag(ordinal)
                                        
                    # loesche eventuell vorhandene Kind Bereiche
                    ch_sections = []
                    def get_childs(childs):
                        for i in range (len(childs)):
                            ch_sections.append(childs[i])
                            if childs[i].ChildSections != None:
                                get_childs(childs[i].ChildSections)

                    get_childs(sec.ChildSections)
                        
                    for child_sec in ch_sections:
                        child_sec.dispose()
                    
                    trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
                    
                    # Loesche Bereich
                    sec.IsVisible = True          
                    textSectionCursor = self.mb.doc.Text.createTextCursorByRange(sec.Anchor)
                    textSectionCursor.setString('')
                    sec.dispose()
                    
                    # sec_helfer wieder auf invisible setzen, damit er nicht geloescht wird
                    self.mb.sec_helfer.IsVisible = False
                                        
                    # Trenner loeschen
                    if trenner_name in self.mb.doc.TextSections.ElementNames:
                        trenner = self.mb.doc.TextSections.getByName(trenner_name)
                        textSectionCursor.gotoRange(trenner.Anchor,True)
                        trenner.dispose()
                        textSectionCursor.setString('')
                  
                    while textSectionCursor.TextSection == None:
                        textSectionCursor.goLeft(1,True)
                        
                    textSectionCursor.setString('')
            except:
                log(inspect.stack,tb())
        
        
    def erneuere_dict_bereiche(self):
        if self.mb.debug: log(inspect.stack)
        
        sections = self.mb.doc.TextSections

        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot() 
        alle_Zeilen = root.findall('.//')
                
        # Fuer die Neuordnung des dict_bereiche
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        
        for i in range (len(alle_Zeilen)):
            sec = sections.getByName('OrganonSec'+str(i))
            ordinal = alle_Zeilen[i].tag
            
            #Path = 'file:///' + self.mb.pfade['odts'] + '/%s.odt' % ordinal   
            Path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)  
            
            Bereichsname_dict.update({'OrganonSec'+str(i):Path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(i)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(i):ordinal})
             
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.props[T.AB].dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        
        
from com.sun.star.style.ParagraphAdjust import CENTER 
from com.sun.star.awt import XMouseListener,XMouseMotionListener,XFocusListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT
from com.sun.star.awt.MouseButton import RIGHT as MB_RIGHT

class Zeilen_Listener (unohelper.Base, XMouseListener,XMouseMotionListener,XFocusListener):
    def __init__(self,ctx,mb):
        if mb.debug: log(inspect.stack)
        self.pfeil = False
        self.ctx = ctx
        self.mb = mb
        # fuer den erzeugten Pfeil
        self.first_time = True
        # beschreibt die Art der Aktion
        self.einfuegen = None
        # fuer das gezogene Textfeld
        self.colored = False        
        self.edit_text = False

        self.mb.props[T.AB].selektierte_zeile_alt = None
        self.dragged = False
        self.menu_geklickt = False
        
        try:
            self.SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            self.SFLink2 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            
        except:
            log(inspect.stack)
            
        
        # nur fuers logging
        self.log_selbstruf = False
              
    def mouseMoved(self,ev):  
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener)
        
        # Mauslistener starten
        try:
            self.mb.mausrad_an = True
            self.mb.class_Mausrad.starte_mausrad(called_from_treeview=True)
        except:
            log(inspect.stack,tb())

        try:
            # die gesamte Zeile, control (ordinal)
            self.mb.props[T.AB].selektierte_zeile = ev.Source.Context.AccessibleContext.AccessibleName
            # control 'textfeld'   
            zeile = ev.Source

            # selektierte Zeile einfaerben, ehem. sel. Zeile zuruecksetzen
            zeile.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE 
            if self.mb.props[T.AB].selektierte_zeile_alt != None:
                if self.mb.props[T.AB].selektierte_zeile != self.mb.props[T.AB].selektierte_zeile_alt:
                    ctrl = self.mb.props[T.AB].Hauptfeld.getControl(self.mb.props[T.AB].selektierte_zeile_alt).getControl('textfeld')
                    ctrl.Model.BackgroundColor = KONST.FARBE_HF_HINTERGRUND

            # bei bearbeitetem Bereich: speichern  
            if self.mb.props[T.AB].selektierte_zeile_alt != None: 
                if self.mb.props[T.AB].tastatureingabe == True:
                    bereichsname_zeile_alt = self.mb.props[T.AB].dict_bereiche['ordinal'][self.mb.props[T.AB].selektierte_zeile_alt ]
                    bereich_zeile_alt = self.mb.doc.TextSections.getByName(bereichsname_zeile_alt)
    
                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(bereich_zeile_alt.FileLink.FileURL,bereichsname_zeile_alt)
                    self.mb.props[T.AB].tastatureingabe = False
                    
            self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)      
            self.mb.props[T.AB].selektierte_zeile_alt = self.mb.props[T.AB].selektierte_zeile
            self.mb.class_Sidebar.passe_sb_an(zeile)
            
            # Bei Doppelclick Zeileneintrag bearbeiten
            if ev.Buttons == MB_LEFT:   
                if ev.ClickCount == 2: 
                    
                    # Projektordner von der Umbenennung ausnehmen
                    root = self.mb.props[T.AB].xml_tree.getroot()
                    zeile_xml = root.find('.//' + self.mb.props[T.AB].selektierte_zeile_alt )
                    art = zeile_xml.attrib['Art']
                    
                    if art == 'prj':
                        self.mb.nachricht('Der Projektordner kann nicht umbenannt werden',"infobox")
                        return False
                    else:
                        zeile.Model.ReadOnly = False 
                        zeile.Model.BackgroundColor = KONST.FARBE_EDITIERTE_ZEILE
                        self.edit_text = True   
                        return False
            
            return False
        except:
            log(inspect.stack,tb())

    def mouseReleased(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        if self.menu_geklickt:
            self.menu_geklickt = False
            return
        
        # Sichtbarkeit der Bereiche umschalten
        if self.dragged == False:  
            self.schalte_sichtbarkeit_der_Bereiche()
               
        else:
            self.dragged = False

        # wenn maus gezogen und pfeil erzeugt wurde
        if self.pfeil == True:
            self.pfeil = False
            self.first_time = True
            
            if self.colored:
                # Farbe gezogenes Textfeld
                ev.Source.Model.BackgroundColor = self.old_color#16777215 #weiss
                self.colored = False 
            
            
            s = self.zielordner.getControl('symbol')

            if s != None:
                source = (ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName)#,ev.Source.Context.AccessibleContext.Location.Y)
                target = self.zielordner.AccessibleContext.AccessibleName
                action = self.einfuegen                          
                s.dispose()
                if source != target:
                    self.zeilen_neu_ordnen(source,target,action)   

        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener)
            
        zeilenordinal =  self.mb.props[T.AB].selektierte_zeile
        bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][zeilenordinal]
        bereich = self.mb.doc.TextSections.getByName(bereichsname)

        self.mb.viewcursor.gotoRange(bereich.Anchor,False)
        self.mb.viewcursor.gotoStart(False)
            
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)
        
        self.mb.loesche_undo_Aktionen()
                           
    def focusGained(self,ev):
        return False  
    
    def focusLost(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        # Listener fuers Mausrad beenden
        self.mb.mausrad_an = False
        
        # Bearbeitung des Zeileneintrags wieder ausschalten
        if self.edit_text == True:

            ev.Source.Model.ReadOnly = True 
            self.edit_text = False
            # neuen Text in die xml Dateien eintragen
            zeile_ord = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
            
            for tab_name in self.mb.props:
                
                tree = self.mb.props[tab_name].xml_tree
                root = tree.getroot() 
                
                xml_elem = root.find('.//'+zeile_ord)
                
                if xml_elem != None:
                
                    xml_elem.attrib['Name'] = ev.Source.Text
                    
                    if tab_name == 'Projekt':
                        Path1 = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
                    else:
                        Path1 = os.path.join(self.mb.pfade['tabs'],tab_name + '.xml')
                        
                    self.mb.tree_write(tree,Path1)
                    
                    # Textfeld in allen Tabs anpassen
                    textfeld = self.mb.props[tab_name].Hauptfeld.getControl(zeile_ord).getControl('textfeld')
                    textfeld.Model.Text = ev.Source.Text
            
        return False

    def mouseDragged(self,ev):
#         if self.mb.debug: log(inspect.stack)
        
        # Papierkorb darf nicht verschoben werden
        ordinal = ev.Source.Context.AccessibleContext.AccessibleName
        if ordinal in (self.mb.props[T.AB].Papierkorb,self.mb.props[T.AB].Projektordner):
            self.mb.nachricht(LANG.NICHT_VERSCHIEBBAR,"infobox")
            return

        self.dragged = True
        if self.edit_text == False:
            self.pfeil = True
            X = ev.Source.AccessibleContext.AccessibleParent.PosSize.X + ev.X
            # -1: da icon 24x24, Container aber nur 22 hoch
            Y = ev.Source.AccessibleContext.AccessibleParent.PosSize.Y -1 + ev.Y
            
            self.erzeuge_pfeil(X,Y,ev)
            
            # Textfeld waehrend des Ziehens einfaerben
            if not self.colored:
                self.colored = True
                self.old_color = ev.Source.Model.BackgroundColor
                ev.Source.Model.BackgroundColor = KONST.FARBE_GEZOGENE_ZEILE 
                           
    def erzeuge_pfeil(self,X,Y,ev):
#         if self.mb.debug: log(inspect.stack)
        
        try:        
            if self.pfeil == True:
                if self.first_time == True:
                    self.zielordner,info = self.berechne_pos(Y)
                    zeileY,art,lvl,nachfolger,ord_erster,ordinal = info

                    lvl_nf,art_nf,ord_nf = nachfolger
                    
                    self.first_time = False
                    self.Y = Y

                    # zur Sicherheit - wenn keine Abfrage greift, wird ein falsches Bild dargestellt
                    ImageURL = KONST.IMAGE_GESCHEITERT

                    sourceZeile = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext
                    sourceName = sourceZeile.AccessibleName
                    sourceYPos = sourceZeile.Location.Y

                    sourceArt = self.mb.props[T.AB].dict_zeilen_posY[sourceYPos][4]

                    if sourceArt in ('dir','prj'):
                        subelements = self.mb.props[T.AB].dict_ordner[sourceName]
                    else:
                        subelements = 'quatsch'
                    
                    if sourceArt in ('dir','prj') and ordinal in subelements:
                        ImageURL = KONST.IMAGE_GESCHEITERT
                        self.einfuegen = 'gescheitert'  

                    # Wenn Erster
                    elif ordinal == ord_erster:
                        if zeileY < 10:
                            ImageURL = KONST.IMAGE_PFEIL_HOCH
                            self.einfuegen = 'drueber'
                        elif zeileY < 16:
                            ImageURL = KONST.IMAGE_PFEIL_RECHTS
                            self.einfuegen = 'inOrdnerEinfuegen'
                        else:
                            ImageURL = KONST.IMAGE_PFEIL_RUNTER
                            self.einfuegen = 'drunter'

                    # Wenn 'pg'
                    elif art == 'pg':
                        if lvl_nf < lvl:
                            if zeileY > 15:
                                lvl = lvl_nf
                                self.einfuegen = 'vorNachfolger',ord_nf                               
                            else:
                                self.einfuegen = 'drunter'                                
                        else:
                            self.einfuegen = 'drunter'
                        ImageURL = KONST.IMAGE_PFEIL_RUNTER
                     

                    # Wenn 'dir'
                    elif (art in ('dir','prj')) :
                        if lvl_nf < lvl:
                            if zeileY > 15:
                                lvl = lvl_nf
                                ImageURL = KONST.IMAGE_PFEIL_RUNTER
                                self.einfuegen = 'vorNachfolger',ord_nf                               
                            else:
                                ImageURL = KONST.IMAGE_PFEIL_RECHTS
                                self.einfuegen = 'inOrdnerEinfuegen'                                                   
                        elif art_nf in ('dir','prj'):
                            if zeileY > 15:
                                ImageURL = KONST.IMAGE_PFEIL_RUNTER
                                self.einfuegen = 'drunter'   
                            else:
                                ImageURL = KONST.IMAGE_PFEIL_RECHTS
                                self.einfuegen = 'inOrdnerEinfuegen'                                                             
                        else:
                            ImageURL = KONST.IMAGE_PFEIL_RECHTS
                            self.einfuegen = 'inOrdnerEinfuegen' 
                                                                            
                    # Wenn 'waste'                  
                    elif (art == 'waste') :
                        ImageURL = KONST.IMAGE_PFEIL_RECHTS
                        self.einfuegen = 'inPapierkorbEinfuegen'
                   

                    # Pfeil erstellen
                    Color__Container = 10202
                    Attr = (int(lvl)*16,3,16,16,'eintrag23', Color__Container)    
                    PosX,PosY,Width,Height,Name,Color = Attr
                     
                    control, model = self.mb.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )      
                    model.Border = False                    
                    model.ImageURL = ImageURL
 
                    self.zielordner.addControl("symbol",control)
 
                else:      
                                    
                    if not (self.Y -5 < Y < self.Y+5):   
                        zielordner,info = self.berechne_pos(Y)
                        zeileY,art,lvl,nachfolger,ord_erster,ordinal = info
                        lvl_nf,art_nf,ord_nf = nachfolger
                       
                        if zielordner.AccessibleContext.AccessibleName == ord_erster:

                            self.zielordner.getControl("symbol").dispose()
                            self.first_time = True
                            self.Y = Y 

                        
                        # Code zur Beruhigung der Anzeige - weniger dispose()
                        # self.zielordner wird in first_time == True gesetzt
                        if zielordner == self.zielordner:
                            
                            if art == 'pg':
                                if lvl_nf < lvl:
                                    self.zielordner.getControl("symbol").dispose()
                                    self.first_time = True
                                    self.Y = Y  
                                else:                                    
                                    pass
                            if art in ('dir','prj'):
                                if lvl_nf < lvl:
                                    self.zielordner.getControl("symbol").dispose()
                                    self.first_time = True
                                    self.Y = Y 
                                elif art_nf in ('dir','prj'): 
                                    self.zielordner.getControl("symbol").dispose()
                                    self.first_time = True
                                    self.Y = Y 
                                else:                                    
                                    pass
                            else: 
                                pass                    
                        else:      
                            pfeil = self.zielordner.getControl("symbol")  
                            if pfeil != None:
                                pfeil.dispose()
                                self.first_time = True
                                self.Y = Y  
                    
        except:
            # Der Fehler braucht nicht geloggt zu werden, da er 
            # staendig und folgenlos auftritt
            pass #print(tb())
                
    def berechne_pos(self,Y):
#         if self.mb.debug: log(inspect.stack)
        
        y = (math_floor(Y/KONST.ZEILENHOEHE)) 
        
        # Position des Papierkorbs 
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()   
        #ord_papierkorb = root.find(".//*[@Art='waste']").tag
        ord_papierkorb = self.mb.props['Projekt'].Papierkorb
        zeile_papierkorb = self.mb.props[T.AB].Hauptfeld.getControl(ord_papierkorb)
        pos_papierkorb = zeile_papierkorb.PosSize.Y/KONST.ZEILENHOEHE

        # abfangen: wenn Mauspos ueber Hauptfeld oder unter Papierkorb hinausgeht
        if y < 0:
            y = 0
        elif y > pos_papierkorb:
            y = pos_papierkorb

        posY = (y * KONST.ZEILENHOEHE)       
        posY_nachf = ((y+1) * KONST.ZEILENHOEHE)
        
        ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = self.mb.props[T.AB].dict_zeilen_posY[posY]
        ord_erster = self.mb.props[T.AB].dict_zeilen_posY[0][0]
        
        # try/except um nur eine Abfrage an das Dict zu stellen
        try:
            nf = self.mb.props[T.AB].dict_zeilen_posY[posY_nachf]
            lvl_nf = nf[3]
            art_nf = nf[4]
            ord_nf = nf[0]
            nachfolger = (lvl_nf,art_nf,ord_nf)
        except:
            # Papierkorb: Nachfolger ist immer auf hoeherem lvl
            nachfolger = (100,'keiner','sdfsf')
        
        zeileY = Y - posY
        info = (zeileY,art,lvl,nachfolger,ord_erster,ordinal)
        zielordner = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)
         
        return (zielordner,info)
 
           
    def zeilen_neu_ordnen(self,source,target,action):
        if self.mb.debug: log(inspect.stack)

        if action != 'gescheitert':

            
            ok = self.wird_datei_in_papierkorb_verschoben(source,target)
            if not ok:
                return

            if 'vorNachfolger' in action:
                # wenn der Nachfolger gleich der Quelle ist, nicht verschieben
                if source == action[1]:
                    return
 
            eintraege = self.xml_neu_ordnen(source,target,action)
            self.posY_in_tv_anpassen(eintraege)
            
            # dict_ordner updaten
            self.mb.class_Projekt.erzeuge_dict_ordner()

            # Bereiche neu verlinken
            sections = self.mb.doc.TextSections
            self.verlinke_Bereiche(sections)
            # Sichtbarkeit der Bereiche umschalten
            self.schalte_sichtbarkeit_der_Bereiche()      

            if T.AB == 'Projekt':   
                Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
            else:
                Path = os.path.join(self.mb.pfade['tabs'] , T.AB + '.xml' )
            self.mb.tree_write(self.mb.props[T.AB].xml_tree,Path)
            
                
    def wird_datei_in_papierkorb_verschoben(self,source,target):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if T.AB != 'Projekt':
                return True
            
            if target != self.mb.props['Projekt'].Papierkorb:
                return True
                    
            if source in self.mb.props['Projekt'].dict_ordner:
                ordinale = self.mb.props['Projekt'].dict_ordner[source]
            else:
                ordinale = source,
            
            
            for ordn in ordinale:
                for tab in self.mb.props:
                    if tab != 'Projekt':
                        if ordn in self.mb.props[tab].dict_bereiche['ordinal']:
                            
                            root = self.mb.props[tab].xml_tree.getroot()
                            dateiname = root.find('.//'+ordn).attrib['Name']
                                                         
                            self.mb.nachricht(LANG.KANN_NICHT_VERSCHOBEN_WERDEN %(dateiname,tab),"warningbox")
                            return False
        except:
            log(inspect.stack,tb())
                    
        return True
        
    def posY_in_tv_anpassen(self,eintraege): 
        if self.mb.debug: log(inspect.stack)
        
        # ordnen des dict_zeilen_posY
        self.mb.props[T.AB].dict_zeilen_posY = {}
        index = 0
        
        for eintrag in eintraege:
            
            ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            
            if sicht == 'ja':
                # dict_zeilen_posY updaten            
                self.mb.props[T.AB].dict_zeilen_posY.update({KONST.ZEILENHOEHE*index:eintrag})  
                                            
                # Y_Wert sichtbarer Eintraege setzen
                contr_zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)
                                
                # der X-Wert ALLER Eintraege wird in xml_m neu gesetzt
                y = KONST.ZEILENHOEHE*index
                if contr_zeile.PosSize.Y != y:
                    contr_zeile.setPosSize(0,y,0,0,2)# 2: Flag fuer: nur Y Wert aendern
                          
                index += 1  

        
    def xml_neu_ordnen(self,source,target,action):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        C_XML = self.mb.class_XML
        target_xml = root.find('.//'+target)
        
        # bei Verschieben in einen geschlossenen Ordner
        if 'inOrdnerEinfuegen' in action and target_xml.attrib['Zustand'] == 'zu':
            target_xml.attrib['Zustand'] = 'auf'
            self.schalte_sichtbarkeit_des_hf(target,target_xml,'zu')
            
            tar_cont = self.mb.props[T.AB].Hauptfeld.getControl(target)
            tar = tar_cont.getControl('icon')
            tar.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
        
        
        if 'drunter' in action:
            C_XML.drunter_einfuegen(source,target)
        elif 'inOrdnerEinfuegen' in action:
            C_XML.in_Ordner_einfuegen(source,target)
        elif 'vorNachfolger' in action:
            C_XML.vor_Nachfolger_einfuegen(source,nachfolger = action[1])
        elif 'inPapierkorbEinfuegen' in action:
            C_XML.in_Papierkorb_einfuegen(source,target)
        elif 'drueber' in action:
            C_XML.drueber_einfuegen(source,target)
        
                  
        # bei Verschieben in den Papierkorb
        if 'inPapierkorbEinfuegen' in action and target_xml.attrib['Zustand'] == 'zu':
            # dict_ordner muss vorher erneuert werden
            self.mb.class_Projekt.erzeuge_dict_ordner()               
            self.schalte_sichtbarkeit_des_hf(target,target_xml,'auf')
            
        
        eintraege = []
        # selbstaufruf nur fuer den debug
        C_XML.selbstaufruf = False
        C_XML.get_tree_info(root,eintraege)
        
        return eintraege
        
   
    def schalte_sichtbarkeit_des_hf(self,selbst,selbst_xml,zustand,zeige_projektordner = False):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
         
        def durchlaufe_baum(dir):  
            for child in dir:
                #print(child.tag,child.attrib['Name'])
                if child.attrib['Art'] in ('dir','waste','prj'):  
                    #print('wird sichtbar 1', child.attrib['Name'],child.attrib['Zustand'],child.attrib['Parent'])  
                    tar = self.mb.props[T.AB].Hauptfeld.getControl(child.tag)
                    tar.Visible = True   
                    
                    elem = root.find('.//'+ child.tag)
                    elem.attrib['Sicht'] = 'ja'
                           
                    if child.attrib['Zustand'] == 'auf':                        
                        durchlaufe_baum(child)
                else:
                    if dir.attrib['Zustand'] == 'auf':
                        #print('wird sichtbar 2', child.attrib['Name'],'pg')
                        tar = self.mb.props[T.AB].Hauptfeld.getControl(child.tag)
                        tar.Visible = True           
                        
                        elem = root.find('.//'+ child.tag)
                        elem.attrib['Sicht'] = 'ja'   
        
        if zeige_projektordner:
            selbst_xml.attrib['Zustand'] = 'auf'
            durchlaufe_baum(selbst_xml)
        
        if zustand == 'auf':
            dir = self.mb.props[T.AB].dict_ordner[selbst]
            
            for child in dir:
                if child != selbst:
                    tar = self.mb.props[T.AB].Hauptfeld.getControl(child)
                    tar.Visible = False
                    elem = root.find('.//'+ child)
                    elem.attrib['Sicht'] = 'nein'
                    #print('wird unsichtbar',elem.attrib['Name'],elem.attrib['Sicht'])
        else:   
            durchlaufe_baum(selbst_xml)
            
        self.positioniere_elemente_im_baum_neu()          
        self.update_dict_zeilen_posY() 
              
        
    def verlinke_Bereiche(self,sections):
        if self.mb.debug: log(inspect.stack)
        
        # langsame und sichere Loesung: es werden alle Bereiche neu verlinkt, 
        # nicht nur die verschobenen
        # NEU UND SCHNELLER: Es werden nur noch sichtbare Bereiche neu verlinkt in: self.verlinke_Sektion

        # Der VC Listener wird von IsVisible ausgeloest,
        # daher wird er vorher ab- und hinterher wieder angeschaltet
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 
                
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot() 
        alle_Zeilen = root.findall('.//')
        
        # Fuer die Neuordnung des dict_bereiche
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}

        try:
            # Blendet den Papierkorb aus, wenn neuer Bereich eingefuegt wurde
            letzte_zeile = sections.getByIndex(sections.Count - 1)
            letzte_zeile.IsVisible = False   
        except:
            log(inspect.stack,tb())
                
        for i in range (len(alle_Zeilen)):
            ordinal = alle_Zeilen[i].tag
            Path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)  
            
            Bereichsname_dict.update({'OrganonSec'+str(i):Path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(i)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(i):ordinal})
            
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.props[T.AB].dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)         
        
    def get_links(self):
        if self.mb.debug: log(inspect.stack)
        
        dict = {}
        sections = self.mb.doc.TextSections
        names = sections.ElementNames
         
        for name in names:
            if 'OrganonSec' in name:
                sec = sections.getByName(name)

                if sec.FileLink.FileURL != '':
                    dict.update({sec.FileLink.FileURL:sec.Name})
         
        return dict
    
    
    def get_sections(self):
        
        dict = {}
        sections = self.mb.doc.TextSections
        names = sections.ElementNames

        for name in names:
            if 'OrganonSec' in name:
                sec = sections.getByName(name)
                dict.update({name:sec})
        return dict
    
    
    def get_links2(self,sections_uno):
        if self.mb.debug: log(inspect.stack)
 
        dict = {}
        sections = self.mb.doc.TextSections
        names = sections.ElementNames

        for name in sections_uno:
            sec = sections_uno[name]
            dict.update({sec.FileLink.FileURL:sec.Name})

        return dict       
    
    def finde_eltern_bereich(self,sec):
        if self.mb.debug: log(inspect.stack)
        
        while sec.ParentSection != None:
            sec = sec.ParentSection
        return sec 
        
    def schalte_sichtbarkeit_der_Bereiche(self,zeilenordinal = None):
        if self.mb.debug: log(inspect.stack)
        # sec_trenner.Anchor.BreakType = 'NONE'
        #self.mb.use_UM_Listener = False
        try:
            # Der VC Listener wird von IsVisible ausgeloest,
            # daher wird er vorher ab- und hinterher wieder angeschaltet
            self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 
            
            if zeilenordinal == None:
                if self.mb.props[T.AB].selektierte_zeile != None:
                    zeilenordinal = self.mb.props[T.AB].selektierte_zeile
                else:
                    return

            # Viewcursor an den Anfang setzen, damit beim Umschalten der Bereiche
            # nur die Helfer-sec2 gezeigt wird
            sec = self.finde_eltern_bereich(self.mb.viewcursor.TextSection)
            self.mb.viewcursor.gotoRange(sec.Anchor.Start,False)
                        
            self.mb.sec_helfer2.IsVisible = True
            self.mb.viewcursor.gotoStart(False)
            
            sections_uno = self.get_sections()
               
            # Ordner
            if zeilenordinal in self.mb.props[T.AB].dict_ordner:
                zeilen_in_ordner_ordinal = self.mb.props[T.AB].dict_ordner[zeilenordinal]
                #print(zeilen_in_ordner_ordinal)
                
                hinzugefuegte = []
                
                # alle Zeilen im Ordner einblenden
                for z in zeilen_in_ordner_ordinal:
                    
                    ordnereintrag_name = self.mb.props[T.AB].dict_bereiche['ordinal'][z]
                    #z_in_ordner = self.mb.doc.TextSections.getByName(ordnereintrag_name)
                    z_in_ordner = sections_uno[ordnereintrag_name]
                    hinzugefuegte.append(z_in_ordner.Name)
                    #pd()
                    self.verlinke_Sektion(ordnereintrag_name,z_in_ordner,sections_uno)

                    #print('blende ein:',z_in_ordner.Name,'##################')
                    z_in_ordner.IsVisible = True
                    self.mache_Kind_Bereiche_sichtbar(z_in_ordner)
                    
                    # Wenn mehr als nur ein geschlossener Ordner zu sehen ist
                    if len(zeilen_in_ordner_ordinal) > 1:
                        # Wenn die Zeile im Papierkorb ist, darf der letzte trenner nicht gezeigt werden, da
                        # er nicht existiert
                        if zeilenordinal in self.mb.props[T.AB].dict_ordner[self.mb.props[T.AB].Papierkorb]:
                            if zeilen_in_ordner_ordinal.index(z) != len(zeilen_in_ordner_ordinal)-1:
                                self.zeige_Trenner(z_in_ordner,zeilenordinal)
                        else:
                            self.zeige_Trenner(z_in_ordner,zeilenordinal)
                    # Wenn der Ordner keine Kinder hat
                    else:
                        self.entferne_Trenner(z_in_ordner)
                       
                # uebrige noch sichtbare ausblenden
                sichtbare_bereiche = self.mb.props['Projekt'].sichtbare_bereiche

                for bereich in (sichtbare_bereiche): 

                    if bereich not in hinzugefuegte:                            
                        #sec = self.mb.doc.TextSections.getByName(bereich)
                        sec = sections_uno[bereich]
                        sec.IsVisible = False 

                        #print('blende aus:',sec.Name,bereich_ord,'---------------')
                        self.entferne_Trenner(sec)
                                       
                # sichtbare Bereiche wieder in Prop eintragen
                self.mb.props['Projekt'].sichtbare_bereiche = []  
                for b in zeilen_in_ordner_ordinal:
                    self.mb.props['Projekt'].sichtbare_bereiche.append(self.mb.props[T.AB].dict_bereiche['ordinal'][b]) 
                    
                                   
            else:
            # Seiten 
                #self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 
                selekt_bereich_name = self.mb.props[T.AB].dict_bereiche['ordinal'][zeilenordinal]
                selekt_bereich = self.mb.doc.TextSections.getByName(selekt_bereich_name)
                
                self.mb.sec_helfer.IsVisible = True
                selekt_bereich.IsVisible = False
                
                self.verlinke_Sektion(selekt_bereich_name,selekt_bereich,sections_uno)
                
                self.mache_Kind_Bereiche_sichtbar(selekt_bereich)
                self.entferne_Trenner(selekt_bereich)
                
                for bereich in (self.mb.props['Projekt'].sichtbare_bereiche):
                    if bereich != selekt_bereich_name:
                        section = self.mb.doc.TextSections.getByName(bereich)  
                        section.IsVisible = False
                        self.entferne_Trenner(section)

                selekt_bereich.IsVisible = True 
                self.mb.sec_helfer.IsVisible = False
                self.mb.props['Projekt'].sichtbare_bereiche = [selekt_bereich_name]
                
                # HACK ! Manchmal wird die Ansicht des Dokumentenfensters nicht vollstaendig erneuert.
                # Scheint mit Formatierungen des Textes zusammenzuhaengen.
                # Daher wird die Anzeige von Textbereichen gesetzt, um die 
                # Ansicht zu aktualisieren.
                oBool = self.mb.current_Contr.ViewSettings.ShowTextBoundaries
                self.mb.current_Contr.ViewSettings.ShowTextBoundaries = oBool
            
            self.mb.sec_helfer2.IsVisible = False
            self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener) 
            #self.mb.use_UM_Listener = True
        except:
            log(inspect.stack,tb())
            
    
    def schalte_sichtbarkeit_des_ersten_Bereichs(self):
        if self.mb.debug: log(inspect.stack)
   
        # Der VC Listener wird von IsVisible ausgeloest,
        # daher wird er vorher ab- und hinterher wieder angeschaltet
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 

        zeilenordinal =  self.mb.props[T.AB].selektierte_zeile
        bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][zeilenordinal]
        
        sections_uno = self.get_sections()
        
        # Ordner
        if zeilenordinal in self.mb.props[T.AB].dict_ordner:
            zeilen_in_ordner_ordinal = self.mb.props[T.AB].dict_ordner[zeilenordinal][0],
            
            # alle Zeilen im Ordner einblenden
            for z in zeilen_in_ordner_ordinal:
                ordnereintrag_name = self.mb.props[T.AB].dict_bereiche['ordinal'][z]
                z_in_ordner = self.mb.doc.TextSections.getByName(ordnereintrag_name)
                self.verlinke_Sektion(ordnereintrag_name,z_in_ordner,sections_uno)    

            # uebrige noch sichtbare ausblenden
            for bereich in self.mb.props['Projekt'].sichtbare_bereiche:                        
                bereich_ord = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][bereich]                     
                if bereich_ord not in zeilen_in_ordner_ordinal:                            
                    sec = self.mb.doc.TextSections.getByName(bereich)
                    sec.IsVisible = False 
                    self.entferne_Trenner(sec)      
                               
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener) 
    
    def verlinke_Sektion(self,name,bereich,sections_uno):
        if self.mb.debug: log(inspect.stack)
        
        dict_filelinks = self.get_links2(sections_uno)

        try:
            url_in_dict = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][name])
            url_sec = bereich.FileLink.FileURL
            
            if url_in_dict != url_sec:
            
                self.SFLink.FileURL = url_in_dict

                if url_in_dict in dict_filelinks:
                    sec = sections_uno[dict_filelinks[url_in_dict]]
                    
                    empty_file = os.path.join(self.mb.pfade['odts'],'empty_file.odt') 
                    self.SFLink2.FileURL = uno.systemPathToFileUrl(empty_file)
                    sec.setPropertyValue('FileLink',self.SFLink2)
                    #print('link neu gesetzt ##############',sec.FileLink.FileURL)
                 
                if url_in_dict != bereich.FileLink.FileURL:
                    bereich.setPropertyValue('FileLink',self.SFLink)

        except:
            log(inspect.stack,tb())
    

    ## NUR ZU TESTZWECKEN ##
    def pruefe_vorkommen(self,url_in_dict,child):  
        dict_filelinks = self.get_links()
        if url_in_dict in dict_filelinks:
            print('url immer noch vorhanden') 
            pd()   
        else:
            print('link nicht vorhanden')
            
        sections = self.mb.doc.TextSections
        names = sections.ElementNames
        
        if child in names:
            sec = sections.getByName(child)
            print(child,'ist in den Sections: als Kind von:', sec.ParentSection.Name)
            
            
            for name in names:
                if 'OrganonSec' in name:
                    sec = sections.getByName(name)
                    time.sleep(0.04)
                    if sec.IsVisible:
                        print(sec.Name,'ist sichtbar')
                    else:
                        print(sec.Name,'ist unsichtbar')
            
        else:
            print(child,'ist nicht in den Sections')
            
    def mache_Kind_Bereiche_sichtbar(self,sec):
        # self.log_selbstruf: um mache_Kind_Bereiche_sichtbar nur
        # 1x zu loggen
        if not self.log_selbstruf:
            if self.mb.debug: log(inspect.stack)
        
        self.log_selbstruf = True

        for kind in sec.ChildSections:
            kind.IsVisible = True
            #print('blende Kind ein:',kind.Name)
            if len(kind.ChildSections) > 0:
                self.mache_Kind_Bereiche_sichtbar(kind)
        
        self.log_selbstruf = False
    
    def zeige_Trenner(self,sec,source_ordinal):  
        if self.mb.debug: log(inspect.stack)

        try:
            trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
            sec_name =  sec.Name

            sec_ordinal = os.path.basename(sec.FileLink.FileURL).split('.')[0]
            sec_nachfolger_name = 'OrganonSec' + str(int(sec.Name.split('OrganonSec')[1])+1)

            if trenner_name in self.mb.doc.TextSections.ElementNames:
                trennerSec = self.mb.doc.TextSections.getByName(trenner_name)
                trennerSec.IsVisible = True
                if len(sec.ChildSections) != 0:
                    trennerSec.BackColor = sec.ChildSections[0].BackColor
                    self.setze_Trenner_Formatierung(trennerSec,sec_nachfolger_name,source_ordinal,sec_ordinal)
                return trennerSec
            
            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                     
            newSection.setName(trenner_name)
            newSection.IsProtected = True
            if len(sec.ChildSections) != 0:
                newSection.BackColor = sec.ChildSections[0].BackColor
    
            
            sec_nachfolger = self.mb.doc.TextSections.getByName(sec_nachfolger_name)
    
            cur = sec_nachfolger.Anchor.Text.createTextCursor()
            cur.gotoRange(sec_nachfolger.Anchor,False)
            cur.gotoRange(cur.TextSection.Anchor.Start,False)        
            
            sec.Anchor.End.Text.insertTextContent(cur, newSection, False)
            
            cur.goLeft(1,False)
            cur.ParaStyleName = 'Standard'
            #cur.PageDescName = ""
            #cur.Text.insertString(cur,'a',False)
            cur.PageDescName = ""
            
            self.setze_Trenner_Formatierung(newSection,sec_nachfolger_name,source_ordinal,sec_ordinal)
            
            return newSection
        except:
            log(inspect.stack,tb())
            
            
    def setze_Trenner_Formatierung(self,sec_trenner,sec_nachfolger_name,source_ordinal,sec_ordinal):  
        if self.mb.debug: log(inspect.stack)

        nachfolger_ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][sec_nachfolger_name]

        if nachfolger_ordinal == self.mb.props[T.AB].Papierkorb:
            sec_trenner.setPropertyValue('BackGraphicURL','')
            sec_trenner.Anchor.setString('')
            
            return
                
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        nachfolger_xml = root.find('.//'+nachfolger_ordinal)
        source_xml = root.find('.//'+source_ordinal)
        sec_xml = root.find('.//'+sec_ordinal)
        
        
        def get_dateiname():
            lvl = sec_xml.attrib['Lvl']
            list_root = root.findall('.//')
            anz = len(list_root)
            ind = list_root.index(sec_xml)+1

            i = range(ind,anz)
            if anz == ind:
                return
            for i in range(ind,anz):
                if list_root[i].attrib['Lvl'] <= lvl:
                    name = list_root[i].attrib['Name']
                    return name
            return '?'

        if int(nachfolger_xml.attrib['Lvl']) <= int(source_xml.attrib['Lvl']):
            sec_trenner.setPropertyValue('BackGraphicURL','')
            sec_trenner.Anchor.setString('')
            return
        
        sett_trenner = self.mb.settings_orga['trenner']
        trenner_format = sett_trenner['trenner']

        if trenner_format == 'farbe':
            
            if sec_xml.attrib['Zustand'] == 'zu':
                datei_name = get_dateiname()
            else:
                datei_name = nachfolger_xml.attrib['Name']
            sec_trenner.Anchor.setString(datei_name)
            sec_trenner.Anchor.ParaAdjust = CENTER
            sec_trenner.BackColor = KONST.FARBE_TRENNER_HINTERGRUND
            sec_trenner.Anchor.CharColor = KONST.FARBE_TRENNER_SCHRIFT 
            
            sec_trenner.setPropertyValue('BackGraphicURL','')

        elif trenner_format == 'strich':
            bgl = sec_trenner.BackGraphicLocation
            bgl.value = 'MIDDLE_BOTTOM'
             
            KONST.URL_TRENNER = 'vnd.sun.star.extension://xaver.roemers.organon/img/trenner.png'
            sec_trenner.Anchor.setString('')
            sec_trenner.setPropertyValue('BackGraphicURL',KONST.URL_TRENNER)
            sec_trenner.setPropertyValue("BackGraphicLocation", bgl)
        
        elif trenner_format == 'keiner':
            sec_trenner.Anchor.setString('')
            sec_trenner.setPropertyValue('BackGraphicURL','')
        
        elif trenner_format == 'user':
            sec_trenner.Anchor.setString('')
            sec_trenner.setPropertyValue('BackGraphicURL',sett_trenner['trenner_user_url'])

                
    def entferne_Trenner(self,sec):
        if self.mb.debug: log(inspect.stack)

        trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
        
        if trenner_name not in self.mb.doc.TextSections.ElementNames:
            return
#         log(inspect.stack,None,str(trenner_name+';'+sec.Name))
        trenner = self.mb.doc.TextSections.getByName(trenner_name)
        trenner.IsVisible = False
                
        
    def update_dict_zeilen_posY(self):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        C_XML = self.mb.class_XML
           
        eintraege = []
        C_XML.selbstaufruf = False
        C_XML.get_tree_info(root,eintraege)
        self.mb.props[T.AB].dict_zeilen_posY = {}
 
        i = 0
        for eintrag in eintraege:
            ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            if sicht == 'ja':
                self.mb.props[T.AB].dict_zeilen_posY.update({i*KONST.ZEILENHOEHE:eintrag})
                i += 1
        
        
    def positioniere_elemente_im_baum_neu(self):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()

        alle_sichtbaren = root.findall(".//*[@Sicht='ja']")

        index = 0
        for elem in alle_sichtbaren:
            zeile = self.mb.props[T.AB].Hauptfeld.getControl(elem.tag)
            zeile.setPosSize(0,index*KONST.ZEILENHOEHE,0,0,2)
            index += 1
            
    def disposing(self,ev):
        return False
            
            
class TreeView_Symbol_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,ctx,mb):
        if mb.debug: log(inspect.stack)
        self.ctx = ctx
        self.mb = mb
        
    def disposing(self,ev):
        return False

    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.props[T.AB].selektierte_zeile = ev.Source.Context.AccessibleContext.AccessibleName
        selbst = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
        
        if selbst in self.mb.props['Projekt'].dict_ordner:
            ist_ordner = True
        else:
            ist_ordner = False

        if ev.Buttons == MB_RIGHT:
            self.menu_rightclick(ev,ist_ordner,selbst)

        
        if ev.Buttons == MB_LEFT and ist_ordner:    
            if ev.ClickCount == 2: 
                
                try:
                    tree = self.mb.props[T.AB].xml_tree
                    root = tree.getroot()
                    selbst_xml = root.find('.//'+selbst)
                    zustand = selbst_xml.attrib['Zustand']
                    
                    if zustand == 'zu':
                        selbst_xml.attrib['Zustand'] = 'auf'
                        if selbst_xml.attrib['Art'] in ('dir','prj'):
                            ev.Source.Model.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16
                        if selbst_xml.attrib['Art'] == 'waste':
                            ev.Source.Model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_offen.png'
                        self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_hf(selbst,selbst_xml,zustand)    
                    else:
                        selbst_xml.attrib['Zustand'] = 'zu'
                        if selbst_xml.attrib['Art'] in ('dir','prj'):
                            bild_ordner = KONST.IMG_ORDNER_16
                            childs = list(selbst_xml)
                            if len(childs) > 0:
                                bild_ordner = KONST.IMG_ORDNER_VOLL_16
                            ev.Source.Model.ImageURL = bild_ordner
                        if selbst_xml.attrib['Art'] == 'waste':
                            ev.Source.Model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png'
                        self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_hf(selbst,selbst_xml,zustand)
                       
                    #Path = self.mb.pfade['settings'] + '/ElementTree.xml' 
                    Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
                    self.mb.props[T.AB].xml_tree
                    self.mb.class_Projekt.erzeuge_dict_ordner() 
                    self.mb.class_Baumansicht.korrigiere_scrollbar()
                except:
                    log(inspect.stack,tb())
                
            return False
        
    def mouseReleased(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def menu_rightclick(self,ev,ist_ordner,ordinal):
        try:
            
           
            papierkorb = self.mb.props[T.AB].Papierkorb
            projektordner = self.mb.props[T.AB].Projektordner
                
                
            
            controls = []
            maus_listener = Symbol_Popup_Mouse_Listener(self.mb)
            
            if ordinal == papierkorb:
                prop_names = ('Label',)
                prop_values = ('Papierkorb leeren',)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", 0, 0, 0,0, prop_names, prop_values)
                control.addMouseListener(maus_listener)
                controls.append(control)
            else:
            
                # IN PAPIERKORB VERSCHIEBEN
                if ordinal not in(papierkorb,projektordner):
                    prop_names = ('Label',)
                    prop_values = (LANG.IN_PAPIERKORB_VERSCHIEBEN,)
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", 0, 0, 0,0, prop_names, prop_values)
                    control.addMouseListener(maus_listener)
                    controls.append(control)
                
                # PROJEKTORDNER AUSKLAPPEN
                if ist_ordner:
                    if ordinal != papierkorb:
                        prop_names = ('Label',)
                        prop_values = (LANG.ORDNER_AUSKLAPPEN,)
                        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 0, 0, 0,0, prop_names, prop_values)
                        control.addMouseListener(maus_listener)
                        controls.append(control)
                
            b = 0
            h = 0
            
            for ctrl in controls:
                breite,hoehe = self.mb.kalkuliere_und_setze_Control(ctrl)
                if breite > b:
                    b = breite
                
                ctrl.setPosSize(30,h+5,breite+10,0,7)
                h += hoehe
            
            if len(controls) == 0:
                return
            
            # SEPERATOR
            control, model = self.mb.createControl(self.mb.ctx, "FixedLine", 20, 5, 5,h,(),())
            model.Orientation = 1
            controls.append(control)
            
            # Position finden
            BREITE = b
            HOEHE = h

            loc_cont = self.mb.current_Contr.Frame.ContainerWindow.AccessibleContext.LocationOnScreen
            pos_hf = self.mb.props['Projekt'].Hauptfeld.AccessibleContext.LocationOnScreen
            
            y = pos_hf.Y - loc_cont.Y
            y +=  ev.Source.Context.PosSize.Y 
            x = self.mb.dialog.AccessibleContext.LocationOnScreen.X - loc_cont.X + ev.Source.PosSize.X
            posSize = x+20,y,BREITE +50,HOEHE +10
            
            # Fenster erzeugen
            win,cont = self.mb.erzeuge_Dialog_Container(posSize,1+16)
            
            # Listener fuers Dispose des Fensters
            from menu_bar import Schliesse_Menu_Listener
            listener = Schliesse_Menu_Listener(self.mb)
            cont.addMouseListener(listener) 
            listener.ob = win
            

            cont.Model.BorderColor = 16580864
            
            maus_listener.window = win
            
            for ctrl in controls:
                cont.addControl(ctrl.Model.Label,ctrl)
        except:
            log(inspect.stack,tb())
            
            
class Symbol_Popup_Mouse_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.window = None
        
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        label = ev.Source.Text

        ordinal = self.mb.props[T.AB].selektierte_zeile
        
        if label == LANG.ORDNER_AUSKLAPPEN:
            self.mb.class_Funktionen.projektordner_ausklappen(ordinal)
        
        if label == LANG.IN_PAPIERKORB_VERSCHIEBEN:
            papierkorb = self.mb.props[T.AB].Papierkorb
            self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(ordinal,papierkorb,'inPapierkorbEinfuegen')
            
        if label == 'Papierkorb leeren':
            self.mb.class_Baumansicht.leere_Papierkorb()   

        self.window.dispose()
                
    def mouseEntered(self,ev):
        ev.value.Source.Model.FontWeight = 150
        return False
    def mouseExited(self,ev):
        ev.value.Source.Model.FontWeight = 100
        return False
    def disposing(self,ev):
        return False

class Tag1_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        if ev.Buttons == MB_LEFT and ev.ClickCount == 2 or ev.Buttons == MB_RIGHT:    
            self.mb.class_Funktionen.erzeuge_Tag1_Container(ev)
                
    def mouseEntered(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def disposing(self,ev):
        return False
    
class Tag2_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        if ev.Buttons == MB_LEFT and ev.ClickCount == 2 or ev.Buttons == MB_RIGHT:  
            ordinal = ev.Source.Context.AccessibleContext.AccessibleName  
            self.mb.class_Funktionen.erzeuge_Tag2_Container(ev,ordinal)
                
    def mouseEntered(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False


  
