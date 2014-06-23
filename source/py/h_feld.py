# -*- coding: utf-8 -*-

import unohelper


class Main_Container():
    
    def __init__(self,mb):
        global pd
        pd = mb.pd
        
        self.dialog = mb.dialog
        self.ctx = mb.ctx
        self.mb = mb
        self.listenerDir = Dir_Listener(self.ctx,self.mb)
        self.tag1_listener = Tag1_Listener(self.mb)
        
        
    def start(self):
        if self.mb.debug: print(self.mb.debug_time(),'start')

        self.erzeuge_Navigations_Hauptfeld()  
        self.erzeuge_Scrollbar(self.dialog,self.ctx)

        
    def erzeuge_Navigations_Hauptfeld(self):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Navigations_Hauptfeld')
        # Das aeussere Hauptfeld wird fuers Scrollen benoetigt. Das innere und eigentliche
        # Hauptfeld scrollt dann in diesem Hauptfeld_aussen

        # Hauptfeld_Aussen
        Attr = (22,KONST.ZEILENHOEHE+2,1000,1800,'Hauptfeld_aussen')    
        PosX,PosY,Width,Height,Name = Attr
        
        control1, model1 = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        
             
        self.dialog.addControl('Hauptfeld_aussen',control1)  
        
        # eigentliches Hauptfeld
        Attr = (0,0,1000,10000,)    
        PosX,PosY,Width,Height = Attr
         
        control2, model2 = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  

        model2.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
        control1.addControl('Hauptfeld',control2)  
    
        self.mb.Hauptfeld = control2

  
    def erzeuge_Verzeichniseintrag(self,eintrag,class_Zeilen_Listener,index=0):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Verzeichniseintrag')
        # wird in projects aufgerufen
        ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag        

        ##### Aeusserer Container #######
        
        Attr = (2,KONST.ZEILENHOEHE*index,600,20)    
        PosX,PosY,Width,Height = Attr
        
        control, model = self.mb.createControl(self.ctx,"Container",PosX,PosY,Width,Height,(),() )  
        model.Text = ordinal
        model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
            
        self.mb.Hauptfeld.addControl(ordinal,control)
        
        
        ### einzelne Elemente #####
        tag1X,tag2X,tag3X = 0,0,0
        if int(self.mb.settings_proj['tag1']):
            tag1X = 16
        if int(self.mb.settings_proj['tag2']):
            tag2X = 16
        if int(self.mb.settings_proj['tag3']):
            tag3X = 16
        
        
        # Textfeld
        Farbe_Textfeld = 102023
        Attr = (32+int(lvl)*16+tag1X+tag2X+tag3X,0,400,20,'egal', Farbe_Textfeld)   
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control1, model1 = self.mb.createControl(self.ctx,"Edit",PosX,PosY,Width,Height,(),() )  
        model1.Text = name
        model1.Border = False
        model1.BackgroundColor = KONST.FARBE_ZEILE_STANDARD
        model1.ReadOnly = True

        control1.addMouseListener(class_Zeilen_Listener) 
        control1.addMouseMotionListener(class_Zeilen_Listener)
        control1.addFocusListener(class_Zeilen_Listener)

        control.addControl('textfeld',control1)
        
        # Icon
        Color__Container = 10202
        Attr = (16+int(lvl)*16,2,16,16,'egal', Color__Container)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control2, model2 = self.mb.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  

        if art == 'waste':
            control2.addMouseListener(self.listenerDir)
            self.mb.Papierkorb = ordinal
        
        if art == 'prj':
            self.mb.Projektordner = ordinal
             
        
        # icons sind unter: C:\Program Files (x86)\LibreOffice 4\share\config ... \images\res
        # richtige Icons finden
        if art in ('dir','prj'):
            control2.addMouseListener(self.listenerDir)

            if zustand == 'auf':
                model2.ImageURL = KONST.IMG_ORDNER_GEOEFFNET_16 
            else:
                bild_ordner = KONST.IMG_ORDNER_16

                tree = self.mb.xml_tree 
                root = tree.getroot()

                ordner_xml = root.find('.//'+ordinal)
                if ordner_xml != None:
                    childs = list(ordner_xml)
                    if len(childs) > 0:
                        bild_ordner = KONST.IMG_ORDNER_VOLL_16

                model2.ImageURL = bild_ordner
                
        elif art == 'pg':
            model2.ImageURL = 'private:graphicrepository/res/sx03150.png' 
        elif art == 'waste':
            if zustand == 'zu':
                model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png' 
            else:
                model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_offen.png' 
                
        model2.Border = 0

        control.addControl('icon',control2)                       
        # return ist nur fuer neu angelegte Dokumente nutzbar 
        
        
        if int(self.mb.settings_proj['tag1']):
            
            # Tag1 Farbe
            Color__Container = 10202
            Attr = (32+int(lvl)*16,2,16,16,'egal', Color__Container)    
            PosX,PosY,Width,Height,Name,Color = Attr
            
            control_tag1, model_tag1 = self.mb.createControl(self.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
            model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
            model_tag1.Border = 0
            control_tag1.addMouseListener(self.tag1_listener)
            
            control.addControl('tag1',control_tag1)
            
        
        if sicht == 'nein':
            control.Visible = False
        else:              
            index += 1
        return index  
     
    def erzeuge_Scrollbar(self,dialog,ctx):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Scrollbar')
        
        nav_cont_aussen = dialog.getControl('Hauptfeld_aussen')
        nav_cont = nav_cont_aussen.getControl('Hauptfeld')
          
        MenuBar = dialog.getControl('Organon_Menu_Bar')
        MBHoehe = MenuBar.PosSize.value.Height + MenuBar.PosSize.value.Y
        NCHoehe = 0 #nav_cont.PosSize.value.Height
        NCPosY  = nav_cont.PosSize.value.Y
        Y =  NCHoehe + NCPosY + MBHoehe
        Height = dialog.PosSize.value.Height - Y - 25

        Attr = (0,Y,20,Height,'scrollbar', None)    
        PosX,PosY,Width,Height,Name,Color = Attr
        
        control, model = self.mb.createControl(ctx,"ScrollBar",PosX,PosY,Width,Height,(),() )  
        #model.Name = 'ScrollBar_Hauptfeld'
        model.Orientation = 1
        model.LiveScroll = True        
        hoehe_Hauptfeld = nav_cont.Size.value.Height 
        model.ScrollValueMax = hoehe_Hauptfeld/4
        
        self.mb.scrollbar = control
        
        dialog.addControl('ScrollBar',control)  

        listener = ScrollBarListener(nav_cont)
        control.addAdjustmentListener(listener) 
        
    def korrigiere_scrollbar(self):
        if self.mb.debug: print(self.mb.debug_time(),'korrigiere_scrollbar')
        
        dialog = self.mb.dialog
        #SB = dialog.getControl('ScrollBar')
        SB = self.mb.scrollbar
        
        
        hoehe = sorted(list(self.mb.dict_zeilen_posY))
        if hoehe != []:
            max =  hoehe[-1] - (self.mb.dialog.PosSize.Height - 100)
            if max > 0:
                SB.Visible = True
                SB.Maximum = max
                
            else:
                SB.Maximum  = 1
                SB.Visible = False
                nav_cont_aussen = dialog.getControl('Hauptfeld_aussen')
                nav_cont = nav_cont_aussen.getControl('Hauptfeld')
                nav_cont.setPosSize(0, 0,0,0,2)
                #pd()
    
    def erzeuge_neue_Zeile(self,ordner_oder_datei):
        
        #self.mb.timer_start = self.mb.time.clock()
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_neue_Zeile')
        
        if self.mb.selektierte_zeile == None:       
            self.mb.Mitteilungen.nachricht(self.mb.lang.ZEILE_AUSWAEHLEN,'infobox')
            return None
        else:
            try:
                StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
                StatusIndicator.start(self.mb.lang.ERZEUGE_NEUE_ZEILE,2)
                StatusIndicator.setValue(2)
                
                self.mb.doc.lockControllers()
                
                ord_sel_zeile = self.mb.selektierte_zeile.AccessibleName
                
                # XML TREE
                tree = self.mb.xml_tree
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
                self.erzeuge_Verzeichniseintrag(eintrag,self.mb.class_Zeilen_Listener)
                self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)
                            
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
                tb()
            StatusIndicator.end()

            
            
    def kontrolle(self):
        root2 = self.mb.xml_tree.getroot()
        alle = root2.findall('.//')
        asd = []
        for i in alle:
            asd.append((i.tag,i.attrib['Lvl'],i.attrib['Sicht'],i))


    def leere_Papierkorb(self):
        #self.mb.timer_start = self.mb.time.clock()
        if self.mb.debug: print(self.mb.debug_time(),'leere_Papierkorb')
        try:
            self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener)
            
            tree = self.mb.xml_tree
            root = tree.getroot()
            C_XML = self.mb.class_XML
    
            papierkorb_xml = root.find(".//"+self.mb.Papierkorb)  
            papierkorb_xml.attrib['Zustand'] = 'zu'
            papierkorb_inhalt1 = []
            C_XML.selbstaufruf = False
            C_XML.get_tree_info(papierkorb_xml,papierkorb_inhalt1)        
            selektierter_ist_im_papierkorb = False
    
            # Zeilen im Hauptfeld loeschen
            for eintrag in papierkorb_inhalt1:        
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
    
                if not selektierter_ist_im_papierkorb:
                    if self.mb.selektierte_zeile.AccessibleName == ordinal:
                        selektierter_ist_im_papierkorb = True
                
                if art != 'waste':
                    contr = self.mb.Hauptfeld.getControl(ordinal)               
                    contr.dispose()
    
            # Eintraege in XML Tree loeschen
            papierkorb_inhalt = list(papierkorb_xml)
            for verworfene in papierkorb_inhalt:
                papierkorb_xml.remove(verworfene)
                                
            Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
            tree.write(Path)
    
            self.erneuere_selektierungen(selektierter_ist_im_papierkorb) 
    
            # loesche Bereich(e) und Datei(en)
            self.loesche_Bereiche_und_Dateien(papierkorb_inhalt1,papierkorb_inhalt)
           
            # XML,Ansicht und Dicts neu ordnen
            # vielleicht gibt es noch eine andere Moeglichkeit?
            # hier wird der gesamte Baum fuer jede Neuordnung nochmals durchlaufen
            # -> Performance?
            # Insgesamt wird der Baum 4x durchlaufen
            self.mb.class_Projekt.erzeuge_dict_ordner()
            self.mb.class_Zeilen_Listener.update_dict_zeilen_posY()    # <- dict wird durch tree aufgebaut. schon erneuert?    
            
            self.mb.Papierkorb_geleert = True
            self.erneuere_dict_bereiche()
            self.korrigiere_scrollbar()
    
            self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)
        except:
            tb()
        
                
    def erneuere_selektierungen(self,selektierter_ist_im_papierkorb):
        if self.mb.debug: print(self.mb.debug_time(),'erneuere_selektierungen')
        zeile_Papierkorb = self.mb.Hauptfeld.getControl(self.mb.Papierkorb)
        textfeld_Papierkorb = zeile_Papierkorb.getControl('textfeld')
        icon_Papierkorb = zeile_Papierkorb.getControl('icon')
        icon_Papierkorb.Model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png'
        
        if selektierter_ist_im_papierkorb:
            self.mb.selektierte_zeile = None
            textfeld_Papierkorb.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE
            self.mb.selektierte_Zeile_alt = textfeld_Papierkorb
            papierkorb_bereichsname = self.mb.dict_bereiche['ordinal'][self.mb.Papierkorb]
            bereich_papierkorb = self.mb.doc.TextSections.getByName(papierkorb_bereichsname)
            
            bereich_papierkorb.IsVisible = True
            self.mb.sichtbare_bereiche = [papierkorb_bereichsname]
            
            

    def loesche_Bereiche_und_Dateien(self,papierkorb_inhalt1,papierkorb_inhalt):
        if self.mb.debug: print(self.mb.debug_time(),'loesche_Bereiche_und_Dateien')

        for inhalt in papierkorb_inhalt1:
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = inhalt
            
            try:
            
                if art != 'waste':
                    zahl = ordinal.split('nr')[1]
                    # loesche datei ordinal
                    Path = os.path.join(self.mb.pfade['odts'], 'nr%s.odt' % zahl)
                    os.remove(Path)
                    
                    bereichsname = self.mb.dict_bereiche['ordinal'][ordinal]
                    
                    # loesche text der Datei im Dokument
                    sections = self.mb.doc.TextSections
                    sec = sections.getByName(bereichsname) 
                    
                    
#                     # neu TEST  ########
#                     SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
#                     sec.setPropertyValue('FileLink',SFLink)
                    
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
                                   
                    textSectionCursor = self.mb.doc.Text.createTextCursorByRange(sec.Anchor)
                    #textSectionCursor.gotoRange(sec.End,True)
                    textSectionCursor.setString('')
    
                    trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
                    sec.dispose()
                    
                    # Trenner loeschen
                    if trenner_name in self.mb.doc.TextSections.ElementNames:
                        trenner = self.mb.doc.TextSections.getByName(trenner_name)
                        textSectionCursor.gotoRange(trenner.Anchor,True)
                        trenner.dispose()
                        textSectionCursor.setString('')
                  
                    textSectionCursor.gotoEnd(False)
                    while textSectionCursor.TextSection == None:
                        textSectionCursor.goLeft(1,True)
                    #textSectionCursor.goLeft(1,True)
                    textSectionCursor.setString('')
    
                    
            except:
                tb()
        
    def erneuere_dict_bereiche(self):
        if self.mb.debug: print(self.mb.debug_time(),'erneuere_dict_bereiche')
        sections = self.mb.doc.TextSections

        tree = self.mb.xml_tree
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
             
        self.mb.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        
        
    
from com.sun.star.awt import XMouseListener,XMouseMotionListener,XFocusListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT
from com.sun.star.awt.MouseButton import RIGHT as MB_RIGHT

class Zeilen_Listener (unohelper.Base, XMouseListener,XMouseMotionListener,XFocusListener):
    def __init__(self,hf,ctx,mb):
        
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
        self.mb.selektierte_Zeile_alt = None
        self.dragged = False
        
              
    def mouseMoved(self,ev):  
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self, ev):
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener)
        
        try:
            # die gesamte Zeile, control (ordinal)
            self.mb.selektierte_zeile = ev.Source.Context.AccessibleContext  
            # control 'textfeld'   
            zeile = ev.Source
            # selektierte Zeile einfaerben, ehem. sel. Zeile zuruecksetzen
            zeile.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE 
            if self.mb.selektierte_Zeile_alt != None:
                if zeile != self.mb.selektierte_Zeile_alt:
                    self.mb.selektierte_Zeile_alt.Model.BackgroundColor = KONST.FARBE_ZEILE_STANDARD
            
            # bei bearbeitetem Bereich: speichern  
            if self.mb.selektierte_Zeile_alt != None: 
                if self.mb.bereich_wurde_bearbeitet == True:
                    zeilenordinal_zeile_alt = self.mb.selektierte_Zeile_alt.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName   
                    bereichsname_zeile_alt = self.mb.dict_bereiche['ordinal'][zeilenordinal_zeile_alt]
                    bereich_zeile_alt = self.mb.doc.TextSections.getByName(bereichsname_zeile_alt)
    
                    self.mb.bereich_wurde_bearbeitet = False
                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(bereich_zeile_alt.FileLink.FileURL,bereichsname_zeile_alt)
                    
            self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)      
            self.mb.selektierte_Zeile_alt = zeile
            
            # Bei Doppelclick Zeileneintrag bearbeiten
            if ev.Buttons == MB_LEFT:   
                if ev.ClickCount == 2: 
                    
                    # Projektordner von der Umbenennung ausnehmen
                    zeilenordinal = self.mb.selektierte_Zeile_alt.Context.AccessibleContext.AccessibleName
                    root = self.mb.xml_tree.getroot()
                    zeile_xml = root.find('.//'+zeilenordinal)
                    art = zeile_xml.attrib['Art']
                    
                    if art == 'prj':
                        self.mb.Mitteilungen.nachricht('Der Projektordner kann nicht umbenannt werden',"infobox")
                        return False
                    else:
                        zeile.Model.ReadOnly = False 
                        zeile.Model.BackgroundColor = KONST.FARBE_EDITIERTE_ZEILE
                        self.edit_text = True   
                        return False
            
            return False
        except:
            tb()
    
    def mouseReleased(self, ev):
        
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
            
        zeilenordinal =  self.mb.selektierte_zeile.AccessibleName
        bereichsname = self.mb.dict_bereiche['ordinal'][zeilenordinal]
        bereich = self.mb.doc.TextSections.getByName(bereichsname)

        self.mb.viewcursor.gotoRange(bereich.Anchor,False)
        self.mb.viewcursor.gotoStart(False)
            
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)
        
        self.mb.loesche_undo_Aktionen()
                           
    def focusGained(self,ev):
        return False  
    
    def focusLost(self, ev):
        # Bearbeitung des Zeileneintrags wieder ausschalten
        if self.edit_text == True:

            ev.Source.Model.ReadOnly = True 
            self.edit_text = False
            # neuen Text in die xml Datei eintragen
            tree = self.mb.xml_tree
            root = tree.getroot() 
            zeile_ord = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
            xml_elem = root.find('.//'+zeile_ord)
            xml_elem.attrib['Name'] = ev.Source.Text
            
            Path1 = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
            tree.write(Path1)
        return False

    def mouseDragged(self,ev):
        # Papierkorb darf nicht verschoben werden
        ordinal = ev.Source.Context.AccessibleContext.AccessibleName
        if ordinal in (self.mb.Papierkorb,self.mb.Projektordner):
            self.mb.Mitteilungen.nachricht(self.mb.lang.NICHT_VERSCHIEBBAR,"infobox")
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

                    sourceArt = self.mb.dict_zeilen_posY[sourceYPos][4]

                    if sourceArt in ('dir','prj'):
                        subelements = self.mb.dict_ordner[sourceName]
                    else:
                        subelements = 'quatsch'
                    
                    if sourceArt in ('dir','prj') and ordinal in subelements:
                        ImageURL = KONST.IMAGE_GESCHEITERT
                        self.einfuegen = 'gescheitert'  

                    # Wenn Erster
                    elif ordinal == ord_erster:
                        if zeileY < 12:
                            ImageURL = KONST.IMAGE_PFEIL_HOCH
                            self.einfuegen = 'drueber'
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
            tb()
                
    def berechne_pos(self,Y):
        
        y = (math.floor(Y/KONST.ZEILENHOEHE)) 
        
        # Position des Papierkorbs 
        tree = self.mb.xml_tree
        root = tree.getroot()   
        ord_papierkorb = root.find(".//*[@Art='waste']").tag
        zeile_papierkorb = self.mb.Hauptfeld.getControl(ord_papierkorb)
        pos_papierkorb = zeile_papierkorb.PosSize.Y/KONST.ZEILENHOEHE

        # abfangen: wenn Mauspos ueber Hauptfeld oder unter Papierkorb hinausgeht
        if y < 0:
            y = 0
        elif y > pos_papierkorb:
            y = pos_papierkorb

        posY = (y * KONST.ZEILENHOEHE)       
        posY_nachf = ((y+1) * KONST.ZEILENHOEHE)
        
        ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = self.mb.dict_zeilen_posY[posY]
        ord_erster = self.mb.dict_zeilen_posY[0][0]
        
        # try/except um nur eine Abfrage an das Dict zu stellen
        try:
            nf = self.mb.dict_zeilen_posY[posY_nachf]
            lvl_nf = nf[3]
            art_nf = nf[4]
            ord_nf = nf[0]
            nachfolger = (lvl_nf,art_nf,ord_nf)
        except:
            # Papierkorb: Nachfolger ist immer auf hoeherem lvl
            nachfolger = (100,'keiner','sdfsf')
        
        zeileY = Y - posY
        info = (zeileY,art,lvl,nachfolger,ord_erster,ordinal)
        zielordner = self.mb.Hauptfeld.getControl(ordinal)
         
        return (zielordner,info)
 
           
    def zeilen_neu_ordnen(self,source,target,action):
        #self.mb.timer_start = self.mb.time.clock()
        if self.mb.debug: print(self.mb.debug_time(),'zeilen_neu_ordnen')
        
        if action != 'gescheitert':
            if 'vorNachfolger' in action:
                # wenn der Nachfolger gleich der Quelle ist, nicht verschieben
                if source == action[1]:
                    return
                else:  
                    eintraege = self.xml_neu_ordnen(source,target,action)
                    self.hf_neu_ordnen(eintraege)
                    if self.mb.debug: print(self.mb.debug_time(),'nach hf')
                    
            else:  
                eintraege = self.xml_neu_ordnen(source,target,action)
                self.hf_neu_ordnen(eintraege)
                if self.mb.debug: print(self.mb.debug_time(),'nach hf')
                
            Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
            if self.mb.debug: print(self.mb.debug_time(),'xml_tree.write')
            self.mb.xml_tree.write(Path)
            
                
 
    def hf_neu_ordnen(self,eintraege): 
        if self.mb.debug: print(self.mb.debug_time(),'hf_neu_ordnen')
        
        tree = self.mb.xml_tree
        root = tree.getroot()
        
        # ordnen des dict_zeilen_posY
        self.mb.dict_zeilen_posY = {}
        index = 0
        
        tag1X,tag2X,tag3X = 0,0,0
        if int(self.mb.settings_proj['tag1']):
            tag1X = 16
        if int(self.mb.settings_proj['tag2']):
            tag2X = 16
        if int(self.mb.settings_proj['tag3']):
            tag3X = 16
        
        
        for eintrag in eintraege:
            
            ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            if sicht == 'ja':
                
                # Y_Wert sichtbarer Eintraege setzen
                cont = self.mb.Hauptfeld.getControl(ordinal)
                # der X-Wert ALLER Eintraege wird in xml_m neu gesetzt
                cont.Peer.setPosSize(0,KONST.ZEILENHOEHE*index,0,0,2)# 2: Flag fuer: nur Y Wert aendern
                
                iconArt = cont.getControl('icon')
                iconArt.Peer.setPosSize(16+int(lvl)*16,0,0,0,1)   
                
                if int(self.mb.settings_proj['tag1']):
                    tag1_cont = cont.getControl('tag1')
                    tag1_cont.Peer.setPosSize(32+int(lvl)*16,0,0,0,1)  
                if int(self.mb.settings_proj['tag2']):
                    tag2_cont = cont.getControl('tag2')
                    tag2_cont.Peer.setPosSize(32+int(lvl)*16+tag1X,0,0,0,1)   
                if int(self.mb.settings_proj['tag3']):
                    tag3_cont = cont.getControl('tag3')
                    tag3_cont.Peer.setPosSize(32+int(lvl)*16+tag1X+tag2X,0,0,0,1)    
                
                textfeld = cont.getControl('textfeld')
                textfeld.Peer.setPosSize(32+int(lvl)*16+tag1X+tag2X+tag3X,0,0,0,1)        
                 
                # dict_zeilen_posY updaten            
                self.mb.dict_zeilen_posY.update({KONST.ZEILENHOEHE*index:eintrag})                   
                index += 1  
             
        # dict_ordner updaten
        self.mb.class_Projekt.erzeuge_dict_ordner()
        
        sections = self.mb.doc.TextSections

        # Bereiche neu verlinken
        self.verlinke_Bereiche(sections)
        # Sichtbarkeit der Bereiche umschalten
        self.schalte_sichtbarkeit_der_Bereiche()        

        
    def xml_neu_ordnen(self,source,target,action):
        if self.mb.debug: print(self.mb.debug_time(),'xml_neu_ordnen')
        
        tree = self.mb.xml_tree
        root = tree.getroot()
        C_XML = self.mb.class_XML
        target_xml = root.find('.//'+target)
        
        # bei Verschieben in einen geschlossenen Ordner
        if 'inOrdnerEinfuegen' in action and target_xml.attrib['Zustand'] == 'zu':
            target_xml.attrib['Zustand'] = 'auf'
            self.schalte_sichtbarkeit_des_hf(target,target_xml,'zu')
            
            tar_cont = self.mb.Hauptfeld.getControl(target)
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
        if self.mb.debug: print(self.mb.debug_time(),'schalte_sichtbarkeit_des_hf')
        
        tree = self.mb.xml_tree
        root = tree.getroot()
         
        def durchlaufe_baum(dir):  
            for child in dir:
                #print(child.tag,child.attrib['Name'])
                if child.attrib['Art'] in ('dir','waste','prj'):  
                    #print('wird sichtbar 1', child.attrib['Name'],child.attrib['Zustand'],child.attrib['Parent'])  
                    tar = self.mb.Hauptfeld.getControl(child.tag)
                    tar.Visible = True   
                    
                    elem = root.find('.//'+ child.tag)
                    elem.attrib['Sicht'] = 'ja'
                           
                    if child.attrib['Zustand'] == 'auf':                        
                        durchlaufe_baum(child)
                else:
                    if dir.attrib['Zustand'] == 'auf':
                        #print('wird sichtbar 2', child.attrib['Name'],'pg')
                        tar = self.mb.Hauptfeld.getControl(child.tag)
                        tar.Visible = True           
                        
                        elem = root.find('.//'+ child.tag)
                        elem.attrib['Sicht'] = 'ja'   
        
        if zeige_projektordner:
            selbst_xml.attrib['Zustand'] = 'auf'
            durchlaufe_baum(selbst_xml)
        
        if zustand == 'auf':
            dir = self.mb.dict_ordner[selbst]
            
            for child in dir:
                if child != selbst:
                    tar = self.mb.Hauptfeld.getControl(child)
                    tar.Visible = False
                    elem = root.find('.//'+ child)
                    elem.attrib['Sicht'] = 'nein'
                    #print('wird unsichtbar',elem.attrib['Name'],elem.attrib['Sicht'])
        else:   
            durchlaufe_baum(selbst_xml)
            
        self.positioniere_elemente_im_baum_neu()          
        self.update_dict_zeilen_posY() 
        
        
        
        
    def verlinke_Bereiche(self,sections):
        if self.mb.debug: print(self.mb.debug_time(),'verlinke_Bereiche')
        
        # langsame und sichere Loesung: es werden alle Bereiche neu verlinkt, 
        # nicht nur die verschobenen
        # NEU UND SCHNELLER: Es werden nur noch sichtbare Bereiche neu verlinkt in: self.verlinke_Sektion

        # Der VC Listener wird von IsVisible ausgeloest,
        # daher wird er vorher ab- und hinterher wieder angeschaltet
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 
                
        tree = self.mb.xml_tree
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
            tb()
                
        for i in range (len(alle_Zeilen)):
            ordinal = alle_Zeilen[i].tag
            Path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)  
            
            Bereichsname_dict.update({'OrganonSec'+str(i):Path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(i)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(i):ordinal})
            
        self.mb.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)         
        
        
    def schalte_sichtbarkeit_der_Bereiche(self):
        if self.mb.debug: print(self.mb.debug_time(),'schalte_sichtbarkeit_der_Bereiche')
        try:
        
            # Der VC Listener wird von IsVisible ausgeloest,
            # daher wird er vorher ab- und hinterher wieder angeschaltet
            self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 
    
            zeilenordinal =  self.mb.selektierte_zeile.AccessibleName
            bereichsname = self.mb.dict_bereiche['ordinal'][zeilenordinal]
    
            # Ordner
            if zeilenordinal in self.mb.dict_ordner:
                zeilen_in_ordner_ordinal = self.mb.dict_ordner[zeilenordinal]
                
                # alle Zeilen im Ordner einblenden
                for z in zeilen_in_ordner_ordinal:
                    ordnereintrag_name = self.mb.dict_bereiche['ordinal'][z]
                                            
                    z_in_ordner = self.mb.doc.TextSections.getByName(ordnereintrag_name)

                    self.verlinke_Sektion(ordnereintrag_name,z_in_ordner)
                    
                    z_in_ordner.IsVisible = True
                    self.mache_Kind_Bereiche_sichtbar(z_in_ordner)
                    # Wenn mehr als nur ein geschlossener Ordner zu sehen ist
                    if len(zeilen_in_ordner_ordinal) > 1:
                        if zeilen_in_ordner_ordinal.index(z) != len(zeilen_in_ordner_ordinal)-1:
                            self.zeige_Trenner(z_in_ordner)
                        else:
                            # Sollte der letzte Trenner noch sichtbar sein, entfernen
                            name_sec = self.mb.dict_bereiche['ordinal'][z]
                            orga_sec = self.mb.doc.TextSections.getByName(name_sec)
                            self.entferne_Trenner(orga_sec)
        
                # uebrige noch sichtbare ausblenden
                for bereich in reversed(self.mb.sichtbare_bereiche):                        
                    bereich_ord = self.mb.dict_bereiche['Bereichsname-ordinal'][bereich]                     
                    if bereich_ord not in zeilen_in_ordner_ordinal:                            
                        sec = self.mb.doc.TextSections.getByName(bereich)
                        sec.IsVisible = False 
                        self.entferne_Trenner(sec)
                                           
                # sichtbare Bereiche wieder in Prop eintragen
                self.mb.sichtbare_bereiche = []  
                for b in zeilen_in_ordner_ordinal:
                    self.mb.sichtbare_bereiche.append(self.mb.dict_bereiche['ordinal'][b])                    
            else:
            # Seiten 
                selekt_bereich_name = self.mb.dict_bereiche['ordinal'][zeilenordinal]
                selekt_bereich = self.mb.doc.TextSections.getByName(selekt_bereich_name)

                self.verlinke_Sektion(selekt_bereich_name,selekt_bereich)
                
                selekt_bereich.IsVisible = True
                self.mache_Kind_Bereiche_sichtbar(selekt_bereich)
                self.entferne_Trenner(selekt_bereich)
                
                for bereich in reversed(self.mb.sichtbare_bereiche):
                    if bereich != selekt_bereich_name:
                        section = self.mb.doc.TextSections.getByName(bereich)  
                        section.IsVisible = False
                        self.entferne_Trenner(section)
                        
                self.mb.sichtbare_bereiche = [selekt_bereich_name]

            self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener) 
        except:
            tb()
        
    def schalte_sichtbarkeit_des_ersten_Bereichs(self):
        if self.mb.debug: print(self.mb.debug_time(),'schalte_sichtbarkeit_des_ersten_Bereichs')
   
        # Der VC Listener wird von IsVisible ausgeloest,
        # daher wird er vorher ab- und hinterher wieder angeschaltet
        self.mb.current_Contr.removeSelectionChangeListener(self.mb.VC_selection_listener) 

        zeilenordinal =  self.mb.selektierte_zeile.AccessibleName
        bereichsname = self.mb.dict_bereiche['ordinal'][zeilenordinal]

        # Ordner
        if zeilenordinal in self.mb.dict_ordner:
            zeilen_in_ordner_ordinal = self.mb.dict_ordner[zeilenordinal][0],
            
            # alle Zeilen im Ordner einblenden
            for z in zeilen_in_ordner_ordinal:
                ordnereintrag_name = self.mb.dict_bereiche['ordinal'][z]
                z_in_ordner = self.mb.doc.TextSections.getByName(ordnereintrag_name)
                self.verlinke_Sektion(ordnereintrag_name,z_in_ordner)    

            # uebrige noch sichtbare ausblenden
            for bereich in self.mb.sichtbare_bereiche:                        
                bereich_ord = self.mb.dict_bereiche['Bereichsname-ordinal'][bereich]                     
                if bereich_ord not in zeilen_in_ordner_ordinal:                            
                    sec = self.mb.doc.TextSections.getByName(bereich)
                    sec.IsVisible = False 
                    self.entferne_Trenner(sec)                 
        
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener) 

    def verlinke_Sektion(self,name,bereich):
        #if self.mb.debug: print(self.mb.debug_time(),'verlinke_Sektion')
        url_in_dict = uno.systemPathToFileUrl(self.mb.dict_bereiche['Bereichsname'][name])
        url_sec = bereich.FileLink.FileURL

        if url_in_dict != url_sec:
        
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FilterName = 'writer8'
            SFLink.FileURL = url_in_dict

            bereich.setPropertyValue('FileLink',SFLink)
          
    
    def mache_Kind_Bereiche_sichtbar(self,sec):
        for kind in sec.ChildSections:
            kind.IsVisible = True
            if len(kind.ChildSections) > 0:
                self.mache_Kind_Bereiche_sichtbar(kind)
    
    def zeige_Trenner(self,sec):  
        
        trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
         
        if trenner_name in self.mb.doc.TextSections.ElementNames:
            trennerSec = self.mb.doc.TextSections.getByName(trenner_name)
            trennerSec.IsVisible = True
            if len(sec.ChildSections) != 0:
                trennerSec.BackColor = sec.ChildSections[0].BackColor
            return
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                 
        newSection.setName(trenner_name)
        newSection.IsProtected = True
        if len(sec.ChildSections) != 0:
            newSection.BackColor = sec.ChildSections[0].BackColor

        
        sec_nachfolger_name = 'OrganonSec' + str(int(sec.Name.split('OrganonSec')[1])+1)
        sec_nachfolger = self.mb.doc.TextSections.getByName(sec_nachfolger_name)

        cur = sec_nachfolger.Anchor.Text.createTextCursor()
        cur.gotoRange(sec_nachfolger.Anchor,False)
        cur.gotoRange(cur.TextSection.Anchor.Start,False)        
        
        sec.Anchor.End.Text.insertTextContent(cur, newSection, False)
        
        cur.goLeft(1,False)
        cur.ParaStyleName = 'Standard'
        #cur.PageDescName = ""
        #cur.Text.insertString(cur,'a',False)

        
        bgl = newSection.BackGraphicLocation
        bgl.value = 'MIDDLE_BOTTOM'
         
        KONST.URL_TRENNER = 'vnd.sun.star.extension://xaver.roemers.organon/img/trenner.png'
 
        newSection.setPropertyValue('BackGraphicURL',KONST.URL_TRENNER)
        newSection.setPropertyValue("BackGraphicLocation", bgl)
        cur.PageDescName = ""
        #pd() #newSection.setPropertyValue("ParaStyleName", 'Standard')
        
    def entferne_Trenner(self,sec):
        
        #print('entferne', sec.Name)
        trenner_name = 'trenner' + sec.Name.split('OrganonSec')[1]
        
        if trenner_name not in self.mb.doc.TextSections.ElementNames:
            return
        
        trenner = self.mb.doc.TextSections.getByName(trenner_name)
        trenner.IsVisible = False
        return
        cur = self.mb.doc.Text.createTextCursor()
        cur.gotoRange(trenner.Anchor,False)
        cur.collapseToEnd()
        trenner.dispose()

        cur.goLeft(1,True)
        cur.setString('')
        
    def update_dict_zeilen_posY(self):
        if self.mb.debug: print(self.mb.debug_time(),'update_dict_zeilen_posY')
        
        tree = self.mb.xml_tree
        root = tree.getroot()
        C_XML = self.mb.class_XML
           
        eintraege = []
        C_XML.selbstaufruf = False
        C_XML.get_tree_info(root,eintraege)
        self.mb.dict_zeilen_posY = {}
 
        i = 0
        for eintrag in eintraege:
            ordinal,parent,text,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            if sicht == 'ja':
                self.mb.dict_zeilen_posY.update({i*KONST.ZEILENHOEHE:eintrag})
                i += 1
        
        
    def positioniere_elemente_im_baum_neu(self):
        if self.mb.debug: print(self.mb.debug_time(),'positioniere_elemente_im_baum_neu')
        
        tree = self.mb.xml_tree
        root = tree.getroot()

        alle_sichtbaren = root.findall(".//*[@Sicht='ja']")

        index = 0
        for elem in alle_sichtbaren:
            zeile = self.mb.Hauptfeld.getControl(elem.tag)
            zeile.setPosSize(0,index*KONST.ZEILENHOEHE,0,0,2)
            index += 1
            
    def disposing(self,ev):
        return False
        
            
            
class Dir_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,ctx,MenuBar):
        self.ctx = ctx
        self.mb = MenuBar
        
    def disposing(self,ev):
        return False

    def mousePressed(self, ev):
        self.mb.selektierte_zeile = ev.Source.Context.AccessibleContext

        if ev.Buttons == MB_LEFT:    
            if ev.ClickCount == 2: 
                
                try:
                    selbst = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
                    
                    tree = self.mb.xml_tree
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
                    tree.write(Path)
                    self.mb.class_Projekt.erzeuge_dict_ordner() 
                    self.mb.class_Hauptfeld.korrigiere_scrollbar()
                except:
                    tb()
                
            return False
        
    def mouseReleased(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False

class Tag1_Listener (unohelper.Base, XMouseListener):
    
    def __init__(self,mb):
        self.mb = mb
        
    def mousePressed(self, ev):
        
        if ev.Buttons == MB_LEFT:    
            if ev.ClickCount == 2: 
                self.mb.class_Funktionen.erzeuge_Tag1_Container(ev)
    def mouseEntered(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def disposing(self,ev):
        return False

from com.sun.star.awt import XAdjustmentListener
class ScrollBarListener (unohelper.Base,XAdjustmentListener):
    
    def __init__(self,nav_cont):        
        self.nav_cont = nav_cont
            
    def adjustmentValueChanged(self,ev):
        self.nav_cont.setPosSize(0, -ev.value.Value,0,0,2)
        
    def disposing(self,ev):
        return False

    


  
