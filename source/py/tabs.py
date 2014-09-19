# -*- coding: utf-8 -*-

import unohelper



class Tabs():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        global pd
        pd = mb.pd

        self.mb = mb
        
    
    def start(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        try:
            if not in_tab_einfuegen:
                self.berechne_ordinale_in_baum_und_tab(False)
                self.erzeuge_Fenster(False)    
            else:     
                self.berechne_ordinale_in_baum_und_tab(True)
                self.erzeuge_Fenster(True)    
        except:
            if self.mb.debug: log(inspect.stack,tb())
    
    def berechne_ordinale_in_baum_und_tab(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        if in_tab_einfuegen:
            tab = 'Projekt'
            
            tree_tab = self.mb.props[T.AB].xml_tree
            root_tab = tree_tab.getroot()
            
            baum_tab = []
            self.mb.class_XML.get_tree_info(root_tab,baum_tab)
            
            self.im_tab_vorhandene = []
            
            # Ordinale eintragen
            for t in baum_tab:
                self.im_tab_vorhandene.append(t[0])
            
        else:
            tab = T.AB
            self.im_tab_vorhandene = []
            
            
        tree = self.mb.props[tab].xml_tree
        root = tree.getroot()
        
        self.baum = []
        self.mb.class_XML.get_tree_info(root,self.baum)
        
    
    def berechne_ordinal_nach_auswahl(self,in_tab_einfuegen):    
        if self.mb.debug: log(inspect.stack)
        
        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        
        ordinale = []
        ordinale_seitenleiste = []
        ordinale_baumansicht = []
        ordinale_auswahl = []
        

       
        if tab_auswahl.eigene_auswahl_use == 1:
            ordinale_auswahl = self.get_ordinale_eigene_auswahl()
        if tab_auswahl.seitenleiste_use == 1:
            ordinale_seitenleiste = self.get_ordinale_seitenleiste(in_tab_einfuegen)
        if tab_auswahl.baumansicht_use == 1:
            ordinale_baumansicht = self.get_ordinale_baumansicht(in_tab_einfuegen)
        if tab_auswahl.suche_use == 1:
            pass
        
        
        ############# LOGIK ###############
        # Alle Ordinale in list eintragen
        
        for ordi in ordinale_auswahl:
            if ordi not in ordinale:
                ordinale.append(ordi)
        for ordi in ordinale_seitenleiste:
            if ordi not in ordinale:
                ordinale.append(ordi)
        for ordi in ordinale_baumansicht:
            if ordi not in ordinale:
                ordinale.append(ordi)
        
        helfer = ordinale[:]
        
        # Wenn Logik auf UND steht, alle im entsprechenden Tag 
        # nicht vorhandeneOrdinale wieder loeschen  
        if tab_auswahl.seitenleiste_log != 'V':
            for ordin in ordinale:
                if ordin not in ordinale_seitenleiste:
                    helfer.remove(ordin)
        if tab_auswahl.baumansicht_log != 'V':
            for ordi in ordinale:
                if ordi not in ordinale_baumansicht:
                    helfer.remove(ordi)
        
        ordinale = helfer
                         
        if len(ordinale) == 0:
            return []
         
        # Ordinale sortieren 
        ordinale = self.sortiere_ordinale(ordinale)
        if tab_auswahl.zeitlich_anordnen == 1:
            ordinale = self.sortiere_ordinale_zeitlich(ordinale,tab_auswahl)
       
        return ordinale
        
    
    def erzeuge_neuen_tab(self,ordinale):
        if self.mb.debug: log(inspect.stack)

        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        
        try:
            
            tab_name = tab_auswahl.tab_name
            T.AB = tab_name
             
            self.erzeuge_props(tab_name)
            Eintraege = self.erzeuge_Eintraege(tab_name,ordinale)        
 
            self.mb.tab_umgeschaltet = True
            self.mb.tab_id_old = self.mb.active_tab_id
            win,tab_id = self.get_tab(tab_name)
 
            self.mb.tabs.update({tab_id:(win,tab_name)})
 
            self.mb.erzeuge_Menu(win)
 
            self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
 
            self.setze_selektierte_zeile('nr0')
            self.mb.class_Baumansicht.korrigiere_scrollbar()
            
            # zur Sicherung, damit der projekt xml nicht ueberschrieben wrd
            if T.AB == 'Projekt' or self.mb.active_tab_id == 1:
                self.mb.Mitteilungen.nachricht('ERROR',"warningbox",16777216)
                return
             
            tree = self.mb.props[T.AB].xml_tree
            Path = os.path.join(self.mb.pfade['tabs'] , T.AB +'.xml' )
            self.mb.tree_write(tree,Path)
        except:
            if self.mb.debug: log(inspect.stack,tb())

    def fuege_ausgewaehlte_in_tab_ein(self,ordinale):
        if self.mb.debug: log(inspect.stack)

        tab_name = copy.deepcopy(T.AB)
        tab_xml = copy.deepcopy(self.mb.props[T.AB].xml_tree)
        ord_selektierter = self.mb.props[T.AB].selektierte_zeile.AccessibleName
        
        self.schliesse_Tab(False)
     
        try:
            
            T.AB = tab_name
             
            self.erzeuge_props(tab_name)

            Eintraege = self.erzeuge_neue_Eintraege_im_tab(tab_name,ordinale,tab_xml, ord_selektierter)        

            self.mb.tab_umgeschaltet = True
            self.mb.tab_id_old = self.mb.active_tab_id
            win,tab_id = self.get_tab(tab_name)
 
            self.mb.tabs.update({tab_id:(win,tab_name)})
 
            self.mb.erzeuge_Menu(win)
            self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
            self.setze_selektierte_zeile('nr0')
            self.mb.class_Baumansicht.korrigiere_scrollbar()
            
            tree = self.mb.props[T.AB].xml_tree
            Path = os.path.join(self.mb.pfade['tabs'] , T.AB +'.xml' )
            self.mb.tree_write(tree,Path)

            self.mb.write_tab = False
        except:
            if self.mb.debug: log(inspect.stack,tb())
 
        
    def lade_tabs(self):
        if self.mb.debug: log(inspect.stack)
                                   
        try:
            
            gespeicherte_tabs = self.get_gespeicherte_tabs()
            
            for tab_name in gespeicherte_tabs:

                T.AB = tab_name
                
                self.erzeuge_props(tab_name)
                Eintraege = self.lade_tab_Eintraege(tab_name)        
            
                self.mb.tab_umgeschaltet = True
                self.mb.tab_id_old = self.mb.active_tab_id
                win,tab_id = self.get_tab(tab_name)
    
                self.mb.tabs.update({tab_id:(win,tab_name)})
                
                self.mb.erzeuge_Menu(win)
                self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
                
                self.setze_selektierte_zeile('nr0')
                self.mb.class_Baumansicht.korrigiere_scrollbar()
            
        except:
            if self.mb.debug: log(inspect.stack,tb())

    
    def lade_tab_Eintraege(self,tab_name):
        if self.mb.debug: log(inspect.stack)

        pfad = os.path.join(self.mb.pfade['tabs'], tab_name+'.xml')      
        self.mb.props[tab_name].xml_tree = self.mb.ET.parse(pfad)
        root = self.mb.props[tab_name].xml_tree.getroot()

        self.mb.props[tab_name].kommender_Eintrag = int(root.attrib['kommender_Eintrag'])
        
        Elements = root.findall('.//')       
        Eintraege = []
        
        for elem in Elements:
             
            ordinal = elem.tag
            parent  = elem.attrib['Parent']
            name    = elem.attrib['Name']
            art     = elem.attrib['Art']
            lvl     = elem.attrib['Lvl'] 
            zustand = elem.attrib['Zustand'] 
            sicht   = elem.attrib['Sicht'] 
            tag1   = elem.attrib['Tag1'] 
            tag2   = elem.attrib['Tag2'] 
            tag3   = elem.attrib['Tag3'] 
            
            Eintraege.append((ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3))
            
        return Eintraege
    
    def get_gespeicherte_tabs(self):
        if self.mb.debug: log(inspect.stack)
        
        tab_ordner = self.mb.pfade['tabs']
        tab_names = []
        
        for root, dirs, files in os.walk(tab_ordner):
            for file in files:
                tab_names.append(file.split('.xml')[0])

        return tab_names
    
    def setze_selektierte_zeile(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)        
        self.mb.props[T.AB].selektierte_zeile = zeile.AccessibleContext
    
    def get_ordinale_seitenleiste(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        try:

            tab_auswahl = self.mb.props[T.AB].tab_auswahl
            selektierte_tags = tab_auswahl.seitenleiste_tags.split(', ')
            
            ordinale = []
            
            if len(selektierte_tags) == 0:
                return ordinale
            
            alle_tag_eintraege = self.mb.dict_sb_content['ordinal']

            if tab_auswahl.seitenleiste_log_tags == 'V':
                # Alle Eintraege
                UND = False
            else:
                # Nur Eintraege, in denen alle Tags vorkommen
                UND = True
                        
            for tag_eintrag in alle_tag_eintraege:
                if in_tab_einfuegen:
                    if tag_eintrag  in self.im_tab_vorhandene:
                        continue
                    
                tag_ist_drin = False
                s_tag_ist_drin = []
                
                for s_tag in selektierte_tags:
                    
                    if s_tag in alle_tag_eintraege[tag_eintrag]['Tags_general']:
                        tag_ist_drin = True
                        s_tag_ist_drin.append(True)
                    else:
                        s_tag_ist_drin.append(False)
                        
                    if UND == True:
                        if False in s_tag_ist_drin:
                            tag_ist_drin = False
                    else:    
                        if tag_ist_drin:
                            break
                        
                if tag_ist_drin:
                    ordinale.append(tag_eintrag)
        except:
            if self.mb.debug: log(inspect.stack,tb())

        return sorted(ordinale)
    
    def get_ordinale_baumansicht(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        ordinale = []
        
        if in_tab_einfuegen:
            tab = 'Projekt'
        else:
            tab = T.AB
        
        tree = self.mb.props[tab].xml_tree
        root = tree.getroot()
        all_el = root.findall('.//')

        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        ausgew_icons = tab_auswahl.baumansicht_tags
        
        for eintrag in all_el:
            if in_tab_einfuegen:
                if eintrag.tag in self.im_tab_vorhandene:
                    continue
            if eintrag.attrib['Tag1'] in ausgew_icons:
                ordinale.append(eintrag.tag)

        return ordinale
    
    def get_ordinale_eigene_auswahl(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinale = []

        keys = self.mb.settings_exp['ausgewaehlte'].keys()
        for key in keys:
            if self.mb.settings_exp['ausgewaehlte'][key][1] == 1:
                ordinale.append(key)

        return ordinale
    
    
    def erzeuge_Fenster(self,in_tab_einfuegen = False):
        if self.mb.debug: log(inspect.stack)

        LANG = self.mb.lang
        
        if in_tab_einfuegen:
            hoehe = 640
        else:
            hoehe = 690
        
        
        posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
        posSize = (posSize_main.X + 20, posSize_main.Y + 20,310,hoehe)
        win,cont = self.mb.erzeuge_Dialog_Container(posSize)
        
        button_listener = Auswahl_Button_Listener(self.mb,win,cont,in_tab_einfuegen)

        x1 = 10
        x2 = 280
        x3 = 10
        x4 = 20
        x5 = 140
        
        width = 120
        width2 = 80
        
        y = 20
        
        
        if in_tab_einfuegen:
            Ueberschrift = LANG.IMPORTIERE_IN_TAB
        else:
            Ueberschrift = LANG.ERZEUGE_NEUEN_TAB_AUS
        
        prop_names = ('Label',)
        prop_values = (Ueberschrift,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x1, y, 200, 20, prop_names, prop_values)  
        cont.addControl('Titel', control)
        model.FontWeight = 200.0
        
        y += 20
        
        
        prop_names = ('Label',)
        prop_values = (LANG.MEHRFACHE_AUSWAHL,1)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x1, y, 200, 20, prop_names, prop_values)  
        cont.addControl('R1', control)

        
        #################### EIGENE AUSWAHL ########################
        y += 40
        
        prop_names = ('Label','State')
        prop_values = (LANG.EIGENE_AUSWAHL,0)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x3, y, width, 20, prop_names, prop_values)  
        cont.addControl('Eigene_Auswahl_use', control)
        #control.Enable = False
        
        
        prop_names = ('Label',)
        prop_values = (LANG.AUSWAHL,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x5, y, width2, 20, prop_names, prop_values)  
        cont.addControl('Eigene_Auswahl', control)
        control.addActionListener(button_listener) 
        control.setActionCommand('Eigene Auswahl')
        
        y += 30
        
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x4, y, 210, 20, (), ())  
        cont.addControl('Line', control)
        
        #####################




        
        #################### TAGS SEITENLEISTE ########################
        y += 30
        
        prop_names = ('Label',)
        prop_values = (u'V',)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x2, y- 2, 16, 16, prop_names, prop_values)  
        model.HelpText = LANG.TAB_HELP_TEXT_NOT_IMPLEMENTED
        control.addActionListener(button_listener) 
        control.setActionCommand('V')
        cont.addControl('but1_Seitenleiste', control)
        
        
        
        prop_names = ('Label',)
        prop_values = (LANG.TAGS_SEITENLEISTE,)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x3, y, width, 20, prop_names, prop_values)  
        cont.addControl('CB_Seitenleiste', control)
        
        
        prop_names = ('Label',)
        prop_values = (LANG.AUSWAHL,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x5, y, width2, 20, prop_names, prop_values) 
        control.addActionListener(button_listener) 
        control.setActionCommand('Tags Seitenleiste')
        cont.addControl('but2_Seitenleiste', control)
        
        y += 30
        
        prop_names = ('Label',)
        prop_values = (u'V',)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x4, y- 2, 16, 16, prop_names, prop_values)  
        model.HelpText = LANG.TAB_HELP_TEXT
        control.addActionListener(button_listener) 
        control.setActionCommand('V')
        cont.addControl('but3_Seitenleiste', control)
        
        
        prop_names = ('Label',)
        prop_values = ('',)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x4 + 20, y, 200, 20, prop_names, prop_values)  
        cont.addControl('txt_Seitenleiste', control)
        
        
        
        y += 30
        
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x4, y, 210, 20, (), ())  
        cont.addControl('Line', control)
                
        ####################
        
        #################### TAGS BAUMANSICHT ########################
        y += 30
        
        prop_names = ('Label',)
        prop_values = (u'V',)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x2, y- 2, 16, 16, prop_names, prop_values)  
        model.HelpText = LANG.TAB_HELP_TEXT_NOT_IMPLEMENTED
        control.addActionListener(button_listener) 
        control.setActionCommand('V')
        cont.addControl('but1_Baumansicht', control)
        
        
        
        prop_names = ('Label',)
        prop_values = (LANG.TAGS_BAUMANSICHT,)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x3, y, width, 20, prop_names, prop_values)  
        cont.addControl('CB_Baumansicht', control)
        control.Enable = True
        
        
        prop_names = ('Label',)
        prop_values = (LANG.AUSWAHL,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x5, y, width2, 20, prop_names, prop_values) 
        control.addActionListener(button_listener) 
        control.setActionCommand('Tags Baumansicht')
        cont.addControl('but2_Baumansicht', control)
        
        y += 30        
        
        control, model = self.mb.createControl(self.mb.ctx, "Container", x4, y, 200, 20, (), ())  
        cont.addControl('icons_Baumansicht', control)
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        
        
        y += 30
        
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x4, y, 210, 20, (), ())  
        cont.addControl('Line', control)
                
        ####################
        
        #################### SUCHE ########################
        y += 30
        
        prop_names = ('Label',)
        prop_values = (u'V',)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x2, y- 2, 16, 16, prop_names, prop_values)  
        model.HelpText = LANG.TAB_HELP_TEXT_NOT_IMPLEMENTED
        control.addActionListener(button_listener) 
        control.setActionCommand('V')
        cont.addControl('but1_Suche', control)
        
        
        
        prop_names = ('Label',)
        prop_values = (LANG.SUCHE,)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x3, y, width, 20, prop_names, prop_values)  
        cont.addControl('CB_Suche', control)
        control.Enable = False
        
        
        prop_names = ('Label',)
        prop_values = (LANG.AUSWAHL,)
        control, model = self.mb.createControl(self.mb.ctx, "Edit", x5, y, width2, 20, prop_names, prop_values) 
        control.addTextListener(button_listener) 
        cont.addControl('edit_Suche', control)
        
        y += 30
        
        prop_names = ('Label',)
        prop_values = ('',)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x4, y, 200, 20, prop_names, prop_values)  
        cont.addControl('txt_Suche', control)
        
        
        
        y += 30
        
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x4, y, 210, 20, (), ())  
        cont.addControl('Line', control)
                
        ####################

                
        y += 50
        
        prop_names = ('Label',)
        prop_values = (LANG.ZEITLICH_ANORDNEN,)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x3, y, 290, 20, prop_names, prop_values)  
        cont.addControl('Zeit', control)
        
        y += 20
        
        prop_names = ('Label',)
        prop_values = (LANG.NUTZE_ZEIT,)
        control, model = self.mb.createControl(self.mb.ctx, "RadioButton", x4 + 7, y, 290, 20, prop_names, prop_values)  
        cont.addControl('z1', control)
        control.State = 1
        
        y += 20
        
        prop_names = ('Label',)
        prop_values = (LANG.NUTZE_DATUM,)
        control, model = self.mb.createControl(self.mb.ctx, "RadioButton", x4 + 7, y, 290, 20, prop_names, prop_values)  
        cont.addControl('z2', control)
        
        y += 20
        
        prop_names = ('Label',)
        prop_values = (LANG.NUTZE_ZEIT_UND_DATUM,)
        control, model = self.mb.createControl(self.mb.ctx, "RadioButton", x4 + 7, y, 290, 20, prop_names, prop_values)  
        cont.addControl('z3', control)
        
        y += 20
        
        prop_names = ('Label',)
        prop_values = (LANG.UNAUSGEZEICHNETE,)
        control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x4 + 7, y, 290, 20, prop_names, prop_values)  
        cont.addControl('Zeit2', control)
        
        

   
        y += 30
        
        ###########################  TRENNER #####################################################
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x1, y, 290, 20, (), ())  
        cont.addControl('Line', control)
        
        
        if not in_tab_einfuegen:
            y += 30
            
            prop_names = ('Label',)
            prop_values = (LANG.TABNAME,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", x1, y, 80, 20, prop_names, prop_values)  
            cont.addControl('tab_name_eingabe', control)
            
            y += 20
            
            prop_names = ('Text',)
            prop_values = ('Tab %s' %str(len(self.mb.tabs)+1),)
            control, model = self.mb.createControl(self.mb.ctx, "Edit", x3, y, width2, 20, prop_names, prop_values) 
            cont.addControl('tab_name', control)
  
        y += 50
        
        prop_names = ('Label',)
        prop_values = (LANG.OK,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 210, y, width2, 30, prop_names, prop_values)  
        control.addActionListener(button_listener)  
        if in_tab_einfuegen:
            control.setActionCommand('ok_tab')
        else:          
            control.setActionCommand('ok')
        cont.addControl('but_ok', control)
        
        
    
    def erzeuge_props(self,tab_name):
        if self.mb.debug: log(inspect.stack)
        
        x = Props()
        self.mb.props.update({tab_name :x})
        
        # erzeuge neuen xml_tree
        et = ElementTree
        root = et.Element('Tabs')
        
        tree = et.ElementTree(root)
        self.mb.props[tab_name].xml_tree = tree
        
        root.attrib['Name'] = 'root'
        root.attrib['Programmversion'] = self.mb.programm_version
        
        
 
    def erzeuge_Eintraege(self,tab_name,ordinale):
        if self.mb.debug: log(inspect.stack)
        
        # hier sollen die Ergebnisse von Suche oder Tags erzeugt werden
        
#         parent_eintrag = ('nr0','root',self.mb.projekt_name,0,'prj','auf','ja','leer','leer','leer')
#         papierkorb = ('nr5','root',lang.PAPIERKORB,0,'waste','zu','ja','leer','leer','leer')
                
        xml_tree = self.mb.props['Projekt'].xml_tree
        root = xml_tree.getroot()

        Eintraege = []
        
        # Ordinal des Projekts muss enthalten sein und an erster Stelle stehen 
        if 'nr0' not in ordinale:
            ordinale.insert(0,'nr0')
        else:
            ordinale.remove('nr0')
            ordinale.insert(0,'nr0')
        
        
        papierkorb = self.mb.props['Projekt'].Papierkorb
        ordinale.append(papierkorb)
                
        for ordi in ordinale:
            elem = root.find('.//'+ordi)

            ordinal = elem.tag
            name    = elem.attrib['Name']
            art     = elem.attrib['Art']
            if ordinal not in ('nr0',papierkorb):
                lvl     = 1 
                parent  = 'nr0'
            else:
                lvl = 0 #elem.attrib['Lvl'] 
                parent  = 'root'
            zustand = elem.attrib['Zustand'] 
            sicht   = 'ja' 
            tag1   = elem.attrib['Tag1'] 
            tag2   = elem.attrib['Tag2'] 
            tag3   = elem.attrib['Tag3'] 

            eintrag = (ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3)
            Eintraege.append(eintrag)      
            self.erzeuge_tab_XML_Eintrag(eintrag,tab_name)

        return Eintraege
        
    def erzeuge_neue_Eintraege_im_tab(self,tab_name,neue_ordinale,tab_xml, ord_selekt):
        if self.mb.debug: log(inspect.stack)        
             
        xml_tree = self.mb.props['Projekt'].xml_tree
        root = xml_tree.getroot()
        root_tab = tab_xml.getroot()
        
        ordinale = []
        self.mb.class_XML.get_tree_info(root_tab,ordinale)
        
        ord_selektierter = ord_selekt
        selekt_xml = root_tab.find('.//'+ord_selekt)
        ziel_xml = selekt_xml
        
        # wenn ord_selektierter ein Ordner ist,
        # letzten Kindeintrag suchen und ord_selektierter neu setzen
        suche = False
        for ordn in ordinale:
            if ordn[0] == ord_selektierter:
                childs = list(selekt_xml)
                # selektiert das letzte Kind eine Ebene tiefer,
                # wenn selektierter ein Ordner ist
                if childs != []:
                    ziel_xml = childs[-1]

                alle_Kinder = []
                self.mb.class_XML.get_tree_info(selekt_xml,alle_Kinder)
                # Selektiert das allerletzte Kind aller Unterordnereintraege
                ord_selektierter = alle_Kinder[len(alle_Kinder)-1][0]
                
                break

        Eintraege = []

        for ordi in ordinale:
            
            elem = root_tab.find('.//'+ordi[0])
        
            ordinal = elem.tag
            name    = elem.attrib['Name']
            art     = elem.attrib['Art']
            lvl     = elem.attrib['Lvl']
            parent  = elem.attrib['Parent']
            zustand = elem.attrib['Zustand'] 
            sicht   = elem.attrib['Sicht']
            tag1   = elem.attrib['Tag1'] 
            tag2   = elem.attrib['Tag2'] 
            tag3   = elem.attrib['Tag3'] 
            
            eintrag = (ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3)
            Eintraege.append(eintrag)      
            self.erzeuge_tab_XML_Eintrag(eintrag,tab_name)
            
            if ordi[0] == ord_selektierter:
                    
                for o in neue_ordinale:
            
                    elem2 = root.find('.//'+ o)

                    ordinal = elem2.tag
                    name    = elem2.attrib['Name']
                    art     = elem2.attrib['Art']
                    lvl     = ziel_xml.attrib['Lvl']
                    parent  = ziel_xml.attrib['Parent']
                    zustand = elem2.attrib['Zustand'] 
                    sicht   = 'ja'
                    tag1   = elem2.attrib['Tag1'] 
                    tag2   = elem2.attrib['Tag2'] 
                    tag3   = elem2.attrib['Tag3'] 
                    
                    eintrag = (ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3)
                    Eintraege.append(eintrag)      
                    self.erzeuge_tab_XML_Eintrag(eintrag,tab_name,parent,lvl)
                

        return Eintraege
        
    
    def sortiere_ordinale(self,ordinale):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['Projekt']
        root = props.xml_tree.getroot()
        all_ordinals = [elem.tag for elem in root.iter()]

        sorted_ordinals = []
        
        for ordn in all_ordinals:
            if ordn in ordinale:
                sorted_ordinals.append(ordn)

        return sorted_ordinals
    
    
    def sortiere_ordinale_zeitlich(self,ordinale,tab_auswahl):
        if self.mb.debug: log(inspect.stack)
      
        dict_sb = self.mb.dict_sb_content['ordinal']
        
        if tab_auswahl.kein_tag_einbeziehen == 1: 
            nutze_alle = True
        else:
            nutze_alle = False
        
        
        def berechne_ords(attribut):
            i = 0
            list_zeit = []
            for ordi in ordinale:
                zeit = dict_sb[ordi]['Tags_time'][attribut]
                if zeit == None:
                    if not nutze_alle:
                        continue
                    zeit = 'None'+str(i)
                    i += 1
                list_zeit.append((zeit,ordi))
            
            return sorted(list_zeit)
        
        
        if tab_auswahl.nutze_zeit == 1:                
            sortierte_liste = berechne_ords('zeit')
            
        elif tab_auswahl.nutze_datum == 1:
            sortierte_liste = berechne_ords('datum')
            
        else:
            i = 0
            list_zeit = []
            
            for ordi in ordinale:
                zeit = dict_sb[ordi]['Tags_time']['zeit']
                datum = dict_sb[ordi]['Tags_time']['datum']
                
                # Wenn d + t gewaehlt wurde, sollte zumindest ein Datum angegeben worden sein
                # Wenn nicht, wird das Ordinal hier entfernt
                if not nutze_alle:
                    if datum == None:
                        continue
                
                if zeit == None:
                    zeit = str(23599999)
                elif zeit == 0:
                    zeit = '00000000'
                    
                zeit = str(zeit)[:5]
                
                
                if datum == None:
                    datum = 99991231
                    
                datum = str(datum)
                datumzeit = int(datum+zeit)
                list_zeit.append((datumzeit,ordi))
                
            sortierte_liste = sorted(list_zeit)

        sort_list = list(x[1] for x in sortierte_liste)
        return sort_list
        
        
    def erzeuge_tab_XML_Eintrag(self,eintrag,tab_name,parent_neu = None,lvl_neu = None):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[tab_name].xml_tree
        root = tree.getroot()
        et = self.mb.ET             
        ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
        
        if parent_neu != None:
            par = root.find('.//'+parent_neu)
            lvl = lvl_neu
            sicht = 'ja'
            zustand = 'auf'
        elif parent == 'root':
            par = root
        elif parent == 'Tabs':
            par = root
        else:
            par = root.find('.//'+parent)
        

        el = et.SubElement(par,ordinal)
        el.attrib['Parent'] = parent
        el.attrib['Name'] = name
        el.attrib['Art'] = art
        
        el.attrib['Lvl'] = str(lvl)
        el.attrib['Zustand'] = zustand
        el.attrib['Sicht'] = sicht
        el.attrib['Tag1'] = tag1
        el.attrib['Tag2'] = tag2
        el.attrib['Tag3'] = tag3
                    
        self.mb.props[tab_name].kommender_Eintrag = int(self.mb.props[tab_name].kommender_Eintrag) + 1
        root.attrib['kommender_Eintrag'] = str(self.mb.props[tab_name].kommender_Eintrag)   

 
    
    def erzeuge_Hauptfeld(self,win,tab_name,Eintraege):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.props[tab_name].Hauptfeld = self.mb.class_Baumansicht.erzeuge_Feld_Baumansicht(win)  
            self.mb.class_Baumansicht.erzeuge_Scrollbar(win,self.mb.ctx)   

            self.erzeuge_Eintraege_und_Bereiche(Eintraege,tab_name)    
        except:
            if self.mb.debug: log(inspect.stack,tb())

        
    def get_tab(self,tab_name):
        if self.mb.debug: log(inspect.stack)
        
        try:

            tabsX = self.mb.tabsX
             
            from com.sun.star.beans import NamedValue
            dialog1 = "vnd.sun.star.extension://xaver.roemers.organon/factory/Dialog1.xdl"
            tab_id = tabsX.insertTab() # Create new tab, return value is tab id
            # Valid properties are: 
            # Title, ToolTip, PageURL, EventHdl, Image, Disabled.
            v1 = NamedValue("PageURL", dialog1)
            v2 = NamedValue("Title", tab_name)
            v3 = NamedValue("EventHdl", self.mb.factory.CWHandler)
            tabsX.setTabProps(tab_id, (v1, v2, v3))
            tabsX.activateTab(tab_id)             
              
            win = self.mb.factory.CWHandler.window2
            win.Model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD

            win.addWindowListener(self.mb.w_listener)
            
            self.setMenuBar(win,self.mb.ctx)

            return win,tab_id              
  
        except:
            if self.mb.debug: log(inspect.stack,tb())
            
    def setMenuBar(self,window,ctx):
        try:
            smgr = ctx.ServiceManager
            oMenuBar = smgr.createInstanceWithContext('com.sun.star.awt.MenuBar', ctx)
            oPopUp = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", ctx)
            oMenuBar.insertItem(0, "First0", 4, 0)
            oMenuBar.insertItem(1, "First1firlefanz", 4, 0)
            oMenuBar.insertItem(2, "First2", 4, 0)
            oMenuBar.insertItem(3, "First3zuckerhut", 4, 0)
            oMenuBar.insertItem(4, "First4", 4, 0)
            window.setMenuBar(oMenuBar)
        except:
            if self.mb.debug: log(inspect.stack,tb())
            
    def erzeuge_Eintraege_und_Bereiche(self,Eintraege,tab_name):
        if self.mb.debug: log(inspect.stack)        
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0

        for eintrag in Eintraege:
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                     
            index = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,self.mb.class_Zeilen_Listener,index,tab_name)

            if sicht == 'ja':
                # index wird in erzeuge_Zeile_in_der_Baumansicht bereits erhoeht, daher hier 1 abziehen
                self.mb.props[tab_name].dict_zeilen_posY.update({(index-1)*KONST.ZEILENHOEHE:eintrag})
                self.mb.props[tab_name].sichtbare_bereiche.append('OrganonSec'+str(index2))
                
            # Bereiche   
            inhalt = name
            path = os.path.join(self.mb.pfade['odts'],ordinal+'.odt') 
            
            Bereichsname_dict.update({'OrganonSec'+str(index2):path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
            
            index2 += 1
                    
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.props[T.AB].dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        self.mb.class_Projekt.erzeuge_dict_ordner()

 
    def entferne_alle_listener(self,win):
        if self.mb.debug: log(inspect.stack)
        
        #return
        win.removeWindowListener(self.mb.w_listener)
#         self.listener.dispose()
#         self.listener2.dispose()
#         self.current_Contr.removeSelectionChangeListener(self.VC_selection_listener) 
#         self.current_Contr.removeKeyHandler(self.keyhandler)
#         win.removeWindowListener(self.w_listener)
#         self.undo_mgr.removeUndoManagerListener(self.undo_mgr_listener)
        
    def schliesse_Tab(self,abfrage = True):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if abfrage:
                # Frage: Soll Tab wirklich geschlossen werden?
                entscheidung = self.mb.Mitteilungen.nachricht(self.mb.lang.TAB_SCHLIESSEN %T.AB ,"warningbox",16777216)
                # 3 = Nein oder Cancel, 2 = Ja
                if entscheidung == 3:
                    return
                #print('active tab id', self.mb.active_tab_id,self.mb.tabs[self.mb.active_tab_id][1])
                if T.AB == 'Projekt':
                    self.mb.Mitteilungen.nachricht("Project tab can't be closed" ,"warningbox",16777216)
                    return
            
            # loesche tab listener
            win = self.mb.tabs[self.mb.active_tab_id][0]
            self.entferne_alle_listener(win)
            
            # loesche tab.xml
            tab_name = self.mb.tabs[self.mb.active_tab_id][1]
            Path = os.path.join(self.mb.pfade['tabs'], '%s.xml' % tab_name)

            os.remove(Path)
            
            # loesche Tab
            self.mb.tabsX.removeTab(self.mb.active_tab_id)
                        
            # loesche props[tab]
            del self.mb.props[T.AB]
            
            T.AB = 'Projekt'
            
        except:
            if self.mb.debug: log(inspect.stack,tb())


 

from com.sun.star.awt import XActionListener,XTextListener
class Auswahl_Button_Listener(unohelper.Base, XActionListener,XTextListener):
    def __init__(self,mb,win,fenster_cont,in_tab_einfuegen = False):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = win
        self.fenster_cont = fenster_cont
        self.in_tab_einfuegen = in_tab_einfuegen
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        if ev.ActionCommand == 'Tags Seitenleiste':
            self.erzeuge_tag_auswahl_seitenleiste(ev)
        elif ev.ActionCommand == 'Tags Baumansicht':
            self.erzeuge_tag_auswahl_baumansicht(ev)
        elif ev.ActionCommand == 'Eigene Auswahl':
            self.erzeuge_eigene_auswahl(ev)
        elif ev.ActionCommand == 'ok':
            try:
                self.erstelle_auswahl_dict(ev)
                ok = self.pruefe_tab_namen()
                if ok:
                    self.win.dispose()
                    ordinale = self.mb.class_Tabs.berechne_ordinal_nach_auswahl(False)
                    if ordinale == []:
                        return
                    self.mb.class_Tabs.erzeuge_neuen_tab(ordinale)
            except:
                if self.mb.debug: log(inspect.stack,tb())
                
        elif ev.ActionCommand == 'ok_tab':
            try:
                self.erstelle_auswahl_dict(ev)
                self.win.dispose()
                ordinale = self.mb.class_Tabs.berechne_ordinal_nach_auswahl(True)
                #ordinale =  self.schliesse_vorhandene_ordinale_aus(ordin)
                if ordinale == []:
                    return
                self.mb.class_Tabs.fuege_ausgewaehlte_in_tab_ein(ordinale)
            except:
                if self.mb.debug: log(inspect.stack,tb())
                
        elif ev.ActionCommand == 'V':
            if ev.Source.Model.Label == 'V':
                ev.Source.Model.Label = u'\u039B'
            else:
                ev.Source.Model.Label = 'V'
    
    
    def textChanged(self,ev):
        main_win = ev.Source.Context
        main_win.getControl('txt_Suche').Model.Label = ev.Source.Text
        
    def get_fenster_position(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        par_win = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext
        x = par_win.LocationOnScreen.X + par_win.Size.Width + 20
        y = par_win.LocationOnScreen.Y
        return x,y        
        
    def erzeuge_tag_auswahl_baumansicht(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        x,y = self.get_fenster_position(ev)
        posSize = (x,y,220,400)
        
        win,cont = self.mb.erzeuge_Dialog_Container(posSize)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.AUSGEWAEHLTE,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 100, 20, prop_names, prop_values) 
        model.FontWeight = 200.0 
        cont.addControl('ausgewaehlte_XXX', control)
        
        lb_alle_icons,lb_ausgew_icons = self.erzeuge_ListBox_Tag1()
        
        tag_item_listener = Tag1_Listener(self.mb,self.fenster_cont,lb_alle_icons,lb_ausgew_icons)
        lb_alle_icons.addItemListener(tag_item_listener)
        lb_ausgew_icons.addItemListener(tag_item_listener)
        
        cont.addControl('Eintraege_Tag1', lb_alle_icons)
        cont.addControl('Ausgewaehlte_Tag1', lb_ausgew_icons)
            
       
    def erzeuge_ListBox_Tag1(self):
        if self.mb.debug: log(inspect.stack)
        
        # alle Punkte
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 120, 10, KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control.setMultipleMode(False)
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        model.Border = 0
        
        items = ('blau',
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
        
        
        # ausgewaehlte Punkte
        control_ausgewaehlte, model = self.mb.createControl(self.mb.ctx, "ListBox", 10, 40, KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control.setMultipleMode(False)
        model.Border = 0
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        
        return control,control_ausgewaehlte
    
        
    def erzeuge_tag_auswahl_seitenleiste(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:

            x,y = self.get_fenster_position(ev)
            posSize = (x,y,970,400)
            
            win,cont = self.mb.erzeuge_Dialog_Container(posSize)
            
            tags = self.mb.dict_sb_content['tags']
            
            dict_panels = self.mb.class_Sidebar.sb_panels1
            ausgew_tags = 'Tags_characters','Tags_objects','Tags_locations','Tags_user1','Tags_user2','Tags_user3'
            
            prop_names = ('Label',)
            prop_values = (self.mb.lang.AUSGEWAEHLTE,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 100, 20, prop_names, prop_values) 
            model.FontWeight = 200.0 
            cont.addControl('ausgewaehlte_XXX', control)
            

            #Tags_general
            self.controls = []
            auswahl_listener = Auswahl_Tags_Listener(self.mb,win,cont,self.controls,ev.Source)
            alle_tags = tags['Tags_general'][:]

            x = 150
            width = 100

            for tag in ausgew_tags:

                prop_names = ('Label','Align')
                prop_values = (dict_panels[tag],1)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", x, 10, width, 20, prop_names, prop_values)  
                cont.addControl(dict_panels[tag], control)
                
                y = 0
                for t in tags[tag]:
                    prop_names = ('Label',)
                    prop_values = (t,)
                    control, model = self.mb.createControl(self.mb.ctx, "Button", x + 10, y + 30, width - 20, 20, prop_names, prop_values)  
                    cont.addControl(t, control)
                    control.setActionCommand(t)
                    control.addActionListener(auswahl_listener)
                    
                    if t in alle_tags:
                        alle_tags.remove(t)
                    
                    y += 25
                
                
                x += (width + 10)
                
            ############## TAGS ALLGEMEIN #####################
            prop_names = ('Label','Align')
            prop_values = (dict_panels['Tags_general'],1)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", x, 10, width, 20, prop_names, prop_values)  
            cont.addControl(dict_panels['Tags_general'], control)
            
            y = 0
            for t in alle_tags:
                prop_names = ('Label',)
                prop_values = (t,)
                control, model = self.mb.createControl(self.mb.ctx, "Button", x + 10, y + 30, width - 20, 20, prop_names, prop_values)  
                cont.addControl(t, control)
                control.setActionCommand(t)
                control.addActionListener(auswahl_listener)
                
                y += 25
            
        except:
            if self.mb.debug: log(inspect.stack,tb())
        
    def erstelle_auswahl_dict(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        main_win = ev.Source.Context

        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        
        tab_auswahl.eigene_auswahl = None
        tab_auswahl.eigene_auswahl_use = main_win.getControl('Eigene_Auswahl_use').State
        
        tab_auswahl.seitenleiste_use = main_win.getControl('CB_Seitenleiste').State
        tab_auswahl.seitenleiste_log = main_win.getControl('but1_Seitenleiste').Model.Label
        tab_auswahl.seitenleiste_log_tags = main_win.getControl('but3_Seitenleiste').Model.Label
        tab_auswahl.seitenleiste_tags = main_win.getControl('txt_Seitenleiste').Model.Label
        
        tab_auswahl.baumansicht_use = main_win.getControl('CB_Baumansicht').State
        tab_auswahl.baumansicht_log = main_win.getControl('but1_Baumansicht').Model.Label
        tab_auswahl.baumansicht_tags = self.get_baumansicht_icons()
        
        tab_auswahl.suche_use = main_win.getControl('CB_Suche').State
        tab_auswahl.suche_log = main_win.getControl('but1_Suche').Model.Label
        tab_auswahl.suche_term = main_win.getControl('txt_Suche').Model.Label

        tab_auswahl.zeitlich_anordnen = main_win.getControl('Zeit').State
        tab_auswahl.kein_tag_einbeziehen = main_win.getControl('Zeit2').State
        tab_auswahl.nutze_zeit = main_win.getControl('z1').State
        tab_auswahl.nutze_datum = main_win.getControl('z2').State
        tab_auswahl.nutze_zeit_und_datum = main_win.getControl('z3').State
        
        if self.in_tab_einfuegen:
            tab_auswahl.tab_name = T.AB
        else:
            tab_auswahl.tab_name = main_win.getControl('tab_name').Model.Text
        
    
    def pruefe_tab_namen(self):
        if self.mb.debug: log(inspect.stack)
        
        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        tab_name = tab_auswahl.tab_name
        
        path = os.path.join(self.mb.pfade['tabs'],tab_name+'.xml')

        if os.path.exists(path):
            self.mb.Mitteilungen.nachricht(self.mb.lang.TAB_EXISTIERT_SCHON,'infobox')
            return False
        
        else:
            return True
    
    def get_baumansicht_icons(self):
        if self.mb.debug: log(inspect.stack)
        
        container = self.fenster_cont.getControl('icons_Baumansicht')
        ausgew_icons = []
        
        for cont in container.Controls:
            name = cont.Model.ImageURL.split('punkt_')[1]
            name = name.replace('.png','')
            ausgew_icons.append(name)

        return ausgew_icons
    
    #################
    
    
    def erzeuge_eigene_auswahl(self,ev):
        if self.mb.debug: log(inspect.stack)
        # das Fenster fuer die eigene Auswahl entspricht fast dem fuer die Auswahl beim Export.
        # -> der Code ist daher doppelt vorhanden und koennte vereinfacht werden
        #
        # umfasst:
        # erzeuge_Fenster_fuer_eigene_Auswahl
        # setze_hoehe_und_scrollbalken
        # erzeuge_auswahl
        #
        # und die Listener: 
        # Auswahl_ScrollBar_Listener
        # Auswahl_CheckBox_Listener
        
        self.erzeuge_Fenster_fuer_eigene_Auswahl(ev)
        
        
    
    def erzeuge_Fenster_fuer_eigene_Auswahl(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        x,y = self.get_fenster_position(ev)
        posSize = x,y,400,600
        
        lang = self.mb.lang
        sett = self.mb.settings_exp
 
        # Dict von alten Eintraegen bereinigen
        eintr = []
        for ordinal in sett['ausgewaehlte']:
            if self.in_tab_einfuegen:
                eintr.append(ordinal)
            else:
                if ordinal not in self.mb.props[T.AB].dict_bereiche['ordinal']:
                    eintr.append(ordinal)
                    
        for ordn in eintr:
            del sett['ausgewaehlte'][ordn]

        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)

        fenster_cont.Model.Text = lang.AUSWAHL
        fenster_cont.Model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE        
        
        control_innen, model = self.mb.createControl(self.mb.ctx,"Container",20,0,posSize[2],posSize[3],(),() )  
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        fenster_cont.addControl('Huelle', control_innen)
        
        y = self.erzeuge_treeview(control_innen)
        control_innen.setPosSize(0, 0,0,y + 20,8)

        self.setze_hoehe_und_scrollbalken(y,posSize[3],fenster,fenster_cont,control_innen)
    
      
    def setze_hoehe_und_scrollbalken(self,y,y_desk,fenster,fenster_cont,control_innen):  
        if self.mb.debug: log(inspect.stack)
        
        if y < y_desk-20:
            fenster.setPosSize(0,0,0,y + 20,8) 
        else:

            Attr = (0,0,20,y_desk,'scrollbar', None)    
            PosX,PosY,Width,Height,Name,Color = Attr
            
            # SCROLLBAR
            control, model = self.mb.createControl(self.mb.ctx,"ScrollBar",PosX,PosY,Width,Height,(),() )  
            #
            model.Orientation = 1
            model.BorderColor = KONST.EXPORT_DIALOG_FARBE
            model.Border = 0
            model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
            
            
            model.LiveScroll = True        
            #model.ScrollValueMax = y/2
            control.Maximum = y
            
            listener = Auswahl_ScrollBar_Listener(self.mb,control_innen)
            control.addAdjustmentListener(listener) 
            
            fenster_cont.addControl('ScrollBar',control)  
  
    def berechne_eigene_auswahl(self):
        if self.mb.debug: log(inspect.stack)
        
        if self.in_tab_einfuegen:
            tab = 'Projekt'
            
            tree_tab = self.mb.props[T.AB].xml_tree
            root_tab = tree_tab.getroot()
            
            baum_tab = []
            self.mb.class_XML.get_tree_info(root_tab,baum_tab)
            
            im_tab_vorhandene = []
            
            # Ordinale eintragen
            for t in baum_tab:
                im_tab_vorhandene.append(t[0])
            
        else:
            tab = T.AB
            im_tab_vorhandene = []
            
            
        tree = self.mb.props[tab].xml_tree
        root = tree.getroot()
        
        baum = []
        self.mb.class_XML.get_tree_info(root,baum)
        
        return baum,im_tab_vorhandene
        
    def erzeuge_treeview(self,fenster_cont):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sett = self.mb.settings_exp
            lang = self.mb.lang


            #baum,im_tab_vorhandene = self.berechne_eigene_auswahl()
            tabs = self.mb.class_Tabs
            baum = tabs.baum
            im_tab_vorhandene = tabs.im_tab_vorhandene

            y = 10
            x = 10
            
            listener = Auswahl_CheckBox_Listener(self.mb,fenster_cont)
            
            #Titel
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x,y ,300,20,(),() )  
            control.Text = lang.AUSWAHL_TIT
            model.FontWeight = 200.0
            fenster_cont.addControl('Titel', control)
            
            y += 30
            
            # Untereintraege auswaehlen
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40,y ,300,20,(),() )  
            control.Text = lang.ORDNER_CLICK
            model.FontWeight = 150.0
            fenster_cont.addControl('ausw', control)
            
            control, model = self.mb.createControl(self.mb.ctx,"CheckBox",x+20,y ,20,20,(),() )  
            control.State = sett['auswahl']
            control.ActionCommand = 'untereintraege_auswaehlen'
            control.addActionListener(listener)
            fenster_cont.addControl('Titel', control)
            
            y += 30

            for eintrag in baum:
    
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
                
                if art == 'waste':
                    break
                
                # Checkbox
                control, model = self.mb.createControl(self.mb.ctx,"CheckBox",x+20*int(lvl),y ,20,20,(),() )  
                control.ActionCommand = ordinal+'xxx'+name
                
                if ordinal in im_tab_vorhandene:
                    control.Enable = False
                    model.State = 0
                else:
                    control.addActionListener(listener)
                    if ordinal in sett['ausgewaehlte']:
                        model.State = sett['ausgewaehlte'][ordinal][1]
                        
                fenster_cont.addControl(ordinal, control)
                
                # Symbol
                control, model = self.mb.createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
                model.Border = False
                if art in ('dir','prj'):
                    model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
                else:
                    model.ImageURL = 'private:graphicrepository/res/sx03150.png' 
                fenster_cont.addControl('Titel', control)   
                if ordinal in im_tab_vorhandene:
                    control.Enable = False
                    
                # Name
                control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,200,20,(),() )  
                control.Text = name
                fenster_cont.addControl('Titel', control)
                if ordinal in im_tab_vorhandene:
                    control.Enable = False
                y += 20 
                
            return y 
        except:
            if self.mb.debug: log(inspect.stack,tb())
            


from com.sun.star.awt import XItemListener
class Tag1_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,win,lb_alle_icons,lb_ausgew_icons):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = win
        self.lb_alle_icons = lb_alle_icons
        self.lb_ausgew_icons = lb_ausgew_icons
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
        
        try: 
            container_baumansicht = self.win.getControl('icons_Baumansicht')
            item = ev.Source.Items[ev.Selected]
            

            if ev.Source == self.lb_alle_icons:
                
                if self.lb_ausgew_icons.ItemCount == 0:
                    Bedingung = True
                elif item not in self.lb_ausgew_icons.Items:
                    Bedingung = True
                else:
                    Bedingung = False
                    
                if Bedingung:
                    self.lb_ausgew_icons.addItem(item, 0)
                    
                    for it in self.lb_ausgew_icons.Items:
                        pos = self.lb_ausgew_icons.Items.index(it)       
                        self.lb_ausgew_icons.Model.setItemImage(pos,KONST.URL_IMGS+'punkt_%s.png' %it)
                        
            elif ev.Source == self.lb_ausgew_icons:
                pos = self.lb_ausgew_icons.Items.index(item) 
                self.lb_ausgew_icons.Model.removeItem(pos)
            
            
            container_controls = container_baumansicht.getControls()
            
            for con in container_controls:
                con.dispose()
            
            x = 0  
            for it in self.lb_ausgew_icons.Model.AllItems:
                
                prop_names = ('ImageURL','Border')
                prop_values = (it.Second,0)
                control, model = self.mb.createControl(self.mb.ctx, "ImageControl", x, 0, 20, 20, prop_names, prop_values) 
                container_baumansicht.addControl(it.First, control)
                x += 20
            
        except:
            if self.mb.debug: log(inspect.stack,tb())


        
class Auswahl_Tags_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,win,cont,controls,source):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = win
        self.cont = cont
        self.controls = controls
        self.source = source
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            txt_control = self.source.Context.getControl('txt_Seitenleiste')
                    
            
            if 'del' in ev.ActionCommand:
                tag = ev.ActionCommand.split('del')[1]

                for c in self.controls:
                    cont = self.cont.getControl(c+'button')
                    cont.dispose()
                    
                self.controls.remove(tag)
                    
                for c in self.controls:
                    self.erzeuge_button(c)
                
                txt_control.Model.Label = self.erzeuge_text()

            else:
                if ev.ActionCommand not in self.controls:
                    self.controls.append(ev.ActionCommand)
                    self.erzeuge_button(ev.ActionCommand)

                    txt_control.Model.Label = self.erzeuge_text()
                    
        except:
            if self.mb.debug: log(inspect.stack,tb())
            
    def erzeuge_button(self,ActionCommand):
        if self.mb.debug: log(inspect.stack)
        
        ind = self.controls.index(ActionCommand)
        y = ind*30
                    
        width = 100
        
        prop_names = ('Label',)
        prop_values = (ActionCommand,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 10, y + 50, width - 20, 20, prop_names, prop_values) 
        control.addActionListener(self)
        control.setActionCommand('del'+ActionCommand)
        self.cont.addControl(ActionCommand+'button', control)
    
    def erzeuge_text(self):
        if self.mb.debug: log(inspect.stack)
        
        text = ''
                    
        for t in self.controls:
            if self.controls.index(t) > 0:
                z = ', '
            else:
                z = ''
            text += z + t
            
        return text

from com.sun.star.awt import XMouseListener, XItemListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT 
     



from com.sun.star.awt import XAdjustmentListener
class Auswahl_ScrollBar_Listener (unohelper.Base,XAdjustmentListener):
    
    def __init__(self,mb,fenster_cont):  
        if mb.debug: log(inspect.stack)      
        self.mb = mb
        self.fenster_cont = fenster_cont
        
    def adjustmentValueChanged(self,ev):
        self.fenster_cont.setPosSize(0, -ev.value.Value,0,0,2)
        
    def disposing(self,ev):
        return False

            
            
class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,fenster_cont):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.fenster_cont = fenster_cont
    
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
            sett['ausgewaehlte'].update({ordinal:(titel,state)})

            if sett['auswahl']:
                if ordinal in self.mb.props[T.AB].dict_ordner:
                    
                    tree = self.mb.props[T.AB].xml_tree
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
                        if ordn != self.mb.props[T.AB].Papierkorb:
                            control = self.fenster_cont.getControl(ordn)
                            control.Model.State = state
                            zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordn)
                            titel = zeile.getControl('textfeld').Text
                            sett['ausgewaehlte'].update({ordn:(titel,state)}) 


    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:              
            return 1           
        
            
                       

             
     
