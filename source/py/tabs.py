# -*- coding: utf-8 -*-

import unohelper



class Tabs():
    
    def __init__(self,mb):

        global pd
        pd = mb.pd

        self.mb = mb
        
    
    def start(self):
        
        try:

            self.erzeuge_Fenster()            
            
        except:
            tb()

    
    def erzeuge_neuen_tab(self):
        if self.mb.debug: log(inspect.stack)
            
        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        
        ordinale = []
        ordinale_seitenleiste = []
        ordinale_baumansicht = []
        ordinale_auswahl = []
        
        try:
            if tab_auswahl.rb:
                if tab_auswahl.eigene_auswahl_use == 1:
                    ordinale_auswahl = self.get_ordinale_eigene_auswahl()
                if tab_auswahl.seitenleiste_use == 1:
                    ordinale_seitenleiste = self.get_ordinale_seitenleiste()
                if tab_auswahl.baumansicht_use == 1:
                    ordinale_baumansicht = self.get_ordinale_baumansicht()
                if tab_auswahl.suche_use == 1:
                    pass
            
            else:
                # TAB Date
                return
            
            ############# ToDo: LOGIK ###############
            UND = True
            
            if UND:
                for ordi in ordinale_auswahl:
                    if ordi not in ordinale:
                        ordinale.append(ordi)
                for ordi in ordinale_seitenleiste:
                    if ordi not in ordinale:
                        ordinale.append(ordi)
                for ordi in ordinale_baumansicht:
                    if ordi not in ordinale:
                        ordinale.append(ordi)
        

            if len(ordinale) == 0:
                return
            
            tab_name = tab_auswahl.tab_name
            T.AB = tab_name
            
            self.erzeuge_props(tab_name)
            Eintraege = self.erzeuge_Eintraege(tab_name,ordinale)        

            self.mb.tab_umgeschaltet = True
            self.mb.tab_id_old = self.mb.active_tab_id
            win,tab_id = self.get_tab(tab_name)

            self.mb.tabs.update({tab_id:(win,tab_name)})

            self.erzeuge_Menu(win,tab_id)

            self.erzeuge_Hauptfeld(win,tab_name,Eintraege)

            self.setze_selektierte_zeile('nr0')
            self.mb.class_Hauptfeld.korrigiere_scrollbar()
            
            if T.AB == 'Projekt' or self.mb.active_tab_id == 1:
                self.mb.Mitteilungen.nachricht('ERROR',"warningbox",16777216)
                return
            
            tree = self.mb.props[T.AB].xml_tree
            Path = os.path.join(self.mb.pfade['tabs'] , T.AB +'.xml' )
            tree.write(Path)
        except:
            tb()
    
    
    def ok(self):
        if T.AB == 'Projekt' or self.mb.active_tab_id == 1:
            self.mb.Mitteilungen.nachricht('ERROR',"warningbox",16777216)
            return False
        else:
            return True
    
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
                
                self.erzeuge_Menu(win,tab_id)
                self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
                
                self.setze_selektierte_zeile('nr0')
                self.mb.class_Hauptfeld.korrigiere_scrollbar()
            
        except:
            tb()

    
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
    
    def get_ordinale_seitenleiste(self):
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
            tb()

        return sorted(ordinale)
    
    def get_ordinale_baumansicht(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinale = []
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        all_el = root.findall('.//')

        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        ausgew_icons = tab_auswahl.baumansicht_tags
        
        for eintrag in all_el:
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
    
    
    def erzeuge_Fenster(self):
        if self.mb.debug: log(inspect.stack)
        
        LANG = self.mb.lang
        
        posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
        posSize = (posSize_main.X + 20, posSize_main.Y + 20,380,650)
        win,cont = self.mb.erzeuge_Dialog_Container(posSize)
        
        button_listener = Auswahl_Button_Listener(self.mb,win,cont)

        x1 = 10
        x2 = 70
        x3 = 120
        x4 = 140
        x5 = 270
        
        width = 120
        width2 = 80
        
        y = 20
        
        
        prop_names = ('Label',)
        prop_values = (LANG.ERZEUGE_NEUEN_TAB_AUS,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x2, y, 200, 20, prop_names, prop_values)  
        cont.addControl('Titel', control)
        model.FontWeight = 200.0
        
        y += 40
        
        
        prop_names = ('Label','State')
        prop_values = (LANG.MEHRFACHE_AUSWAHL,1)
        control, model = self.mb.createControl(self.mb.ctx, "RadioButton", x1, y, 200, 20, prop_names, prop_values)  
        cont.addControl('R1', control)
        
        prop_names = ('Label',)
        prop_values = (LANG.TAG_DATE,)
        control_RB2, model_RB2 = self.mb.createControl(self.mb.ctx, "RadioButton", x1, y, 200, 20, prop_names, prop_values)  
        cont.addControl('R2', control_RB2)
        control_RB2.Enable = False
        
        
        
        #################### EIGENE AUSWAHL ########################
        y += 30
        
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

        ################## RADIOBUTTON 2 ###################
        control_RB2.setPosSize(x1,y,0,0,3)
        ################## RADIOBUTTON 2 ###################
        
        
        
        y += 30
        
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedLine", x1, y, 340, 20, (), ())  
        cont.addControl('Line', control)
        
        y += 40
        
        
        ################## Einstellungen ###################
#         prop_names = ('Label',)
#         prop_values = ('Behalte Hierarchie bei',)
#         control, model = self.mb.createControl(self.mb.ctx, "CheckBox", x1, y, 200, 20, prop_names, prop_values) 
#         control.Enable = False 
#         cont.addControl('Hierarchie', control)
        
        y += 30
        
        prop_names = ('Label',)
        prop_values = (LANG.TABNAME,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", x1, y, 80, 20, prop_names, prop_values)  
        cont.addControl('tab_name_eingabe', control)
        
        prop_names = ('Text',)
        prop_values = ('Tab %s' %str(len(self.mb.tabs)+1),)
        control, model = self.mb.createControl(self.mb.ctx, "Edit", x3, y, width2, 20, prop_names, prop_values) 
        cont.addControl('tab_name', control)
  
        y += 50
        
        prop_names = ('Label',)
        prop_values = (LANG.OK,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", x5, y, width2, 30, prop_names, prop_values)  
        control.addActionListener(button_listener)            
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
        
        if 'nr0' not in ordinale:
            ordinale.insert(0,'nr0')
        
        papierkorb = self.mb.props['Projekt'].Papierkorb
        ordinale.append(papierkorb)
        
        ordinale = self.sortiere_ordinale(ordinale)
        
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
        
    def erzeuge_tab_XML_Eintrag(self,eintrag,tab_name):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[tab_name].xml_tree
        root = tree.getroot()
        et = self.mb.ET             
        ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
        
        if parent == 'root':
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
            self.mb.props[tab_name].Hauptfeld = self.mb.class_Hauptfeld.erzeuge_Navigations_Hauptfeld(win)  
            self.mb.class_Hauptfeld.erzeuge_Scrollbar(win,self.mb.ctx)   

            self.erzeuge_Eintraege_und_Bereiche(Eintraege,tab_name)    
        except:
            tb()

        
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
            
            return win,tab_id              
  
        except:
            tb()
       
    
    def erzeuge_Eintraege_und_Bereiche(self,Eintraege,tab_name):
        if self.mb.debug: log(inspect.stack)        
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0

        for eintrag in Eintraege:
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                     
            index = self.mb.class_Hauptfeld.erzeuge_Verzeichniseintrag(eintrag,self.mb.class_Zeilen_Listener,index,tab_name)

            if sicht == 'ja':
                # index wird in erzeuge_Verzeichniseintrag bereits erhoeht, daher hier 1 abziehen
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

        
    
    
    def erzeuge_Menu(self,win,tab_id):
        if self.mb.debug: log(inspect.stack)
        
        try:             
            self.listener = Menu_Kopf_Listener(self) 
            self.listener2 = Menu_Kopf_Listener2(self.mb,self) 
            self.erzeuge_MenuBar_Container(win)
            
            self.erzeuge_Menu_Kopf_Datei(win,self.listener)
            self.erzeuge_Menu_Kopf_Bearbeiten(win,self.listener)
            self.erzeuge_Menu_Kopf_Optionen(win,self.listener)
            
            if self.mb.debug:
                self.erzeuge_Menu_Kopf_Test(win,self.listener)
            
            #self.erzeuge_Menu_neuer_Ordner(win,listener2)
            #self.erzeuge_Menu_Kopf_neues_Dokument(win,listener2)
            self.erzeuge_Menu_Kopf_Papierkorb_leeren(win,self.listener2)
            
        except Exception as e:
                self.mb.Mitteilungen.nachricht('erzeuge_Menu ' + str(e),"warningbox")
                tb()

    
    def erzeuge_MenuBar_Container(self,win):
        menuB_control, menuB_model = self.mb.createControl(self.mb.ctx, "Container", 2, 2, 1000, 20, (), ())          
        menuB_model.BackgroundColor = KONST.Color_MenuBar_Container
         
        win.addControl('Organon_Menu_Bar', menuB_control)


    def erzeuge_Menu_Kopf_Datei(self,win,listener):
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 0, 2, 35, 20, (), ())           
        model.Label = self.mb.lang.FILE           
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar') 
        MenuBarCont.addControl('Datei', control)
        
        
    def erzeuge_Menu_Kopf_Bearbeiten(self,win,listener):
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 37, 2, 60, 20, (), ())           
        model.Label = self.mb.lang.BEARBEITEN_M             
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar') 
        MenuBarCont.addControl('Bearbeiten', control)
    
    
    def erzeuge_Menu_Kopf_Optionen(self,win,listener):         
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 100, 2, 55, 20, (), ())           
        model.Label = self.mb.lang.OPTIONS           
        control.addMouseListener(listener)
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')      
        MenuBarCont.addControl('Optionen', control)
        
        
    def erzeuge_Menu_Kopf_Test(self,win,listener):
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 300, 2, 50, 20, (), ())           
        model.Label = 'Test'                     
        control.addMouseListener(listener) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')   
        MenuBarCont.addControl('Projekt', control)
  
        
    def erzeuge_Menu_Kopf_Papierkorb_leeren(self,win,listener2):
        control, model = self.mb.createControl(self.mb.ctx, "ImageControl", 240, 0, 20, 20, (), ())           
        model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leeren.png'
        model.HelpText = self.mb.lang.CLEAR_RECYCLE_BIN
        model.Border = 0                       
        control.addMouseListener(listener2) 
        
        MenuBarCont = win.getControl('Organon_Menu_Bar')     
        MenuBarCont.addControl('Papierkorb_leeren', control)


    def erzeuge_Menu_DropDown_Container(self,ev,BREITE = 0, HOEHE = 0):
        if BREITE == 0:
            BREITE = KONST.Breite_Menu_DropDown_Container 
        if HOEHE == 0:
            HOEHE = KONST.Hoehe_Menu_DropDown_Container
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
        gnDefaultWindowAttributes = 1   # + 16 + 32 + 64 + 128 
    #         com_sun_star_awt_WindowAttribute_SHOW + \
    #         com_sun_star_awt_WindowAttribute_BORDER + \
    #         com_sun_star_awt_WindowAttribute_MOVEABLE + \
    #         com_sun_star_awt_WindowAttribute_CLOSEABLE + \
    #         com_sun_star_awt_WindowAttribute_SIZEABLE
           
        X = ev.Source.AccessibleContext.LocationOnScreen.value.X 
        Y = ev.Source.AccessibleContext.LocationOnScreen.value.Y + ev.Source.Size.value.Height
     
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = BREITE
        oRect.Height = HOEHE
         
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
        cont_model.BackgroundColor = KONST.MENU_DIALOG_FARBE  
        cont.setModel(cont_model)
        # need createPeer just only the container
        cont.createPeer(toolkit, oWindow)
        cont.setPosSize(0, 0, 0, 0, 15)
        
        oFrame.setComponent(cont, None)
         
        # create Listener
        listener = DropDown_Container_Listener(self)
        cont.addMouseListener(listener) 
        listener.ob = oWindow
        Name = ev.value.Source.AccessibleContext.AccessibleName
 
        self.menu_fenster = oWindow
 
        if Name == self.mb.lang.FILE:
            self.erzeuge_Menu_DropDown_Eintraege_Datei(oWindow, cont)
        if Name == self.mb.lang.BEARBEITEN_M:
            self.erzeuge_Menu_DropDown_Eintraege_Bearbeiten(oWindow, cont)
        if Name == self.mb.lang.OPTIONS:
            self.erzeuge_Menu_DropDown_Eintraege_Optionen(oWindow, cont)
        return Name
   
    
    def erzeuge_Menu_DropDown_Eintraege_Datei(self,window,cont):

        lang = self.mb.lang
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 10 ,  10 , 
                                        KONST.Breite_Menu_DropDown_Eintraege-6, 
                                        KONST.Hoehe_Menu_DropDown_Eintraege-6, (), ())   
        control.setMultipleMode(False)
        
        items = (
                lang.EXPORT_2, 
                lang.IMPORT_2)
        
        control.addItems(items, 0)
        model.BackgroundColor = KONST.MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Datei', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)
        
    
    def erzeuge_Menu_DropDown_Eintraege_Optionen(self,window,cont):
        try:
            if self.projekt_name != None:
                tag1 = self.settings_proj['tag1']
                tag2 = self.settings_proj['tag2']
                tag3 = self.settings_proj['tag3']
            else:
                tag1 = 0
                tag2 = 0
                tag3 = 0
                
            y = 10
            
            # Titel Baumansicht
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, y,KONST.BREITE_DROPDOWN_OPTIONEN-16, 30-6, (), ())   
            model.Label = self.mb.lang.SICHTBARE_TAGS_BAUMANSICHT
            model.FontWeight = 200
            cont.addControl('Titel_Baumansicht',control)
            
            y += 20
            
            # Tag1
            control_tag1, model_tag1 = self.mb.createControl(self.mb.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag1.Label = self.mb.lang.SHOW_TAG1
            model_tag1.State = tag1
            
            
            y += 16
            
            # Tag2
            control_tag2, model_tag2 = self.mb.createControl(self.mb.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag2.Label = self.mb.lang.SHOW_TAG2
            control_tag2.Enable = False
            model_tag2.State = tag2
            
            y += 16
                
            # Tag3
            control_tag3, model_tag3 = self.mb.createControl(self.mb.ctx, "CheckBox", 10, y, 
                                                      KONST.BREITE_DROPDOWN_OPTIONEN-6, 30-6, (), ())   
            model_tag3.Label = self.mb.lang.SHOW_TAG3
            control_tag3.Enable = False
            model_tag3.State = tag3
                
                
            tag1_listener = Tag1_Item_Listener(self,model_tag1)
            control_tag1.addItemListener(tag1_listener)
            cont.addControl('Checkbox_Tag1', control_tag1)
            cont.addControl('Checkbox_Tag2', control_tag2)
            cont.addControl('Checkbox_Tag3', control_tag3)
            
            
            y += 24
            
            
            # Titel Sidebar
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, y,KONST.BREITE_DROPDOWN_OPTIONEN-20, 30-6, (), ())   
            model.Label = self.mb.lang.SICHTBARE_TAGS_SEITENLEISTE
            model.FontWeight = 200
            cont.addControl('Titel_Baumansicht',control)
    
            y += 20
            
            
            sb_panels_tup = self.class_Sidebar.sb_panels_tup
            sb_panels1 = self.class_Sidebar.sb_panels1
            # Tags Sidebar
            
            listener_SB = Tag_SB_Item_Listener(self)
            
            for panel in sb_panels_tup:
            
                control, model = self.mb.createControl(self.mb.ctx, "CheckBox", 10, y, 
                                                          KONST.BREITE_DROPDOWN_OPTIONEN-20, 20, (), ())   
                model.Label = sb_panels1[panel]
                if panel in self.dict_sb['sichtbare']:
                    model.State = True
                else:
                    model.State = False
                cont.addControl(panel, control)
                control.addItemListener(listener_SB)
                y += 16
    
            y += 24
            
            HOEHE_LISTBOX = 50
            # ListBox
            control, model = self.mb.createControl(self.mb.ctx, "ListBox", 10, y, KONST.BREITE_DROPDOWN_OPTIONEN-20, 
                                            HOEHE_LISTBOX, (), ())   
            control.setMultipleMode(False)
            
            items = (  
                      self.mb.lang.ZEIGE_TEXTBEREICHE,
                     '-------',
                     'Homepage')
            
            control.addItems(items, 0)
            model.BackgroundColor = KONST.MENU_DIALOG_FARBE
            model.Border = False
            
            cont.addControl('Eintraege_Optionen', control)
            
            listener = DropDown_Item_Listener(self)  
            listener.window = window    
            control.addItemListener(listener)  
            
            y += HOEHE_LISTBOX + 10
            
            window.setPosSize(0,0,0,y,8)
        except:
            tb()

    
    def erzeuge_Menu_DropDown_Eintraege_Bearbeiten(self,window,cont):
            
        y = 10
                
        # ListBox
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 10, y, KONST.Breite_Menu_DropDown_Eintraege-6, 
                                        KONST.Hoehe_Menu_DropDown_Eintraege - 30, (), ())   
        control.setMultipleMode(False)
        
        items = ( self.mb.lang.UNFOLD_PROJ_DIR, 
                  '',
                  '*Suche',
                  '*Nach Tags sortieren')
                  
        
        control.addItems(items, 0)
        model.BackgroundColor = KONST.MENU_DIALOG_FARBE
        model.Border = False
        
        cont.addControl('Eintraege_Bearbeiten', control)
        
        listener = DropDown_Item_Listener(self)  
        listener.window = window    
        control.addItemListener(listener)  
        

    
    def erzeuge_neue_Projekte(self):
        try:
            self.class_Projekt.test()
        except:
            traceback.print_exc()
             
    def erzeuge_Zeile(self,ordner_oder_datei):
        try:
            self.class_Hauptfeld.erzeuge_neue_Zeile(ordner_oder_datei)    
        except:
            tb()                      
 
 
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
        
    def schliesse_Tab(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
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
            tb()


 

from com.sun.star.awt import XActionListener,XTextListener
class Auswahl_Button_Listener(unohelper.Base, XActionListener,XTextListener):
    def __init__(self,mb,win,fenster_cont):
        self.mb = mb
        self.win = win
        self.fenster_cont = fenster_cont
        
    def actionPerformed(self,ev):
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
                    self.mb.class_Tabs.erzeuge_neuen_tab()
            except:
                tb()
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
            tb()
        
    def erstelle_auswahl_dict(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        main_win = ev.Source.Context

        tab_auswahl = self.mb.props[T.AB].tab_auswahl
        
        tab_auswahl.rb = main_win.getControl('R1').State
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

        tab_auswahl.behalte_hierarchie_bei = 0#main_win.getControl('Hierarchie').State
        tab_auswahl.tab_name = main_win.getControl('tab_name').Model.Text
    
    
    def pruefe_tab_namen(self):

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
        
        x,y = self.get_fenster_position(ev)
        posSize = x,y,400,600
        
        lang = self.mb.lang
        sett = self.mb.settings_exp
 
        # Dict von alten Eintraegen bereinigen
        eintr = []
        for ordinal in sett['ausgewaehlte']:
            if ordinal not in self.mb.props[T.AB].dict_bereiche['ordinal']:
                eintr.append(ordinal)
        for ordn in eintr:
            del sett['ausgewaehlte'][ordn]

        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
        # Listener um Position zu bestimmen
        fenster_cont.Model.Text = lang.AUSWAHL
        fenster_cont.Model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
#         listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
#         fenster_cont.addEventListener(listenerF)
#         self.cl_exp.auswahl_fenster = fenster
        
        
        control_innen, model = self.mb.createControl(self.mb.ctx,"Container",20,0,posSize[2],posSize[3],(),() )  
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        fenster_cont.addControl('Huelle', control_innen)
        
        
        y = self.erzeuge_auswahl(control_innen)
        control_innen.setPosSize(0, 0,0,y + 20,8)

        self.setze_hoehe_und_scrollbalken(y,posSize[3],fenster,fenster_cont,control_innen)
    
      
    def setze_hoehe_und_scrollbalken(self,y,y_desk,fenster,fenster_cont,control_innen):  
        
        
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
  
        
    def erzeuge_auswahl(self,fenster_cont):
        try:
            sett = self.mb.settings_exp
            lang = self.mb.lang
            
            tree = self.mb.props[T.AB].xml_tree
            root = tree.getroot()
            
            baum = []
            self.mb.class_XML.get_tree_info(root,baum)
            
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
                
                control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,200,20,(),() )  
                control.Text = name
                fenster_cont.addControl('Titel', control)
                
                
                control, model = self.mb.createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
                model.Border = False
                if art in ('dir','prj'):
                    model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
                else:
                    model.ImageURL = 'private:graphicrepository/res/sx03150.png' 
                fenster_cont.addControl('Titel', control)   
                  
                    
                control, model = self.mb.createControl(self.mb.ctx,"CheckBox",x+20*int(lvl),y ,20,20,(),() )  
                control.addActionListener(listener)
                control.ActionCommand = ordinal+'xxx'+name
                if ordinal in sett['ausgewaehlte']:
                    model.State = sett['ausgewaehlte'][ordinal][1]
                fenster_cont.addControl(ordinal, control)
                
                y += 20 
                
            return y 
        except:
            tb()
            


from com.sun.star.awt import XItemListener
class Tag1_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,win,lb_alle_icons,lb_ausgew_icons):
        self.mb = mb
        self.win = win
        self.lb_alle_icons = lb_alle_icons
        self.lb_ausgew_icons = lb_ausgew_icons
        
    # XItemListener    
    def itemStateChanged(self, ev):  
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
            tb()


        
class Auswahl_Tags_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,win,cont,controls,source):
        self.mb = mb
        self.win = win
        self.cont = cont
        self.controls = controls
        self.source = source
        
    def actionPerformed(self,ev):
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
            tb()
            
    def erzeuge_button(self,ActionCommand):
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
     
class Menu_Kopf_Listener (unohelper.Base, XMouseListener):
    def __init__(self,tb):
        #self.tabs = tabs
        self.tb = tb
        self.menu_Kopf_Eintrag = 'None'
        self.tb.geoeffnetesMenu = None
        self.mb = self.tb.mb
         
    def mousePressed(self, ev):
        if ev.Buttons == MB_LEFT:

            try:
                if self.menu_Kopf_Eintrag == self.mb.lang.FILE:
                    self.tb.geoeffnetesMenu = self.tb.erzeuge_Menu_DropDown_Container(ev)
                elif self.menu_Kopf_Eintrag == 'Test':
                    self.mb.Test()          
                elif self.menu_Kopf_Eintrag == self.mb.lang.OPTIONS:
                    BREITE = KONST.BREITE_DROPDOWN_OPTIONEN
                    self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev,BREITE)
                elif self.menu_Kopf_Eintrag == self.mb.lang.BEARBEITEN_M:
                    self.mb.geoeffnetesMenu = self.mb.erzeuge_Menu_DropDown_Container(ev)
                 
                     
                self.mb.loesche_undo_Aktionen()
            except:
                tb()
            return False
 
    def mouseEntered(self, ev):
        
        ev.value.Source.Model.FontUnderline = 1     
        if self.menu_Kopf_Eintrag != ev.value.Source.Text:
            self.menu_Kopf_Eintrag = ev.value.Source.Text  
            if None not in (self.menu_Kopf_Eintrag,self.tb.geoeffnetesMenu):
                if self.menu_Kopf_Eintrag != self.tb.geoeffnetesMenu:
                    self.tb.menu_fenster.dispose()
  
        return False
    
    def mouseExited(self, ev):
        ev.value.Source.Model.FontUnderline = 0
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
 
class Menu_Kopf_Listener2 (unohelper.Base, XMouseListener):
    def __init__(self,mb,tabs):
        self.tabs = tabs
        self.mb = mb
        self.geklickterMenupunkt = None
         
    def mousePressed(self, ev):

        #print('mousePressed, Menu_Kopf_Listener2')
        if ev.Buttons == 1:
            if self.mb.projekt_name != None:
#                 if ev.Source.Model.HelpText == self.mb.lang.INSERT_DOC:            
#                     self.mb.erzeuge_Zeile('dokument')
#                 if ev.Source.Model.HelpText == self.mb.lang.INSERT_DIR:            
#                     self.mb.erzeuge_Zeile('Ordner')
                if ev.Source.Model.HelpText == self.mb.lang.CLEAR_RECYCLE_BIN:            
                    self.mb.class_Hauptfeld.leere_Papierkorb(True)  
                     
                self.mb.loesche_undo_Aktionen()
                return False
         
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False
 
class DropDown_Container_Listener (unohelper.Base, XMouseListener):
        def __init__(self,mb):
            #self.tabs = tabs
            self.ob = None
            self.mb = mb
            
        def mousePressed(self, ev):
            #print('mousePressed,DropDown_Container_Listener')  
            if ev.Buttons == MB_LEFT:
                return False
        
        def mouseExited(self, ev): 
            #print('mouseExited') 


            if self.enthaelt_Punkt(ev):
                pass
            else:            
                self.ob.dispose() 
                self.mb.geoeffnetesMenu = None   
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
 
        def mouseEntered(self,ev):
            return False
     
class DropDown_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,tb):
        #self.tabs = tabs
        self.mb = tb.mb
        self.window = None
         
    # XItemListener    
    def itemStateChanged(self, ev):        
        sel = ev.value.Source.Items[ev.value.Selected]

        try:
            if sel == self.mb.lang.NEW_PROJECT:
                self.do()
                self.mb.class_Projekt.erzeuge_neues_Projekt()
            elif sel == self.mb.lang.OPEN_PROJECT:
                self.do()
                self.mb.class_Projekt.lade_Projekt()
            elif sel == self.mb.lang.NEW_DOC:
                self.do()
                self.mb.erzeuge_Zeile('dokument')
            elif sel == self.mb.lang.NEW_DIR:
                self.do()
                self.mb.erzeuge_Zeile('Ordner')
            elif sel == self.mb.lang.EXPORT_2:
                self.do()
                self.mb.class_Export.export()
            elif sel == self.mb.lang.IMPORT_2:
                self.do()
                self.mb.class_Import.importX()
            elif sel == self.mb.lang.UNFOLD_PROJ_DIR:
                self.do()
                self.mb.class_Funktionen.projektordner_ausklappen()
            elif sel == self.mb.lang.ZEIGE_TEXTBEREICHE:
                self.do()
                oBool = self.mb.current_Contr.ViewSettings.ShowTextBoundaries
                self.mb.current_Contr.ViewSettings.ShowTextBoundaries = not oBool   
            elif sel == 'Homepage':
                self.do()
                import webbrowser
                webbrowser.open('https://github.com/XRoemer/Organon')
     
     
            self.mb.loesche_undo_Aktionen()
        except:
            tb()
         
    def do(self): 
        self.window.dispose()
        self.mb.geoeffnetesMenu = None
 
 
class Tag1_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,model,tabs):
        self.tabs = tabs
        self.model = model
         
    # XItemListener    
    def itemStateChanged(self, ev):  

        try:
            sett = self.mb.settings_proj
             
            if self.model.State == 1:
                sett['tag1'] = 1
            else:
                sett['tag1'] = 0
             
            if not sett['tag1']:
                sett['tag1'] = 0
                self.mache_tag1_sichtbar(False)
            else:
                sett['tag1'] = 1
                self.mache_tag1_sichtbar(True) 
             
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        except:
            tb()
     
    def mache_tag1_sichtbar(self,sichtbar):
     
        # alle Zeilen
        controls_zeilen = self.mb.props[T.AB].Hauptfeld.Controls
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
         
        if not sichtbar:
            for contr_zeile in controls_zeilen:
                tag1_contr = contr_zeile.getControl('tag1')
                text_contr = contr_zeile.getControl('textfeld')
                posSizeX = text_contr.PosSize.X
                 
                text_contr.setPosSize(posSizeX-16,0,0,0,1)
 
                if self.mb.settings_proj['tag2']:
                    tag2_contr = contr_zeile.getControl('tag2')
                if self.mb.settings_proj['tag3']:
                    tag3_contr = contr_zeile.getControl('tag3')
                     
                tag1_contr.dispose()
                 
        if sichtbar:
            for contr_zeile in controls_zeilen:
                text_contr = contr_zeile.getControl('textfeld')
                text_posX = text_contr.PosSize.X
                text_contr.setPosSize(text_posX + 16 ,0,0,0,1)                  
 
                icon_contr = contr_zeile.getControl('icon')
                icon_posX_end = icon_contr.PosSize.X + icon_contr.PosSize.Width 
 
                Color__Container = 10202
                Attr = (icon_posX_end,2,16,16,'egal', Color__Container)    
                PosX,PosY,Width,Height,Name,Color = Attr
 
                ord_zeile = contr_zeile.AccessibleContext.AccessibleName
                zeile_xml = root.find('.//'+ord_zeile)
                tag1 = zeile_xml.attrib['Tag1']
 
                control_tag1, model_tag1 = self.mb.createControl(self.mb.ctx,"ImageControl",PosX,PosY,Width,Height,(),() )  
                model_tag1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/punkt_%s.png' % tag1
                model_tag1.Border = 0
                control_tag1.addMouseListener(self.mb.class_Hauptfeld.tag1_listener)
 
                contr_zeile.addControl('tag1',control_tag1)
 
class Tag_SB_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,tabs):
        self.mb = mb
        self.tabs = tabs
         
    def itemStateChanged(self, ev):        
        name = ev.Source.AccessibleContext.AccessibleName
        state = ev.Source.State
        
         
        panels = self.mb.class_Sidebar.sb_panels2
        if state == 1:
            self.mb.dict_sb['sichtbare'].append(panels[name])
            #self.mb.class_Sidebar.erzeuge_sb_layout(name,'ein')
        else:
            self.mb.dict_sb['sichtbare'].remove(panels[name])
            #self.mb.class_Sidebar.erzeuge_sb_layout(name,'aus')
 


from com.sun.star.awt import XAdjustmentListener
class Auswahl_ScrollBar_Listener (unohelper.Base,XAdjustmentListener):
    
    def __init__(self,mb,fenster_cont):        
        self.mb = mb
        self.fenster_cont = fenster_cont
        
    def adjustmentValueChanged(self,ev):
        self.fenster_cont.setPosSize(0, -ev.value.Value,0,0,2)
        
    def disposing(self,ev):
        return False

            
            
class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,fenster_cont):
        self.mb = mb
        self.fenster_cont = fenster_cont
    
    def disposing(self,ev):
        return False

        
    def actionPerformed(self,ev):

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
        
            
                       

             
     
