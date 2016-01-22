# -*- coding: utf-8 -*-

import unohelper



class Tabs():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)

        self.mb = mb
        
        self.win_eigene_auswahl = False
        self.win_seitenleiste = False
        self.win_baumansicht = False
        
            
    def start(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)

        try:
            # ausgewaehlte zeile ueberpruefen
            props = self.mb.props[T.AB]
            selektiert = props.selektierte_zeile
            papierkorb_inhalt = self.mb.class_XML.get_papierkorb_inhalt()
            
            if selektiert in papierkorb_inhalt:
                self.mb.nachricht(LANG.NICHT_IM_PAPIERKORB_ERSTELLEN,'infobox')
                return
            
            self.berechne_ordinale_in_baum_und_tab(in_tab_einfuegen)
            self.dialog_tabs(in_tab_einfuegen)   
        except:
            log(inspect.stack,tb())
    
    def berechne_ordinale_in_baum_und_tab(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        if in_tab_einfuegen:
            tab = 'ORGANON'
            
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
        
        ordinale = set()
        ordinale_auswahl = []
        ordinale_seitenleiste = []
        ordinale_baumansicht = []
        
        if tab_auswahl.eigene_auswahl_use == 1:
            ordinale_auswahl = self.get_ordinale_eigene_auswahl()
        if tab_auswahl.seitenleiste_use == 1:
            ordinale_seitenleiste = self.get_ordinale_seitenleiste(in_tab_einfuegen)
        if tab_auswahl.baumansicht_use == 1:
            ordinale_baumansicht = self.get_ordinale_baumansicht(in_tab_einfuegen)
        
        
        ############# LOGIK ###############
        # Alle Ordinale in list eintragen
        ordinale |= set(ordinale_auswahl)
        ordinale |= set(ordinale_seitenleiste)
        ordinale |= set(ordinale_baumansicht)
        
        helfer = ordinale
        
        # Wenn Logik auf UND steht, alle im entsprechenden Tag 
        # nicht vorhandene Ordinale wieder loeschen  
        if tab_auswahl.seitenleiste_use and tab_auswahl.seitenleiste_log != 'V':
            ordinale = ordinale.intersection(set(ordinale_seitenleiste))
        if tab_auswahl.baumansicht_use and tab_auswahl.baumansicht_log != 'V':
            ordinale = ordinale.intersection(set(ordinale_baumansicht))
        if tab_auswahl.eigene_auswahl_use and tab_auswahl.eigene_auswahl_log != 'V':
            ordinale = ordinale.intersection(set(ordinale_auswahl))
          
        if len(ordinale) == 0:
            return []
        
        # Ordinale sortieren 
        ordinale = self.sortiere_ordinale(ordinale)
        if tab_auswahl.zeitlich_anordnen == 1:
            ordinale = self.sortiere_ordinale_zeitlich(ordinale,tab_auswahl)
        
        return ordinale
        

    def get_ordinale_seitenleiste(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        try:
            tags = self.mb.tags
            if in_tab_einfuegen:
                props = self.mb.props['ORGANON']
            else:
                props = self.mb.props[T.AB]
                
            tab_auswahl = self.mb.props[T.AB].tab_auswahl
            selektierte_tags = tab_auswahl.seitenleiste_tags.split(', ')
            
            if len(selektierte_tags) == 0:
                return []
            
            
            alle_ordinale = list(k[0] for i,k in props.dict_zeilen_posY.items() if k[2] not in props.dict_ordner[props.Papierkorb])
            tag_panels = list(tags['sammlung'])
            alle_tags_in_ordinal = {j : 
                              [ k for i in tag_panels for k in tags['ordinale'][j][i] ] 
                              for j in tags['ordinale']} 


            ordinale = []
                       
            for ordi in alle_ordinale:
                if in_tab_einfuegen:
                    if ordi in self.im_tab_vorhandene:
                        continue
                    
                tag_ist_drin = False
                s_tag_ist_drin = []
                
                for s_tag in selektierte_tags:

                    if s_tag in alle_tags_in_ordinal[ordi]:
                        tag_ist_drin = True
                        s_tag_ist_drin.append(True)
                    else:
                        s_tag_ist_drin.append(False)
                        
                    if tab_auswahl.seitenleiste_log_tags != 'V':
                        if False in s_tag_ist_drin:
                            tag_ist_drin = False
                    else:    
                        if tag_ist_drin:
                            break
                        
                if tag_ist_drin:
                    ordinale.append(ordi)
        except:
            log(inspect.stack,tb())
        
        return sorted(ordinale)
    
    def get_ordinale_baumansicht(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        ordinale = []
        
        if in_tab_einfuegen:
            tab = 'ORGANON'
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
            if tab_auswahl.baumansicht_log_tags == 'V':
                if eintrag.attrib['Tag1'] in ausgew_icons or eintrag.attrib['Tag2'] in ausgew_icons:
                    ordinale.append(eintrag.tag)
            else:
                if eintrag.attrib['Tag1'] in ausgew_icons and eintrag.attrib['Tag2'] in ausgew_icons:
                    ordinale.append(eintrag.tag)

        return ordinale
    
    def get_ordinale_eigene_auswahl(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinale = []

        for key in self.mb.settings_exp['ausgewaehlte']:
            if self.mb.settings_exp['ausgewaehlte'][key] == 1:
                ordinale.append(key)

        return ordinale
    
    
    def dialog_tabs_elemente(self,button_listener,in_tab_einfuegen = False):
        if self.mb.debug: log(inspect.stack)

        width = 20
        width2 = 20
                
        if in_tab_einfuegen:
            Ueberschrift = LANG.IMPORTIERE_IN_TAB
        else:
            Ueberschrift = LANG.ERZEUGE_NEUEN_TAB_AUS


        #################### nach Zeit / Datum ordnen ####################################
        tags = self.mb.tags
        kat_zeit = tuple(v[0] for i,v in tags['nr_name'].items() if v[1] == 'time')
        kat_datum = tuple(v[0] for i,v in tags['nr_name'].items() if v[1] == 'date')
        
        
        enable_zeit = kat_zeit != ()
        sel_zeit = int(kat_zeit != ())
        
        enable_datum  = kat_datum != ()
        sel_datum = int(kat_datum != ())
              
        enable_zd = enable_datum or enable_zeit

        tab_auswahl = self.mb.props[T.AB].tab_auswahl

        abst1 = 15
        
        controls = [
        
        10,
        ('Titel',"FixedText",1, 
                        'tab0', 0, 20, 20,
                        ('Label','FontWeight'),
                        (Ueberschrift,150),
                        {} 
                        ),
        20,
        ('R1',"FixedText",1, 
                        'tab0', 0, 200, 20, 
                        ('Label',),
                        (LANG.MEHRFACHE_AUSWAHL,),
                        {} 
                        ),
        #################### EIGENE AUSWAHL ########################
        40,
        ('Eigene_Auswahl_use',"CheckBox",1, 
                        'tab0', 0, width, 20,  
                        ('Label','State'),
                        (LANG.EIGENE_AUSWAHL,0),
                        {} 
                        ),
        0,
        ('Eigene_Auswahl',"Button",1, 
                        'tab1', 0, width2, 20,  
                        ('Label',),
                        (LANG.AUSWAHL,),
                        {'setActionCommand':'Eigene Auswahl','addActionListener':(button_listener)} 
                        ),
        ('but1_eigene_auswahl',"Button",0, 
                        'tab2', 0, 16, 16, 
                        ('Label','HelpText'),
                        (tab_auswahl.eigene_auswahl_log,LANG.TAB_HELP_TEXT),
                        {'setActionCommand':'V auswahl','addActionListener':(button_listener)} 
                        ),
        30,
        ('Line1',"FixedLine",0, 
                        'tab0x+{}-tab2'.format(abst1), 0, 50, 20, 
                        (), 
                        (),
                        {} 
                        ),
        #################### TAGS SEITENLEISTE ########################
        30,
        ('but1_Seitenleiste',"Button",0, 
                        'tab2', 0, 16, 16, 
                        ('Label','HelpText'),
                        (tab_auswahl.seitenleiste_log,LANG.TAB_HELP_TEXT),
                        {'setActionCommand':'V sl','addActionListener':(button_listener)} 
                        ),
        0,
        ('CB_Seitenleiste',"CheckBox",1, 
                        'tab0', 0, width, 20,  
                        ('Label',),
                        (LANG.TAGS_SEITENLEISTE,),
                        {} 
                        ),
        0,
        ('but2_Seitenleiste',"Button",1, 
                        'tab1', 0, width2, 20, 
                        ('Label',),
                        (LANG.AUSWAHL,),
                        {'setActionCommand':'Tags Seitenleiste','addActionListener':(button_listener)} 
                        ),
        30,
        ('but3_Seitenleiste',"Button",0, 
                        'tab0x+{}'.format(abst1), 0, 16, 16,  
                        ('Label','HelpText'),
                        (tab_auswahl.seitenleiste_log_tags,LANG.TAB_HELP_TEXT),
                        {'setActionCommand':'V sl tags','addActionListener':(button_listener)} 
                        ),
        0,
        ('txt_Seitenleiste',"FixedText",0, 
                        'tab0x+{}'.format(abst1+20), 0, 200, 50,  
                        ('Label','MultiLine'),
                        ('',True),
                        {} 
                        ),
        45,
        ('Line2',"FixedLine",0, 
                        'tab0x+{}-tab2'.format(abst1), 0, 50, 20, 
                        (), 
                        (),
                        {} 
                        ),  
        #################### TAGS BAUMANSICHT ########################
        30,
        ('but1_Baumansicht',"Button",0, 
                        'tab2', 0, 16, 16,  
                        ('Label','HelpText'),
                        (tab_auswahl.baumansicht_log,LANG.TAB_HELP_TEXT),
                        {'setActionCommand':'V baum','addActionListener':(button_listener)}
                        ), 
        0,
        ('CB_Baumansicht',"CheckBox",1, 
                        'tab0', 0, width, 20,  
                        ('Label','Enable'),
                        (LANG.TAGS_BAUMANSICHT,True),
                        {} 
                        ),
        ('but2_Baumansicht',"Button",1, 
                        'tab1', 0, width2, 20, 
                        ('Label',),
                        (LANG.AUSWAHL,),
                        {'setActionCommand':'Tags Baumansicht','addActionListener':(button_listener)}
                        ),
        30,
        ('but3_Baumansicht',"Button",0, 
                        'tab0x+{}'.format(abst1), 0, 16, 16,  
                        ('Label','HelpText'),
                        (tab_auswahl.baumansicht_log_tags,LANG.TAB_HELP_TEXT),
                        {'setActionCommand':'V baum tags','addActionListener':(button_listener)} 
                        ),  
        0,      
        ('icons_Baumansicht',"Container",0, 
                        'tab0x+{}-max'.format(abst1+20), 0, 200, 20, 
                        (), 
                        (),
                        {}
                        ),  
        30,
        ('Line3',"FixedLine",0, 
                        'tab0x+{}-tab2'.format(abst1), 0, 50, 20, 
                        (), 
                        (),
                        {}
                        ),
#         #################### SUCHE ########################
#         30,
#         ('but1_Suche',"Button",0, 
#                         'tab2', 0, 16, 16,  
#                         ('Label','HelpText'),
#                         (u'V',LANG.TAB_HELP_TEXT),
#                         {'setActionCommand':'V','addActionListener':(button_listener),'Enable':False}
#                         ),
#         0,
#         ('CB_Suche',"CheckBox",1, 
#                         'tab0', 0, width, 20,  
#                         ('Label',),
#                         (LANG.SUCHE,),
#                         {'Enable':False} 
#                         ),
#         0,
#         ('edit_Suche',"Edit",0, 
#                         'tab1-tab1-E', 0, width2, 20, 
#                         ('Label',),
#                         (LANG.AUSWAHL,),
#                         {'addTextListener':(button_listener),'Enable':False }
#                         ),
#         20,
#         ('txt_Suche',"FixedText",1, 
#                         'tab0x+{}'.format(abst1), 0, 200, 20,  
#                         ('Label',),
#                         ('',),
#                         {'Enable':False} 
#                         ),
#         20,
#         ('Line4',"FixedLine",0, 
#                         'tab0x+{}-tab2'.format(abst1), 0, 50, 20, 
#                         (), 
#                         (),
#                         {}
#                         ),  
        ####################################################################
        50,
        ('Zeit',"CheckBox",1, 
                        'tab0', 0, 290, 20,  
                        ('Label',),
                        (LANG.ZEITLICH_ANORDNEN,),
                        {'Enable':enable_datum or enable_zeit} 
                        ),        
        
        # RadioButtons
        30,
        ('z1',"RadioButton",1, 
                        'tab0x+{}'.format(abst1), 0, 30, 20, 
                        ('Label','State'),
                        (LANG.NUTZE_ZEIT,int(enable_zeit)) ,
                        {'Enable':enable_zeit}
                        ),
        20,
        ('z2',"RadioButton",1, 
                        'tab0x+{}'.format(abst1), 0, 30, 20,
                        ('Label','State'),
                        (LANG.NUTZE_DATUM,int(enable_datum and not enable_zeit)),
                        {'Enable':enable_datum}
                        ),      
        20,
        ('z3',"RadioButton",1, 
                        'tab0x+{}'.format(abst1), 0, 30, 20,  
                        ('Label',),
                        (LANG.NUTZE_ZEIT_UND_DATUM,),
                        {'Enable':enable_datum and enable_zeit}
                        ),
        -40,
        ('zeit_lb',"ListBox",1, 
                        'tab1-tab1-E', 0, width2, 18,  
                        ('Border','Dropdown','LineCount',),
                        (2,True,5),
                        {'addItems':(kat_zeit,0),'SelectedItems':(0,),'Enable': enable_zeit}
                        ),
        20,
        ('datum_lb',"ListBox",1, 
                        'tab1-tab1-E', 0, width2, 18, 
                        ('Border','Dropdown','LineCount',),
                        (2,True,5), 
                        {'addItems':(kat_datum,0),'SelectedItems':(0,),'Enable':enable_datum}
                        ),
        50,
        ('Zeit2',"CheckBox",1, 
                        'tab0x+{}'.format(abst1), 0, 20, 20, 
                        ('Label',),
                        (LANG.UNAUSGEZEICHNETE,), 
                        {'Enable':enable_datum or enable_zeit}
                        ),
        20,
        ###########################  TRENNER #####################################################
        ('Line5',"FixedLine",0, 
                        'tab0x-max', 0, 50, 20, 
                        (), 
                        (),
                        {}
                        ),  
        
        ]
        
        if not in_tab_einfuegen:
            
            controls.extend([
            30,
            
            
            ('tab_name_eingabe',"FixedText",1, 
                            'tab0', 0, 80, 20,  
                            ('Label',),
                            (LANG.TABNAME,),
                            {}
            
            
            ),
            20,
            
            
            ('tab_name',"Edit",0, 
                            'tab0-tab1', 0, width2, 20, 
                            ('Text',),
                            ('Tab %s' %str(len(self.mb.tabsX.Hauptfelder)+1),),
                            {}
                            ),
            

            ])
            
        controls.extend([
        50,
        ('but_ok',"Button",1,
                        'tab1-tab2', 0, width2, 30, 
                        ('Label',),
                        (LANG.OK,), 
                        {'setActionCommand':'ok_tab' if in_tab_einfuegen else 'ok','addActionListener':(button_listener)}
                        ),
        20,
        ])


        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 0),
                 1 : (None, 20),
                 2 : (None, 5),
                 }
        
        abstand_links = 10
        
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)    
        
        return controls2,max_breite
        
    
    def dialog_tabs(self,in_tab_einfuegen):
        if self.mb.debug: log(inspect.stack)
        
        try:

            button_listener = Auswahl_Button_Listener(self.mb,in_tab_einfuegen)

            controls,max_breite = self.dialog_tabs_elemente(button_listener,in_tab_einfuegen)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)

            X = self.mb.dialog.Size.Width
            posSize = X,30,max_breite,max_hoehe
            
            win,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            button_listener.win = win
            button_listener.fenster_cont = cont
            
            # Controls in Hauptfenster eintragen
            for c in ctrls:
                cont.addControl(c,ctrls[c])
            
        except:
            log(inspect.stack,tb())
    
 
    def sortiere_ordinale(self,ordinale):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['ORGANON']
        root = props.xml_tree.getroot()
        all_ordinals = [elem.tag for elem in root.iter()]

        return [o for o in all_ordinals if o in ordinale]


    def sortiere_ordinale_zeitlich(self,ordinale,tab_auswahl):
        if self.mb.debug: log(inspect.stack)
        try:
            dat_format = self.mb.settings_proj['datum_format']
            dat_trenner = self.mb.settings_proj['datum_trenner']

            if tab_auswahl.nutze_datum or tab_auswahl.nutze_zeit_und_datum:
                panel_nr_datum = [i for i,v in self.mb.tags['nr_name'].items() if v[0] == tab_auswahl.sel_datum][0]
            if tab_auswahl.nutze_zeit or tab_auswahl.nutze_zeit_und_datum:    
                panel_nr_zeit = [i for i,v in self.mb.tags['nr_name'].items() if v[0] == tab_auswahl.sel_zeit][0]
            
            dict_tags = self.mb.tags['ordinale']
            
            if tab_auswahl.kein_tag_einbeziehen == 1: 
                nutze_alle = True
            else:
                nutze_alle = False
            
            
            def berechne_ords(panel_nr):
                i = 0
                list_zeit = []
                for ordi in ordinale:
                    zeit = dict_tags[ordi][panel_nr]
                    if zeit == None:
                        if not nutze_alle:
                            continue
                        zeit = 'None'+str(i)
                        i += 1
                    list_zeit.append((zeit,ordi))
                
                return sorted(list_zeit)
            
            
            if tab_auswahl.nutze_zeit == 1:                
                sortierte_liste = berechne_ords(panel_nr_zeit)
                
            elif tab_auswahl.nutze_datum == 1:
                sortierte_liste = berechne_ords(panel_nr_datum)
                
            else:
                i = 0
                list_zeit = []
                
                
                for ordi in ordinale:
                    zeit = dict_tags[ordi][panel_nr_zeit]
                    datum_dict = dict_tags[ordi][panel_nr_datum]
                    
                    if datum_dict != None:
                        jahr = datum_dict['yyyy']
                        monat = datum_dict['mm']
                        tag = datum_dict['dd']
                        
                        datum = jahr + monat + tag
                    else:
                        datum = '9999999999999999999999'
                    
                    # Wenn d + t gewaehlt wurde, sollte zumindest ein Datum angegeben worden sein
                    # Wenn nicht, wird der Ordinal hier entfernt
                    if not nutze_alle:
                        if datum == None:
                            continue
                    
                    if zeit == None:
                        zeit = str(23599999)
                    elif zeit == 0:
                        zeit = '00000000'
                    else:
                        zeit = zeit.replace(':','')
                        
                    zeit2 = str(zeit)[:5]
                    
                    datumzeit = int(datum+zeit2)
                    list_zeit.append((datumzeit,ordi))
                    
                sortierte_liste = sorted(list_zeit)
    
            sort_list = list(x[1] for x in sortierte_liste)
            
            return sort_list
        except:
            log(inspect.stack,tb())
            
 

from com.sun.star.awt import XActionListener,XTextListener
class Auswahl_Button_Listener(unohelper.Base, XActionListener,XTextListener):
    def __init__(self,mb,in_tab_einfuegen = False):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = None
        self.fenster_cont = None
        self.in_tab_einfuegen = in_tab_einfuegen
        
        
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        try:
            cmd = ev.ActionCommand
            
            if cmd == 'Tags Seitenleiste':
                if not self.mb.class_Tabs.win_seitenleiste:
                    self.erzeuge_tag_auswahl_seitenleiste(ev)
                    self.mb.class_Tabs.win_seitenleiste = True
                
            elif cmd == 'Tags Baumansicht':
                if not self.mb.class_Tabs.win_baumansicht:
                    self.erzeuge_tag_auswahl_baumansicht(ev)
                    self.mb.class_Tabs.win_baumansicht = True
                
            elif cmd == 'Eigene Auswahl':
                if not self.mb.class_Tabs.win_eigene_auswahl:
                    self.erzeuge_Fenster_fuer_eigene_Auswahl(ev)
                    self.mb.class_Tabs.win_eigene_auswahl = True
                    
            elif cmd == 'ok':
                try:
                    self.erstelle_auswahl_dict(ev)
                    ok = self.pruefe_tab_namen()
                    if ok:
                        self.win.dispose()
                        
                        if len(self.mb.undo_mgr.AllUndoActionTitles) > 0:
                            ordinal = self.mb.props[T.AB].selektierte_zeile
                            bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                            # nachfolgende Zeile erzeugt bei neuem Tab Fehler - 
                            path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                            self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)
                        
                        ordinale = self.mb.class_Tabs.berechne_ordinal_nach_auswahl(False)
                        if ordinale == []:
                            self.mb.popup(LANG.AUSWAHL_KEIN_ERGEBNIS,2)
                            return
                        self.mb.tabsX.erzeuge_neuen_tab2(ordinale)

                except:
                    log(inspect.stack,tb())
                    
            elif cmd == 'ok_tab':
                try:
                    self.erstelle_auswahl_dict(ev)
                    self.win.dispose()
                    
                    ordinale = self.mb.class_Tabs.berechne_ordinal_nach_auswahl(True)
    
                    if ordinale == []:
                        self.mb.popup(LANG.AUSWAHL_KEIN_ERGEBNIS,2)
                        return
                    self.mb.tabsX.fuege_ausgewaehlte_in_tab_ein(ordinale)
                    
                except:
                    log(inspect.stack,tb())
                    
            elif cmd[0] == 'V':
                
                tab_auswahl = self.mb.props[T.AB].tab_auswahl

                if ev.Source.Model.Label == 'V':
                    logic = u'\u039B'
                else:
                    logic = 'V'
                    
                ev.Source.Model.Label = logic
                    
                if cmd == 'V sl':
                    tab_auswahl.seitenleiste_log = logic
                elif cmd == 'V sl tags':
                    tab_auswahl.seitenleiste_log_tags = logic
                elif cmd == 'V baum':
                    tab_auswahl.baumansicht_log = logic
                elif cmd == 'V baum tags':
                    tab_auswahl.baumansicht_log_tags = logic 
                elif cmd == 'V auswahl':
                    tab_auswahl.eigene_auswahl_log = logic 

        except:
            log(inspect.stack,tb())
    
        
    def get_fenster_position(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        x = self.mb.dialog.Size.Width + ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.Size.Width +20
        return x,30      
        
    def erzeuge_tag_auswahl_baumansicht(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # BENUTZTE TAGS
            url_nutzer_tags = []
            nutzer_tags = []
            farb_tags = []
            
            tree = self.mb.props['ORGANON'].xml_tree
            root = tree.getroot()
            alle_elem = root.findall('.//')            
            
            for el in alle_elem:
                farbe = el.attrib['Tag1']
                url = el.attrib['Tag2']
                
                if farbe not in ('','leer') and farbe not in farb_tags:
                    farb_tags.append(el.attrib['Tag1'])
                
                if url not in ('','leer') and url not in url_nutzer_tags:
                    url_nutzer_tags.append(el.attrib['Tag2'])
                    name = os.path.basename(el.attrib['Tag2']).split('.')[0]
                    nutzer_tags.append((name,el.attrib['Tag2']))
            
            
            # TITEL
            prop_names = ('Label',)
            prop_values = (LANG.AUSGEWAEHLTE,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 100, 20, prop_names, prop_values) 
            model.FontWeight = 150 
                        
            prop_names = ('Label',)
            prop_values = (LANG.BENUTZTE,)
            controlT2, modelT2 = self.mb.createControl(self.mb.ctx, "FixedText", 120, 10, 300, 20, prop_names, prop_values) 
            modelT2.FontWeight = 150 
            
            # TRENNER
            controlTrenner, modelTrenner = self.mb.createControl(self.mb.ctx, "FixedLine", 100, 40, 10, 340, (), ()) 
            modelTrenner.Orientation = 1
            
            farb_icons,ausgew_icons,user_icons,breite = self.erzeuge_ListBox_Tag1(tuple(farb_tags),nutzer_tags) 
            
            breite1 = controlT2.PreferredSize.Width + controlT2.PosSize.X
            if breite1 > breite:
                breite = breite1           
            
            tag_item_listener = Tag1_Listener(self.mb,self.fenster_cont,farb_icons,ausgew_icons,user_icons)
            farb_icons.addItemListener(tag_item_listener)
            ausgew_icons.addItemListener(tag_item_listener)
            user_icons.addItemListener(tag_item_listener)
            
            
        except:
            log(inspect.stack,tb())        
        
        # DIALOG FENSTER
        x,y = self.get_fenster_position(ev)
        posSize = (x,y,breite + 20,400)
        
        win,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)   
                    
        cont.addControl('ausgewaehlte_XXX', control)
        cont.addControl('ausgewaehlte_YYY', controlT2)
        cont.addControl('Trenner', controlTrenner)
        cont.addControl('Eintraege_Tag1', farb_icons)
        cont.addControl('Ausgewaehlte_Tag1', ausgew_icons)
        cont.addControl('Ausgewaehlte_Tag1', user_icons)
        
        dispose_listener = Listener_for_Win_dispose(self.mb,'baumansicht')
        win.addEventListener(dispose_listener)

            
       
    def erzeuge_ListBox_Tag1(self,farb_tags,nutzer_tags):
        if self.mb.debug: log(inspect.stack)
        
        # FARB_TAGS
        control, model = self.mb.createControl(self.mb.ctx, "ListBox", 120, 40, 80 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control.setMultipleMode(False)
        model.Border = 0          
        
        for item in farb_tags:
            model.insertItem(control.ItemCount,item,KONST.URL_IMGS+'punkt_%s.png' %item)
        
        
        # NUTZER_TAGS
        control2, model2 = self.mb.createControl(self.mb.ctx, "ListBox", 220, 40, 280 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control2.setMultipleMode(False)
        model2.Border = 0

        for (item,url) in nutzer_tags:
            model2.insertItem(0,item,url)
        
        breite = control2.PreferredSize.Width + control2.PosSize.X
        
        # Listbox fuer ausgewaehlte Punkte
        control_ausgewaehlte, model = self.mb.createControl(self.mb.ctx, "ListBox", 10, 40, KONST.BREITE_TAG1_CONTAINER -8 , KONST.HOEHE_TAG1_CONTAINER -8 , (), ())   
        control_ausgewaehlte.setMultipleMode(False)
        model.Border = 0
        
        return control,control_ausgewaehlte,control2, breite
    
        
    def erzeuge_tag_auswahl_seitenleiste(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:

            sammlung = self.mb.tags['sammlung']
            tag_panels = [[i,v[0]] for i,v in self.mb.tags['nr_name'].items() if v[1] == 'tag']            
            
            self.controls = []
            auswahl_listener = Auswahl_Tags_Listener(self.mb,self.controls,ev.Source)

            x = 150
            width = 100
            ctrls = {}
            y_max = 0
            
            for nr,name in tag_panels:

                prop_names = ('Label','Align')
                prop_values = (name,1)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", x, 10, width, 20, prop_names, prop_values)  
                
                ctrls.update({name:control})
                
                y = 0
                for t in sammlung[nr]:
                    prop_names = ('Label',)
                    prop_values = (t,)
                    control, model = self.mb.createControl(self.mb.ctx, "Button", x + 10, y + 30, width - 20, 20, prop_names, prop_values)  
                    control.setActionCommand(t)
                    control.addActionListener(auswahl_listener)
                    ctrls.update({t+'###':control})
                    
                    y += 25
                    
                    if y > y_max:
                        y_max = y
                
                x += (width + 10)
            
            
            x1,y1 = self.get_fenster_position(ev)
            posSize = (x1,y1,x,y_max + 250)

            win,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            
            auswahl_listener.win = win
            auswahl_listener.cont = cont
            
            for c,control in ctrls.items():
                cont.addControl(c, control)
                
                
            prop_names = ('Label',)
            prop_values = (LANG.AUSGEWAEHLTE,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 100, 20, prop_names, prop_values) 
            model.FontWeight = 200.0 
            cont.addControl('ausgewaehlte_XXX', control)
            
               
            dispose_listener = Listener_for_Win_dispose(self.mb,'seitenleiste')
            win.addEventListener(dispose_listener)
            
        except:
            log(inspect.stack,tb())
            
            
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

        tab_auswahl.zeitlich_anordnen = main_win.getControl('Zeit').State
        tab_auswahl.kein_tag_einbeziehen = main_win.getControl('Zeit2').State
        tab_auswahl.nutze_zeit = main_win.getControl('z1').State
        tab_auswahl.sel_zeit = main_win.getControl('zeit_lb').SelectedItem
        tab_auswahl.nutze_datum = main_win.getControl('z2').State
        tab_auswahl.sel_datum = main_win.getControl('datum_lb').SelectedItem
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
            self.mb.nachricht(LANG.TAB_EXISTIERT_SCHON,'infobox')
            return False
        
        else:
            return True
    
    def get_baumansicht_icons(self):
        if self.mb.debug: log(inspect.stack)
        
        container = self.fenster_cont.getControl('icons_Baumansicht')
        ausgew_icons = []
        
        for cont in container.Controls:
            
            if 'punkt_' in cont.Model.ImageURL:
                name = cont.Model.ImageURL.split('punkt_')[1]
                name = name.replace('.png','')
            else:
                name = cont.Model.ImageURL
            
            ausgew_icons.append(name)
        
        return ausgew_icons

    
    def erzeuge_Fenster_fuer_eigene_Auswahl(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        x,y = self.get_fenster_position(ev)
        posSize = x,30,400,600
        pos = x,y
        
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
        
        if self.in_tab_einfuegen:
            tab_name = 'ORGANON'
        else:
            tab_name = T.AB
        
        (y,
         fenster,
         fenster_cont,
         control_innen,
         ctrls) = self.mb.class_Fenster.erzeuge_treeview_mit_checkbox(
                                                                tab_name=tab_name,
                                                                pos=pos,
                                                                auswaehlen=True)
         
        if self.in_tab_einfuegen:
            vorhandene_dateien = list(self.mb.props[T.AB].dict_bereiche['ordinal'])

            for v in vorhandene_dateien:
                if v in ctrls:
                    for c in ctrls[v]:
                        c.setEnable(False)
                
        
        dispose_listener = Listener_for_Win_dispose(self.mb,'eigene auswahl')
        fenster_cont.addEventListener(dispose_listener)
    
    def disposing(self,ev):
        return False

      

            

from com.sun.star.lang import XEventListener

class Listener_for_Win_dispose(unohelper.Base,XEventListener):
    def __init__(self,mb,win):
        self.mb = mb
        self.win = win
    
    def disposing(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            tabs = self.mb.class_Tabs
            if self.win == 'baumansicht':
                tabs.win_baumansicht = False
            elif self.win == 'eigene auswahl':
                tabs.win_eigene_auswahl = False
            elif self.win == 'seitenleiste':
                tabs.win_seitenleiste = False
        except:
            log(inspect.stack,tb())



from com.sun.star.awt import XItemListener
class Tag1_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,win,farb_icons,ausgew_icons,user_icons):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = win
        self.farb_icons = farb_icons
        self.ausgew_icons = ausgew_icons
        self.user_icons = user_icons
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
        
        try: 
            container_baumansicht = self.win.getControl('icons_Baumansicht')
            item = ev.Source.Items[ev.Selected]

            if ev.Source != self.ausgew_icons:
                
                if self.ausgew_icons.ItemCount == 0:
                    Bedingung = True
                elif item not in self.ausgew_icons.Items:
                    Bedingung = True
                else:
                    Bedingung = False
                    
                if Bedingung:

                    url = ev.Source.Model.AllItems[ev.Selected].Second
                    self.ausgew_icons.Model.insertItem(self.ausgew_icons.ItemCount,item, url)
                       
            else:
                pos = self.ausgew_icons.Items.index(item) 
                self.ausgew_icons.Model.removeItem(pos)
            
            
            container_controls = container_baumansicht.getControls()
            
            for con in container_controls:
                con.dispose()
            
            x = 0  
            for it in self.ausgew_icons.Model.AllItems:
                
                prop_names = ('ImageURL','Border')
                prop_values = (it.Second,0)
                control, model = self.mb.createControl(self.mb.ctx, "ImageControl", x, 0, 16, 16, prop_names, prop_values) 
                container_baumansicht.addControl(it.First, control)
                x += 20

        except:
            log(inspect.stack,tb())
            
    
    def disposing(self,ev):
        return False


        
class Auswahl_Tags_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,controls,source):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = None
        self.cont = None
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
            log(inspect.stack,tb())
            
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
    
    def disposing(self,ev):
        return False


                      
      
        
            


class TabsX():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
        self.organon_fenster = None
        self.breite_sichtbar = None
        self.breite_sichtbar = None
        self.tab_listener = None
#         (kante,h_tabs,mindestgroesse,breite_sichtbar,
#          breite_hauptfeld,hoehe_hauptfeld) = abmessungen
        
        
        self.kante = 2
        self.h_tabs = 20
        self.mindestgroesse = 50
        
        self.breite_hauptfeld = 1800
        self.hoehe_hauptfeld = 2000
        
        
        self.Hauptfelder = {}
        self.breiten_tabs = []
        
        
        self.tableiste = None
        self.tableiste_hoehe = 0
                
        self.prj_ord = None
        
    
    def run(self,organon_fenster):
        if self.mb.debug: log(inspect.stack)
        
        self.organon_fenster = organon_fenster
        self.breite_sichtbar = self.organon_fenster.PosSize.Width
        self.breite_sichtbar -= 2 * self.kante
        self.tab_listener = Tab_Leiste_Listener(self.mb,self.Hauptfelder,self.organon_fenster)
        
        self.prj_ord = 'nr0'
        
        self.tableiste,tab_model = self.mb.createControl(self.mb.ctx,'Container',0,0,
                                    self.breite_sichtbar + 2 * self.kante,0,
                                   ('BackgroundColor',),(KONST.FARBE_TABS_TRENNER,))
        return self.tableiste
    
    def erzeuge_neuen_tab(self,tab_name,loesche=False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if tab_name in self.Hauptfelder:
                if loesche:
                    del (self.Hauptfelder[tab_name])
                else:
                    return
            
            tab_fenster = self.erzeuge_tabeintrag_und_fenster(tab_name)
            
            self.Hauptfelder.update({tab_name:tab_fenster})
            self.tableiste_hoehe = self.layout_tab_zeilen()
            self.mb.dialog.addControl(tab_name,tab_fenster)

            return tab_fenster
            
        except Exception as e:
            log(inspect.stack,tb())

        
    def erzeuge_neuen_tab2(self,ordinale, tab_name=None):
        if self.mb.debug: log(inspect.stack)

        try:
            if not tab_name:
                tab_auswahl = self.mb.props[T.AB].tab_auswahl
                tab_name = tab_auswahl.tab_name
            
            # zur Sicherung, damit der projekt xml nicht ueberschrieben wrd
            if tab_name == 'ORGANON':
                self.mb.nachricht('ERROR',"warningbox",16777216)
                return
            
            self.erzeuge_props(tab_name)
            Eintraege = self.erzeuge_Eintraege(tab_name,ordinale)        

            win = self.erzeuge_neuen_tab(tab_name)
 
            self.mb.erzeuge_Menu(win,tab=True)
 
            self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
            erste_datei =   self.get_erste_datei(tab_name)          
            self.setze_selektierte_zeile(erste_datei,tab_name)
            self.mb.class_Baumansicht.korrigiere_scrollbar()            
            
            tree = self.mb.props[tab_name].xml_tree
            Path = os.path.join(self.mb.pfade['tabs'] , tab_name +'.xml' )
            self.mb.tree_write(tree,Path)
            
            self.tab_umschalten(tab_name)
            
        except:
            log(inspect.stack,tb())
            
        
    def pref_size(self,w,ctrl):
        pref =  ctrl.PreferredSize.Width
        if pref < self.mindestgroesse:
            pref = self.mindestgroesse
        ctrl.setPosSize(0,0,pref,0,4)
        self.breiten_tabs.append([w,pref,ctrl])
    
    
    def layout_tab_zeilen(self,breite=None):
        if self.mb.debug: log(inspect.stack)
        
        if breite != None:
            self.breite_sichtbar = breite - 2 * self.kante
            self.tableiste.setPosSize(0,0,breite,0,4)
        
        zeilen,mehrraum = self.berechne_tab_zeilen()
        hoehe = self.setze_tab_umbruch(zeilen, mehrraum)
                    
        self.tableiste.setPosSize(0,0,0,hoehe,8)
        
        for t_name in self.Hauptfelder:
            self.Hauptfelder[t_name].setPosSize(0,hoehe,0,0,2)
        
        self.tableiste_hoehe = hoehe    
        return hoehe
      
        
    def erzeuge_tabeintrag_und_fenster(self,tab_name):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # Eintrag in Tableiste
            ctrl,model = self.mb.createControl(self.mb.ctx,'FixedText',0,0,0,self.h_tabs,
                        ('Label','Border','BackgroundColor',
                         'TextColor',
                         'Align','VerticalAlign'),
                        (tab_name,0,KONST.FARBE_TABS_HINTERGRUND,
                         KONST.FARBE_TABS_SCHRIFT,
                         1,1))
            
            ctrl.addMouseListener(self.tab_listener)
            self.tableiste.addControl(tab_name,ctrl)
            
            # Groesse zurechtschneiden
            self.pref_size(tab_name,ctrl)
            
            
            # Tabfenster
            container_hf,model_hf = self.mb.createControl(self.mb.ctx,'Container',
                                                          0,0,self.breite_hauptfeld,self.hoehe_hauptfeld,
                                           ('BackgroundColor',),(KONST.FARBE_HF_HINTERGRUND,))
        except:
            log(inspect.stack,tb())
            
        return container_hf
    

    def berechne_tab_zeilen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            x = 0
            zeilen = {1:[]}
            mehrraum = []
            z = 1
            for b in self.breiten_tabs:
                x += b[1] 
                if x < self.breite_sichtbar -10:
                    zeilen[z].append(b)
                    continue
                else:
                    z += 1
                    zeilen.update({z:[b]})
                    mehrraum.append(self.breite_sichtbar-x + b[1])
                    x = b[1]
            mehrraum.append(self.breite_sichtbar-x)
            
            if zeilen[1] == []:
                zeilen = {x-1:zeilen[x] for x in zeilen}
                del zeilen[0]
                mehrraum.pop(0)

            return zeilen,mehrraum
        except:
            print(tb())
    

    def setze_tab_umbruch(self,zeilen,mehrraum):
        if self.mb.debug: log(inspect.stack)
        
        # Tabzeilen setzen    
        for zeile in sorted(zeilen):
            x = 0
            X = 0
            try:
                if len(zeilen[zeile]) > 1:
                    for z in zeilen[zeile]:
                        if zeilen[zeile].index(z) != len(zeilen[zeile])-1:
                            mehr = mehrraum[zeile-1]
                            mehr -= self.kante * (len(zeilen[zeile]) -1 )
                            mehr = int( mehr / len(zeilen[zeile]) )
                            X = z[1] + mehr 
                            y = ( self.h_tabs + self.kante ) * (zeile-1) + self.kante
                            z[2].setPosSize(x + self.kante,y,X,0,7)
                            x += X + self.kante
                        else:
                            # letzter Eintrag
                            # um ein gleichmaessiges Ende zu bekommen
                            X = self.breite_sichtbar  - x
                            y = ( self.h_tabs + self.kante ) * (zeile-1) + self.kante
                            z[2].setPosSize(x + self.kante,y,X,0,7)
                            
                else:
                    z = zeilen[zeile][0]
                    mehr = mehrraum[zeile-1]
                    X = z[1] + mehr 
                    y = ( self.h_tabs + self.kante) * (zeile-1) + self.kante
                    z[2].setPosSize(x + self.kante,y,X,0,7)
                    x += X + self.kante
            except:
                print(tb())
        
        hoehe = zeile * (self.h_tabs + self.kante) + self.kante
        return hoehe
        
        
    def tab_umschalten(self,tab_name,wurde_geloescht=False):
        if self.mb.debug: log(inspect.stack)
        
        try:

            if tab_name != T.AB:
                
                neu = self.Hauptfelder[tab_name]
                
                if len(neu.Controls) == 0:
                    # wenn Tab noch nicht angewaehlt wurde
                    # muss er noch erzeugt werden
                    self.mb.erzeuge_Menu(neu,tab=True)
                    Eintraege = self.get_tab_Eintraege(tab_name) 
                    self.erzeuge_Hauptfeld(neu,tab_name,Eintraege)
     
                    erste_datei = self.get_erste_datei(tab_name)
                    self.setze_selektierte_zeile(erste_datei,tab_name)
                    self.mb.class_Baumansicht.korrigiere_scrollbar()
                    

                neu.setVisible(True)
                
                tab_icon = self.tableiste.getControl(tab_name)
                tab_icon.Model.BackgroundColor = KONST.FARBE_TABS_SEL_HINTERGRUND
                tab_icon.Model.TextColor = KONST.FARBE_TABS_SEL_SCHRIFT
                
                if not wurde_geloescht:
                    alt = self.Hauptfelder[T.AB]
                    alt.setVisible(False)
    
                    tab_icon_alt = self.tableiste.getControl(T.AB)
                    tab_icon_alt.Model.BackgroundColor = KONST.FARBE_TABS_HINTERGRUND
                    tab_icon_alt.Model.TextColor = KONST.FARBE_TABS_SCHRIFT
                
                if wurde_geloescht:
                    pass
                
                elif len(self.mb.undo_mgr.AllUndoActionTitles) > 0:
                    
                    ordinal = self.mb.props[T.AB].selektierte_zeile
                    bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]

                    path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])

                    self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)      

                T.AB = tab_name
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche()
                self.mb.class_Baumansicht.korrigiere_scrollbar()
                
                if not wurde_geloescht:
                    self.mb.class_Sidebar.erzeuge_sb_layout()
                
        except:
            log(inspect.stack,tb())
            
        
    def schliesse_Tab(self,abfrage = True):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            if T.AB == 'ORGANON':
                return
            
            if abfrage:
                # Frage: Soll Tab wirklich geschlossen werden?
                entscheidung = self.mb.nachricht(LANG.TAB_SCHLIESSEN %T.AB ,"warningbox",16777216)
                # 3 = Nein oder Cancel, 2 = Ja
                if entscheidung == 3:
                    return
                #print('active tab id', self.mb.tabsX.active_tab_id,self.mb.tabs[self.mb.tabsX.active_tab_id][1])
                if T.AB == 'ORGANON':
                    self.mb.nachricht("Project tab can't be closed" ,"warningbox",16777216)
                    return


            # loesche tab.xml
            tab_name = T.AB
            Path = os.path.join(self.mb.pfade['tabs'], '%s.xml' % tab_name)

            os.remove(Path)
                        
            tab_container = self.Hauptfelder[tab_name]
            tab_container.dispose()
            
            for t in self.breiten_tabs:
                if t[0] == tab_name:
                    ctrl = t[2]
                    ctrl.dispose()
                    index = self.breiten_tabs.index(t)
                    break
                
            self.breiten_tabs.pop(index)
            self.layout_tab_zeilen()

            del self.mb.props[T.AB]
            del self.mb.tabsX.Hauptfelder[T.AB]
            
            self.tab_umschalten('ORGANON',wurde_geloescht=True)
        except:
            log(inspect.stack,tb())
        
        
    def fuege_ausgewaehlte_in_tab_ein(self,ordinale):
        if self.mb.debug: log(inspect.stack)

        tab_name = T.AB
        tab_xml = copy.deepcopy(self.mb.props[T.AB].xml_tree)
        ord_selektierter = self.mb.props[T.AB].selektierte_zeile
        
        self.schliesse_Tab(False)
     
        try:
                        
            self.erzeuge_props(tab_name)

            Eintraege = self.erzeuge_neue_Eintraege_im_tab(tab_name,ordinale,tab_xml, ord_selektierter)        

            win = self.erzeuge_neuen_tab(tab_name,loesche=True)
  
            self.mb.erzeuge_Menu(win)
            self.erzeuge_Hauptfeld(win,tab_name,Eintraege)
            self.mb.class_Baumansicht.korrigiere_scrollbar()
            
            tree = self.mb.props[tab_name].xml_tree
            Path = os.path.join(self.mb.pfade['tabs'] , tab_name +'.xml' )
            self.mb.tree_write(tree,Path)
            
            self.setze_selektierte_zeile(ord_selektierter,tab_name)
            self.tab_umschalten(tab_name)
        except:
            log(inspect.stack,tb())
            
    
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
#         papierkorb = ('nr5','root',LANG.PAPIERKORB,0,'waste','zu','ja','leer','leer','leer')
                
        xml_tree = self.mb.props['ORGANON'].xml_tree
        root = xml_tree.getroot()

        Eintraege = []
        
        
        # Ordinal des Projekts muss enthalten sein und an erster Stelle stehen 
        if self.prj_ord not in ordinale:
            ordinale.insert(0,self.prj_ord)
        else:
            ordinale.remove(self.prj_ord)
            ordinale.insert(0,self.prj_ord)
        
        
        papierkorb = self.mb.props['ORGANON'].Papierkorb
        ordinale.append(papierkorb)
                
        for ordi in ordinale:
            elem = root.find('.//'+ordi)

            ordinal = elem.tag
            name    = elem.attrib['Name']
            art     = elem.attrib['Art']
            if ordinal not in (self.prj_ord,papierkorb):
                lvl     = 1 
                parent  = self.prj_ord
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
             
        xml_tree = self.mb.props['ORGANON'].xml_tree
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
    
    
    def erzeuge_tab_XML_Eintrag(self,eintrag,tab_name,parent_neu = None,lvl_neu = None):
        if self.mb.debug: log(inspect.stack)
        
        try:
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
        except:
            log(inspect.stack,tb())
 
    
    def erzeuge_Hauptfeld(self,win,tab_name,Eintraege):
        if self.mb.debug: log(inspect.stack)
        
        try:
            props = self.mb.props
            self.mb.props[tab_name].Hauptfeld = self.mb.class_Baumansicht.erzeuge_Feld_Baumansicht(win)  
            self.mb.class_Fenster.erzeuge_Scrollbar2(win)  
            self.erzeuge_Eintraege_und_Bereiche(Eintraege,tab_name) 
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_hf_ctrls()   
        except:
            log(inspect.stack,tb())

            
    def erzeuge_Eintraege_und_Bereiche(self,Eintraege,tab_name):
        if self.mb.debug: log(inspect.stack)        
        props = self.mb.props[tab_name]
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0
        
        if self.mb.settings_proj['tag3']:
            tree = props.xml_tree
            root = tree.getroot()
            gliederung = self.mb.class_Gliederung.rechne(tree)
        else:
            gliederung = None
        
        
        for eintrag in Eintraege:
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                  
            index,ctrl = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,
                                                                               gliederung,index,tab_name,
                                                                               neuer_tab=True)
            
            if sicht == 'ja':
                # index wird in erzeuge_Zeile_in_der_Baumansicht bereits erhoeht, daher hier 1 abziehen
                pos_Y = (index-1) * KONST.ZEILENHOEHE
                props.dict_zeilen_posY.update({ pos_Y : eintrag })
                self.mb.sichtbare_bereiche.append( 'OrganonSec' + str(index2) )
                props.dict_posY_ctrl.update({ pos_Y : ctrl })
                
            # Bereiche   
            inhalt = name
            path = os.path.join(self.mb.pfade['odts'],ordinal+'.odt') 
            
            Bereichsname_dict.update({'OrganonSec'+str(index2):path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
            
            index2 += 1

        props.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        props.dict_bereiche.update({'ordinal':ordinal_dict})
        props.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        self.mb.class_Projekt.erzeuge_dict_ordner(tab_name)

        
    def lade_tabs(self):
        if self.mb.debug: log(inspect.stack)
                                   
        try:
            gespeicherte_tabs = self.get_gespeicherte_tabs()
            
            for tab_name in gespeicherte_tabs:

                T.AB = tab_name
                
                self.erzeuge_props(tab_name)
                Eintraege = self.get_tab_Eintraege(tab_name)        

                win = self.erzeuge_neuen_tab(tab_name)  
                self.Hauptfelder[tab_name].setVisible(False)
            
            T.AB = 'ORGANON'
            
        except:
            log(inspect.stack,tb())

    
    def get_erste_datei(self,tab_name):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[tab_name].xml_tree
        root = tree.getroot()
        
        ordinale = []
        self.mb.class_XML.get_tree_info(root,ordinale)

        erste_datei = self.prj_ord
        
        for o in ordinale:
            if o[2] == 'Papierkorb':
                return erste_datei
            elif o[4] == 'pg':
                return o[0]
            
        return erste_datei
    
    def get_tab_Eintraege(self,tab_name):
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
    
    def setze_selektierte_zeile(self,ordinal,tab_name):
        if self.mb.debug: log(inspect.stack)
        try:
            
            zeile = self.mb.props[tab_name].Hauptfeld.getControl(ordinal)  
            self.mb.props[tab_name].selektierte_zeile = ordinal 
            self.mb.props[tab_name].selektierte_zeile_alt = ordinal
            textfeld = zeile.getControl('textfeld') 
            textfeld.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE
        except:
            log(inspect.stack,tb())        
        
        
from com.sun.star.awt import XMouseListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT 
    
class Tab_Leiste_Listener (unohelper.Base, XMouseListener):
    def __init__(self,mb,Hauptfelder,win=None):
        if mb.debug: log(inspect.stack)  
        
        self.mb = mb 
        self.Hauptfelder = Hauptfelder
        self.sichtbar = 'ORGANON'
        self.win = win
        
    def mousePressed(self,ev):
        if self.mb.debug: log(inspect.stack)

        tabsX = self.mb.tabsX
        tab_name = ev.Source.Model.Label
        tabsX.tab_umschalten(tab_name)
            
    def mouseReleased(self,ev):pass 
    def mouseEntered(self,ev):pass 
    def mouseExited(self,ev):pass
    def disposing(self,ev):pass                       



