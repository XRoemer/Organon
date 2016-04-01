# -*- coding: utf-8 -*-

import unohelper


class Suche():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
        self.state = {
                    'rb1_alles' : 1,
                    'rb1_sichtbare' : 0,
                    'rb1_auswahl' : 0,
                           
                    'auswahl' : [],
                    'suchbegriffe' : '',
                    
                    'alle' : 0,
                    'regex' : 0,
                    'klein_gross' : 0,
                    'ganzes_wort' : 0,
                    }
        
        self.state2 = {
                    'neuer_tab' : 1,
                    'tab_name' : '',
                    
                    'funde_taggen' : 0,
                    'tag_name' : '',
                    'kat_auswahl' : 0,
                    'in_vh_tags_eintragen' : 0,
                    
                    'mark_funde' : 0,
                    'mark_farbe' : 16275544
                    }
        
        self.titel_hoehe = 25
        self.einrueckung = 0
        
        self.dict_suche = {}
        self.dict_funde = {}
                    
            
    def dialog_elemente_suche(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            state = self.state            
            
            state['tab_name'] = LANG.TAB + ' {}'.format(len(self.mb.props) + 1)
            info_url = 'vnd.sun.star.extension://xaver.roemers.organon/img/Organon Icons/info.png'

            tags = self.mb.tags
            tags_items = tuple(v[0] for v in tags['nr_name'].values() if v[1] == 'tag')
            
            y = 0
            
            controls = (
                # Wird nicht angezeigt, dient nur der Tabberechnung
                ('xxxy',"FixedText",0,        
                                        'tab0',0,25,0,  
                                        (),
                                        (),       
                                        {} 
                                        ),
                0,
                ('rb1_alles',"RadioButton",1,     
                                        'tab0',y,120,22, 
                                        ('Label','VerticalAlign','State'),
                                        (LANG.ALLES,1,state['rb1_alles']),              
                                        {} 
                                        ),    
                0,
                ('info',"ImageButton",0,     
                                        'tab4+20',y,16,16, 
                                        ('ImageURL','Border'),
                                        (info_url,0),              
                                        {'addMouseListener':self.listener_suche} 
                                        ),
                0,                                                  
                ('rb1_sichtbare',"RadioButton",1,    
                                        'tab1',y,120,22,  
                                        ('Label','VerticalAlign','State'),
                                        (LANG.SICHTBARE,1,state['rb1_sichtbare']),       
                                        {} 
                                        ),  
                0,
                ('rb1_auswahl',"RadioButton",1,     
                                        'tab2',y,120,22,
                                        ('Label','VerticalAlign','State'),
                                        (LANG.AUSWAHL,1,state['rb1_auswahl']),      
                                        {} 
                                        ),  
                20,                                                      
                ('auswahl',"Button",1,           
                                        'tab2',y,70,20,
                                        ('Label',),
                                        (LANG.AUSWAHL,),                                 
                                        {'setActionCommand':'auswahl','addActionListener':self.listener_suche} 
                                        ), 
                
                ######################################################################################################
                20,
                ('suchbegriffe_tit',"FixedText",1,                      
                                        'tab0x',y ,50,20,        
                                        ('Label','FontWeight'), 
                                        (LANG.SUCHBEGRIFFE ,150),                 
                                        {}                      
                                        ),
                20,
                ('suchbegriffe',"Edit",1,       
                                        'tab0x-tab3-E',y ,150,25,
                                        ('HelpText',),
                                        (LANG.SUCHBEGRIFFE_HELP_TEXT,),        
                                        {} 
                                        ),
                30, 
                ('alle',"Button",0,     
                                        'tab3', 0, 20, 20,
                                        ('Label','HelpText'),
                                        (u'\u039B' if state['alle'] else 'V' , LANG.TAB_HELP_TEXT),
                                        {'setActionCommand':'V','addActionListener':self.listener_suche} 
                                        ),
                0, 
                ('klein_gross',"CheckBox",1,     
                                        'tab0x+10',y,110,22,
                                        ('Label', 'State'),
                                        (LANG.GROSS_KLEIN, state['klein_gross']),           
                                        {} 
                                        ),
                20,
                ('ganzes_wort',"CheckBox",1,     
                                        'tab0x+10',y,110,22,
                                        ('Label', 'State'),
                                        (LANG.GANZES_WORT, state['ganzes_wort']),           
                                        {} 
                                        ),
                25,
                ('regex',"CheckBox",1,      
                                        'tab0x+10',y,110,22, 
                                        ('Label', 'State'),
                                        (LANG.REGEX, state['regex']),     
                                        {} 
                                        ), 
                30,
                ('suche',"Button",1,       
                                        'tab0x',y,110,25,
                                        ('Label',),
                                        (LANG.SUCHE,),     
                                        {'setActionCommand':'suchen','addActionListener':self.listener_suche}
                                        ),
                20,
                )
            
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 10),
                     1 : (None, 10),
                     2 : (None, 0),
                     3 : (None, 0),
                     4 : (None, 0),
                     }
            
            abstand_links = 10
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite
        except:
            log(inspect.stack,tb())  
            
            
    def dialog_elemente_bearbeiten(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:

            state = self.mb.class_Suche.state2
            
            listener = Suche_Dialog_Listener(self.mb)
            listener.state = state
            
            state['tab_name'] = LANG.TAB + ' {}'.format(len(self.mb.props) + 1)

            tags = self.mb.tags
            tags_items = tuple(v[0] for v in tags['nr_name'].values() if v[1] == 'tag')
                        
            y = 0
            
            controls = (
                # Wird nicht angezeigt, dient nur der Tabberechnung
                ('xxxy',"FixedText",0,        
                                        'tab0',0,25,0,  
                                        (),
                                        (),       
                                        {} 
                                        ),
                10,  
                
                ######################################################################################################
                ('liste',"ListBox",0,     
                                        'tab0x-tab0-E',y,110,24,
                                        ('Border','Dropdown','LineCount','SelectedItem'),
                                        ( 0,True,16,1),       
                                        {} 
                                        ),
                4,
                ('auswahl_txt',"FixedText",1,        
                                        'tab1',0,25,24,  
                                        ('Label',),
                                        (LANG.AUSWAHL,),       
                                        {} 
                                        ),
                26,
                ('controlFL5',"FixedLine",0,      
                                        'tab0x-max',y ,40,1,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                10,              
                
                ######################################################################################################
                ('neuer_tab',"Button",1,     
                                        'tab0-tab0-E',y,110,24,
                                        ('Label', 'State'),
                                        (LANG.IN_NEUEM_TAB, state['neuer_tab']),           
                                        {'setActionCommand':'neuer_tab','addActionListener':listener}
                                        ),
                0,
                ('tab_name',"Edit",0,           
                                        'tab1-tab1-E',y ,100,20, 
                                        ('Text',),
                                        (state['tab_name'],),               
                                        {} 
                                        ),
                30,
                ('controlFL3',"FixedLine",0,      
                                        'tab0x-max',y ,40,1,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                10,
                ('funde_taggen',"Button",1,     
                                        'tab0-tab0-E',y,110,24,
                                        ('Label','State'),
                                        (LANG.FUNDE_TAGGEN, state['funde_taggen']),           
                                        {'setActionCommand':'funde_taggen','addActionListener':listener}
                                        ),
                30,
                ('in_vh_tags_eintragen',"CheckBox",1,     
                                        'tab0',y,110,22,
                                        ('Label','State'),
                                        (LANG.IN_VORHANDENE_TAGS_EINTRAGEN, state['in_vh_tags_eintragen']),           
                                        {} 
                                        ),
                25,
                ('tag_name',"Edit",0,           
                                        'tab0',y ,100,20, 
                                        (),
                                        (),               
                                        {} 
                                        ),
                0,
                ('kat_neu_name',"Edit",0,           
                                        'tab1-tab1-E',y ,100,20, 
                                        (),
                                        (),               
                                        {} 
                                        ),
                25,
                ('kat_auswahl',"ListBox",0,          
                                        'tab0',y ,100,20, 
                                        ('Border','Dropdown','LineCount'),
                                        ( 2,True,10),                         
                                        {'addItems':(tags_items,0),'SelectedItems':(state['kat_auswahl'],)} 
                                        ),
                0,
                ('neue_kat',"Button",1,       
                                        'tab1',y,20,22,
                                        ('Label',),
                                        (LANG.NEUE_KATEGORIE,),     
                                        {'setActionCommand':'neue_kat','addActionListener':listener}
                                        ),
                30,
                ('controlFL4',"FixedLine",0,      
                                        'tab0x-max',y ,40,1,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                10,
                ('mark_funde',"Button",1,     
                                        'tab0-tab0-E',y,110,24,
                                        ('Label',),
                                        (LANG.MARK_FUNDE,),           
                                        {'setActionCommand':'mark_funde','addActionListener':listener}
                                        ),
                0,
                ('farbe',"FixedText",0,        
                                        'tab1',0,40,24,  
                                        ('BackgroundColor','Label','Border'),
                                        (state['mark_farbe'],'    ',1),       
                                        {'addMouseListener':listener} 
                                        ),
                0,
                ('unmark_funde',"Button",0,     
                                        'tab1+52',y,24,24,
                                        ('Label',),
                                        ('X',),           
                                        {'setActionCommand':'unmark_funde','addActionListener':listener}
                                        ),
                20,
                )
            
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 20),
                     1 : (None, 20),
                     2 : (None, 0),
                     }
            
            abstand_links = 10
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite
        except:
            log(inspect.stack,tb())  
    
    
    def erzeuge_panel(self, name, suche = False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
#             if self.mb.programm == 'LibreOffice':
#                 y1,dicke = 0, 1
#             else:
#                 y1,dicke = -4, 10
            
            farben = self.mb.settings_orga['organon_farben']
            
            cont, model1 = self.mb.createControl(self.mb.ctx, "Container", 0, 0, self.breite_sidebar, 0, (), ()) 
            
            cont_titelleiste, model1 = self.mb.createControl(self.mb.ctx, "Container", 0, 0, self.breite_sidebar, self.titel_hoehe, 
                                                             ('BackgroundColor','TextColor'), 
                                                             (farben['menu_hintergrund'], farben['menu_schrift'])) 
            cont.addControl(name,cont_titelleiste)
            
            line, model = self.mb.createControl(self.mb.ctx, "FixedLine", 0, 0, self.breite_sidebar , 1, (), ()) 
            cont_titelleiste.addControl('Line',line)
                                            
            cont_icon, model = self.mb.createControl(self.mb.ctx, "ImageControl", 8,9 , 10, 10,
                                                     ('ImageURL', 'Border'), (self.url_collapse, 0) )  
            # falls kein icon in der graphic repository vorhanden ist
            if cont_icon.PreferredSize.Height == 0:
                model.ImageURL = KONST.IMG_MINUS
                
            cont_icon.addMouseListener(self.collapse_listener)
            cont_titelleiste.addControl('icon',cont_icon)  
            
            if suche:
                but, model = self.mb.createControl(self.mb.ctx, "Button", 30, 5, 14, 14, (), ()) 
                cont_titelleiste.addControl('delete',but)
                listener = Suchergebnisse_Listener(self.mb)
                but.addActionListener(listener)
                but.setActionCommand('sucheloeschen_' + name)
                pos = 50
            else:
                pos = 30
            
            
            cont_label, model = self.mb.createControl(self.mb.ctx, "FixedText", pos, 0 , self.breite_sidebar, self.titel_hoehe,
                                                      ('Label','VerticalAlign','FontWeight'), 
                                                      (name, 1, 150 ) 
                                                      )  
            cont_titelleiste.addControl('label',cont_label)
            
            cont_feld, model1 = self.mb.createControl(self.mb.ctx, "Container", 10, self.titel_hoehe , self.breite_sidebar, 0, (), ()) 
            cont.addControl(name + '_feld', cont_feld)
            
            odic = {
                    name: {
                           'container' : cont,
                           'titel' : cont_label,
                           'feld' : cont_feld,
                           'offen' : True,
                           'hoehe' : 0,
                           }
                    }
            
            return odic
    
        except:
            log(inspect.stack,tb())
            
            
    def setze_panel_pos(self, odic, y, feld_hoehe):
        if self.mb.debug: log(inspect.stack)
        
        cont = odic['container']
        feld = odic['feld']

        cont.setPosSize(0, y, 0, feld_hoehe + self.titel_hoehe, 10)
        feld.setPosSize(0, 0, 0, feld_hoehe, 8)
        
        odic['hoehe'] = feld_hoehe
        
        return feld_hoehe + self.titel_hoehe
            
            
    def erzeuge_sb_search_layout(self):
        if self.mb.debug: log(inspect.stack)
        
        dict_sb = self.mb.dict_sb
        
        if not dict_sb['design_gesetzt']:
            # faerbt nur noch die Schrift in der Menuleiste 
            if self.mb.settings_orga['organon_farben']['design_office']:
                dict_sb['setze_sidebar_design']()

        try:
            
            xUIElement = dict_sb['controls']['organon_search'][0]
            sb = dict_sb['controls']['organon_search'][1]
            self.dict_suche['xUIElement'] = xUIElement
            self.dict_suche['sb'] = sb
            panelWin = xUIElement.Window           
            
            self.url_expand = xUIElement.Theme.Image_Expand
            self.url_collapse = xUIElement.Theme.Image_Collapse

            orga_sb,seitenleiste = self.mb.class_Sidebar.get_seitenleiste('Organon Search')
             
            try:
                self.breite_sidebar = seitenleiste.PosSize.Width
            except:
                self.breite_sidebar = 400
                
                
            self.breite_sidebar = 1000
            
            self.collapse_listener = Search_Collapse_Button_Listener(self.mb)
        
        
            # CONTAINER
            container, model0 = self.mb.createControl(self.mb.ctx, "Container", 0, 0, self.breite_sidebar, 0, (), ())  
            self.dict_suche['container'] = container 
             
                
            # SUCHEN    
            self.listener_suche = Suche_Dialog_Listener(self.mb)
            #self.listener_suche.state = self.state
               
            ctrls_suche, max_breite = self.mb.class_Suche.dialog_elemente_suche()
            self.ctrls_suche,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(ctrls_suche) 
            
            odic = self.erzeuge_panel('Suchen')
            hoehe = self.setze_panel_pos(odic['Suchen'], 0, max_hoehe)
            self.dict_suche.update(odic)

            for name,c in sorted(self.ctrls_suche.items()):
                self.dict_suche['Suchen']['feld'].addControl(name,c)
            
            container.addControl('Suchen',self.dict_suche['Suchen']['container'])

            # BEARBEITEN
            ctrls_bearbeiten, max_breite = self.mb.class_Suche.dialog_elemente_bearbeiten()
            self.ctrls_bearbeiten,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(ctrls_bearbeiten) 
            
            odic = self.erzeuge_panel('Bearbeiten')
            hoehe2 = self.setze_panel_pos(odic['Bearbeiten'], hoehe, max_hoehe)
            self.dict_suche.update(odic)
            
            for name,c in sorted(self.ctrls_bearbeiten.items()):
                self.dict_suche['Bearbeiten']['feld'].addControl(name,c)
                
            container.addControl('Bearbeiten',self.dict_suche['Bearbeiten']['container'])
                
            # ERGEBNISSE
            odic = self.erzeuge_panel('Ergebnisse')
            hoehe3 = self.setze_panel_pos(odic['Ergebnisse'], hoehe + hoehe2, 0)
            self.dict_suche.update(odic)
            self.dict_suche['Ergebnisse']['fundstellen'] = {}
            container.addControl('Ergebnisse',self.dict_suche['Ergebnisse']['container'])
            
            panelWin.addControl('Container' , container)

            xUIElement.height = hoehe + hoehe2 + hoehe3 
            container.setPosSize(0,0,0, hoehe + hoehe2 + hoehe3, 8)
                        
            sb.requestLayout()
            
            self.mb.class_Suche.suchergebnisse_laden()
            
        except:
            log(inspect.stack,tb())       
            
    
    def suche_loeschen(self, suchbegriffe):        
        if self.mb.debug: log(inspect.stack)
        
        try:            
            dict_suche = self.mb.class_Suche.dict_suche
            dict_fundstelle = dict_suche['Ergebnisse']['fundstellen'][suchbegriffe]
            panels = [ f for f in dict_suche['Ergebnisse']['fundstellen'].values() 
                      if f['container'].PosSize.Y > dict_fundstelle['container'].PosSize.Y]
            
            hoehe = dict_fundstelle['container'].PosSize.Height
            
            setPosSize = lambda x,y : x.setPosSize(0,0,0, x.PosSize.Height + y, 8 )
            setPosSize2 = lambda x,y : x.setPosSize(0, x.PosSize.Y + y, 0,0,2 )
            
            dict_suche['Ergebnisse']['hoehe'] -= hoehe
            setPosSize( dict_fundstelle['container'], - hoehe )
            
            dict_suche['xUIElement'].height -= hoehe
            
            for p in panels:
                setPosSize2(p['container'], - hoehe)

            dict_fundstelle['container'].dispose()
            del dict_suche['Ergebnisse']['fundstellen'][suchbegriffe]
            del self.dict_funde[ suchbegriffe ]
            
            dict_suche['sb'].requestLayout()
            self.suchergebnisse_speichern()
        except:
            log(inspect.stack,tb())
            
        
    def suchergebnisse_speichern(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad = os.path.join( self.mb.pfade['settings'], 'searches.json' )
            with open(pfad, 'w',) as outfile:
                json.dump(self.dict_funde, outfile, indent=4, separators=(',', ': '))                     
        except:
            log(inspect.stack,tb())
        
        
    def suchergebnisse_laden(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            if self.dict_funde == {}:
                try:
                    pfad = os.path.join( self.mb.pfade['settings'], 'searches.json' )
                    if os.path.exists(pfad):
                        
                        dict_funde = self.mb.class_Funktionen.oeffne_json(pfad)
                        
                        if dict_funde != None:
                            self.dict_funde = dict_funde
                        else:
                            dict_funde = {}
                except:
                    self.dict_funde = {}
            
            if self.dict_funde != {}:
                
                zu_loeschen = []
                
                for sw,odict in self.dict_funde.items():
                    if odict['state']['regex']:
                        ordis = list(odict['ergebnisse'])
                    else:
                        ordis = set( [item for row in odict['ergebnisse'] for item in odict['ergebnisse'][row]] )   
                        
                    if len(ordis) == 0:
                        zu_loeschen.append(sw)
                        continue      
                       
                    ord_sortiert = self.mb.class_Tabs.sortiere_ordinale(ordis)     
                                   
                    self.listener_suche.ergebnisse_anzeigen(ord_sortiert, odict['ergebnisse'], sw, odict['state']['regex'])
                
                for z in zu_loeschen:
                    del self.dict_funde[z]
                
        except:
            log(inspect.stack,tb())
              
        
from com.sun.star.awt import XActionListener, XMouseListener    
class Suche_Dialog_Listener(unohelper.Base, XActionListener, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
        self.ordinale_ausgewaehlte = []
        self.suche = self.mb.class_Suche

    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        if ev.Source == self.suche.ctrls_bearbeiten['farbe']:
            farbe = self.mb.class_Funktionen.waehle_farbe()
            self.state['mark_farbe'] = farbe
            ev.Source.Model.BackgroundColor = farbe
        elif ev.Source == self.suche.ctrls_suche['info']:
            self.oeffne_suche_info()
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            cmd = ev.ActionCommand
            
            if cmd == 'auswahl':
                self.erzeuge_Fenster_fuer_eigene_Auswahl(ev)
                
            elif cmd == 'V':
                label = ev.Source.Model.Label
                ev.Source.Model.Label = u'\u039B' if label == 'V' else 'V'
                self.state['alle'] = 1 if label == 'V' else 0
                
            elif cmd == 'suchen':
                ok, sw = self.dialog_suchen_auswerten()
                if ok:
                    self.suchen(sw)
            
            
            else:
                
                ctrls = self.suche.ctrls_bearbeiten
                term = ctrls['liste'].SelectedItem
                
                if term == '':
                    Popup(self.mb, 'warning').text = LANG.KEINE_SUCHE_AUSGEWAEHLT
                    return
                
                funde = self.suche.dict_funde[term] 
                ordis = set( [item for row in funde['ergebnisse'] for item in funde['ergebnisse'][row]] ) 
                
                if cmd == 'neuer_tab':
                    
                    ord_sortiert = self.mb.class_Tabs.sortiere_ordinale(ordis)
                    if self.mb.props[T.AB].Projektordner in ord_sortiert:
                        ord_sortiert.remove(self.mb.props[T.AB].Projektordner)

                    tab_name = ctrls['tab_name'].Model.Text
                    
                    if tab_name in self.mb.props:
                        Popup(self.mb, 'warning').text = LANG.TAB_EXISTIERT_SCHON
                        return
                        
                    self.mb.tabsX.erzeuge_neuen_tab2(ord_sortiert, tab_name=tab_name)
                 
                elif cmd == 'neue_kat':
                    
                    name = ctrls['kat_neu_name'].Model.Text.strip()
                    
                    if name in self.mb.tags['name_nr']:
                        Popup(self.mb, 'warning').text = LANG.KATEGORIE_EXISTIERT.format(name)
                        return
                    
                    self.neue_tag_kat_anlegen(name)
                    ctrls['kat_auswahl'].addItem(name,0)
                    ctrls['kat_auswahl'].selectItem(name,True)
                
                elif cmd == 'funde_taggen':
                    self.funde_taggen(ordis, funde, ctrls )
                    
                elif cmd == 'mark_funde':
                    farbe = self.suche.ctrls_bearbeiten['farbe'].Model.BackgroundColor
                    self.fundstellen_markieren(list(ordis), funde, ctrls, farbe)
                    
                elif cmd == 'unmark_funde':
                    farbe = -1
                    self.fundstellen_markieren(list(ordis), funde, ctrls, farbe)
                                    
        except:
            log(inspect.stack,tb())


    def mouseEntered(self, ev):
        return False
    def mouseExited(self, ev):  
        return False
    def mouseReleased(self,ev):
        return False
    def disposing(self,ev):
        return False            
            
    
    def erzeuge_Fenster_fuer_eigene_Auswahl(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            parent = ev.Source.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleParent
            posSize = ev.Source.AccessibleContext.AccessibleParent.PosSize
            
            x = posSize.X + posSize.Width + 20
            y = posSize.Y
            pos = x,y
            
            # Auswahl korrespondiert mit Auswahl fuer Export
            sett = self.mb.settings_exp
     
            # Dict von alten Eintraegen bereinigen
            sett['ausgewaehlte'] = { ordi:v for ordi,v in sett['ausgewaehlte'].items() 
                                    if ordi in self.mb.props[T.AB].dict_bereiche['ordinal'] }

            (y,
             fenster,
             fenster_cont,
             control_innen,
             ctrls) = self.mb.class_Fenster.erzeuge_treeview_mit_checkbox(
                                                                    tab_name=T.AB,
                                                                    pos=pos,
                                                                    auswaehlen=True,
                                                                    parent=parent
                                                                    )
      
        except:
            log(inspect.stack,tb())
            

    def dialog_suchen_auswerten(self):
        if self.mb.debug: log(inspect.stack)

        try:
            
            ctrls = self.mb.class_Suche.ctrls_suche
            state = {}
            
            ausgewaehlte = [ordi for ordi,v in self.mb.settings_exp['ausgewaehlte'].items() if v]
            
            
            state['rb1_alles'] = int(ctrls['rb1_alles'].State)
            state['rb1_sichtbare'] = int(ctrls['rb1_sichtbare'].State)
            state['rb1_auswahl'] = int(ctrls['rb1_auswahl'].State)
                   
            state['auswahl'] = ausgewaehlte
                        
            state['regex'] = ctrls['regex'].State
            state['klein_gross'] = ctrls['klein_gross'].State
            state['ganzes_wort'] = ctrls['ganzes_wort'].State
            
            state['alle'] = 1 if ctrls['alle'].Model.Label == 'V' else 0
            
            
            if state['rb1_auswahl'] and ausgewaehlte == []:
                Popup(self.mb, 'info').text = LANG.KEINE_DATEIEN_AUSGEWAEHLT
                return False
            
            self.ordinale_ausgewaehlte = self.get_ordinale_ausgewaehlte(ausgewaehlte, state)
            
            if not self.ordinale_ausgewaehlte:
                Popup(self.mb, 'info').text = LANG.KEINE_DATEIEN_AUSGEWAEHLT
                return False
            
            elif ctrls['suchbegriffe'].Model.Text == '':
                Popup(self.mb, 'info').text = LANG.KEINE_SUCHBEGRIFFE
                return False, None            

            term = self.mb.class_Suche.ctrls_suche['suchbegriffe'].Model.Text
                     
            if state['regex']:
                state['suchbegriffe'] = term
                sw = term
            else:
                state['suchbegriffe'] = self.get_suchbegriffe(term)
                sw = ', '.join(state['suchbegriffe'])
                
            while sw in self.mb.class_Suche.dict_funde:
                sw = sw + 'I'
            
            self.mb.class_Suche.dict_funde[ sw ] = {'state' : state}
            
            return True, sw
        except:
            log(inspect.stack,tb())
            return False, None
        
        
    def get_suchbegriffe(self, term):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if term[-1] == ',':
                term = term[:len(term)-1]
            
            begriffe = [a.strip() for a in term.split(',')]

            verbinden = []
            for b in begriffe:
                if b[-1] == '\\':
                    verbinden.append(b)

            for v in reversed(verbinden):
                index = begriffe.index(v)
                begriffe[index] = ', '.join([begriffe[index].replace('\\',''), begriffe[index + 1]])
                del begriffe[index + 1]
            
            return begriffe
        except:
            log(inspect.stack,tb())
            
    
    def get_ordinale_ausgewaehlte(self, ausgewaehlte, state):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props[T.AB]
        
        if state['rb1_auswahl']:
            return ausgewaehlte
        
        elif state['rb1_alles']:
            return [ordi for ordi in props.dict_bereiche['ordinal'] if ordi not in props.dict_ordner[props.Papierkorb] ]
        
        elif state['rb1_sichtbare']:
            ordis = [props.dict_bereiche['Bereichsname-ordinal'][name] for name in self.mb.sichtbare_bereiche]
            return [o for o in ordis if o in props.dict_bereiche['ordinal'] and o not in props.dict_ordner[props.Papierkorb] ]


    def plain_txt_auslesen(self,ordinale):
        if self.mb.debug: log(inspect.stack)
        
        try:      
            texte = {}
            
            for name in ordinale:
                path = os.path.join(self.mb.pfade['plain_txt'], name + '.txt')
                
                with codecs_open(path , "r","utf-8") as f:
                    texte.update( { name: f.read() } )
             
            return texte  
        except:
            log(inspect.stack,tb())
            return {}
    
    
    def suchen(self, sw):
        if self.mb.debug: log(inspect.stack)
        
        try:
            state = self.mb.class_Suche.dict_funde[sw]['state']        
            plain_text = self.plain_txt_auslesen(self.ordinale_ausgewaehlte)
            
            if state['regex']:
                ergebnisse = self.suchen_regex(plain_text,state['suchbegriffe'])
                alle_funde = list(ergebnisse)
                gemeinsam = alle_funde
            else:
                ergebnisse, gemeinsam = self.suchen_python(plain_text,state)
                alle_funde = list( set( [x for a in ergebnisse.values() for x in a] ) )
                
                    
            if state['regex']:
                ordis = alle_funde
            
            elif not state['alle']:
                ordis = gemeinsam
            else:
                ordis = alle_funde
            
            if not ordis:
                Popup(self.mb, 'info').text = LANG.SUCHE_KEINE_ERGEBNISSE
                return
            
            ord_sortiert = self.mb.class_Tabs.sortiere_ordinale(ordis)
            self.ergebnisse_anzeigen(ord_sortiert, ergebnisse, sw, state['regex'])
            self.mb.class_Suche.suchergebnisse_speichern()
        except:
            log(inspect.stack,tb())

    
    def ergebnisse_anzeigen(self,ord_sortiert,ergebnisse,titel,regex=False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            dict_suche = self.mb.class_Suche.dict_suche         
            dict_funde = self.mb.class_Suche.dict_funde
               
            props = self.mb.props[T.AB]
            root = props.xml_tree.getroot()
            dateiname = { o : root.find('.//' + o).attrib['Name'] for o in ord_sortiert}
            
            listener = Suchergebnisse_Listener(self.mb)
            state = self.mb.class_Suche.state
            
            funde = { o : ['dir',[]] if o in props.dict_ordner else ['pg',[]] for o in ord_sortiert}

            
            if regex:                
                for ordi, sw in ergebnisse.items():
                    if ordi in funde:
                        funde[ordi][1] = sw
            else:                
                for sw,ordis in ergebnisse.items():
                    for o in ordis:
                        if o in funde:
                            funde[o][1].append(sw)
            
            
            dict_funde[titel]['ergebnisse'] = ergebnisse
            
            controls = []
            
            for o in ord_sortiert:                    
                
                if o in props.dict_ordner:
                    farbe = self.mb.settings_orga['organon_farben']['schrift_ordner']
                else:
                    farbe = self.mb.settings_orga['organon_farben']['schrift_datei']
                
                controls.extend([
                4,
                (o+'_X',"Button",0,       
                                'tab0',0,14,14,
                                (),
                                (),     
                                {'setActionCommand':o,'addActionListener':listener}
                                ),
                0,
                (o,"FixedText",1,       
                                'tab1',0,110,20,
                                ('Label','TextColor'),
                                (dateiname[o], farbe),     
                                {}
                                ),
                20,                
                ])
            
                for sw in funde[o][1]:

                    controls.extend([
                    
                    (o+'_' + sw + '_' + titel,"FixedText",1,       
                                    'tab1+20',0,110,20,
                                    ('Label','Border'),
                                    (sw,2),     
                                    {'addMouseListener':listener}
                                    ),
                    20,
                                    ])
            
                controls.extend([4])
            
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 5),
                     1 : (None, 0),
                     }
            
                        
            abstand_links = 0
            controls2, tabs3, max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
            ctrls, hoehe_feld = self.mb.class_Fenster.erzeuge_fensterinhalt(controls2) 
            
            hoehe_cont = hoehe_feld + self.mb.class_Suche.titel_hoehe
            pos = dict_suche['Ergebnisse']['hoehe']
            
            odic = self.mb.class_Suche.erzeuge_panel(titel, True)
                        
            odic[titel]['hoehe'] = hoehe_feld
            dict_suche['Ergebnisse']['fundstellen'].update(odic)
            
            setPosSize = lambda x,y : x.setPosSize(0,0,0, x.PosSize.Height + y, 8 )
            
            setPosSize(dict_suche['container'], hoehe_cont)
            setPosSize(dict_suche['Ergebnisse']['container'], hoehe_cont) 
            setPosSize(dict_suche['Ergebnisse']['feld'], hoehe_cont)
            dict_suche['Ergebnisse']['hoehe'] += hoehe_cont 
            dict_suche['xUIElement'].height += hoehe_cont 
                                    
            # Controls in Hauptfenster eintragen
            for name,c in sorted(ctrls.items()):
                dict_suche['Ergebnisse']['fundstellen'][titel]['feld'].addControl(name,c)
            
            self.mb.class_Suche.setze_panel_pos(odic[titel], pos, hoehe_feld)
            dict_suche['Ergebnisse']['feld'].addControl(titel,dict_suche['Ergebnisse']['fundstellen'][titel]['container'])
           
            listener.ctrls = ctrls
            listener.suchbegriffe = titel   
                
            dict_suche['sb'].requestLayout()
            
            list_ctrl = self.suche.ctrls_bearbeiten['liste']
            list_ctrl.addItem(titel,list_ctrl.ItemCount)
            list_ctrl.selectItem(titel,True)
        except:
            log(inspect.stack,tb())

    
    def funde_taggen(self, ordis, funde, ctrls):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ergebnisse = funde['ergebnisse']
            state = funde['state']
            
            tag_name = ctrls['tag_name'].Model.Text
            in_vh_tags_eintragen = ctrls['in_vh_tags_eintragen'].State
            
            
            if state['regex'] and tag_name == '':
                Popup(self.mb, 'info').text = LANG.REGEX_BRAUCHT_TAGNAMEN
                return 

            alle_tags = [a for b,v in self.mb.tags['sammlung'].items() for a in v] 
           
            if not in_vh_tags_eintragen: 
                
                doppelter_tag = None
                
                if tag_name in alle_tags and tag_name != '':
                    doppelter_tag = tag_name
                else:
                    if not state['regex'] and tag_name == '':
                        for s in state['suchbegriffe']:
                            if s in alle_tags:
                                doppelter_tag = s
                                break

                if doppelter_tag:
                    Popup(self.mb, 'warning').text = LANG.TAG_VORHANDEN_WAEHLEN.format(doppelter_tag)
                    return False
            
            
            
            tags = self.mb.tags            
            alle_tags = [a for b,v in self.mb.tags['sammlung'].items() for a in v]
            
            
            # nur ein neuer Tag wird eingetragen
            if tag_name:
                
                new_tag = tag_name
                
                if state['regex']:
                    ergebnisse = {'xxx':list(ergebnisse)}
                
                if new_tag in alle_tags:
                    panel_nr = [nr for nr,v in tags['sammlung'].items() if new_tag in v][0]
                else:
                    kat_name = ctrls['kat_auswahl'].SelectedItem
                    panel_nr = tags['name_nr'][kat_name]
                    
                for swort,ords in ergebnisse.items():
                    for o in ords: 
                        if new_tag not in tags['ordinale'][o][panel_nr]:
                            tags['ordinale'][o][panel_nr].append(new_tag)
                    
                    if new_tag not in tags['sammlung'][panel_nr]:
                        tags['sammlung'][panel_nr].append(new_tag)
                      
            else:
                # verschiedene Tags werden eingetragen
                vorhandene = {}
                if in_vh_tags_eintragen:
                    
                    zu_loeschende = []
                    
                    for new_tag,ordis in ergebnisse.items():
                        for panel_nr,ts in tags['sammlung'].items():
                            if new_tag in ts:
                                vorhandene.update({panel_nr: [ordis,new_tag] })
                                zu_loeschende.append(new_tag)
                                break
                    for z in zu_loeschende:
                        del ergebnisse[z]
                
                for panel_nr,[ordis,new_tag] in vorhandene.items():
                    for o in ordis:
                        if new_tag not in tags['ordinale'][o][panel_nr]:
                            tags['ordinale'][o][panel_nr].append(new_tag)
                
                
                kategorie = ctrls['kat_auswahl'].SelectedItem
                panel_nr = tags['name_nr'][kategorie]
                
                for new_tag, ordis in ergebnisse.items():
                    for o in ordis:
                        if new_tag not in tags['ordinale'][o][panel_nr]:
                            tags['ordinale'][o][panel_nr].append(new_tag)
                        if new_tag not in tags['sammlung'][panel_nr]:
                            tags['sammlung'][panel_nr].append(new_tag)
            
            
            Popup(self.mb, 'info').text = LANG.FUNDE_WURDEN_GETAGGT
            
        except:
            log(inspect.stack,tb())
        
        
    def suchen_regex(self,texte,begriff):
        if self.mb.debug: log(inspect.stack)
        
        try:
            results = {}
            comp = re.compile(begriff)

            for ordinal,txt in texte.items():
                if comp.search(txt):
                    results.update({ordinal:[begriff,]})
            return results  
        except:
            log(inspect.stack,tb())
            return []
        
        
    def suchen_python(self,otexte,state):
        if self.mb.debug: log(inspect.stack)
        
        obegriffe = state['suchbegriffe']
        
        if state['klein_gross']:
            texte = copy.deepcopy(otexte)
            begriffe = {o:o for o in obegriffe}
        else:
            texte = {o:t.lower() for o,t in otexte.items() }
            begriffe = {o:o.lower() for o in obegriffe}

        ergebnisse = {}
        gemeinsam = set()
        
        if not state['ganzes_wort']:
            
            for wort,wort_lower in begriffe.items():
                ergebnisse[wort] = []
                
                for ordinal,txt in texte.items():
                    if wort_lower in txt:
                        ergebnisse[wort].append(ordinal)
        else:
            for wort,wort_lower in begriffe.items():
                ergebnisse[wort] = []
                
                for ordinal,txt in texte.items():
                    if re.search(r'\b' + re.escape(wort_lower) + r'\b', txt):
                        ergebnisse[wort].append(ordinal)
        
        if not state['alle']:
            
            alle_funde = [x for x in ergebnisse.values()]
            gemeinsam = set(alle_funde[0])
            for b in alle_funde[1:]:
                gemeinsam = gemeinsam.intersection(b)

        return ergebnisse, list(gemeinsam)
    
        
    def fundstellen_markieren(self, ordinale, funde, ctrls, farbe): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            ergebnisse = funde['ergebnisse']
            state = funde['state']
            props = self.mb.props[T.AB]
            
            anzahl = len(ordinale)
            StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
            StatusIndicator.start(LANG.ERZEUGE_NEUE_ZEILE,anzahl)
            StatusIndicator.setValue(2)
            
            for ordi in ordinale:
                
                stelle = ordinale.index(ordi)
                StatusIndicator.setValue(stelle)
                StatusIndicator.setText('{0}/{1}'.format(stelle,anzahl))
                
                pfad = os.path.join(self.mb.pfade['odts'], ordi + '.odt')
                pfad_url = uno.systemPathToFileUrl(pfad)
                doc = self.mb.class_Funktionen.lade_hidden_doc(pfad_url)
                
                sd = doc.createSearchDescriptor()
                
                if state['klein_gross']:
                    sd.SearchCaseSensitive = True
                if state['ganzes_wort']:
                    sd.SearchWords = True
                
                                
                if not state['regex']:

                    for swort,ergs in ergebnisse.items():
                        if ordi in ergebnisse[swort]:
                            sd.SearchString = swort
                            funde = doc.findAll(sd)
                            
                            for e in range(funde.Count):
                                funde.getByIndex(e).CharBackColor = farbe
                
                else:
                    sd.SearchRegularExpression = True
                    sd.SearchString = state['suchbegriffe']
                    
                    funde = doc.findAll(sd)
                            
                    for e in range(funde.Count):
                        funde.getByIndex(e).CharBackColor = farbe
                
                
                doc.store()
                doc.close(True)
                
                # Damit die markierten Fundstellen sichtbar werden, muessen bereits
                # verlinkte Sections neu verlinkt werden.
                
                sec_name = props.dict_bereiche['ordinal'][ordi]
                sec = self.mb.doc.TextSections.getByName(sec_name)
                
                if sec.FileLink.FileURL != '':
                    file_link = sec.FileLink
                    path = file_link.FileURL
                    
                    fl = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                    sec.setPropertyValue('FileLink',fl)
                    
                    fl.FileURL = path
                    sec.setPropertyValue('FileLink',fl)
            
        except:
            log(inspect.stack,tb())
            try:
                doc.close(False)
            except:
                pass
            
        StatusIndicator.end()
        
        
    def neue_tag_kat_anlegen(self,name):   
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        neue_nr = len(tags['abfolge'])  
         
        tags['abfolge'].append(neue_nr)
        tags['name_nr'].update( {name : neue_nr} )
        tags['nr_breite'].update( {neue_nr : 3.0} )
        tags['nr_name'].update( {neue_nr : [name,'tag']} )
        tags['sammlung'].update( {neue_nr : []} )
        tags['sichtbare'].append( neue_nr )
         
        for o in tags['ordinale']:
            tags['ordinale'][o].update( {neue_nr : []} )
                 
    
    def oeffne_suche_info(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            lang,hb,ordi = ['de','Organon Handbuch.organon',
                            'nr80.odt'] if self.mb.language == 'de' else ['en','Organon Manual.organon','nr81.odt']
            
            url = os.path.join(self.mb.path_to_extension,'description','Handbuecher',lang,hb,'Files','odt',ordi)
            URL = uno.systemPathToFileUrl(url)
            self.mb.class_Fenster.oeffne_dokument_in_neuem_fenster(URL)
        
        except:
            log(inspect.stack,tb())

        
        
class Suchergebnisse_Listener(unohelper.Base, XActionListener, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
                
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        if 'sucheloeschen_' in ev.ActionCommand:
            suchbegriffe = ev.ActionCommand.split('_',1)[1]
            self.mb.class_Suche.suche_loeschen(suchbegriffe)
        else:
            self.suche_eintrag_loeschen(ev)
  
    def disposing(self,ev):
        return False
        
    def mouseReleased(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self,ev):
        if self.mb.debug: log(inspect.stack)
        self.suchergebnis_aufrufen(ev)
        
    def suchergebnis_aufrufen(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ordinal, tag, titel = [ n for n,c in self.ctrls.items() if c == ev.Source ][0].split('_',2)
            
            state = self.mb.class_Suche.dict_funde[titel]['state']
            
            if ordinal != self.mb.props[T.AB].selektierte_zeile:
                self.mb.class_Baumansicht.selektiere_zeile(ordinal)
            
            transliterateFlags = 1024 if state['klein_gross'] else 1280
            searchFlags = 65552 if state['ganzes_wort'] else 65536
            
            if tag:
                kwargs = {
                        'SearchItem.StyleFamily':2,
                        'SearchItem.CellType':0,
                        'SearchItem.RowDirection':True,
                        'SearchItem.AllTables':False,
                        'SearchItem.Backward':False,
                        'SearchItem.Pattern':False,
                        'SearchItem.Content':False,
                        'SearchItem.AsianOptions':False,
                        'SearchItem.AlgorithmType':state['regex'],
                        'SearchItem.SearchFlags': searchFlags,
                        'SearchItem.SearchString': tag,
                        'SearchItem.ReplaceString':"",
                        'SearchItem.Locale':255,
                        'SearchItem.ChangedChars':2,
                        'SearchItem.DeletedChars':2,
                        'SearchItem.InsertedChars':2,
                        'SearchItem.TransliterateFlags': transliterateFlags,
                        'SearchItem.Command':1,
                        'Quiet':True
                        
                        }
                frame = self.mb.doc.CurrentController.Frame    
                self.mb.dispatch(frame,'ExecuteSearch',
                                 **kwargs
                                 )
        except:
            log(inspect.stack,tb())
            
            
    def suche_eintrag_loeschen(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ordinal = ev.ActionCommand
            dict_suche = self.mb.class_Suche.dict_suche
            
            dict_fundstelle = dict_suche['Ergebnisse']['fundstellen'][self.suchbegriffe]
            ctrls = [name for name, c in self.ctrls.items() if ordinal in name]

            omax = []
            omin = []
            
            for name in ctrls:
                c = self.ctrls[name]
                omin.append(c.PosSize.Y)
                omax.append(c.PosSize.Y + c.PosSize.Height)
                c.dispose()
                del (self.ctrls[name])
            
            omin = min(omin)
            omax = max(omax)
            diff = omax-omin + 8
            
            for c in self.ctrls.values():
                if c.PosSize.Y > omax:
                    y = c.PosSize.Y
                    c.setPosSize(0, y - diff, 0, 0, 2)
            
            setPosSize = lambda x,y : x.setPosSize(0,0,0, x.PosSize.Height + y, 8 )
            setPosSize2 = lambda x,y : x.setPosSize(0, x.PosSize.Y + y, 0,0,2 )
                        
            dict_suche['Ergebnisse']['hoehe'] -= diff
            dict_fundstelle['hoehe'] -= diff
            setPosSize( dict_fundstelle['feld'], - diff )
            setPosSize( dict_fundstelle['container'], - diff )
            
            dict_suche['xUIElement'].height -= diff
            
            
            for name, odict in dict_suche['Ergebnisse']['fundstellen'].items():
                if name == self.suchbegriffe:
                    continue
                if odict['container'].PosSize.Y > dict_fundstelle['container'].PosSize.Y:
                    setPosSize2( odict['container'], - diff )
            
            dict_funde_erg = self.mb.class_Suche.dict_funde[self.suchbegriffe]['ergebnisse']
            
            for v in dict_funde_erg.values():
                if ordinal in v:
                    v.remove(ordinal)            
            
            self.mb.class_Suche.suchergebnisse_speichern()
            dict_suche['sb'].requestLayout()
        except:
            log(inspect.stack,tb())
        
        
class Search_Collapse_Button_Listener(unohelper.Base, XMouseListener):
    def __init__(self,mb):
        self.mb = mb
        
    def disposing(self,ev):
        return False
    def mouseReleased(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        source = ev.Source
        try:
            txt = source.Context.Controls[2].Text
        except:
            txt = source.Context.Controls[3].Text
            
        self.collapse_expand_panels(txt,source)


    def collapse_expand_panels(self, name, source):   
        if self.mb.debug: log(inspect.stack)
        
        try:
            dict_suche = self.mb.class_Suche.dict_suche
            
            if name == 'Suchen':
                panels = (dict_suche['Bearbeiten'], dict_suche['Ergebnisse'])
            elif name == 'Bearbeiten':
                panels = (dict_suche['Ergebnisse'],)
            elif name not in ('Suchen', 'Ergebnisse', 'Bearbeiten'):
                panels = [ dict_suche['Ergebnisse']['fundstellen'][n] for n in dict_suche['Ergebnisse']['fundstellen'] if n != name ]
            else:
                panels = ()
                            
            setPosSize = lambda x,y : x.setPosSize(0,0,0, x.PosSize.Height + y, 8 )
            setPosSize2 = lambda x,y : x.setPosSize(0, x.PosSize.Y + y, 0,0,2 )

            def klappen(odic):
                
                if odic['offen']:
                    odic['offen'] = False
                    hoehe = - odic['hoehe']
                    odic['feld'].setPosSize(0,0,0,0,8)
                else:
                    odic['offen'] = True
                    hoehe = odic['hoehe']
                    odic['feld'].setPosSize(0,0,0,hoehe,8)
                    
                setPosSize(odic['container'], hoehe)
                
                for p in panels:
                    setPosSize2(p['container'], hoehe)
            
                dict_suche['xUIElement'].height += hoehe
                setPosSize(dict_suche['container'], hoehe)
            
            
            if name in ('Suchen', 'Bearbeiten', 'Ergebnisse'):
                klappen(dict_suche[name])
            else:
                odic = dict_suche['Ergebnisse']['fundstellen'][name]
                
                if odic['offen']:
                    odic['offen'] = False
                    hoehe = - odic['hoehe']
                    odic['feld'].setPosSize(0,0,0,0,8)
                    
                else:
                    odic['offen'] = True
                    hoehe = odic['hoehe']
                    odic['feld'].setPosSize(0,0,0,hoehe,8)
                    
                    
                setPosSize(odic['container'], hoehe)
                
                for p in panels:
                    if p['container'].PosSize.Y > odic['container'].PosSize.Y:
                        setPosSize2(p['container'], hoehe)
                
                dict_suche['xUIElement'].height += hoehe
                setPosSize(dict_suche['container'], hoehe)
                setPosSize(dict_suche['Ergebnisse']['container'], hoehe)
                setPosSize(dict_suche['Ergebnisse']['feld'], hoehe)
                dict_suche['Ergebnisse']['hoehe'] += hoehe
                    
                                    
            dict_suche['sb'].requestLayout()
            
        except:
            log(inspect.stack,tb())
        
                
        
        
        
            