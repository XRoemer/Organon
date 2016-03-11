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
                    
                    'neuer_tab' : 1,
                    'tab_name' : '',
                    
                    'funde_taggen' : 0,
                    'tag_name' : '',
                    'kat_auswahl' : 0,
                    'in_vh_tags_eintragen' : 0,
                    
                    'mark_funde' : 0,
                    'mark_farbe' : 16275544
                    }
        

    def suche(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.dialog_suche()
                
        except Exception as e:
            Popup(self.mb, 'error').text = 'ERROR: '+ str(e)
            log(inspect.stack,tb())


    def dialog_suche_elemente(self,listener): 
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
                15,
                ('rb1_alles',"RadioButton",1,     
                                        'tab0x',y,120,22, 
                                        ('Label','VerticalAlign','State'),
                                        (LANG.ALLES,1,state['rb1_alles']),              
                                        {} 
                                        ),    
                0,
                ('info',"ImageButton",0,     
                                        'tab2',y,16,16, 
                                        ('ImageURL','Border'),
                                        (info_url,0),              
                                        {'addMouseListener':listener} 
                                        ),
                20,                                                  
                ('rb1_sichtbare',"RadioButton",1,    
                                        'tab0x',y,120,22,  
                                        ('Label','VerticalAlign','State'),
                                        (LANG.SICHTBARE,1,state['rb1_sichtbare']),       
                                        {} 
                                        ),  
                20,
                ('rb1_auswahl',"RadioButton",1,     
                                        'tab0x',y,120,22,
                                        ('Label','VerticalAlign','State'),
                                        (LANG.AUSWAHL,1,state['rb1_auswahl']),      
                                        {} 
                                        ),  
                0,                                                      
                ('auswahl',"Button",1,           
                                        'tab2',y,70,25,
                                        ('Label',),
                                        (LANG.AUSWAHL,),                                 
                                        {'setActionCommand':'auswahl','addActionListener':listener} 
                                        ),
                40,
                ('controlFL0',"FixedLine",0,      
                                        'tab0x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
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
                                        'tab0x-max',y ,150,25,
                                        ('HelpText',),
                                        (LANG.SUCHBEGRIFFE_HELP_TEXT,),        
                                        {} 
                                        ),
                30, 
                ('alle',"Button",0,     
                                        'tab2', 0, 20, 20,
                                        ('Label','HelpText'),
                                        (u'\u039B' if state['alle'] else 'V' , LANG.TAB_HELP_TEXT),
                                        {'setActionCommand':'V','addActionListener':listener} 
                                        ),
                0, 
                ('klein_gross',"CheckBox",1,     
                                        'tab1',y,110,22,
                                        ('Label', 'State'),
                                        (LANG.GROSS_KLEIN, state['klein_gross']),           
                                        {} 
                                        ),
                20,
                ('ganzes_wort',"CheckBox",1,     
                                        'tab1',y,110,22,
                                        ('Label', 'State'),
                                        (LANG.GANZES_WORT, state['ganzes_wort']),           
                                        {} 
                                        ),
                25,
                ('regex',"CheckBox",1,      
                                        'tab1',y,110,22, 
                                        ('Label', 'State'),
                                        (LANG.REGEX, state['regex']),     
                                        {} 
                                        ), 
                30,
                ('controlFL2',"FixedLine",0,      
                                        'tab0x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                
                ######################################################################################################
                40,
                ('neuer_tab',"CheckBox",1,     
                                        'tab0x',y,110,22,
                                        ('Label', 'State'),
                                        (LANG.IN_NEUEM_TAB, state['neuer_tab']),           
                                        {} 
                                        ),
                20,
                ('tab_name',"Edit",0,           
                                        'tab1',y ,100,20, 
                                        ('Text',),
                                        (state['tab_name'],),               
                                        {} 
                                        ),
                30,
                ('controlFL3',"FixedLine",0,      
                                        'tab1x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                20,
                ('funde_taggen',"CheckBox",1,     
                                        'tab0x',y,110,22,
                                        ('Label','State'),
                                        (LANG.FUNDE_TAGGEN, state['funde_taggen']),           
                                        {} 
                                        ),
                25,
                ('in_vh_tags_eintragen',"CheckBox",1,     
                                        'tab1x',y,110,22,
                                        ('Label','State'),
                                        (LANG.IN_VORHANDENE_TAGS_EINTRAGEN, state['in_vh_tags_eintragen']),           
                                        {} 
                                        ),
                25,
                ('tag_name',"Edit",0,           
                                        'tab1',y ,100,20, 
                                        (),
                                        (),               
                                        {} 
                                        ),
                0,
                ('kat_neu_name',"Edit",0,           
                                        'tab2-max',y ,100,20, 
                                        (),
                                        (),               
                                        {} 
                                        ),
                25,
                ('kat_auswahl',"ListBox",0,          
                                        'tab1',y ,100,20, 
                                        ('Border','Dropdown','LineCount'),
                                        ( 2,True,10),                         
                                        {'addItems':(tags_items,0),'SelectedItems':(state['kat_auswahl'],)} 
                                        ),
                0,
                ('neue_kat',"Button",1,       
                                        'tab2',y,20,25,
                                        ('Label',),
                                        (LANG.NEUE_KATEGORIE,),     
                                        {'setActionCommand':'neue_kat','addActionListener':listener}
                                        ),
                30,
                ('controlFL4',"FixedLine",0,      
                                        'tab1x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                20,
                ('mark_funde',"CheckBox",1,     
                                        'tab0x',y,110,22,
                                        ('Label','State'),
                                        (LANG.MARK_FUNDE, state['mark_funde']),           
                                        {} 
                                        ),
                0,
                ('farbe',"FixedText",0,        
                                        'tab2',0,40,20,  
                                        ('BackgroundColor','Label','Border'),
                                        (state['mark_farbe'],'    ',1),       
                                        {'addMouseListener':listener} 
                                        ),
                30,
                ('controlFL5',"FixedLine",0,      
                                        'tab0x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                20,
                ('suche',"Button",1,       
                                        'tab2+40-max',y,110,30,
                                        ('Label',),
                                        (LANG.SUCHE,),     
                                        {'setActionCommand':'suchen','addActionListener':listener}
                                        ),
                20,
                )
            
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 0),
                     1 : (None, 20),
                     2 : (None, 0),
                     }
            
            abstand_links = 10
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite
        except:
            log(inspect.stack,tb())

            
    def dialog_suche(self):
        if self.mb.debug:log(inspect.stack)
        
        try:
            
            listener = Suche_Dialog_Listener(self.mb)
            
            controls,max_breite = self.dialog_suche_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)             
             
            
            # Hauptfenster erzeugen
            posSize = None,None,max_breite,max_hoehe
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            
            
            # Controls in Hauptfenster eintragen
            for name,c in sorted(ctrls.items()):
                fenster_cont.addControl(name,c)
            
            listener.ctrls = ctrls
            listener.state = self.state
             
        except:
            log(inspect.stack,tb())            
            

from com.sun.star.awt import XActionListener, XMouseListener    
class Suche_Dialog_Listener(unohelper.Base, XActionListener, XMouseListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = None
        self.state = None
        self.ordinale = []

    def mousePressed(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        if ev.Source == self.ctrls['farbe']:
            farbe = self.mb.class_Funktionen.waehle_farbe()
            self.state['mark_farbe'] = farbe
            ev.Source.Model.BackgroundColor = farbe
        elif ev.Source == self.ctrls['info']:
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
                ok = self.dialog_suchen_auswerten()
                if ok:
                    self.suchen()
            elif cmd == 'neue_kat':
                name = self.ctrls['kat_neu_name'].Model.Text.strip()
                if name in self.mb.tags['name_nr']:
                    Popup(self.mb, 'warning').text = LANG.KATEGORIE_EXISTIERT.format(name)
                    return
                self.neue_tag_kat_anlegen(name)
                self.ctrls['kat_auswahl'].addItem(name,0)
                self.ctrls['kat_auswahl'].selectItem(name,True)
                
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
                                                                    parent=parent)
      
        except:
            log(inspect.stack,tb())
            

    def dialog_suchen_auswerten(self):
        if self.mb.debug: log(inspect.stack)

        try:
            
            ctrls = self.ctrls
            state = self.state
            
            ausgewaehlte = [ordi for ordi,v in self.mb.settings_exp['ausgewaehlte'].items() if v]
            

            state['rb1_alles'] = int(ctrls['rb1_alles'].State)
            state['rb1_sichtbare'] = int(ctrls['rb1_sichtbare'].State)
            state['rb1_auswahl'] = int(ctrls['rb1_auswahl'].State)
                   
            state['auswahl'] = ausgewaehlte
                        
            state['regex'] = ctrls['regex'].State
            state['klein_gross'] = ctrls['klein_gross'].State
            state['ganzes_wort'] = ctrls['ganzes_wort'].State
            
            state['neuer_tab'] = ctrls['neuer_tab'].State
            state['tab_name'] = ctrls['tab_name'].Model.Text
            
            state['in_vh_tags_eintragen'] = ctrls['in_vh_tags_eintragen'].State
            state['funde_taggen'] = ctrls['funde_taggen'].State
            state['tag_name'] = ctrls['tag_name'].Model.Text.strip()
            state['kat_auswahl'] = ctrls['kat_auswahl'].SelectedItemPos
            
            state['mark_funde'] = ctrls['mark_funde'].State
            
            alle_tags = [a for b,v in self.mb.tags['sammlung'].items() for a in v]
            
            if state['rb1_auswahl'] and ausgewaehlte == []:
                Popup(self.mb, 'info').text = LANG.KEINE_DATEIEN_AUSGEWAEHLT
                return False
            
            self.ordinale = self.get_ordinale_ausgewaehlte(ausgewaehlte)
            
            if not self.ordinale:
                Popup(self.mb, 'info').text = LANG.KEINE_DATEIEN_AUSGEWAEHLT
                return False
            
            elif not (state['neuer_tab'] or state['funde_taggen'] or state['mark_funde']):
                Popup(self.mb, 'info').text = LANG.KEINE_AKTION_GEWAEHLT
                return False
                
            elif state['regex'] and state['funde_taggen'] and state['tag_name'] == '':
                Popup(self.mb, 'info').text = LANG.REGEX_BRAUCHT_TAGNAMEN
                return False
            
            elif state['alle'] and state['funde_taggen'] and state['tag_name'] == '':
                Popup(self.mb, 'info').text = LANG.OPT_ALLE_BRAUCHT_TAGNAMEN
                return False
            
            elif ctrls['suchbegriffe'].Model.Text == '':
                Popup(self.mb, 'info').text = LANG.KEINE_SUCHBEGRIFFE.format(state['tag_name'])
                return False
            
            elif state['neuer_tab'] and state['tab_name'] in self.mb.props:
                Popup(self.mb, 'info').text = LANG.TAB_EXISTIERT_SCHON
                return False
            
            elif state['neuer_tab'] and state['tab_name'] == '':
                Popup(self.mb, 'info').text = LANG.TABNAMEN_EINGEBEN
                return False
            

            
            alle_tags = [a for b,v in self.mb.tags['sammlung'].items() for a in v] 
            state['suchbegriffe'] = self.get_suchbegriffe()
            
            if state['funde_taggen']:
                if not state['in_vh_tags_eintragen']: 
                    
                    doppelter_tag = None
                    
                    if state['tag_name'] in alle_tags and state['tag_name'] != '':
                        doppelter_tag = state['tag_name']
                    else:
                        if not state['regex'] and state['tag_name'] == '':
                            for s in state['suchbegriffe']:
                                if s in alle_tags:
                                    doppelter_tag = s
                                    break

                    if doppelter_tag:
                        Popup(self.mb, 'warning').text = LANG.TAG_VORHANDEN_WAEHLEN.format(doppelter_tag)
                        return False
                
            return True
        except:
            log(inspect.stack,tb())
        
        
        
    def get_suchbegriffe(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            term = self.ctrls['suchbegriffe'].Model.Text
            
            if self.state['regex']:
                return term
            
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
            
    
    def get_ordinale_ausgewaehlte(self, ausgewaehlte):
        if self.mb.debug: log(inspect.stack)
        
        state = self.state
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
    
    
    def suchen(self):
        if self.mb.debug: log(inspect.stack)
        
        state = self.state
        
        plain_text = self.plain_txt_auslesen(self.ordinale)
        
        if state['regex']:
            ergebnisse = self.suchen_regex(plain_text,state['suchbegriffe'])
            alle_funde = list(ergebnisse)
            gemeinsam = alle_funde
        else:
            ergebnisse, gemeinsam = self.suchen_python(plain_text,state['suchbegriffe'])
            alle_funde = list( set( [x for a in ergebnisse.values() for x in a] ) )
            
                
        if state['regex']:
            ordis = alle_funde
        
        elif state['alle']:
            ordis = gemeinsam
        else:
            ordis = alle_funde
        
        if not ordis:
            Popup(self.mb, 'info').text = LANG.SUCHE_KEINE_ERGEBNISSE
            return
        
        if state['neuer_tab']:
            ord_sortiert = self.mb.class_Tabs.sortiere_ordinale(ordis)
            if self.mb.props[T.AB].Projektordner in ord_sortiert:
                ord_sortiert.remove(self.mb.props[T.AB].Projektordner)
            self.mb.tabsX.erzeuge_neuen_tab2(ord_sortiert, tab_name=state['tab_name'])
        
        if state['mark_funde']:
            self.fundstellen_markieren(ordis,ergebnisse)
        
        if state['funde_taggen']:
            self.funde_taggen(ordis,ergebnisse)
        
    
    def funde_taggen(self,ordis,ergebnisse):
        if self.mb.debug: log(inspect.stack)
        
        try:
            tags = self.mb.tags
            state = self.state
            
            alle_tags = [a for b,v in self.mb.tags['sammlung'].items() for a in v]
            
            
            # nur ein neuer Tag wird eingetragen
            if state['tag_name']:
                
                new_tag = state['tag_name']
                
                if state['regex']:
                    ergebnisse = {'xxx':list(ergebnisse)}
                
                if new_tag in alle_tags:
                    panel_nr = [nr for nr,v in tags['sammlung'].items() if new_tag in v][0]
                else:
                    kat_name = self.ctrls['kat_auswahl'].SelectedItem
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
                if state['in_vh_tags_eintragen']:
                    
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
                
                
                kategorie = self.ctrls['kat_auswahl'].SelectedItem
                panel_nr = tags['name_nr'][kategorie]
                
                for new_tag, ordis in ergebnisse.items():
                    for o in ordis:
                        if new_tag not in tags['ordinale'][o][panel_nr]:
                            tags['ordinale'][o][panel_nr].append(new_tag)
                        if new_tag not in tags['sammlung'][panel_nr]:
                            tags['sammlung'][panel_nr].append(new_tag)
            
            
            self.mb.class_Sidebar.erzeuge_sb_layout()

        except:
            log(inspect.stack,tb())
        
        
    def suchen_regex(self,texte,begriff):
        if self.mb.debug: log(inspect.stack)
        # re.escape
        try:
            results = {}
            comp = re.compile(begriff)

            for ordinal,txt in texte.items():
                if comp.search(txt):
                    results.update({ordinal:None})
            return results  
        except:
            log(inspect.stack,tb())
            return []
        
        
    def suchen_python(self,otexte,obegriffe):
        if self.mb.debug: log(inspect.stack)
        
        state = self.state
        
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
        
        if state['alle']:
            
            alle_funde = [x for x in ergebnisse.values()]
            gemeinsam = set(alle_funde[0])
            for b in alle_funde[1:]:
                gemeinsam = gemeinsam.intersection(b)
            
            
        
        return ergebnisse, list(gemeinsam)
    
    
        
    def fundstellen_markieren(self,ordinale,ergebnisse): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            state = self.state

            props = self.mb.props[T.AB]
            sichtbare = [props.dict_bereiche['Bereichsname-ordinal'][s] for s in self.mb.sichtbare_bereiche]
            
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
                                funde.getByIndex(e).CharBackColor = state['mark_farbe']
                
                else:
                    sd.SearchRegularExpression = True
                    sd.SearchString = state['suchbegriffe']
                    
                    funde = doc.findAll(sd)
                            
                    for e in range(funde.Count):
                        funde.getByIndex(e).CharBackColor = state['mark_farbe']
                
                
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
         
        self.mb.class_Sidebar.erzeuge_sb_layout()
        
    
    def oeffne_suche_info(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            lang,hb,ordi = ['de','Organon Handbuch.organon',
                            'nr74.odt'] if self.mb.language == 'de' else ['en','Organon Manual.organon','nr81.odt']
            
            url = os.path.join(self.mb.path_to_extension,'description','Handbuecher',lang,hb,'Files','odt',ordi)
            URL = uno.systemPathToFileUrl(url)
            self.mb.class_Fenster.oeffne_dokument_in_neuem_fenster(URL)
        
        except:
            log(inspect.stack,tb())

        
        
        
        
            