# -*- coding: utf-8 -*-

import unohelper


class Export():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.haupt_fenster = None
        self.trenner_fenster = None
        self.auswahl_fenster = None
        
        
    def export(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            if self.mb.filters_export == None:
                self.mb.class_Import.erzeuge_filter()
                
            self.dialog_exportfenster()
                
        except Exception as e:
            self.mb.nachricht('Export.export '+ str(e),"warningbox")
            log(inspect.stack,tb())


    def dialog_exportfenster_elemente(self): 
        if self.mb.debug: log(inspect.stack)
        try:
            breite = 400
            
            
            sett = self.mb.settings_exp
            
            
            # benoetigte Listner erzeugen            
            listenerDis = AB_Fenster_Dispose_Listener(self.mb,self)
            listener = Speicherordner_Button_Listener(self.mb)
            listener6 = Fenster_Export_Listener1(self.mb)
            listener3 = B_Auswahl_Button_Listener(self.mb,self)
            listener4 = Export_Button_Listener(self.mb)        
            listener5 = A_Trenner_Button_Listener(self.mb,self)        
            listener2 = Fenster_Export_Listener2(self.mb)
            f_listener = ExportFilter_Item_Listener(self.mb)
            
            listener_dic = {
                            'listenerDis' : listenerDis,
                            'listener' : listener,
                            'listener6' : listener6,
                            'listener3' : listener3,
                            'listener4' : listener4,
                            'listener5' : listener5,
                            'listener2' : listener2,
                            'f_listener' : f_listener
                            }
            
            
            
            # FILTER
            filters = self.mb.filters_export
            filter_items = tuple(filters[x][0] for x in filters)
            sel = tuple(list(filters).index(f) for f in filters if f == sett['typ'])[0]
            

            # SPEICHERORT
            label_speicherort = decode_utf(sett['speicherort'])
            try:
                l = uno.fileUrlToSystemPath(label_speicherort)
            except:
                l = label_speicherort
            label_speicherort = l
            
            abstand = 18
            y = 0
            
            controls = (
                10,
                ('controlE',"FixedText",1,                      
                                        'tab0',y ,50,20,        
                                        ('Label','FontWeight'), 
                                        (LANG.EXPORT ,150),                 
                                        {}                      
                                        ), 
                30,
                ('control_CB1',"CheckBox",1,     
                                        'tab0',y,120,22, 
                                        ('Label','State','VerticalAlign'),
                                        (LANG.ALLES ,sett['alles'],1),              
                                        {'addItemListener':(listener6)} 
                                        ),    
                20,                                                  
                ('control_CB2',"CheckBox",1,    
                                        'tab0',y,120,22,  
                                        ('Label','State','VerticalAlign'),
                                        (LANG.SICHTBARE ,sett['sichtbar'],1),       
                                        {'addItemListener':(listener6)} 
                                        ),  
                20,
                ('control_CB3',"CheckBox",1,     
                                        'tab0',y,120,22,
                                        ('Label','State','VerticalAlign'),
                                        (LANG.AUSWAHL ,sett['eigene_ausw'],1),      
                                        {'addItemListener':(listener6)} 
                                        ),  
                0,                                                      
                ('controlA',"Button",1,           
                                        'tab1-tab1-E',y,70,20,
                                        ('Label',),
                                        (LANG.AUSWAHL,),                                 
                                        {'Enable':not(sett['alles'] or sett['sichtbar']),'addActionListener':(listener3,)} 
                                        ),
                30,
                ('controlT',"FixedLine",0,      
                                        'tab0x-max',y ,40,5,
                                        (),
                                        (),                                                
                                        {} 
                                        ),  
                ######################################################################################################
                40,
                ('controlex',"FixedText",1,       
                                        'tab0',y ,150,20,
                                        ('Label','FontWeight'),
                                        (LANG.EXPORTIEREN_ALS,150.0),        
                                        {} 
                                        ),
                30,  
                ('controlcb1',"CheckBox",1,      
                                        'tab0',y,110,22, 
                                        ('Label','State'),
                                        (LANG.EIN_DOKUMENT,sett['einz_dok']),     
                                        {'addItemListener':(listener2)} 
                                        ),  
                20,
                ('controlcb2',"CheckBox",1,     
                                        'tab0+{}'.format(abstand),y,110,22,
                                        ('Label','State'),
                                        (LANG.TRENNER,sett['trenner']),           
                                        {'addItemListener':(listener2),'Enable':not(sett['einz_dat'] or sett['neues_proj']) } 
                                        ),
                0,
                ('controlTr',"Button",1,          
                                        'tab1-tab1-E',y ,70,20, 
                                        ('Label',),
                                        (LANG.BEARBEITEN,),                        
                                        {'Enable':(sett['trenner'] and sett['einz_dok']),'addActionListener':(listener5,) } 
                                        ),
                40,
                ('controlcb3',"CheckBox",1,       
                                        'tab0',y,110,22,
                                        ('Label','State'),
                                        (LANG.EINZ_DATEIEN,sett['einz_dat']),     
                                        {'addItemListener':(listener2)} 
                                        ),
                20,
                ('controlcb4',"CheckBox",1,      
                                        'tab0x+{}'.format(abstand),y,110,22,
                                        ('Label','State'),
                                        (LANG.ORDNERSTRUKTUR,sett['ordner_strukt']),     
                                        {'addItemListener':(listener2),'Enable':not(sett['einz_dok'] or sett['neues_proj']) } 
                                        ),
                30,
                ('controlcb5',"CheckBox",1,       
                                        'tab0',y,110,22, 
                                        ('Label','State'),
                                        (LANG.NEW_PROJECT2,sett['neues_proj']),   
                                        {'addItemListener':(listener2)} 
                                        ),
                25,
                ('controlPN',"Edit",0,           
                                        'tab0+{}'.format(abstand),y ,100,20, 
                                        ('HelpText',),
                                        (LANG.PROJEKT_NAMEN_EINGEBEN,),               
                                        {'Enable':not(sett['einz_dat'] or sett['einz_dok']) } 
                                        ),
                30,
                ('controlFL2',"FixedLine",0,        
                                        'tab0x-max',y,40,5,
                                        (),
                                        (),                                                    
                                        {} 
                                        ),
                #######################################################################################################################
                40,
                ('controlf',"FixedText",1,        
                                        'tab0',y ,80,20,
                                        ('Label','FontWeight'),
                                        (LANG.DATEITYP,150),                 
                                        {} 
                                        ),
                0,
                ('controlL',"ListBox",0,       
                                        'tab1x-max',y ,150,20,
                                        ('LineCount','Dropdown'),
                                        (150,True), 
                                        {'addItems':(filter_items,0), 'SelectedItems':(sel,) ,'addItemListener':(f_listener)}
                                        ),
                30,
                ('controlSO',"FixedText",1,      
                                        'tab0',y ,80,20, 
                                        ('Label','FontWeight'),
                                        (LANG.SPEICHERORT,150),              
                                        {} 
                                        ),
                0,
                ('controlSD',"Button",1,          
                                        'tab1-tab1-E',y,70,20,
                                        ('Label',),
                                        (LANG.WAEHLEN,),                              
                                        {'setActionCommand':'speicherort','addActionListener':(listener,)} 
                                        ) ,
                30,
                ('controlFO',"FixedText",0,      
                                        'tab0-max',y ,80,60, 
                                        ('HelpText','Label','MultiLine'),
                                        ('URL',label_speicherort,True),             
                                        {} 
                                        ),
                60,
                ('controlB',"Button",1,         
                                        'tab2' ,y,80,30,  
                                        ('Label',),
                                        (LANG.EXPORTIEREN,),                 
                                        {'setActionCommand':'speicherort','addActionListener':(listener4,)} 
                                        ) ,
                20
                )
            
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 15),
                     1 : (None, 1),
                     2 : (None, 5),
                     3 : (None, 5),
                     4 : (None, 5),
                     }
            
            abstand_links = 10
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite,listener_dic
        except:
            log(inspect.stack,tb())
            
            
    def dialog_exportfenster(self):
        if self.mb.debug:log(inspect.stack)
        
        try:
            controls,max_breite,listener_dic = self.dialog_exportfenster_elemente()
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)             
                          
        
            
            # Custom Properties eintragen
            buttons = [ctrls['control_CB1'],
                       ctrls['control_CB2'],
                       ctrls['control_CB3']]
            listener_dic['listener6'].buttons = buttons
            
            buttons2 = [ctrls['controlcb1'],
                        ctrls['controlcb2'],
                        ctrls['controlcb3'],
                        ctrls['controlcb4'],
                        ctrls['controlcb5'],
                        ctrls['controlPN']]
            
            listener_dic['listener2'].buttons2 = buttons2
            
            listener_dic['listener2'].trenner = ctrls['controlTr']
            listener_dic['listener6'].but_Auswahl = ctrls['controlA']
            listener_dic['listener'].model = ctrls['controlFO'].Model
            listener_dic['listener4'].feld_projekt_name = ctrls['controlPN']
            
            

            # Hauptfenster erzeugen
            posSize = None,None,max_breite,max_hoehe
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            fenster_cont.Model.Text = LANG.EXPORT
            
            fenster_cont.addEventListener(listener_dic['listenerDis'])
            self.haupt_fenster = fenster
            
            # Listener fuer Hauptfenster
            listener_dic['listener3'].exp_fenster = fenster
            listener_dic['listener5'].exp_fenster = fenster
            
            # Controls in Hauptfenster eintragen
            for name,c in ctrls.items():
                fenster_cont.addControl(name,c)
                
        except:
            log(inspect.stack,tb())
            
     
    def kopiere_projekt(self,neuer_projekt_name,pfad_zu_neuem_ordner,
                        ordinale,tree,tags_neu,backup = False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            # zunaechst den zuletzt bearbeiteten Bereich speichern
            if self.mb.props[T.AB].tastatureingabe:
                ordinal = self.mb.props[T.AB].selektierte_zeile
                bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
                # nachfolgende Zeile erzeugt bei neuem Tab Fehler - 
                path = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)
            
            
            from shutil import copy 
            
            self.neuer_pfad = ''

            ordinale.append('empty_file')
            ordinale_files = list(ordi + '.odt' for ordi in ordinale)
            
            
            # neuen Projektordner erstellen
            if not os.path.exists(pfad_zu_neuem_ordner):
                os.makedirs(pfad_zu_neuem_ordner)
            
            def get_path(pfad):
                if os.path.split(pfad)[1] == self.base:  
                    pf = ''
                    for p in self.neuer_pfad_list:
                        pf = os.path.join(pf,p)
                    self.neuer_pfad = pf        
                else:
                    self.neuer_pfad_list.insert(0,os.path.split(pfad)[1])
                    split_pfad = os.path.split(pfad)[0]
                    get_path(split_pfad)
                    

            Path = self.mb.pfade['projekt']
            alter_projekt_name = self.mb.projekt_name
            
            first_time = True
            self.neuer_pfad_list = []
            
            # Den Pfad durchlaufen, bevor geschrieben wird, da beim Schreiben in denselben Pfad
            # eine Endlosschleife entsteht
            path_walk = []
            for a,b,c in os.walk(Path):
                if 'Backups' in a:
                    continue
                path_walk.append((a,b,c))  
            
            
            for root,dirs,files in path_walk:
                
                if not backup:
                    if os.path.basename(root) == 'Tabs':
                        continue
                
                # Variable fuer den relativen Pfad in get_path zuruecksetzen   
                self.neuer_pfad_list = []
                
                if first_time:
                    self.base = os.path.basename(root)
                    first_time = False
                else:
                    get_path(root)


                if self.neuer_pfad != '':
                    ziel = os.path.join(pfad_zu_neuem_ordner,self.neuer_pfad)
                else:
                    ziel = pfad_zu_neuem_ordner
                
                
                for d in dirs:
                    z_pfad = os.path.join(ziel,d)
                    os.makedirs(z_pfad)
                    
                for f in files:
                    q_pfad = os.path.join(root,f)
                    z_pfad = os.path.join(ziel,f)
                    
                    try:
                        if '~lock' in f: continue
                        
                        ### AENDERUNGEN GEGENUEBER DEM ORIGINAL AUSFUEHREN
                        elif 'ElementTree' in f: 
                            if backup:
                                root_xml = tree.getroot()
                                projekt_xml = root_xml.find(".//*[@Art='prj']")
                                projekt_xml.attrib['Name'] = neuer_projekt_name
                            self.mb.tree_write(tree,z_pfad)
                            continue
                        
                        elif f == 'tags.pkl': 
                            from pickle import dump as pickle_dump
                            with open(z_pfad, 'wb') as fi:
                                pickle_dump(tags_neu, fi,2)
                        
                        # ungenutzte odts vom Kopieren aussschliessen  
                        elif 'odt' == os.path.basename(root):
                            if f not in ordinale_files:
                                continue
                        
                        # ungenutzte Bilder vom Kopieren aussschliessen  
                        elif 'Images' == os.path.basename(root):
                            bild_vorhanden = False
                            
                            tags = self.mb.tags
                            bild_panels = [i for i,v in tags['nr_name'].items() if v[1] == 'img']
                            
                            for ordi in tags_neu['ordinale']:
                                for b in bild_panels:
                                    if tags_neu['ordinale'][ordi][b] != '':
                                        b_pfad = uno.fileUrlToSystemPath(tags_neu['ordinale'][ordi][b])
                                    else:
                                        b_pfad = ''
                                    if os.path.basename(b_pfad) == os.path.basename(z_pfad):
                                        bild_vorhanden = True
                            if not bild_vorhanden:
                                continue
                    except:
                        # Bei Fehlern wird die Datei einfach kopiert
                        pass

                    copy(q_pfad, z_pfad)
                    
                    # Startdatei .organon umbenennen
                    if f == alter_projekt_name + '.organon':
                        neuer_name = os.path.join(os.path.split(z_pfad)[0],neuer_projekt_name + '.organon')
                        os.rename(z_pfad,neuer_name)
                      
        except Exception as e:
            self.mb.nachricht('kopiere_Projekt ' + str(e),"warningbox")
            log(inspect.stack,tb())

        

def decode_utf(term):
    
    if isinstance(term, str):
        return term
    else:
        return term.decode('utf8')   



from com.sun.star.awt import XItemListener, XActionListener, XFocusListener    
class ExportFilter_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
              
        sel = ev.value.Source.Items[ev.value.Selected] 
        filters = self.mb.filters_export

        for f in filters:
            if filters[f][0] == sel:

                self.mb.settings_exp['typ'] = f
                self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp)
            
    def disposing(self,ev):
        return False


class Fenster_Export_Listener1(unohelper.Base, XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.buttons = None
        self.but_Auswahl = None
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
           
        sett = self.mb.settings_exp   
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
        # alle anderen CheckBoxen auf 0 setzen
        for cont in self.buttons:
            if cont != ev.Source:
                cont.State = 0
                
        if ev.Source.Model.Label == LANG.AUSWAHL:
            self.but_Auswahl.Enable = True
        else:
            self.but_Auswahl.Enable = False
        
        # settings_exp neu setzen
        sett['alles'] = 0
        sett['sichtbar'] = 0
        sett['eigene_ausw'] = 0
        
        if ev.Source.Model.Label == LANG.ALLES:
            sett['alles'] = 1
        elif ev.Source.Model.Label == LANG.SICHTBARE:
            sett['sichtbar'] = 1
        elif ev.Source.Model.Label == LANG.AUSWAHL:
            sett['eigene_ausw'] = 1
            
    def disposing(self,ev):
        return False

              
class Fenster_Export_Listener2(unohelper.Base, XItemListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.buttons2 = None
        self.trenner = None
        

    def itemStateChanged(self, ev):    
        if self.mb.debug: log(inspect.stack)    

        if ev.Source.Model.Label in (LANG.EIN_DOKUMENT,LANG.EINZ_DATEIEN,LANG.NEW_PROJECT2):
            # um sich nicht selbst abzuwaehlen
            if ev.Source.State == 0:
                ev.Source.State = 1

            if ev.Source.Model.Label == LANG.EIN_DOKUMENT:
                self.buttons2[1].Enable = True
                self.buttons2[2].State = False
                self.buttons2[3].Enable = False
                self.buttons2[4].State = False
                self.buttons2[5].Enable = False
                if self.mb.settings_exp['trenner']:
                    self.trenner.Enable = True
            elif ev.Source.Model.Label == LANG.EINZ_DATEIEN:
                self.buttons2[0].State = False
                self.buttons2[1].Enable = False
                self.buttons2[3].Enable = True
                self.buttons2[4].State = False
                self.buttons2[5].Enable = False
                self.trenner.Enable = False
            else:
                self.buttons2[0].State = False
                self.buttons2[2].State = False
                self.buttons2[1].Enable = False
                self.buttons2[3].Enable = False
                self.buttons2[5].Enable = True
                self.trenner.Enable = False
            
            self.mb.settings_exp['einz_dat'] = 0
            self.mb.settings_exp['einz_dok'] = 0
            self.mb.settings_exp['neues_proj'] = 0
            
            if ev.Source.Model.Label == LANG.EIN_DOKUMENT:
                self.mb.settings_exp['einz_dok'] = 1
            elif ev.Source.Model.Label == LANG.EINZ_DATEIEN:
                self.mb.settings_exp['einz_dat'] = 1   
            elif ev.Source.Model.Label == LANG.NEW_PROJECT2:
                self.mb.settings_exp['neues_proj'] = 1   
                
        elif ev.Source.Model.Label == LANG.ORDNERSTRUKTUR:
            if self.mb.settings_exp['ordner_strukt']:
                self.mb.settings_exp['ordner_strukt'] = 0
            else:
                self.mb.settings_exp['ordner_strukt'] = 1
        
        elif ev.Source.Model.Label == LANG.TRENNER: 
            if self.mb.settings_exp['trenner']:
                self.mb.settings_exp['trenner'] = 0
                self.trenner.Enable = False
            else:
                self.mb.settings_exp['trenner'] = 1
                self.trenner.Enable = True
            
                
    def disposing(self,ev):
        return False



class Export_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.feld_projekt_name = None
        self.ungespeichert = []
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_exp
        
        sections = self.get_ausgewaehlte_bereiche()
        
        if len(sections) < 1:
            self.mb.nachricht(LANG.NICHTS_AUSGEWAEHLT,"infobox")
            return
        
        if sett['einz_dok']:
            ok = self.exp_in_ein_dokument(sections)
        elif sett['einz_dat']:
            ok = self.exp_in_einzel_dat(sections)
        elif sett['neues_proj']:
            ok = self.exp_in_neues_proj(sections)
        
        if ok:
            self.ueberpruefe_auf_ungespeicherte()    
            self.mb.nachricht(LANG.EXPORT_ERFOLGREICH,"infobox")
        else:
            self.mb.nachricht(LANG.EXPORT_NICHT_ERFOLGREICH,"warningbox")

    def ueberpruefe_auf_ungespeicherte(self):
        if self.mb.debug: log(inspect.stack)

        if self.ungespeichert != []:
            self.mb.nachricht(LANG.UNGESPEICHERTE_PFADE,"warningbox")
        
        self.ungespeichert = []
        
        
          
    def get_ausgewaehlte_bereiche(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_exp

        # selektiert nur die sichtbaren Bereiche
        if sett['sichtbar']:
            sections = []
            for sec_name in self.mb.props['ORGANON'].sichtbare_bereiche:
                sections.append(sec_name)
        
        else:
            # selektiert alle Bereiche
            secs = self.mb.doc.TextSections.ElementNames
            sections = []
            
            for sec_name in secs:
                if 'OrganonSec' in sec_name:
                    sections.append(sec_name)
        
        # selektiert aus allen Bereichen die ausgewaehlten
        if sett['eigene_ausw']:
            sects = [] 
            for sec_name in sections:
                if 'OrganonSec' in sec_name:  
                    # Abfrage, ob sec in Tab
                    if sec_name in self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal']:
                        sec_ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][sec_name]
                        if sec_ordinal in sett['ausgewaehlte']:
                            if sett['ausgewaehlte'][sec_ordinal] == 1:
                                sects.append(sec_name)

            sections = sects

        return sections
            
            
    def exp_in_ein_dokument(self,sections):
        if self.mb.debug: log(inspect.stack)
        
        st_ind = self.mb.current_Contr.Frame.createStatusIndicator()        
        
        try:
            sett = self.mb.settings_exp
            self.mb.class_Bereiche.starte_oOO()
            
            oOO = self.mb.class_Bereiche.oOO
            cur = oOO.Text.createTextCursor()
            text = oOO.Text
            cur.gotoEnd(False)
            
            
            anz_sections = len(sections)
            st_ind.start('exportiere, bitte warten',anz_sections)
            st_ind.setValue(anz_sections/2)
            zaehler = 1
            
            # Pruefen, ob die Trenndatei noch existiert
            trennerDat_existiert = True
            if sett['dok_einfuegen']:
                URL = sett['url']
                if URL != '':
                    if not os.path.exists(URL):
                        trennerDat_existiert = False
                    
                
            for sec_name in sections:         
                    
                zaehler += 1
                 
                cur.gotoEnd(False)
                
                sec_ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][sec_name]
                
                if sec_ordinal == self.mb.props[T.AB].Papierkorb:
                    break  
                
                
                if sett['trenner']:
                
                    if sett['seitenumbruch_ord']:
                        if sec_ordinal in self.mb.props[T.AB].dict_ordner:
                            from com.sun.star.style.BreakType import PAGE_BEFORE
                            cur.BreakType = PAGE_BEFORE
                            text.insertControlCharacter(cur, 0, True)
                            
                    if sett['seitenumbruch_dat']:
                        if sec_ordinal not in self.mb.props[T.AB].dict_ordner:
                            from com.sun.star.style.BreakType import PAGE_BEFORE
                            cur.BreakType = PAGE_BEFORE
                            text.insertControlCharacter(cur, 0, True)   
                    
                    if sett['ordnertitel']:
                        
                        if sec_ordinal in self.mb.props[T.AB].dict_ordner:
                            
                            contr = self.mb.props[T.AB].Hauptfeld.getControl(sec_ordinal)
                            tf = contr.getControl('textfeld')
                            titel = tf.Model.Text
                            
                            if sett['format_ord']:
                                oldStyle = cur.ParaStyleName
                                cur.ParaStyleName = self.mb.settings_exp['style_ord'] 
                                
                            cur.setString(titel)
                            cur.gotoEnd(False)
                            
                            if sett['format_ord']:
                                text.insertControlCharacter(cur,0,False)
                                cur.ParaStyleName = oldStyle
                                
                            cur.gotoEnd(False)
                            
                    
                    if sett['dateititel']:
                        if sec_ordinal not in self.mb.props[T.AB].dict_ordner:
                            
                            contr = self.mb.props[T.AB].Hauptfeld.getControl(sec_ordinal)
                            tf = contr.getControl('textfeld')
                            titel = tf.Model.Text
                            
                            if sett['format_dat']:
                                oldStyle = cur.ParaStyleName
                                cur.ParaStyleName = self.mb.settings_exp['style_dat'] 
                                
                            cur.setString(titel)
                            cur.gotoEnd(False)
                            
                            if sett['format_dat']:
                                text.insertControlCharacter(cur,0,False)
                                cur.ParaStyleName = oldStyle
                            
                            cur.gotoEnd(False)

                                
                cur.gotoEnd(False)
                cur.gotoEndOfParagraph(False)

                fl = self.mb.props[T.AB].dict_bereiche['Bereichsname'][sec_name]
                link = uno.systemPathToFileUrl(fl)
                
                SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                SFLink.FilterName = 'writer8'
                SFLink.FileURL = link
                                
                newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                newSection.setPropertyValue('FileLink',SFLink)
                
                oOO.Text.insertTextContent(cur,newSection,False)   
                oOO.Text.removeTextContent(newSection)                   
                
                if sett['trenner']:
                
                    if sett['leerzeilen_drunter']:
                        anz = int(sett['anz_drunter'])
                        for i in range(anz):
                            cur.gotoEnd(False)
                            text.insertControlCharacter(cur,0,False)
                            cur.gotoEnd(False)
                  
                    
                    if sett['dok_einfuegen'] and trennerDat_existiert:
                        cur.gotoEnd(False)
                        URL = sett['url']
                        SFLink2 = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                        SFLink2.FileURL = uno.systemPathToFileUrl(URL)

                        newSection2 = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                        newSection2.setPropertyValue('FileLink',SFLink2)
                        
                        oOO.Text.insertTextContent(cur,newSection2,False)
                        oOO.Text.removeTextContent(newSection2)
                        
                        cur.gotoEnd(False)
           
           
            # entferne OrgInnerSec
            tsecs = oOO.TextSections
            for tname in tsecs.ElementNames:
                if 'OrgInnerSec' in tname:
                    sec = tsecs.getByName(tname)
                    sec.dispose()
                    
            
            path = uno.fileUrlToSystemPath(decode_utf(self.mb.settings_exp['speicherort']))
            Path2 = os.path.join(path, self.mb.projekt_name)
            
            filters = self.mb.filters_export
    
            ofilter = [(f,filters[f]) for f in filters if f == sett['typ']]
            filterName = ofilter[0][0]
            extension = '.' + ofilter[0][1][1][0]
            
            
            if filterName == 'LaTex':
                extension = '.tex'
            elif filterName == 'HtmlO':
                extension = '.html'
            
            if os.path.exists(Path2+extension):
                Path2 = self.pruefe_dateiexistenz(Path2,extension)
            
            Path1 = Path2+extension
            Path3 = uno.systemPathToFileUrl(Path1)   
            
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'FilterName'
            prop.Value = filterName
            

            if filterName == 'LaTex':
                self.mb.class_Latex.greek2latex(oOO.Text,Path2,self.mb.projekt_name)
            elif filterName == 'HtmlO':
                self.mb.class_Html.export2html(oOO.Text,Path1,self.mb.projekt_name)
            else:
                oOO.storeToURL(Path3,(prop,))

            ok = True
        except:
            log(inspect.stack,tb())
            ok = False
            
        self.mb.class_Bereiche.schliesse_oOO()   
        st_ind.end() 
        
        return ok
        
    def exp_in_einzel_dat(self,sections):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sett = self.mb.settings_exp
            st_ind = self.mb.current_Contr.Frame.createStatusIndicator()    
    
            
            # pruefen, ob speicherordner existiert; Namen aendern
            speicherort = decode_utf(self.mb.settings_exp['speicherort'])
            speicherordner = os.path.join(uno.fileUrlToSystemPath(speicherort), self.mb.projekt_name)
            if os.path.exists(speicherordner):
                speicherordner = self.pruefe_dateiexistenz(speicherordner)
            
            
            def berechne_tree():
                 
                tree2 = copy.deepcopy(self.mb.props[T.AB].xml_tree)
                root2 = tree2.getroot()
                
                all_el = root2.findall('.//')
                
                for el in all_el:
                    parent = root2.find('.//'+el.tag+'/..')
                    doppelte = parent.findall("./*[@Name='%s']" %el.attrib['Name'])
                    
                    if len(doppelte) > 1:
                        for i in range(len(doppelte)-1):
                            elem = doppelte[i+1]
                            elem.attrib['Name'] = elem.attrib['Name'] + '(%s)' %i
                            
                return tree2
             
             
            # Pfade zum Speichern in Ordnerstruktur berechnen        
            if sett['ordner_strukt']:
    
                tree = berechne_tree()
                root = tree.getroot()
                
                baum = []
                self.mb.class_XML.get_tree_info(root,baum)
    
                pfade = {}
                dict_baum = {}
    
                for eintrag in baum:
                    ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag  
                    dict_baum.update({ordinal:(self.mb.class_Funktionen.verbotene_buchstaben_austauschen(parent),
                                               self.mb.class_Funktionen.verbotene_buchstaben_austauschen(name),
                                               lvl,art,zustand,sicht)})
                
                def suche_parent(ord_kind):
                    ord_parent = dict_baum[ord_kind][0]
                    return ord_parent
                
                for eintrag in baum:
                    pfad2 = speicherordner
                    ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
                    
                    name = self.mb.class_Funktionen.verbotene_buchstaben_austauschen(name)
                    
                    ordner = []
                    ordn = ordinal
                    for i in range(int(lvl)-1):
                        ordn = suche_parent(ordn)
                        ordner.append(dict_baum[ordn][1])
                    
                    ordner.reverse()  

                    for ordn in ordner:
                        pfad2 = os.path.join(pfad2, ordn)

                    pfad2 = os.path.join(pfad2, dict_baum[ordinal][1])

                    if art in ('dir'):
                        pfad2 = os.path.join(pfad2, name)
                    
                    pfade.update({ordinal:(name,pfad2,art)})
            
            
            # Statusindicator
            anz_sections = len(sections)
            st_ind.start(LANG.EXPORT_BITTE_WARTEN,anz_sections)
            zaehler = 0
            
            
            ungespeichert = []
            
            while zaehler < anz_sections:
                 
                self.mb.class_Bereiche.starte_oOO()
                oOO = self.mb.class_Bereiche.oOO
                cur = oOO.Text.createTextCursor()
                text = oOO.Text
                
                    
                # Speichern     
                for i in range(3):
                    
                    if zaehler  > len(sections) - 1:  
                        break
                     
                    sec_name = sections[zaehler]
                     
                    if 'OrganonSec' in sec_name:  
                         
                        zaehler += 1       
                         
                        sec_ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal'][sec_name]
                        if sec_ordinal == self.mb.props[T.AB].Papierkorb:
                            break  
                        
                        # evt vorhandene Sections loeschen, die mit dem Cursor nicht erreicht werden
                        for el_n in range(oOO.TextSections.Count):
                            el = oOO.TextSections.getByIndex(0)
                            el.dispose()
    
                        cur.gotoStart(False)
                        cur.gotoEnd(True)
                        cur.setString('')
                        
                        
                        fl = self.mb.props[T.AB].dict_bereiche['Bereichsname'][sec_name]
                        link = uno.systemPathToFileUrl(fl)
                        
                        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                        SFLink.FilterName = 'writer8'
                        SFLink.FileURL = link
                                        
                        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                        newSection.setPropertyValue('FileLink',SFLink)
                        
                      
                        oOO.Text.insertTextContent(cur,newSection,False)                   
                        
                        # Den durch Organon angelegten Textbereich wieder loeschen
                        try:
                            cur.gotoRange(newSection.Anchor,False)  
                        except:
                            # aendern
                            # Tabellen erzeugen hier einen Fehler
                            # allerdings ist nicht klar, ob diese Zeile ueberhaupt gebraucht
                            # wird. Evtl. reicht removeTextContent auch aus.
                            # pruefen !!!
                            pass
                            
                        
                        # entferne OrgInnerSec
                        tsecs = oOO.TextSections
                        for tname in tsecs.ElementNames:
                            if 'OrgInnerSec' in tname:
                                sec = tsecs.getByName(tname)
                                sec.dispose()

                        cur.gotoEnd(False)
                        cur.goLeft(1,True)
                        cur.setString('')
                            
                        oOO.Text.removeTextContent(newSection)
                         
                        if not sett['ordner_strukt']:
                            contr = self.mb.props[T.AB].Hauptfeld.getControl(sec_ordinal)
                            tf = contr.getControl('textfeld')
                            titel = tf.Model.Text
                            titel = "".join(i for i in titel if i not in r'\/:*?"<>|').strip()
                            pfad = os.path.join(speicherordner, titel)
                        else:
                            pfad = pfade[sec_ordinal][1]
                            
                        # Prop fuer Filter erstellen    
                        filters = self.mb.filters_export
                        ofilter = [(f,filters[f]) for f in filters if f == sett['typ']]
                        filterName = ofilter[0][0]
                        extension = '.' + ofilter[0][1][1][0] 
                        
                        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
                        prop.Name = 'FilterName'
                        prop.Value = filterName
                        
                        if filterName == 'LaTex':
                            extension = '.tex'
                        elif filterName == 'HtmlO':
                            extension = '.html'
                        
                        
                        # pruefen, ob datei existiert; Namen aendern       
                        if os.path.exists(pfad+extension):
                            pfad = self.pruefe_dateiexistenz(pfad,extension)
                            
                        path = pfad + extension 
                        
                        
                        # Auf eventuelle Fehler untersuchen.
                        # Die in Windows und Linux verbotenen Buchstaben in Pfaden sollten
                        # bereits getauscht worden sein. Falls doch einer fehlt, wird die Datei
                        # nicht gespeichert und ein Fehler in die Log Datei geschrieben
                        try:
                            path2 = uno.systemPathToFileUrl(path) 
                        except Exception as e:
                            self.ungespeichert.append(path)
                            continue

                        if filterName == 'LaTex':
                            self.mb.class_Latex.greek2latex(oOO.Text,pfad,self.mb.projekt_name)
                        elif filterName == 'HtmlO':
                            self.mb.class_Html.export2html(oOO.Text,path,self.mb.projekt_name)
                        else:
                            oOO.storeToURL(path2,(prop,))
                        
                        
                    # unterbricht while Schleife, wenn nur Trenner und keine keine OrganonSec mehr uebrig sind 
                    if self.teste_auf_verbliebene_bereiche(sections[zaehler::]):
                        zaehler = anz_sections
                         
                self.mb.class_Bereiche.schliesse_oOO()    
                st_ind.setValue(zaehler)
                
            ok = True
        except:
            log(inspect.stack,tb())
            ok = False

        st_ind.end()  
        return ok

    def exp_in_neues_proj(self,sections):
        if self.mb.debug: log(inspect.stack)
        
        # ToDo: Sicherstellen, dass die Rechte zum Schreiben existieren

        try:
            neuer_projekt_name = self.feld_projekt_name.Text
            ok = self.pruefe_neuen_projekt_namen()
            
            if not ok:
                return
            
            tree,ordinale = self.et_und_ordinale_berechnen(sections,neuer_projekt_name)
            tags_neu = self.passe_tags_an(ordinale)

            speicherort = uno.fileUrlToSystemPath(self.mb.settings_exp['speicherort'])
            pfad_zu_neuem_ordner = os.path.join(speicherort,neuer_projekt_name + '.organon')
            
            self.mb.class_Export.kopiere_projekt(neuer_projekt_name,pfad_zu_neuem_ordner,ordinale,tree,tags_neu)
            
            return True       
        except Exception as e:
            self.mb.nachricht('exp_in_neues_proj ' + str(e),"warningbox")
            log(inspect.stack,tb())
            return False
    
    def et_und_ordinale_berechnen(self,sections,projektname):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props['ORGANON']
        dict_BO = props.dict_bereiche['Bereichsname-ordinal']
        
        # Ordinale der Sections bestimmen
        ordinale = []
        for bereich in dict_BO:
            if bereich in sections:
                ordinale.append(dict_BO[bereich])
    
        
        # Deepcopy des ElementTrees zum Bearbeiten oeffnen
        tree = copy.deepcopy(self.mb.props['ORGANON'].xml_tree)
        root = tree.getroot()
        
        all_el = root.findall('.//')            
        
        # alle Seiten, die nicht mehr im ET vorkommen, loeschen
        for el in all_el:
            if el.tag not in ordinale:
                if el.attrib['Art'] == 'pg':
                    parent = root.find('.//'+el.tag+'/..')
                    child = root.find('.//'+el.tag)
                    parent.remove(child)

        alle_ordner = root.findall(".//*[@Art='dir']")
        
        # hoechsten Lvl berechnen
        lvl = 0
        for ordner in alle_ordner:
            if int(ordner.attrib['Lvl']) > lvl:
                lvl = int(ordner.attrib['Lvl'])

        # alle Ordner, die nicht mehr im ET vorkommen und kein Kind mehr haben, loeschen
        # Schleife nach lvl von hoch nach niedrig durchlaufen
        for l in reversed(range(int(lvl)+1)):
            ordner_lvl = root.findall(".//*[@Art='dir'][@Lvl='%s']" %l)
            for o_lvl in ordner_lvl:
                childs = list(c for c in o_lvl)
                if len(childs) == 0:
                    parent = root.find('.//'+o_lvl.tag+'/..')
                    child = root.find('.//'+o_lvl.tag)
                    parent.remove(child)
        
        # Projektnamen neu eintragen
        prj = root.find(".//*[@Art='prj']")
        prj.attrib['Name'] = projektname
        
        ordinale = []
        all_el = root.findall('.//')  
        
        for el in all_el:
            ordinale.append(el.tag)
        
        return tree,ordinale

  
    def passe_tags_an(self,ordinale):
        if self.mb.debug: log(inspect.stack)
        
        tags_neu = copy.deepcopy(self.mb.tags)
        
        new_dict = {}
        
        for ordn in tags_neu['ordinale']:
            if ordn in ordinale:
                new_dict.update({ordn : tags_neu['ordinale'][ordn]})
        
        tags_neu['ordinale'] = new_dict
        
        return tags_neu
        
        
            
    
    def pruefe_neuen_projekt_namen(self):
        if self.mb.debug: log(inspect.stack)
       
        neuer_projekt_name = self.feld_projekt_name.Text
        if neuer_projekt_name == '':
            self.mb.nachricht(LANG.KEIN_NAME,"warningbox")
            return False
        
        speicherort = self.mb.settings_exp['speicherort']
        
        if speicherort == '':
            self.mb.nachricht(LANG.KEIN_SPEICHERORT,"warningbox")
            return False
    
        sysPath = uno.fileUrlToSystemPath(speicherort)
        
        path = os.path.join(sysPath,neuer_projekt_name+'.organon')
        
        if os.path.exists(path):
            self.mb.nachricht(LANG.ORDNER_EXISTIERT_SCHON %neuer_projekt_name,"warningbox")
            return False
        else:
            return True

        
    
    def teste_auf_verbliebene_bereiche(self,sections):
        if self.mb.debug: log(inspect.stack)
        
        regex = re.compile('OrganonSec')
        matches = [string for string in sections if re.match(regex, string)]
        
        if len(matches) == 0:
            return True
        else:
            return False


    def pruefe_dateiexistenz(self,pfad,dateierweiterung = None):
        if self.mb.debug: log(inspect.stack)
        
        if dateierweiterung != None:
            i = 0
            while os.path.exists(pfad+dateierweiterung):
                pfadX = pfad.split(dateierweiterung)[0]
                if pfadX[-1] == ')':                                    
                    sub = re.sub(r"\([0-9]+\)\Z",'('+str(i)+')', pfadX) 
                    if sub == pfad:
                        pfad = pfad + '(%s)' %i
                    else:
                        pfad = sub
                else:
                    pfad = pfad + '(%s)' %i
                i += 1
            return pfad
        
        else:
            i = 0
            while os.path.exists(pfad):
                if pfad[-1] == ')':                                    
                    sub = re.sub(r"\([0-9]*\)\Z",'('+str(i)+')', pfad) 
                    if sub == pfad:
                        pfad = pfad + '(%s)' %i
                    else:
                        pfad = sub
                else:
                    pfad = pfad + '(%s)' %i
                i += 1
            return pfad
        
        
    def disposing(self,ev):
        return False


                        
from com.sun.star.lang import XEventListener
class AB_Fenster_Dispose_Listener(unohelper.Base, XEventListener):
    # Listener um Position zu bestimmen
    def __init__(self,mb,cl_exp):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.cl_exp = cl_exp
        
    def disposing(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        if ev.Source.Model.Text == LANG.TRENNER_TIT:
            self.cl_exp.trenner_fenster = None
        if ev.Source.Model.Text == LANG.AUSWAHL:
            self.cl_exp.auswahl_fenster = None
            
        if ev.Source.Model.Text == LANG.EXPORT:
            self.cl_exp.haupt_fenster = None
            
            if self.cl_exp.auswahl_fenster != None:
                self.cl_exp.auswahl_fenster.dispose()
                self.cl_exp.auswahl_fenster = None
            if self.cl_exp.trenner_fenster != None:
                self.cl_exp.trenner_fenster.dispose()
                self.cl_exp.trenner_fenster = None

            # Settings speichern
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
            
    
class A_Trenner_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,cl_exp):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.exp_fenster = None
        self.cl_exp = cl_exp
        
    def disposing(self,ev):
        return False

        
    def actionPerformed(self,ev):      
        if self.mb.debug: log(inspect.stack)
        
        if self.cl_exp.trenner_fenster != None:
            self.cl_exp.trenner_fenster.toFront()
            return
        self.dialog_dokument_exp()
        
    
    def dialog_dokument_exp_elemente(self,listener,listenerLB):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_exp
        
        cb_listener = A_Trenner_CheckBox_Listener(self.mb)  
        listenerLZ = A_Anz_Leerzeilen_Focus_Listener(self.mb) 
        
          

        pStyles = self.mb.doc.StyleFamilies.ParagraphStyles
        style_names = pStyles.ElementNames
        index = style_names.index(sett['style_ord'])
        index2 = style_names.index(sett['style_dat'])

        if self.mb.settings_exp['url'] != '':
            try:
                label_url = uno.fileUrlToSystemPath(decode_utf(sett['url']))
            except:
                label_url = ''
        else:
            label_url = ''

        y = 0
        
        # Titel
        controls = [
        10,
        ('controlE',"FixedText",1,
                            'tab0x',y ,80,20,
                            ('Label','FontWeight'),
                            (LANG.TRENNER_TIT,150),
                            {}
                            ),   

        40,

        # Ordner
        ('controlO',"CheckBox",1,
                            'tab0' ,y,80,22,
                            ('Label','State'),
                            (LANG.ORDNERTITEL,sett['ordnertitel']),
                            {'setActionCommand':'ordnertitel','addActionListener':(cb_listener)} 
                            ),  
        0,
        
        ('controlF',"CheckBox",1,
                            'tab1' ,y,80,22,
                            ('Label','State'),
                            (LANG.FORMAT,sett['format_ord']),
                            {'setActionCommand':'format_ord','addActionListener':(cb_listener)} 
                            ),  


        0,
        
        # Liste der Formate
        ('controlL',"ListBox",0,
                            'tab2',y ,100,20,
                            ('Dropdown',),
                            (True,),
                            {'addItems':(style_names,0),'SelectedItems':(index,),'addItemListener':listenerLB}
                            ),  

        
        
        30,
        
        # Datei
        ('controlD',"CheckBox",1,
                            'tab0' ,y,80,22,
                            ('Label','State'),
                            (LANG.DATEITITEL,sett['dateititel']),
                            {'setActionCommand':'dateititel','addActionListener':(cb_listener)} 
                            ),  
        
        0,
        
        ('controlF2',"CheckBox",1,
                            'tab1' ,y,160,22,
                            ('Label','State'),
                            (LANG.FORMAT,sett['format_dat']),
                            {'setActionCommand':'format_dat','addActionListener':(cb_listener)} 
                            ),  

        0,
        
        # Liste der Formate
        ('controlL2',"ListBox",0,
                            'tab2',y ,100,20,
                            ('Dropdown',),
                            (True,),
                            {'addItems':(style_names,0),'SelectedItems':(index2,),'addItemListener':listenerLB}
                            ),  
        
        50,
        # DOKUMENT
        ('controlDFix',"FixedText",1,
                            'tab0',y ,200,20,
                            ('Label','FontWeight'),
                            (LANG.ORT_DES_DOKUMENTS,150),
                            {}
                            ),  
        50,
        
        ('controlL3',"CheckBox",1,
                            'tab0' ,y,160,22,
                            ('Label','State'),
                            (LANG.LEERZEILEN,sett['leerzeilen_drunter']),
                            {'setActionCommand':'leerzeilen_drunter','addActionListener':(cb_listener)} 
                            ),  

        0,
        
        ('controlA2',"Edit",1,
                            'tab1' ,y,20,20,
                            ('HelpText','Text'),
                            (LANG.ANZAHL_LEERZEILEN,str(sett['anz_drunter'])),
                            {'addFocusListener':(listenerLZ)}
                            ),  

        50,
        
        ('controlDo',"CheckBox",1,
                            'tab0x' ,y,160,22,
                            ('Label','State'),
                            (LANG.DOK_EINFUEGEN,sett['dok_einfuegen']),
                            {'setActionCommand':'dok_einfuegen','addActionListener':(cb_listener)} 
                            ),  
        0,
        
        # Button
        ('controlD2',"Button",1,
                            'tab2',y-3,100,22,
                            ('Label',),
                            (LANG.WAEHLEN,),
                            {'addActionListener':(listener)}
                            ),  
        
        
        25,
        
        ('controlF3',"FixedText",0,
                            'tab0x+20' ,y,300,55,
                            ('HelpText','Label','MultiLine'),
                            ('URL',label_url,True),
                            {}
                            ),  

        55,
        
        ('controlSB',"CheckBox",1,
                            'tab0x' ,y,200,22,
                            ('Label','State'),
                            (LANG.SEITENUMBRUCH_ORD,sett['seitenumbruch_ord']),
                            {'setActionCommand':'seitenumbruch_ord','addActionListener':(cb_listener)} 
                            ),  
        
        20,
        
        ('controlSb2',"CheckBox",1,
                            'tab0x' ,y,200,22,
                            ('Label','State'),
                            (LANG.SEITENUMBRUCH_DAT,sett['seitenumbruch_dat']),
                            {'setActionCommand':'seitenumbruch_dat','addActionListener':(cb_listener)} 
                            ),  
        10,
        ]
        
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 10),
                 1 : (None, 10),
                 2 : (None, 5),
                 3 : (None, 5),
                 4 : (None, 5),
                 }
        
        abstand_links = 10
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
          
        return controls2,max_breite
        
    def dialog_dokument_exp(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = A_TrennDatei_Button_Listener(self.mb)   
            listenerLB = A_ParaStyle_Item_Listener(self.mb)
        
            controls,max_breite = self.dialog_dokument_exp_elemente(listener,listenerLB)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)  
            
            
            posSize = berechne_pos(self.mb,self.cl_exp,self.exp_fenster,'Trenner')
            posSize = posSize[0],posSize[1],max_breite,max_hoehe
            
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            fenster_cont.Model.Text = LANG.TRENNER_TIT
            
            listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
            fenster_cont.addEventListener(listenerF)
            
            self.cl_exp.trenner_fenster = fenster
            
            listener.model = ctrls['controlF3'].Model
            listenerLB.cont_ord = ctrls['controlL']
            listenerLB.cont_dat = ctrls['controlL2']
            
            # Controls in Hauptfenster eintragen
            for name,c in ctrls.items():
                fenster_cont.addControl(name,c)
        except:
            log(inspect.stack,tb())
            
        
        
    def oeffne_Trenner_fenster(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        
        
        sett = self.mb.settings_exp
        cb_listener = A_Trenner_CheckBox_Listener(self.mb)      
          
        posSize = berechne_pos(self.mb,self.cl_exp,self.exp_fenster,'Trenner')
        posSize = posSize[0],posSize[1],320,360
        fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
        fenster_cont.Model.Text = LANG.TRENNER_TIT
        
        listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
        fenster_cont.addEventListener(listenerF)
        
        self.cl_exp.trenner_fenster = fenster

        y = 10
        
        # Titel
        controlE, modelE = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,80,20,(),() )  
        controlE.Text = LANG.TRENNER_TIT
        modelE.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlE)
        
        y += 40
        
        # Ordner
        controlO, modelO = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,80,22,(),() )  
        modelO.Label = LANG.ORDNERTITEL
        modelO.State = sett['ordnertitel']
        controlO.ActionCommand = 'ordnertitel'
        controlO.addActionListener(cb_listener)
        fenster_cont.addControl('Ordnertitel', controlO)
        
        controlF, modelF = self.mb.createControl(self.mb.ctx,"CheckBox",20 + 100 ,y,160,22,(),() )  
        modelF.Label = LANG.FORMAT
        modelF.State = sett['format_ord']
        controlF.ActionCommand = 'format_ord'
        controlF.addActionListener(cb_listener)
        fenster_cont.addControl('Format', controlF)
        
            # Liste der Formate
        controlL, modelL = self.mb.createControl(self.mb.ctx,"ListBox",20 + 180,y -3 ,100,20,(),() )  
        #controlL.setMultipleMode(False)
        pStyles = self.mb.doc.StyleFamilies.ParagraphStyles
        style_names = pStyles.ElementNames
        controlL.addItems(style_names,0)
        modelL.Dropdown = True
        index = style_names.index(sett['style_ord'])
        modelL.SelectedItems = index,
        fenster_cont.addControl('Liste_Ord', controlL)
        
        y += 30
        
        # Datei
        controlD, modelD = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,80,22,(),() )  
        modelD.Label = LANG.DATEITITEL
        modelD.State = sett['dateititel']
        controlD.ActionCommand = 'dateititel'
        controlD.addActionListener(cb_listener)
        fenster_cont.addControl('Dateititel', controlD)
        
        controlF2, modelF2 = self.mb.createControl(self.mb.ctx,"CheckBox",20 + 100 ,y,160,22,(),() )  
        modelF2.Label = LANG.FORMAT
        modelF2.State = sett['format_dat']
        controlF2.ActionCommand = 'format_dat'
        controlF2.addActionListener(cb_listener)
        fenster_cont.addControl('Format2', controlF2)
        
        # Liste der Formate
        controlL2, modelL2 = self.mb.createControl(self.mb.ctx,"ListBox",20 + 180,y -3 ,100,20,(),() )  
        controlL2.addItems(style_names,0)
        modelL2.Dropdown = True
        index = style_names.index(sett['style_dat'])
        modelL2.SelectedItems = index,
        fenster_cont.addControl('Liste_Dat', controlL2)
        
        # Listener fuer beide Stylelisten
        listenerLB = A_ParaStyle_Item_Listener(self.mb,controlL,controlL2)
        controlL.addItemListener(listenerLB)
        controlL2.addItemListener(listenerLB)
        
        y += 50
        
        # DOKUMENT
        controlD, modelD = self.mb.createControl(self.mb.ctx,"FixedText",100,y ,200,20,(),() )  
        controlD.Text = LANG.ORT_DES_DOKUMENTS
        modelD.FontWeight = 200.0
        fenster_cont.addControl('Titel2', controlD)
        
        y += 50
        
        controlL2, modelL2 = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,160,22,(),() )  
        modelL2.Label = LANG.LEERZEILEN
        modelL2.State = sett['leerzeilen_drunter']
        controlL2.ActionCommand = 'leerzeilen_drunter'
        controlL2.addActionListener(cb_listener)
        fenster_cont.addControl('Leerzeilen2', controlL2)     
        
        controlA2, modelA2 = self.mb.createControl(self.mb.ctx,"Edit",120 ,y,20,30,(),() )  
        modelA2.HelpText = LANG.ANZAHL_LEERZEILEN
        modelA2.Text = sett['anz_drunter']
        listenerLZ = A_Anz_Leerzeilen_Focus_Listener(self.mb)
        controlA2.addFocusListener(listenerLZ)
        fenster_cont.addControl('Anzahl', controlA2) 
        
        y += 50
        
        controlDo, modelDo = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,160,22,(),() )  
        modelDo.Label = LANG.DOK_EINFUEGEN
        modelDo.State = sett['dok_einfuegen']
        controlDo.ActionCommand = 'dok_einfuegen'
        controlDo.addActionListener(cb_listener)
        fenster_cont.addControl('Dokument', controlDo)
        
        # Button
        controlD, modelD = self.mb.createControl(self.mb.ctx,"Button",200,y-3,100,22,(),() )  
        controlD.Label = LANG.WAEHLEN
        fenster_cont.addControl('Dok', controlD)
        
        y += 20
        
        controlF, modelF = self.mb.createControl(self.mb.ctx,"FixedText",40 ,y,500,22,(),() )  
        modelF.HelpText = 'URL'

        if self.mb.settings_exp['url'] != '':
            modelF.Label = uno.fileUrlToSystemPath(decode_utf(sett['url']))
        fenster_cont.addControl('Anzahl', controlF) 
        
        listener = A_TrennDatei_Button_Listener(self.mb,modelF)
        controlD.addActionListener(listener)
        
        y += 40
        
        controlSB, modelSB = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,200,22,(),() )  
        modelSB.Label = LANG.SEITENUMBRUCH_ORD
        modelSB.State = sett['seitenumbruch_ord']
        controlSB.ActionCommand = 'seitenumbruch_ord'
        controlSB.addActionListener(cb_listener)
        fenster_cont.addControl('seitenumbruch_ord', controlSB) 
        
        y += 20
        
        controlSb2, modelSb2 = self.mb.createControl(self.mb.ctx,"CheckBox",20 ,y,200,22,(),() )  
        modelSb2.Label = LANG.SEITENUMBRUCH_DAT
        modelSb2.State = sett['seitenumbruch_dat']
        controlSb2.ActionCommand = 'seitenumbruch_dat'
        controlSb2.addActionListener(cb_listener)
        fenster_cont.addControl('seitenumbruch_dat', controlSb2) 
        

class A_Trenner_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def disposing(self,ev):
        return False

        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_exp
        sett[ev.ActionCommand] = self.toggle(sett[ev.ActionCommand])

    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:
            return 1


class A_TrennDatei_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model = None
        
    def disposing(self,ev):
        return False

        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        filepath,ok = self.mb.class_Funktionen.filepicker2()
        
        if not ok:
            return
        
        self.mb.settings_exp['url'] = filepath
        self.model.Label = uno.fileUrlToSystemPath(filepath)



class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model = None
        
    def disposing(self,ev):
        return False

        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        Filepicker.setDisplayDirectory(self.mb.settings_exp['speicherort'])
        Filepicker.execute()
        
        if Filepicker.Directory == '':
            return
        
        filepath = Filepicker.getDirectory()
        
        self.mb.settings_exp['speicherort'] = filepath
        self.model.Label = uno.fileUrlToSystemPath(filepath)
        
            
        
class A_ParaStyle_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.cont_ord = None
        self.cont_dat = None
    
    def disposing(self,ev):
        return False

        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
        
        if ev.Source == self.cont_ord:
            self.mb.settings_exp['style_ord'] = ev.Source.Items[ev.Selected] 
        elif ev.Source == self.cont_dat:
            self.mb.settings_exp['style_dat'] = ev.Source.Items[ev.Selected]     
    
    def disposing(self,ev):
        return False
      
       
    
class A_Anz_Leerzeilen_Focus_Listener(unohelper.Base, XFocusListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
    
    def disposing(self,ev):
        return False

        
    # XItemListener    
    def focusLost(self, ev): 
        if self.mb.debug: log(inspect.stack)

        if ev.Source.Model.Text.isdigit():
            self.mb.settings_exp['anz_drunter'] = int(ev.Source.Model.Text)
        else:
            ev.Source.Model.Text = self.mb.settings_exp['anz_drunter']
                
    def focusGained(self,ev):
        return False  



def berechne_pos(mb,cl_exp,exp_fenster,Rufer):

    anderes_fenster = None
    
    if Rufer == 'Auswahl':
        if cl_exp.trenner_fenster != None:
            anderes_fenster = cl_exp.trenner_fenster
    elif Rufer == 'Trenner':
        if cl_exp.auswahl_fenster != None:
            anderes_fenster = cl_exp.auswahl_fenster
         
    if anderes_fenster != None:
        posSizePlus = anderes_fenster.PosSize
        XPlus = posSizePlus.Width + 20
    else:
        XPlus = 0
    
    posSize_main = mb.desktop.ActiveFrame.ContainerWindow.PosSize
    posSize_expWin = exp_fenster.PosSize
    
    X = posSize_expWin.X + XPlus
    Y = posSize_expWin.Y
    Width = posSize_expWin.Width
    Height = posSize_main.Height - 40

    posSize = X + Width + 20,Y,Width,Height
    
    return posSize



class B_Auswahl_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,cl_exp):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.exp_fenster = None
        self.cl_exp = cl_exp
        
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        try:
            if self.cl_exp.auswahl_fenster != None:
                return
            
            sett = self.mb.settings_exp
            
            posSize = berechne_pos(self.mb,self.cl_exp,self.exp_fenster,'Auswahl')
            
            # Dict von alten Eintraegen bereinigen
            eintr = []
            for ordinal in sett['ausgewaehlte']:
                if ordinal not in self.mb.props[T.AB].dict_bereiche['ordinal']:
                    eintr.append(ordinal)
            for ordn in eintr:
                del sett['ausgewaehlte'][ordn]
            
            
            
            pos = posSize[0],posSize[1]
            
            (y,
             fenster,
             fenster_cont,
             control_innen,
             ctrls) = self.mb.class_Fenster.erzeuge_treeview_mit_checkbox(
                                                                    tab_name=T.AB,
                                                                    pos=pos,
                                                                    auswaehlen=True)
            
            # Listener um Position zu bestimmen
            listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
            
            fenster_cont.addEventListener(listenerF)
            self.cl_exp.auswahl_fenster = fenster
            

        except:
            log(inspect.stack,tb())

    
    def disposing(self,ev):
        return False
            
            
      
        
            

