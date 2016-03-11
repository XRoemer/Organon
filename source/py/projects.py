# -*- coding: utf-8 -*-

import unohelper
from com.sun.star.beans import PropertyValue

'''
PFADE:

extern gespeichert:
alle in Sidebar,export_settings,import_settings: FileURL
 // pfade.txt in OO...uno_packages ... Organon: SystemPath


intern:
mb.path_to_extension,dict mb.pfade, mb.projekt_path : SystemPath

'''

class Projekt():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.ctx = mb.ctx
        self.mb = mb
        
        self.mb.settings_proj['use_template'] = (False,None)
        self.mb.settings_proj['use_template_organon'] = (False,None)           
        self.first_time = True


    def dialog_neues_projekt_anlegen_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.settings_proj['use_template'] = 0,''
            
            if self.mb.speicherort_last_proj != None:
                # try/except fuer Ubuntu: U meldet Fehler: couldn't convert fileUrlTo ...
                # -> gespeicherten Pfad ueberpruefen!
                try:
                    speicherort = uno.fileUrlToSystemPath(self.mb.speicherort_last_proj)
                except:
                    #modelU.Label = '-' 
                    speicherort = self.mb.speicherort_last_proj
            else:
                speicherort = '-' 
            
            
            if self.mb.writer_vorlagen == {}:
                writer_vorlagen = (LANG.NO_TEMPLATES,)
                writer_enable = False
            else:
                writer_vorlagen = tuple(self.mb.writer_vorlagen)
                writer_enable = True
                
            
            templs = self.mb.settings_orga['templates_organon']
            organon_enable = len(templs['templates']) > 0
            templs_Org = tuple(templs['templates'])
            
            
            controls = (
            20,
            ('control',"FixedText",1,        
                                'tab0',0,250,20,  
                                ('Label','FontWeight'),
                                (LANG.ENTER_PROJ_NAME ,150),          
                                {} 
                                ),
            30,
            ('prj_name',"Edit",0,            
                                'tab0x-tab0-E',0,200,20,   
                                (),
                                (),                                                       
                                {} 
                                ) ,
            30,
            ('controlT3',"FixedLine",0,      
                                'tab0x-max',0,360,40,   
                                (),
                                (),                                                       
                                {} 
                                ), 
            43,

            ('controlP',"FixedText",1,       
                                'tab0',0,20,20,  
                                ('Label','FontWeight'),
                                (LANG.SPEICHERORT,150),               
                                {} 
                                ),  
            0,
            ('controlW',"Button",1,          
                                'tab2',0,80,20,   
                                ('Label',),
                                (LANG.AUSWAHL,),                                  
                                {'setActionCommand':LANG.WAEHLEN,'addActionListener':(listener)}
                                ),              
            30,
            ('speicherort',"FixedText",0,        
                                'tab0x-max',0,300,50,   
                                ('Label','MultiLine'),
                                (speicherort,True),           
                                {} 
                                ), 
            30,
            ('controlT',"FixedLine",0,       
                                'tab0x-max',0,360,40,   
                                (),
                                (),                                                       
                                {} 
                                ), 
            40,
            
            ('controlFormO',"FixedText",1,    
                                 'tab0',0,80,20,   
                                 ('Label','FontWeight','Enabled','HelpText'),
                                 (LANG.TEMPLATES_ORGANON,150,organon_enable,LANG.ORG_TEMPLATES_SETZEN),              
                                 {} 
                                 ),
            0,  
            ('organon_cb',"CheckBox",1,    
                                 'tab2',0,200,20,   
                                 ('Label','Enabled'),
                                 (LANG.NUTZEN,organon_enable),                                  
                                 {'setActionCommand':'organon','addActionListener':(listener)} 
                                 ) ,
            22,
            ('organon_lb',"ListBox",1,       
                                 'tab2',0,100,20,    
                                 ('Dropdown','Enabled','HelpText'),
                                 (True,organon_enable,LANG.ORG_TEMPLATES_SETZEN),                                       
                                 {'addItems':(templs_Org,0),'SelectedItems':(0,)}
                                 ),
            20,
            ('controlTO4',"FixedLine",0,      
                                 'tab0x-max',0,360,40,   
                                 (),
                                 (),                                                       
                                 {} 
                                 ), 
            40,
            
            ('controlForm',"FixedText",1,     
                                 'tab0',0,80,20,   
                                 ('Label','FontWeight','HelpText','Enabled'),
                                 (LANG.TEMPLATES_WRITER,150,LANG.WRITER_TEMPLATES_SETZEN,writer_enable),              
                                 {} 
                                 ),
            0,  
            ('writer_cb',"CheckBox",1,    
                                 'tab2',0,200,20,   
                                 ('Label','Enabled'),
                                 (LANG.NUTZEN,writer_enable),                                   
                                 {'setActionCommand':'writer','addActionListener':(listener)} 
                                 ) ,
            22,
            ('writer_lb',"ListBox",1,   
                                 'tab2',0,100,20,    
                                 ('Dropdown','HelpText','Enabled'),
                                 (True,LANG.WRITER_TEMPLATES_SETZEN,writer_enable),                                       
                                 {'addItems':(writer_vorlagen,0),'SelectedItems':(0,)}
                                 ),
            20,
            ('controlT4',"FixedLine",0,      
                                 'tab0x-max',0,360,40,   
                                 (),
                                 (),                                                       
                                 {} 
                                 ), 
            40,
            
            ('ok',"Button",1,          
                                 'tab2-tab2-E',0,80,30,    
                                 ('Label',),
                                 (LANG.OK,),                                       
                                 {'setActionCommand':LANG.OK,'addActionListener':(listener)} 
                                 ), 
            0,
            )
        
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 50),
                     1 : (None, 0),
                     2 : (None, 0),
                     }
            
            abstand_links = 10
            
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite
        except:
            log(inspect.stack,tb())
            
    
    def dialog_neues_projekt_anlegen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # prueft, ob eine Organon Datei geladen ist
            if len(self.mb.props[T.AB].dict_bereiche) != 0:
                Popup(self.mb, 'warning').text = LANG.PRUEFE_AUF_GELADENES_ORGANON_PROJEKT
                return 
            
            self.get_writer_vorlagen()
            
            # LISTENER
            #listenerS = Speicherordner_Button_Listener(self.mb)
            listener = neues_Projekt_Dialog_Listener(self.mb) 
            
            controls,max_breite = self.dialog_neues_projekt_anlegen_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls)
            
            # Hauptfenster erzeugen
            posSize = None, None, max_breite, max_hoehe + 20
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            #fenster_cont.Model.Text = LANG.EXPORT

            # Controls in Hauptfenster eintragen
            for name,c in ctrls.items():
                fenster_cont.addControl(name,c)

            listener.ctrls = ctrls
            listener.fenster = fenster
        except:
            log(inspect.stack,tb())
                    
        
    def erzeuge_neues_Projekt(self,is_template, templ_art, templ_pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ok = self.erzeuge_Ordner_Struktur() 
            if not ok:
                return
            
            self.erzeuge_import_Settings()
            self.erzeuge_export_Settings()  
            self.erzeuge_proj_Settings()
            #self.mb.class_Querverweise.erzeuge_organon_meta_graph()
            
            self.mb.class_Bereiche.leere_Dokument()        
            self.mb.class_Baumansicht.start()             
            Eintraege = self.beispieleintraege2()
            
            self.erzeuge_Projekt_xml_tree() 
            self.mb.class_Bereiche.erzeuge_leere_datei()               
            self.erzeuge_Eintraege_und_Bereiche(Eintraege)
            
            Path1 = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
            self.mb.tree_write(self.mb.props[T.AB].xml_tree,Path1)
            
            if is_template and templ_art == 'writer':
                self.template_kopieren(templ_pfad)
                
            self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
            
            erste_datei = self.mb.tabsX.get_erste_datei(T.AB)
            props = self.mb.props[T.AB]
            props.selektierte_zeile = erste_datei
            props.selektierte_zeile_alt = erste_datei
            self.mb.Listener.VC_selection_listener.bereichsname_alt = props.dict_bereiche['ordinal'][erste_datei]
            
            self.mb.class_Tags.lege_tags_an()
            self.mb.class_Tags.speicher_tags()

            self.mb.class_Baumansicht.korrigiere_scrollbar()

            dateiname = "{}.organon".format(self.mb.projekt_name)
            filepath = os.path.join(self.mb.pfade['projekt'], dateiname)
            self.trage_projekt_in_zuletzt_geladene_Projekte_ein(dateiname,filepath)
            
            if is_template:
                self.mb.class_Projekt.mit_template_oeffnen(templ_pfad)
                return
            else:
                erste_datei = self.mb.tabsX.get_erste_datei(T.AB)
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche(erste_datei,action='lade_projekt')
                
                # ausgewaehlte zeile einfaerben
                zeile = props.Hauptfeld.getControl(erste_datei)
                textfeld = zeile.getControl('textfeld')
                textfeld.Model.BackgroundColor = KONST.FARBE_AUSGEWAEHLTE_ZEILE 
            
                self.mb.Listener.starte_alle_Listener()
                self.mb.use_UM_Listener = True
                
                
        except Exception as e:
            log(inspect.stack,tb())
            Popup(self.mb, 'error').text = 'erzeuge_neues_Projekt ' + str(e)
            
            
    def beispieleintraege2(self):
        if self.mb.debug: log(inspect.stack)
        
        Eintraege = [
                ('nr0','root',self.mb.projekt_name,0,'prj','auf','ja','leer','leer','leer'),
                ('nr1','nr0',LANG.TITEL,1,'pg','-','ja','leer','leer','leer'),
                ('nr2','nr0',LANG.KAPITEL+' 1',1,'dir','auf','ja','leer','leer','leer'),
                ('nr3','nr2',LANG.SZENE + ' 1',2,'pg','-','ja','leer','leer','leer'),
                ('nr4','nr2',LANG.SZENE + ' 2',2,'pg','-','ja','leer','leer','leer'),
                ('nr5','root',LANG.PAPIERKORB,0,'waste','zu','ja','leer','leer','leer')]
        
        return Eintraege        

    
    def besitzt_template(self):
        '''
        Die Vorlage haette auch aus settings_proj ausgelesen werden koennen.
        Die Methode "besitzt_template" hat jedoch den Vorteil, dass Vorlagen
        auch im Nachhinein eingefuegt, geaendert oder entfernt werden koennen.
        '''
        if self.mb.debug: log(inspect.stack)
        
        args = self.mb.doc.Args
        pruef_pfad = ''
        for a in args:
            if a.Name == 'URL':
                pruef_pfad = a.Value
        
        # Wenn der Pfad nicht leer ist, wurde das Dokument
        # als Template von Organon gestartet
        if pruef_pfad != '':
            return False,''
        
        odt_pfad = self.mb.pfade['odts']
        
        files = []
        for (dirpath, dirnames, filenames) in os.walk(odt_pfad):
            files.extend(filenames)
            break
        
        if 'template.ott' in filenames:
            templ_pfad = os.path.join(dirpath,'template.ott')
            self.mb.settings_proj['use_template'] = [1,templ_pfad]
            return True,templ_pfad
         
        self.mb.settings_proj['use_template'] = [0,'']
        return False,''
            
    
    def template_kopieren(self,templ_pfad):
        if self.mb.debug: log(inspect.stack)
        
        neuer_templ_pfad = os.path.join(self.mb.pfade['odts'],'template.ott')
        
        import shutil
        shutil.copyfile(templ_pfad, neuer_templ_pfad)
        
        self.mb.settings_proj['use_template'] = [1,neuer_templ_pfad]
    
                      
    def setze_pfade(self): 
        if self.mb.debug: log(inspect.stack)
        
        paths = self.mb.createUnoService( "com.sun.star.util.PathSettings" )
        pHome = paths.Work_writable
        if sys.platform == 'linux':
            os.chdir( '//')        

        pOrganon = self.mb.projekt_path

        pProjekt =  os.path.join(pOrganon , '%s.organon' % self.mb.projekt_name)
        pFiles =    os.path.join(pProjekt , 'Files')
        pOdts =     os.path.join(pFiles , 'odt')
        pImages =   os.path.join(pFiles , 'Images')
        pIcons =    os.path.join(pFiles , 'Icons')
        pSettings = os.path.join(pProjekt , 'Settings')
        pTabs =     os.path.join(pSettings , 'Tabs')
        pPlain_Text=os.path.join(pFiles , 'plain_txt')
        
        
        self.mb.pfade.update({'home':pHome}) 
        self.mb.pfade.update({'projekt':pProjekt})      
        self.mb.pfade.update({'organon':pOrganon})
        self.mb.pfade.update({'files':pFiles})
        self.mb.pfade.update({'odts':pOdts})
        self.mb.pfade.update({'settings':pSettings}) 
        self.mb.pfade.update({'images':pImages}) 
        self.mb.pfade.update({'tabs':pTabs}) 
        self.mb.pfade.update({'icons':pIcons}) 
        self.mb.pfade.update({'plain_txt':pPlain_Text}) 

    
    def lade_settings(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            pfad = os.path.join(self.mb.pfade['settings'],'export_settings.txt')            
            with codecs_open(pfad , "r","utf-8") as f:
                txt = f.read()
            self.mb.settings_exp = eval(txt)
            self.mb.settings_exp['ausgewaehlte'] = {}

            pfad = os.path.join(self.mb.pfade['settings'],'import_settings.txt')
            with codecs_open(pfad , "r","utf-8") as f:
                txt = f.read()
            self.mb.settings_imp = eval(txt)
        
            pfad = os.path.join(self.mb.pfade['settings'],'project_settings.txt')
            with codecs_open(pfad , "r","utf-8") as f:
                txt = f.read()
            self.mb.settings_proj = eval(txt)
        
        except:
            log(inspect.stack,tb())
        

    def erzeuge_Ordner_Struktur(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
        
            pfade = self.mb.pfade
            # Organon
            if not os.path.exists(pfade['organon']):
                os.makedirs(pfade['organon'])
            # Organon/<Projekt Name>
            if not os.path.exists(pfade['projekt']):
                os.makedirs(pfade['projekt'])
            # Organon/<Projekt Name>/Files
            if not os.path.exists(pfade['files']):
                os.makedirs(pfade['files'])
            # Organon/<Projekt Name>/Files
            if not os.path.exists(pfade['odts']):
                os.makedirs(pfade['odts'])   
            # Organon/<Projekt Name>/Files
            if not os.path.exists(pfade['images']):
                os.makedirs(pfade['images'])  
            # Organon/<Projekt Name>/Settings
            if not os.path.exists(pfade['settings']):
                os.makedirs(pfade['settings'])
            # Organon/<Projekt Name>/Settings/Tags
            if not os.path.exists(pfade['tabs']):
                os.makedirs(pfade['tabs'])
            # Organon/<Projekt Name>/Settings/Tags
            if not os.path.exists(pfade['icons']):
                os.makedirs(pfade['icons'])
            if not os.path.exists(pfade['plain_txt']):
                os.makedirs(pfade['plain_txt'])
    
            # Datei anlegen, die bei lade_Projekt angesprochen werden soll
            path = os.path.join(pfade['projekt'],"%s.organon" % self.mb.projekt_name)
            with open(path, "w") as f:
                f.write('Dies ist eine Organon Datei. Goennen Sie ihr ihre Existenz.') 
                 
            # Setzen einer UserDefinedProperty um Projekt identifizieren zu koennen
            UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
            has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
            if has_prop:
                UD_properties.setPropertyValue('ProjektName',self.mb.projekt_name) 
            else:
                UD_properties.addProperty('ProjektName',1,self.mb.projekt_name) 
                     
            # damit von den Bereichen in die Datei verlinkt wird, muss sie gespeichert werden   
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name))
            Path2 = uno.systemPathToFileUrl(Path1)
    
            self.mb.doc.storeAsURL(Path2,())  
             
        except Exception as e:
            Popup(self.mb, 'error').text = "ERROR: " + str(e)
            return False
        
        return True
        

    def get_writer_vorlagen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            paths = self.mb.createUnoService( "com.sun.star.util.PathSettings" )
            template = paths.Template
            if ';' in template:
                template = template.split(';')
            
            if self.mb.programm == 'LibreOffice':
                temp = []
                for path in template:
                    if 'vnd.sun.star.expand' not in path:
                        if ('/../') in path:
                            t = ''.join(path.split('program/../'))
                            temp.append(t)
                        else:
                            temp.append(path)
                        
                template = temp    
    
            writer_vorlagen = {}
            
            for path in template:
                if 'Roaming' in path:
                    self.mb.user_template_path = path
                
                if os.path.exists(uno.fileUrlToSystemPath(path)):
                    files = os.listdir(uno.fileUrlToSystemPath( path ))
                    for filename in files:
                        
                        pfad = os.path.join(uno.fileUrlToSystemPath(path),filename)
                        erweiterung = os.path.splitext(pfad)[1]
                        if erweiterung == '.ott':
                            dateiname = os.path.split(pfad)[1]
                            dateiname1 = dateiname.split(erweiterung)[0]
                            
                            writer_vorlagen.update({ dateiname1 : pfad })
                            
            self.mb.writer_vorlagen = writer_vorlagen
        except:
            log(inspect.stack,tb())


    def lade_Projekt(self,filepicker = True, filepath = ''):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if self.pruefe_auf_geladenes_organon_projekt():
                return

            if filepicker:
                ofilter = ('Organon Project','*.organon')
                filepath,ok = self.mb.class_Funktionen.filepicker2(ofilter=ofilter,url_to_sys=True)
                
                if not ok:
                    return
            
            dateiname = os.path.basename(filepath)
            dateiendung = os.path.splitext(filepath)[1]
            
            self.mb.projekt_name = dateiname.split(dateiendung)[0]
            self.mb.projekt_path = os.path.dirname(os.path.dirname(filepath))  

            self.setze_pfade()
            
            self.mb.class_Bereiche.leere_Dokument() 
            self.lade_settings() 
                        
            has_template,templ_pfad = self.besitzt_template()
            
            if has_template:
                self.mit_template_oeffnen(templ_pfad,True)
                return

            Eintraege = self.lese_xml_datei()
            
            self.mb.class_Version.pruefe_version()
            
            self.mb.props[T.AB].Hauptfeld = self.mb.class_Baumansicht.erzeuge_Feld_Baumansicht(self.mb.prj_tab) 
            
            self.erzeuge_Eintraege_und_Bereiche2(Eintraege) 
            
            if self.mb.win.PosSize.Height > 20:
                # Wenn ein Template geoeffnet wurde, ist self.mb.win
                # an dieser Stelle noch nicht geoeffnet, daher Height == 0
                # schalte Sichtbarkeit wird daher uebersprungen, da ansonsten
                # alle Controls ausgeblendet werden
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_hf_ctrls()
            
            # setzt die selektierte Zeile auf die erste Datei
            erste_datei = self.mb.tabsX.get_erste_datei(T.AB)
            self.mb.props[T.AB].selektierte_zeile = erste_datei
            self.mb.props[T.AB].selektierte_zeile_alt = erste_datei
            self.mb.tabsX.setze_selektierte_zeile(erste_datei,T.AB)
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_ersten_Bereichs()
            self.mb.Listener.VC_selection_listener.bereichsname_alt = self.mb.props[T.AB].dict_bereiche['ordinal'][erste_datei]
            
            self.mb.class_Fenster.erzeuge_Scrollbar2()    
            self.mb.class_Mausrad.registriere_Maus_Focus_Listener(self.mb.props['ORGANON'].Hauptfeld.Context.Context)

            # Wenn die UDProp verloren gegangen sein sollte, wieder setzen
            UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
            has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
            if not has_prop:
                UD_properties.addProperty('ProjektName',1,self.mb.projekt_name) 
            
            # damit von den Bereichen in die Datei verlinkt wird, muss sie gespeichert werden   
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name))
            Path2 = uno.systemPathToFileUrl(Path1)
            self.mb.doc.storeAsURL(Path2,()) 
            
            self.mb.tabsX.lade_tabs()
            
            erste_datei = self.mb.tabsX.get_erste_datei(T.AB)
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche(erste_datei,action='lade_projekt')
            
            self.mb.class_Tags.lade_tags() 
             
            self.trage_projekt_in_zuletzt_geladene_Projekte_ein(dateiname,filepath)
            
            prj_ctrl = self.mb.tabsX.tableiste.getControl('ORGANON')
            prj_ctrl.Model.BackgroundColor = KONST.FARBE_TABS_SEL_HINTERGRUND
            prj_ctrl.Model.TextColor = KONST.FARBE_TABS_SEL_SCHRIFT
            
            self.mb.class_Baumansicht.korrigiere_scrollbar()
            
            self.mb.Listener.starte_alle_Listener()
            
            self.mb.class_Sidebar.erzeuge_sb_layout()
            self.mb.use_UM_Listener = True  
                        
        except Exception as e:
            log(inspect.stack,tb())
            log(inspect.stack,extras='Projekt nicht geladen\r\n' + str(e))
            if e.typeName == 'com.sun.star.task.ErrorCodeIOException':
                Popup(self.mb, 'error').text = LANG.ERROR_PROJECT_LOCKED.format(filepath) + str(e)
            else:
                Popup(self.mb, 'error').text = LANG.ERROR_LOAD_PROJECT + str(e)
        
    
    def lade_Projekt2(self):
        if self.mb.debug: log(inspect.stack)

        try:
            props = self.mb.props[T.AB]
            
            selektierte_zeile = props.selektierte_zeile
            
            self.mb.class_Funktionen.leere_hf()
            self.setze_pfade()
            self.mb.class_Bereiche.leere_Dokument() 
            self.lade_settings() 
                 
            props.Hauptfeld = self.mb.class_Baumansicht.erzeuge_Feld_Baumansicht(self.mb.prj_tab) 
               
            Eintraege = self.lese_xml_datei()
            self.erzeuge_Eintraege_und_Bereiche2(Eintraege) 
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_hf_ctrls()
            
            self.mb.class_Baumansicht.selektiere_zeile(selektierte_zeile, speichern = False)
            self.mb.Listener.VC_selection_listener.bereichsname_alt = self.mb.props[T.AB].dict_bereiche['ordinal'][selektierte_zeile]
            
            self.mb.class_Fenster.erzeuge_Scrollbar2()    
            self.mb.class_Baumansicht.korrigiere_scrollbar()
                         
            # damit von den Bereichen in die Datei verlinkt wird, muss sie gespeichert werden   
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name))
            Path2 = uno.systemPathToFileUrl(Path1)
            self.mb.doc.storeAsURL(Path2,()) 
                                     
        except Exception as e:
            Popup(self.mb, 'error').text = 'ERROR: could not load project\r\n ' + str(e)
            log(inspect.stack,tb())
            
            
    def trage_projekt_in_zuletzt_geladene_Projekte_ein(self,dateiname,filepath):
        if self.mb.debug: log(inspect.stack)
        
        zuletzt = self.mb.settings_orga['zuletzt_geladene_Projekte']
        
        try:
            # Fuer projekte erstellt vor v0.9.9.8b
            if isinstance(zuletzt, dict):
                list_proj = list(zuletzt)
                zuletzt = [[p,zuletzt[p]] for p in list_proj]
            
            name = dateiname.split('.organon')[0]
            
            # Handbuecher ausschliessen
            if name in ['Organon Handbuch','Organon Manual']:
                return
                                    
            neu = [name,filepath]
            
            if neu not in zuletzt:
                zuletzt.insert(0,[name,filepath])        
            else:
                index = zuletzt.index(neu)
                del(zuletzt[index])
                zuletzt.insert(0,[name,filepath])

            if len(zuletzt) > 9:
                del(zuletzt[-1])
            
            self.mb.settings_orga['zuletzt_geladene_Projekte'] = zuletzt
            self.mb.schreibe_settings_orga()
        except:
            log(inspect.stack,tb())

    
    def pruefe_auf_geladenes_organon_projekt(self):
        if self.mb.debug: log(inspect.stack)
        
        # prueft, ob eine Organon Datei geladen ist
        if len(self.mb.props[T.AB].dict_bereiche) == 0:
            return False
        else:
            Popup(self.mb, 'warning').text = LANG.PRUEFE_AUF_GELADENES_ORGANON_PROJEKT
            return True
        
        
      
    def erzeuge_Projekt_xml_tree(self):
        if self.mb.debug: log(inspect.stack)
        
        et = ElementTree  
        root = et.Element('ORGANON')
        tree = et.ElementTree(root)
        self.mb.props[T.AB].xml_tree = tree
        root.attrib['Name'] = 'root'
        root.attrib['kommender_Eintrag'] = self.mb.props[T.AB].kommender_Eintrag
        # Version fuer eventuelle Kompabilitaetspruefung speichern
        # wird nur an dieser Stelle verwendet
        root.attrib['Programmversion'] = self.mb.programm_version
           

                           
    def erzeuge_Eintraege_und_Bereiche(self,Eintraege):
        if self.mb.debug: log(inspect.stack)
        
        try:
            props = self.mb.props[T.AB]
            
            CB = self.mb.class_Bereiche
            CB.leere_Dokument()    ################################  rausnehmen
            CB.starte_oOO()
            
            Bereichsname_dict = {}
            ordinal_dict = {}
            Bereichsname_ord_dict = {}
            index = 0
            index2 = 0 
            
            
            if self.mb.settings_proj['tag3']:
                tree = props.xml_tree
                gliederung = self.mb.class_Gliederung.rechne(tree)
            else:
                gliederung = None
            
            
            for eintrag in Eintraege:
                # Navigation
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                         
                index, ctrl = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,gliederung,index)
                self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)  
    
                if sicht == 'ja':
                    # index wird in erzeuge_Zeile_in_der_Baumansicht bereits erhoeht, daher hier 1 abziehen
                    pos_Y = (index-1)*KONST.ZEILENHOEHE
                    props.dict_zeilen_posY.update({ pos_Y : eintrag })
                    self.mb.sichtbare_bereiche.append('OrganonSec' + str(index2) )
                    props.dict_posY_ctrl.update({ pos_Y : ctrl })
                    
                # Bereiche   
                inhalt = name
                path = CB.erzeuge_neue_Datei(index2,inhalt)
                path2 = uno.systemPathToFileUrl(path)
                
                if art == 'waste':
                    CB.erzeuge_bereich_papierkorb(index2,path2) 
                else:
                    CB.erzeuge_bereich(index2,path2,sicht) 
    
                Bereichsname_dict.update({'OrganonSec'+str(index2):path})
                ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
                Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
                
                index2 += 1
            
            props.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
            props.dict_bereiche.update({'ordinal':ordinal_dict})
            props.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
            
            self.erzeuge_helfer_bereich()
            CB.loesche_leeren_Textbereich_am_Ende()  
            CB.schliesse_oOO()
            self.erzeuge_dict_ordner()
        
        except:
            log(inspect.stack,tb())

    def erzeuge_helfer_bereich(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")       
            newSection.setName('Organon_Sec_Helfer')
            self.mb.sec_helfer = newSection
            
            text = self.mb.doc.Text
            textSectionCursor = text.createTextCursor()
            
            # Helfer Section 1
            textSectionCursor.gotoEnd(False)
            text.insertTextContent(textSectionCursor, newSection, False)
            newSection.IsVisible = False

        except:
            log(inspect.stack,tb())

    def erzeuge_Eintraege_und_Bereiche2(self,Eintraege):
        if self.mb.debug: log(inspect.stack)

        CB = self.mb.class_Bereiche
        props = self.mb.props[T.AB]
        
        self.erzeuge_dict_ordner()
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0 
        
        first_time = True
        
        
        if self.mb.settings_proj['tag3']:
            tree = props.xml_tree
            gliederung = self.mb.class_Gliederung.rechne(tree)
        else:
            gliederung = None
            
        
        for eintrag in Eintraege:
            # Navigation
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag   
                     
            index, ctrl = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,gliederung,index)
            
            if sicht == 'ja':
                pos_Y = (index-1) * KONST.ZEILENHOEHE
                props.dict_zeilen_posY.update({ pos_Y : eintrag})
                props.dict_posY_ctrl.update({ pos_Y : ctrl })
                self.mb.sichtbare_bereiche.append( 'OrganonSec' + str(index2) )
                
            # Bereiche   
            path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)
            path2 = uno.systemPathToFileUrl(path)
            # Der Papierkorb muss mit einem File verlinkt werden, damit die Bereiche richtig eingefuegt werden koennen
            if art == 'waste':
                CB.erzeuge_bereich_papierkorb(index2,path2) 
            else:
                CB.erzeuge_bereich(index2,path2,'nein') 

            if first_time:       
                # Viewcursor an den Anfang setzen, damit 
                # der Eindruck eines schnell geladenen Dokuments entsteht   
                self.mb.viewcursor.gotoStart(False)
                first_time = False
            
            Bereichsname_dict.update({ 'OrganonSec' + str(index2) : path })
            ordinal_dict.update({ ordinal : 'OrganonSec' + str(index2) })
            Bereichsname_ord_dict.update({ 'OrganonSec' + str(index2) : ordinal })
            
            index2 += 1

        props.dict_bereiche.update({ 'Bereichsname' : Bereichsname_dict})
        props.dict_bereiche.update({ 'ordinal' : ordinal_dict})
        props.dict_bereiche.update({ 'Bereichsname-ordinal' : Bereichsname_ord_dict})
        
        self.erzeuge_helfer_bereich()
        CB.loesche_leeren_Textbereich_am_Ende() 
           
                   
    def erzeuge_dict_ordner(self, tab_name=None):
        if self.mb.debug: log(inspect.stack)

        if tab_name:
            TAB = tab_name
        else:
            TAB = T.AB
            
        tree = self.mb.props[TAB].xml_tree
        root = tree.getroot()
        
        self.mb.props[TAB].dict_ordner = {}
        
        alle_eintraege = root.findall('.//')
        ordner = []
        
        ### Vielleicht gibt es eine Moeglichkeit, den Baum nur einmal zu durchlaufen?
        ### Statt 1) Baum komplett durchlaufen 2) jeden Eintrag nochmals rekursiv durchlaufen
        ### Ziel: einmal durchlaufen und jedes Kind bei allen Elternelementen eintragen
        
        # Liste aller Ordner erstellen
        for eintrag in alle_eintraege:
            if eintrag.attrib['Art'] in ('dir','waste','prj'):
                ordner.append(eintrag.tag)
        
        
        def get_tree_info(elem, odict,tag,helfer):
            helfer.append(elem.tag)
            # hier wird self.mb.props[T.AB].dict_ordner geschrieben
            odict[tag] = helfer
            if elem.attrib['Zustand'] == 'auf':
                for child in list(elem):
                    get_tree_info(child, odict,tag,helfer)

        # Fuer alle Ordner eine Liste ihrer Kinder erstellen -> self.mb.props[T.AB].dict_ordner       
        for tag in sorted(ordner):
            odir = root.find('.//'+tag)
            helfer = []
            get_tree_info(odir,self.mb.props[TAB].dict_ordner,tag,helfer)
        
        
    def lese_xml_datei(self):
        if self.mb.debug: log(inspect.stack)

        pfad = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')      
        self.mb.props[T.AB].xml_tree = ElementTree.parse(pfad)
        root = self.mb.props[T.AB].xml_tree.getroot()

        self.mb.props[T.AB].kommender_Eintrag = int(root.attrib['kommender_Eintrag'])
        
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
   
           
    def erzeuge_proj_Settings(self):
        if self.mb.debug: log(inspect.stack)
        
        if self.mb.language == 'de':
                datum_format = ['dd','mm','yyyy']
        else:
            datum_format = ['mm','dd','yyyy']
                
        settings_proj = {
            'tag1' : 0, 
            'tag2' : 0,
            'tag3' : 0,
            'use_template' : self.mb.settings_proj['use_template'],
            'use_template_organon' : self.mb.settings_proj['use_template_organon'],
            'formatierung' : 'Standard',
            'datum_trenner' : '.',
            'datum_format' : datum_format,
            }
            
        self.mb.speicher_settings("project_settings.txt", settings_proj)    
        self.mb.settings_proj = settings_proj
        
    
    def erzeuge_export_Settings(self):
        if self.mb.debug: log(inspect.stack)
        
        settings_exp = {
            # Export Dialog
            'alles' : 0, 
            'sichtbar' : 0,
            'eigene_ausw' : 1,
            
            'einz_dok' : 1,
            'trenner' : 1,
            'einz_dat' : 0,
            'ordner_strukt' : 0,
            'typ' : 'writer8',
            'speicherort' : uno.systemPathToFileUrl(self.mb.pfade['projekt']),
            'neues_proj' : 0,
            
            # Trenner
            'ordnertitel': 1,
            'format_ord': 1,
            'style_ord': 'Heading',
            'dateititel': 1,
            'format_dat': 1,
            'style_dat': 'Heading',
            'dok_einfuegen': 0,
            'url': '',
            'leerzeilen_drunter': 1,
            'anz_drunter' : 2,
            'seitenumbruch_ord' : 1,
            'seitenumbruch_dat' : 0,
            
            # Auswahl
            'auswahl' : 1,
            'ausgewaehlte' : {},
            
            # HTML Export
            'html_export' : {
                            'FETT' : 1,
                            'KURSIV' : 1,
                            'UEBERSCHRIFT' : 1,
                            'FUSSNOTE' : 1,
                            'FARBEN' : 1,
                            'PARA' : 1,
                            'AUSRICHTUNG' : 1,
                            'LINKS' : 1,
                            'ZITATE' : 0,
                            'SCHRIFTGROESSE' : 0,
                            'CSS' : 0
                            },
            }  
        
        self.mb.speicher_settings("export_settings.txt", settings_exp)        
        self.mb.settings_exp = settings_exp
    
    
    def erzeuge_import_Settings(self):
        if self.mb.debug: log(inspect.stack)
        
        Settings = {'imp_dat' : '1',
                'ord_strukt' : '1',
                
                'url_dat' : '',
                'url_ord' : '',
                
                #Filter
                'odt' : '1',
                'doc' : '0',
                'docx' : '0',
                'rtf' : '0',
                'txt' : '0',
                'auswahl' : '0',
                'filterauswahl' : {},
                }  
        
        self.mb.speicher_settings("import_settings.txt", Settings)
        self.mb.settings_imp = Settings
      
    def erzeuge_plain_txt(self):
        if self.mb.debug: log(inspect.stack)
        
        
        def auslesen(fkt_schreiben,mb,uno,codecs_open,pd):
                
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True

            texte = {}
            
            for name,pfad in mb.props['ORGANON'].dict_bereiche['Bereichsname'].items():

                try:
                    pfad2 = uno.systemPathToFileUrl(pfad)
                    doc = mb.doc.CurrentController.Frame.loadComponentFromURL(pfad2,'_blank',0,(prop,))
                    txt = doc.Text.String
                    texte.update({ name: txt})
                    doc.close(False)
                except Exception as e:
                    doc.close(False)
            
            fkt_schreiben(texte,codecs_open)
                        
            
        def schreiben(t,codecs_open):
            for name,txt in t.items():
                
                path = 'C:\\Users\\Homer\\Desktop\\Neuer Ordner\\txt test\\' + name + '.txt'
                
                with codecs_open(path , "w","utf-8") as f:
                    f.write(txt)
            
            print('fertig')
        
        
        try:
            from threading import Thread  

            t = Thread(target=auslesen,args=(schreiben,self.mb,uno,codecs_open,pd))
            t.start()
                
        except Exception as e:
            log(inspect.stack,tb())
        
        
        
        
        
          
    def get_flags(self,x):
        if self.mb.debug: log(inspect.stack)
              
        x_bin_rev = bin(x).split('0b')[1][::-1]

        flags = []
        
        for i in range(len(x_bin_rev)):
            z = 2**int(i)*int(x_bin_rev[i])
        
            if z != 0:
                flags.append(z)
        
        return flags
    
    def mit_template_oeffnen(self,templ_pfad,ist_vorhandenes_prj = False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # damit das Template geoeffnet werden kann, muss das Dokument unter einem
            # anderen Namen gespeichert werden
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name+'wird_geloescht'))
            Path2 = uno.systemPathToFileUrl(Path1)

            self.mb.doc.storeAsURL(Path2,()) 
            
            if ist_vorhandenes_prj:
                url = templ_pfad
            else:
                url = self.mb.settings_proj['use_template'][1]
            url = uno.systemPathToFileUrl(url)
 
            self.mb.Listener.entferne_alle_Listener()            
            
            prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop2.Name = 'AsTemplate'
            prop2.Value = True
            
            prop3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop3.Name = 'DocumentTitle'
            prop3.Value = 'opened by Organon;'+Path2.replace('wird_geloescht','')
                        
            doc = self.mb.desktop.ActiveFrame.loadComponentFromURL(url,'_top','',(prop2,prop3))
            
            self.mb.doc.close(False)
            os.remove(Path1)

        except:
            log(inspect.stack,tb())
        

    def find_titel(self,panellist,tag,lang):
        if self.mb.debug: log(inspect.stack)

        try:
            tag_xml = panellist.find(".//*[@{{http://openoffice.org/2001/registry}}name='{0}']".format(tag))
            title_xml = tag_xml.find(".//*[@{http://openoffice.org/2001/registry}name='Title']")
            name = title_xml.find(".//*[@{{http://www.w3.org/XML/1998/namespace}}lang='{0}']".format(lang))
            return name.text
        except:
            return None
    
       
    def test(self):

        try:
            
            pass

            
            props = self.mb.props[T.AB]
            doc = self.mb.doc
            vc = self.mb.viewcursor
   
            #desc = self.mb.createUnoService("com.sun.star.comp.framework.UICommandDescription")    
            #cmi = ContextMenuInterceptor(self.mb)
            #contr.registerContextMenuInterceptor(cmi)
                        
            
            vc = self.mb.viewcursor
            


            ordinal = self.mb.class_Bereiche.get_ordinal(vc)
            self.mb.class_Bereiche.datei_speichern(ordinal)
            self.mb.class_Tools.zeige_content_xml(ordinal)  
            
            import imp, menu_bar
            imp.reload(menu_bar)
            
            #for i in range(1,185,10):
            
            text = 31 * ' werter herr gesangs verein # ' 
            #text = LANG.ORGANIZER_INFO
            

            attribs = dir(LANG)
            
            texte = [getattr(LANG, a) for a in attribs][33:51]
            
#             for a in attribs:
#                 txt = getattr(LANG, a)
#                 if len(txt.split())>0:
#                     menu_bar.Popup(self.mb, 'error').text = txt
#                     break
            
            #p.text = 'Hi!'
            #p.text = 'so' 
            
            
            
        except:
            log(inspect.stack,tb())
            pd()
        pd() 
        


    
    



#             from com.sun.star.beans import StringPair
#             sp = StringPair()
#             sp.First = "content.xml"
#             sp.Second =  "id1720227130"
#             el = doc.getElementByMetadataReference(sp)
#             
#             from com.sun.star.rdf import XURI
#             
#             rdf = doc.RDFRepository
#             gr = rdf.GraphNames[0]
#             
#             xuri = self.mb.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.sun.star.rdf.URI",
#                             ("http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile",
#                              ),self.mb.ctx)
# 
# 
#             
#             
#             from com.sun.star.beans import PropertyValue
#             prop = PropertyValue()
#             
#             prop.Name = u"text/xml#"
#             prop.Value = u'content.xml'
#             graph = doc.RDFRepository.getGraph(xuri)
#             
#             footnotes = [doc.Footnotes.getByIndex(i) for i in range(doc.Footnotes.Count)]
#             
#             nested = vc.NestedTextContent
#             statements = graph.getStatements(None,None,None)
#             
#             inhalt = []
#             while statements.hasMoreElements():
#                 inhalt.append(statements.nextElement())
#                 
#                 
#             xuri2 = self.mb.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.sun.star.rdf.URI",
#                             ('file:///C:/Users/Homer/Documents/organon%20projekte/unsinn.organon/Files/odt/template.ott/',
#                              'content.xml'),self.mb.ctx)
#                 
#             graph2 = doc.RDFRepository.getGraph(xuri2)
            
#             smgr = self.mb.ctx.ServiceManager
#             inh = u"\n".join(sorted(smgr.AvailableServiceNames) )
#             pfad_plain_txt = u'C:\\Users\\Homer\\Desktop\\Neuer Ordner\\odict.txt'
#             with codecs_open(pfad_plain_txt , "w","utf-8") as f:
#                 f.write(inh)
            #uri = self.mb.createUnoService('com.sun.star.rdf.URI')
            
            #value_uri = uri.createKnown(com.sun.star.rdf.URIs.RDF_VALUE)



# from com.sun.star.ui import XContextMenuInterceptor
# from com.sun.star.ui.ContextMenuInterceptorAction import (
#     IGNORED, CONTINUE_MODIFIED, EXECUTE_MODIFIED)
# class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
#     
#     INSERT_SHEET = "slot:26269"
#     
#     def __init__(self, mb):
#         self.mb = mb
#     
#     def notifyContextMenuExecute(self, ev):
#         pd()
#         cont = ev.ActionTriggerContainer
# #         if cont.getByIndex(0).CommandURL == self.INSERT_SHEET:
# #             #item = cont.createInstance("com.sun.star.ui.ActionTriggerSeparator")
# #             #item.SeparatorType = LINE
# #             #cont.insertByIndex(8, item)
# #             items = ActionTriggerContainer()
# # #             items.insertByIndex(0, ActionTriggerItem(
# # #                 ".uno:SelectTables", "Sheet...", 
# # #                 "", None))
# # #             
# # #             item = ActionTriggerItem("GoTo", "Go to", "123", items)
# # #             cont.insertByIndex(7, item)
# #             
# #             return EXECUTE_MODIFIED
#         return IGNORED
# 
# 




from com.sun.star.awt import XActionListener
class neues_Projekt_Dialog_Listener(unohelper.Base,XActionListener): 
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctrls = None
        self.fenster = None
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:            
            cmd = ev.ActionCommand  
            
            if cmd == 'organon':
                self.toggle_checkbox(cmd)
                
            elif cmd == 'writer':
                self.toggle_checkbox(cmd)

            elif cmd == LANG.WAEHLEN:
                self.file_aussuchen()

            elif cmd == LANG.OK:
                self.dialog_neues_projekt_auswerten()
                
        except:
            log(inspect.stack,tb())
            
    
    def toggle_checkbox(self,cmd):
        if self.mb.debug: log(inspect.stack)
        
        if cmd == 'writer':
            self.ctrls['organon_cb'].setState(False)
        else:
            self.ctrls['writer_cb'].setState(False)
            
    
    def file_aussuchen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            filepath = None
            pfad = os.path.join(self.mb.path_to_extension, "pfade.txt")
            
            if os.path.exists(pfad):            
                with codecs_open( pfad, "r","utf-8") as f:
                    filepath = f.read() 
    
            filepath = self.mb.class_Funktionen.folderpicker(filepath=filepath, sys=True)
            
            if filepath:
                with codecs_open( pfad, "w","utf-8") as f:
                    f.write(filepath)
                    
                self.mb.speicherort_last_proj = filepath
                self.ctrls['speicherort'].Model.Label = filepath
            
        except:
            log(inspect.stack,tb())
            
            
    def dialog_neues_projekt_auswerten(self):
        if self.mb.debug: log(inspect.stack)
        
        namen_pruefen = self.mb.class_Funktionen.verbotene_buchstaben_austauschen
        
        prj_name = self.ctrls['prj_name'].Model.Text
        speicherort = self.ctrls['speicherort'].Model.Label
        
        # Projektname und Speicherort ueberpruefen
        if prj_name == '':
            Popup(self.mb, 'warning').text = LANG.KEIN_NAME
            return
        
        elif prj_name != namen_pruefen(prj_name):
            Popup(self.mb, 'warning').text = LANG.UNGUELTIGE_ZEICHEN
            return
             
        elif speicherort == None:
            Popup(self.mb, 'warning').text = LANG.KEIN_SPEICHERORT
            return
        
        self.mb.projekt_path = speicherort
        
        # Templates 
        templ_pfad, templ_art = None, None
        is_template = self.ctrls['organon_cb'].State or self.ctrls['writer_cb'].State
        
        if is_template:
            templ_art = 'organon' if self.ctrls['organon_cb'].State else 'writer'
            
            listbox = self.ctrls['{}_lb'.format(templ_art)]
            sel = listbox.SelectedItem
            
            if templ_art == 'writer':
                templ_pfad = self.mb.writer_vorlagen[sel]
            else:
                dir_pfad = self.mb.settings_orga['templates_organon']['pfad']
                templ_pfad = os.path.join(dir_pfad,sel + '.organon')
                
        # Pfade setzen       
        self.mb.projekt_name = prj_name
        self.mb.class_Projekt.setze_pfade()
        
        # Wahrscheinlich ueberfluessig und nur zur Sicherheit:
        # Pruefen ob neues Projekt mit gleichem Namen im Vorlagenordner
        # erstellt werden soll
        if is_template and templ_art == 'organon':
            if templ_pfad == self.mb.pfade['projekt']:
                Popup(self.mb, 'warning').text = LANG.GLEICHER_PFAD
                return
        
        # Pruefen ob aktuelles Dokument den gleichen Namen wie das neue besitzt
        if self.mb.projekt_name == self.mb.doc.Title.split('.odt')[0]:
            Popup(self.mb, 'warning').text = LANG.DOUBLE_PROJ_NAME
            return
        
        # Wenn das Projekt schon existiert, Abfrage, ob Projekt ueberschrieben werden soll
        elif os.path.exists(self.mb.pfade['projekt']):
            # 16777216 Flag fuer YES_NO
            entscheidung = self.mb.entscheidung(LANG.PROJ_EXISTS,"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return
            elif entscheidung == 2:
                try:
                    import shutil
                    # entfernt das vorhandene Projekt
                    shutil.rmtree(self.mb.pfade['projekt'])
                except:
                    pass
                
        if is_template and templ_art == 'organon':
                    
            try:
                self.fenster.dispose()
                self.mb.class_Funktionen.projekt_umbenannt_speichern(templ_pfad,self.mb.pfade['projekt'],self.mb.projekt_name)
                new_path = os.path.join(self.mb.pfade['projekt'],self.mb.projekt_name + '.organon')
                self.mb.class_Projekt.lade_Projekt(filepicker = False, filepath = new_path)
                return
            except Exception as e:
                log(inspect.stack,tb())  
                Popup(self.mb, 'error').text = LANG.TEMPLATE_NICHT_GELADEN.format(str(e))
                return
        
        self.fenster.dispose()
        self.mb.class_Projekt.erzeuge_neues_Projekt(is_template, templ_art, templ_pfad)
  
    def disposing(self,ev):
        return False
 

        
   
        
