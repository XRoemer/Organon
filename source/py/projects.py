# -*- coding: utf-8 -*-

import unohelper

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
        
        self.first_time = True
   
    def erzeuge_neues_Projekt(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            if self.pruefe_auf_geladenes_organon_projekt():
                return
            
            geglueckt,self.mb.projekt_name = self.dialog_neues_projekt_anlegen()  
            
            if geglueckt:
                self.setze_pfade()

                if self.mb.projekt_name == self.mb.doc.Title.split('.odt')[0]:
                    
                    self.mb.nachricht(LANG.DOUBLE_PROJ_NAME,"warningbox")
                    return
                
                # Wenn das Projekt schon existiert, Abfrage, ob Projekt ueberschrieben werden soll
                # funktioniert das unter Linux?? ############
                elif os.path.exists(self.mb.pfade['projekt']):
                    # 16777216 Flag fuer YES_NO
                    entscheidung = self.mb.nachricht(LANG.PROJ_EXISTS,"warningbox",16777216)
                    # 3 = Nein oder Cancel, 2 = Ja
                    if entscheidung == 3:
                        return
                    elif entscheidung == 2:
                        try:
                            import shutil
                            # entfernt das vorhandene Projekt
                            shutil.rmtree(self.mb.pfade['projekt'])
                        except:
                            # scheint trotz Fehlermeldung zu funktionieren win7 OO/LO
                            pass #log(inspect.stack,tb())
                  
            if geglueckt:
                
                ok = self.erzeuge_Ordner_Struktur() 
                if not ok:
                    return
                self.erzeuge_import_Settings()
                self.erzeuge_export_Settings()  
                self.erzeuge_proj_Settings()
                          
                self.mb.class_Bereiche.leere_Dokument()        
                self.mb.class_Baumansicht.start()             
                Eintraege = self.beispieleintraege2()
                
                self.erzeuge_Projekt_xml_tree() 
                self.mb.class_Bereiche.erzeuge_leere_datei()               
                self.erzeuge_Eintraege_und_Bereiche(Eintraege)
                
                Path1 = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
                self.mb.tree_write(self.mb.props[T.AB].xml_tree,Path1)
                
                self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
                
                self.mb.props[T.AB].selektierte_zeile = self.mb.props[T.AB].Hauptfeld.getByIdentifier(0).AccessibleContext.AccessibleName
                self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_ersten_Bereichs()
                #self.mb.doc.addDocumentEventListener(self.mb.doc_listener)
                
                self.mb.class_Sidebar.lege_dict_sb_content_an()
                self.mb.class_Sidebar.lade_sidebar()
                self.mb.class_Sidebar.speicher_sidebar_dict()
                
                self.mb.class_Baumansicht.korrigiere_scrollbar()
                
                self.mb.use_UM_Listener = True
                
                filepath = os.path.join(self.mb.pfade['projekt'],"%s.organon" % self.mb.projekt_name)
                dateiname = "%s.organon" % self.mb.projekt_name
                self.trage_projekt_in_zuletzt_geladene_Projekte_ein(dateiname,filepath)
                
        except Exception as e:
            self.mb.nachricht('erzeuge_neues_Projekt ' + str(e),"warningbox")
            log(inspect.stack,tb())
            pd()

                        
    def setze_pfade(self): 
        if self.mb.debug: log(inspect.stack)
        
        paths = self.mb.smgr.createInstance( "com.sun.star.util.PathSettings" )
        pHome = paths.Work_writable
        if sys.platform == 'linux':
            os.chdir( '//')        

        pOrganon = self.mb.projekt_path

        pProjekt =  os.path.join(pOrganon , '%s.organon' % self.mb.projekt_name)
        pFiles =    os.path.join(pProjekt , 'Files')
        pOdts =     os.path.join(pFiles , 'odt')
        pImages =   os.path.join(pFiles , 'Images')
        pIcons =   os.path.join(pFiles , 'Icons')
        pSettings = os.path.join(pProjekt , 'Settings')
        pTabs =     os.path.join(pSettings , 'Tabs')
        
        
        self.mb.pfade.update({'home':pHome}) 
        self.mb.pfade.update({'projekt':pProjekt})      
        self.mb.pfade.update({'organon':pOrganon})
        self.mb.pfade.update({'files':pFiles})
        self.mb.pfade.update({'odts':pOdts})
        self.mb.pfade.update({'settings':pSettings}) 
        self.mb.pfade.update({'images':pImages}) 
        self.mb.pfade.update({'tabs':pTabs}) 
        self.mb.pfade.update({'icons':pIcons}) 
        
        # Pfad fuer Fehler Debugging des Projektes setzen
        self.mb.class_Log.path_to_project_settings = pSettings

    
    def lade_settings(self):
        if self.mb.debug: log(inspect.stack)
        
        pfad = os.path.join(self.mb.pfade['settings'],'export_settings.txt')
        self.mb.settings_exp = eval(open(pfad).read())

        pfad = os.path.join(self.mb.pfade['settings'],'import_settings.txt')
        self.mb.settings_imp = eval(open(pfad).read())
        
        pfad = os.path.join(self.mb.pfade['settings'],'project_settings.txt')
        self.mb.settings_proj = eval(open(pfad).read())
        

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
    
            # Datei anlegen, die bei lade_Projekt angesprochen werden soll
            path = os.path.join(pfade['projekt'],"%s.organon" % self.mb.projekt_name)
            with open(path, "w") as file:
                file.write('Dies ist eine Organon Datei. Goennen Sie ihr ihre Existenz.') 
                 
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
             
#         except PermissionError as e:
#             self.mb.nachricht("You don't have the permission to write into this folder " + str(e),"warningbox")
#             return False
        except Exception as e:
            self.mb.nachricht("ERROR: " + str(e),"warningbox")
            return False
        
        return True
        
    def dialog_neues_projekt_anlegen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:

            if self.mb.speicherort_last_proj != None:
                # try/except fuer Ubuntu: U meldet Fehler: couldn't convert fileUrlTo ...
                # -> gespeicherten Pfad ueberpruefen!
                try:
                    modelU_Label = uno.fileUrlToSystemPath(self.mb.speicherort_last_proj)
                except:
                    #modelU.Label = '-' 
                    modelU_Label = self.mb.speicherort_last_proj
            else:
                modelU_Label = '-' 
            

            modelForm1_Label = LANG.TEMPLATE_WRITER
            if not self.mb.settings_proj:
                state = 0#False
                not_state = 1
            else:
                if self.mb.settings_proj['use_template'][0]:
                    state = 1
                    not_state = 0
                else:
                    state = 0
                    not_state = 1
                    
            modelForm1_State = not_state
            modelForm2_State = state
            
            if self.mb.user_styles == ():
                user_styles = (LANG.NO_TEMPLATES,)
                controlLBF2_Enable = 0
                controlForm2_Enable = 0
            else:
                user_styles = self.mb.user_styles
                controlLBF2_Enable = 1

            self.mb.user_styles,pfade = self.get_user_styles()
            templates = ('Minimal','Standard','Maximum')
            
            # LISTENER
            listenerS = Speicherordner_Button_Listener(self.mb)
            listener = neues_Projekt_Dialog_Listener(self.mb) 
            listenerCB = Neues_Projekt_CheckBox_Listener(self.mb)
            listener_info = Neues_Projekt_InfoButton_Listener(self.mb)

            y = 0
            
            tab0 = tab0x = 25
            tab1 = tab1x = 115
            tab2 = tab2x = 0
            tab3 = tab3x = 80
            tab4 = tab4x = 130
            
            tabs = [tab0,tab1,tab2,tab3,tab4]
            
            design = self.mb.class_Design
            design.set_default(tabs)

                
            controls = (
            20,
            ('control',"FixedText",         'tab0x',y,250,20,  ('Label','FontWeight'),(LANG.ENTER_PROJ_NAME ,150),          {'addKeyListener':(listener)} ),
            30,
            ('control1',"Edit",             'tab0',y,200,20,   (),(),                                                       {} ) ,
            30,
            ('controlT3',"FixedLine",       'tab0',y,360,40,   (),(),                                                       {} ), 
            43,
            ('controlP',"FixedText",        'tab0x',y,120,20,  ('Label','FontWeight'),(LANG.SPEICHERORT,150),               {} ),  
            0,
            ('controlW',"Button",           'tab2x',y,80,20,   ('Label',),(LANG.AUSWAHL,),                                  {'setActionCommand':LANG.WAEHLEN,'addActionListener':(listenerS,listener)}),              
            30,
            ('controlU',"FixedText",        'tab0',y,300,20,   ('Label','State'),(modelU_Label,modelForm1_State),           {} ), 
            30,
            ('controlT',"FixedLine",        'tab0',y,360,40,   (),(),                                                       {} ), 
            40,
#             ('controlForm',"FixedText",     'tab0x',y,80,20,   ('Label','FontWeight'),(LANG.FORMATIERUNG,150),              {} ),
#             0,  
#             ('controlForm1',"CheckBox",     'tab1',y,200,20,   ('Label','State'),(LANG.TEMPLATE_WRITER,modelForm1_State),   {'setActionCommand':'standard','addActionListener':(listenerCB,)} ),
#             0,
#             ('controlHelp',"Button",        'tab4',y ,30,30,   ('ImageURL',),('vnd.sun.star.extension://xaver.roemers.organon/img/info_16.png',),{'setActionCommand':'formatierung','addActionListener':(listener_info,)} ), 
#             25,
#             ('controlForm2',"CheckBox",     'tab1',y,200,20,   ('Label','State'),(LANG.TEMPLATE_USER,modelForm2_State),     {'setActionCommand':'user','addActionListener':(listenerCB,)} ) ,
#             22,
#             ('controlLBF2',"ListBox",       'tab2',y,80,20,    ('Dropdown',),(True,),                                       {'Enable':controlLBF2_Enable,'addItems':user_styles,'SelectedItems':0,'addItemListener':(listenerCB)}),
#             20,
#             ('controlT5',"FixedLine",       'tab2',y,238,40,   (),(),                                                       {} ),
#             50,
#             ('controlFormLBF4',"FixedText", 'tab2',y,300,20,   ('Label',),(LANG.EIGENES_TEMPL_ERSTELLEN,),                  {} ), 
#             25,
#             ('controlLBF5',"FixedText",     'tab2',y,50,20,    ('Label',),(LANG.NAME,),                                     {}),  
#             0,
#             ('controlLBF6',"Edit",          'tab3',y,130,20,   (),(),                                                       {} ),
#             25,
#             ('controlER',"Button",          'tab2x',y ,80,20,  ('Label',),(LANG.ERSTELLEN,),                                {'setActionCommand':'vorlage_erstellen','addActionListener':(listener,)} ),  
#             30,
#             ('controlT2',"FixedLine",       'tab0',y,360,40,   (),(),                                                       {} ), 
#             40,
#             ('controlTemp',"FixedText",     'tab0x',y,80,20,   ('Label',),(LANG.TEMPLATE,),                                 {}) ,
#             -3,
#             ('controlTempL',"ListBox",      'tab2',y,80,20,    ('Dropdown',),(True,),                                       {'addItems':templates,'Enable':False,'SelectedItems':0}) ,
#             -10,
#             ('controlHelpT',"Button",       'tab4',y,30,30,    ('ImageURL',),('vnd.sun.star.extension://xaver.roemers.organon/img/info_16.png',),{'setActionCommand':'formatierung','addActionListener':(listener_info,)}  ),
#             30,
#             ('controlT4',"FixedLine",       'tab0',y,360,40,   (),(),                                                       {} ), 
#             40,
            ('control2',"Button",           'tab3',y,80,30,    ('Label',),(LANG.OK,),                                       {'setActionCommand':LANG.OK,'addActionListener':(listener,)} ) , 
            0,
#             ('control3',"Button",           tab2+120,y,80,30,('Label',),(LANG.CANCEL,),                                  {'setActionCommand':LANG.CANCEL,'addActionListener':(listener,)} )  
            )
            
            
            # HAUPTFENSTER    
            # create the dialog model and set the properties
            dialogModel = self.mb.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
            dialogModel.Width = 215 
            dialogModel.Height = 320
            dialogModel.Title = LANG.CREATE_NEW_PROJECT
                      
            # create the dialog control and set the model
            controlContainer = self.mb.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx);
            controlContainer.setModel(dialogModel);
                     
            # create a peer
            toolkit = self.mb.smgr.createInstanceWithContext("com.sun.star.awt.ExtToolkit", self.ctx);       
            controlContainer.setVisible(False);       
            controlContainer.createPeer(toolkit, None);
            # ENDE HAUPTFENSTER
            
            
            pos_y = 0
            
            for ctrl in controls:
                if isinstance(ctrl,int):
                    pos_y += ctrl
                else:
                    name,unoCtrl,X,Y,width,height,prop_names,prop_values,extras = ctrl
                    pos_x = locals()[X]
                    
                    locals()[name],locals()[name.replace('control','model')] = self.mb.createControl(self.ctx,unoCtrl,pos_x,pos_y,width,height,prop_names,prop_values)
                            
                    if 'x' in X:
                        w,h = self.mb.kalkuliere_und_setze_Control(locals()[name],'w')
                        design.setze_tab(X,w)
                    
                    if 'setActionCommand' in extras:
                        locals()[name].setActionCommand(extras['setActionCommand'])
                    if 'addItems' in extras:
                        locals()[name].addItems(extras['addItems'],0)
                    if 'Enable' in extras:
                        locals()[name].Enable = extras['Enable']
                    if 'addActionListener' in extras:
                        for l in extras['addActionListener']:
                            locals()[name].addActionListener(l)
                    if 'addKeyListener' in extras:
                        locals()[name].addKeyListener(extras['addKeyListener'])
                    if 'addItemListener' in extras:
                        locals()[name].addItemListener(extras['addItemListener'])

                    controlContainer.addControl(name,locals()[name])
            
            # Tabs x-Position neu berechnen
            design.kalkuliere_tabs()

            for i in range(len(tabs)):
                locals()['tab%sx'%i] = design.new_tabs['tab%sx'%(i)]                
            
            for ctrl in controls:
                if isinstance(ctrl,int):
                    pass
                else:
                    name,unoCtrl,X,Y,width,height,prop_names,prop_values,extras = ctrl                    
                    pos_x = design.new_tabs[X]
                    
                    # Sonderregeln
                    if X == 'tab1':
                        pos_x -= 15
                    
                    locals()[name].setPosSize(pos_x,0,0,0,1)
            
            controlContainer.setPosSize(0,0,400,pos_y + 50,12)

            # UEBERGABE AN LISTENER
            listenerS.control = locals()['controlU']
            listener.model_proj_name = locals()['model1']
            #listener.model_neue_vorl = locals()['modelLBF6']
            #listener.control_CB = locals()['controlForm2']
            #listener.control_LB = locals()['controlLBF2']
            #listenerCB.modelStandard = locals()['modelForm1']
            #listenerCB.modelUser = locals()['modelForm2']
            #listenerCB.modelListBox = locals()['modelLBF2']


            controlContainer.addTopWindowListener(listener)
        
            geglueckt = controlContainer.execute()       
            controlContainer.dispose() 
    
            return geglueckt,locals()['model1'].Text
        
        except:
            log(inspect.stack,tb())

                       
   
    
    def get_user_styles(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            paths = self.mb.smgr.createInstance( "com.sun.star.util.PathSettings" )
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
    
            benutzervorlagen = []
            pfade = []
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
                        
                            benutzervorlagen.append(dateiname1)
                            pfade.append(pfad)
                        
            self.mb.user_styles_pfade = tuple(pfade)
            return tuple(benutzervorlagen),tuple(pfade)  
        except:
            log(inspect.stack,tb())


    def lade_Projekt(self,filepicker = True, filepath = ''):
        if self.mb.debug: log(inspect.stack)

        try:
            if self.pruefe_auf_geladenes_organon_projekt():
                return
    
            if filepicker:
                Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
                Filepicker.appendFilter('Organon Project','*.organon')
                #ofilter = Filepicker.getCurrentFilter()
                Filepicker.execute()
                # see: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Basic/File_Control
    
                if Filepicker.Files == '':
                    return
    
                filepath =  uno.fileUrlToSystemPath(Filepicker.Files[0])
                
            dateiname = os.path.basename(filepath)
            dateiendung = os.path.splitext(filepath)[1]
    
            # Wenn keine .organon Datei gewaehlt wurde
            if dateiendung  != '.organon':
                return
            
            self.mb.projekt_name = dateiname.split(dateiendung)[0]
            self.mb.projekt_path = os.path.dirname(os.path.dirname(filepath))  

            self.setze_pfade()
            self.mb.class_Bereiche.leere_Dokument() 
            self.lade_settings()  
            Eintraege = self.lese_xml_datei()

            self.mb.class_Version.pruefe_version()

            self.mb.props[T.AB].Hauptfeld = self.mb.class_Baumansicht.erzeuge_Feld_Baumansicht(self.mb.dialog) 
            self.erzeuge_Eintraege_und_Bereiche2(Eintraege) 
            
            # setzt die selektierte Zeile auf die erste Zeile
            self.mb.props[T.AB].selektierte_zeile = self.mb.props[T.AB].Hauptfeld.getByIdentifier(0).AccessibleContext.AccessibleName
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_des_ersten_Bereichs()
            
            self.mb.class_Baumansicht.erzeuge_Scrollbar()    
            self.mb.class_Baumansicht.korrigiere_scrollbar()
            self.mb.class_Mausrad.registriere_Maus_Focus_Listener(self.mb.props['Projekt'].Hauptfeld.Context.Context)

            # Wenn die UDProp verloren gegangen sein sollte, wieder setzen
            UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
            has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
            if not has_prop:
                UD_properties.addProperty('ProjektName',1,self.mb.projekt_name) 
            
            # damit von den Bereichen in die Datei verlinkt wird, muss sie gespeichert werden   
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name))
            Path2 = uno.systemPathToFileUrl(Path1)
            self.mb.doc.storeAsURL(Path2,()) 
            
            self.mb.class_Tabs.lade_tabs()
            
            self.mb.class_Sidebar.lade_sidebar_dict()
            self.mb.class_Sidebar.lade_sidebar()
            self.selektiere_ersten_Bereich()
            self.mb.use_UM_Listener = True    
            self.trage_projekt_in_zuletzt_geladene_Projekte_ein(dateiname,filepath)
            
        except Exception as e:
            self.mb.nachricht('lade_Projekt ' + str(e),"warningbox")
            log(inspect.stack,tb())
    
    
    
    def trage_projekt_in_zuletzt_geladene_Projekte_ein(self,dateiname,filepath):
        if self.mb.debug: log(inspect.stack)
        
        zuletzt = self.mb.settings_orga['zuletzt_geladene_Projekte']
        
        name = dateiname.split('.organon')[0]

        if name not in zuletzt:
            zuletzt.update({name:filepath})
            
        # Fehlt: Wenn der Dateiname an einem anderen Ort nochmal existiert, wir er nicht in die Liste aufgenommen
        self.mb.class_Funktionen.schreibe_settings_orga()

        
    def pruefe_auf_geladenes_organon_projekt(self):
        if self.mb.debug: log(inspect.stack)
        
        # prueft, ob eine Organon Datei geladen ist
        UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
        has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
        #self.mb.entferne_alle_listener() 
        if len(self.mb.props[T.AB].dict_bereiche) == 0:
            return False
        else:
#         if has_prop:
            self.mb.nachricht(LANG.PRUEFE_AUF_GELADENES_ORGANON_PROJEKT,"warningbox")
            return True
        
        
      
    def erzeuge_Projekt_xml_tree(self):
        if self.mb.debug: log(inspect.stack)
        
        et = self.mb.ET    
        #prj_name = self.mb.projekt_name.replace(' ','_')   
        root = et.Element('Projekt')
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
            CB = self.mb.class_Bereiche
            CB.leere_Dokument()    ################################  rausnehmen
            CB.starte_oOO()
            
            Bereichsname_dict = {}
            ordinal_dict = {}
            Bereichsname_ord_dict = {}
            index = 0
            index2 = 0 
            
            if self.mb.settings_proj['tag3']:
                tree = self.mb.props[T.AB].xml_tree
                root = tree.getroot()
                gliederung = self.mb.class_Gliederung.rechne(tree)
            else:
                gliederung = None
            
            
            for eintrag in Eintraege:
                # Navigation
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                         
                index = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,self.mb.class_Zeilen_Listener,gliederung,index)
                self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)  
    
                if sicht == 'ja':
                    # index wird in erzeuge_Zeile_in_der_Baumansicht bereits erhoeht, daher hier 1 abziehen
                    self.mb.props[T.AB].dict_zeilen_posY.update({(index-1)*KONST.ZEILENHOEHE:eintrag})
                    self.mb.props['Projekt'].sichtbare_bereiche.append('OrganonSec'+str(index2))
                    
                # Bereiche   
                inhalt = name
                path = CB.erzeuge_neue_Datei(index2,inhalt)
                path2 = uno.systemPathToFileUrl(path)
                
                if art == 'waste':
                    CB.erzeuge_bereich(index2,path2,sicht,True) 
                else:
                    CB.erzeuge_bereich(index2,path2,sicht) 
    
                Bereichsname_dict.update({'OrganonSec'+str(index2):path})
                ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
                Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
                
                index2 += 1
            
            self.mb.props[T.AB].dict_bereiche.update({'Bereichsname':Bereichsname_dict})
            self.mb.props[T.AB].dict_bereiche.update({'ordinal':ordinal_dict})
            self.mb.props[T.AB].dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
            
            self.erzeuge_helfer_bereich()
            CB.loesche_leeren_Textbereich_am_Ende()  
            CB.schliesse_oOO()
            self.erzeuge_dict_ordner()
        
        except:
            log(inspect.stack,tb())

    def erzeuge_helfer_bereich(self):
        if self.mb.debug: log(inspect.stack)
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")       
        newSection.setName('Organon_Sec_Helfer')
        
        newSection2 = self.mb.doc.createInstance("com.sun.star.text.TextSection")       
        newSection2.setName('Organon_Sec_Helfer2')
        
        text = self.mb.doc.Text
        textSectionCursor = text.createTextCursor()
        
        # Helfer Section 1
        textSectionCursor.gotoEnd(False)
        text.insertTextContent(textSectionCursor, newSection, False)
        newSection.IsVisible = False
        
        # Helfer Section 2
        textSectionCursor.gotoStart(False)
        text.insertTextContent(textSectionCursor, newSection2, False)
        newSection2.IsVisible = False
        
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = os.path.join(self.mb.pfade['odts'],'empty_file.odt')
        newSection2.setPropertyValue('FileLink',SFLink)
        
        
        self.mb.sec_helfer = newSection
        self.mb.sec_helfer2 = newSection2        
        

    def erzeuge_Eintraege_und_Bereiche2(self,Eintraege):
        if self.mb.debug: log(inspect.stack)

        CB = self.mb.class_Bereiche
        
        self.erzeuge_dict_ordner()
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0 
        
        first_time = True
        
        
        if self.mb.settings_proj['tag3']:
            tree = self.mb.props[T.AB].xml_tree
            root = tree.getroot()
            gliederung = self.mb.class_Gliederung.rechne(tree)
        else:
            gliederung = None
            
        
        for eintrag in Eintraege:
            # Navigation
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag   
                     
            index = self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,self.mb.class_Zeilen_Listener,gliederung,index)
            
            if sicht == 'ja':
                # index wird in erzeuge_Zeile_in_der_Baumansicht bereits erhoeht, daher hier 1 abziehen
                self.mb.props[T.AB].dict_zeilen_posY.update({(index-1)*KONST.ZEILENHOEHE:eintrag})
                self.mb.props['Projekt'].sichtbare_bereiche.append('OrganonSec'+str(index2))
                
            # Bereiche   
            path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)
            path2 = uno.systemPathToFileUrl(path)
            # Der Papierkorb muss mit einem File verlinkt werden, damit die Bereiche richtig eingefuegt werden koennen
            if art == 'waste':
                CB.erzeuge_bereich(index2,path2,sicht,True) 
            else:
                CB.erzeuge_bereich(index2,path2,sicht) 

            if first_time:       
                # Viewcursor an den Anfang setzen, damit 
                # der Eindruck eines schnell geladenen Dokuments entsteht   
                self.mb.viewcursor.gotoStart(False)
                first_time = False
            
            Bereichsname_dict.update({'OrganonSec'+str(index2):path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
            
            index2 += 1

        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.props[T.AB].dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.props[T.AB].dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
        self.erzeuge_helfer_bereich()
        CB.loesche_leeren_Textbereich_am_Ende() 
           
                   
    def erzeuge_dict_ordner(self):
        if self.mb.debug: log(inspect.stack)

        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        
        ordner = []
        self.mb.props[T.AB].dict_ordner = {}
        
        alle_eintraege = root.findall('.//')
        
        
        ### Vielleicht gibt es eine Moeglichkeit, den Baum nur einmal zu durchlaufen?
        ### Statt 1) Baum komplett durchlaufen 2) jeden Eintrag nochmals rekursiv durchlaufen
        ### Ziel: einmal durchlaufen und jedes Kind bei allen Elternelementen eintragen
        
        # Liste aller Ordner erstellen
        for eintrag in alle_eintraege:
            if eintrag.attrib['Art'] in ('dir','waste','prj'):
                ordner.append(eintrag.tag)
        
        
        def get_tree_info(elem, dict,tag,helfer):
            helfer.append(elem.tag)
            # hier wird self.mb.props[T.AB].dict_ordner geschrieben
            dict[tag] = helfer
            if elem.attrib['Zustand'] == 'auf' or elem.attrib['Art'] == 'waste':
                for child in list(elem):
                    get_tree_info(child, dict,tag,helfer)

        
        # Fuer alle Ordner eine Liste ihrer Kinder erstellen -> self.mb.props[T.AB].dict_ordner       
        for tag in sorted(ordner):
            dir = root.find('.//'+tag)
            helfer = []
            get_tree_info(dir,self.mb.props[T.AB].dict_ordner,tag,helfer)
        


    def lese_xml_datei(self):
        if self.mb.debug: log(inspect.stack)

        pfad = os.path.join(self.mb.pfade['settings'], 'ElementTree.xml')      
        self.mb.props[T.AB].xml_tree = self.mb.ET.parse(pfad)
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
   
    def selektiere_ersten_Bereich(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinal = self.mb.props[T.AB].dict_bereiche['Bereichsname-ordinal']['OrganonSec0']
        zeile = self.mb.props[T.AB].Hauptfeld.getControl(ordinal)
        textfeld = zeile.getControl('textfeld')
        self.mb.props[T.AB].selektierte_zeile_alt = textfeld.Context.AccessibleContext.AccessibleName
        
    def erzeuge_proj_Settings(self):
        if self.mb.debug: log(inspect.stack)
        
        settings_proj = {
            'tag1' : 0, 
            'tag2' : 0,
            'tag3' : 0,
            'use_template' : self.mb.settings_proj['use_template'],
            'user_styles' : (),
            'formatierung' : 'Standard',
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
            # .encode("utf-8") <- wird das in der folgenden Zeile gebraucht?
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
            'hoehe_auswahlfenster' : 200,
            
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
      
    
    
       
    def beispieleintraege(self):
        
        Eintraege = [('nr0','root','Vorbemerkung',0,'pg','-','ja'),
                ('nr1','root','Projekt',0,'dir','auf','ja'),
                ('nr2','nr1','Titelseite',1,'pg','-','ja'),
                ('nr3','nr1','Kapitel1',1,'dir','auf','ja'),
                ('nr4','nr3','Szene1',2,'pg','-','ja'),
                ('nr5','nr3','Szene2',2,'pg','-','ja'),
                ('nr6','nr1','Kapitel2',1,'dir','auf','ja'),
                ('nr7','nr6','Szene1b',2,'pg','-','ja'),
                ('nr8','nr6','Szene2b',2,'pg','-','ja'),
                ('nr9','nr1','Interlude',1,'pg','-','ja'),
                ('nr10','nr1','Kapitel3',1,'dir','auf','ja'),
                ('nr11','nr1','Kapitel4',1,'dir','auf','ja'),
                ('nr12','nr11','UnterKapitel',2,'dir','zu','ja'),
                ('nr13','nr12','Szene3a',3,'pg','-','nein'),
                ('nr14','nr12','Szene3b',3,'pg','-','nein'),
                ('nr15','nr11','Szene3c',2,'pg','-','ja'),
                ('nr16','nr11','Szene3d',2,'pg','-','ja'),
                ('nr17','nr11','Kapitel4a',2,'dir','auf','ja'),
                ('nr18','nr1','Kapitel4b',1,'dir','auf','ja'),
                ('nr19','nr18','Szene4',2,'pg','-','ja'),
                ('nr20','root','Papierkorb',0,'waste','zu','ja')]
        
        return Eintraege

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
        
        
    def get_flags(self,x):
        if self.mb.debug: log(inspect.stack)
              
        x_bin_rev = bin(x).split('0b')[1][::-1]

        flags = []
        
        for i in range(len(x_bin_rev)):
            z = 2**int(i)*int(x_bin_rev[i])
        
            if z != 0:
                flags.append(z)
        
        return flags
    
       
    def test(self):

        try:
            
            pass

            

            
            
            ## PILLOW ##
            
#             pass
#             pyPath = 'C:/Program Files (x86)/OpenOffice 4/program/python-core-2.7.6/lib/site-packages/Pillow-2.7.0-py2.7-win32.egg'
#             
#             if pyPath not in sys.path:
#                 sys.path.append(pyPath)
#             import PIL
#             from PIL import Image,ImageDraw,ImageFont
            
#             ft_size = 56
#             
#             image = Image.new("RGBA", (800,200), (255,255,255))
# 
#             usr_font = ImageFont.truetype('C:\\Program Files\\Scribus 1.4.5\\share\\fonts\\Lato2\\Lato-Light.ttf', ft_size) 
#             msg = u'Beispieltext für Schrifttype:'
#             msg2 = 'Lato-Light.ttf' + ' {}pt'.format(ft_size)
#             draw = ImageDraw.Draw(image)
#      
#             draw.text( (5,0) , msg, (0,0,0), font=usr_font)
#             w, h = draw.textsize(msg,font=usr_font)
#              
#             draw.text( (5,h) , msg2, (0,0,0), font=usr_font)
#             w2, h2 = draw.textsize(msg2,font=usr_font)
#              
#             wmax = max(w,w2)
#             hmax = h+h2+5
#      
#             image.size = (wmax+10,hmax)
#             print(wmax+10,hmax)
#             pfad = os.path.join(self.mb.pfade['projekt'],'mein bild.png')
#             image.save(pfad)
#             
#             
#             faktor = ((wmax+10)/hmax)/2
#             print(faktor)
#             maxsize = (wmax*faktor, hmax*faktor)
#             #maxsize = (100,100)
#             #image.thumbnail(maxsize, PIL.Image.ANTIALIAS)
#             #im = image.resize((image.size[0] / 3, image.size[1] / 2))
#             
#             #print(image.size,wmax/2+5, hmax/2)
#             pfad = os.path.join(self.mb.pfade['projekt'],'mein bild2.png')
#             #im.save(pfad)
#             
#              
#             basewidth = 300
#             wpercent = (basewidth/float(image.size[0]))
#             hsize = int((float(image.size[1])*float(wpercent)))
#             im = image.resize((basewidth,hsize), PIL.Image.ANTIALIAS)
#             im.save(pfad)
            
            
            
            
            
            
            
            
#             # StringIO
#             output = StringIO.StringIO()
#             image.save(output, format="GIF")
#             contents = output.getvalue()
#             output.close()
             
#             photo = PhotoImage(data=contents)
#             self.beispieltext = Label(self, image=photo,width =W)
#             self.beispieltext.grid(column=6, row=30)
#             self.beispieltext.image = photo
            
#             pass
#             pyPath = self.mb.pfade['projekt']
#              
#             if pyPath not in sys.path:
#                 sys.path.append(pyPath)
# #                 
# #                 
#             pyPath = os.path.join(self.mb.pfade['projekt'],'PythonMagick')
#             if pyPath not in sys.path:
#                 sys.path.append(pyPath)
#                 
#             cur_version = sys.version_info
#             #print(cur_version)    
#             
#             #import cairo
#             #import pango
#             
#             import PythonMagick
#             
# #             path = self.mb.pfade['projekt']
# #             draw(path)
#             print('klar')



############################     Text als Bild einfuegen




#             cont1,model1 = self.mb.createControl(self.mb.ctx,'ImageControl',10,10,191,14,('ImageURL',),(os.path.join(self.mb.pfade['projekt'],'kopp.jpg'),))
#             path = uno.systemPathToFileUrl('C:\\Python276\\testbild.png')
#             model1.ImageURL = path
#             model1.Border = 0
#             #model1.setPropertyValue('ImageURL',os.path.join(self.mb.pfade['projekt'],'kopp.jpg'))
#             
#             
#             
#             #model1.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/papierkorb_leer.png'
#             
#             cont2,model2 = self.mb.createControl(self.mb.ctx, "FixedText", 10, 30, 300,30, (), ())
#             cont2.Text = 'hier ist ein Quantum AV Beispieltext!'
# 
#             
#             cont3,model3 = self.mb.createControl(self.mb.ctx,'ImageControl',10,50,215,18,('ImageURL',),(os.path.join(self.mb.pfade['projekt'],'kopp.jpg'),))
#             path = uno.systemPathToFileUrl('C:\\Python276\\testbild2.png')
#             model3.ImageURL = path
#             #model3.Border = 0
#             
#             cont4,model4 = self.mb.createControl(self.mb.ctx,'ImageControl',10,70,192,12,('ImageURL',),(os.path.join(self.mb.pfade['projekt'],'kopp.jpg'),))
#             path = uno.systemPathToFileUrl('C:\\Python276\\testbild3.png')
#             model4.ImageURL = path
#             model4.Border = 0
#             
#             
# 
#             posSize = 200,200,600,300
# 
#             win,cont = self.mb.erzeuge_Dialog_Container(posSize)
#             
#            
#             
#             
#             
#             cont.addControl('',cont1)
#             cont.addControl('',cont2)
#             #cont.addControl('',cont3)
#             cont.addControl('',cont4)
            
#############################################################################################


#             attributes = inspect.getmembers(win.StyleSettings, lambda a:not(inspect.isroutine(a)))
#              
#             for a in attributes:
#                 try:
#                     setattr(win.StyleSettings, a[0], 502)
#                 except:
#                     pass
            
#             cmp_win = self.mb.desktop.ActiveFrame.ComponentWindow
#             win = cmp_win.Windows[0]
#             
#             width,height = win.Size.Width,win.Size.Height
#             loc = win.AccessibleContext.LocationOnScreen
#             
#             loc_cont = self.mb.current_Contr.Frame.ContainerWindow.AccessibleContext.LocationOnScreen
#             cont = self.mb.current_Contr.Frame.ContainerWindow
#                     
# #             x = self.mb.dialog.AccessibleContext.LocationOnScreen.X - loc_cont.X + self.mb.dialog.PosSize.Width
# #             y = self.mb.dialog.AccessibleContext.LocationOnScreen.Y - loc_cont.Y 
#     
#             from com.sun.star.util.MeasureUnit import POINT
#             
#             
#             x,y = loc.X,loc.Y
#             x,y = win.PosSize.X,62 + 31
#             
#             #x,y = new_point.X,new_point.Y
#             frame = self.mb.desktop.ActiveFrame
#             lmgr = frame.LayoutManager
#             
#             
#             
#             y = cmp_win.AccessibleContext.AccessibleParent.PosSize.Y + cmp_win.AccessibleContext.Location.Y
#             
#             cw = lmgr.DockingAreaAcceptor.ContainerWindow
#             
# #             point = cw.AccessibleContext.LocationOnScreen
# #             y = cont.convertPointToPixel(point,POINT).Y
#             y = cmp_win.PosSize.Y + 31
#             posSize = x,y,width,height
#             
#             #win2,cont = self.mb.erzeuge_Dialog_Container(posSize,1+32+64+128)#Flags=1+512)#1+32+64+128)
#             
#             
#             
#             cont_win = frame.ContainerWindow
#             
#             for w in cont_win.Windows:
#                 p = w.PosSize
#                 #print(p.X,p.Y,p.Width,p.Height)
#             
#             lmgr.DockingAreaAcceptor.ContainerWindow.Background = 0
#             
#             for e in lmgr.Elements:
#                 print(e.ResourceURL)
#                 try:
#                     p = e.RealInterface.PosSize
#                     
#                     print(p.X,p.Y,p.Width,p.Height)
#                 except:
#                     print('Fehler')
                    
            #pd()     

            #pd()
            
            
#             
            #self.mb.btn.State = 1
            #self.org.erzeuge_steuerung()
#             ctx = self.mb.ctx
#             smgr = self.mb.ctx.ServiceManager
#                
#             config_provider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider",ctx)
#       
#             prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
#             prop.Name = "nodepath"
#             prop.Value = "/org.openoffice.Office.Views/Windows/"
#                    
#             config_access = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (prop,))
                 
            #config_access.Persona = nutze_persona
            #config_access.PersonaSettings = persona_url
                    
            #config_access.commitChanges()
            #decklist = config_access.Content.DeckList
            #decklist.replaceByName('GalleryDeck',decklist.OrganonsToolDeck)
            
            
#             from organizer import Organizer
#             self.org = Organizer(self.mb,pd,T,LANG,KONST)
#             self.org.run()
            
            
            
            
            
            
            
#             #self.myDialog()
#             RESOURCE_URL = "private:resource/dockingwindow/9809"
#             #self.Org.calc_frame.LayoutManager.showElement(RESOURCE_URL)
#             #self.mb.desktop.ActiveFrame.LayoutManager.showElement(RESOURCE_URL)
# #             self.mb.dialog.endDialog(0)
# #             self.mb.dialog.endExecute()
#             
#             
#             self.mb.win.invalidate(1)
#             #self.mb.win.dispose()
#             win = self.mb.win
#             #win.setVisible(False)
#             #self.mb.tabsX.Window.dispose()
#             tabs = self.mb.tabsX
#             #tabs.removeTab(0)
#             win.enableClipSiblings(True)
#             win.enableDialogControl(True)
#             
#             attributes = inspect.getmembers(win.StyleSettings, lambda a:not(inspect.isroutine(a)))
#             attributes2 = inspect.getmembers(win.StyleSettings)
#             
#             ctx = uno.getComponentContext()
#             smgr = ctx.ServiceManager
#             desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
#             doc = desktop.getCurrentComponent() 
#             current_Contr = doc.CurrentController
#             viewcursor = current_Contr.ViewCursor
#             
#             object = doc
#             max_lvl = 3
#             
#             def get_attribs(obj,lvl):
#                 results = {}
#                 for key in dir(obj):
#                     
#                     try:
#                         value = getattr(obj, key)
#                         if 'callable' in str(type(value)):
#                             continue
#                     except :
#                         #print(key)
#                         continue
# 
#                     if key not in results:
#                         if type(value) in (
#                                            type(None),
#                                            type(True),
#                                            type(1),
#                                            type(.1),
#                                            type('string'),
#                                            type(()),
#                                            type([]),
#                                            type(b''),
#                                            type(r''),
#                                            type(u'')
#                                            ):
#                             results.update({key: value})
#                             
#                         elif lvl < max_lvl:
#                             try:
#                                 results.update({key: get_attribs(value,lvl+1)})
#                             except:
#                                 pass
#                 
#                 return results
#             
#             
#             diff = []
#             
#             def findDiff(d1, d2, path=""):
#                 for k in d1.keys():
#                     if not d2.has_key(k):
#                         print (path, ":")
#                         print (k + " as key not in d2", "\n")
#                     else:
#                         if type(d1[k]) is dict:
#                             if path == "":
#                                 path = k
#                             else:
#                                 path = path + "->" + k
#                             findDiff(d1[k],d2[k], path)
#                         else:
#                             if d1[k] != d2[k]:
#                                 diff.append((path,k,d1[k],d2[k]))
#                                 path = ''
#             
#             
#             self.res1 = get_attribs(object,1)
#             self.mb.viewcursor.gotoEnd(False)
#             self.mb.viewcursor.setString('Test ')
#             self.res2 = get_attribs(object,1)
#             
#             findDiff(self.res1, self.res2)
#             
#             for d in diff:
#                 print(d)
#                 time.sleep(.4)
            
#             for x in range(lvl-1):
#                 
#                 res = get_attribs(object)
            
            
            dic = self.mb.dict_sb_content




            
        except:
            log(inspect.stack,tb())
            print(tb())
            pd()
        pd()
    
    
    
    
    

            
    

    def myDialog(self):
        psm = uno.getComponentContext().ServiceManager
        dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    
        dlg = dp.createDialog("vnd.sun.star.extension://xaver.roemers.organon//factory/Dialog1.xdl")     
        dlg.Title = "Mitteilung"
        
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 0, 0, 400, 40, (), ())           
        model.Label = 'Mitteilung'
        model.MultiLine = True
        
        
        
        dlg.addControl('mitteilung',control)
        
        pref_size = control.PreferredSize
        
        
        
        #dlg.getControl("TextField1").Text = " Here you can read your message "
        dlg.execute()
        #pd()
#         control.setVisible(True)
#         dlg.setVisible(True)
#         time.sleep(3) ## 5 sec
        dlg.dispose()   
 


def spiral(x, y):
    theta0 = -math.pi * 3 / 4
    theta = x / Width * math.pi * 2 + theta0
    radius = y + 200 - x/7
    xnew = radius*math.cos(theta)
    ynew = radius*math.sin(-theta)
    return xnew + Width/2, ynew + Height/2

def curl(x, y):
    xn = x - Textwidth/2
    #yn = y - Textheight/2
    xnew = xn
    ynew = y + xn ** 3 / ((Textwidth/2)**3) * 70
    return xnew + Width/2, ynew + Height*2/5

def draw(path):
    import cairo
    import math
    
    global Width, Height, Textwidth
    Width, Height = 512, 512
    surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, Width, Height)
    ctx = cairo.Context(surface)
    solidpattern = ctx.get_source()
    
    # background
    pat = cairo.LinearGradient (1, 1, 1, Height)
    #pat.add_color_stop_rgba (1, 0, 0, 0, 1)
    #pat.add_color_stop_rgba (0, 1, 1, 1, 1)
    
    ctx.rectangle (0,0,Width,Height)
    ctx.set_source (pat)
    ctx.fill ()
    
    # foreground
    ctx.set_source (solidpattern)
    ctx.set_source_rgb (1,1,1)
    
    ctx.select_font_face("Sans")
    ctx.set_font_size(12)
    
    dir_ctx = dir(ctx)
    
    # curly text
    ctx.new_path()
    ctx.move_to(20, 80)
    ctx.set_source_rgb(0, 0, 0)
    text = u"für vérenö\nNiemand hört auf mich!\Aber daß das so ist\wird sich ändern!"
    ctx.text_path(text)
    Textwidth, Textheight = ctx.text_extents(text)[2:4]
    #warpPath(ctx, curl)
    ctx.fill()
    import os
    os.chdir(path)
    
    surface.write_to_png('cairo_test.png')
    
            


from com.sun.star.awt import XActionListener,XTopWindowListener,XKeyListener,XItemListener
class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model = None
        self.control = None
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        try:
            filepath = None
            pfad = os.path.join(self.mb.path_to_extension,"pfade.txt")
            
            if os.path.exists(pfad):            
                with codecs_open( pfad, "r","utf-8") as file:
                    filepath = file.read() 
    
            Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
            if filepath != None:
                Filepicker.setDisplayDirectory(filepath)
            Filepicker.execute()
            
            if Filepicker.Directory == '':
                return
            
            filepath = Filepicker.getDirectory()
            
            with codecs_open( pfad, "w","utf-8") as file:
                file.write(uno.fileUrlToSystemPath(filepath))
            
            self.mb.speicherort_last_proj = filepath
            self.control.Model.Label = uno.fileUrlToSystemPath(filepath)
            self.mb.kalkuliere_und_setze_Control(self.control)
            
            if self.mb.debug: log(inspect.stack,None,filepath)
            
        except:
            if 'filepath' in locals():
                if self.mb.debug: log(inspect.stack,tb(),filepath)
            else:
                log(inspect.stack,tb())
        
        
        
    def disposing(self,ev):
        return False



from com.sun.star.awt.Key import RETURN
class neues_Projekt_Dialog_Listener(unohelper.Base,XActionListener,XTopWindowListener,XKeyListener): 
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model_proj_name = None
        self.model_neue_vorl = None
        self.control_CB = None
        self.control_LB = None
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            namen_pruefen = self.mb.class_Funktionen.verbotene_buchstaben_austauschen
            
            parent = ev.Source.AccessibleContext.AccessibleParent 
            cmd = ev.ActionCommand  
            if cmd == 'vorlage_erstellen':
                self.vorlage_auswaehlen()
            elif cmd == LANG.WAEHLEN:
                return
            elif cmd == LANG.CANCEL:
                parent.endDialog(0)
            elif self.model_proj_name.Text == '':
                self.mb.nachricht(LANG.KEIN_NAME,"warningbox")
            elif self.mb.speicherort_last_proj == None:
                self.mb.nachricht(LANG.KEIN_SPEICHERORT,"warningbox")
            elif self.model_proj_name.Text != namen_pruefen(self.model_proj_name.Text):
                self.mb.nachricht(LANG.UNGUELTIGE_ZEICHEN,"warningbox")
            elif cmd == LANG.OK:
                self.get_path()
                parent.endDialog(1)
                
        except:
            log(inspect.stack,tb())
        
            
    def keyPressed(self,ev):
        if ev.KeyCode == RETURN:
            parent = ev.Source.AccessibleContext.AccessibleParent 
            if self.model.Text == '':
                parent.endDialog(0)
            else:
                self.get_path()
                parent.endDialog(1)

    def get_path(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            filepath = None
            pfad = os.path.join(self.mb.path_to_extension,"pfade.txt")
            
            if os.path.exists(pfad):            
                with codecs_open( pfad, "r","utf-8") as file:
                    filepath = file.read() 
                self.mb.projekt_path = filepath
        except:
            log(inspect.stack,tb())
            
            
    def vorlage_auswaehlen(self):        
        if self.mb.debug: log(inspect.stack)
        
        if self.model_neue_vorl.Text == '':
            self.mb.nachricht(LANG.WARNUNG_NAME_TEMPLATE,"warningbox")
            return
        
        else:
            name = self.model_neue_vorl.Text
            
            if name in self.mb.user_styles:
                self.mb.nachricht(LANG.WARNUNG_TEMPLATE_EXISTS   ,"warningbox")
                return
            
            else:
                if self.mb.settings_proj['use_template'][0] == False:
                    URL = "private:factory/swriter"
                else:
                    URL = uno.systemPathToFileUrl(self.mb.settings_proj['use_template'][1])
                    
                self.vorlage_erstellen(name,URL)
                self.control_LB.Enable = True
                self.control_CB.Enable = True
                
                if self.control_LB.Items[0] == LANG.NO_TEMPLATES:
                    
                    self.control_LB.Model.removeAllItems()
                    self.control_LB.addItems((name,),0)
                    self.control_LB.Model.SelectedItems = 0,
                    
                    self.mb.user_styles = (name,)
                    
                else:
                    self.control_LB.addItem(name,0)
                    self.control_LB.Model.SelectedItems = 0,
                    
                    st = []
                    for style in self.mb.user_styles:
                        st.append(style)
                    st.append(name)         
                    
                    self.mb.user_styles = tuple(st)       


        
    def vorlage_erstellen(self,name,URL):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True                            
            
            prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop2.Name = 'FilterName'
            prop2.Value = 'writer8_template'
            
            #URL2 = "private:factory/swriter"
            newDoc = self.mb.desktop.loadComponentFromURL(URL,'_blank',8+32,(prop,))
            #newDoc.DocumentProperties.TemplateURL = URL
            newDoc.DocumentProperties.Title = name
            
            p1 = uno.fileUrlToSystemPath(self.mb.user_template_path)  
            p2 = os.path.join(p1,name+'.ott')          
            Path2 = uno.systemPathToFileUrl(p2)
            newDoc.storeToURL(Path2,(prop2,))
            
        except:
            log(inspect.stack,tb())
        newDoc.close(False)
        self.mb.nachricht(LANG.NEUES_TEMPLATE + '\n%s   ' % p1,"infobox")
        
        # var DefaultTemplate = ((XDocumentPropertiesSupplier)oDoc).getDocumentProperties().TemplateURL;
    
    
    def keyReleased(self,ev):
        return False
    
    # XTopWindowListener
    def windowOpened(self,ev):
        return False
    def windowActivated(self,ev):
        return False
    def windowDeactivated(self,ev):
        return False
    def windowClosing(self,ev):
        return False
    def windowClosed(self,ev):
        return False
    def windowMinimized(self,ev):
        return False
    def windowNormalized(self,ev):
        return False
    def disposing(self,ev):
        return False
    
        
class Neues_Projekt_CheckBox_Listener(unohelper.Base, XActionListener,XItemListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.modelStandard = None
        self.modelUser = None
        self.modelListBox = None
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
            return
        elif ev.ActionCommand == 'standard':
            self.modelUser.State = False
            self.mb.settings_proj['use_template'] = (False,None)
        elif ev.ActionCommand == 'user':
            self.modelStandard.State = False
            pfad = self.get_template_pfad()
            self.mb.settings_proj['use_template'] = (True,pfad)
        
    
    def get_template_pfad(self):
        if self.mb.debug: log(inspect.stack)
        
        pfade = self.mb.user_styles_pfade
        gewaehlt = self.modelListBox.SelectedItems[0]
        pfad = pfade[gewaehlt]

        return pfad
        
    
    def itemStateChanged(self, ev):  
        if self.mb.debug: log(inspect.stack)
        
        use,pfad = self.mb.settings_proj['use_template']
        self.mb.settings_proj['use_template'] = (use,self.get_template_pfad())
    
    def disposing(self,ev):
        return False



class Neues_Projekt_InfoButton_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if ev.ActionCommand == 'formatierung':
                path = os.path.join(self.mb.path_to_extension,'languages','info_format_%s.odt' % self.mb.language)
                URL = uno.systemPathToFileUrl(path)
            else:
                path = os.path.join(self.mb.path_to_extension,'languages','info_template_%s.odt' % self.mb.language)
                URL = uno.systemPathToFileUrl(path)
            
            self.new_doc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,())
            
            contWin = self.new_doc.CurrentController.Frame.ContainerWindow               
            contWin.setPosSize(0,0,870,900,12)
            
            lmgr = self.new_doc.CurrentController.Frame.LayoutManager
            for elem in lmgr.Elements:
            
                if lmgr.isElementVisible(elem.ResourceURL):
                    lmgr.hideElement(elem.ResourceURL)
                    
            lmgr.HideCurrentUI = True  
            
            
            viewSettings = self.new_doc.CurrentController.ViewSettings
            viewSettings.ZoomType = 3
            viewSettings.ZoomValue = 100
            viewSettings.ShowRulers = False
            
        except:
            log(inspect.stack,tb())

        
    def disposing(self,ev):
        return False
        
   
        
