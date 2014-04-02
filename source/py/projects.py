# -*- coding: utf-8 -*-

#print('Projekte')
import traceback
import uno
import unohelper
import os
import sys
import codecs
tb = traceback.print_exc

ZEILENHOEHE = 22



class Projekt():
    
    def __init__(self,mb,pydevBrk):
        
        global pd
        pd = pydevBrk
        
        self.dialog = mb.dialog
        self.ctx = mb.ctx
        self.mb = mb
        self.pd = pydevBrk
        
        global lang
        lang = self.mb.lang
   
   
   ##########   ausgesuchten Pfad neu setzen
   
   
    def erzeuge_neues_Projekt(self):
        try:
            geglueckt,self.mb.projekt_name = self.dialog_neues_projekt_anlegen()  
            
            if geglueckt:
                self.setze_pfade()
                #return
                if self.mb.projekt_name == self.mb.doc.Title.split('.odt')[0]:
                    
                    self.mb.Mitteilungen.nachricht(lang.DOUBLE_PROJ_NAME,"warningbox")
                    return
                    # Wenn das Projekt schon existiert, Abfrage, ob Projekt ueberschrieben werden soll
                    # funktioniert das unter Linux?? ############
                elif os.path.exists(self.mb.pfade['projekt']):
                    # 16777216 Flag fuer YES_NO
                    entscheidung = self.mb.Mitteilungen.nachricht(lang.PROJ_EXISTS,"warningbox",16777216)
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
                            pass
                        
        except Exception as e:
            self.mb.Mitteilungen.nachricht(str(e),"warningbox")
            tb()
        
        if geglueckt:
            try:
                self.erzeuge_Settings()
                self.erzeuge_Ordner_Struktur() 
                self.erzeuge_import_Settings()
                self.erzeuge_export_Settings()  
                          
                self.mb.class_Bereiche.leere_Dokument()        
                self.mb.class_Hauptfeld.start()             
                Eintraege = self.beispieleintraege2()
                self.erzeuge_Projekt_xml_tree()
                self.erzeuge_Eintraege_und_Bereiche(Eintraege)
                  
                Path1 = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
                self.mb.xml_tree.write(Path1)
                Path2 = os.path.join(self.mb.pfade['settings'],'settings.xml')
                self.mb.xml_tree_settings.write(Path2)
                
            except Exception as e:
                self.mb.Mitteilungen.nachricht(str(e),"warningbox")
                tb()

                        
    def setze_pfade(self): 
  
        paths = self.mb.smgr.createInstance( "com.sun.star.util.PathSettings" )
        pHome = paths.Work_writable
        if sys.platform == 'linux':
            os.chdir( '//')

#         retval = os.getcwd()
#         print ("Current working directory %s" % retval)          

        pNavi = self.mb.projekt_path
        pOrganon = uno.fileUrlToSystemPath(pNavi)

        pProjekt =  os.path.join(pOrganon , '%s.organon' % self.mb.projekt_name)
        pFiles =    os.path.join(pProjekt , 'Files')
        pOdts =     os.path.join(pFiles , 'odt')
        pSettings = os.path.join(pProjekt , 'Settings')
        
        self.mb.pfade.update({'home':pHome}) 
        self.mb.pfade.update({'projekt':pProjekt})      
        self.mb.pfade.update({'organon':pOrganon})
        self.mb.pfade.update({'files':pFiles})
        self.mb.pfade.update({'odts':pOdts})
        self.mb.pfade.update({'settings':pSettings}) 

    
    def lade_imp_exp_settings(self):

        pyPath = self.mb.pfade['settings']
        sys.path.append(pyPath)

        import export_settings
        self.mb.exp_settings = export_settings.exp_settings[0]

        pfad = os.path.join(self.mb.pfade['settings'],'imp_settings.xml')
        self.mb.imp_settings_xml = self.mb.ET.parse((pfad))

        
    def erzeuge_Ordner_Struktur(self):
        
        
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
        # Organon/<Projekt Name>/Settings
        if not os.path.exists(pfade['settings']):
            os.makedirs(pfade['settings'])

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

                       
    def dialog_neues_projekt_anlegen(self):
        
        y = 20
        # Projektname eingeben
        control, model = createControl(self.ctx,"FixedText",25,y,250,20,(),() )  
        model.Label = lang.ENTER_PROJ_NAME
        model.FontWeight = 150
        
        y += 30
        # Eingabefeld
        control1, model1 = createControl(self.ctx,"Edit",25,y,200,20,(),() ) 
        
        
        
        
        y += 40
        
        # Trenner 
        controlT3, model3 = createControl(self.mb.ctx,"FixedLine",20,y-10 ,360,40,(),() ) 
        
        
        y += 40
        # speicherort
        controlP, modelP = createControl(self.ctx,"FixedText",25,y + 3,120,20,(),() )  
        modelP.Label = lang.SPEICHERORT
        modelP.FontWeight = 150
        
        # waehlen 
        controlW, modelW = createControl(self.ctx,"Button",142,y,80,20,(),() )  
        modelW.Label = lang.AUSWAHL 
        controlW.setActionCommand(lang.WAEHLEN)
        
        y += 30
        # url
        controlU, modelU = createControl(self.ctx,"FixedText",25,y,300,20,(),() ) 
        if self.mb.speicherort_last_proj != None:
            modelU.Label = uno.fileUrlToSystemPath(self.mb.speicherort_last_proj)
        else:
            modelU.Label = '-' 
        
        listenerS = Speicherordner_Button_Listener(self.mb,modelU)
        controlW.addActionListener(listenerS)
        
        y += 40
        
        # Trenner 
        controlT, modelT = createControl(self.mb.ctx,"FixedLine",20,y-10 ,360,40,(),() )  
        
        
        
        
        y += 40
        # Formatierung
        controlForm, modelForm = createControl(self.ctx,"FixedText",25,y,80,20,(),() )  
        modelForm.Label = lang.FORMATIERUNG #lang.ENTER_PROJ_NAME
        modelForm.FontWeight = 150
        
        
        controlForm1, modelForm1 = createControl(self.ctx,"CheckBox",125,y,200,20,(),() )  
        modelForm1.Label = lang.TEMPLATE_WRITER
        modelForm1.State = not self.mb.use_template[0]
        

        controlHelp, modelHelp = createControl(self.ctx,"Button",350,y ,30,30,(),() )  
        modelHelp.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/info_16.png'
        
        y += 25
        
        controlForm2, modelForm2 = createControl(self.ctx,"CheckBox",125,y,200,20,(),() ) 
        modelForm2.Label = lang.TEMPLATE_USER
        modelForm2.State = self.mb.use_template[0]
        
        y += 25
            # Liste der Formate
        controlLBF2, modelLBF2 = createControl(self.mb.ctx,"ListBox",142,y -3 ,80,20,(),() )  

        self.mb.user_styles,pfade = self.get_user_styles()

        if self.mb.user_styles == ():
            user_styles = ('lang.NO_TEMPLATES',)
            controlLBF2.Enable = False
            controlForm2.Enable = False
        else:
            user_styles = self.mb.user_styles
        controlLBF2.addItems(user_styles,0)
        modelLBF2.Dropdown = True
        #index = style_names.index(set['style_ord'])
        #modelLBF.SelectedItems = index,
        modelLBF2.SelectedItems = 0,
        
        
        # Trenner 
        controlT5, modelT5 = createControl(self.mb.ctx,"FixedLine",142,y+15 ,238,40,(),() )
        
        
        y += 50
        
        controlFormLBF4, modelFormLBF4 = createControl(self.ctx,"FixedText",142,y,300,20,(),() )  
        modelFormLBF4.Label = lang.EIGENES_TEMPL_ERSTELLEN

        
        y += 25
        
        controlLBF5, modelLBF5 = createControl(self.ctx,"FixedText",142,y,50,20,(),() )  
        modelLBF5.Label = lang.NAME
        
        controlLBF6, modelLBF6 = createControl(self.ctx,"Edit",222,y,158,20,(),() ) 
        
        
        y += 25
        
        # waehlen 
        controlER, modelER = createControl(self.ctx,"Button",142,y ,80,20,(),() )  
        modelER.Label = lang.ERSTELLEN
        controlER.setActionCommand('vorlage_erstellen')
        
        
        y += 40
        
        # Trenner 
        controlT2, modelT2 = createControl(self.mb.ctx,"FixedLine",20,y-10 ,360,40,(),() )  
        
        
        y += 40
        # Formatierung
        controlTemp, modelTemp = createControl(self.ctx,"FixedText",25,y,80,20,(),() )  
        modelTemp.Label = lang.TEMPLATE
        modelTemp.FontWeight = 150
        
        
        controlTempL, modelTempL = createControl(self.mb.ctx,"ListBox",142,y -3 ,80,20,(),() )  
        templates = ('Minimal','Standard','Maximum')
        controlTempL.addItems(templates,0)
        modelTempL.Dropdown = True
        #index = style_names.index(set['style_ord'])
        #modelLBF.SelectedItems = index,
        modelTempL.SelectedItems = 0,
        controlTempL.Enable = False
        
        
        controlHelpT, modelHelpT = createControl(self.ctx,"Button",350,y - 10 ,30,30,(),() )  
        modelHelpT.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/info_16.png'
        
        
        
        y += 40
        
        # Trenner 
        controlT4, modelT4 = createControl(self.mb.ctx,"FixedLine",20,y-10 ,360,40,(),() )  
        
        
        y += 40
        x = 142
        # ok button
        control2, model2 = createControl(self.ctx,"Button",x,y,80,30,(),() )  
        model2.Label = lang.OK
        control2.setActionCommand(lang.OK)
        
        # cancel button  
        control3, model3 = createControl(self.ctx,"Button",x + 120,y,80,30,(),() )  
        model3.Label = lang.CANCEL
        control3.setActionCommand(lang.CANCEL)
        
            
            
        # HAUPTFENSTER    
        smgr = self.mb.smgr
        
        # create the dialog model and set the properties
        dialogModel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
           
        dialogModel.PositionX = 65
        dialogModel.PositionY = 65
        dialogModel.Width = 215 # if lang == en:195

        dialogModel.Height = 320
        dialogModel.Title = lang.CREATE_NEW_PROJECT
                  
        # create the dialog control and set the model
        controlContainer = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx);
        controlContainer.setModel(dialogModel);
                 
        # create a peer
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.ExtToolkit", self.ctx);       
        controlContainer.setVisible(False);       
        controlContainer.createPeer(toolkit, None);
        # ENDE HAUPTFENSTER
        
        # LISTENER 1
        listener = neues_Projekt_Dialog_Listener(self.mb,model1,modelLBF6,controlForm2,controlLBF2) 
        control1.addKeyListener(listener) 
        controlW.addActionListener(listener) 
        controlER.addActionListener(listener) 
        control2.addActionListener(listener) 
        control3.addActionListener(listener) 
        
        
        
        # LISTENER 2
        listenerCB = Neues_Projekt_CheckBox_Listener(self.mb,modelForm1,modelForm2,modelLBF2)
        controlForm1.addActionListener(listenerCB)
        controlForm1.ActionCommand = 'standard'
        controlForm2.addActionListener(listenerCB)
        controlForm2.ActionCommand = 'user'
        controlLBF2.addItemListener(listenerCB)
        
        
        # CONTROLS HINZUFUEGEN
        controlContainer.addControl('text',control)
        controlContainer.addControl('name',control1)
        
        controlContainer.addControl('trenner',controlT3)
        
        controlContainer.addControl('projektordner',controlP)
        controlContainer.addControl('waehlen',controlW)
        controlContainer.addControl('url',controlU)
        
        
        controlContainer.addControl('trenner',controlT)
        
        controlContainer.addControl('Form',controlForm)
        controlContainer.addControl('Form1',controlForm1)       
#         controlContainer.addControl('FormLBF',controlLBF)
#         controlContainer.addControl('FormLBF1',controlLBF1)
        
        controlContainer.addControl('Help',controlHelp)
        
        controlContainer.addControl('Form2',controlForm2)       
        controlContainer.addControl('FormLB2',controlLBF2)
        controlContainer.addControl('FormLBF3',controlT5)
        controlContainer.addControl('FormLBF4',controlFormLBF4)
        controlContainer.addControl('FormLBF5',controlLBF5)
        controlContainer.addControl('FormLBF6',controlLBF6)
        controlContainer.addControl('FormLBF6',controlER)
        
        controlContainer.addControl('trenner',controlT2)
        
        controlContainer.addControl('template',controlTemp)        
        controlContainer.addControl('templateL',controlTempL)
        controlContainer.addControl('templateLH',controlHelpT)

        
        controlContainer.addControl('trenner',controlT4)
        
        controlContainer.addControl('button',control2)
        controlContainer.addControl('button2',control3)
        
        controlContainer.addTopWindowListener(listener)
    
        geglueckt = controlContainer.execute()       
        controlContainer.dispose() 

        return geglueckt,model1.Text
    
    
    def get_user_styles(self):
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
                    
                        benutzervorlagen.append(dateiname1)  #,pfad))
                        pfade.append(pfad)
                    
        self.mb.user_styles_pfade = tuple(pfade)
        return tuple(benutzervorlagen),tuple(pfade)

# ===============================================================================================================  


    def lade_Projekt(self,filepicker = True):
        if self.mb.debug: print(self.mb.debug_time(),'lade_Projekt')
       
        # fehlt: wenn bereits ein Projekt geladen wurde, stuerzt oOO ab
        # daher: alles entfernen !!
        
        if filepicker:
            Filepicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
            Filepicker.appendFilter('Organon Project','*.organon')
            filter = Filepicker.getCurrentFilter()
            Filepicker.execute()
            # see: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Basic/File_Control

            if Filepicker.Files == '':
                return
            
            filepath = Filepicker.Files[0]
            # Wenn keine .organon Datei gewaehlt wurde
            if filepath.split('/')[-1].split('.')[1]  != 'organon':
                return
            
            self.mb.projekt_name = filepath.split('/')[-1].split('.')[0] 
            proj = os.path.dirname(filepath) 
            self.mb.projekt_path = os.path.dirname(proj)   
            

#         # prueft, ob eine Organon Datei geladen ist
#         UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
#         has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
# 
#         if has_prop:
#             dialog_contr = self.mb.dialog.Controls
#             for contr in dialog_contr:
#                 contr.dispose()


        try:
            self.setze_pfade()
            self.mb.class_Bereiche.leere_Dokument() 
            self.lese_xml_settings_datei() 
            self.lade_imp_exp_settings()      
            self.mb.class_Hauptfeld.erzeuge_Navigations_Hauptfeld() 
            self.mb.class_Hauptfeld.erzeuge_Scrollbar(self.mb.dialog,self.mb.ctx)       
            Eintraege = self.lese_xml_datei()
            self.erzeuge_Eintraege_und_Bereiche2(Eintraege) 
            
            # setzt die selektierte Zeile auf die erste Zeile
            self.mb.selektierte_zeile = self.mb.Hauptfeld.getByIdentifier(0).AccessibleContext
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche()
            
    
            # Wenn die UDProp verloren gegangen sein sollte, wieder setzen
            UD_properties = self.mb.doc.DocumentProperties.UserDefinedProperties
            has_prop = UD_properties.PropertySetInfo.hasPropertyByName('ProjektName')
            if not has_prop:
                UD_properties.addProperty('ProjektName',1,self.mb.projekt_name) 
            
            # damit von den Bereichen in die Datei verlinkt wird, muss sie gespeichert werden   
            Path1 = (os.path.join(self.mb.pfade['odts'],'%s.odt' % self.mb.projekt_name))
            Path2 = uno.systemPathToFileUrl(Path1)
            self.mb.doc.storeAsURL(Path2,()) 
            
        except Exception as e:
            self.mb.Mitteilungen.nachricht(str(e),"warningbox")
            tb()
      
    def erzeuge_Projekt_xml_tree(self):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Projekt_xml_tree')
        
        et = self.mb.ET       
        root = et.Element(self.mb.projekt_name)
        tree = et.ElementTree(root)
        self.mb.xml_tree = tree
        root.attrib['Name'] = 'root'
           
                       
    def erzeuge_Eintraege_und_Bereiche(self,Eintraege):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Eintraege_und_Bereiche')
        CB = self.mb.class_Bereiche
        CB.leere_Dokument()    ################################  rausnehmen
        CB.starte_oOO()
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0 
        for eintrag in Eintraege:
            # Navigation
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag2 = eintrag   
                     
            index = self.mb.class_Hauptfeld.erzeuge_Verzeichniseintrag(eintrag,self.mb.class_Zeilen_Listener,index)
            self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)  

            if sicht == 'ja':
                # index wird in erzeuge_Verzeichniseintrag bereits erhoeht, daher hier 1 abziehen
                self.mb.dict_zeilen_posY.update({(index-1)*ZEILENHOEHE:eintrag})
                self.mb.sichtbare_bereiche.append('OrganonSec'+str(index2))
                
            # Bereiche   
            inhalt = name
            path = CB.erzeuge_neue_Datei(index2,inhalt)
            CB.erzeuge_bereich(index2,path,sicht)            

            Bereichsname_dict.update({'OrganonSec'+str(index2):path})
            ordinal_dict.update({ordinal:'OrganonSec'+str(index2)})
            Bereichsname_ord_dict.update({'OrganonSec'+str(index2):ordinal})
            
            index2 += 1
        
        self.mb.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})

        CB.loesche_leeren_Textbereich_am_Ende()  
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener) 
        CB.schliesse_oOO()
        self.erzeuge_dict_ordner()


    def erzeuge_Eintraege_und_Bereiche2(self,Eintraege):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Eintraege_und_Bereiche2')

        CB = self.mb.class_Bereiche
        CB.leere_Dokument()    ################################  rausnehmen
        
        self.erzeuge_dict_ordner()
        
        Bereichsname_dict = {}
        ordinal_dict = {}
        Bereichsname_ord_dict = {}
        index = 0
        index2 = 0 
        
        first_time = True
        
        for eintrag in Eintraege:
            # Navigation
            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag   
                     
            index = self.mb.class_Hauptfeld.erzeuge_Verzeichniseintrag(eintrag,self.mb.class_Zeilen_Listener,index)
            
            if sicht == 'ja':
                # index wird in erzeuge_Verzeichniseintrag bereits erhoeht, daher hier 1 abziehen
                self.mb.dict_zeilen_posY.update({(index-1)*ZEILENHOEHE:eintrag})
                self.mb.sichtbare_bereiche.append('OrganonSec'+str(index2))
                
            # Bereiche   
            path = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)
            path2 = uno.systemPathToFileUrl(path)
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

        self.mb.dict_bereiche.update({'Bereichsname':Bereichsname_dict})
        self.mb.dict_bereiche.update({'ordinal':ordinal_dict})
        self.mb.dict_bereiche.update({'Bereichsname-ordinal':Bereichsname_ord_dict})
        
         
        CB.loesche_leeren_Textbereich_am_Ende() 
       
        self.mb.current_Contr.addSelectionChangeListener(self.mb.VC_selection_listener)         
    
                   
    def erzeuge_dict_ordner(self):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_dict_ordner')

        tree = self.mb.xml_tree
        root = tree.getroot()
        
        ordner = []
        self.mb.dict_ordner = {}
        
        alle_eintraege = root.findall('.//')
        
        
        ### Vielleicht gibt es eine Moeglichkeit, den Baum nur einmal zu durchlaufen?
        ### Statt 1) Baum komplett durchlaufen 2) jeden Eintrag nochmals rekursiv durchlaufen
        ### Ziel: einmal durchlaufen und jedes Kind bei allen Elternelementen eintragen
        
        # Liste aller Ordner erstellen
        for eintrag in alle_eintraege:
            if eintrag.attrib['Art'] in ('dir','waste'):
                ordner.append(eintrag.tag)
        
        
        def get_tree_info(elem, dict,tag,helfer):
            helfer.append(elem.tag)
            # hier wird self.mb.dict_ordner geschrieben
            dict[tag] = helfer
            if elem.attrib['Zustand'] == 'auf':
                for child in list(elem):
                    get_tree_info(child, dict,tag,helfer)
        
        
        # Fuer alle Ordner eine Liste ihrer Kinder erstellen -> self.mb.dict_ordner       
        for tag in ordner:
            dir = root.find('.//'+tag)
            helfer = []
            get_tree_info(dir,self.mb.dict_ordner,tag,helfer)


    def lese_xml_datei(self):
        if self.mb.debug: print(self.mb.debug_time(),'lese_xml_datei')

        pfad = self.mb.pfade['settings'] + '/ElementTree.xml'        
        self.mb.xml_tree = self.mb.ET.parse(pfad)
        root = self.mb.xml_tree.getroot()
        
        self.mb.kommender_Eintrag = int(root.attrib['kommender_Eintrag'])
        
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
    
    
    def lese_xml_settings_datei(self):
        if self.mb.debug: print(self.mb.debug_time(),'lese_xml_settings_datei')
        
        pfad = os.path.join(self.mb.pfade['settings'], 'settings.xml')      
        self.mb.xml_tree_settings = self.mb.ET.parse(pfad)
        root = self.mb.xml_tree_settings.getroot()
      
        self.mb.tag1_visible = (root.find(".//tag1").attrib['sichtbar'] == 'ja')
        self.mb.tag2_visible = (root.find(".//tag2").attrib['sichtbar'] == 'ja')
        self.mb.tag3_visible = (root.find(".//tag3").attrib['sichtbar'] == 'ja')
        
    
    def erzeuge_Settings(self):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_Settings')
        
        et = self.mb.ET       
        root = et.Element('Settings')
        tree = et.ElementTree(root)
         
        et.SubElement(root,'tag1',sichtbar='ja')
        et.SubElement(root,'tag2',sichtbar='nein')
        et.SubElement(root,'tag3',sichtbar='nein')
        
        self.mb.tag1_visible = True
        self.mb.tag2_visible = False
        self.mb.tag3_visible = False
        
        self.mb.xml_tree_settings = tree
    
    
    def erzeuge_export_Settings(self):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_export_Settings')
        
        exp_settings = {
            # Export Dialog
            'alles' : 0, 
            'sichtbar' : 0,
            'eigene_ausw' : 1,
            
            'einz_dok' : 1,
            'trenner' : 1,
            'einz_dat' : 0,
            'ordner_strukt' : 0,
            'typ' : '.odt',
            'speicherort' : self.mb.pfade['projekt'].encode("utf-8"),
            
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
            }  
        
        
            
        path = os.path.join(self.mb.pfade['settings'],"export_settings.py")
        
        with codecs.open(path , "w", 'utf-8') as file:
            file.writelines('# -*- coding: utf-8 -*- \r\nexp_settings = ')          
            for line in str(exp_settings).split(','):
                file.writelines(line+(',\n'))
        
        self.mb.exp_settings = exp_settings
        
        
    def erzeuge_import_Settings(self):
        
        Settings = [('imp_dat','1'),
                ('ord_strukt','1'),
                
                ('url_dat',''),
                ('url_ord',''),
                
                #Filter
                ('odt','1'),
                ('doc','0'),
                ('docx','0'),
                ('rtf','0'),
                ('txt','0'),]
        
        
        
        et = self.mb.ET       
        root = et.Element('Import_Settings')

        for set in Settings:
            el = et.SubElement(root, set[0])
            el.text = set[1]
        
        tree = et.ElementTree(root)
        self.mb.imp_settings_xml = tree
        
        pfad = os.path.join(self.mb.pfade['settings'],'imp_settings.xml')
        tree.write(pfad)
        

          
       
        
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
            
            Eintraege = [('nr0','root','Vorbemerkung',0,'pg','-','ja','leer','leer','leer'),
                    ('nr1','root','Projekt',0,'dir','auf','ja','leer','leer','leer'),
                    ('nr2','nr1','Titelseite',1,'pg','-','ja','leer','leer','leer'),
                    ('nr3','nr1','Kapitel1',1,'dir','auf','ja','leer','leer','leer'),
                    ('nr4','nr3','Szene1',2,'pg','-','ja','leer','leer','leer'),
                    ('nr5','nr3','Szene2',2,'pg','-','ja','leer','leer','leer'),
#                     ('nr6','nr1','Kapitel2',1,'dir','auf','ja'),
#                     ('nr7','nr6','Szene1b',2,'pg','-','ja'),
#                     ('nr8','nr6','Szene2b',2,'pg','-','ja'),
#                     ('nr9','nr1','Interlude',1,'pg','-','ja'),
#                     ('nr10','nr1','Kapitel3',1,'dir','auf','ja'),
#                     ('nr11','nr1','Kapitel4',1,'dir','auf','ja'),
#                     ('nr12','nr11','UnterKapitel',2,'dir','zu','ja'),
#                     ('nr13','nr12','Szene3a',3,'pg','-','nein'),
#                     ('nr14','nr12','Szene3b',3,'pg','-','nein'),
#                     ('nr15','nr11','Szene3c',2,'pg','-','ja'),
#                     ('nr16','nr11','Szene3d',2,'pg','-','ja'),
#                     ('nr17','nr11','Kapitel4a',2,'dir','auf','ja'),
#                     ('nr18','nr1','Kapitel4b',1,'dir','auf','ja'),
#                     ('nr19','nr18','Szene4',2,'pg','-','ja'),
                    ('nr6','root','Papierkorb',0,'waste','zu','ja','leer','leer','leer')]
            
            return Eintraege
        
    def test(self):
        print('test')
        try:       
           

            
            pd()
            
        except:
            tb()
        
        
    
from com.sun.star.awt import XActionListener,XTopWindowListener,XKeyListener,XItemListener
class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,model):
        self.mb = mb
        self.model = model
        
    def actionPerformed(self,ev):
        
        filepath = None
        pfad = os.path.join(self.mb.path_to_extension,"pfade.txt")
        
        if os.path.exists(pfad):            
            with codecs.open( pfad, "r","utf-8") as file:
                filepath = file.read() 

        Filepicker = createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        if filepath != None:
            Filepicker.setDisplayDirectory(filepath)
        Filepicker.execute()
        
        if Filepicker.Directory == '':
            return
        
        filepath = Filepicker.getDirectory()
        
        with open( pfad, "w") as file:
            file.write(filepath)  
        
        self.mb.speicherort_last_proj = filepath
        self.model.Label = uno.fileUrlToSystemPath(filepath)




from com.sun.star.awt.Key import RETURN
class neues_Projekt_Dialog_Listener(unohelper.Base,XActionListener,XTopWindowListener,XKeyListener): 
    def __init__(self,mb,model_proj_name,model_neue_vorl,control_CB,control_LB):
        self.mb = mb
        self.model_proj_name = model_proj_name
        self.model_neue_vorl = model_neue_vorl
        self.control_CB = control_CB
        self.control_LB = control_LB
        
    def actionPerformed(self,ev):
        parent = ev.Source.AccessibleContext.AccessibleParent 
        cmd = ev.ActionCommand  
        if cmd == 'vorlage_erstellen':
            self.vorlage_auswaehlen()
        elif cmd == lang.WAEHLEN:
            return
        elif cmd == lang.CANCEL:
            parent.endDialog(0)
        elif self.model_proj_name.Text == '':
            self.mb.Mitteilungen.nachricht(lang.KEIN_NAME,"warningbox")
        elif self.mb.speicherort_last_proj == None:
            self.mb.Mitteilungen.nachricht(lang.KEIN_SPEICHERORT,"warningbox")
        elif cmd == lang.OK:
            self.get_path()
            parent.endDialog(1)
        
             
    def windowClosed(self,ev):
        pass
            
    def keyPressed(self,ev):
        if ev.KeyCode == RETURN:
            parent = ev.Source.AccessibleContext.AccessibleParent 
            if self.model.Text == '':
                parent.endDialog(0)
            else:
                self.get_path()
                parent.endDialog(1)

    def get_path(self):
        
        filepath = None
        pfad = os.path.join(self.mb.path_to_extension,"pfade.txt")
        
        if os.path.exists(pfad):            
            with codecs.open( pfad, "r","utf-8") as file:
                filepath = file.read() 
            self.mb.projekt_path = filepath
            
    def vorlage_auswaehlen(self):        
                    
        if self.model_neue_vorl.Text == '':
            self.mb.Mitteilungen.nachricht(lang.WARNUNG_NAME_TEMPLATE,"warningbox")
            return
        
        else:
            name = self.model_neue_vorl.Text
            
            if name in self.mb.user_styles:
                self.mb.Mitteilungen.nachricht(lang.WARNUNG_TEMPLATE_EXISTS   ,"warningbox")
                return
            
            else:
                if self.mb.use_template[0] == False:
                    URL = "private:factory/swriter"
                else:
                    URL = uno.systemPathToFileUrl(self.mb.use_template[1])
                    
                self.vorlage_erstellen(name,URL)
                self.control_LB.Enable = True
                self.control_CB.Enable = True
                
                if self.control_LB.Items[0] == lang.NO_TEMPLATES:
                    
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
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True                
            print(URL)
            
            
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
            
            #
        except:
            tb()
        newDoc.close(False)
        self.mb.Mitteilungen.nachricht(lang.NEUES_TEMPLATE + '\n%s   ' % p1,"infobox")
        
        # var DefaultTemplate = ((XDocumentPropertiesSupplier)oDoc).getDocumentProperties().TemplateURL;
        
        
class Neues_Projekt_CheckBox_Listener(unohelper.Base, XActionListener,XItemListener):
    def __init__(self,mb,modelStandard,modelUser,modelListBox):
        self.mb = mb
        self.modelStandard = modelStandard
        self.modelUser = modelUser
        self.modelListBox = modelListBox
        
    def actionPerformed(self,ev):
        #set = self.mb.exp_settings
        #set[ev.ActionCommand] = self.toggle(set[ev.ActionCommand])
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
            return
        elif ev.ActionCommand == 'standard':
            self.modelUser.State = False
            self.mb.use_template = (False,None)
        elif ev.ActionCommand == 'user':
            self.modelStandard.State = False
            pfad = self.get_template_pfad()
            self.mb.use_template = (True,pfad)
        
    
    def get_template_pfad(self):
        pfade = self.mb.user_styles_pfade
        gewaehlt = self.modelListBox.SelectedItems[0]
        pfad = pfade[gewaehlt]

        return pfad
        
    
    def itemStateChanged(self, ev):  

        use,pfad = self.mb.use_template
        self.mb.use_template = (use,self.get_template_pfad())





        
        
        
        
################ TOOLS ################################################################

# Handy function provided by hanya (from the OOo forums) to create a control, model.
def createControl(ctx,type,x,y,width,height,names,values):
   smgr = ctx.getServiceManager()
   ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
   ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
   ctrl_model.setPropertyValues(names,values)
   ctrl.setModel(ctrl_model)
   ctrl.setPosSize(x,y,width,height,15)
   return (ctrl, ctrl_model)


def createUnoService(serviceName):
  sm = uno.getComponentContext().ServiceManager
  return sm.createInstanceWithContext(serviceName, uno.getComponentContext())
                   

