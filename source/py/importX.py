# -*- coding: utf-8 -*-

import unohelper


class ImportX():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.fenster_import = None
        self.fenster_filter = None
        
        self.auszuschliessende_filter = ('org.openoffice.da.writer2bibtex','org.openoffice.da.writer2latex','BibTeX_Writer','LaTeX_Writer')
        
        
    def importX(self):
        if self.mb.debug: log(inspect.stack)
        try:
   
            if self.mb.filters_import == None:
                self.erzeuge_filter()
            
            if self.fenster_import != None:
                self.fenster_import.toFront()
                
                if self.fenster_filter != None:
                    self.fenster_filter.toFront()
            else:
                self.erzeuge_importfenster()
                
        except Exception as e:
            self.mb.nachricht('ImportX ' + str(e),"warningbox")
            log(inspect.stack,tb())

 
    def erzeuge_importfenster(self): 
        if self.mb.debug: log(inspect.stack)
        
        breite = 400
        hoehe = 460
        
        imp_set = self.mb.settings_imp
        
        

        posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
        X = self.mb.dialog.Size.Width 
        Y = posSize_main.Y 
        Width = breite
        Height = hoehe
        
        posSize = X,Y,Width,Height
        fenster,fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
        fenster_cont.Model.Text = LANG.IMPORT
        
        self.fenster_import = fenster
        listenerDisp = Fenster_Dispose_Listener(self.mb,self)
        fenster_cont.addEventListener(listenerDisp)
        
        y = 10
        
        # Titel
        controlE, modelE = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,50,20,(),() )  
        controlE.Text = LANG.IMPORT
        modelE.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlE)
        self.mb.kalkuliere_und_setze_Control(controlE,'w')
        
        y += 30
        
        control1, model1 = self.mb.createControl(self.mb.ctx,"CheckBox",20,y,120,22,(),() )  
        model1.Label = LANG.IMPORT_DATEI
        model1.State = int(imp_set['imp_dat'])
        control1.ActionCommand = 'model1'
        fenster_cont.addControl('ImportD', control1)
        self.mb.kalkuliere_und_setze_Control(control1,'w')
        
        y += 30
        
        control2, model2 = self.mb.createControl(self.mb.ctx,"CheckBox",20,y,120,22,(),() )  
        model2.Label = LANG.IMPORT_ORDNER
        model2.State = not int(imp_set['imp_dat'])
        control2.ActionCommand = 'model2'
        fenster_cont.addControl('ImportO', control2)
        self.mb.kalkuliere_und_setze_Control(control2,'w')
        
        y += 20
        
        control3, model3 = self.mb.createControl(self.mb.ctx,"CheckBox",40,y,180,22,(),() )  
        model3.Label = LANG.ORDNERSTRUKTUR
        model3.State =  int(imp_set['ord_strukt'])
        control3.Enable = not int(imp_set['imp_dat'])
        control3.ActionCommand = 'struktur'
        fenster_cont.addControl('struktur', control3)
        self.mb.kalkuliere_und_setze_Control(control3,'w')
        
        # CheckBox Listener
        listenerA = Auswahl_CheckBox_Listener(self.mb,model1,model2,control3)
        control1.addActionListener(listenerA)
        control2.addActionListener(listenerA)
        control3.addActionListener(listenerA)
                    
        y += 40
        
        # Trenner -----------------------------------------------------------------------------
        controlT, modelT = self.mb.createControl(self.mb.ctx,"FixedLine",20,y-10 ,breite - 40,40,(),() )  
        fenster_cont.addControl('Trenner', controlT)
        
        y += 40
        
        controlE, modelE = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,50,20,(),() )  
        controlE.Text = LANG.DATEIFILTER
        fenster_cont.addControl('Filter', controlE)
        self.mb.kalkuliere_und_setze_Control(controlE,'w')
        
        buttons = {}
        filters = ('odt','doc','docx','rtf','txt')
        filter_CB_listener = Filter_CheckBox_Listener(self.mb,buttons,fenster)
        for filt in filters:
            control, model = self.mb.createControl(self.mb.ctx,"CheckBox",100 ,y,80,22,(),() ) 
            control.ActionCommand = filt 
            model.Label = filt
            model.State = imp_set[filt]    
            control.addActionListener(filter_CB_listener)                 
            fenster_cont.addControl(filt, control)
            buttons.update({filt:control})
            if imp_set['auswahl'] == 1:
                control.Enable = False
            y += 16

        
        control, model = self.mb.createControl(self.mb.ctx,"CheckBox",180 ,y-5*16,100,22,(),() )  
        buttons.update({'auswahl':control})
        control.ActionCommand = 'auswahl'
        model.Label = LANG.EIGENE_AUSWAHL
        model.State = imp_set['auswahl']   
        control.addActionListener(filter_CB_listener)                    
        fenster_cont.addControl('auswahl', control)      
        self.mb.kalkuliere_und_setze_Control(control,'w')
        
        
        control, model = self.mb.createControl(self.mb.ctx,"Button",180 + 16,y - 3*16  ,80,20,(),() )  ###
        control.Label = LANG.AUSWAHL
        fenster_cont.addControl('Filter_Auswahl', control)  
        self.mb.kalkuliere_und_setze_Control(control,'w')
        
        filter_Ausw_B_listener = Filter_Auswahl_Button_Listener(self.mb,fenster)
        control.addActionListener(filter_Ausw_B_listener)  
        
        
        y += 20   
        
        # Trenner -----------------------------------------------------------------------------
        controlT, modelT = self.mb.createControl(self.mb.ctx,"FixedLine",20,y-10 ,breite - 40,40,(),() )  
        fenster_cont.addControl('Trenner', controlT)
        
        y += 40

        controlA, modelA = self.mb.createControl(self.mb.ctx,"Button",20,y  ,110,20,(),() )  ###
        controlA.Label = LANG.DATEIAUSWAHL
        controlA.ActionCommand = 'Datei'
        fenster_cont.addControl('Auswahl', controlA) 
        self.mb.kalkuliere_und_setze_Control(controlA,'w')
        
        y += 25
        
        controlE, modelE = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,600,20,(),() )  
        text = imp_set['url_dat']
        if text == None or text == '': 
            text = '-'
        else:
            text = uno.fileUrlToSystemPath(decode_utf(text))
        controlE.Text = text
        fenster_cont.addControl('Filter', controlE)    
        self.mb.kalkuliere_und_setze_Control(controlE,'w')
        
        y += 40   
        
        controlA2, modelA2 = self.mb.createControl(self.mb.ctx,"Button",20,y  ,110,20,(),() )  ###
        controlA2.Label = LANG.ORDNERAUSWAHL
        controlA2.ActionCommand = 'Ordner'
        fenster_cont.addControl('Auswahl2', controlA2)             
        self.mb.kalkuliere_und_setze_Control(controlA2,'w')
        
        y += 25
        
        controlE2, modelE2 = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,600,20,(),() )  
        text = imp_set['url_ord']
        if text == None or text == '': 
            text = '-'
        else:
            text = uno.fileUrlToSystemPath(decode_utf(text))
        controlE2.Text = text
        fenster_cont.addControl('Filter2', controlE2) 
        
        listenerA2 = Auswahl_Button_Listener(self.mb,modelE,modelE2)
        controlA.addActionListener(listenerA2)
        controlA2.addActionListener(listenerA2)   
        
        
        y += 60
        
        controlI, modelI = self.mb.createControl(self.mb.ctx,"Button",breite - 80 - 20,y  ,80,30,(),() )  ###
        controlI.Label = LANG.IMPORTIEREN
        listener_imp = Import_Button_Listener(self.mb,fenster)
        controlI.addActionListener(listener_imp)
        self.mb.kalkuliere_und_setze_Control(controlI,'w')
        fenster_cont.addControl('importieren', controlI) 
        
        fenster.setPosSize(0,0,0,y + 40,8)
    
    
    def get_flags(self,x):
        #if self.mb.debug: log(inspect.stack)
        try:
            x_bin_rev = bin(x).split('0b')[1][::-1]
    
            flags = []
            
            for i in range(len(x_bin_rev)):
                z = 2**int(i)*int(x_bin_rev[i])
            
                if z != 0:
                    flags.append(z)
            
            return flags
        except:
            return 'xxxxxx'
    
    
    def erzeuge_filter(self):
        if self.mb.debug: log(inspect.stack)
        try:
            typeDet = self.mb.createUnoService("com.sun.star.document.TypeDetection")
            FF = self.mb.createUnoService( "com.sun.star.document.FilterFactory" )
            FilterNames = FF.getElementNames()
            
            # da in OO und LO die Positionen der Eintraege im MediaDescriptor verschieden sind,
            # wird zuerst nach den richtigen Positionen gesucht
            
            fil = FF.getByName(FilterNames[0])
            fil2 = typeDet.getByName(typeDet.ElementNames[0])

            i = 0
            
            for fi in fil:
                if fi.Name == 'UIName':
                    uiName_pos = i
                elif fi.Name == 'Name':
                    name_pos = i
                elif fi.Name == 'Type':
                    type_pos = i
                elif fi.Name == 'DocumentService':
                    docSer_pos = i
                elif fi.Name == 'Flags':
                    flags_pos = i
                    
                i += 1          
            
            i = 0
            
            for filx in fil2:
                if filx.Name == 'Extensions':
                    extensions_pos = i
                     
                i += 1   
    
            def formatiere(term):
                term = term.replace("'","")
                term = term.replace('[','')
                term = term.replace(']','')
                term = term.replace(", "," (",1)
                term = term.replace(",","")
                term = term +')'
                return term
            
            def formatiere2(term):
                ter = []
                for t in term:
                    ter.append('*.'+str(t).replace("'",""))
                return list(ter)
            
            
            self.mb.filters_import = {}   
            self.mb.filters_export = {}    
            
            for filt in FilterNames:

                if filt in self.auszuschliessende_filter:
                    continue

                f = FF.getByName(filt)
                if f[docSer_pos].Value == 'com.sun.star.text.TextDocument':
                    # doppelt s.o. // sollte geaendert werden
                    flags_pos = [f.index(i) for i in f if i.Name == 'Flags'][0]
                    uiName_pos = [f.index(i) for i in f if i.Name == 'UIName'][0]
                    type_pos = [f.index(i) for i in f if i.Name == 'Type'][0]
                    
                    flags = self.get_flags(f[flags_pos].Value)
                    
                    if 1 in flags or 2 in flags:
                        uiName = f[uiName_pos].Value  
                        filter_typeDet = typeDet.getByName(f[type_pos].Value)
                        
                        extensions = filter_typeDet[extensions_pos].Value
                        label2 = formatiere2(extensions)
                        label2.insert(0,str(uiName))
                        
                        if 1 in flags:
                            # FilterName: Filter als Label, Extensions Endungen
                            self.mb.filters_import.update({filt:(formatiere(str(label2)),extensions)})
                        
                        if 2 in flags:  
                            if 'bib' in extensions:
                                pass 
#                             time.sleep(0.04)
#                             print(extensions)                    
                            # FilterName: Filter als Label, Extensions Endungen
                            self.mb.filters_export.update({filt:(formatiere(str(label2)),extensions)})

            self.mb.filters_export.update({'LaTex':('LaTex (Organon) (*.tex)','tex','*tex')})
            self.mb.filters_export.update({'HtmlO':('HTML (Organon) (*.html)','html','*html')})
            
        except:
            log(inspect.stack,tb())
                        
    def warning(self,mb):
        if self.mb.debug: log(inspect.stack)
        
        LANG.IMPORT_WARNING = "Don't try to import files, which hold any links to other files, ole-ojects, internet-links etc. inside."
        entscheidung = mb.nachricht(LANG.IMPORT_WARNING,"warningbox",16777216)
        # 3 = Nein oder Cancel, 2 = Ja
        if entscheidung == 3:
            return False
        elif entscheidung == 2:
            return True


def encode_utf(term):
    if isinstance(term, str):
        return term
    else:
        return term.encode('utf8')

def decode_utf(term):
    if isinstance(term, str):
        return term
    else:
        return term.decode('utf8')   

def escape_xml(term):
    Zeichen = {
               ' ' : '_Leerzeichen_',
               '(' : '_KlammerAuf_',
               ')' : '_KlammerZu_',
               '.' : '_PuNkt_'}
    
    for z in Zeichen:        
        term = term.replace(z, Zeichen[z])
    return term

def unescape_xml(term):
    Zeichen = {
               '_Leerzeichen_' : ' ',
               '_KlammerAuf_' : '(',
               '_KlammerZu_' : ')',
               '_PuNkt_' : '.'}
    
    for z in Zeichen:        
        term = term.replace(z, Zeichen[z])
    return term


from com.sun.star.lang import XEventListener
class Fenster_Dispose_Listener(unohelper.Base, XEventListener):
    # Listener um Position zu bestimmen
    def __init__(self,mb,class_Import):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.class_Import = class_Import
        
    def disposing(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        self.class_Import.fenster_import = None

        if self.class_Import.fenster_filter != None:
            self.class_Import.fenster_filter.dispose()
            self.class_Import.fenster_filter = None
        # Settings speichern
        self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp) 
        

from com.sun.star.awt import XItemListener, XActionListener, XFocusListener  
from com.sun.star.style.BreakType import NONE as BREAKTYPE_NONE  
class Import_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,fenster):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.fenster = fenster
        
        

        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp
        try:
            if int(imp_set['imp_dat']) == 1:
                
                if imp_set['url_dat'] == '':
                    self.mb.nachricht(LANG.ERST_DATEI_AUSWAEHLEN,"warningbox")
                    self.fenster.toFront()
                    return
                
                self.mb.Listener.remove_Undo_Manager_Listener()
                self.datei_importieren('dokument',imp_set['url_dat'])
                self.mb.Listener.add_Undo_Manager_Listener()
                
            else:
                if imp_set['url_ord'] == '':
                    self.mb.nachricht(LANG.ERST_ORDNER_AUSWAEHLEN,"warningbox")
                    self.fenster.toFront()
                    return
    
                self.fenster.dispose()
                self.mb.class_Import.fenster_import = None
                
                self.mb.Listener.remove_Undo_Manger_Listener()
                lade = self.ordner_importieren()
    
                if lade:
                    self.mb.class_Projekt.lade_Projekt2()
                    self.mb.nachricht(LANG.IMPORT_ABGESCHLOSSEN,'infobox')
                    
                self.mb.Listener.add_Undo_Manager_Listener()
        except:
            log(inspect.stack,tb())

            
    def datei_importieren(self,typ,url_dat):
        if self.mb.debug: log(inspect.stack)
        
        zeile_nr,zeile_pfad = self.erzeuge_neue_Zeile()
        
        ordinal_neuer_Eintrag = 'nr'+str(zeile_nr)
        bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal_neuer_Eintrag]
        sections = self.mb.doc.TextSections
        neuer_Bereich = sections.getByName(bereichsname) 
                            
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
    
        self.oOO = self.mb.desktop.loadComponentFromURL(url_dat,'_blank',8+32,(prop,))
        
        # moeglicherweise vorhandene Links entfernen
        self.entferne_links(self.oOO)
        self.kapsel_in_Bereich(self.oOO,str(zeile_nr))
        self.entferne_seitenumbrueche_am_anfang(self.oOO)
                
        prop3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop3.Name = 'FilterName'
        prop3.Value = 'writer8'

        Path2 = uno.systemPathToFileUrl(zeile_pfad)
        self.oOO.storeToURL(Path2,(prop3,))
        self.oOO.close(False)
        
        #SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink = neuer_Bereich.FileLink
         
        neuer_Bereich.setPropertyValue('FileLink',SFLink)
        
        # zeilenname der neuen Datei setzen
        zeile_hf = self.mb.props[T.AB].Hauptfeld.getControl('nr'+str(zeile_nr))
        zeile_textfeld = zeile_hf.getControl('textfeld')

        link_quelle1 = uno.fileUrlToSystemPath(url_dat)
        endung = os.path.splitext(link_quelle1)[1]
        name1 = os.path.basename(link_quelle1)
        name2 = name1.split(endung)[0]

        zeile_textfeld.Model.Text = name2
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot() 
        
        xml_elem = root.find('.//'+ ordinal_neuer_Eintrag)
        xml_elem.attrib['Name'] = name2

        Path = os.path.join(self.mb.pfade['settings'] , 'ElementTree.xml' )
        self.mb.tree_write(tree,Path)
        
        self.mb.class_Sidebar.lege_dict_sb_content_ordinal_an(ordinal_neuer_Eintrag)
        
        return ordinal_neuer_Eintrag,bereichsname
     
    
    def entferne_seitenumbrueche_am_anfang(self,oOO):
        if self.mb.debug: log(inspect.stack)
        
        try:
            enum = oOO.Text.createEnumeration()
            para = enum.nextElement() 
            para.BreakType = BREAKTYPE_NONE
            para.PageDescName = ''
        except:
            log(inspect.stack,tb())
        
    
            
    def entferne_links(self,oOO):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.mb.class_Bereiche.verlinkte_Bilder_einbetten(oOO)
            
            all_links = []
            
            for name in oOO.Links.ElementNames:
                all_links.append( oOO.Links.getByName(name))
    
            for l in all_links:
                if len(l.ElementNames) != 0:
                    
                    # Sections (=region) entfernen
                    if 'region' in l.ElementNames[0]:
                        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                        for li in l.ElementNames:
                            link = l.getByName(li)
                            link.setPropertyValue('FileLink',SFLink)
                            link.IsVisible = True
                            if 'OrganonSec' in link.Name:
                                link.Name = 'OldOrgSec'+link.Name.split('OrganonSec')[1]
    
        except:
            log(inspect.stack,tb())
    
    def kapsel_in_Bereich(self,oOO,ordn):
        if self.mb.debug: log(inspect.stack)
        
        cur = oOO.Text.createTextCursor()
        cur.gotoEnd(False)
        
        if cur.TextSection != None:
            try:
                # Zur Sicherheit, falls der Text nur aus einem Textbereich besteht,
                # wird eine Zeile angehangen
                prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
                prop.Name = 'ParaStyleName'
                prop.Value = 'Standard'   
                oOO.Text.appendParagraph((prop,))
            except:
                pass
        
        cur.gotoStart(False)
        cur.gotoEnd(True)
        cur.goRight(1,True)
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        newSection.setName('OrgInnerSec'+ordn)
        
        oOO.Text.insertTextContent(cur,newSection,True)
        
        
            
    def erzeuge_neue_Zeile(self):
        if self.mb.debug: log(inspect.stack)
        
        if self.mb.props[T.AB].selektierte_zeile == None:       
            self.mb.nachricht(self.mb.LANG.ZEILE_AUSWAEHLEN,'infobox')
            return None
        else:
            
            
                      
            ord_sel_zeile = self.mb.props[T.AB].selektierte_zeile
            
            # XML TREE
            tree = self.mb.props[T.AB].xml_tree
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
            art = 'pg'
            zustand = '-'            
            eintrag = ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 
            
            # neue Zeile / neuer XML Eintrag
            self.mb.class_XML.erzeuge_XML_Eintrag(eintrag)
            
            if self.mb.settings_proj['tag3']:
                gliederung = self.mb.class_Gliederung.rechne(tree)
            else:
                gliederung = None
            
            self.mb.class_Baumansicht.erzeuge_Zeile_in_der_Baumansicht(eintrag,self.mb.class_Zeilen_Listener,gliederung)
            
                        
            # neue Datei / neuen Bereich anlegen           
            # kommender Eintrag wurde in erzeuge_XML_Eintrag schon erhoeht
            nr = int(root.attrib['kommender_Eintrag']) - 1          
            inhalt = ordinal

            path = os.path.join(self.mb.pfade['odts'] , 'nr%s.odt' %nr)             
            self.erzeuge_bereich3(nr,path,sicht)

            # Zeilen anordnen
            source = ordinal
            target = xml_sel_zeile.tag
            action = 'drunter'  
            
            eintraege = self.mb.class_Zeilen_Listener.xml_neu_ordnen(source,target,action)
            self.mb.class_Zeilen_Listener.posY_in_tv_anpassen(eintraege)
            
            # dict_ordner updaten
            self.mb.class_Projekt.erzeuge_dict_ordner()

            # Bereiche neu verlinken
            sections = self.mb.doc.TextSections
            self.mb.class_Zeilen_Listener.verlinke_Bereiche(sections)
            # Sichtbarkeit der Bereiche umschalten
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_der_Bereiche() 

            return nr,path    
        
    
    def erzeuge_bereich3(self,i,path,sicht):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 
        
        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
        
        alle_hotzenploetze = []
        for sec_name in sections.ElementNames:
            if 'OrganonSec' in sec_name:
                alle_hotzenploetze.append(sec_name)
        anzahl_hotzenploetze = len(alle_hotzenploetze)
                
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = path
        SFLink.FilterName = 'writer8'
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        newSection.setName('OrganonSec'+str(anzahl_hotzenploetze))
        
        if sicht == 'nein':
            newSection.IsVisible = False

        if sections.Count == 0:
            # bei leerem Projekt
            textSectionCursor = text.createTextCursor()
        else:
            sectionN = sections.getByIndex(sections.Count-1)
            textSectionCursor = text.createTextCursorByRange(sectionN.Anchor)
            textSectionCursor.gotoEnd(False)
        
        text.insertTextContent(textSectionCursor, newSection, False)
 
    
    def ordner_importieren(self):
        if self.mb.debug: log(inspect.stack)
        
        importXml,links_und_filter,erfolg = self.erzeuge_elementTree()
        
        if erfolg:
            self.fuege_importXml_in_xml_ein(importXml)
            
            path = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
            self.mb.tree_write(self.mb.props[T.AB].xml_tree,path)
            
            self.neue_Dateien_erzeugen(importXml,links_und_filter)
            
            return True
        else:
            return False

 
    def neue_Dateien_erzeugen(self,importXml,links_und_filter):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.Listener.remove_VC_selection_listener()
        
        speicherordner = os.path.join(self.mb.pfade['odts'])
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
                    
        anzahl_links = len(links_und_filter)
        index = 1

        for ordn,link_filt in links_und_filter.items():
            
            link,filt = link_filt
            
            StatusIndicator = self.mb.desktop.getCurrentFrame().createStatusIndicator()
            StatusIndicator.start(LANG.ERZEUGE_DATEI %(index,anzahl_links),anzahl_links)
            StatusIndicator.setValue(index)
            index += 1
        
            if filt != None:
                prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
                prop2.Name = 'FilterName'
                prop2.Value = filt[0]
                  
                props = prop,prop2
            else:
                props = prop,
              

            self.oOO = self.mb.desktop.loadComponentFromURL(link,'_blank',8+32,(props))

            self.entferne_links(self.oOO)
            self.kapsel_in_Bereich(self.oOO,ordn)
            self.entferne_seitenumbrueche_am_anfang(self.oOO)
            
            prop3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop3.Name = 'FilterName'
            prop3.Value = 'writer8'
            
            pfad = os.path.join(speicherordner,ordn +'.odt')
            pfad2 = uno.systemPathToFileUrl(pfad)
            
            self.oOO.storeToURL(pfad2,(prop3,))
            self.oOO.close(False)
            
            self.mb.class_Sidebar.lege_dict_sb_content_ordinal_an(ordn)
            
            StatusIndicator.end()

 
    def fuege_importXml_in_xml_ein(self,importXml):
        if self.mb.debug: log(inspect.stack)
        
        root = self.mb.props[T.AB].xml_tree.getroot()
        
        name_selek_zeile = self.mb.props[T.AB].selektierte_zeile
        xml_selekt_zeile = root.find('.//'+name_selek_zeile)
        
        parent = root.find('.//'+xml_selekt_zeile.tag+'/..')
        
        liste = list(parent)
        index_sel = liste.index(xml_selekt_zeile)
        parent.insert(index_sel+1,importXml)       
        
    
    def erzeuge_elementTree(self):
        if self.mb.debug: log(inspect.stack)

        imp_set = self.mb.settings_imp
        path = uno.fileUrlToSystemPath(imp_set['url_ord'])
        
        if not os.path.exists(path):
            self.mb.nachricht(LANG.NO_FILES,"infobox")
            return None,None,False
        
        Verzeichnis,AusgangsOrdner = self.durchlaufe_ordner(path)
        anz = len(Verzeichnis)
        
        if anz == 0:
            self.mb.nachricht(LANG.NO_FILES,"infobox")
            return None,None,False
        elif anz > 10:
            entscheidung = self.mb.nachricht(LANG.IMP_FORTFAHREN %anz,"warningbox",16777216)
            # 3 = Nein oder Cancel, 2 = Ja
            if entscheidung == 3:
                return None,None,False
            
                    
        ordinal_neuer_Eintrag = 'nr%s' %self.mb.props[T.AB].kommender_Eintrag
        self.mb.props[T.AB].kommender_Eintrag += 1

        et = self.mb.ET
        root_xml = et.Element(ordinal_neuer_Eintrag)
        tree_xml = et.ElementTree(root_xml)
        
        name_selek_zeile = self.mb.props[T.AB].selektierte_zeile
        xml_selekt_zeile = self.mb.props[T.AB].xml_tree.getroot().find('.//'+name_selek_zeile)
        parent_sel_zeile = self.mb.props[T.AB].xml_tree.getroot().find('.//'+name_selek_zeile+'/..')
        lvl = int(xml_selekt_zeile.attrib['Lvl'])
        
        # Der Hauptordner
        self.setze_attribute(root_xml,AusgangsOrdner,'dir',lvl,parent_sel_zeile)
        
        url_links = {}
        url_links.update({ordinal_neuer_Eintrag:("private:factory/swriter",None)})
        
        if int(imp_set['ord_strukt']) == 0:               
            
            for eintr in Verzeichnis:
                File,Pfad,Ordner,filt = eintr
                
                ordinal_neuer_Eintrag = 'nr%s' %self.mb.props[T.AB].kommender_Eintrag
                self.mb.props[T.AB].kommender_Eintrag += 1
                
                element = et.SubElement(root_xml,ordinal_neuer_Eintrag)
                
                lvl = int(root_xml.attrib['Lvl']) + 1
                name = os.path.splitext(File)[0]
                self.setze_attribute(element,name,'pg',lvl,root_xml)
                                    
                file_url = uno.systemPathToFileUrl(Pfad)
                url_links.update({ordinal_neuer_Eintrag:(file_url,None)})
                
            root = self.mb.props[T.AB].xml_tree.getroot()
            root.attrib['kommender_Eintrag'] = str(self.mb.props[T.AB].kommender_Eintrag)
            
            return root_xml,url_links,True                    
            
        else:
            for eintr in Verzeichnis:
                File,Pfad,Ordner,filt = eintr
                elem = root_xml
                
                for o in Ordner[::-1]:
                    o = escape_xml(o)
       
                    if unescape_xml(o) != AusgangsOrdner:

                        x = elem.find(o)
                        if x == None:
                            
                            ordinal_neuer_Eintrag = 'nr%s' %self.mb.props[T.AB].kommender_Eintrag
                            self.mb.props[T.AB].kommender_Eintrag += 1
                            
                            url_links.update({ordinal_neuer_Eintrag:("private:factory/swriter",None)})
                            name = unescape_xml(o)
                            
                            #print('schreibe Ordner',name,'in Ordner',elem.attrib['Name']) 
                            et.SubElement(elem,o)
                            element = elem.find(o)
                            element.text = ordinal_neuer_Eintrag
                            lvl = int(elem.attrib['Lvl']) + 1
                            self.setze_attribute(element,name,'dir',lvl,elem)
                                                            
                        elem = elem.find(o)
                
                #print('schreibe Datei',File,'in Ordner',o) 
                file_url = uno.systemPathToFileUrl(Pfad)
                ordinal_neuer_Eintrag = 'nr%s' %self.mb.props[T.AB].kommender_Eintrag
                self.mb.props[T.AB].kommender_Eintrag += 1
                url_links.update({ordinal_neuer_Eintrag:(file_url,filt)})

                et.SubElement(elem,ordinal_neuer_Eintrag)
                
                element = elem.find(ordinal_neuer_Eintrag)
                lvl = int(elem.attrib['Lvl']) + 1
                name = os.path.splitext(File)[0]
                self.setze_attribute(element,name,'pg',lvl,elem)
                
            
            # Tag der Ordner auf deren Text setzen
            for elemen in root_xml.findall('.//'):
                if elemen.text != None:
                    elemen.tag = elemen.text
                    for child in list(elemen):
                        child.attrib['Parent'] = elemen.tag

            root = self.mb.props[T.AB].xml_tree.getroot()
            root.attrib['kommender_Eintrag'] = str(self.mb.props[T.AB].kommender_Eintrag)
            
            return root_xml,url_links,True


    
            
    def setze_attribute(self,element,name,art,level,parent):
        if self.mb.debug: log(inspect.stack)
        
        if art == 'dir':
            zustand = 'auf'
        else:
            zustand = '-'
        
        element.attrib['Name'] = name
        element.attrib['Zustand'] = zustand
        element.attrib['Sicht'] = 'ja'
        element.attrib['Parent'] = parent.tag
        element.attrib['Lvl'] = str(level)
        element.attrib['Art'] = art
        
        element.attrib['Tag1'] = 'leer'
        element.attrib['Tag2'] = 'leer'
        element.attrib['Tag3'] = 'leer'
    
           
    def get_filter_endungen(self):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp  
        filterendungen = []

        if not int(imp_set['auswahl']):
            endungen = 'odt','doc','docx','rtf','txt'
            for end in endungen:
                if int(imp_set[end]) == 1:
                    filterendungen.append(end)
        else:
            filterauswahl = imp_set['filterauswahl']
            imp_filter = self.mb.filters_import
            
            filt = []
            filterendungen = []

            for f in filterauswahl:
                if filterauswahl[f] == 1:
                    filt.append(f)

            for f in filt:
                if f in list(imp_filter):
                    end = imp_filter[f][1]
                    for e in end:
                        if e not in filterendungen:
                            filterendungen.append(e)

        return filterendungen

    
    def durchlaufe_ordner(self,path):
        if self.mb.debug: log(inspect.stack)
        
        # Vorsortierung, um nicht fuer jede Datei deep scan durchzufuehren
        filterendungen = self.get_filter_endungen()

        def find_parents(ordn,AusgangsOrdner,Liste):
            drueber = os.path.dirname(ordn)
            uebergeordneter = os.path.split(drueber)[1]
            
            if uebergeordneter != AusgangsOrdner:
                Liste.append(uebergeordneter)
                find_parents(drueber,AusgangsOrdner,Liste)
        
        Verzeichnis = []
        first_time = True
        
        for root, dirs, files in os.walk(path):

            if first_time:
                AusgangsOrdner = os.path.split(root)[1]
                first_time = False
                drueber = os.path.dirname(root)            
            
            Ordner = os.path.split(root)[1]
    
            for file in files:
                # Vorsortierung, um nicht fuer jede Datei deep scan durchzufuehren
                if file.split('.')[-1] in filterendungen:
                    bool,filt = self.deep_scan(root,file)

                    if bool:
                        if Ordner != AusgangsOrdner:
                            Liste = []
                            find_parents(root,AusgangsOrdner,Liste)
                            
                            pfad = os.path.join(root,file)
                            Liste.insert(0,Ordner)
                            Verzeichnis.append((file,pfad,Liste,filt)) 
                        else:
                            Liste = [AusgangsOrdner]
                            pfad = os.path.join(root,file)
                            Verzeichnis.append((file,pfad,Liste,filt)) 

        return Verzeichnis,AusgangsOrdner

  
    def deep_scan(self,root,file):
        if self.mb.debug: log(inspect.stack)
        
        if int(self.mb.settings_imp['auswahl']) == 0:
            return True,None
        else:
            pfad = os.path.join(root,file)
            
            url = uno.systemPathToFileUrl(pfad)
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'URL'
            prop.Value = url
            
            typeDet = self.mb.createUnoService("com.sun.star.document.TypeDetection")
            filter_det = typeDet.queryTypeByDescriptor((prop,),True)     
            
            for i in range(len(filter_det[1])): 
                if filter_det[1][i].Name == 'FilterName':
                    pos_name = i     
            
            gefunden = None
            for filt in self.mb.filters_import:
                if filter_det[1][pos_name].Value == filt:
                    gefunden = filt
                    break
                
            if gefunden != None:
                if gefunden in self.mb.settings_imp['filterauswahl']:
                    if self.mb.settings_imp['filterauswahl'][gefunden]:
                        return True,filter_det

            return False,None
   
    def disposing(self,ev):
        return False

class Import_CheckBox_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,buttons,but_Auswahl):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.buttons = buttons
        self.but_Auswahl = but_Auswahl
        
    # XItemListener    
    def itemStateChanged(self, ev):   
        if self.mb.debug: log(inspect.stack)
             
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
       
     

class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,model1,model2,contr_strukt):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model1 = model1
        self.model2 = model2
        self.contr_strukt = contr_strukt
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp

        if ev.ActionCommand == 'struktur':
            if imp_set['ord_strukt'] == 0:
                imp_set['ord_strukt'] = 1
            else:
                imp_set['ord_strukt'] = 0                
        elif ev.Source.State == 0:
            ev.Source.State = 1
            return

        if ev.ActionCommand == 'model1':
            self.model2.State = False
            imp_set['imp_dat'] = '1'
            self.contr_strukt.Enable = False
        elif ev.ActionCommand == 'model2':
            self.model1.State = False
            imp_set['imp_dat'] = '0'
            self.contr_strukt.Enable = True
        else:
            pass

        self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  
    
    def disposing(self,ev):
        return False
        
           
class Filter_Auswahl_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,importfenster):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.importfenster = importfenster
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        imp_set = self.mb.settings_imp  
        
        if self.mb.programm == 'OpenOffice':
            # OO erzeugt unregelmaessig einen Fehler, den ich nicht finden konnte, und stuerzt ab
            # OO verliert die Werte fuer den Eintrag 'writer8' im dict filters_import
            # daher werden die filter hier noch mal berechnet
            self.erzeuge_filter()
            filters = self.filters_import
        else:
            filters = self.mb.filters_import
        
    
        ps = self.importfenster.PosSize
        posSize = ps.X+ps.Width+20,ps.Y,400,len(filters)*16 + 70
        
        fenster, fenster_cont = self.mb.erzeuge_Dialog_Container(posSize)
        
        self.mb.class_Import.fenster_filter = fenster
        
        buttons = {}
        f_listener = Filter_CheckBox_Listener2(self.mb,buttons)

        y = 20
        
        # Titel
        control, model = self.mb.createControl(self.mb.ctx,"FixedText",20,y ,80,20,(),() )  
        control.Text = LANG.FILTER
        model.FontWeight = 200.0
        fenster_cont.addControl('Titel', control)

        control, model = self.mb.createControl(self.mb.ctx,"Button",160,y-3,100,20,(),() )  
        control.Label = LANG.ALLE_WAEHLEN
        control.ActionCommand = 'alle_waehlen'
        control.addActionListener(f_listener)
        fenster_cont.addControl('Alle', control)  

        y += 30
        
        filter_name = sorted(filters)
        
        for name in filter_name:
            
            control, model = self.mb.createControl(self.mb.ctx,"CheckBox",20,y,400,22,(),() )  
            model.Label = filters[name][0]             
            control.ActionCommand = name
            
            control.addActionListener(f_listener)
            fenster_cont.addControl(name, control)
              
  
            if name in imp_set['filterauswahl'].keys():
                control.State = imp_set['filterauswahl'][name]
            else:
                imp_set['filterauswahl'].update({name:0})
            buttons.update({name:control})
 
            y += 16
        
      
    def disposing(self,ev):
        return False
    
    def get_flags(self,x):
        if self.mb.debug: log(inspect.stack)
        
        # Nur fuer OO
        x_bin_rev = bin(x).split('0b')[1][::-1]

        flags = []
        
        for i in range(len(x_bin_rev)):
            z = 2**int(i)*int(x_bin_rev[i])
        
            if z != 0:
                flags.append(z)
        
        return flags
    
    
    def erzeuge_filter(self):
        if self.mb.debug: log(inspect.stack)
        
        # Nur fuer OO
        typeDet = self.mb.createUnoService("com.sun.star.document.TypeDetection")
        FF = self.mb.createUnoService( "com.sun.star.document.FilterFactory" )
        FilterNames = FF.getElementNames()
        
        # da in OO und LO die Positionen der Eintraege im MediaDescriptor verschieden sind,
        # wird zuerst nach den richtigen Positionen gesucht
    
        fil = FF.getByName(FilterNames[0])

        i = 0
        
        for fi in fil:
            if fi.Name == 'UIName':
                uiName_pos = i
            elif fi.Name == 'Name':
                name_pos = i
            elif fi.Name == 'Type':
                type_pos = i
            elif fi.Name == 'DocumentService':
                docSer_pos = i
            elif fi.Name == 'Flags':
                flags_pos = i
                
            i += 1          

        
        def formatiere(term):
            term = term.replace("'","")
            term = term.replace('[','')
            term = term.replace(']','')
            term = term.replace(", "," (",1)
            term = term.replace(",","")
            term = term +')'
            return term
        
        def formatiere2(term):
            ter = []
            for t in term:
                ter.append('*.'+str(t).replace("'",""))
            return list(ter)
        
        
        self.filters_import = {}   
        self.filters_export = {}    
        
        
        cont = self.mb.doc.getControllers()
        for filt in FilterNames:
            
            if filt in self.mb.class_Import.auszuschliessende_filter:
                continue
            
            f = FF.getByName(filt)

            if f[docSer_pos].Value == 'com.sun.star.text.TextDocument':
                flags = self.get_flags(f[flags_pos].Value)

                if 1 in flags or 2 in flags:
                    uiName = f[uiName_pos].Value  
                    filter_typeDet = typeDet.getByName(f[type_pos].Value)
                    
                    # in LO und OO gleich
                    extensions = filter_typeDet[5].Value
                    label2 = formatiere2(extensions)
                    label2.insert(0,str(uiName))
                    
                    if 1 in flags:
                        # FilterName: Filter als Label, Extensions Endungen
                        self.filters_import.update({filt:(formatiere(str(label2)),extensions)})
                    
                    if 2 in flags:                        
                        # FilterName: Filter als Label, Extensions Endungen
                        self.filters_export.update({filt:(formatiere(str(label2)),extensions)})

        
       
class Filter_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,buttons,fenster):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.buttons = buttons
        self.fenster = fenster
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp  
        imp_set[ev.ActionCommand] = self.buttons[ev.ActionCommand].State
        
        if ev.ActionCommand == 'auswahl':
            if ev.Source.State == True:
                self.mb.nachricht(LANG.IMPORT_FILTER_WARNUNG,"warningbox")
                self.fenster.toFront()
            for but in self.buttons:
                if but != 'auswahl':
                    self.buttons[but].Enable = not (self.buttons['auswahl'].State)
                    
        self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  
        
    def disposing(self,ev):
        return False
        
class Filter_CheckBox_Listener2(unohelper.Base, XActionListener):
    def __init__(self,mb,buttons):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.buttons = buttons
        self.state = False
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp  
        cmd = ev.ActionCommand

        if cmd == 'alle_waehlen':
            self.state = not self.state
            for but in self.buttons:
                self.buttons[but].State = self.state
                imp_set['filterauswahl'][but] = self.state
        else:
            imp_set['filterauswahl'][cmd] = self.buttons[cmd].State   
                                 
        self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  

    def disposing(self,ev):
        return False

class Auswahl_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,model1,model2):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.model1 = model1
        self.model2 = model2
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        imp_set = self.mb.settings_imp 

        if ev.ActionCommand == 'Datei':
        
            filepath,ok = self.mb.class_Funktionen.filepicker2()
            
            if not ok:
                return
            
            self.model1.Label = uno.fileUrlToSystemPath(filepath)
            imp_set['url_dat'] = filepath
            
            self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  

        
        elif ev.ActionCommand == 'Ordner':
        
            Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
            if imp_set['url_ord'] != None:
                Filepicker.setDisplayDirectory(imp_set['url_ord'])
            Filepicker.execute()
         
            if Filepicker.Directory == '':
                return
             
            filepath = Filepicker.Directory
            self.model2.Label = uno.fileUrlToSystemPath(filepath)
            imp_set['url_ord'] = filepath
            
            self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  

    def disposing(self,ev):
        return False
           
