# -*- coding: utf-8 -*-

from com.sun.star.style.BreakType import NONE as BREAKTYPE_NONE
from com.sun.star.beans import UnknownPropertyException

class Bereiche():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)

        self.mb = mb
        self.doc = mb.doc       
        self.oOO = None
        
        self.services = {
                        'frame' : 'com.sun.star.text.TextFrame',
                        'table' : 'com.sun.star.text.CellProperties',
                        'fn'    : 'com.sun.star.text.Footnote',
                        'en'    : 'com.sun.star.text.Endnote',
                        }
    
    def starte_oOO(self, URL=None):
        if self.mb.debug: log(inspect.stack)
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        
        prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop2.Name = 'AsTemplate'
        prop2.Value = True
                
        if URL == None:
            URL="private:factory/swriter"
            if self.mb.settings_proj['use_template'][0] == True:
                if self.mb.settings_proj['use_template'][1] != '':
                    URL = uno.systemPathToFileUrl(self.mb.settings_proj['use_template'][1])
                
        self.oOO = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,prop2))
        
        
    def schliesse_oOO(self):
        if self.mb.debug: log(inspect.stack)
        self.oOO.close(False)
        
    
    def erzeuge_neue_Datei(self,i,inhalt):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 

        text = self.oOO.Text
        
        inhalt = 'nr. ' + nr + '\t' + inhalt
        
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        text.insertString( cursor, '', True )

        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        newSection.setName('OrgInnerSec'+nr)
        text.insertTextContent(cursor, newSection, False)
        cursor.goLeft(1,True)
        
        text.insertString( cursor, inhalt, True )
                
        Path1 = os.path.join(self.mb.pfade['odts'] , 'nr{}.odt'.format(nr) )
        Path2 = uno.systemPathToFileUrl(Path1)      
          
        self.oOO.storeToURL(Path2,())
        self.plain_txt_speichern(inhalt, 'nr{}'.format(nr))
        
        newSection.dispose()
        
        return Path1
    
    
    def erzeuge_neue_Datei2(self,i,inhalt):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
            
            URL="private:factory/swriter"
    
            if self.mb.settings_proj['use_template'][0] == True:
                URL = uno.systemPathToFileUrl(self.mb.settings_proj['use_template'][1])
            self.oOO = self.mb.desktop.loadComponentFromURL(URL,'_blank',8+32,(prop,))
            
            nr = str(i) 
            
            text = self.oOO.Text
            
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)

            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection.setName('OrgInnerSec'+nr)
            text.insertTextContent(cursor, newSection, False)
            
            cursor.goLeft(2,True)
            cursor.collapseToStart()
            
            if not self.mb.debug:
                inhalt = ' '
                
            text.insertString( cursor, inhalt, True )
            
            cursor.collapseToEnd()
            cursor.goRight(2,True)
            cursor.setString('')
            
            Path1 = os.path.join(self.mb.pfade['odts'] , 'nr%s.odt' %nr )
            Path2 = uno.systemPathToFileUrl(Path1)    
    
            self.oOO.storeToURL(Path2,())
            self.plain_txt_speichern(inhalt, 'nr{}'.format(nr))
            
            self.oOO.close(False)

        except:
            log(inspect.stack,tb())
    
    
    def erzeuge_leere_datei(self):
        if self.mb.debug: log(inspect.stack)
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        
        URL="private:factory/swriter"
        
        dokument = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,))
        
        text = dokument.Text
        cur = text.createTextCursor()
        cur.gotoStart(False)
        
        from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
        for i in range(100):
            text.insertControlCharacter( cur, PARAGRAPH_BREAK, 0 )
            text.insertControlCharacter( cur, PARAGRAPH_BREAK, 0 )

        url = os.path.join(self.mb.pfade['odts'], 'empty_file.odt')
        dokument.storeToURL(uno.systemPathToFileUrl(url),())
        dokument.close(False)
        
        
    def leere_Dokument(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            all_sections_Namen = self.doc.TextSections.ElementNames
            if self.doc.TextSections.Count != 0:
                for name in all_sections_Namen:
                    sec = self.doc.TextSections.getByName(name)
                    sec.setPropertyValue('IsProtected',False)
                    sec.dispose()

            cursor = self.mb.viewcursor
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            cursor.setString('')

        except:
            log(inspect.stack,tb())
            
    
    def erzeuge_bereich_papierkorb(self,i,path):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 
        
        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = path
        SFLink.FilterName = 'writer8'
        newSection.setPropertyValue('FileLink',SFLink)
            
        newSection.setName('OrganonSec'+nr)
        
        sectionN = sections.getByIndex(sections.Count-1)
        textSectionCursor = text.createTextCursorByRange(sectionN.Anchor)
        textSectionCursor.gotoEnd(False)
        
        text.insertTextContent(textSectionCursor, newSection, False)
    
                     
    def erzeuge_bereich(self,i,path,sicht):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 
        
        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            
        newSection.setName('OrganonSec'+nr)
        
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
        
    
    def erzeuge_bereich2(self,i,sicht):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 

        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
        
        alle_hotzenploetze = []
        for sec_name in sections.ElementNames:
            if 'OrganonSec' in sec_name:
                alle_hotzenploetze.append(sec_name)
        anzahl_hotzenploetze = len(alle_hotzenploetze)
        
        path = os.path.join(self.mb.pfade['odts'] , 'nr%s.odt' %nr) 
        path = uno.systemPathToFileUrl(path)
                 
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = path#_to_Papierkorb
        SFLink.FilterName = 'writer8'

        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        newSection.setPropertyValue('FileLink',SFLink)
        newSection.setName('OrganonSec'+str(anzahl_hotzenploetze))

        if sicht == 'nein':
            newSection.IsVisible = False

        textSectionCursor = text.createTextCursor()
        textSectionCursor.gotoEnd(False)

        text.insertTextContent(textSectionCursor, newSection, False)

        # Der richtige Link fuer den letzten Bereich wird erst hier gesetzt, da die sec ansonsten
        # falsch eingefuegt wird (weiss der Henker, warum)
        path_to_empty = uno.systemPathToFileUrl(os.path.join(self.mb.pfade['odts'],'empty_file.odt'))
        SFLink.FileURL = path_to_empty
        newSection.setPropertyValue('FileLink',SFLink)
        
        newSection.Anchor.BreakType = BREAKTYPE_NONE


    def loesche_leeren_Textbereich_am_Ende(self):
        if self.mb.debug: log(inspect.stack)
        
        text = self.doc.Text
        cur = text.createTextCursor()
        
        cur.gotoEnd(False)
        cur.goLeft(1,True)
        cur.setString('')
    
    def datei_speichern(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
        self.datei_nach_aenderung_speichern(bereichsname,speichern = True)        
        
        
    def datei_nach_aenderung_speichern(self, bereichsname = None, anderer_pfad = None, speichern = False):
        
        if bereichsname == None:
            ordinal = self.mb.props[T.AB].selektierte_zeile
            bereichsname = self.mb.props[T.AB].dict_bereiche['ordinal'][ordinal]
        
        props = self.mb.props[T.AB]
        if props.dict_bereiche['Bereichsname-ordinal'][bereichsname] == props.Papierkorb:
            # Papierkorb wird nicht gespeichert
            return
        
        if (len(self.mb.undo_mgr.AllUndoActionTitles) > 0 and bereichsname != None) or speichern:
            # Nur loggen, falls tatsaechlich gespeichert wurde
            if self.mb.debug: log(inspect.stack)

            try:
                # Damit das Handbuch nicht geaendert wird:
                if self.mb.anleitung_geladen:
                    return
                
                if anderer_pfad:
                    pfad = anderer_pfad
                else:
                    pfad = uno.systemPathToFileUrl(self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname])
                                
                self.verlinkte_Bilder_einbetten(self.mb.doc)
                projekt_path = self.mb.doc.URL
                
                prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
                prop.Name = 'Hidden'
                prop.Value = True
                
                newDoc =  self.mb.desktop.loadComponentFromURL("private:factory/swriter",'_blank',8+32,(prop,))
                cur = newDoc.Text.createTextCursor()
                cur.gotoStart(False)
                cur.gotoEnd(True)
    
                SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
                SFLink.FileURL = projekt_path
    
                newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                newSection.setPropertyValues(("LinkRegion",'FileLink'),(bereichsname,SFLink))
                                    
                newDoc.Text.insertTextContent(cur, newSection, True)
                newDoc.Text.removeTextContent(newSection)
                newDoc.storeToURL(pfad,())
                
                if '.odthelfer' not in pfad:
                    # plain_txt speichern
                    plain_txt = newDoc.Text.String
                    dict_bereiche = self.mb.props[T.AB].dict_bereiche
                    
                    os_path = uno.fileUrlToSystemPath(pfad)
                    bereich = [ n for n,p in dict_bereiche['Bereichsname'].items() if p == os_path ][0]
                    ordinal = dict_bereiche['Bereichsname-ordinal'][bereich]
                    
                    self.plain_txt_speichern(plain_txt, ordinal)

                newDoc.close(False)
    
                self.mb.loesche_undo_Aktionen()
            except:
                log(inspect.stack,tb())
                Popup(self.mb, 'warning').text = LANG.DATEI_NICHT_GESPEICHERT
                try:
                    newDoc.close(False)
                except:
                    pass
        
    
    def plain_txt_speichern(self,plain_txt,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad_plain_txt = os.path.join(self.mb.pfade['plain_txt'], ordinal + '.txt')
    
            with codecs_open(pfad_plain_txt , "w","utf-8") as f:
                f.write(plain_txt)
        except:
            log(inspect.stack,tb())
            
    
    def plain_txt_loeschen(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad_plain_txt = os.path.join(self.mb.pfade['plain_txt'], ordinal + '.txt')
            os.remove(pfad_plain_txt)
        except:
            log(inspect.stack,tb())
        

    def verlinkte_Bilder_einbetten(self,doc):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.Listener.VC_selection_listener.selbstruf = True
        
        try:
            bilder = self.mb.doc.GraphicObjects
            bitmap = self.mb.doc.createInstance( "com.sun.star.drawing.BitmapTable" )
            
            for i in range(bilder.Count):
                bild = bilder.getByIndex(i)
                if 'vnd.sun.star.GraphicObject' not in bild.GraphicURL:
                    # Wenn die Grafik nicht gespeichert werden kann,
                    # (weil noch nicht geladen oder User hat zu frueh weitergeklickt)
                    # muesste sie eigentlich bei der naechsten Aenderung
                    # im Bereich gespeichert werden
                    try:
                        bitmap.insertByName( "TempI"+str(i), bild.GraphicURL )
                        #if self.mb.debug: print(bild.GraphicURL)
                        try:
                            internalUrl = bitmap.getByName( "TempI"+str(i) ) 
                            bild.GraphicURL = internalUrl 
                        except:
                            pass
                        bitmap.removeByName( "TempI"+str(i) ) 
                    except:
                        log(inspect.stack,tb())
        except:
            log(inspect.stack,tb())
        self.mb.Listener.VC_selection_listener.selbstruf = False   
        
        
    def speicher_odt(self,pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True        
        
            prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop2.Name = 'FilterName'
            prop2.Value = "HTML (StarWriter)"
    
            pfad = uno.systemPathToFileUrl(os.path.join(self.speicherordner,self.titel+'.html'))
            doc = self.doc.CurrentController.Frame.loadComponentFromURL(pfad,'_blank',0,(prop,prop2))

            prop1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop1.Name = 'Overwrite'
            prop1.Value = True
            
            pfad = uno.systemPathToFileUrl(os.path.join(self.speicherordner,self.titel+'.odt'))
            doc.storeToURL(pfad,(prop1,))
            doc.close(False)
            
        except:
            log(inspect.stack,tb())


    def speicher_pdf(self):
        if self.mb.debug: log(inspect.stack)
        
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True        
    
        prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop2.Name = 'FilterName'
        prop2.Value = "HTML (StarWriter)"

        pfad = uno.systemPathToFileUrl(os.path.join(self.speicherordner,self.titel+'.html'))
        doc = self.doc.CurrentController.Frame.loadComponentFromURL(pfad,'_blank',0,(prop,prop2))

        prop1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop1.Name = 'Overwrite'
        prop1.Value = True
        
        
        prop3 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop3.Name = 'FilterName'
        prop3.Value = 'writer_pdf_Export'
        
        
        pfad = uno.systemPathToFileUrl(os.path.join(self.speicherordner,self.titel+'.pdf'))
        doc.storeToURL(pfad,(prop1,prop3))
        doc.close(False)

    
    def get_ordinal(self, text_range, is_in_frame=None):
        #if self.mb.debug: log(inspect.stack)
        
        try:
            if is_in_frame:
                sec, sec_name, info = self.get_organon_section(text_range, ['frame',])
            else:
                sec, sec_name = self.get_organon_section(text_range)
                
            ordinal = 'nr{}'.format(sec.Name.split('OrgInnerSec')[1])
            
            if is_in_frame:
                return ordinal, info['frame']
            else:
                return ordinal
        except:
            log(inspect.stack,tb())
            return None
    
    
    def get_organon_section(self, text_range, info=[]):
        #if self.mb.debug: log(inspect.stack)
                
        try:
            info2 = {i:False for i in info}
            
            def get_info(txt_inst):
                for i in info:
                    if self.services[i] in txt_inst.SupportedServiceNames:
                        info2[i] = True
                        return

            def get_parsec(s):
                try:
                    while 'OrgInnerSec' not in s.Name:
                        if s.ParentSection != None:
                            s = s.ParentSection
                        else:
                            s = s.Anchor.Text.Anchor.TextSection
                    return s
                except:
                    log(inspect.stack,tb())
                    return None, ''
            
            par_sec = None  
              
            if text_range.TextSection != None:
                sec = text_range.TextSection
                
                if 'trenner' in sec.Name:
                    return sec, sec.Name
                else:
                    par_sec = get_parsec(sec)
                    if info: get_info(text_range.Text) 

            else:
                # Fuss- oder Endnote, TextFrame
                try:
                    sec = text_range.Text.Anchor.TextSection
                    par_sec = get_parsec(sec)
                    if info: get_info(text_range.Text)

                except UnknownPropertyException:
                    # TextFrame
                    # Je nach Verschachtelung gibt es mehrere Moeglichkeiten
                    sec = text_range.TextFrame.Anchor.TextSection
                    if sec == None:
                        sec = text_range.TextFrame.Anchor.TextFrame.Anchor.TextSection
                    if sec != None:
                        par_sec = get_parsec(sec)
                        if info: get_info(text_range.TextFrame.Text)
                    
            if info:
                return par_sec, par_sec.Name, info2
            
            return par_sec, par_sec.Name
        except:
            log(inspect.stack,tb())
            return None, ''
        
            
            
            
            
            
            
            
            
            
            
           