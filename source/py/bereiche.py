# -*- coding: utf-8 -*-

import unohelper
from com.sun.star.style.BreakType import NONE as BREAKTYPE_NONE

class Bereiche():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)

        # Konstanten
        self.mb = mb
        self.doc = mb.doc       
        
        # Klassen
        self.oOO = None
    
    
    def starte_oOO(self,URL=None):
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
                
        Path1 = os.path.join(self.mb.pfade['odts'] , 'nr%s.odt' %nr )
        Path2 = uno.systemPathToFileUrl(Path1)        
        self.oOO.storeToURL(Path2,())
        
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
            #inhalt = 'nr. ' + nr + '\t' + inhalt
            
            cursor = text.createTextCursor()
            #cursor = self.oOO.CurrentController.ViewCursor
            cursor.gotoStart(False)
            cursor.gotoEnd(True)

            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            newSection.setName('OrgInnerSec'+nr)
            text.insertTextContent(cursor, newSection, False)
            
            cursor.goLeft(2,True)
            cursor.collapseToStart()
            
            if self.mb.debug:
                text.insertString( cursor, inhalt, True )
            else:
                text.insertString( cursor, ' ', True )
            
            cursor.collapseToEnd()
            cursor.goRight(2,True)
            cursor.setString('')
            
            Path1 = os.path.join(self.mb.pfade['odts'] , 'nr%s.odt' %nr )
            Path2 = uno.systemPathToFileUrl(Path1)    
    
            self.oOO.storeToURL(Path2,())
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
        #cur.setPropertyValue('BreakType',6)
        leer = ' ' * 60 * 30 * 400
        leer = 'hallo'
        #text.setString(leer)
        
        from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
        for i in range(100):
            text.insertControlCharacter( cur, PARAGRAPH_BREAK, 0 )
            text.insertControlCharacter( cur, PARAGRAPH_BREAK, 0 )
        
#         enum = text.createEnumeration()
#         paras = []
#         while enum.hasMoreElements():
#             paras.append(enum.nextElement())
            
#         for par in paras:
#             cur.gotoRange(par.Start,False)
#             cur.setPropertyValue('BreakType',4)
        
        url = os.path.join(self.mb.pfade['odts'], 'empty_file.odt')
        dokument.storeToURL(uno.systemPathToFileUrl(url),())
        dokument.close(False)
        
    def leere_Dokument(self):
        if self.mb.debug: log(inspect.stack)
        
        text = self.doc.Text        
        try:
            all_sections_Namen = self.doc.TextSections.ElementNames
            if self.doc.TextSections.Count != 0:
                for name in all_sections_Namen:
                    sec = self.doc.TextSections.getByName(name)
                    sec.dispose()
    
            
            cursor = self.mb.viewcursor
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            cursor.setString('')
        except:
            log(inspect.stack,tb())
            

                     
    def erzeuge_bereich(self,i,path,sicht,papierkorb=False):
        if self.mb.debug: log(inspect.stack)
        
        nr = str(i) 
        
        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        
        if papierkorb:
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
            SFLink.FileURL = path
            SFLink.FilterName = 'writer8'
            newSection.setPropertyValue('FileLink',SFLink)
            
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

        bereichsname_Papierkorb = self.mb.props[T.AB].dict_bereiche['ordinal'][self.mb.props[T.AB].Papierkorb]
        path_to_Papierkorb = self.mb.props[T.AB].dict_bereiche['Bereichsname'][bereichsname_Papierkorb]
        
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
        sections = self.doc.TextSections
        cur = text.createTextCursor()
        
        cur.gotoEnd(False)
        cur.goLeft(1,True)
        cur.setString('')
       
        
    def datei_nach_aenderung_speichern(self,zu_speicherndes_doc_path,bereichsname = None):

        if self.mb.props[T.AB].tastatureingabe == True and bereichsname != None:
            # Nur loggen, falls tatsaechlich gespeichert wurde
            if self.mb.debug: log(inspect.stack)

            try:
                # Damit das Handbuch nicht geaendert wird:
                if self.mb.anleitung_geladen:
                    return
                                
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
                newDoc.storeToURL(zu_speicherndes_doc_path,())
                newDoc.close(False)
    
                self.mb.props[T.AB].tastatureingabe = False
                self.mb.loesche_undo_Aktionen()
            except:
                log(inspect.stack,tb())


    def verlinkte_Bilder_einbetten(self,doc):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.selbstruf = True
        
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
                
        self.mb.selbstruf = False   
        
        
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

