# -*- coding: utf-8 -*-

import traceback
import uno
import unohelper
from threading import Thread
import time

tb = traceback.print_exc

FARBE_AUSGEWAEHLTE_ZEILE = 14866636 #creme 16777075 # hellgelb
FARBE_ZEILE_STANDARD = 15790320 #LOO grau 16777215 # weiss  



class Bereiche():
    
    def __init__(self,mb):

        global pd
        pd = mb.pd
             
        # Konstanten
        self.mb = mb
        self.doc = mb.doc       
        self.viewcursor = mb.viewcursor
        
        # Klassen
        self.oOO = None
    
    
    def starte_oOO(self,URL="private:factory/swriter"):
        if self.mb.debug: print(self.mb.debug_time(),'starte_oOO') 
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        if self.mb.debug: print(self.mb.debug_time(),'inside oOO')
        
        self.oOO = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,))
        
        if self.mb.debug: print(self.mb.debug_time(),'oOO geladen')
        
    def schliesse_oOO(self):
        if self.mb.debug: print(self.mb.debug_time(),'schliesse_oOO')
        self.oOO.close(False)
        
    
    def erzeuge_neue_Datei(self,i,inhalt):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_neue_Datei')
        
        nr = str(i) 

        text = self.oOO.Text
        
        inhalt = 'nr. ' + nr + '\t' + inhalt
        
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)

        text.insertString( cursor, inhalt, True )
        Path = 'file:///' + self.mb.pfade['odts'] + '/nr%s.odt' %nr         
        self.oOO.storeToURL(Path,())
        
        return Path
    
    def erzeuge_neue_Datei2(self,i,inhalt):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_neue_Datei2')
 
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        
        URL="private:factory/swriter"
        self.oOO = self.mb.desktop.loadComponentFromURL(URL,'_blank',8+32,(prop,))
        
        nr = str(i) 
        
        text = self.oOO.Text
        inhalt = 'nr. ' + nr + '\t' + inhalt
        
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)

        text.insertString( cursor, inhalt, True )
        
        Path = 'file:///' + self.mb.pfade['odts'] + '/nr%s.odt' %nr  

        self.oOO.storeToURL(Path,())
        self.oOO.close(False)
    
    
    def speichern(self,args):
        Path,oOO2 = args
        if self.mb.debug: print(self.mb.debug_time(),'vorm speichern')
        oOO2.storeToURL(Path,())
        #self.oOO2.close(False)
        if self.mb.debug: print(self.mb.debug_time(),'nachm speichern')
        
    def leere_Dokument(self):
        if self.mb.debug: print(self.mb.debug_time(),'leere_Dokument')
        
        text = self.doc.Text        
        try:
            all_sections_Namen = self.doc.TextSections.ElementNames
            for name in all_sections_Namen:
                sec = self.doc.TextSections.getByName(name)
                sec.dispose()
    
            inhalt = ''
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            text.insertString( cursor, inhalt, True )
        except:
            print('Dokument ist schon leer')
                     
    def erzeuge_bereich(self,i,path,sicht):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_bereich')
        
        nr = str(i) 
        
        text = self.mb.doc.Text
        sections = self.mb.doc.TextSections  
                
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = path
        SFLink.FilterName = 'nr'+nr
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
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
        
    
    def erzeuge_bereich2(self,i,path,sicht):
        if self.mb.debug: print(self.mb.debug_time(),'erzeuge_bereich2')
        
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
        SFLink.FilterName = 'nr'+nr
        
        newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
        newSection.setPropertyValue('FileLink',SFLink)
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
        
    
        
    def loesche_leeren_Textbereich_am_Ende(self):
        if self.mb.debug: print(self.mb.debug_time(),'loesche_leeren_Textbereich_am_Ende')
        
        text = self.doc.Text
        sections = self.doc.TextSections
        cur = text.createTextCursor()
        
        cur.gotoEnd(False)
        cur.goLeft(1,True)
        cur.setString('')
       
                       
    def verlinke_bereiche(self,quellbereich_name,zielbereich_name):
        if self.mb.debug: print(self.mb.debug_time(),'verlinke_bereiche')
        
        sections = self.doc.TextSections
        
        for nr in range(sections.Count):            
            sec = sections.getByIndex(nr)
            print(sec.FileLink.FileURL)
                
        link_quelle = self.mb.dict_bereiche['Bereichsname'][quellbereich_name]# <- der Link
        zielbereich = sections.getByName(zielbereich_name)

        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink.FileURL = link_quelle
        SFLink.FilterName = 'filter'
        
        zielbereich.setPropertyValue('FileLink',SFLink)
        
    def datei_nach_aenderung_speichern(self,zu_speicherndes_doc_path,bereichsname = None):
        
        if self.mb.tastatureingabe == True and bereichsname != None:
            if self.mb.debug: print(self.mb.debug_time(),'datei_nach_aenderung_speichern')
            
            self.pruefe_auf_unverlinkte_bilder()
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

            self.mb.tastatureingabe = False

    def pruefe_auf_unverlinkte_bilder(self):
        if self.mb.debug: print(self.mb.debug_time(),'pruefe_auf_unverlinkte_bilder')
        self.mb.selbstruf = True
        
        bilder = self.mb.doc.GraphicObjects
        bitmap = self.mb.doc.createInstance( "com.sun.star.drawing.BitmapTable" )
        
        for i in range(bilder.Count):
            bild = bilder.getByIndex(i)
            if 'vnd.sun.star.GraphicObject' not in bild.GraphicURL:
                # Wenn die Grafik nicht gespeichert werden kann,
                # (weil noch nicht geladen oder User hat zu früh weitergeklickt)
                # müsste sie eigentlich bei der nächsten Änderung
                # im Bereich gespeichert werden
                try:
                    bitmap.insertByName( "TempI"+str(i), bild.GraphicURL )
                    print(bild.GraphicURL)
                    try:
                        internalUrl = bitmap.getByName( "TempI"+str(i) ) 
                        bild.GraphicURL = internalUrl 
                    except:
                        pass
                    bitmap.removeByName( "TempI"+str(i) ) 
                except:
                    print('insert unverlinkte bilder gescheitert')
                
        self.mb.selbstruf = False      

from com.sun.star.view import XSelectionChangeListener
class ViewCursor_Selection_Listener(unohelper.Base, XSelectionChangeListener):
    
    def __init__(self,mb):
        self.mb = mb
        self.text_section_old = 'nicht vorhanden'
        self.mb.selbstruf = False
        
    def disposing(self,ev):
        return False
    
    def selectionChanged(self,ev):
        
        if self.mb.selbstruf:
            print('selection selbstruf')
            return
        print('selection')

        selected_text_section = self.mb.current_Contr.ViewCursor.TextSection
        if selected_text_section == None:
            return False
        
        s_name = selected_text_section.Name
        
        # stellt sicher, dass nur selbst erzeugte Bereiche angesprochen werden
        # und der Trenner übersprungen wird
        if 'trenner'  in s_name:
            if self.mb.zuletzt_gedrueckte_taste == None:
                try:
                    self.mb.viewcursor.goDown(1,False)
                except:
                    self.mb.viewcursor.goUp(1,False)
                return False
            # 1024,1027 Pfeil runter,rechts
            elif self.mb.zuletzt_gedrueckte_taste.KeyCode in (1024,1027):  
                self.mb.viewcursor.goDown(1,False)       
            else:
                self.mb.viewcursor.goUp(1,False)
            # sollte der viewcursor immer noch auf einem Trenner stehen,
            # befindet er sich im letzten Bereich -> goUp    
            if 'trenner' in self.mb.viewcursor.TextSection.Name:
                    self.mb.viewcursor.goUp(1,False)
            return False 
        
        # test ob ausgewählter Bereich ein Kind-Bereich ist -> Selektion wird auf Parent gesetzt
        elif 'trenner' not in s_name and 'OrganonSec' not in s_name:
            sec = []
            self.test_for_parent_section(selected_text_section,sec)
            selected_text_section = sec[0]
            
        self.so_name =  None   
             
        if self.mb.selektierte_Zeile_alt != None:
            text_section_old_ordinal = self.mb.selektierte_Zeile_alt.AccessibleContext.AccessibleParent.AccessibleContext.AccessibleName
            text_section_old_bereichsname = self.mb.dict_bereiche['ordinal'][text_section_old_ordinal]
            self.text_section_old = self.mb.doc.TextSections.getByName(text_section_old_bereichsname)            
            self.so_name = self.mb.dict_bereiche['ordinal'][text_section_old_ordinal]
    
        if self.text_section_old == 'nicht vorhanden':
            #print('selek gewechs, old nicht vorhanden')
            self.mb.bereich_wurde_bearbeitet = True
            self.text_section_old = selected_text_section 
            return False 
        elif self.mb.Papierkorb_geleert == True:
            #print('selek gewechs, Papierkorb_geleert')
            # fehlt: nur speichern, wenn die Datei nicht im Papierkorb gelandet ist
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.text_section_old.FileLink.FileURL,self.so_name)
            self.mb.bereich_wurde_bearbeitet = True
            self.text_section_old = selected_text_section 
            self.mb.Papierkorb_geleert = False 
            return False       
        else:
            if self.text_section_old == selected_text_section:
                #print('selek nix gewechs',self.so_name , s_name)
                self.mb.bereich_wurde_bearbeitet = True
                return False                
            else:
                #print('selek gewechs',self.so_name , s_name)
                self.mb.class_Bereiche.datei_nach_aenderung_speichern(self.text_section_old.FileLink.FileURL,self.so_name)
                self.mb.bereich_wurde_bearbeitet = True
                self.text_section_old = selected_text_section 
                self.farbe_der_selektion_aendern(selected_text_section.Name)
                return False 
                
   
    def test_for_parent_section(self,selected_text_sectionX,sec):
        if selected_text_sectionX.ParentSection != None:
            selected_text_sectionX = selected_text_sectionX.ParentSection
            self.test_for_parent_section(selected_text_sectionX,sec)
        else:
            sec.append(selected_text_sectionX)
        
                  
    def farbe_der_selektion_aendern(self,bereichsname):      

        ordinal = self.mb.dict_bereiche['Bereichsname-ordinal'][bereichsname]
        zeile = self.mb.Hauptfeld.getControl(ordinal)
        textfeld = zeile.getControl('textfeld')
        
        self.mb.selektierte_zeile = zeile.AccessibleContext
        
        # selektierte Zeile einfärben, ehem. sel. Zeile zurücksetzen
        textfeld.Model.BackgroundColor = FARBE_AUSGEWAEHLTE_ZEILE 
        if self.mb.selektierte_Zeile_alt != None:                
            self.mb.selektierte_Zeile_alt.Model.BackgroundColor = FARBE_ZEILE_STANDARD

        self.mb.selektierte_Zeile_alt = textfeld
  
     

from com.sun.star.awt import XWindowListener
class Dialog_Window_Listener(unohelper.Base,XWindowListener):
    
    def __init__(self,mb):
        self.mb = mb
        #print('window listener init')
    def windowResized(self,ev):
        self.korrigiere_hoehe_des_scrollbalkens()
        #print('windowResized')
        return False
    def windowMoved(self,ev):
        #print('windowMoved')
        return False
    def windowShown(self,ev):
        self.korrigiere_hoehe_des_scrollbalkens()
        #print('windowShown')
        return False
    def windowHidden(self,ev):
        #pd()
        if self.mb.bereich_wurde_bearbeitet:
            ordinal = self.mb.selektierte_zeile.AccessibleName
            bereichsname = self.mb.dict_bereiche['ordinal'][ordinal]
            path = self.mb.dict_bereiche['Bereichsname'][bereichsname]
            self.mb.class_Bereiche.datei_nach_aenderung_speichern(path,bereichsname)
        self.mb.entferne_alle_listener() 
        #print('windowHidden')
        return False
    
    def korrigiere_hoehe_des_scrollbalkens(self):
       
        nav_cont_aussen = self.mb.dialog.getControl('Hauptfeld_aussen')
        nav_cont = nav_cont_aussen.getControl('Hauptfeld')
          
        MenuBar = self.mb.dialog.getControl('Organon_Menu_Bar')
        MBHoehe = MenuBar.PosSize.value.Height + MenuBar.PosSize.value.Y
        NCHoehe = 0 #nav_cont.PosSize.value.Height
        NCPosY  = nav_cont.PosSize.value.Y
        Y =  NCHoehe + NCPosY + MBHoehe
        Height = self.mb.dialog.PosSize.value.Height - Y -25
        
        scrll = self.mb.dialog.getControl('ScrollBar')
        scrll.setPosSize(0,0,0,Height,8)

        