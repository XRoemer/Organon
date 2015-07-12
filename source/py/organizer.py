# -*- coding: utf-8 -*-

import unohelper

class Organizer():
    '''
    Es fehlt:
    - keine zwei Instanzen des Organizers oeffnen
    - beim Schließen von Calc evt. User Icons loeschen
    ( modul funktionen: Tag2_Item_Listener.galerie_icon_im_prj_ordner_evt_loeschen() )
    - Dateien umbenennen
    - Tags Zeit/Datum
    - Tags Allgemein
    - Einfuegen von Dateien/Ordnern und
    - Verschieben (dann muesste allerdings die Baumstruktur abgebildet werden. Evt. zu kompliziert)
    - Loeschen (in den Papierkorb verschieben)
    - Optionen: anzuzeigende Tags und deren Breite einstellen
    '''
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.dict_sb_content = None
        self.calc = None
        self.sheet = None
        self.sheet_controller = None
        self.calc_frame = None
        
        self.btn = []
        self.data_array = []
        self.first_time_info = True
        
        self.spalten = [
           'Synopsis',
           'Notes', 
#                        'Images', 
           'Tags_characters', 
           'Tags_locations', 
           'Tags_objects', 
           #'Tags_time', 
           'Tags_user1', 
           'Tags_user2', 
           'Tags_user3']
            
        self.spalten2 = [
                LANG.DATEI,
                LANG.SYNOPSIS,
                LANG.NOTIZEN,
#                     LANG.BILDER,
                LANG.CHARAKTERE,
                LANG.ORTE,
                LANG.OBJEKTE,
                #LANG.ZEIT,
                LANG.BENUTZER1,
                LANG.BENUTZER2,
                LANG.BENUTZER3
                ]
        
        # self.pos bestimmt den Ort der Tabelle
        # Wenn eine Spalte/Zeile mehr benoetigt werden sollte,
        # können die ersten beiden Werte geaendert werden
        self.pos = [ 3, 3, 3 + len(self.spalten2), 3 ]
        
        # Die Color items sollten mal uebersetzt werden!
        self.color_items = (
                'blau',
                'braun',
                'creme',
                'gelb',
                'grau',
                'gruen',
                'hellblau',
                'hellgrau',
                'lila',
                'ocker',
                'orange',
                'pink',
                'rostrot',
                'rot',
                'schwarz',
                'tuerkis',
                'weiss')
        
        
    def run(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if self.mb.programm == 'OpenOffice':
                if self.first_time_info:
                    self.first_time_info = False
                    self.mb.nachricht(LANG.OO_ORGANIZER_INFO,'infobox')
            
            self.dict_sb_content = copy.deepcopy(self.mb.dict_sb_content)
            
            self.oeffne_calc()
            self.get_farben()
            self.setze_eintraege()
            self.erzeuge_ansicht()
            self.erzeuge_steuerung()
            
            self.setze_icons()            
            self.schuetze_document()
            
            handler = Enhanced_MC_Handler(self.mb,self)
            self.sheet_controller.addEnhancedMouseClickHandler(handler)
            
            rangex = self.sheet.getCellRangeByPosition(0,0,0,0)
            self.sheet_controller.select(rangex)
                        
        except:
            log(inspect.stack,tb())

        
    def schuetze_document(self):
        if self.mb.debug: log(inspect.stack)
        
        # Schützen, Aufheben des Schutzes, Schützen wird ausgeführt,
        # damit die Icons auf die richtige Höhe gesetzt werden.
        # Eventuell fehlt hier der richtige Befehl (self.calc.render?)
        self.sheet.protect('')
        
        dispatcher = self.mb.createUnoService("com.sun.star.frame.DispatchHelper") 
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
           
        prop.Name = 'Protect'
        prop.Value = False
        dispatcher.executeDispatch(self.calc_frame, ".uno:Protect", "", 0, (prop,))
        
        self.sheet.protect('')
    
    def get_farben(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_orga['organon_farben']

        farben = ['hf_hintergrund',
                  'schrift_datei',
                  'schrift_ordner',
                  'menu_hintergrund',
                  'menu_schrift']
        
        for f in farben:
            setattr(self, f, sett[f])
            
        if sett['design_office'] and sett['office']['nutze_dok_farbe']:
            self.hintergrund = sett['office']['dok_hintergrund']
        else:
            self.hintergrund = 15790320 # LO grau
        
    
    def erzeuge_button(self,lbl,btn_name,pos,listener):
        if self.mb.debug: log(inspect.stack)
        
        try:
            form = self.calc.createInstance("com.sun.star.form.component.Form")
            form.Name = btn_name
            
            forms = self.sheet.DrawPage.Forms
            if btn_name not in forms.ElementNames:
                forms.insertByIndex (0,form)
            else:
                form = forms.getByName(btn_name)
            
            oButton = self.mb.createUnoService("com.sun.star.form.component.CommandButton")
            oButton.Label = lbl
            
            oButton.Name = btn_name
            oButton.BackgroundColor = self.menu_hintergrund
            oButton.TextColor = self.menu_schrift
            
            form.insertByName(btn_name,oButton)
            
            shape = self.calc.createInstance("com.sun.star.drawing.ControlShape")
            
            cell = self.sheet.getCellByPosition(3,1)
            
            p = shape.Position
            p.X = cell.Position.X + pos * 3500
            p.Y = cell.Position.Y
            shape.setPosition(p)
            
            s = shape.Size
            s.Height = 750
            s.Width = 3000
            shape.setSize(s)
            
            shape.Control = oButton
            oButton.Toggle = 1
            
            oButton.addPropertiesChangeListener(('State',),listener)

            return shape
        except:
            log(inspect.stack,tb())
            
            
            
    def erzeuge_img_button(self,pos,btn_name,url,listener):
        if self.mb.debug: log(inspect.stack)
        
        try:
            form = self.calc.createInstance("com.sun.star.form.component.Form")
            form.Name = btn_name
            
            forms = self.sheet.DrawPage.Forms
            forms.insertByIndex (0,form)
            
            oButton = self.mb.createUnoService("com.sun.star.form.component.ImageButton")
            
            if url in self.color_items:
                pfad = self.mb.path_to_extension
                url = uno.systemPathToFileUrl(os.path.join(pfad,'img','punkt_%s.png' %url))
            elif url == 'leer':
                url = ''
            
            oButton.setPropertyValues(
                                      ('BackgroundColor','Name','Border','ScaleImage','ImageURL'),
                                      (self.hintergrund,btn_name,0,False,url)
                                      )

            form.insertByName(btn_name,oButton)
            
            shape = self.calc.createInstance("com.sun.star.drawing.ControlShape")
                        
            p = shape.Position
            p.X = pos[0] 
            p.Y = pos[1]
            shape.setPosition(p)
            
            s = shape.Size
            s.Height = 400
            s.Width = 400
            shape.setSize(s)
            
            shape.Control = oButton
            
            return shape
        except:
            log(inspect.stack,tb())
            
    
    def erzeuge_steuerung(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Button_Click_Listener(self.mb,self)
            self.listener = listener
            
            
            self.sheet_controller.setFormDesignMode(True)
            draw_page = self.sheet.DrawPage
            self.draw_page = draw_page
            
            shape1 = self.erzeuge_button(LANG.MENU,'Menu',0,listener)
            shape2 = self.erzeuge_button(LANG.INFO,'Info',1,listener)
            shape3 = self.erzeuge_button(LANG.UEBERNEHMEN,'Uebernehmen',2,listener)
            draw_page.add(shape1)
            draw_page.add(shape2)
            draw_page.add(shape3)
            
            dp = draw_page.getByIndex(0)
            
            self.sheet_controller.setFormDesignMode(False)
 
        except:
            log(inspect.stack,tb())
            
    
    def oeffne_calc(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
            
            prop2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop2.Name = 'AsTemplate'
            prop2.Value = True
                    
            URL="private:factory/scalc"
                    
            self.calc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,())
            
            frames = self.mb.desktop.Frames
            for i in range(frames.Count):                
                if frames.getByIndex(i).Controller == self.calc.CurrentController:
                    self.calc_frame = frames.getByIndex(i)
            
            self.sheet_controller = self.calc_frame.Controller
            self.sheet = self.sheet_controller.ActiveSheet
            
            self.calc_frame.setPropertyValue('Title','Organon Organizer')
            
#             self.mb.calc = {'calc_frame':self.calc_frame,
#                             'calc':self.calc,
#                             'sheet_controller':self.sheet_controller,
#                             'sheet':self.sheet}
            

            RESOURCE_URL = "private:resource/dockingwindow/9809"
            self.calc_frame.LayoutManager.hideElement(RESOURCE_URL)
#             pd()
#             listener = Close_Listener(self.mb,self)
#             self.calc_frame.addCloseListener(listener)
        except:
            log(inspect.stack,tb())
            
        
    def erzeuge_ansicht(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            x0,y0,x1,y1 = self.pos
            
            sett = self.mb.settings_orga['organon_farben']
            
            lmgr = self.calc_frame.LayoutManager
            
            for el in lmgr.Elements:
                lmgr.hideElement(el.ResourceURL)
                        
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Sidebar'
            prop.Value = False
            
            dispatcher = self.mb.createUnoService("com.sun.star.frame.DispatchHelper") 
            dispatcher.executeDispatch(self.calc_frame, ".uno:Sidebar", "", 0, (prop,))
            
            prop.Name = 'ViewRowColumnHeaders'
            prop.Value = False
            dispatcher.executeDispatch(self.calc_frame, ".uno:ViewRowColumnHeaders", "", 0, (prop,))
                      
            self.calc_frame.LayoutManager.AutomaticToolbars = False
            
            # Hintergrund faerben
            rangex = self.sheet.getCellRangeByPosition( 0, 0, x0 - 1, self.sheet.Rows.Count-1 )
            rangex.CellBackColor = self.hintergrund
            rangex = self.sheet.getCellRangeByPosition( 0, 0, self.sheet.Columns.Count-1, y0 - 1 )
            rangex.CellBackColor = self.hintergrund
            
            rangex = self.sheet.getCellRangeByPosition( x1, 0, self.sheet.Columns.Count-1, self.sheet.Rows.Count-1 )
            rangex.CellBackColor = self.hintergrund
            rangex = self.sheet.getCellRangeByPosition( 0, y1, self.sheet.Columns.Count-1, self.sheet.Rows.Count-1 )
            rangex.CellBackColor = self.hintergrund
            
            # Tabellenhintergruende
            rangex = self.sheet.getCellRangeByPosition( 0, y1, self.sheet.Columns.Count-1, self.sheet.Rows.Count-1 )
            rangex.CellBackColor = self.hintergrund
            
            # WRAP TEXT
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'WrapText'
            prop.Value = True            
            
            dispatcher.executeDispatch(self.calc_frame, ".uno:WrapText", "", 0, (prop,))

        except:
            log(inspect.stack,tb())
            
        
        
    def setze_eintraege(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            # WRAP TEXT
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'WrapText'
            prop.Value = True            
            
            dispatcher = self.mb.createUnoService("com.sun.star.frame.DispatchHelper") 
            dispatcher.executeDispatch(self.calc_frame, ".uno:WrapText", "", 0, (prop,))
            
            x0,y0,x1,y1 = self.pos
            y1 = y0
            
            xEnde = x0 + len(self.spalten2)

            def set_farben(c,b,hg):
                c.TopBorder = b
                c.BottomBorder = b
                c.LeftBorder = b
                c.RightBorder = b
                c.CellBackColor = hg
                                
            
            dict_sb = self.dict_sb_content
            tags_ord = dict_sb['ordinal']

            if T.AB == 'Projekt':
                Eintraege = self.mb.class_Projekt.lese_xml_datei()
            else:
                Eintraege = self.mb.class_Tabs.lade_tab_Eintraege(T.AB)
                        
            
            self.Eintraege = Eintraege
            self.modify_listener = Modify_Listener(self.mb,self,self.Eintraege)
            self.modify_listener.pos = self.pos
            
            # Border    
            cell = self.sheet.getCellByPosition(0,0)
            border = cell.TopBorder
            border.Color = self.hintergrund
            border.OuterLineWidth = 100
            
            # KOPFZEILE
            for s in range(len(self.spalten2)):
                cell = self.sheet.getCellByPosition(x0 + s ,y0)
                text = self.spalten2[s]
                cell.String = text
                
                cell.CharWeight = 150
                cell.CharColor = self.menu_schrift
                
            
            rangex = self.sheet.getCellRangeByPosition(x0,y0,xEnde,y0)
            set_farben(rangex,border,self.menu_hintergrund)
            
            
            # Zellenabstaende innen            
            abstand = 200
            rangex.ParaLeftMargin = abstand
            rangex.ParaRightMargin = abstand
            rangex.ParaBottomMargin = abstand
            rangex.ParaTopMargin = abstand
            
            # PROTECTION
            prot = cell.CellProtection
            prot.IsLocked = False
                
            from com.sun.star.table.CellVertJustify import CENTER
            
            rangex = self.sheet.getCellRangeByPosition(x0, y0+1, xEnde, y0+len(Eintraege))
            rangex.CharColor = self.schrift_datei

            # EINTRAEGE
            for e in Eintraege:
                pass
                if e[0] == self.mb.props['Projekt'].Papierkorb:
                    y1 += 1
                    break

                cell = self.sheet.getCellByPosition(x0 ,y1 + 1)
                cell.String = e[2]
                
                cell.Text.VertJustify = CENTER
                
                if e[4] != 'pg':
                    cell.CharColor = self.schrift_ordner
                
                
                for x_sp in range(len(self.spalten)):
 
                    cell = self.sheet.getCellByPosition(x0 + x_sp + 1 ,y1 + 1)
 
                    text = tags_ord[e[0]][self.spalten[x_sp]]
                    
                    if isinstance(text, list):
                        if text != []:
                            text = ',\n'.join(text)
                            cell.String = text
                    elif text != '':
                        cell.String = text
                        
                y1 += 1

            rangex = self.sheet.getCellRangeByPosition(x0, y0+1, x0+len(self.spalten2), y0+len(Eintraege))
            self.rangex = rangex
            
            border.OuterLineWidth = 10
            set_farben(rangex,border,self.hf_hintergrund)
            
            rangex.setPropertyValue('CellProtection',prot)
            rangex.addModifyListener(self.modify_listener)
            
            self.data_array = [list(zeile) for zeile in rangex.DataArray]
            abc = [list(z.replace('\n','') if zeile.index(z) not in (1,2) else z for z in zeile ) 
                   for zeile in rangex.DataArray]
            
            self.data_array_orig = copy.deepcopy(rangex.DataArray)
            
            # Zellenabstaende innen            
            abstand = 200
            rangex.ParaLeftMargin = abstand
            rangex.ParaRightMargin = abstand
            rangex.ParaBottomMargin = abstand
            rangex.ParaTopMargin = abstand
            

            self.sheet_controller.select(rangex)
            dispatcher.executeDispatch(self.calc_frame, ".uno:WrapText", "", 0, (prop,))
            
            # Die Breiten sollten in eine proj prop ausgelagert werden,
            # um sie vom Nutzer bei Bedarf anpassen lassen zu können
            breiten = [.5,.3,.3,
                       1.5,
                       2.5,
                       2.5,
                       1,1,1,1,1,1,1]
    
            for b in range(len(breiten)):
                spalte = self.sheet.Columns.getByIndex(b)
                spalte.Width = breiten[b] * 2540
                
            self.pos[3] = y1
        
        except:
            log(inspect.stack,tb())
        
        
    def setze_icons(self): 
        if self.mb.debug: log(inspect.stack)
        
        x0,y0,x1,y1 = self.pos 
        y1 = y0
        
        self.icons = {}
        self.modify_listener.icons = self.icons
                
        def berechne_pos(c):
            pos = c.Position
            h = c.Size.Height
            hh = (h - 500)/2

            y = pos.Y + hh
            return pos.X,y
        
        try: 
            draw_page = self.sheet.DrawPage
            
            for e in self.Eintraege:
                
                if e[0] == self.mb.props['Projekt'].Papierkorb:
                    y1 += 1
                    break
                
                cell = self.sheet.getCellByPosition(x0 - 1 ,y1 + 1)
                
                name = 'IMG_{}'.format(e[0])
                
                shape1 = self.erzeuge_img_button(berechne_pos(cell),name,e[7],self.listener)
                draw_page.add(shape1)
                self.icons.update({name:shape1})
                
                cell = self.sheet.getCellByPosition(x0 - 2 ,y1 + 1)

                name2 = 'IMGU_{}'.format(e[0])
                
                shape2 = self.erzeuge_img_button(berechne_pos(cell),name2,e[8],self.listener)
                draw_page.add(shape2)
                self.icons.update({name2:shape2})
                
                y1 += 1

        except:
            log(inspect.stack,tb())
            

 

from com.sun.star.beans import XPropertiesChangeListener      
class Button_Click_Listener(unohelper.Base, XPropertiesChangeListener):

    def __init__(self, mb,Org):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.pressed = False
        self.Org = Org
        
    def propertiesChange(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
#             RESOURCE_URL = "private:resource/dockingwindow/9809"
#             if self.Org.calc_frame.LayoutManager.isElementVisible(RESOURCE_URL):
#                 self.Org.calc_frame.LayoutManager.hideElement(RESOURCE_URL)
#             else:
#                 self.Org.calc_frame.LayoutManager.showElement(RESOURCE_URL)
#             return
            btn = ev[0].Source
            
            # MENULEISTE EIN/AUSBLENDEN
            if btn.Label == LANG.MENU:
                lmgr = self.Org.calc_frame.LayoutManager
                
                ResourceURL ='private:resource/menubar/menubar'
                
                if lmgr.isElementVisible(ResourceURL):
                    lmgr.hideElement(ResourceURL)
                else:
                    lmgr.showElement(ResourceURL)
                    
            # INFO ANZEIGEN
            elif btn.Label == LANG.INFO:
                self.mb.nachricht(LANG.ORGANIZER_INFO.format(
                         LANG.UEBERNEHMEN,
                         LANG.MENU,
                         LANG.SYNOPSIS,
                         LANG.NOTIZEN,
                         LANG.CHARAKTERE,
                         LANG.ORTE),
                         "infobox")
            # TAGS UEBERNEHMEN
            else:
                self.tags_uebernehmen()
                
        except:
            log(inspect.stack,tb())
            
    
    
    def tags_uebernehmen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sichtbare = self.mb.dict_sb_content['sichtbare']
            self.mb.dict_sb_content = self.Org.dict_sb_content
            
            for tag in sichtbare:
                self.mb.class_Sidebar.erzeuge_sb_layout(tag,'focus_lost')
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','focus_lost')
            
            self.mb.nachricht(LANG.UEBERNOMMEN,'infobox')            
        except:
            log(inspect.stack,tb())
            
            

    def aendere_dateinamen(self):
        # fehlt
        pass
    
    def disposing(self,ev):
        return False
        
        
from com.sun.star.util import XModifyListener      
class Modify_Listener(unohelper.Base, XModifyListener):

    def __init__(self, mb,Org,eintraege):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.Org = Org
        self.eintraege = eintraege
        self.pos = None
        self.icons = None
        
        self.new_data_array = []
        
         
    def modified(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            aenderungen = self.finde_veraenderung()
            self.pruefe_veraenderung(aenderungen)
            self.Org.data_array = self.new_data_array
            self.update_icon_pos()
        except:
            log(inspect.stack,tb())

    
    def finde_veraenderung(self):
        if self.mb.debug: log(inspect.stack)
        
        x0,y0,x1,y1 = self.pos
        rangex = self.Org.sheet.getCellRangeByPosition(x0, 
                                                       y0+1, 
                                                       x0+len(self.Org.spalten2), 
                                                       y0+len(self.Org.Eintraege))
        
        self.new_data_array = [list(z) for z in rangex.DataArray]
        
        aenderungen = []
        
        for d in range(len(self.new_data_array)):
            if self.new_data_array[d] != self.Org.data_array[d]:
                
                for j in range(len(self.new_data_array[d])):
                    if self.new_data_array[d][j] != self.Org.data_array[d][j]:
                        aenderungen.append([ j + x0, d + y0 + 1, self.new_data_array[d][j] ])
                        
        return aenderungen
    
    
    def zelle_nach_liste(self,inhalt):
        if self.mb.debug: log(inspect.stack)
        inhalt = inhalt.replace('\n','').split(',')
        return [i.strip() for i in inhalt if i.strip() != '']
    
    
    def liste_nach_zelle(self,inhalt):
        if self.mb.debug: log(inspect.stack)
        return ',\n'.join(inhalt)
    
    
    def setze_zelle(self,cell,inhalt):    
        if self.mb.debug: log(inspect.stack)
        
        # Listener entfernen und wieder adden, damit er bei Aenderung
        # der Zelle nicht angesprochen wird.
        self.Org.rangex.removeModifyListener(self.Org.modify_listener)
        cell.String = inhalt
        self.Org.rangex.addModifyListener(self.Org.modify_listener)
        
        
    def wieder_loeschen(self,col,row,kateg):
        if self.mb.debug: log(inspect.stack)
                
        cell = self.Org.sheet.getCellByPosition(col,row)
        c,r = col-self.pos[0],row-self.pos[1]-1
        alter_eintrag = self.Org.data_array[r][c]

        self.setze_zelle(cell, alter_eintrag)
        self.new_data_array[r][c] = alter_eintrag
        
            
    def get_infos(self,col,row):
        if self.mb.debug: log(inspect.stack)
        
        dict_sb = self.Org.dict_sb_content
        c,r = col-self.pos[0],row-self.pos[1]-1
        ordinal = self.eintraege[r][0]
        return dict_sb,c,r,ordinal
    
    
    def setze_synopsis_und_notizen(self,col,row,kateg,inhalt):
        if self.mb.debug: log(inspect.stack)
        
        dict_sb,c,r,ordinal = self.get_infos(col,row)
        dict_sb['ordinal'][ordinal][kateg] = inhalt
        self.new_data_array[r][c] = inhalt
                
    
    def tag_auf_vorkommen_testen(self,tag,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        # pruefen, ob tag noch in anderen dateien benutzt wird
        dict_sb = self.Org.dict_sb_content
        
        for ordi in dict_sb['ordinal']:
            if ordi != ordinal:
                if tag in dict_sb['ordinal'][ordi]['Tags_general']:
                    dic = dict_sb['ordinal'][ordi]
                    return True
        return False
    
    
    def loesche_tags_aus_dict(self,tags,ordinal,kateg):
        if self.mb.debug: log(inspect.stack)
        
        def loesche(fkt,value):
            try:
                fkt(value)
            except:
                log(inspect.stack,tb())
        
        dict_sb = self.Org.dict_sb_content
        
        for tag in tags:
            # aus den eigenen tags loeschen
            loesche(dict_sb['ordinal'][ordinal]['Tags_general'].remove,tag)
            loesche(dict_sb['ordinal'][ordinal][kateg].remove,tag)
            # aus den allg. Tags loeschen
            if not self.tag_auf_vorkommen_testen(tag,ordinal):
                loesche(dict_sb['tags']['Tags_general'].remove,tag)
                loesche(dict_sb['tags'][kateg].remove,tag)     
     
     
    def setze_leeren_inhalt(self,col,row,kateg):
        if self.mb.debug: log(inspect.stack)
        
        try:
            dict_sb,c,r,ordinal = self.get_infos(col,row)
            alte_eintraege = self.zelle_nach_liste(self.Org.data_array[r][c])
            
            if kateg in ('Tags_characters','Tags_locations','Tags_objects',
                         'Tags_user1','Tags_user2','Tags_user3'):
                self.loesche_tags_aus_dict(alte_eintraege,ordinal,kateg)
            self.new_data_array[r][c] = ''
        except:
            log(inspect.stack,tb())
    
    
    def tag_hinzufuegen(self,tag,ordinal,kateg):
        if self.mb.debug: log(inspect.stack)
        
        dict_sb = self.Org.dict_sb_content
        
        if tag not in dict_sb['ordinal'][ordinal]['Tags_general']:
            dict_sb['ordinal'][ordinal]['Tags_general'].append(tag)
        if tag not in dict_sb['ordinal'][ordinal][kateg]:
            dict_sb['ordinal'][ordinal][kateg].append(tag)
        if tag not in dict_sb['tags'][kateg]:
            dict_sb['tags'][kateg].append(tag)
        if tag not in dict_sb['tags']['Tags_general']:
            dict_sb['tags']['Tags_general'].append(tag)

 
    def setze_aenderung(self,col,row,kateg,inhalt,geloeschte,hinzugefuegte):
        if self.mb.debug: log(inspect.stack)
        
        try:
            dict_sb,c,r,ordinal = self.get_infos(col,row)
            
            self.loesche_tags_aus_dict(geloeschte,ordinal,kateg)
            
            for tag in hinzugefuegte:
                self.tag_hinzufuegen(tag,ordinal,kateg)
        
            cell = self.Org.sheet.getCellByPosition(col,row)
            self.setze_zelle(cell, self.liste_nach_zelle(inhalt))
        
            self.new_data_array[r][c] = self.liste_nach_zelle(inhalt)
        except:
            log(inspect.stack,tb())        
        
        
    def pruefe_veraenderung(self,aenderungen):
        if self.mb.debug: log(inspect.stack)
        
        try:

            for a in aenderungen:
                
                col,row,inhalt = a
                dict_sb,c,r,ordinal = self.get_infos(col,row)
                
                if col - self.pos[0] == 0:
                    kateg = 'datei'
                else:
                    kateg = self.Org.spalten[col - self.pos[0] - 1]

                # PROJEKTNAME
                if (col,row) == (self.Org.pos[0],self.Org.pos[1] + 1):
                    self.wieder_loeschen(col,row,kateg,new_data_array)
                    self.mb.nachricht(LANG.PRJ_NAME_KEINE_AENDERUNG,"infobox")
                    continue
                # DATEINAME
                if kateg == 'datei':
                    r = row-self.pos[1]-1
                    self.eintraege[r][2] = inhalt
                    continue
                # LEERER INHALT
                if inhalt == '':
                    self.setze_leeren_inhalt(col,row,kateg)
                    continue
                # SYNOPSIS, NOTIZEN            
                if kateg in ('Synopsis','Notes'):
                    self.setze_synopsis_und_notizen(col,row,kateg,inhalt)
                    continue
                
                
                # PERSONEN, ORTE, OBJEKTE, ect.
                # Art der Aenderung bestimmen
                inhalt = self.zelle_nach_liste(inhalt)
                
                alte_eintraege = self.zelle_nach_liste(self.Org.data_array[r][c])
                
                set_inhalt = set(inhalt)
                set_alt = set(alte_eintraege)
                set_kategorie = set(dict_sb['tags'][kateg])
                set_general = set(dict_sb['tags']['Tags_general'])
                
                tags_gen_ohne_kateg = set_general.difference(set_kategorie)
                bereits_in_anderen_tags_vorhanden = len(set_inhalt.intersection(tags_gen_ohne_kateg)) > 0
                
                keine_aenderung = set_inhalt == set_alt

                geloeschte = set_alt.difference(set_inhalt)
                hinzugefuegte = set_inhalt.difference(set_alt)


                if keine_aenderung:
                    self.wieder_loeschen(col,row,kateg)
                
                elif bereits_in_anderen_tags_vorhanden:
                    self.mb.nachricht(LANG.EIN_TAG_BEREITS_VORHANDEN,"infobox")
                    self.wieder_loeschen(col,row,kateg)                   
                else:
                    self.setze_aenderung(col,row,kateg,inhalt,geloeschte,hinzugefuegte)
                    
        except:
            log(inspect.stack,tb())
    
        
    def update_icon_pos(self):
        if self.mb.debug: log(inspect.stack)
        
        x0,y0,x1,y1 = self.pos 
        y1 = y0
        
        def berechne_pos(c):
            pos = c.Position
            h = c.Size.Height
            hh = (h - 500)/2

            y = pos.Y + hh
            pos.Y = y
            return pos
        
        try:
            for e in self.eintraege:
                
                if e[0] == self.mb.props['Projekt'].Papierkorb:
                    y1 += 1
                    break
                
                cell = self.Org.sheet.getCellByPosition(x0 - 1 ,y1 + 1)
                icon = self.icons['IMG_{}'.format(e[0])]

                icon.setPosition(berechne_pos(cell))
                
                cell = self.Org.sheet.getCellByPosition(x0 - 2 ,y1 + 1)
                icon = self.icons['IMGU_{}'.format(e[0])]

                icon.setPosition(berechne_pos(cell))
                
                y1 += 1
        except:
            print(tb())
    
    def disposing(self,ev):
        return False
    
    
from com.sun.star.awt import XEnhancedMouseClickHandler       
class Enhanced_MC_Handler(unohelper.Base, XEnhancedMouseClickHandler):

    def __init__(self, mb, Org):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.Org = Org
        
    def disposing(self, ev):
        return False

    def mousePressed(self, ev):
        # nur loggen, wenn tatsaechlich verwendet
        
        if hasattr(ev.Target, 'Control'):
            if self.mb.debug: log(inspect.stack)
            
            print(ev.Target.Control.Name)
        else:
            x0,y0,x1,y1 = self.Org.pos
            
            range_col = list(range(x0-2,x0))
            range_row = list(range(y0 + 1 ,y0 + len(self.Org.Eintraege) - 1))

            addr = ev.Target.CellAddress
            col,row = addr.Column, addr.Row

            if col in range_col and row in range_row:
                if self.mb.debug: log(inspect.stack)
                
                reihe = row - y0 -1
                ord_source = self.Org.Eintraege[reihe][0]
                window_parent = self.Org.calc_frame.getComponentWindow()
                X,Y = ev.X,ev.Y
                
                # FARB ICONS
                if col == y0 - 1:
                    self.mb.class_Funktionen.erzeuge_Tag1_Container(ord_source,X,Y,window_parent=window_parent)
                else:
                    self.mb.class_Funktionen.erzeuge_Tag2_Container(ord_source,X,Y,window_parent=window_parent)

        return True
   
    def mouseReleased(self, ev):
        return True
    

from com.sun.star.util import XCloseListener       
class Close_Listener(unohelper.Base, XCloseListener):

    def __init__(self, mb, Org):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.Org = Org
        
    def queryClosing(self, ev):
        pd()
        RESOURCE_URL = "private:resource/dockingwindow/9809"
        self.Org.calc_frame.LayoutManager.showElement(RESOURCE_URL)
        self.mb.desktop.ActiveFrame.LayoutManager.showElement(RESOURCE_URL)
        #pd()
        return True  
    
    def notifyClosing(self, ev):
        #pd()
        RESOURCE_URL = "private:resource/dockingwindow/9809"
        self.Org.calc_frame.LayoutManager.showElement(RESOURCE_URL)
        self.mb.desktop.ActiveFrame.LayoutManager.showElement(RESOURCE_URL)
        #pd()
        return True          
            
#                 
# from com.sun.star.awt import XMouseClickHandler       
# class Handler(unohelper.Base, XMouseClickHandler, object):
#     """ Handles mouse click on the document. """
#     def __init__(self, ctx, doc):
#         self.ctx = ctx
#         self.doc = doc
#         self._register()
#         
#     def _register(self):
#         self.doc.getCurrentController().addMouseClickHandler(self)
#         self.ctx = uno.getComponentContext() 
#       
#     def unregister(self):
#         """ Remove myself from broadcaster. """
#         self.doc.getCurrentController().removeMouseClickHandler(self)
# 
#     def disposing(self, ev):
#         return False
#         global handler
#         handler = None
# 
#     def mousePressed(self, ev):
#         return False
#    
#     def mouseReleased(self, ev):
#         
#         try:
#             selected = self.doc.getCurrentSelection()
#         except:
#             print(tb())
#             
#         
#         return False
#         if ev.Buttons == MB_LEFT:
#             
#             addr = selected.getRangeAddress()
#  
#         return False        
#         
#         
#   
# import timeit
# class Timer():
#     
#     def __init__(self):
#         self.start_time = timeit.default_timer()     
#     
#     def t(self,extras = ''):
#         
#         print('{0}  {1}'.format("{:10.2f}".format(timeit.default_timer()-self.start_time),extras))
#         
#         
        
        
        