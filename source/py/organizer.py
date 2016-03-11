# -*- coding: utf-8 -*-

import unohelper

class Organizer():
    '''
    Es fehlt:
    - keine zwei Instanzen des Organizers oeffnen
    - beim Schließen von Calc evt. User Icons loeschen
    ( modul funktionen: Tag2_Item_Listener.galerie_icon_im_prj_ordner_evt_loeschen() )
    - Einfuegen von Dateien/Ordnern und
    - Verschieben (dann muesste allerdings die Baumstruktur abgebildet werden. Evt. zu kompliziert)
    - Loeschen (in den Papierkorb verschieben)
    '''
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.tags_org = None
        self.calc = None
        self.sheet = None
        self.sheet_controller = None
        self.calc_frame = None
        
        self.btn = []
        self.data_array = []
        self.first_time_info = True
        self.rangex = None
        self.bilder_reihe = {}
        
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
                    Popup(self.mb, 'info').text = LANG.OO_ORGANIZER_INFO
                    
            if len(self.mb.props['ORGANON'].dict_zeilen_posY) == 0:
                # Kein Projekt geladen
                return
            
            
            self.initialisiere()
            self.erzeuge_listener_und_event()
            
            self.oeffne_calc()
            self.get_farben()
            self.get_eintraege()
            self.erzeuge_ansicht()
            
            rangex = self.sheet.getCellRangeByPosition(0,0,0,0)
            self.sheet_controller.select(rangex)
            
            self.erzeuge_steuerung()
            self.setze_eintraege()
            self.setze_icons()   
            # in den hintergrund         
            self.schuetze_document()
            
            if self.mb.programm == 'LibreOffice':
                for br in self.bilder_reihe:
                    self.update_bilder_pos(br)
            
            handler = Enhanced_MC_Handler(self.mb,self)
            self.sheet_controller.addEnhancedMouseClickHandler(handler)
            
            self.calc.CurrentController.freezeAtPosition(0,4)
            
            self.calc.setModified(False)
     
        except:
            log(inspect.stack,tb())


    def initialisiere(self):
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        self.bilder_reihe = {}
        self.tags_org = copy.deepcopy(self.mb.tags)
        
        self.spalten = [a for a in tags['abfolge'] if a in tags['sichtbare']]
        self.spalten2 = [tags['nr_name'][a][0] for a in tags['abfolge'] if a in tags['sichtbare']]
        self.spalten2.insert(0, LANG.DATEI)
        
        self.bild_panels = [t for t in self.spalten if self.tags_org['nr_name'][t][1] == 'img']
        self.txt_panels = [t for t in self.spalten if self.tags_org['nr_name'][t][1] == 'txt']
        self.date_panels = [t for t in self.spalten if self.tags_org['nr_name'][t][1] == 'date']
        self.time_panels = [t for t in self.spalten if self.tags_org['nr_name'][t][1] == 'time']
        # self.pos bestimmt den Ort der Tabelle
        # Wenn eine Spalte/Zeile mehr benoetigt werden sollte,
        # koennen die ersten beiden Werte geaendert werden
        self.pos = [ 3, 3, 3 + len(self.spalten2), 3 ]
        
        
    def erzeuge_listener_und_event(self):
        if self.mb.debug: log(inspect.stack)
        
        self.listener_btn_click = Button_Click_Listener(self.mb,self)
        self.listener_icons = Icons_Listener(self.mb,self)
        
        oEvent = uno.createUnoStruct('com.sun.star.script.ScriptEventDescriptor')
        oEvent.ListenerType = "com.sun.star.awt.XMouseListener"
        oEvent.EventMethod = "mousePressed"
        oEvent.ScriptType = "Script"
        # Script Code existiert nicht. Das Event wird nur benutzt, um 
        # firing des Listeners auszuloesen
        oEvent.ScriptCode = "vnd.sun.star.script:NotExistingMacro?language=Basic&location=document" 
        
        self.event = oEvent
        
    def schuetze_document(self):
        if self.mb.debug: log(inspect.stack)
        
        # Schuetzen, Aufheben des Schutzes, Schuetzen wird ausgefuehrt,
        # damit die Icons auf die richtige Hoehe gesetzt werden.
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
            
        self.hintergrund = sett['office']['dok_hintergrund']
        
    
    def erzeuge_button(self,lbl,btn_name,pos):
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
            #oButton.Toggle = 1
            
            oButton.addPropertiesChangeListener(('State',),self.listener_btn_click)
            
            form.registerScriptEvent( 0, self.event )
            form.addScriptListener(self.listener_icons) 
            
            return shape
        except:
            log(inspect.stack,tb())
            

    def bild_einfuegen(self,col,row,panel_nr,url=None):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            c,r = col-self.pos[0], row-self.pos[1]-1
            ordinal = self.Eintraege[r][0]
            
            if not url:
                url,ok = self.mb.class_Funktionen.filepicker2(url_to_sys=False)
                if not ok:
                    return True
            
            cell = self.sheet.getCellByPosition(col,row)
            
            self.setze_zelle_und_tag(c, r, cell, url, ordinal, panel_nr)
                            
            zell_breite = cell.Size.Width
            
            name = 'panel{0}_{1}'.format(panel_nr,ordinal)
            
            btn,hoehe = self.erzeuge_bild_button(name,url,zell_breite)

            shape = self.erzeuge_bild_shape(name, btn, cell,hoehe)
            self.sheet.DrawPage.add(shape)
            
            if row not in self.bilder_reihe:
                self.bilder_reihe[row] = {}
            if panel_nr not in self.bilder_reihe[row]:
                self.bilder_reihe[row][panel_nr] = {}
            
            self.bilder_reihe[row][panel_nr]['shape'] = shape
            self.bilder_reihe[row][panel_nr]['hoehe'] = hoehe
            self.bilder_reihe[row][panel_nr]['values'] = [ ordinal,url,panel_nr,cell ]  
            
            self.update_bilder_pos(row) 
            self.update_icon_pos(col, row)
            self.tags_org['ordinale'][ordinal][panel_nr] = url
        except:
            log(inspect.stack,tb())            
            
        
    def erzeuge_bild_button(self,name,url,zell_breite):    
        if self.mb.debug: log(inspect.stack)
                
        btn = self.mb.createUnoService("com.sun.star.form.component.ImageButton")
        btn.setPropertyValues(('Name','Border','ScaleImage','ImageURL'),
                              (name,0,True,url))
        graphic = btn.Graphic
        
        # BERECHNE HOEHE
        h = graphic.Size.Height
        b = graphic.Size.Width
        quotient = float(b)/float(h)
        hoehe = int(zell_breite / quotient)
        
        return btn,hoehe
    
    
    def erzeuge_bilder_reihe_btns(self,row):
        if self.mb.debug: log(inspect.stack)
        
        btns = {}
        h_max = 0
        
        for pannel_nr in self.bilder_reihe[row]:
            ordinal,bild_name,panel_nr,cell = self.bilder_reihe[row][pannel_nr]['values']
            bildpfad = uno.systemPathToFileUrl(os.path.join(self.mb.pfade['images'], bild_name))
            name = 'panel{0}_{1}'.format(panel_nr,ordinal) 
            btn,hoehe = self.erzeuge_bild_button(name,bildpfad,cell.Size.Width)
            h_max = hoehe if hoehe > h_max else h_max
            
            btns[panel_nr] = [btn,hoehe,name]
        
        return btns,h_max
        
    
    def erzeuge_bild_shape(self,name,btn,cell,hoehe):    
        if self.mb.debug: log(inspect.stack)
        
        forms = self.sheet.DrawPage.Forms
        
        form = self.calc.createInstance("com.sun.star.form.component.Form")
        form.setName(name)
        form.insertByName('btn_' + name, btn)
        
        shape = self.calc.createInstance("com.sun.star.drawing.ControlShape")
            
        # Das Bild wird 10% kleiner als die Panelbreite dargestellt
        # Abstand auf jeder Seite: 5%
        
        self.setze_shape_pos(shape, cell, hoehe)
        
        shape.Control = btn
        shape.Name = name
        
        form.registerScriptEvent( 0, self.event )
        form.addScriptListener(self.listener_icons) 
        
        forms.insertByIndex (0,form)
        
        return shape
    
    
    def setze_shape_pos(self,shape,cell,hoehe):
        if self.mb.debug: log(inspect.stack)
        
        breite = cell.Size.Width
                        
        p = shape.Position
        p.X = cell.Position.X + breite/20
        p.Y = cell.Position.Y + (cell.Size.Height - hoehe/10*9) /2
        shape.setPosition(p)
        
        s = shape.Size
        s.Height = hoehe/10*9
        s.Width = breite/10*9
        shape.setSize(s)
        
        
    def setze_icon_pos(self,shape,cell):
        if self.mb.debug: log(inspect.stack)
        
        breite = cell.Size.Width
                        
        p = shape.Position
        p.X = cell.Position.X + breite/20
        p.Y = cell.Position.Y + (cell.Size.Height - hoehe/10*9) /2
        shape.setPosition(p)
        
        s = shape.Size
        s.Height = hoehe/10*9
        s.Width = breite/10*9
        shape.setSize(s)
        
        
    def setze_zell_hoehe(self,cell,hoehe):  
        if self.mb.debug: log(inspect.stack)
        
        # je nach Hoehe der Eintraege Zellhoehe anpassen         
        if cell.Size.Height < hoehe:
            cell.Rows.setPropertyValue('Height',hoehe) 
        else:
            cell.Rows.setPropertyValue('OptimalHeight',True)
            # Bei einem Update des Bildes muss erneut geprueft werden
            if cell.Size.Height < hoehe:
                cell.Rows.setPropertyValue('Height',hoehe)   

 
    def update_bilder_pos(self,row):
        if self.mb.debug: log(inspect.stack)
        
        def setze_bild(r):
            shapes = [v for a,v in self.bilder_reihe[r].items()]
            
            cell = self.sheet.getCellByPosition(0,r)
                
            cell.Rows.setPropertyValue('OptimalHeight',True)
            
            if shapes != []:
                h_max = sorted([v['hoehe'] for a,v in self.bilder_reihe[r].items()])[-1]
                self.setze_zell_hoehe(cell,h_max)
                
                for v in shapes:
                    shape = v['shape']
                    hoehe = v['hoehe']
                    cell = v['values'][3]
                    self.setze_shape_pos(shape,cell,hoehe)
        
        
        if self.mb.programm == 'LibreOffice':
            for row in self.bilder_reihe:
                setze_bild(row)
        else:
            if row in self.bilder_reihe:
                setze_bild(row)
            
    def update_icon_pos_OO(self,col,row):
        if self.mb.debug: log(inspect.stack)
        
        def berechne_pos(c):
            pos = c.Position
            h = c.Size.Height
            h2 = 500#i.Size.Height
            hh = int((float(h) - h2)/2)
            
            y = pos.Y + hh
            pos.Y = y
            return pos

        def do(name,cell):
            
            icon = self.icons[name]
            pos = berechne_pos(cell)
            
            size = icon.Size
            size.Height = 500
            icon.setSize(size)
            
            icon.setPosition(berechne_pos(cell))

        
        try:
            visible_tags_tv = self.mb.settings_proj['tag1'],self.mb.settings_proj['tag2']
            ordinal = self.Eintraege[row-self.pos[1]-1][0]

            if visible_tags_tv[0]:
                name = 'IMG_{}'.format(ordinal)
                cell = self.sheet.getCellByPosition(self.pos[0]-1,row)
                do(name,cell)
                
            if visible_tags_tv[1]:
                name = 'IMGU_{}'.format(ordinal)
                cell = self.sheet.getCellByPosition(self.pos[0]-2,row)
                do(name,cell)
            
        except:
            log(inspect.stack,tb())    

    def update_icon_pos(self,col,row):
        if self.mb.debug: log(inspect.stack)
        
        visible_tags_tv = self.mb.settings_proj['tag1'],self.mb.settings_proj['tag2']
        OO = self.mb.programm == 'OpenOffice'
        
        
        def berechne_pos(c):
            pos = c.Position
            h = c.Size.Height
            h2 = 500#i.Size.Height
            hh = int((float(h) - h2)/2)
            
            y = pos.Y + hh
            pos.Y = y
            return pos

        def do(name,cell):
            icon = self.icons[name]
            pos = berechne_pos(cell)
            icon.setPosition(berechne_pos(cell))
            
            if OO:
                size = icon.Size
                size.Height = 500
                icon.setSize(size)

        
        def aendere_pos(ordi,r):
            if visible_tags_tv[0]:
                name = 'IMG_{}'.format(ordi)
                cell = self.sheet.getCellByPosition(self.pos[0]-1,r)
                do(name,cell)
                
            if visible_tags_tv[1]:
                name = 'IMGU_{}'.format(ordi)
                cell = self.sheet.getCellByPosition(self.pos[0]-2,r)
                do(name,cell)


        try:
            ordinal = self.Eintraege[row-self.pos[1]-1][0]
            
            if OO:
                aendere_pos(ordinal,row)
            else:
                for i,e in enumerate(self.Eintraege):
                    ordinal = e[0]
                    aendere_pos(ordinal,self.pos[1] + i + 1)
        except:
            log(inspect.stack,tb())
            
                    
    def erzeuge_bilder_reihen(self):
        if self.mb.debug: log(inspect.stack)
        
        draw_page = self.sheet.DrawPage

        for row in self.bilder_reihe:
            btns,h_max = self.erzeuge_bilder_reihe_btns(row)
            
            erstes_panel = list(self.bilder_reihe[row].keys())[0]
            
            cell = self.bilder_reihe[row][erstes_panel]['values'][3]
            self.setze_zell_hoehe(cell, h_max)
                            
            for panel_nr,v in btns.items():
                
                cell = self.bilder_reihe[row][panel_nr]['values'][3]
                btn,hoehe,name = btns[panel_nr]
                shape = self.erzeuge_bild_shape(name, btn, cell,hoehe)
                
                draw_page.add(shape)
                self.bilder_reihe[row][panel_nr]['shape'] = shape
                self.bilder_reihe[row][panel_nr]['hoehe'] = hoehe
                
            
           
    def erzeuge_img_button(self,pos,btn_name,url):
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
            
            if url == '':
                bordercolor = 4147801
            else:
                bordercolor = self.hintergrund
            
            oButton.setPropertyValues(
                                      ('BackgroundColor','Name','Border',
                                       'ScaleImage','ImageURL','BorderColor'),
                                      (self.hintergrund,btn_name,2,
                                       False,url,bordercolor)
                                      )
            
            form.insertByName('btn_'+btn_name,oButton)

            shape = self.calc.createInstance("com.sun.star.drawing.ControlShape")
                        
            p = shape.Position
            p.X = pos[0] 
            p.Y = pos[1]
            shape.setPosition(p)
            
            s = shape.Size
            s.Height = 500
            s.Width = 500
            shape.setSize(s)
            
            shape.Control = oButton

            form.registerScriptEvent( 0, self.event )
            form.addScriptListener(self.listener_icons) 

            return shape
        except:
            log(inspect.stack,tb())
            
    
    def erzeuge_steuerung(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.listener_btn_click = Button_Click_Listener(self.mb,self)            
            
            self.sheet_controller.setFormDesignMode(True)
            draw_page = self.sheet.DrawPage
            self.draw_page = draw_page
            
            shape1 = self.erzeuge_button(LANG.MENU,'Menu',0)
            shape2 = self.erzeuge_button(LANG.INFO,'Info',1)
            shape3 = self.erzeuge_button(LANG.UEBERNEHMEN,'Uebernehmen',2)
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
                    
            URL="private:factory/scalc"
                    
            self.calc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,())
            
            ctx = uno.getComponentContext()
            toolkit = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", self.mb.ctx)    
            self.win = toolkit.getActiveTopWindow() 
            
            frames = self.mb.desktop.Frames
            for i in range(frames.Count):                
                if frames.getByIndex(i).Controller == self.calc.CurrentController:
                    self.calc_frame = frames.getByIndex(i)
                    
            self.win = self.calc_frame.ComponentWindow
            self.sheet_controller = self.calc_frame.Controller
            self.sheet = self.sheet_controller.ActiveSheet
            
            self.calc_frame.setPropertyValue('Title','Organon Organizer')

            RESOURCE_URL = "private:resource/dockingwindow/9809"
            self.calc_frame.LayoutManager.hideElement(RESOURCE_URL)
            
            self.close_listener = Organizer_Close_Listener(self.mb,self,self.calc)
            self.calc.addDocumentEventListener(self.close_listener)
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
            rangex = self.sheet.getCellRangeByPosition( 0, y1+1, self.sheet.Columns.Count-1, self.sheet.Rows.Count-1 )
            rangex.CellBackColor = self.hintergrund
#             
#             # Tabellenhintergruende
#             rangex = self.sheet.getCellRangeByPosition( 0, y1, self.sheet.Columns.Count-1, self.sheet.Rows.Count-1 )
#             rangex.CellBackColor = self.hintergrund
            
            
            y1 = y0
            
            xEnde = x0 + len(self.spalten2) -1

            def set_farben(c,b,hg):
                c.TopBorder = b
                c.BottomBorder = b
                c.LeftBorder = b
                c.RightBorder = b
                c.CellBackColor = hg
            
            # Border    
            cell = self.sheet.getCellByPosition(0,0)
            border = cell.TopBorder
            border.Color = self.hintergrund
            border.OuterLineWidth = 100
            
            # Kopfzeile
            rangex = self.sheet.getCellRangeByPosition(x0,y0,xEnde,y0)
            set_farben(rangex,border,self.menu_hintergrund)
            rangex.CharWeight = 150
            rangex.CharColor = self.menu_schrift
            
            # Zellenabstaende innen            
            abstand = 200
            rangex.ParaLeftMargin = abstand
            rangex.ParaRightMargin = abstand
            rangex.ParaBottomMargin = abstand
            rangex.ParaTopMargin = abstand
            
            # Eintraege
            border.OuterLineWidth = 10
            rangex = self.sheet.getCellRangeByPosition(x0, y0+1, xEnde, y0+len(self.Eintraege))
            set_farben(rangex,border,self.hf_hintergrund)
            # setzt NumberFormat auf Text
            rangex.setPropertyValue('NumberFormat',100)

            # PROTECTION
            prot = cell.CellProtection
            prot.IsLocked = False
            rangex.setPropertyValue('CellProtection',prot)
            
            # Zellenabstaende innen            
            abstand = 200
            rangex.ParaLeftMargin = abstand
            rangex.ParaRightMargin = abstand
            rangex.ParaBottomMargin = abstand
            rangex.ParaTopMargin = abstand
            
            # Zellbreiten
            tags = self.mb.tags
            
            breiten = [1, .6, .6, 4]
            breiten2 = [tags['nr_breite'][nr] for nr in tags['abfolge'] if nr in tags['sichtbare']]    
            breiten.extend(breiten2) 
            
            for b in range(len(breiten)):
                spalte = self.sheet.Columns.getByIndex(b)
                spalte.Width = breiten[b] * 1000

            # WRAP TEXT
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'WrapText'
            prop.Value = True 
            
            self.sheet_controller.select(rangex)
            dispatcher.executeDispatch(self.calc_frame, ".uno:WrapText", "", 0, (prop,))
            
            
            
#             # Bildspalten
#             if self.bild_panels != []:
#                 for v in self.bild_panels:
#                     index = self.spalten.index(v) + 1
#                     rangex = self.sheet.getCellRangeByPosition( x0 + index, y0 +1, x0 + index, self.sheet.Rows.Count-1 )
#                     #rangex.CellBackColor = self.hintergrund
#             
#             #
        except:
            log(inspect.stack,tb())
            
    
    
    def get_eintraege(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB == 'ORGANON':
            Eintraege = self.mb.class_Projekt.lese_xml_datei()
        else:
            Eintraege = self.mb.tabsX.get_tab_Eintraege(T.AB)
        
        weiter = True
        zu_loeschen = []
        papierkorb_ord = self.mb.props[T.AB].Papierkorb
        
        for e in range(len(Eintraege)):
            if Eintraege[e][0] == papierkorb_ord:
                weiter = False
            if weiter:
                Eintraege[e] = list(Eintraege[e])
            else:
                zu_loeschen.append(e)
        
        for e in zu_loeschen:
            del Eintraege[-1]
        
        self.Eintraege = Eintraege
        self.pos[3] = self.pos[1] + len(Eintraege)
        
        
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
                  
            tags_ord = self.tags_org['ordinale']

            
            self.modify_listener = Modify_Listener(self.mb,self,self.Eintraege)
            self.modify_listener.pos = self.pos
            
            
            
            # KOPFZEILE
            for s in range(len(self.spalten2)):
                cell = self.sheet.getCellByPosition(x0 + s ,y0)
                text = self.spalten2[s]
                cell.String = text
                

            from com.sun.star.table.CellVertJustify import CENTER
            
            rangex = self.sheet.getCellRangeByPosition(x0, y0+1, xEnde, y0+len(self.Eintraege))
            rangex.CharColor = self.schrift_datei
            
            tags = self.tags_org
            
            array = []
            
            self.bilder = []
            self.bilder_reihe = {}
            
            # EINTRAEGE
            for e in self.Eintraege:

                olist = [e[2]]
                
                if e[4] != 'pg':
                    cell = self.sheet.getCellByPosition(x0 ,y1 + 1)
                    cell.CharColor = self.schrift_ordner
                
                col = 1
                for panel_nr in self.spalten:
                    text = tags_ord[e[0]][panel_nr]
                    
                    if panel_nr in self.bild_panels:
                        row = y1 + 1
                        cell = self.sheet.getCellByPosition(x0 + col,row)
                        cell.setPropertyValues(('IsTextWrapped','ShrinkToFit','CharColor'),(False,True,self.hf_hintergrund))
                        if text != '':
                            self.bilder.append([ e[0],text,panel_nr,cell ])
                            if row in self.bilder_reihe:
                                self.bilder_reihe[row][panel_nr] = {'values': [ e[0],text,panel_nr,cell ] }
                            else:
                                self.bilder_reihe[row] = {}
                                self.bilder_reihe[row][panel_nr] = {'values': [ e[0],text,panel_nr,cell ] }
                    
                    elif panel_nr in self.date_panels:
                        if text == None:
                            text = ''
                        else:
                            text = self.mb.class_Tags.formatiere_datumdict_nach_text(text)
                            
                    elif isinstance(text, list):
                        if text != []:
                            text = u',\r\n'.join(text)
                        else:
                            text = u''
                    elif text == None:
                        text = ''
                        
                    olist.append(text)
                    col += 1
                    
                array.append(tuple(olist))
                y1 += 1
            
            
            rangex = self.sheet.getCellRangeByPosition(x0, y0+1, x0+len(self.spalten2)-1, y0+len(self.Eintraege))
            self.rangex = rangex   
            
            array = tuple(array)
            self.rangex.setDataArray(array)
            self.data_array = [list(a) for a in array]
            #
            self.rangex.addModifyListener(self.modify_listener)
                       
            self.pos[3] = y1

            self.erzeuge_bilder_reihen()   
            
            # CellVertJustify
            if self.mb.programm == 'LibreOffice':
                self.rangex.setPropertyValue('VertJustify',2)
            else: 
                vertjust = self.rangex.VertJustify
                vertjust.value = 'CENTER'
                self.rangex.setPropertyValue('VertJustify',vertjust)
            
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
            self.listener2 = Icons_Listener(self.mb,self)
            
            sett = self.mb.settings_proj
            visible_tags_tv = sett['tag1'],sett['tag2']
            
            for e in self.Eintraege:
                
                if e[0] == self.mb.props['ORGANON'].Papierkorb:
                    y1 += 1
                    break
                
                if visible_tags_tv[0]:
                    cell = self.sheet.getCellByPosition(x0 - 1 ,y1 + 1)
                    name = 'IMG_{}'.format(e[0])
                      
                    shape1 = self.erzeuge_img_button(berechne_pos(cell),name,e[7])
                    draw_page.add(shape1)
                    self.icons.update({name:shape1})
                    shape1.addEventListener(self.listener2)
                
                if visible_tags_tv[1]:
                    cell = self.sheet.getCellByPosition(x0 - 2 ,y1 + 1)
                    name2 = 'IMGU_{}'.format(e[0])
                      
                    shape2 = self.erzeuge_img_button(berechne_pos(cell),name2,e[8])
                    draw_page.add(shape2)
                    self.icons.update({name2:shape2})
    
                    shape2.addEventListener(self.listener2)
                
                y1 += 1
            
        except:
            log(inspect.stack,tb())
            

    def tags_uebernehmen(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            tags_org = self.tags_org
            to_ord = tags_org['ordinale']
            tags = self.mb.tags
            t_ord = tags['ordinale']
            
            geaenderte = []
            geloeschte = []
            
            # geaenderte Bilder suchen
            for panel_nr in self.bild_panels:
                for ordi in to_ord:
                    
                    if to_ord[ordi][panel_nr] != t_ord[ordi][panel_nr]:
                        if to_ord[ordi][panel_nr] == '':
                            geloeschte.append([ordi,panel_nr])
                        else:
                            geaenderte.append( [ ordi, panel_nr, uno.fileUrlToSystemPath(to_ord[ordi][panel_nr]) ] )
            
            
            for g in geaenderte:
                ordi, panel_nr, pfad = g
                pfad2 = self.mb.class_Sidebar.bild_einfuegen( panel_nr, ordinal=ordi, filepath=pfad, erzeuge_layout=False, loeschen=False)
                tags_org['ordinale'][ordi][panel_nr] = pfad2                
                
            for g in geloeschte:
                ordi, panel_nr = g
                tags_org['ordinale'][ordi][panel_nr] = ''

            
            self.mb.tags = tags_org
            self.mb.class_Sidebar.erzeuge_sb_layout()
            
            for ordi in self.modify_listener.aenderung_dateinamen:
                text = self.modify_listener.aenderung_dateinamen[ordi]
                self.mb.class_Zeilen_Listener.aendere_datei_namen(ordi,text)
            

            Popup(self.mb, zeit=1, parent=self.win).text = LANG.UEBERNOMMEN

            
            self.calc.setModified(False)  
            self.mb.class_Tags.speicher_tags()           
        except:
            log(inspect.stack,tb())
            
        
    def setze_zelle_und_tag(self,c,r,cell,inhalt,ordinal,panel_nr):    
        if self.mb.debug: log(inspect.stack)
        try:
            # Listener entfernen und wieder adden, damit er bei Aenderung
            # der Zelle nicht angesprochen wird.
            self.rangex.removeModifyListener(self.modify_listener)
            
            cell.String = inhalt
            self.data_array[r][c] = inhalt
            self.tags_org['ordinale'][ordinal][panel_nr] = inhalt
            
            self.rangex.addModifyListener(self.modify_listener)
        except:
            log(inspect.stack,tb())
        
    
from com.sun.star.beans import XPropertiesChangeListener      
class Button_Click_Listener(unohelper.Base, XPropertiesChangeListener):

    def __init__(self, mb,Org):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.pressed = False
        self.org = Org
        
    def propertiesChange(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:

            btn = ev[0].Source
            
            # MENULEISTE EIN/AUSBLENDEN
            if btn.Label == LANG.MENU:
                lmgr = self.org.calc_frame.LayoutManager
                
                ResourceURL ='private:resource/menubar/menubar'
                
                if lmgr.isElementVisible(ResourceURL):
                    lmgr.hideElement(ResourceURL)
                else:
                    lmgr.showElement(ResourceURL)
                    
            # INFO ANZEIGEN
            elif btn.Label == LANG.INFO:
                Popup(self.mb, 'info', parent=self.org.win).text = LANG.ORGANIZER_INFO.format(
                         LANG.UEBERNEHMEN,
                         LANG.MENU,
                         LANG.SYNOPSIS,
                         LANG.NOTIZEN,
                         LANG.CHARAKTERE,
                         LANG.ORTE)
                  
            # TAGS UEBERNEHMEN
            else:
                self.org.tags_uebernehmen()
                
        except:
            log(inspect.stack,tb())

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
        self.aenderung_dateinamen = {}
        
        self.new_data_array = []
        
         
    def modified(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            aenderungen = self.finde_veraenderung()
            rows = self.pruefe_veraenderung(aenderungen)
            self.Org.data_array = self.new_data_array
        except:
            log(inspect.stack,tb())

    
    def finde_veraenderung(self):
        if self.mb.debug: log(inspect.stack)
        
        org = self.Org
        
        x0,y0,x1,y1 = self.pos
        rangex = org.sheet.getCellRangeByPosition(x0, 
                                               y0+1, 
                                               x0+len(org.spalten2)-1, 
                                               y0+len(org.Eintraege))
        
        new_data_array = [list(z) for z in rangex.DataArray]
        
        aenderungen = []
        try:
            for d in range(len(new_data_array)):
                if new_data_array[d] != org.data_array[d]:
                    for j in range(len(new_data_array[d])):
                        if new_data_array[d][j] != org.data_array[d][j]:
                            aenderungen.append([ j + x0, d + y0 + 1, new_data_array[d][j] ])
                            
            self.new_data_array = new_data_array
        except:
            log(inspect.stack,tb())
          
        return aenderungen
    
    
    def zelle_nach_liste(self,inhalt):
        if self.mb.debug: log(inspect.stack)
        inhalt = inhalt.replace('\n','').split(',')
        return [i.strip() for i in inhalt if i.strip() != '']
    
    
    def liste_nach_zelle(self,inhalt):
        if self.mb.debug: log(inspect.stack)
        return ',\n'.join(inhalt)
    
    
    def setze_zelle(self,col,row,inhalt):    
        if self.mb.debug: log(inspect.stack)
        
        try:
            tags,c,r,ordinal = self.get_infos(col,row)
            cell = self.Org.sheet.getCellByPosition(col,row)
            
            # Listener entfernen und wieder adden, damit er bei Aenderung
            # der Zelle nicht angesprochen wird.
            self.Org.rangex.removeModifyListener(self.Org.modify_listener)
            
            cell.String = inhalt
            self.new_data_array[r][c] = inhalt
            
            self.Org.rangex.addModifyListener(self.Org.modify_listener)
        except:
            log(inspect.stack,tb())
            
        
    def zuruecksetzen(self,col,row,panel_nr):
        if self.mb.debug: log(inspect.stack)
        
        c,r = col-self.pos[0],row-self.pos[1]-1
        alter_eintrag = self.Org.data_array[r][c]
        self.setze_zelle(col,row,alter_eintrag)        
            
    def get_infos(self,col,row):
        if self.mb.debug: log(inspect.stack)
        
        tags_org = self.Org.tags_org
        c,r = col-self.pos[0],row-self.pos[1]-1
        ordinal = self.eintraege[r][0]
        return tags_org,c,r,ordinal
    
    def setze_txt_panels(self,col,row,panel_nr,ordinal,inhalt):
        if self.mb.debug: log(inspect.stack)
        
        tags,c,r,ordinal = self.get_infos(col,row)
        tags['ordinale'][ordinal][panel_nr] = inhalt
        self.new_data_array[r][c] = inhalt
    
    def setze_zeit_panels(self,col,row,panel_nr,ordinal,inhalt):
        if self.mb.debug: log(inspect.stack)

        neuer_inhalt = self.mb.class_Tags.formatiere_zeit(inhalt)
        if neuer_inhalt == None:
            neuer_inhalt = ''

        self.setze_zelle(col,row,neuer_inhalt)
        self.Org.tags_org['ordinale'][ordinal][panel_nr] = neuer_inhalt
   
    def setze_datum_panels(self,col,row,panel_nr,ordinal,inhalt): 
        if self.mb.debug: log(inspect.stack)
                
        neuer_inhalt,odict = self.mb.class_Tags.formatiere_datum(inhalt)
        if neuer_inhalt == None:
            neuer_inhalt = ''
        
        self.setze_zelle(col,row, neuer_inhalt)
        self.Org.tags_org['ordinale'][ordinal][panel_nr] = odict
     
    def setze_leeren_inhalt(self,col,row,panel_nr,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        try:
            tags_org,c,r,ordinal = self.get_infos(col,row)
            alte_eintraege = self.zelle_nach_liste(self.Org.data_array[r][c])

            if panel_nr in tags_org['sammlung']:
                self.loesche_tags_aus_dict(alte_eintraege,ordinal,panel_nr)
            self.new_data_array[r][c] = ''
        except:
            log(inspect.stack,tb())
            
    def tag_auf_vorkommen_testen(self,tag,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        # pruefen, ob tag noch in anderen dateien benutzt wird
        tags_org = self.Org.tags_org
        
        tag_panels = list(tags_org['sammlung'])
        alle_tags_in_ordinal = {j : 
              [ k for i in tag_panels for k in tags_org['ordinale'][j][i] ] 
              for j in tags_org['ordinale']} 
        
        for ordi in tags_org['ordinale']:
            if ordi != ordinal:
                if tag in alle_tags_in_ordinal[ordi]:
                    return True
        return False
    
    
    def loesche_tags_aus_dict(self,eintrage,ordinal,panel_nr):
        if self.mb.debug: log(inspect.stack)
        
        def loesche(fkt,value):
            try:
                fkt(value)
            except:
                log(inspect.stack,tb())
                
        tags_org = self.Org.tags_org
        
        for eintrag in eintrage:
            # aus den eigenen tags loeschen
            loesche(tags_org['ordinale'][ordinal][panel_nr].remove,eintrag)
            # aus der Sammlung aller Tags loeschen
            if not self.tag_auf_vorkommen_testen(eintrag,ordinal):
                loesche(tags_org['sammlung'][panel_nr].remove,eintrag)     
     

    def tag_hinzufuegen(self,tag,ordinal,panel_nr):
        if self.mb.debug: log(inspect.stack)
        
        tags_org = self.Org.tags_org
        
        if tag not in tags_org['ordinale'][ordinal][panel_nr]:
            tags_org['ordinale'][ordinal][panel_nr].append(tag)
        if tag not in tags_org['sammlung'][panel_nr]:
            tags_org['sammlung'][panel_nr].append(tag)

 
    def setze_aenderung(self,col,row,panel_nr,ordinal,inhalt,geloeschte,hinzugefuegte):
        if self.mb.debug: log(inspect.stack)
        
        try:            
            self.loesche_tags_aus_dict(geloeschte,ordinal,panel_nr)
            for tag in hinzugefuegte:
                self.tag_hinzufuegen(tag,ordinal,panel_nr)
                
            cell = self.Org.sheet.getCellByPosition(col,row)
            i = self.liste_nach_zelle(inhalt)
            self.setze_zelle(col,row,i)
            
        except:
            log(inspect.stack,tb())        
            
        
    def pruefe_veraenderung(self,aenderungen):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            for a in aenderungen:
                
                col,row,inhalt = a
                tags,c,r,ordinal = self.get_infos(col,row)
                
                row_eintrag = row - self.pos[1]
                
                if col - self.pos[0] == 0:
                    panel_nr = 'datei'
                else:
                    panel_nr = self.Org.spalten[col - self.pos[0] - 1]
                
                
                visible_tags_tv = self.mb.settings_proj['tag1'],self.mb.settings_proj['tag2']
 
                if (visible_tags_tv[0] or visible_tags_tv[1]):
                    def bns(r):
                        self.Org.update_bilder_pos(r)
                        self.Org.update_icon_pos(col,r)
                    bild_neu_setzen = bns
                    
                else:
                    def bns2(r):
                        self.Org.update_bilder_pos(r)
                    bild_neu_setzen = bns2
                    

                # PROJEKTNAME
                if (col,row) == (self.Org.pos[0],self.Org.pos[1] + 1):
                    self.zuruecksetzen(col,row,kateg)
                    Popup(self.mb, 'info', parent=self.Org.win).text = LANG.PRJ_NAME_KEINE_AENDERUNG
                    continue
                # DATEINAME
                if panel_nr == 'datei':
                    r = row-self.pos[1]-1
                    self.eintraege[r][2] = inhalt
                    self.aenderung_dateinamen.update({ordinal:inhalt})
                    continue
                # LEERER INHALT
                if inhalt == '':
                    self.setze_leeren_inhalt(col,row,panel_nr,ordinal)
                    bild_neu_setzen(row)
                    continue
                # TEXT PANELS          
                if panel_nr in self.Org.txt_panels:
                    self.setze_txt_panels(col,row,panel_nr,ordinal,inhalt)
                    bild_neu_setzen(row)
                    continue
                # DATUM         
                if panel_nr in self.Org.date_panels:
                    self.setze_datum_panels(col,row,panel_nr,ordinal,inhalt)
                    continue
                # ZEIT        
                if panel_nr in self.Org.time_panels:
                    self.setze_zeit_panels(col,row,panel_nr,ordinal,inhalt)
                    continue
                
                
                # PERSONEN, ORTE, OBJEKTE, ect.
                # Art der Aenderung bestimmen
                inhalt = self.zelle_nach_liste(inhalt)
                
                alte_eintraege = self.zelle_nach_liste(self.Org.data_array[r][c])
                
                set_inhalt = set(inhalt)
                set_alt = set(alte_eintraege)
                
                tags_gen_ohne_kateg = set(a for b in [v for t,v in tags['sammlung'].items() if t != panel_nr] for a in b)
                bereits_in_anderen_tags_vorhanden = len(set_inhalt.intersection(tags_gen_ohne_kateg)) > 0
                
                keine_aenderung = set_inhalt == set_alt

                geloeschte = set_alt.difference(set_inhalt)
                hinzugefuegte = set_inhalt.difference(set_alt)

                if keine_aenderung:
                    pass
                if bereits_in_anderen_tags_vorhanden:
                    Popup(self.mb, 'info', parent=self.Org.win).text = LANG.EIN_TAG_BEREITS_VORHANDEN
                    self.zuruecksetzen(col,row,panel_nr)                   
                else:
                    self.setze_aenderung(col,row,panel_nr,ordinal,inhalt,geloeschte,hinzugefuegte)                    
                    bild_neu_setzen(row)
                        

               
        except:
            log(inspect.stack,tb())
  
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
        return True
   
    def mouseReleased(self, ev):
        if self.mb.debug: log(inspect.stack)
        
        org = self.Org
        try:
            if not hasattr(ev.Target,'CellAddress'):
                return True
            addr = ev.Target.CellAddress
            col,row = addr.Column, addr.Row
            
            try:
                panel_nr = org.spalten[col - org.pos[0] - 1]
            except:
                return True
            
            if panel_nr not in org.bild_panels:
                return True
            elif not (org.pos[1] < row < org.pos[1] + len(org.Eintraege)+1):
                return True
            elif row in org.bilder_reihe and panel_nr in org.bilder_reihe[row]:
                return True

            org.bild_einfuegen(col, row, panel_nr)
            
            if self.mb.programm == 'LibreOffice':
                for br in org.bilder_reihe:
                    org.update_bilder_pos(br)
            else:
                org.update_bilder_pos(row) 
                    
        except:
            log(inspect.stack,tb())
            
        return True
    

                        

from com.sun.star.script import XScriptListener
class Icons_Listener(unohelper.Base,XScriptListener):

    def __init__(self,mb,org):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.org = org

    def firing(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        btn = ev.Arguments[0].Source
        if hasattr(btn.Model,'Label'):
            label = btn.Model.Label
        else:
            label = btn.Model.Name
        # MENULEISTE EIN/AUSBLENDEN
        if label == LANG.MENU:
            lmgr = self.org.calc_frame.LayoutManager
            
            ResourceURL ='private:resource/menubar/menubar'
            
            if lmgr.isElementVisible(ResourceURL):
                lmgr.hideElement(ResourceURL)
            else:
                lmgr.showElement(ResourceURL)
                
        # INFO ANZEIGEN
        elif label == LANG.INFO:

            Popup(self.mb, 'info', parent=self.org.win).text = LANG.ORGANIZER_INFO.format(
                     LANG.UEBERNEHMEN,
                     LANG.MENU,
                     LANG.SYNOPSIS,
                     LANG.NOTIZEN,
                     LANG.CHARAKTERE,
                     LANG.ORTE)

        # TAGS UEBERNEHMEN
        elif label == LANG.UEBERNEHMEN:
            self.org.tags_uebernehmen()
            
        # ICONS AENDERN
        else:
            try:
                ord_source = label.split('_')[2]

                ps = btn.PosSize
                X,Y = ps.X,ps.Y
                
                window_parent = self.org.calc_frame.getComponentWindow()
                
                # FARB ICONS
                if 'IMGU' in label:
                    self.mb.class_Funktionen.erzeuge_Tag2_Container(ord_source,X,Y,window_parent=window_parent)
                elif 'IMG' in label:
                    self.mb.class_Funktionen.erzeuge_Tag1_Container(ord_source,X,Y,window_parent=window_parent)
                elif 'panel' in label:
                    panel_nr = int(label.split('_')[1].split('panel')[1])
                    self.erzeuge_auswahl_bild_tag(X, Y, window_parent,ord_source,panel_nr)
                    
            except:
                log(inspect.stack,tb())
                
            
    def erzeuge_auswahl_bild_tag(self,X,Y,parent,ordinal,panel_nr):       
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            button_listener = Bild_Tag_Button_Listener(self.mb)
                    
            prop_names = ('Label',)
            prop_values = (LANG.LOESCHEN,)
            control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 10, 100, 25, prop_names, prop_values)  
            control.addActionListener(button_listener) 
            control.setActionCommand('loeschen')
            
            prop_names = ('Label',)
            prop_values = (LANG.AENDERN,)
            control2, model2 = self.mb.createControl(self.mb.ctx, "Button", 10, 50, 100, 25, prop_names, prop_values)  
            control2.addActionListener(button_listener) 
            control2.setActionCommand('aendern')
            
            
            posSize = X,Y,120,85
            win,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize,parent=parent)
            cont.addControl('loeschen', control)
            cont.addControl('aendern', control2)
            
            button_listener.win = win
            button_listener.panel_nr = panel_nr
            button_listener.ordinal = ordinal
            
        except:
            log(inspect.stack,tb())
                
    def approveFiring(self,ev): return False
    def disposing(self, ev): return False

        
from com.sun.star.awt import XActionListener
class Bild_Tag_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = None
        self.panel_nr = None
        self.ordinal = None
   
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        cmd = ev.ActionCommand
        
        org = self.mb.class_Organizer
        draw_page = org.sheet.DrawPage
        eintrag = [e for e in org.Eintraege if e[0] == self.ordinal][0]
        
        c = org.spalten.index(self.panel_nr)
        r = org.Eintraege.index(eintrag)
        
        col = org.pos[0] + 1 + c
        row = org.pos[1] + 1 + r
        
        cell = org.sheet.getCellByPosition(col,row)
        
        try:
            if cmd == 'aendern':
   
                url,ok = self.mb.class_Funktionen.filepicker2(url_to_sys=False)
                if not ok:
                    return True
                
                if url:
                    shape = org.bilder_reihe[row][self.panel_nr]['shape']
                    
                    draw_page.remove(shape)
                    shape.dispose()

                    org.bild_einfuegen(col,row,self.panel_nr,url=url)
                    
                    if self.mb.programm == 'LibreOffice':
                        for br in org.bilder_reihe:
                            org.update_bilder_pos(br)
                    else:
                        org.update_bilder_pos(row) 
                        
                    org.update_icon_pos(col,row)
                         
            elif cmd == 'loeschen':
                
                shape = org.bilder_reihe[row][self.panel_nr]['shape']
                
                draw_page.remove(shape)
                shape.dispose()
                 
                del org.bilder_reihe[row][self.panel_nr]
                org.setze_zelle_und_tag(c+1,r,cell,'',self.ordinal,self.panel_nr)

                if self.mb.programm == 'LibreOffice':
                    for br in org.bilder_reihe:
                        org.update_bilder_pos(br)
                else:
                    org.update_bilder_pos(row) 
                
                org.update_icon_pos(col,row)  
                org.tags_org['ordinale'][self.ordinal][self.panel_nr] = ''
                 
        except:
            log(inspect.stack,tb())
            
                        
        self.win.dispose()
        
    def disposing(self,ev):
        return True

            
from com.sun.star.document import XDocumentEventListener
class Organizer_Close_Listener(unohelper.Base,XDocumentEventListener):
    '''
    Lets the Organizer close without warning when no changes had been made.
    Else the user is asked if he wants to save the changes.
    '''
    def __init__(self,mb,Org,calc):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.Org = Org
        self.calc = calc
        
    def documentEventOccured(self,ev):
        try:
            
            if ev.EventName == 'OnPrepareViewClosing':
                if self.mb.debug: log(inspect.stack)
                
                if self.calc.isModified():
                    entscheidung = self.mb.entscheidung(LANG.AENDERUNGEN_NOCH_NICHT_UEBERNOMMEN,"warningbox",16777216,self.mb.doc)
                    # 3 = Nein oder Cancel, 2 = Ja
                    if entscheidung == 3:
                        pass
                    elif entscheidung == 2:
                        try:
                            self.Org.tags_uebernehmen()
                        except:
                            log(inspect.stack,tb())
                            
                        
                tags = self.mb.tags
                bild_panels = self.Org.bild_panels
                
                for root, dirs, files in os.walk(self.mb.pfade['images']):
                    break
                
                genutzte_bilder = []
                
                for ordi in tags['ordinale']:
                    for panel in bild_panels:
                        if tags['ordinale'][ordi][panel] not in genutzte_bilder:
                            genutzte_bilder.append(tags['ordinale'][ordi][panel])
                            
                ungenutzte_bilder = [b for b in files if b not in genutzte_bilder]
                                            
                if ungenutzte_bilder != []: 
                    # Bilder können nicht geloescht werden,
                    # solange der Organizer geoeffnet ist.
                    # Daher wird loeschen 3 Sekunden nach
                    # Schliessen des Organizers aufgerufen
                    from threading import Thread
                     
                    def sleeper(zu_loeschende_bilder,mb,os):  
                        import time
                        time.sleep(3) 
                        for b in zu_loeschende_bilder:
                            path = os.path.join(mb.pfade['images'],b)
                            os.remove(path)
         
                    t = Thread(target=sleeper,args=(ungenutzte_bilder,self.mb,os))
                    t.start()
                
                self.calc.setModified(False)
                
                

        except:
            log(inspect.stack,tb())
            
    



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
        
        
        