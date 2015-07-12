# -*- coding: utf-8 -*-

import unohelper
import pickle
from pickle import HIGHEST_PROTOCOL
from pickle import load as pickle_load
from pickle import dump as pickle_dump

class Sidebar(): 
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb     
        self.design_gesetzt = False   
        self.mb.dict_sb['erzeuge_sb_layout'] = self.erzeuge_sb_layout
        self.mb.dict_sb['optionsfenster'] = self.optionsfenster
        self.mb.dict_sb['dict_sb_zuruecksetzen'] = self.dict_sb_zuruecksetzen
        
        
        self.sb_panels_tup = ('Synopsis',
            'Notes',
            'Images',
            'Tags_general',
            'Tags_characters',
            'Tags_locations',
            'Tags_objects',
            'Tags_time',
            'Tags_user1',
            'Tags_user2',
            'Tags_user3')
        
        self.sb_panels1 = {'Synopsis':LANG.SYNOPSIS,
            'Notes':LANG.NOTIZEN,
            'Images':LANG.BILDER,
            'Tags_general':LANG.ALLGEMEIN,
            'Tags_characters':LANG.CHARAKTERE,
            'Tags_locations':LANG.ORTE,
            'Tags_objects':LANG.OBJEKTE,
            'Tags_time':LANG.ZEIT,
            'Tags_user1':LANG.BENUTZER1,
            'Tags_user2':LANG.BENUTZER2,
            'Tags_user3':LANG.BENUTZER3}
        
        
        
        self.sb_panels2 = {LANG.SYNOPSIS:'Synopsis',
            LANG.NOTIZEN:'Notes',
            LANG.BILDER:'Images',
            LANG.ALLGEMEIN:'Tags_general',
            LANG.CHARAKTERE:'Tags_characters',
            LANG.ORTE:'Tags_locations',
            LANG.OBJEKTE:'Tags_objects',
            LANG.ZEIT:'Tags_time',
            LANG.BENUTZER1:'Tags_user1',
            LANG.BENUTZER2:'Tags_user2',
            LANG.BENUTZER3:'Tags_user3'}
        
        self.sb_tags = ('Tags_general',
                        'Tags_characters',
                        'Tags_locations',
                        'Tags_objects',
                        'Tags_user1',
                        'Tags_user2',
                        'Tags_user3')

       
    
    
    
    def lade_sidebar(self):
        if self.mb.debug: log(inspect.stack)

        if 'empty_project' in self.mb.dict_sb['sichtbare']:
            self.mb.dict_sb['sichtbare'].remove('empty_project')
        
        self.mb.dict_sb['sichtbare'] = self.mb.dict_sb_content['sichtbare']
                      
      
    def passe_sb_an(self,textbereich):
        if self.mb.debug: log(inspect.stack)
        
        ordinal = textbereich.Context.Model.Text
        
        for panel in self.mb.dict_sb['sichtbare']:
            self.erzeuge_sb_layout(panel)

    
    def lege_dict_sb_content_an(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            sb_panels = self.sb_panels_tup
            
            tags = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
            
            
            dict_sb_content = {}
            dict_sb_content.update({'ordinal':{}})
            dict_sb_content.update({'tags':{}})
            # 'sichtbare' definiert hier die Ansicht bei einem neuen Projekt
            dict_sb_content.update({'sichtbare':['Synopsis','Notes','Tags_general','Tags_characters']})
            
            # Anlegen der Verzeichnisstruktur 'ordinal'
            for ordinal in list(self.mb.props[T.AB].dict_bereiche['ordinal']):
                dict_sb_content['ordinal'].update({ordinal:{}})
                for panel in sb_panels:
                    if panel in tags:
                        dict_sb_content['ordinal'][ordinal].update({panel:[]})
                        if panel not in dict_sb_content['tags']:
                            dict_sb_content['tags'].update({panel:[]})
                        
                    else:
                        if panel == 'Tags_time':
                            dict = {}
                            dict.update({'zeit':None})
                            dict.update({'datum':None})
                            dict_sb_content['ordinal'][ordinal].update({panel:dict})
                            
                        else:
                            dict_sb_content['ordinal'][ordinal].update({panel:''})
                        
            
            dict_sb_content.update({'einstellungen':{}})   
            dict_sb_content['einstellungen'].update({'tags_general_loescht_im_ges_dok':0})      
                    
            self.mb.dict_sb_content = dict_sb_content
        except:
            log(inspect.stack,tb())
    
    def lege_dict_sb_content_ordinal_an(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        sb_panels = self.sb_panels_tup
        tags = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
            
        # Anlegen der Verzeichnisstruktur 'ordinal'
        
        self.mb.dict_sb_content['ordinal'].update({ordinal:{}})
        for panel in sb_panels:
            if panel in tags:
                self.mb.dict_sb_content['ordinal'][ordinal].update({panel:[]})
            else:
                if panel == 'Tags_time':
                    dict = {}
                    dict.update({'zeit':None})
                    dict.update({'datum':None})
                    self.mb.dict_sb_content['ordinal'][ordinal].update({panel:dict})
                else:
                    self.mb.dict_sb_content['ordinal'][ordinal].update({panel:''})

    
    def loesche_dict_sb_content_eintrag(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        bildUrl = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
        del(self.mb.dict_sb_content['ordinal'][ordinal])
        
        # Wenn das Bild im Dok nicht mehr vorkommt: loeschen
        if bildUrl != '':
            vorkommen = False
            for ordi in self.mb.dict_sb_content['ordinal']:
                if bildUrl in self.mb.dict_sb_content['ordinal'][ordi]['Images']:
                    vorkommen = True
            if not vorkommen:
                os.remove(uno.fileUrlToSystemPath(bildUrl))
        
        
        
    def speicher_sidebar_dict(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
            with open(pfad, 'wb') as f:
                pickle_dump(self.mb.dict_sb_content, f,2)
        except:
            log(inspect.stack,tb())

    def lade_sidebar_dict(self):
        if self.mb.debug: log(inspect.stack)
        
        pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
        
        dict_exists, backup_exists = False, False

        if os.path.exists(pfad):
            dict_exists = True
        if not os.path.exists(pfad+'.Backup'):
            backup_exists = False
        
        
        try:  
            if dict_exists:          
                with open(pfad, 'rb') as f:
                    self.mb.dict_sb_content =  pickle_load(f)
                self.ueberpruefe_dict_sb_content(not backup_exists)
                
            elif not backup_exists:
                self.lege_dict_sb_content_an()
                self.lade_Backup()
        except:
            log(inspect.stack,tb())
            
            self.lege_dict_sb_content_an()
            if not backup_exists:
                self.lade_Backup()
            
        if not dict_exists:
            self.speicher_sidebar_dict()

        self.erzeuge_dict_sb_content_Backup()
        
        
    def ueberpruefe_dict_sb_content(self,backup_exists):
        if self.mb.debug: log(inspect.stack)
        
        fehlende = []

        for ordinal in list(self.mb.props['Projekt'].dict_bereiche['ordinal']):
            if ordinal not in self.mb.dict_sb_content['ordinal']:
                fehlende.append(ordinal)

        if len(fehlende) > 0:
            if backup_exists:
                self.lade_Backup(fehlende)
            else:
                for f in fehlende:
                    self.lege_dict_sb_content_ordinal_an(f)
                      
    
    def lade_Backup(self,fehlende = 'all'): 
        if self.mb.debug: log(inspect.stack,None,'### Attention ###, sidebar_content.pkl.backup loaded!')
        
        pfad_Backup = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl.Backup')
        
        if fehlende == 'all':
            fehlende = list(self.mb.props['Projekt'].dict_bereiche['ordinal'])
        
        try:
            with open(pfad_Backup, 'rb') as f:
                backup = pickle_load(f)
        except:
            log(inspect.stack,tb())

            for f in fehlende:
                self.lege_dict_sb_content_ordinal_an(f)
            return

        helfer = fehlende[:]  
        
        if backup != None:  
            for f in fehlende:
                if f in backup['ordinal']:
                    self.mb.dict_sb_content['ordinal'].update(backup['ordinal'][f])
                    helfer.remove(f)
        
        fehlende = helfer      
        for f in fehlende:
            self.lege_dict_sb_content_ordinal_an(f)
               
        
     

         
       
    def erzeuge_dict_sb_content_Backup(self):
        if self.mb.debug: log(inspect.stack)
        
        pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
        pfad_Backup = pfad + '.Backup'
        from shutil import copy2
        copy2(pfad, pfad_Backup)
        
    def erzeuge_sb_layout(self,xUIElement_name,rufer = None):
        
        if not self.design_gesetzt:
            if self.mb.settings_orga['organon_farben']['design_office']:
                self.design_gesetzt = True
                self.setze_sidebar_design()
        
        if xUIElement_name == 'empty_project':
            return
        if self.mb.dict_sb['sb_closed']:
            return

        if self.mb.debug: log(inspect.stack)
        
                        
        # Wenn die Sidebar noch nicht geoeffnet wurde,
        # sind noch keine Panels vorhanden
        if xUIElement_name not in self.mb.dict_sb['controls']:
            return
                
        try:
            ctx = self.mb.ctx
            
            xUIElement = self.mb.dict_sb['controls'][xUIElement_name][0]
            sb = self.mb.dict_sb['controls'][xUIElement_name][1]
            panelWin = xUIElement.Window
            
            
            # alte Eintraege im einzelnen Panel vorher loeschen
            if rufer == 'focus_lost':
                pass
            elif rufer != 'factory':
                for conts in panelWin.Controls:
                    conts.dispose()
            
            
            
            ######################################
            #                TAGS                #
            ######################################            
            
            ordinal = self.mb.props[T.AB].selektierte_zeile
            
            if xUIElement_name in self.sb_tags:
                
                remove_or_add_button_listener = Tags_Remove_Button_Listener(self.mb)

                y = 10
                height = 20
                
                prop_names = ('HelpText','MultiLine')
                prop_values = (LANG.ENTER_NEW_TAG,True)
                control, model = self.mb.createControl(ctx, "Edit", 170, y,100, height, prop_names, prop_values)
                panelWin.addControl('Button', control) 

                key_listener = Tags_Key_Listener(self.mb,xUIElement_name)
                control.addKeyListener(key_listener)

                
                #y += height + 10
                y_all_tags = y
                y_all_tags += 30
                
                alle_tags = self.mb.dict_sb_content['tags'][xUIElement_name]
                eintraege = self.mb.dict_sb_content['ordinal'][ordinal][xUIElement_name]
                
                SEP = '_XYX_'
                
                # Leiste mit allen Tags
                for tag_eintrag in alle_tags:
                    if tag_eintrag not in eintraege:
                        
                        height = 18
                        # Tageintrag
                        prop_names = ('Label','Align','TextColor')
                        prop_values = (tag_eintrag,2, 7303024)
                        control, model = self.mb.createControl(ctx, "FixedText", 170, y_all_tags ,100, height, prop_names, prop_values)
                        panelWin.addControl(tag_eintrag, control) 
                        
                        # Add Button
                        prop_names = ('Label','HelpText')
                        prop_values = ('',LANG.TAG_HINZUFUEGEN)
                        control, model = self.mb.createControl(ctx, "Button", 280, y_all_tags ,14, 14, prop_names, prop_values)
                        panelWin.addControl('Remove_'+tag_eintrag, control) 
                        control.setActionCommand(xUIElement_name + SEP + ordinal + SEP + tag_eintrag + SEP + 'hinzufuegen')
                        control.addActionListener(remove_or_add_button_listener)


                        y_all_tags += height 
                
                # Tags
                
                for tag_eintrag in eintraege:
                    
                    height = 18
                    # Tageintrag
                    prop_names = ('Label',)
                    prop_values = (tag_eintrag,)
                    control, model = self.mb.createControl(ctx, "FixedText", 10, y,100, height, prop_names, prop_values)
                    panelWin.addControl(tag_eintrag, control) 
                    
                    
                    if xUIElement_name == 'Tags_general':
                        if self.mb.dict_sb_content['einstellungen']['tags_general_loescht_im_ges_dok'] == 0:
                            helptext = LANG.TAGS_IN_AKT_DAT_LOESCHEN
                        else:
                            helptext = LANG.TAGS_IM_GES_DOK_LOESCHEN
                    else:
                        helptext = LANG.TAG_LOESCHEN
                    # Remove Button
                    prop_names = ('Label','HelpText')
                    prop_values = ('X',helptext)
                    control, model = self.mb.createControl(ctx, "Button", 115, y,14, 14, prop_names, prop_values)
                    panelWin.addControl('Remove_'+tag_eintrag, control) 
                    control.setActionCommand(xUIElement_name + SEP + ordinal + SEP + tag_eintrag + SEP + 'loeschen')
                    control.addActionListener(remove_or_add_button_listener)
                    
                    
                    y += height + 3
                
                
                
                
                if y<y_all_tags:
                    y = y_all_tags
                
                if len(eintraege) == 0:
                    y_all_tags += 21
                
                height = y + 10

            ######################################
            # Synopsis, Notes, Images, Tags_time #
            ######################################                     
            
            elif xUIElement_name == 'Synopsis':
                
                pos_y = 10
                height = 0 
                width = 282 
                 
                text = self.mb.dict_sb_content['ordinal'][ordinal]['Synopsis']
                 
                prop_names = ('MultiLine','Text','MaxTextLen')
                prop_values = (True,text,5000)
                control, model = self.mb.createControl(self.mb.ctx, "Edit", 10, pos_y, width, height, prop_names, prop_values)  
                panelWin.addControl('Synopsis', control)
                
                mS = control.getMinimumSize()
                hoehe = mS.Height + 44

                control.setPosSize(0,0,0,hoehe,8)
                
                listener = Text_Change_Listener_Synopsis(self.mb)
                model.addPropertyChangeListener('Text',listener)
                
                height = hoehe + 20
                
            elif xUIElement_name == 'Notes':
                pos_y = 10
                height = 0 
                width = 282 
                 
                text = self.mb.dict_sb_content['ordinal'][ordinal]['Notes']
                 
                prop_names = ('MultiLine','Text','MaxTextLen')
                prop_values = (True,text,5000)
                control, model = self.mb.createControl(self.mb.ctx, "Edit", 10, pos_y, width, height, prop_names, prop_values)  
                panelWin.addControl('Notes', control)
                
                mS = control.getMinimumSize()
                hoehe = mS.Height + 44

                control.setPosSize(0,0,0,hoehe,8)
                
                listener = Text_Change_Listener_Notizen(self.mb)
                model.addPropertyChangeListener('Text',listener)

                height = hoehe + 20
                

                
            elif xUIElement_name == 'Images':
                pos_y = 10
                height = 200
                

                prop_names = ()
                prop_values = ()
                control, model = self.mb.createControl(self.mb.ctx, "ImageControl", 10, pos_y, 284, height, prop_names, prop_values)  
                panelWin.addControl('Image', control)
                try:
                    if self.mb.dict_sb_content['ordinal'][ordinal]['Images'] != '':
                        model.ImageURL = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
                        breite = self.berechne_bildgroesse(model,height)
                        control.setPosSize(0,0,breite,0,4)
                except:
                    log(inspect.stack,tb())
                
                
                
                height += 20
               
            elif xUIElement_name == 'Tags_time':
                
                try:
                    Background_Color = xUIElement.Window.AccessibleContext.Background
                except:
                    Background_Color = 14804725
                
                if self.mb.language == 'de':
                    date_format = 7
                    time_format = 0
                else:
                    date_format = 8
                    time_format = 2
                
                pos_y = 10
                pos_x = 10
                pos_x2 = 65
                pos_x3 = 170
                y = 20
                
                focus_listener = Tag_Time_Key_Listener(self.mb)
                
                zeit = self.mb.dict_sb_content['ordinal'][ordinal]['Tags_time']['zeit']
                leere_zeit = None
                if self.mb.programm == 'LibreOffice':
                    zeit = self.in_time_struct_wandeln(zeit)
                    leere_zeit = self.in_time_struct_wandeln(leere_zeit)
                    
                prop_names = ('Label',)
                prop_values = (LANG.ZEIT2,)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x, pos_y, 70, y, prop_names, prop_values)  
                panelWin.addControl('Time', control)
                
                prop_names = ('Time','TimeFormat','StrictFormat','Border')
                prop_values = (zeit,time_format,True,0)
                control, model = self.mb.createControl(self.mb.ctx, "TimeField", pos_x2, pos_y, 70, y, prop_names, prop_values)  
                panelWin.addControl('Time', control)
                model.BackgroundColor = Background_Color


                prop_names = ('Time','TimeFormat','StrictFormat')
                prop_values = (leere_zeit,time_format,True)
                control, model = self.mb.createControl(self.mb.ctx, "TimeField", pos_x3, pos_y, 60, y, prop_names, prop_values)  
                panelWin.addControl('Time', control)
                
                control.addKeyListener(focus_listener)
                
                
                pos_y += 30
                
                prop_names = ('Label',)
                prop_values = (LANG.DATUM,)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x, pos_y, 70, y, prop_names, prop_values)  
                panelWin.addControl('Datum_Label', control)
                
                datum = self.mb.dict_sb_content['ordinal'][ordinal]['Tags_time']['datum'] 
                
                prop_names = ('Label',)
                prop_values = (datum,)
                control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x2, pos_y, 70, y, prop_names, prop_values)  
                panelWin.addControl('Time', control)
  
                prop_names = ('Text',)
                prop_values = ('',)
                control, model = self.mb.createControl(self.mb.ctx, "Edit", pos_x3, pos_y, 60, y, prop_names, prop_values)  
                panelWin.addControl('Time', control)

                control.addKeyListener(focus_listener)                 
                
                height = 70
           
                
            xUIElement.height = height
            sb.requestLayout()
            
            
            
            # zum Speichern und Wiederherstellen der sichtbaren Panels
            self.mb.dict_sb_content['sichtbare'] = self.mb.dict_sb['sichtbare']
            
        except:
            log(inspect.stack,tb())        
    def setze_sidebar_design(self):  
        
        try:
            
            def get_farbe(value):
                if isinstance(value, int):
                    return value
                else:
                    return self.mb.settings_orga['organon_farben'][value]
            
            keys = [k for k in self.mb.dict_sb['controls']]
            if len(keys) == 0:
                self.design_gesetzt = False
                return
            
            personen = self.mb.dict_sb['controls'][keys[0]]
            theme = personen[0].Theme
    
            
            sett = self.mb.settings_orga['organon_farben']['office']['sidebar']
            
            hintergrund = get_farbe(sett['hintergrund'])
            
            titel_hintergrund = get_farbe(sett['titel_hintergrund'])
            titel_schrift = get_farbe(sett['titel_schrift'])
            
            panel_titel_hintergrund = get_farbe(sett['panel_titel_hintergrund'])
            panel_titel_schrift = get_farbe(sett['panel_titel_schrift'])
    
            leiste_hintergrund = get_farbe(sett['leiste_hintergrund'])
            leiste_selektiert_hintergrund = get_farbe(sett['leiste_selektiert_hintergrund'])
            leiste_icon_umrandung = get_farbe(sett['leiste_icon_umrandung'])
            
            border_horizontal = get_farbe(sett['border_horizontal'])
            border_vertical = get_farbe(sett['border_vertical'])
            
            # folgende muessen bereits in factory.py gesetzt werden
            # selected_schrift, eigene_fenster_hintergrund, schrift, selected_hintergrund
            
            
            # Tabbar  
            theme.setPropertyValue('Paint_TabBarBackground', leiste_hintergrund)
            theme.setPropertyValue('Paint_TabItemBackgroundNormal', leiste_hintergrund)
            theme.setPropertyValue('Color_TabMenuSeparator', leiste_hintergrund)
            theme.setPropertyValue('Paint_TabItemBackgroundHighlight', leiste_selektiert_hintergrund)
            theme.setPropertyValue('Color_TabItemBorder', leiste_icon_umrandung)            
            theme.setPropertyValue('Int_ButtonCornerRadius', 0)
            
            # Hintergruende
            theme.setPropertyValue('Paint_PanelBackground', hintergrund)
            theme.setPropertyValue('Paint_DeckBackground', hintergrund)
            theme.setPropertyValue('Paint_DeckTitleBarBackground', titel_hintergrund)
            
            tbb = theme.Paint_PanelTitleBarBackground
            tbb.StartColor = panel_titel_hintergrund
            tbb.EndColor = panel_titel_hintergrund
            theme.setPropertyValue('Paint_PanelTitleBarBackground', tbb)
    
            # Borders
            theme.setPropertyValue('Paint_HorizontalBorder', border_horizontal)
            theme.setPropertyValue('Paint_VerticalBorder', border_vertical)
            
            # Schriften
            theme.setPropertyValue('Color_DeckTitleFont', titel_schrift)
            theme.setPropertyValue('Color_PanelTitleFont', panel_titel_schrift)
            
    
            theme.setPropertyValue('Paint_ToolBoxBorderBottomRight', hintergrund) # buttons Umrandung
            theme.setPropertyValue('Paint_ToolBoxBorderTopLeft', hintergrund) # buttons Umrandung
    
            tbb = theme.Paint_ToolBoxBackground # buttons Hintergrund
            tbb.StartColor = hintergrund 
            tbb.EndColor = hintergrund
            theme.setPropertyValue('Paint_ToolBoxBackground', tbb)
            
    #         theme.setPropertyValue('Paint_DropDownBackground', rot)
    #         theme.setPropertyValue('Paint_ToolBoxBorderCenterCorners', rot) #??? 
    #         theme.setPropertyValue('Color_Highlight', rot)
    #         theme.setPropertyValue('Color_HighlightText', rot)
    #         theme.setPropertyValue('Color_DropDownBorder', rot)
        
            sb = personen[1]
            sb.requestLayout()
        
        except:
            log(inspect.stack,tb())
            
        

       
           
    
    def schalte_sidebar_button(self):
        if self.mb.debug: log(inspect.stack)

        try:
            controls = self.mb.dict_sb['controls']
            okey = list(controls)[0]
            window = controls[okey][0].window

            a = window.AccessibleContext.AccessibleParent
            b = a.AccessibleContext.AccessibleParent
            c = b.AccessibleContext.AccessibleParent
            d = c.AccessibleContext.AccessibleParent
            e = d.AccessibleContext.AccessibleParent
            
            Seitenleiste_fenster = e.Windows[0]

            for window in Seitenleiste_fenster.Windows:
                if window.AccessibleContext.AccessibleDescription == 'Organon':
                    o_button = window
                else:
                    n_button = window
 
            n_button.setState(True)
            o_button.setState(True)
            
        except:
            log(inspect.stack,tb())

        
        
    def berechne_bildgroesse(self,model,hoehe):
        if self.mb.debug: log(inspect.stack)
        
        try:
            HOEHE = model.Graphic.Size.Height
            BREITE = model.Graphic.Size.Width

            quotient = float(BREITE)/float(HOEHE)
            BREITE = int(hoehe*quotient)
            return BREITE
            
        except:
            log(inspect.stack,tb())
    
    def dict_sb_zuruecksetzen(self):
        if self.mb.debug: log(inspect.stack)
        
        self.mb.dict_sb['sichtbare']  = ['empty_project'] 
        self.mb.dict_sb['controls'] = {}
        self.mb.dict_sb['erzeuge_Layout'] = None   
        
        
    def optionsfenster(self,cmd):
        if self.mb.debug: log(inspect.stack)
        
        loc_x = self.mb.dict_sb['controls'][cmd][0].Window.Peer.AccessibleContext.LocationOnScreen.X
        loc_y = self.mb.dict_sb['controls'][cmd][0].Window.Peer.AccessibleContext.LocationOnScreen.Y

        if cmd =='Tags_general':
            self.optionsfenster_tags_general(loc_x,loc_y)
        elif cmd =='Images':
            self.optionsfenster_images(loc_x,loc_y)
                   
            
    def optionsfenster_images(self,loc_x,loc_y):
        if self.mb.debug: log(inspect.stack)
        
        win,cont = self.mb.erzeuge_Dialog_Container((loc_x - 20,loc_y,350,110))
 
        listener = Options_Tags_General_And_Images_Listener(self.mb,win)
        
        prop_names = ('Label',)
        prop_values = (LANG.IN_PROJEKTORDNER_IMPORTIEREN,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 340, 20, prop_names, prop_values)  
        cont.addControl('Titel', control)
        
        prop_names = ('Label',)
        prop_values = (LANG.BILD_EINFUEGEN,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 30, 120, 30, prop_names, prop_values)  
        cont.addControl('Titel', control)
        control.setActionCommand('bild_einfuegen')
        control.addActionListener(listener)
        
        prop_names = ('Label',)
        prop_values = (LANG.BILD_LOESCHEN,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 70, 120, 30, prop_names, prop_values)  
        cont.addControl('Titel2', control)
        control.setActionCommand('bild_loeschen')
        control.addActionListener(listener)
        
               
    def optionsfenster_tags_general(self,loc_x,loc_y):
        if self.mb.debug: log(inspect.stack)
        
        win,cont = self.mb.erzeuge_Dialog_Container((loc_x - 20,loc_y,350,80))
        try:
            
            state = self.mb.dict_sb_content['einstellungen']['tags_general_loescht_im_ges_dok']
            
            if state == 1:
                state2 = 0
            else:
                state2 = 1
            
            listener = Options_Tags_General_And_Images_Listener(self.mb)
            
            prop_names = ('Label',)
            prop_values = (LANG.EINSTELLUNGEN_TAGS_GENERAL,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 28, 10, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            
            prop_names = ('Label','State')
            prop_values = (LANG.TAGS_IM_GES_DOK_LOESCHEN,state)
            control, model = self.mb.createControl(self.mb.ctx, "RadioButton", 10, 30, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            control.setActionCommand('1')
            control.addActionListener(listener)
            
            prop_names = ('Label','State')
            prop_values = (LANG.TAGS_IN_AKT_DAT_LOESCHEN,state2)
            control, model = self.mb.createControl(self.mb.ctx, "RadioButton", 10, 50, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            control.setActionCommand('0')
            control.addActionListener(listener)

        except:
            log(inspect.stack,tb())
            
    def in_time_struct_wandeln(self,zeit):
        if self.mb.debug: log(inspect.stack)
        
        prop = uno.createUnoStruct("com.sun.star.util.Time")
        
        if zeit == None:
            prop.Hours = 0
            prop.Minutes = 0
            prop.Seconds = 0
            
        else:
            zeit_str = str(zeit)
    
            prop.Hours = int(zeit_str[0:2])
            prop.Minutes = int(zeit_str[2:4])
            prop.Seconds = int(zeit_str[4:6])

        return prop
    
    def in_date_struct_wandeln(self,datum):
        if self.mb.debug: log(inspect.stack)
        
        prop = uno.createUnoStruct("com.sun.star.util.Date")
        
        if datum == None:
            prop.Year = 0
            prop.Month = 1
            prop.Day = 1
        else:
            date_str = str(datum)

            prop.Year = int(date_str[0:4])
            prop.Month = int(date_str[4:6])
            prop.Day = int(date_str[6:8])

        return prop
    
    def date_time_struct_nach_long_wandeln(self,prop,attribute):
        if self.mb.debug: log(inspect.stack)
        
        if attribute == 'zeit':
            stunden = self.pruefe_format(str(prop.Hours),2)
            minuten = self.pruefe_format(str(prop.Minutes),2)
            sekunden = self.pruefe_format(str(prop.Seconds),2)
            nano = '00'
            
            value = int(stunden+minuten+sekunden+nano)
        
#         if attribute == 'datum':
#             jahr = self.pruefe_format(str(prop.Year),4)
#             monat = self.pruefe_format(str(prop.Month),2)
#             tag = self.pruefe_format(str(prop.Day),2)
#             
#             value = int(jahr+monat+tag)
        
        return value
        
    def pruefe_format(self,value,length):
        if self.mb.debug: log(inspect.stack)
        
        if len(value) != length:
            for i in range(length-len(value)):
                value = '0' + value
        return value

   
from com.sun.star.beans import XPropertyChangeListener
class Text_Change_Listener_Synopsis(unohelper.Base, XPropertyChangeListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def propertyChange(self,ev):
        ordinal = self.mb.props[T.AB].selektierte_zeile
        self.mb.dict_sb_content['ordinal'][ordinal]['Synopsis'] = ev.NewValue
          
          
class Text_Change_Listener_Notizen(unohelper.Base, XPropertyChangeListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb

    def propertyChange(self,ev):
        ordinal = self.mb.props[T.AB].selektierte_zeile
        self.mb.dict_sb_content['ordinal'][ordinal]['Notes'] = ev.NewValue


from com.sun.star.awt import XActionListener
class Options_Tags_General_And_Images_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,win = None):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.win = win
        
    def disposing(self,ev):return False
    
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        # optionsfenster_tags_general
        if ev.ActionCommand in ('1','0'):
            self.mb.dict_sb_content['einstellungen']['tags_general_loescht_im_ges_dok'] = int(ev.ActionCommand) 
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','Options_Tags_General_And_Images_Listener')
            first_element = list(self.mb.dict_sb['controls'])[0]
            self.mb.dict_sb['controls'][first_element][1].requestLayout()
        # optionsfenster_images
        elif ev.ActionCommand == 'bild_einfuegen':
            self.bild_einfuegen()
        elif ev.ActionCommand == 'bild_loeschen':
            self.bild_loeschen_a()
            
    def bild_einfuegen(self):
        if self.mb.debug: log(inspect.stack)
        
        self.win.dispose()
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
#         if filepath != None:
#             Filepicker.setDisplayDirectory(filepath)
        Filepicker.execute()
        
        if Filepicker.Files == '':
                return

        filepath =  Filepicker.Files[0]
        basename = os.path.basename(filepath)
        
        path_image_folder = self.mb.pfade['images']
        complete_path = os.path.join(path_image_folder,basename)
        path = uno.systemPathToFileUrl(complete_path)

        if not os.path.exists(path):
            
            sys_filepath = uno.fileUrlToSystemPath(filepath)
            sys_path = uno.fileUrlToSystemPath(path)
            
            from shutil import copy2
            copy2(sys_filepath, sys_path)
        
        ordinal = self.mb.props[T.AB].selektierte_zeile
        
        old_image_path = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
        
        self.mb.dict_sb_content['ordinal'][ordinal]['Images'] = path
        self.mb.class_Sidebar.erzeuge_sb_layout('Images','Options_Tags_General_And_Images_Listener')
        
        if old_image_path != '' and old_image_path != path:
            self.bild_loeschen(old_image_path)

            
    def bild_loeschen_a(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinal = self.mb.props[T.AB].selektierte_zeile
        old_image_path = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
        self.mb.dict_sb_content['ordinal'][ordinal]['Images'] = ''
        self.bild_loeschen(old_image_path)
        self.mb.class_Sidebar.erzeuge_sb_layout('Images','Options_Tags_General_And_Images_Listener')
        self.win.dispose()

    def bild_loeschen(self,old_image_path):
        if self.mb.debug: log(inspect.stack)
        
        try:
            vorhanden = False
            for ordinal in self.mb.dict_sb_content['ordinal']:
                if old_image_path == self.mb.dict_sb_content['ordinal'][ordinal]['Images']:
                    vorhanden = True
                    
            if not vorhanden:
                os.remove(uno.fileUrlToSystemPath(old_image_path))
        except:
            log(inspect.stack,tb())
            
        


from com.sun.star.awt import XFocusListener,XKeyListener
class Tags_Key_Listener(unohelper.Base, XKeyListener):
    def __init__(self,mb,tag):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.tag = tag
    
    def keyPressed(self,ev):
        return False
        
    def keyReleased(self,ev):
        # Hinzufuegen neuer Tags
        
        if ev.KeyCode != 1280:
            return
        # Nur nach einer Eingabe (=1280) loggen
        if self.mb.debug: log(inspect.stack)
        
        ordinal = self.mb.props[T.AB].selektierte_zeile
        new_tag = ev.Source.Model.Text.replace('\n','')
        
        if new_tag != '':
            if new_tag not in self.mb.dict_sb_content['tags']['Tags_general']:
                self.mb.dict_sb_content['tags']['Tags_general'].append(new_tag)
                
                if new_tag not in self.mb.dict_sb_content['ordinal'][ordinal][self.tag]:
                    self.mb.dict_sb_content['ordinal'][ordinal][self.tag].append(new_tag)
                
                if new_tag not in self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general']:
                    self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general'].append(new_tag)
                
                if new_tag not in self.mb.dict_sb_content['tags'][self.tag]:
                    self.mb.dict_sb_content['tags'][self.tag].append(new_tag)
            
            
                 
            ev.Source.Model.Text = ''

            self.mb.class_Sidebar.erzeuge_sb_layout(self.tag,'focus_lost')
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','focus_lost')


    def disposing(self,ev):pass


    
class Tag_Time_Key_Listener(unohelper.Base, XKeyListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
    
    def keyPressed(self,ev):
        return False
        
    def keyReleased(self,ev):
        # 1280 = Return
        if ev.KeyCode != 1280:
            return False
        # Nur bei Tasteneingabe Return loggen
        if self.mb.debug: log(inspect.stack)
        
        self.mb.doc.UndoManager.undo()
        try:
            ordinal = self.mb.props[T.AB].selektierte_zeile

            if hasattr(ev.Source.Model, 'Time'):
                attribute = 'zeit'
                text = ev.Source.Model.Time

                if self.mb.programm == 'LibreOffice':
                    text = self.mb.class_Sidebar.date_time_struct_nach_long_wandeln(ev.Source.Model.Time,attribute)
            else:
                text = ev.Source.Model.Text
                attribute = 'datum'
                text = self.formatiere_datum(text)
            
            self.mb.dict_sb_content['ordinal'][ordinal]['Tags_time'][attribute] = text
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_time')

        except:
            log(inspect.stack,tb())
        
    def disposing(self,ev):pass
    
    def formatiere_datum(self,datum):
        if self.mb.debug: log(inspect.stack)
        
        gesplittet = datum.split('.')
        
        if len(gesplittet) != 3:
            return None
        
        format_datum = 'de'
        
        if format_datum == 'de':
            tag = gesplittet[0]
            monat = gesplittet[1]
            jahr = gesplittet[2]
        
        else:
            tag = gesplittet[1]
            monat = gesplittet[0]
            jahr = gesplittet[2]
        
        
        
        if 0 in (len(jahr),len(tag),len(monat)):
            return None
        if len(tag)>2 or int(tag) < 1 or int(tag) > 31:
            return None
        if len(monat)>2 or int(monat) < 1 or int(monat) > 12:
            return None
        
        if len(tag) == 1:
            tag = '0'+tag
        if len(monat) == 1:
            monat = '0'+monat
        
        return '%s.%s.%s'%(tag, monat,jahr)
       
class Tags_Remove_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
    
    def disposing(self,ev):
        return False
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            tag,ordinal,tag_eintrag,aktion = ev.ActionCommand.split('_XYX_')
            dict_sb_content = self.mb.dict_sb_content
            
            if aktion == 'hinzufuegen':
                self.hinzufuegen(tag,ordinal,tag_eintrag,dict_sb_content)
            elif aktion == 'loeschen':
                self.loeschen(tag,ordinal,tag_eintrag,dict_sb_content)
            
        except:
            log(inspect.stack,tb())

        
    def pruefe_vorkommen_in_anderen_eintraegen(self,tag,tag_eintrag):  
        if self.mb.debug: log(inspect.stack)
        
        for ordinal in self.mb.dict_sb_content['ordinal']:
            if tag_eintrag in self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general']:
                return True
        return False
    
    def loesche_vorkommen_in_allen_eintraegen(self,tag,tag_eintrag): 
        if self.mb.debug: log(inspect.stack)
        
        tags_kat = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
        
        try: 
            for ordinal in self.mb.dict_sb_content['ordinal']:
                    
                for kat in tags_kat:
                    if tag_eintrag in self.mb.dict_sb_content['ordinal'][ordinal][kat]:
                        self.mb.dict_sb_content['ordinal'][ordinal][kat].remove(tag_eintrag)
            
            for kat in self.mb.dict_sb_content['tags']:
                if tag_eintrag in self.mb.dict_sb_content['tags'][kat]:
                    self.mb.dict_sb_content['tags'][kat].remove(tag_eintrag)    
        except:
            log(inspect.stack,tb())


    def loesche_vorkommen_in_selektierter_datei(self,tag,tag_eintrag,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        tags_kat = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
       
        for tag_kat in tags_kat: 
            if tag_eintrag in self.mb.dict_sb_content['ordinal'][ordinal][tag_kat]:
                self.mb.dict_sb_content['ordinal'][ordinal][tag_kat].remove(tag_eintrag)
                
        vorkommen = self.pruefe_vorkommen_in_anderen_eintraegen(tag,tag_eintrag)
        if not vorkommen:
            for tag_kat in tags_kat:
                if tag_eintrag in self.mb.dict_sb_content['tags'][tag_kat]:
                    self.mb.dict_sb_content['tags'][tag_kat].remove(tag_eintrag)
                    
    
    def loeschen(self,tag,ordinal,tag_eintrag,dict_sb_content): 
        if self.mb.debug: log(inspect.stack)
        
        if tag != 'Tags_general':
            dict_sb_content['ordinal'][ordinal][tag].remove(tag_eintrag)
            dict_sb_content['ordinal'][ordinal]['Tags_general'].remove(tag_eintrag)
            
            vorkommen = self.pruefe_vorkommen_in_anderen_eintraegen(tag,tag_eintrag)
            if not vorkommen:
                dict_sb_content['tags']['Tags_general'].remove(tag_eintrag)
                dict_sb_content['tags'][tag].remove(tag_eintrag)
            
            self.mb.class_Sidebar.erzeuge_sb_layout(tag,'sidebar')
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','sidebar')
            
        else:
            if dict_sb_content['einstellungen']['tags_general_loescht_im_ges_dok'] == 1:
                self.loesche_vorkommen_in_allen_eintraegen(tag,tag_eintrag)
            else:
                self.loesche_vorkommen_in_selektierter_datei(tag,tag_eintrag,ordinal)
            for sichtbare in self.mb.dict_sb['sichtbare']:
                self.mb.class_Sidebar.erzeuge_sb_layout(sichtbare,'sidebar')
        
     
     
    def hinzufuegen(self,tag,ordinal,tag_eintrag,dict_sb_content): 
        if self.mb.debug: log(inspect.stack)

        if tag != 'Tags_general':
            dict_sb_content['ordinal'][ordinal][tag].append(tag_eintrag)
            if tag_eintrag not in dict_sb_content['ordinal'][ordinal]['Tags_general']:
                dict_sb_content['ordinal'][ordinal]['Tags_general'].append(tag_eintrag)
            
            vorkommen = self.pruefe_vorkommen_in_anderen_eintraegen(tag,tag_eintrag)
            if not vorkommen:
                dict_sb_content['tags']['Tags_general'].append(tag_eintrag)
                dict_sb_content['tags'][tag].append(tag_eintrag)
        
        elif tag == 'Tags_general':
            
            if tag_eintrag not in dict_sb_content['tags']['Tags_general']:
                dict_sb_content['tags']['Tags_general'].append(tag_eintrag)
                
            dict_sb_content['ordinal'][ordinal]['Tags_general'].append(tag_eintrag)
        
        self.mb.class_Sidebar.erzeuge_sb_layout(tag,'sidebar')
        self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','sidebar')

    
