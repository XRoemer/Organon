# -*- coding: utf-8 -*-

import unohelper
import pickle
from pickle import HIGHEST_PROTOCOL
from pickle import load as pickle_load
from pickle import dump as pickle_dump

class Sidebar():
    
    def __init__(self,mb,pdk):
        self.mb = mb        
        self.mb.dict_sb['erzeuge_sb_layout'] = self.erzeuge_sb_layout
        self.mb.dict_sb['test'] = self.test
        
        
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
        
        self.sb_panels1 = {'Synopsis':self.mb.lang.SYNOPSIS,
            'Notes':self.mb.lang.NOTIZEN,
            'Images':self.mb.lang.BILDER,
            'Tags_general':self.mb.lang.ALLGEMEIN,
            'Tags_characters':self.mb.lang.CHARAKTERE,
            'Tags_locations':self.mb.lang.ORTE,
            'Tags_objects':self.mb.lang.OBJEKTE,
            'Tags_time':self.mb.lang.ZEIT,
            'Tags_user1':self.mb.lang.BENUTZER1,
            'Tags_user2':self.mb.lang.BENUTZER2,
            'Tags_user3':self.mb.lang.BENUTZER3}
        
        
        
        self.sb_panels2 = {self.mb.lang.SYNOPSIS:'Synopsis',
            self.mb.lang.NOTIZEN:'Notes',
            self.mb.lang.BILDER:'Images',
            self.mb.lang.ALLGEMEIN:'Tags_general',
            self.mb.lang.CHARAKTERE:'Tags_characters',
            self.mb.lang.ORTE:'Tags_locations',
            self.mb.lang.OBJEKTE:'Tags_objects',
            self.mb.lang.ZEIT:'Tags_time',
            self.mb.lang.BENUTZER1:'Tags_user1',
            self.mb.lang.BENUTZER2:'Tags_user2',
            self.mb.lang.BENUTZER3:'Tags_user3'}
        
        self.sb_tags = ('Tags_general',
                        'Tags_characters',
                        'Tags_locations',
                        'Tags_objects',
                        'Tags_user1',
                        'Tags_user2',
                        'Tags_user3')

       
        

        global pd
        pd = pdk
    
    
    
    def lade_sidebar(self):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar lade_sidebar')

        if 'empty_project' in self.mb.dict_sb['sichtbare']:
            self.mb.dict_sb['sichtbare'].remove('empty_project')
        
        self.mb.dict_sb['sichtbare'] = self.mb.dict_sb_content['sichtbare']
              
      
    def passe_sb_an(self,textbereich):
        if self.mb.debug: self.mb.timer_start = self.mb.time.clock()
        if self.mb.debug: print(self.mb.debug_time(),'sidebar passe_sb_an')
        ordinal = textbereich.Context.Model.Text
        
        for panel in self.mb.dict_sb['sichtbare']:
            self.erzeuge_sb_layout(panel,'passe_sb_an')

    
    def lege_dict_sb_content_an(self):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar lege_dict_sb_content_an')
        
        try:
            sb_panels = self.sb_panels_tup
            
            tags = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
            
            
            dict_sb_content = {}
            dict_sb_content.update({'ordinal':{}})
            dict_sb_content.update({'tags':{}})
            # 'sichtbare' definiert hier die Ansicht bei einem neuen Projekt
            dict_sb_content.update({'sichtbare':['Synopsis','Notes','Tags_general','Tags_characters']})
            
            # Anlegen der Verzeichnisstruktur 'ordinal'
            for ordinal in list(self.mb.dict_bereiche['ordinal']):
                dict_sb_content['ordinal'].update({ordinal:{}})
                for panel in sb_panels:
                    if panel in tags:
                        dict_sb_content['ordinal'][ordinal].update({panel:[]})
                        if panel not in dict_sb_content['tags']:
                            dict_sb_content['tags'].update({panel:[]})
                    else:
                        dict_sb_content['ordinal'][ordinal].update({panel:''})
                        
            
            
            dict_sb_content.update({'einstellungen':{}})   
            dict_sb_content['einstellungen'].update({'hoehe_Synopsis':200})  
            dict_sb_content['einstellungen'].update({'breite_Synopsis':284}) 
            dict_sb_content['einstellungen'].update({'hoehe_Notes':200})   
            dict_sb_content['einstellungen'].update({'breite_Notes':284})
            dict_sb_content['einstellungen'].update({'tags_general_loescht_im_ges_dok':0})        
                    
            self.mb.dict_sb_content = dict_sb_content
        except:
            tb()
    
    def lege_dict_sb_content_ordinal_an(self,ordinal):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar lege_dict_sb_content_ordinal_an')
        
        sb_panels = self.sb_panels_tup
        tags = 'Tags_general','Tags_characters','Tags_locations','Tags_objects','Tags_user1','Tags_user2','Tags_user3'
            
        # Anlegen der Verzeichnisstruktur 'ordinal'
        
        self.mb.dict_sb_content['ordinal'].update({ordinal:{}})
        for panel in sb_panels:
            if panel in tags:
                self.mb.dict_sb_content['ordinal'][ordinal].update({panel:[]})
            else:
                self.mb.dict_sb_content['ordinal'][ordinal].update({panel:''})
    
    def loesche_dict_sb_content_eintrag(self,ordinal):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar loesche_dict_sb_content_eintrag')
        
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
        if self.mb.debug: print(self.mb.debug_time(),'sidebar speicher_sidebar_dict')
        
        pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
        with open(pfad, 'wb') as f:
            pickle_dump(self.mb.dict_sb_content, f,2)


    def lade_sidebar_dict(self):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar lade_sidebar_dict')
        
        pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
        
        dict_exists, backup_exists = False, False

        if os.path.exists(pfad):
            dict_exists = True
        if not os.path.exists(pfad+'.Backup'):
            backup_exists = True
        
        
        try:  
            if dict_exists:          
                with open(pfad, 'rb') as f:
                    self.mb.dict_sb_content =  pickle_load(f)
                self.ueberpruefe_dict_sb_content(backup_exists)
                
            elif backup_exists:
                self.lege_dict_sb_content_an()
                self.lade_Backup()
        except:
            self.lege_dict_sb_content_an()
            if backup_exists:
                self.lade_Backup()
            
        if not dict_exists:
            self.speicher_sidebar_dict()

        self.erzeuge_dict_sb_content_Backup()
        
        
    def ueberpruefe_dict_sb_content(self,backup_exists):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar ueberpruefe_dict_sb_content')
        fehlende = []

        for ordinal in list(self.mb.dict_bereiche['ordinal']):
            if ordinal not in self.mb.dict_sb_content['ordinal']:
                fehlende.append(ordinal)
                
        if backup_exists:
            self.lade_Backup(fehlende)
        else:
            for f in fehlende:
                self.lege_dict_sb_content_ordinal_an(f)
                  
    
    def lade_Backup(self,fehlende = 'all'): 
        if self.mb.debug: print(self.mb.debug_time(),'sidebar lade_Backup') 
        
        pfad_Backup = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl.Backup')
        
        if fehlende == 'all':
            fehlende = list(self.mb.dict_bereiche['ordinal'])
        
        try:
            with open(pfad, 'rb') as f:
                backup = pickle_load(pfad_Backup)
        except:
            for f in fehlende:
                self.lege_dict_sb_content_ordinal_an(f)
            return
               
        for f in fehlende:
            if f in backup['ordinal']:
                self.dict_sb_content['ordinal'].update(backup['ordinal'][f])
                fehlende.remove(f)
                
        for f in fehlende:
            self.lege_dict_sb_content_ordinal_an(f)
               
        
     

         
       
    def erzeuge_dict_sb_content_Backup(self):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar erzeuge_dict_sb_content_Backup')
        pfad = os.path.join(self.mb.pfade['files'],'sidebar_content.pkl')
        pfad_Backup = pfad + '.Backup'
        from shutil import copy2
        copy2(pfad, pfad_Backup)
        
    def erzeuge_sb_layout(self,xUIElement_name,rufer):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar erzeuge_sb_layout',rufer)
                
        if xUIElement_name == 'empty_project':
            return
        
        # Wenn die Sidebar noch nicht geoeffnet wurde,
        # sind noch keine Panels vorhanden
        if xUIElement_name not in self.mb.dict_sb['controls']:
            return
                
        try:
            ctx = self.mb.ctx
            
            xUIElement = self.mb.dict_sb['controls'][xUIElement_name][0]
            sb = self.mb.dict_sb['controls'][xUIElement_name][1]
            panelWin = xUIElement.Window
            
            #panelWin.Model.BackgroundColor = KONST.FARBE_NAVIGATIONSFELD
            
            # alte Eintraege im einzelnen Panel vorher loeschen
            if rufer != 'factory':
                for conts in panelWin.Controls:
                    conts.dispose()
            
            
            
            ######################################
            #                TAGS                #
            ######################################            
            
            ordinal = self.mb.selektierte_zeile.AccessibleName
            
            if xUIElement_name in self.sb_tags:
                
                remove_or_add_button_listener = Tags_Remove_Button_Listener(self.mb)

                y = 10
                height = 20
                
                prop_names = ('HelpText',)
                prop_values = (self.mb.lang.ENTER_NEW_TAG,)
                control, model = self.mb.createControl(ctx, "Edit", 170, y,100, height, prop_names, prop_values)
                panelWin.addControl('Button', control) 
                focus_listener = Tags_Focus_Listener(self.mb,xUIElement_name)
                control.addFocusListener(focus_listener)
                
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
                        prop_values = ('',self.mb.lang.TAG_HINZUFUEGEN)
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
                            helptext = self.mb.lang.TAGS_IN_AKT_DAT_LOESCHEN
                        else:
                            helptext = self.mb.lang.TAGS_IM_GES_DOK_LOESCHEN
                    else:
                        helptext = self.mb.lang.TAG_LOESCHEN
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
                height = self.mb.dict_sb_content['einstellungen']['hoehe_Synopsis']
                width = self.mb.dict_sb_content['einstellungen']['breite_Synopsis']
                 
                text = self.mb.dict_sb_content['ordinal'][ordinal]['Synopsis']
                 
                prop_names = ('MultiLine','Text')
                prop_values = (True,text)
                control, model = self.mb.createControl(self.mb.ctx, "Edit", 10, pos_y, width, height, prop_names, prop_values)  
                panelWin.addControl('Synopsis', control)
                
                listener = Text_Change_Listener_Synopsis(self.mb)
                model.addPropertyChangeListener('Text',listener)
                
                height += 20
                
            elif xUIElement_name == 'Notes':
                pos_y = 10
                height = self.mb.dict_sb_content['einstellungen']['hoehe_Notes']
                width = self.mb.dict_sb_content['einstellungen']['breite_Notes']
                 
                text = self.mb.dict_sb_content['ordinal'][ordinal]['Notes']
                 
                prop_names = ('MultiLine','Text')
                prop_values = (True,text)
                control, model = self.mb.createControl(self.mb.ctx, "Edit", 10, pos_y, width, height, prop_names, prop_values)  
                panelWin.addControl('Notes', control)
                
                listener = Text_Change_Listener_Notizen(self.mb)
                model.addPropertyChangeListener('Text',listener)
                
                height += 20
                
                
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
                    tb()
                
                
                
                height += 20
                
            elif xUIElement_name == 'Tags_time':
                height = 20
                
            xUIElement.height = height
            sb.requestLayout()
            
            # zum Speichern und Wiederherstellen der sichtbaren Panels
            self.mb.dict_sb_content['sichtbare'] = self.mb.dict_sb['sichtbare']
            
        except:
            tb()
            #pd()
    

    def berechne_bildgroesse(self,model,hoehe):
        try:
            HOEHE = model.Graphic.Size.Height
            BREITE = model.Graphic.Size.Width

            quotient = float(BREITE)/float(HOEHE)
            BREITE = int(hoehe*quotient)
            return BREITE
            
        except:
            tb()
        #pd()

    def dict_sb_zuruecksetzen(self):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar dict_sb_zuruecksetzen')
        
        self.mb.dict_sb['sichtbare']  = ['empty_project'] 
        self.mb.dict_sb['controls'] = {}
        self.mb.dict_sb['erzeuge_Layout'] = None   
    
    
    def toggle_sicht_sidebar(self):
        frame = self.mb.current_Contr.Frame
        dispatcher = self.mb.createUnoService("com.sun.star.frame.DispatchHelper")
        dispatch = dispatcher.executeDispatch(frame, ".uno:Sidebar" , "", 0, ())
    
    
    def test(self,cmd):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar test')
        
        loc_x = self.mb.dict_sb['controls'][cmd][0].Window.Peer.AccessibleContext.LocationOnScreen.X
        loc_y = self.mb.dict_sb['controls'][cmd][0].Window.Peer.AccessibleContext.LocationOnScreen.Y
        
        if cmd in ('Synopsis','Notes'):
            self.optionsfenster_synopsis_notes(cmd,loc_x,loc_y)
        if cmd =='Tags_general':
            self.optionsfenster_tags_general(loc_x,loc_y)
        if cmd =='Images':
            self.optionsfenster_images(loc_x,loc_y)
        
            
    def optionsfenster_images(self,loc_x,loc_y):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar optionsfenster_images')
        
        win,cont = self.mb.erzeuge_Dialog_Container((loc_x - 20,loc_y,350,110))
 
        listener = Options_Tags_General_And_Images_Listener(self.mb,win)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.IN_PROJEKTORDNER_IMPORTIEREN,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 340, 20, prop_names, prop_values)  
        cont.addControl('Titel', control)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.BILD_EINFUEGEN,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 30, 120, 30, prop_names, prop_values)  
        cont.addControl('Titel', control)
        control.setActionCommand('bild_einfuegen')
        control.addActionListener(listener)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.BILD_LOESCHEN,)
        control, model = self.mb.createControl(self.mb.ctx, "Button", 10, 70, 120, 30, prop_names, prop_values)  
        cont.addControl('Titel2', control)
        control.setActionCommand('bild_loeschen')
        control.addActionListener(listener)
        
            
    def optionsfenster_synopsis_notes(self,cmd,loc_x,loc_y):  
        if self.mb.debug: print(self.mb.debug_time(),'sidebar optionsfenster_synopsis_notes')     
        
        win,cont = self.mb.erzeuge_Dialog_Container((loc_x - 20,loc_y,150,80))
        
        breite = self.mb.dict_sb_content['einstellungen']['breite_' + cmd]
        hoehe = self.mb.dict_sb_content['einstellungen']['hoehe_' + cmd]
        
        
        
        prop_names = ('Label',)
        prop_values = ('Size Text Box '+ cmd,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 10, 120, 20, prop_names, prop_values)  
        cont.addControl('Titel', control)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.BREITE,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 30, 45, 20, prop_names, prop_values)  
        cont.addControl('NumericField', control)
        
        
        prop_names = ('Value','StrictFormat','DecimalAccuracy','ValueMin')
        prop_values = (breite,True,0,50)
        control, model = self.mb.createControl(self.mb.ctx, "NumericField", 50, 30, 30, 16, prop_names, prop_values)  
        cont.addControl('NumericField', control)
        options_syn_note_text_listener = Options_Syn_Note_Text_Listener(self.mb)
        control.addTextListener(options_syn_note_text_listener)
        
        prop_names = ('Label',)
        prop_values = (self.mb.lang.HOEHE,)
        control, model = self.mb.createControl(self.mb.ctx, "FixedText", 10, 50, 45, 20, prop_names, prop_values)  
        cont.addControl('NumericField', control)
        
        
        prop_names = ('Value','StrictFormat','DecimalAccuracy','ValueMin')
        prop_values = (hoehe,True,0,50)
        control, model = self.mb.createControl(self.mb.ctx, "NumericField", 50, 50, 30, 16, prop_names, prop_values)  
        cont.addControl('NumericField', control)
        control.addTextListener(options_syn_note_text_listener)
    
        
    def optionsfenster_tags_general(self,loc_x,loc_y):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar optionsfenster_tags_general')
        
        win,cont = self.mb.erzeuge_Dialog_Container((loc_x - 20,loc_y,350,80))
        try:
            
            state = self.mb.dict_sb_content['einstellungen']['tags_general_loescht_im_ges_dok']
            
            if state == 1:
                state2 = 0
            else:
                state2 = 1
            
            listener = Options_Tags_General_And_Images_Listener(self.mb)
            
            prop_names = ('Label',)
            prop_values = (self.mb.lang.EINSTELLUNGEN_TAGS_GENERAL,)
            control, model = self.mb.createControl(self.mb.ctx, "FixedText", 28, 10, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            
            prop_names = ('Label','State')
            prop_values = (self.mb.lang.TAGS_IM_GES_DOK_LOESCHEN,state)
            control, model = self.mb.createControl(self.mb.ctx, "RadioButton", 10, 30, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            control.setActionCommand('1')
            control.addActionListener(listener)
            
            prop_names = ('Label','State')
            prop_values = (self.mb.lang.TAGS_IN_AKT_DAT_LOESCHEN,state2)
            control, model = self.mb.createControl(self.mb.ctx, "RadioButton", 10, 50, 340, 20, prop_names, prop_values)  
            cont.addControl('Titel', control)
            control.setActionCommand('0')
            control.addActionListener(listener)

        except:
            tb()

from com.sun.star.awt import XTextListener
class Options_Syn_Note_Text_Listener(unohelper.Base, XTextListener):
    def __init__(self,mb):
        self.mb = mb
        
    def textChanged(self,ev):
        
        if self.mb.lang.BREITE in ev.Source.AccessibleContext.AccessibleName:
            cmd1 = 'breite_'
        else:
            cmd1 = 'hoehe_'
        
        titel = ev.Source.Context.getControl('Titel').Text.split(' ')[-1]
        
        
        value = int(ev.Source.Text)
        
        self.mb.dict_sb_content['einstellungen'][cmd1 + titel] = value
        self.mb.class_Sidebar.erzeuge_sb_layout(titel,'options')
            
    
    
from com.sun.star.beans import XPropertyChangeListener
class Text_Change_Listener_Synopsis(unohelper.Base, XPropertyChangeListener):
    def __init__(self,mb):
        self.mb = mb
        
    def propertyChange(self,ev):
        ordinal = self.mb.selektierte_zeile.AccessibleName
        self.mb.dict_sb_content['ordinal'][ordinal]['Synopsis'] = ev.NewValue
          
    
from com.sun.star.beans import XPropertyChangeListener
class Text_Change_Listener_Notizen(unohelper.Base, XPropertyChangeListener):
    def __init__(self,mb):
        self.mb = mb
        
    def propertyChange(self,ev):
        ordinal = self.mb.selektierte_zeile.AccessibleName
        self.mb.dict_sb_content['ordinal'][ordinal]['Notes'] = ev.NewValue
    

from com.sun.star.awt import XActionListener
class Options_Tags_General_And_Images_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,win = None):
        self.mb = mb
        self.win = win
        
    def disposing(self,ev):return False
    
    def actionPerformed(self,ev):

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
        
        ordinal = self.mb.selektierte_zeile.AccessibleName
        
        old_image_path = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
        
        self.mb.dict_sb_content['ordinal'][ordinal]['Images'] = path
        self.mb.class_Sidebar.erzeuge_sb_layout('Images','Options_Tags_General_And_Images_Listener')
        
        if old_image_path != '' and old_image_path != path:
            self.bild_loeschen(old_image_path)

            
    def bild_loeschen_a(self):
        ordinal = self.mb.selektierte_zeile.AccessibleName
        old_image_path = self.mb.dict_sb_content['ordinal'][ordinal]['Images']
        self.mb.dict_sb_content['ordinal'][ordinal]['Images'] = ''
        self.bild_loeschen(old_image_path)
        self.mb.class_Sidebar.erzeuge_sb_layout('Images','Options_Tags_General_And_Images_Listener')
        self.win.dispose()

    def bild_loeschen(self,old_image_path):
        try:
            vorhanden = False
            for ordinal in self.mb.dict_sb_content['ordinal']:
                if old_image_path == self.mb.dict_sb_content['ordinal'][ordinal]['Images']:
                    vorhanden = True
                    
            if not vorhanden:
                os.remove(uno.fileUrlToSystemPath(old_image_path))
        except:
            tb()
            
        


from com.sun.star.awt import XFocusListener
class Tags_Focus_Listener(unohelper.Base, XFocusListener):
    def __init__(self,mb,tag):
        self.mb = mb
        self.tag = tag
    
    def focusGained(self,ev):
        return False
        
    def focusLost(self,ev):
        # Hinzufuegen neuer Tags
        ordinal = self.mb.selektierte_zeile.AccessibleName
        new_tag = ev.Source.Model.Text
        if new_tag != '':
            self.mb.dict_sb_content['ordinal'][ordinal][self.tag].append(new_tag)
            
            if new_tag not in self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general']:
                self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general'].append(new_tag)
            
            if new_tag not in self.mb.dict_sb_content['tags'][self.tag]:
                self.mb.dict_sb_content['tags'][self.tag].append(new_tag)
            
            if new_tag not in self.mb.dict_sb_content['tags']['Tags_general']:
                self.mb.dict_sb_content['tags']['Tags_general'].append(new_tag)
                 
            ev.Source.Model.Text = ''
            
            self.mb.class_Sidebar.erzeuge_sb_layout(self.tag,'sidebar')
            self.mb.class_Sidebar.erzeuge_sb_layout('Tags_general','sidebar')


    def disposing(self,ev):pass

       
class Tags_Remove_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
    
    def disposing(self,ev):
        return False
        
    def actionPerformed(self,ev):
        try:
            tag,ordinal,tag_eintrag,aktion = ev.ActionCommand.split('_XYX_')
            dict_sb_content = self.mb.dict_sb_content
            
            if aktion == 'hinzufuegen':
                self.hinzufuegen(tag,ordinal,tag_eintrag,dict_sb_content)
            elif aktion == 'loeschen':
                self.loeschen(tag,ordinal,tag_eintrag,dict_sb_content)
            
        except:
            tb()

        
    def pruefe_vorkommen_in_anderen_eintraegen(self,tag,tag_eintrag):  
        if self.mb.debug: print(self.mb.debug_time(),'sidebar pruefe_vorkommen_in_anderen_eintraegen')
        
        for ordinal in self.mb.dict_sb_content['ordinal']:
            if tag_eintrag in self.mb.dict_sb_content['ordinal'][ordinal]['Tags_general']:
                return True
        return False
    
    def loesche_vorkommen_in_allen_eintraegen(self,tag,tag_eintrag): 
        if self.mb.debug: print(self.mb.debug_time(),'sidebar loesche_vorkommen_in_allen_eintraegen')
        
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
            tb()
            #pd()

    def loesche_vorkommen_in_selektierter_datei(self,tag,tag_eintrag,ordinal):
        if self.mb.debug: print(self.mb.debug_time(),'sidebar loesche_vorkommen_in_selektierter_datei')
        
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
        if self.mb.debug: print(self.mb.debug_time(),'sidebar loeschen')
        
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
        if self.mb.debug: print(self.mb.debug_time(),'sidebar hinzufuegen')

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
            
            
