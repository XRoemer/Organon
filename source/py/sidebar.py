# -*- coding: utf-8 -*-

import unohelper
import pickle
from pickle import load as pickle_load
from pickle import dump as pickle_dump

class Sidebar(): 
    '''
    Die Methode setze_sidebar_design findet sich im Modul menu_start,
    da sie dort schon gebraucht wird und es keine direkte Verbindung
    der zwei Module gibt. setze_sidebar_design wird via 
    dict_sb['setze_sidebar_design'] uebergeben.
    '''
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb     
        self.mb.dict_sb['erzeuge_sb_layout'] = self.erzeuge_sb_layout

        self.hoehen = [] 
        self.offen = {}
        
        self.url_expand = None
        self.url_collapse = None


    def erzeuge_sb_layout(self):
        if self.mb.dict_sb['sb_closed']:return
        if self.mb.debug: log(inspect.stack)
        
        dict_sb = self.mb.dict_sb
        
        if not dict_sb['design_gesetzt']:
            # faerbt nur noch die Schrift in der Menuleiste 
            if self.mb.settings_orga['organon_farben']['design_office']:
                dict_sb['setze_sidebar_design']()
                
                      
        # hoehen zuruecksetzen
        self.hoehen = []

        try:
            ctx = self.mb.ctx
            
            xUIElement = dict_sb['controls']['organon_sidebar'][0]
            sb = dict_sb['controls']['organon_sidebar'][1]
            panelWin = xUIElement.Window
            
            self.url_expand = xUIElement.Theme.Image_Expand
            self.url_collapse = xUIElement.Theme.Image_Collapse

            # alte Eintraege im einzelnen Panel vorher loeschen
            for conts in panelWin.Controls:
                conts.dispose()

            orga_sb,seitenleiste = self.get_seitenleiste()
            
            try:
                breite_sidebar = seitenleiste.PosSize.Width
            except:
                breite_sidebar = 400
            
            
            def container_erzeugen(panelWin):
                
                dict_felder = {}
                dict_container = {}
                
                pos_y = 0
                
                listener = Tags_Collapse_Button_Listener(self.mb,self)
                
                for panel_nr in self.mb.tags['abfolge']:
                    
                    if panel_nr not in self.mb.tags['sichtbare']:
                        continue
                    
                    panel_name,panel_typ = self.mb.tags['nr_name'][panel_nr]
                    
                    height1 = 30
                    width1 = breite_sidebar
                    
                    container0, model0 = self.mb.createControl(self.mb.ctx, "Container", 0, pos_y, width1, height1, (), ())  
                    
                    if self.mb.programm == 'LibreOffice':
                        line, model = self.mb.createControl(self.mb.ctx, "FixedLine", 0, 0, width1 - 50, 1, (), ()) 
                        container0.addControl('Line',line)
                    
                    pos_y2 = 0
                    height1 = 30
                    container1, model1 = self.mb.createControl(self.mb.ctx, "Container", 0, 0, width1, height1, (), ()) 
                    container0.addControl('Titel',container1)
                    
                    if self.mb.programm == 'OpenOffice':
                        line, model = self.mb.createControl(self.mb.ctx, "FixedLine", 0, -4, width1 - 50, 10, (), ()) 
                        container0.addControl('Line',line)
                    
                    if panel_nr not in self.offen:
                        self.offen[panel_nr] = 1
                    
                    url = self.url_collapse if self.offen[panel_nr] else self.url_expand
                        
                    cont_icon, model = self.mb.createControl(self.mb.ctx, "ImageControl", 8,9 , 10, 10,
                                                             ('ImageURL','Border'), (url,0))  
                    
                    # falls kein icon in der graphic repository vorhanden ist
                    if cont_icon.PreferredSize.Height == 0:
                        model.ImageURL = KONST.IMG_PLUS if self.offen[panel_nr] else KONST.IMG_MINUS
                        
                    cont_icon.addMouseListener(listener)
                    container1.addControl('label',cont_icon)
                    
                    cont_label, model = self.mb.createControl(self.mb.ctx, "FixedText", 30, 0 , width1, height1,
                                                              ('Label','VerticalAlign','FontWeight'), (panel_name,1,150))  
                    container1.addControl('label',cont_label)
                    
                    pos_y2 += 30
                    
                    container2, model2 = self.mb.createControl(self.mb.ctx, "Container", 0, pos_y2 , width1, 0, (), ())  
                    container0.addControl('Feld',container2)
                    
                    pos_y += 200
                    
                    dict_felder.update({panel_nr:container2})
                    dict_container.update({panel_nr:container0})

                    panelWin.addControl('Container' + str(panel_nr), container0)
                        
                  
                return dict_container,dict_felder
                
             
        
            self.dict_container,self.dict_felder = container_erzeugen(panelWin)
               
                    
            ######################################
            #                TAGS                #
            ######################################            
            
            ordinal = self.mb.props[T.AB].selektierte_zeile
            ctrls = self.mb.dict_sb['controls']
            tags = self.mb.tags
            
            height = 0 
            
            for panel_nr in self.mb.tags['abfolge']:
                panel_name,panel_typ = self.mb.tags['nr_name'][panel_nr]

                if panel_nr not in self.mb.tags['sichtbare']:
                    continue

                tag_control = self.dict_felder[panel_nr]
            
            
                if panel_typ == 'tag':
                    
                    remove_or_add_button_listener = Tags_Remove_Button_Listener(self.mb)
     
                    y = 10
                    height = 20
                     
                    prop_names = ('HelpText','MultiLine')
                    prop_values = (LANG.ENTER_NEW_TAG,True)
                    control, model = self.mb.createControl(ctx, "Edit", 170, y,100, height, prop_names, prop_values)
                    tag_control.addControl('Button', control) 
     
                    key_listener = Tags_Key_Listener(self.mb,panel_nr)
                    control.addKeyListener(key_listener)
     
                    y_all_tags = y
                    y_all_tags += 30
                     
                    alle_tags = self.mb.tags['sammlung'][panel_nr]
                    eintraege = self.mb.tags['ordinale'][ordinal][panel_nr]
                     
                    SEP = '_XYX_'
                     
                    # Leiste mit allen Tags
                    for tag_eintrag in alle_tags:
                        if tag_eintrag not in eintraege:
                             
                            height = 18
                            
                            # Tageintrag
                            prop_names = ('Label','Align')
                            prop_values = (tag_eintrag,2)
                            control, model = self.mb.createControl(ctx, "FixedText", 170, y_all_tags ,100, height, prop_names, prop_values)
                            tag_control.addControl(tag_eintrag, control) 
                             
                            # Add Button
                            prop_names = ('Label','HelpText')
                            prop_values = ('',LANG.TAG_HINZUFUEGEN)
                            control, model = self.mb.createControl(ctx, "Button", 280, y_all_tags ,14, 14, prop_names, prop_values)
                            tag_control.addControl('Remove_'+tag_eintrag, control) 
                            control.setActionCommand(str(panel_nr) + SEP + ordinal + SEP + tag_eintrag + SEP + 'hinzufuegen')
                            control.addActionListener(remove_or_add_button_listener)
                            
                            y_all_tags += height 
                            
                            
                    # Tags
                     
                    for tag_eintrag in eintraege:
                         
                        height = 18
                        # Tageintrag
                        prop_names = ('Label',)
                        prop_values = (tag_eintrag,)
                        control, model = self.mb.createControl(ctx, "FixedText", 10, y,100, height, prop_names, prop_values)
                        tag_control.addControl(tag_eintrag, control) 
                    
                        # Remove Button
                        prop_names = ('Label',)
                        prop_values = ('X',)
                        control, model = self.mb.createControl(ctx, "Button", 115, y,14, 14, prop_names, prop_values)
                        tag_control.addControl('Remove_'+tag_eintrag, control) 
                        control.setActionCommand(str(panel_nr) + SEP + ordinal + SEP + tag_eintrag + SEP + 'loeschen')
                        control.addActionListener(remove_or_add_button_listener)
                         
                        y += height + 3
                     
                  
                    if y < y_all_tags:
                        y = y_all_tags
                     
                    if len(eintraege) == 0:
                        y_all_tags += 21
                     
                    height = y + 10 
                    
                    
                    
                ######################################
                # Text, Images, Time, Date           #
                ######################################     
                 
                if panel_typ == 'txt':
                      
                    pos_y = 10
                    width = 282 
                       
                    text = self.mb.tags['ordinale'][ordinal][panel_nr]
                       
                    prop_names = ('MultiLine','Text','MaxTextLen')
                    prop_values = (True,text,5000)
                    control, model = self.mb.createControl(self.mb.ctx, "Edit", 10, pos_y, width, height, prop_names, prop_values)  
                    tag_control.addControl('Synopsis', control)
                      
                    mS = control.getMinimumSize()
                    hoehe = mS.Height + 44
      
                    control.setPosSize(0,0,0,hoehe,8)
                      
                    listener = Text_Change_Listener(self.mb,panel_nr)
                    model.addPropertyChangeListener('Text',listener)
                      
                    height = hoehe + 20
        
                elif panel_typ == 'img':
                    pos_y = 10
                    height = 200
                    breite = 284 
      
                    prop_names = ()
                    prop_values = ()
                    control, model = self.mb.createControl(self.mb.ctx, "ImageControl", 10, pos_y, breite, height, prop_names, prop_values)  
                    tag_control.addControl('Image', control)
                    
                    try:
                        if self.mb.tags['ordinale'][ordinal][panel_nr] != '':
                            model.ImageURL = self.mb.tags['ordinale'][ordinal][panel_nr]
                            breite = self.berechne_bildgroesse(model,height)
                            control.setPosSize(0,0,breite,0,4)
                    except:
                        # Bild ist nicht geladen worden
                        self.mb.tags['ordinale'][ordinal][panel_nr] = ''
                        log(inspect.stack,tb())
                        
                    listener = Images_Listener(self.mb)
                    
                    height += 20

                    prop_names = ('Label',)
                    prop_values = ('+',)
                    control, model = self.mb.createControl(self.mb.ctx, "Button",breite + 15, pos_y , 16, 16, prop_names, prop_values)  
                    tag_control.addControl('Titel', control)
                    control.setActionCommand('bild_einfuegen_XYZ_'+str(panel_nr))
                    control.addActionListener(listener)
                    
                    prop_names = ('Label',)
                    prop_values = ('-',)
                    control, model = self.mb.createControl(self.mb.ctx, "Button", breite +15 , pos_y + 25 , 16, 16, prop_names, prop_values)  
                    tag_control.addControl('Titel2', control)
                    control.setActionCommand('bild_loeschen_XYZ_'+str(panel_nr))
                    control.addActionListener(listener) 
                                        
                    height += 35
                    
                    
                     
                elif panel_typ == 'time':
                      
                    pos_y = 0
                    pos_x = 10
                    pos_x2 = 85
                    pos_x3 = 170
                    y = 20
                      
                    focus_listener = Tag_Time_Key_Listener(self.mb,panel_nr,panel_typ)
                      
                    zeit = self.mb.tags['ordinale'][ordinal][panel_nr]
   
                    prop_names = ('Label',)
                    prop_values = (LANG.ZEIT2,)
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x, pos_y, 70, y, prop_names, prop_values)  
                    tag_control.addControl('Time_label', control)
                      
                    prop_names = ('Label',)
                    prop_values = (zeit,)
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x2, pos_y, 70, y, prop_names, prop_values) 
                    tag_control.addControl('Time_field', control)
      
      
                    prop_names = ('Text',)
                    prop_values = ('',)
                    control, model = self.mb.createControl(self.mb.ctx, "Edit", pos_x3, pos_y, 60, y, prop_names, prop_values)  
                    tag_control.addControl('Time_field_input', control)
                      
                    control.addKeyListener(focus_listener)
                      
                      
                    pos_y += 30
                    height = 30
                    
                elif panel_typ == 'date':

                    pos_y = 0
                    pos_x = 10
                    pos_x2 = 85
                    pos_x3 = 170
                    y = 20
                    
                    focus_listener = Tag_Time_Key_Listener(self.mb,panel_nr,panel_typ)
                      
                    prop_names = ('Label',)
                    prop_values = (LANG.DATUM,)
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x, pos_y, 70, y, prop_names, prop_values)  
                    tag_control.addControl('Datum_Label', control)
                    
                    datum = self.mb.tags['ordinale'][ordinal][panel_nr]
                    if datum:
                        datum_txt = self.mb.class_Tags.formatiere_datumdict_nach_text(datum)
                    else:
                        datum_txt = None    
                    
                    prop_names = ('Label',)
                    prop_values = (datum_txt,)
                    control, model = self.mb.createControl(self.mb.ctx, "FixedText", pos_x2, pos_y, 70, y, prop_names, prop_values)  
                    tag_control.addControl('Datum_txt', control)
        
                    prop_names = ('Text',)
                    prop_values = ('',)
                    control, model = self.mb.createControl(self.mb.ctx, "Edit", pos_x3, pos_y, 60, y, prop_names, prop_values)  
                    tag_control.addControl('Datum_edit', control)
      
                    control.addKeyListener(focus_listener)                 
                      
                    height = 30
                
                self.hoehen.append([panel_nr,height])

            
            y = 0
            y2 = 0
            for panel_nr,hoehe in self.hoehen:

                tag_ctrl = self.dict_container[panel_nr]
                feld_ctrl = self.dict_felder[panel_nr]
                
                
                if self.offen[panel_nr]:
                    tag_ctrl.setPosSize(0,y,0,hoehe+30,10)
                    feld_ctrl.setPosSize(0,0,0,hoehe+30,8)
                    y += hoehe +30
                else:
                    tag_ctrl.setPosSize(0,y,0,30,10)
                    feld_ctrl.setPosSize(0,0,0,0,8)
                    y += 30
                    
                y2 += hoehe +30

            xUIElement.height = y

            sb.requestLayout()
            
        except:
            log(inspect.stack,tb())  
                        
        
    def berechne_bildgroesse(self,model,hoehe):
        if self.mb.debug: log(inspect.stack)
        
        HOEHE = model.Graphic.Size.Height
        BREITE = model.Graphic.Size.Width

        quotient = float(BREITE)/float(HOEHE)
        BREITE = int(hoehe*quotient)
        return BREITE
            
    
    def bild_in_projektordner_kopieren(self,filepath):
        if self.mb.debug: log(inspect.stack)
        
        basename = os.path.basename(filepath)
        path_image_folder = self.mb.pfade['images']
        complete_path = os.path.join(path_image_folder,basename)
        path = uno.systemPathToFileUrl(complete_path)

        if not os.path.exists(complete_path):
            from shutil import copy2
            copy2(filepath, complete_path)
            
        return path
    
    
    def bild_einfuegen(self, panel_nr, ordinal=None, filepath='', erzeuge_layout=True):
        if self.mb.debug: log(inspect.stack)
        
        if filepath == '':
            filepath,ok = self.mb.class_Funktionen.filepicker2()
            if not ok:
                return False

        path = self.bild_in_projektordner_kopieren(filepath)
        
        if ordinal == None:
            ordinal = self.mb.props[T.AB].selektierte_zeile
        
        old_image_path = self.mb.tags['ordinale'][ordinal][panel_nr]
        
        self.mb.tags['ordinale'][ordinal][panel_nr] = path
        
        if erzeuge_layout:
            self.mb.class_Sidebar.erzeuge_sb_layout()
        
        if old_image_path != '' and old_image_path != path:
            self.bild_loeschen(old_image_path,panel_nr)
        
        return path  


    def bild_loeschen_a(self,panel_nr, ordinal=None, erzeuge_layout= True):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if ordinal == None:
                ordinal = self.mb.props[T.AB].selektierte_zeile
                
            old_image_path = self.mb.tags['ordinale'][ordinal][panel_nr]
            
            self.mb.tags['ordinale'][ordinal][panel_nr] = ''
            
            if erzeuge_layout:
                self.mb.class_Sidebar.erzeuge_sb_layout()
                
            self.bild_loeschen(old_image_path,panel_nr)
        except:
            log(inspect.stack,tb())


    def bild_loeschen(self,old_image_path,panel_nr):
        if self.mb.debug: log(inspect.stack)
        
        try:
            vorhanden = False
            for ordinal in self.mb.tags['ordinale']:
                if old_image_path == self.mb.tags['ordinale'][ordinal][panel_nr]:
                    vorhanden = True

            if not vorhanden:
                os.remove(uno.fileUrlToSystemPath(old_image_path))
        except:
            log(inspect.stack,tb())            
        
    
    def get_seitenleiste(self):
        if self.mb.debug: log(inspect.stack)
               
        desk = self.mb.desktop
        contr = desk.CurrentComponent.CurrentController
        wins = contr.ComponentWindow.Windows
        
        childs = []

        for w in wins:
            if not w.isVisible():continue
            
            if w.AccessibleContext.AccessibleChildCount == 0:
                continue
            else:
                child = w.AccessibleContext.getAccessibleChild(0)
                if 'Organon: dockable window' == child.AccessibleContext.AccessibleName:
                    continue
                else:
                    childs.append(child)
                    
        orga_sb = None
        ch = None
        try:
            for c in childs:
                try:
                    for w in c.Windows:
                        try:
                            for w2 in w.Windows:
                                if w2.AccessibleContext.AccessibleDescription == 'Organon':
                                    orga_sb = w2
                                    ch = c
                        except:
                            pass
                except:
                    pass
        except:
            log(inspect_stack,tb())

        return orga_sb,ch
    
   
from com.sun.star.beans import XPropertyChangeListener
class Text_Change_Listener(unohelper.Base, XPropertyChangeListener):
    def __init__(self,mb,panel_nr):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.panel_nr = panel_nr
        
    def propertyChange(self,ev):

        ordinal = self.mb.props[T.AB].selektierte_zeile
        self.mb.tags['ordinale'][ordinal][self.panel_nr] = ev.NewValue


from com.sun.star.awt import XActionListener
class Images_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
    def disposing(self,ev):return False
    
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        cmd,p = ev.ActionCommand.split('_XYZ_')
        panel_nr = int(p)

        # optionsfenster_images
        if cmd == 'bild_einfuegen':
            self.mb.class_Sidebar.bild_einfuegen(panel_nr)
        elif cmd == 'bild_loeschen':
            self.mb.class_Sidebar.bild_loeschen_a(panel_nr)

        


from com.sun.star.awt import XFocusListener,XKeyListener
class Tags_Key_Listener(unohelper.Base, XKeyListener):
    def __init__(self,mb,panel_nr):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.panel_nr = panel_nr
    
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
        tags = self.mb.tags
        
        from itertools import chain
        alle_tags_in_anderen_panels = list(chain.from_iterable(
                                            [v for i,v in tags['sammlung'].items() if i != self.panel_nr ]
                                            ))

        
        if new_tag != '' and new_tag not in alle_tags_in_anderen_panels:
            
            if new_tag not in tags['sammlung'][self.panel_nr]:
                tags['sammlung'][self.panel_nr].append(new_tag)
                
                if new_tag not in tags['ordinale'][ordinal][self.panel_nr]:
                    tags['ordinale'][ordinal][self.panel_nr].append(new_tag)
                
                self.mb.class_Sidebar.erzeuge_sb_layout()
                
            ev.Source.Model.Text = ''
        else:
            ev.Source.Model.Text = ''


    def disposing(self,ev):pass


    
class Tag_Time_Key_Listener(unohelper.Base, XKeyListener):
    def __init__(self,mb,panel_nr,panel_typ):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.panel_nr = panel_nr
        self.panel_typ = panel_typ

    def keyPressed(self,ev):
        return False
        
    def keyReleased(self,ev):
        # 1280 = Return
        if ev.KeyCode != 1280:
            return False
        # Nur bei Tasteneingabe Return loggen
        if self.mb.debug: log(inspect.stack)
        
        # Return wird auch im Dokument ausgefuehrt,
        # daher hier ein Undo
        self.mb.doc.UndoManager.undo()
        
        try:
            ordinal = self.mb.props[T.AB].selektierte_zeile
            text = ev.Source.Model.Text
            
            if self.panel_typ == 'time':
                text = self.mb.class_Tags.formatiere_zeit(text)
            else:
                text,odict = self.mb.class_Tags.formatiere_datum(text)
            
            self.mb.tags['ordinale'][ordinal][self.panel_nr] = odict
            self.mb.class_Sidebar.erzeuge_sb_layout()

        except:
            log(inspect.stack,tb())
            
    def disposing(self,ev):pass
    
       

class Tags_Remove_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
    
    def disposing(self,ev):
        return False
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            panel_nr,ordinal,tag_eintrag,aktion = ev.ActionCommand.split('_XYX_')
            panel_nr = int(panel_nr)
            
            if aktion == 'hinzufuegen':
                self.hinzufuegen(panel_nr,ordinal,tag_eintrag)
            elif aktion == 'loeschen':
                self.loeschen(panel_nr,ordinal,tag_eintrag)
            
        except:
            log(inspect.stack,tb())

        
    def pruefe_vorkommen_in_anderen_eintraegen(self,panel_nr,tag_eintrag):  
        if self.mb.debug: log(inspect.stack)
        
        self.ueberpruefe_dict()
        
        dic = self.mb.tags['ordinale']
        for ordinal in dic:
            if tag_eintrag in dic[ordinal][panel_nr]:
                return True

        return False
                
    
    def loeschen(self,panel_nr,ordinal,tag_eintrag): 
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags
        
        tags['ordinale'][ordinal][panel_nr].remove(tag_eintrag)
        vorkommen = self.pruefe_vorkommen_in_anderen_eintraegen(panel_nr,tag_eintrag)
        
        if not vorkommen:
            tags['sammlung'][panel_nr].remove(tag_eintrag)
        
        self.mb.class_Sidebar.erzeuge_sb_layout()
            
     
    def hinzufuegen(self,panel_nr,ordinal,tag_eintrag): 
        if self.mb.debug: log(inspect.stack)
        
        tags = self.mb.tags

        if tag_eintrag not in tags['ordinale'][ordinal][panel_nr]:
            tags['ordinale'][ordinal][panel_nr].append(tag_eintrag)
        
            if tag_eintrag not in tags['sammlung'][panel_nr]:
                tags['sammlung'][panel_nr].append(tag_eintrag)

            self.mb.class_Sidebar.erzeuge_sb_layout()

    
    def ueberpruefe_dict(self):
        '''
        Fehler im dict wurden wahrscheinlich durch falsches Handling
        von utf8 Characters erzeugt. Zur Sicherheit ist diese Methode
        eingebaut, die richtige Eintraege im dict sicherstellt.
        '''
        dic = self.mb.tags['ordinale']
        zu_loeschende = [d for d in dic if 'nr' not in d]
        
        for d in zu_loeschende:
            del(dic[d])


from com.sun.star.awt import XMouseListener
class Tags_Collapse_Button_Listener(unohelper.Base, XMouseListener):
    def __init__(self,mb,sb):
        self.mb = mb
        self.sb = sb
        
    def disposing(self,ev):
        return False
        
    def mouseReleased(self,ev):
        return False
    def mouseExited(self,ev):
        return False
    def mouseEntered(self,ev):
        return False
    
    def mousePressed(self,ev):
        if self.mb.debug: log(inspect.stack)
        source = ev.Source
        txt = source.Context.Controls[1].Text
        panel_nr = self.mb.tags['name_nr'][txt]
        self.collaps_expand_panels(panel_nr,txt,source)


    def collaps_expand_panels(self,panel_nr,name,source):   
        if self.mb.debug: log(inspect.stack)
        try:
            
            cont = self.sb.dict_container[panel_nr]
            feld = self.sb.dict_felder[panel_nr]
            hoehen = self.sb.hoehen
            
            tag,hoehe = [h for h in hoehen if h[0] == panel_nr][0]
            index = hoehen.index([tag,hoehe])
            
            h_panel_titel = 30

            if self.sb.offen[panel_nr]:
                self.sb.offen[panel_nr] = 0
                cont.setPosSize(0,0,0,h_panel_titel,8)
                feld.setPosSize(0,0,0,0,8)
                source.Model.ImageURL = self.sb.url_expand
                fak = -1

            else:
                self.sb.offen[panel_nr] = 1
                cont.setPosSize(0,0,0,hoehe + h_panel_titel,8)
                feld.setPosSize(0,0,0,hoehe,8)
                source.Model.ImageURL = self.sb.url_collapse
                fak = 1
                
                
            for i in range(index+1,len(hoehen)):
                t = hoehen[i][0]
                cont2 = self.sb.dict_container[t]

                y = cont2.PosSize.Y
                
                cont2.setPosSize(0,y + hoehe * fak,0,0,2)
            
            sb_tup = self.mb.dict_sb['controls']['organon_sidebar']
            xuiElement = sb_tup[0]
            
            gesamthoehe = [h[1] + 30 if self.sb.offen[h[0]] else h_panel_titel for h in hoehen ]
            
            xuiElement.height = sum(gesamthoehe)
            sb_tup[1].requestLayout()
            
        except:
            log(inspect.stack,tb())
        


        
        
        
        
        
        
        
        
        
    
    