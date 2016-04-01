# -*- coding: utf-8 -*-

import unohelper


class Fenster():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.container = None
        #self.geoeffnete_fenster = {}
        
        
    def erzeuge_fensterinhalt(self,controls,pos_y=None):
        
        try:
            # Controls und Models erzeugen
            if pos_y == None:
                pos_y = 0
            ctrls = {}
            
            pos_y_max = [0]
            
            for ctrl in controls:
                if isinstance(ctrl,int):
                    pos_y += ctrl
                elif 'Y=' in ctrl:
                    pos_y_max.append(pos_y)
                    pos_y = int(ctrl.split('Y=')[1])
                else:
                    name,unoCtrl,calc,X,Y,width,height,prop_names,prop_values,extras = ctrl
                    
                    if unoCtrl == 'CheckBox':
                        p = list(prop_names)
                        p.append('VerticalAlign')
                        prop_names = tuple(p)
                        p2 = list(prop_values)
                        p2.append(1)
                        prop_values = tuple(p2)
                    
                    control,model = self.mb.createControl(self.mb.ctx,unoCtrl,X,pos_y+Y,width,height,prop_names,prop_values)
    
                    for k,v in extras.items():
                        if k not in ('SelectedItems'):
                            
                            attr = getattr(control, k)  
                            
                            # funktioniert bei PyUno Objekt nicht:
                            #if hasattr(attr, '__call__'):
                            
                            if 'callable' in str(attr):
                                if isinstance(v, tuple):
                                    attr(*v) 
                                else:
                                    attr(v)
                            else:
                                setattr(control, k, v)
                                
                    if 'SelectedItems' in extras:
                        control.Model.SelectedItems = extras['SelectedItems']

                    ctrls.update({name:control})    
            
            pos_y_max.append(pos_y)
            max_hoehe = max(pos_y_max)
            
            return ctrls,max_hoehe + 20
        except:
            log(inspect.stack,tb())
            
            
    def berechne_tabs(self,controls,tabs,abstand_links):  
        ''' Formatierungen
            
            "-" :     (tab0-tab2)   Breite des Buttons
            "-max" :  (tab-max)     Breite des Buttons bis zum Fensterende
            "x" :     (tab0x)       Buttonbreite wird bei Berechnung des nächsten Tabs ignoriert
            "- -E":   (tab0-tab2-E) Buttonbreite von tab0 bis tab2-Ende
            "+":      (tab0+40)     setzt den Button auf Tab plus Zahlwert
            
            calc = 1    errechnet Preferred Size des Buttons
            
            in tabs (dict) werden die festen Breiten und der Mindestabstand gespeichert.
            
        '''
        if self.mb.debug: log(inspect.stack)  
        
        try:
            tabs_breiten = {t:0 for t in tabs}
            tabs_breite_ctrl_max = {t:0 for t in tabs}
            #mindest_breite = {t:0 if v[0] == None else v[0] for t,v in tabs.items()}
            #abstand = {t:0 if v[1] == None else v[1] for t,v in tabs.items()}
            letzter_tab = sorted(tabs)[-1]
    
            alle_tabs = {}
            breiten = {}
            infos = {}
            
            for c in controls:
                if not isinstance(c, int) and 'Y=' not in c:
                    if c[3] in alle_tabs:
                        alle_tabs[c[3]].append(c)
                    else:
                        alle_tabs[c[3]] = [c]
            

            def get_breite(ct):
                 
                try:
                    name,unoCtrl,calc,tabX,Y,width,height,prop_names,prop_values,extras = ct

                    ctrl = self.mb.createUnoService("com.sun.star.awt.UnoControl%s" % unoCtrl)
                    ctrl_model = self.mb.createUnoService("com.sun.star.awt.UnoControl%sModel" % unoCtrl)
                    ctrl_model.setPropertyValues(prop_names,prop_values)
                    ctrl.setModel(ctrl_model) 
                     
                    if calc:
                        if 'addItems' in extras:
                            ctrl.addItems(extras['addItems'][0],0)
             
                    prefSize = ctrl.getPreferredSize()
                    width = prefSize.Width
                 
                    return width
                except: 
                    log(inspect.stack,tb())  
                    return 0
                
                        
            def get_info(txt):
                eintrag = {
                           'taba' : None,
                           'tabb' : None,
                           'kalk' : True,
                           'rechts' : 0,
                           'tabb_end' : False
                           }
            
                parts = txt.split('-')
                
                for p in range(len(parts)):
                    
                    if p == 0:
                        
                        eintrag['kalk'] = 'x' not in parts[p]
                        
                        term = re.sub('[tabx]', '', parts[p])
                        
                        if '+' in term:
                            a,r = term.split('+')
                            eintrag['taba'],eintrag['rechts'] = int(a),int(r)
                        else:
                            eintrag['taba'] = int(term)
                            
                    elif p == 1:
                        
                        if parts[p] == 'max':
                            eintrag['tabb'] = letzter_tab
                            eintrag['tabb_end'] = True
                        else:
                            eintrag['tabb'] = int(re.sub('[tab]', '', parts[p]))
                            
                    else:
                        eintrag['tabb_end'] = True
                
                return eintrag
            
            
            def berechne_breiten(c,breiten,tab_breite=None):
                for c in ctrls:
                    if c[2]: # calc
                        width = get_breite(c)
                    else:
                        width = c[5] # width
                    
                    breiten[c[0]] = width
                    
                    if tab_breite != None:
                        if width > tab_breite:
                            tab_breite = width
                
                return tab_breite
                        
                    
            # Tabbreiten berechnen 
            for tab,ctrls in alle_tabs.items():
                
                info = get_info(tab)
                infos[tab] = info
                
                if not info['kalk']:
                    berechne_breiten(c, breiten)
                    continue
                 
                nr = info['taba']
                
                
                if tabs[nr][0] != None:
                    tab_breite = tabs[nr][0]
                else:
                    tab_breite = berechne_breiten(c, breiten, tab_breite=0)
                    
                if nr not in tabs_breiten:  
                    tabs_breiten[nr] = tab_breite + tabs[nr][1]
                    tabs_breite_ctrl_max[nr] = tab_breite
                else:
                    if tab_breite + tabs[nr][1] > tabs_breiten[nr]:
                        tabs_breiten[nr] = tab_breite + tabs[nr][1]
                    if tab_breite > tabs_breite_ctrl_max[nr]:
                        tabs_breite_ctrl_max[nr] = tab_breite
            
                        
            # Tabbreiten Summen 
            tabs_breiten_list = [tabs_breiten[v] for v in sorted(tabs_breiten)]
            tabs_breiten_summen = [sum(tabs_breiten_list[:i+1]) for i in range(len(tabs_breiten_list))]
            
             
            neue_tabs = {k+1:tabs_breiten_summen[k]  + abstand_links for k,v in tabs.items()}
            neue_tabs[0] = abstand_links
            
            
            # Tabbreiten setzen 
            controls2 = []
            max_breiten = []
             
            for c in controls:
                if not isinstance(c, int) and 'Y=' not in c:

                    name,unoCtrl,calc,tab_name,Y,width,height,prop_names,prop_values,extras = c
                    c2 = list(c)
                    
                    info = infos[tab_name]
                    
                    # c2[3] : x    c2[5] : width
                    c2[3] = neue_tabs[info['taba']] + info['rechts'] 
                    
                    if info['tabb'] != None:
                        c2[5] = neue_tabs[ info['tabb'] ] - c2[3] 
                        
                        if info['tabb_end']:
                            c2[5] += tabs_breite_ctrl_max[ info['tabb'] ]
                            
                    else:
                        if name in breiten:
                            c2[5] = breiten[name]
                    
                    controls2.append(c2)
                     
                    max_breiten.append(c2[3]+c2[5])
                    
                else:
                    controls2.append(c)
             
            return controls2,neue_tabs,max(max_breiten) + abstand_links
        except:
            log(inspect.stack,tb())
            

    def container_anpassen(self,container,max_breite=None,max_hoehe=None,fenster=None):
        if self.mb.debug: log(inspect.stack)
        
                  
        posSize = container.PosSize
        
        # Breite 
        if max_breite:
            if max_breite > posSize.Width:
                container.setPosSize(0,0,max_breite,0,4)
                if fenster:
                    unterschied = max_breite- posSize.Width
                    fenster.setPosSize(0,0,fenster.PosSize.Width + unterschied,0,4)
            
        # Hoehe
        if max_hoehe:            
            if max_hoehe > posSize.Height:
                container.setPosSize(0,0,0,max_hoehe,8)
                if fenster:
                    unterschied = max_hoehe - posSize.Height
                    fenster.setPosSize(0,0,0,fenster.PosSize.Height + unterschied,8)
  

    def erzeuge_Dialog_Container(self,posSize,Flags=1+32+64+128,parent=None):
        if self.mb.debug: log(inspect.stack)
        
        ctx = self.mb.ctx
        smgr = ctx.ServiceManager
        
        X,Y,Width,Height = posSize

        if parent == None:
            parent = self.mb.topWindow 
            
        if X == None:
            X = parent.PosSize.Width/2-Width/2
        if Y == None:
            Y = parent.PosSize.Height/2-Height/2
        
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
        oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)
    
        # Create Uno Struct
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
        oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
        # global oWindow
        oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
        oWindowDesc.WindowServiceName = ""
        oWindowDesc.Parent = parent
        oWindowDesc.ParentIndex = -1
        oWindowDesc.WindowAttributes = Flags # Flags fuer com.sun.star.awt.WindowAttribute
    
        oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
        oReturnValue, oRect = oXIdlClass.createObject(None)
        oRect.X = X
        oRect.Y = Y
        oRect.Width = Width 
        oRect.Height = Height 
        
        oWindowDesc.Bounds = oRect
    
        # create window
        oWindow = toolkit.createWindow(oWindowDesc)
         
        # create frame for window
        oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame",ctx)
        oFrame.initialize(oWindow)
        oFrame.setCreator(self.mb.desktop)
        oFrame.activate()

        # create new control container
        cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
        cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
        #cont_model.BackgroundColor = KONST.FARBE_ORGANON_FENSTER  # 9225984
        
        cont.setModel(cont_model)
        cont.createPeer(toolkit, oWindow)

        oFrame.setComponent(cont, None)
        
        # PosSize muss erneut gesetzt werden, um die Anzeige zu erneuern,
        # sonst bleibt ein Teil des Fensters schwarz
        oWindow.setPosSize(0,0,Width,Height,12)

        # um das fenster bei sehr vielen Controls schneller zu schliessen,
        # wird es vom Listener auf invisible(True) gesetzt
        dispose_listener = Window_Dispose_Listener(oWindow,self.mb)
        cont.addEventListener(dispose_listener)
        

        if self.mb.settings_orga['organon_farben']['design_organon_fenster']:
            self.mb.class_Organon_Design.set_app_style(cont,self.mb.settings_orga)
        
        return oWindow,cont


    def oeffne_dokument_in_neuem_fenster(self,URL):
        if self.mb.debug: log(inspect.stack)
        
        new_doc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,())
            
        contWin = new_doc.CurrentController.Frame.ContainerWindow               
        contWin.setPosSize(0,0,870,900,12)
        
        lmgr = new_doc.CurrentController.Frame.LayoutManager
        for elem in lmgr.Elements:
        
            if lmgr.isElementVisible(elem.ResourceURL):
                lmgr.hideElement(elem.ResourceURL)
                
        lmgr.HideCurrentUI = True  
        
        viewSettings = new_doc.CurrentController.ViewSettings
        #viewSettings.ZoomType = 3
        viewSettings.ZoomValue = 100
        viewSettings.ShowRulers = False
        
        return new_doc
        
        
    def erzeuge_treeview_mit_checkbox(self,tab_name='ORGANON',listener_innen=None,pos=None,auswaehlen=None,parent=None):
        if self.mb.debug: log(inspect.stack)
        
        control_innen, model = self.mb.createControl(self.mb.ctx,"Container",20,0,400,100,(),() )
        
        if auswaehlen:
            listener_innen = Auswahl_CheckBox_Listener(self.mb)
        
        x,y,ctrls = self.erzeuge_treeview_mit_checkbox_eintraege(tab_name,
                                                                 control_innen,
                                                                 listener=listener_innen,
                                                                 auswaehlen=auswaehlen)
        control_innen.setPosSize(0, 0,x,y + 20,12)
        
        if not pos:
            X,Y = 0,0
        else:
            X,Y = pos
        
        x += 40
        y += 10
        
        
        erzeuge_scrollbar = False
        if y > 800:
            y = 800
            erzeuge_scrollbar = True
            
        posSize = X,Y,x,y
        
        fenster,fenster_cont = self.erzeuge_Dialog_Container(posSize,parent=parent)

        fenster_cont.addControl('Container_innen', control_innen)
        
        if auswaehlen:
            listener_innen.ctrls = ctrls
        
        if erzeuge_scrollbar:
            self.erzeuge_Scrollbar(fenster_cont,(0,0,0,y),control_innen)
            self.mb.class_Mausrad.registriere_Maus_Focus_Listener(fenster_cont)
        
        return y,fenster,fenster_cont,control_innen,ctrls
    
    
    def erzeuge_treeview_mit_checkbox_eintraege(self,tab_name,control_innen,listener=None,auswaehlen=None):
        if self.mb.debug: log(inspect.stack)
        try:
            sett = self.mb.settings_exp
            
            tree = self.mb.props[tab_name].xml_tree
            root = tree.getroot()
            
            baum = []
            self.mb.class_XML.get_tree_info(root,baum)
            
            y = 10
            x = 10
                
            # Titel AUSWAHL
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x,y ,300,20,(),() )  
            control.Text = LANG.AUSWAHL_TIT
            model.FontWeight = 150.0
            control_innen.addControl('Titel', control)
            
            y += 30
            
            # Untereintraege auswaehlen
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40,y ,300,20,(),() )  
            control.Text = LANG.ORDNER_CLICK
            model.FontWeight = 150.0
            control_innen.addControl('ausw', control)
            x_pref = control.getPreferredSize().Width + x + 40
            
            control, model = self.mb.createControl(self.mb.ctx,"CheckBox",x+20,y ,20,20,(),() )  
            control.State = sett['auswahl']
            control.ActionCommand = 'untereintraege_auswaehlen'

            if listener:
                control.addActionListener(listener)
                control.ActionCommand = 'untereintraege_auswaehlen'
            control_innen.addControl('Titel', control)
    
            y += 30
            
            ctrls = {}
            
            
            for eintrag in baum:
    
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
                
                if art == 'waste':
                    break
                
                control1, model1 = self.mb.createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,400,20,(),() )  
                control1.Text = name
                control_innen.addControl('Titel', control1)
                pref = control1.getPreferredSize().Width
                
    
                if x_pref < x + 40+20*int(lvl) + pref:
                    x_pref = x + 40+20*int(lvl) + pref
                
                control2, model2 = self.mb.createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
                model2.Border = False
                
                if art in ('dir','prj'):
                    model2.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
                else:
                    model2.ImageURL = 'private:graphicrepository/res/sx03150.png' 
                control_innen.addControl('Titel', control2)   
                  
                    
                control3, model3 = self.mb.createControl(self.mb.ctx,"CheckBox",x+20*int(lvl),y ,20,20,(),() )  
                control_innen.addControl(ordinal, control3)
                if listener:
                    control3.addActionListener(listener)
                    control3.ActionCommand = ordinal+'xxx'+name
                    if auswaehlen:
                        if ordinal in sett['ausgewaehlte']:
                            model3.State = sett['ausgewaehlte'][ordinal]
                
                ctrls.update({ordinal:[control1,control2,control3]})
                
                y += 20 
                
            return x_pref,y,ctrls
        except:
            log(inspect.stack,tb())  
            
            
    def erzeuge_Scrollbar(self,fenster_cont,PosSize,control_innen,called_from_hf=False):
        if self.mb.debug: log(inspect.stack)
           
        PosX,PosY,Width,Height = PosSize
        Width = 20
        
        control, model = self.mb.createControl(self.mb.ctx,"ScrollBar",PosX,PosY,Width,Height,(),() )  
        model.Orientation = 1
        model.LiveScroll = True        
        model.ScrollValueMax = control_innen.PosSize.Height/4 
                
        control.LineIncrement = fenster_cont.PosSize.Height/Height*50
        control.BlockIncrement = 200
        control.Maximum =  control_innen.PosSize.Height  
        control.VisibleSize = Height      

        listener = ScrollBar_Listener(self.mb,control_innen)
        listener.fenster_cont = control_innen
        control.addAdjustmentListener(listener) 
        
        if called_from_hf:
            listener.called_from_hf = True
        
        fenster_cont.addControl('ScrollBar',control) 
        
        return control 
    
    def erzeuge_Scrollbar2(self,win = None):
        if self.mb.debug: log(inspect.stack)
        
        if win == None:
            win = self.mb.prj_tab
            
        nav_cont_aussen = win.getControl('Hauptfeld_aussen')
        control_innen = nav_cont_aussen.getControl('Hauptfeld')
        
        MBHoehe = 22
        tableiste_hoehe = self.mb.tabsX.tableiste_hoehe 

        Height = self.mb.win.Size.Height - MBHoehe - tableiste_hoehe
        PosSize = 0,MBHoehe,0,Height
        
        self.erzeuge_Scrollbar(win,PosSize,control_innen,called_from_hf=True)
              
    
from com.sun.star.lang import XEventListener
class Window_Dispose_Listener(unohelper.Base,XEventListener):
    '''
    Closing the dialog window holding 50+ controls might
    freeze Writer. This listener closes the window
    explicitly
    
    '''
    def __init__(self,fenster,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.fenster = fenster
    
    def disposing(self,ev):

        if sys.platform == 'win32':
            # Beschleunigt das Dispose ungemein !
            self.fenster.setVisible(False)
            return
        else:
            self.fenster.dispose()

    
    
from com.sun.star.awt import XActionListener   
class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.ctrls = None
    
    def disposing(self,ev):
        return False

    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        sett = self.mb.settings_exp

        if ev.ActionCommand == 'untereintraege_auswaehlen':
            sett['auswahl'] = self.toggle(sett['auswahl'])
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
        else:
            ordinal,titel = ev.ActionCommand.split('xxx')
            state = ev.Source.Model.State
            sett['ausgewaehlte'].update({ordinal:state})
            
            props = self.mb.props[T.AB]
            try:
                if sett['auswahl']:
                    if ordinal in props.dict_ordner:
                        
                        tree = props.xml_tree
                        root = tree.getroot()
                        C_XML = self.mb.class_XML
                        ord_xml = root.find('.//'+ordinal)
                        
                        eintraege = []
                        # selbstaufruf nur fuer den debug
                        C_XML.selbstaufruf = False
                        C_XML.get_tree_info(ord_xml,eintraege)
                        
                        ordinale = []
                        for eintr in eintraege:
                            ordinale.append(eintr[0])
                            

                        
                        for ordn in ordinale:
                            if ordn in self.ctrls:
                                control = self.ctrls[ordn][2]
                                control.Model.State = state
                                sett['ausgewaehlte'].update({ordn:state}) 

            except:
                if self.mb.debug: log(inspect.stack,tb())
                

    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:              
            return 1      
    
    
from com.sun.star.awt import XAdjustmentListener
class ScrollBar_Listener (unohelper.Base,XAdjustmentListener):
    
    def __init__(self,mb,fenster_cont):   
        
        if mb.debug: log(inspect.stack) 
        self.fenster_cont = None
        self.mb = mb
        self.called_from_hf = False
        
    def adjustmentValueChanged(self,ev):
        self.fenster_cont.setPosSize(0, -ev.value.Value,0,0,2)
        if self.called_from_hf:
            self.mb.class_Zeilen_Listener.schalte_sichtbarkeit_hf_ctrls()
            
    def disposing(self,ev):
        return False
    
    def schalte_sichtbarkeit_hf_ctrls(self):
        
        try:
            props = self.mb.props[T.AB]
            co = props.Hauptfeld.PosSize.Y
            tv = self.mb.win.PosSize.Height
            tableiste = self.mb.tabsX.tableiste.Size
    
            untergrenze = -co - 20
            obergrenze = -co + tv - tableiste.Height - 20
            
            Ys = props.dict_zeilen_posY
    
            for y in Ys:
                if untergrenze < y < obergrenze:
                    props.dict_posY_ctrl[y].setVisible(True)
                else:
                    props.dict_posY_ctrl[y].setVisible(False)
        except:
            log(inspect.stack,tb())    
                