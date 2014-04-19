# -*- coding: utf-8 -*-

import unohelper


class Export():
    
    def __init__(self,mb,pdk):
        self.mb = mb
        self.trenner_fenster = None
        self.auswahl_fenster = None
        global pd
        pd = pdk
        
        global lang
        lang = self.mb.lang

    def export(self):
        self.erzeuge_exportfenster()

    def erzeuge_exportfenster(self): 
        

        posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
        X = posSize_main.X +20
        Y = posSize_main.Y +20
        Width = 220
        Height = 420
        
        posSize = X,Y,Width,Height
        fenster,fenster_cont = erzeuge_Dialog_Container(self.mb.smgr,self.mb.ctx,posSize)
        fenster_cont.Model.Text = lang.EXPORT
        listenerDis = AB_Fenster_Dispose_Listener(self.mb,self)
        fenster_cont.addEventListener(listenerDis)
        
        y = 10
        
        # Titel
        controlE, modelE = createControl(self.mb.ctx,"FixedText",20,y ,50,20,(),() )  
        controlE.Text = lang.EXPORT
        modelE.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlE)

        # CheckBoxen
        y += 40
        
        buttons = []
        labels = (lang.ALLES,lang.SICHTBARE,lang.AUSWAHL)
        
        for label in labels:
            control, model = createControl(self.mb.ctx,"CheckBox",20,y,120,22,(),() )  
            model.Label = label
            
            if label == lang.ALLES:
                control.State = self.mb.settings_exp['alles']
            elif label == lang.SICHTBARE:
                control.State = self.mb.settings_exp['sichtbar']
            elif label == lang.AUSWAHL:
                control.State = self.mb.settings_exp['eigene_ausw'] 
                
            fenster_cont.addControl(label, control)
            buttons.append(control)
            y += 20
        
        # Auswahl   
        controlA, modelA = createControl(self.mb.ctx,"Button",130,y - 25 ,70,22,(),() )  ###
        controlA.Label = lang.AUSWAHL
        if self.mb.settings_exp['alles'] or self.mb.settings_exp['sichtbar']:
            controlA.Enable = False
        fenster_cont.addControl('Auswahl', controlA)
        
        y += 20
        
        # Trenner 
        controlT, modelT = createControl(self.mb.ctx,"FixedLine",20,y-10 ,180,40,(),() )  
        fenster_cont.addControl('Trenner', controlT)
        
        y += 40
        
        # text
        controlex, modelex = createControl(self.mb.ctx,"FixedText",20,y ,150,20,(),() )  
        controlex.Text = lang.EXPORTIEREN_ALS
        modelex.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlex)
        
        y += 30
        
        # Exportoptionen
        labels2 = (lang.EIN_DOKUMENT,lang.TRENNER,lang.EINZ_DATEIEN,lang.ORDNERSTRUKTUR)
        buttons2= []
        for label in labels2:
            
            x = 0
            if label in (lang.ORDNERSTRUKTUR,lang.TRENNER):
                x = 20
                
            control1, model1 = createControl(self.mb.ctx,"CheckBox",20 + x,y,110,22,(),() )  
            model1.Label = label
            
            if label == lang.EIN_DOKUMENT:
                control1.State = self.mb.settings_exp['einz_dok']
            elif label == lang.TRENNER:
                control1.State = self.mb.settings_exp['trenner']
                if self.mb.settings_exp['einz_dat']:
                    control1.Enable = False
                else:
                    control1.Enable = True
            elif label == lang.EINZ_DATEIEN:
                control1.State = self.mb.settings_exp['einz_dat']
            elif label == lang.ORDNERSTRUKTUR:
                control1.State = self.mb.settings_exp['ordner_strukt']
                if self.mb.settings_exp['einz_dok']:
                    control1.Enable = False
                else:
                    control1.Enable = True
                
            fenster_cont.addControl(label, control1)
            buttons2.append(control1)
            y += 20
            if label in (lang.ORDNERSTRUKTUR,lang.TRENNER):
                y += 20
        # um 'Ordnerstruktur beibehalten' anzuzeigen
        control1.setPosSize(0,0,200,0,4)
        
        # Auswahl   
        controlTr, modelTr = createControl(self.mb.ctx,"Button",130,y - 105 ,70,22,(),() )  ###
        controlTr.Label = lang.BEARBEITEN
        if self.mb.settings_exp['trenner'] and self.mb.settings_exp['einz_dok']:
            controlTr.Enable = True
        else:
            controlTr.Enable = False
        fenster_cont.addControl('Auswahl', controlTr)
        
        
        
        
        #y += 10
        
        # Exportformat ## fehlt ##
        controlf, modelf = createControl(self.mb.ctx,"FixedText",20,y ,150,20,(),() )  
        controlf.Text = '#' + lang.DATEITYP
        fenster_cont.addControl('Typ', controlf)
        
        # Liste der Formate
        controlL, modelL = createControl(self.mb.ctx,"ListBox",20 + 85,y ,50,20,(),() )  
        #controlL.setMultipleMode(False)
        items = ('odt','doc','pdf')
        controlL.addItems(items,0)
        modelL.Dropdown = True
        modelL.SelectedItems = 0,
        fenster_cont.addControl('Liste', controlL)
        
        controlL.Enable = False
        #pd()
        
        y += 40
        
        controlSO, modelSO = createControl(self.mb.ctx,"FixedText",20 ,y,100,22,(),() )  
        modelSO.Label = lang.SPEICHERORT
        
        fenster_cont.addControl('Dokument', controlSO)
        
        # Button
        controlSD, modelSD = createControl(self.mb.ctx,"Button",130,y-3,70,22,(),() )  ###
        controlSD.Label = lang.WAEHLEN
        controlSD.ActionCommand = 'speicherort'
        fenster_cont.addControl('Dok', controlSD)
        
        y += 20
        
        controlFO, modelFO = createControl(self.mb.ctx,"FixedText",20 ,y,500,22,(),() )  
        modelFO.HelpText = 'URL'
        #modelF.Border = True
        label = decode_utf(self.mb.settings_exp['speicherort'])
        modelFO.Label = uno.fileUrlToSystemPath(label)
        fenster_cont.addControl('Speicherort', controlFO) 
        
        listener = Speicherordner_Button_Listener(self.mb,modelFO)
        controlSD.addActionListener(listener)

        
        y += 40
        
        # Exportbutton
        controlB, modelB = createControl(self.mb.ctx,"Button",120,y,80,30,(),() )  
        controlB.Label = lang.EXPORTIEREN
        fenster_cont.addControl('Export', controlB)
        
        fenster.setPosSize(0,0,0,y + 40,8)
        
        # Listener    
        listener = Fenster_Export_Listener1(self.mb,buttons,controlA)
        for cont in buttons:
            cont.addItemListener(listener)
        
        listener2 = Fenster_Export_Listener2(self.mb,buttons2,controlTr)
        for cont in buttons2:
            cont.addItemListener(listener2)
        
        listener3 = B_Auswahl_Button_Listener(self.mb,self,fenster)
        controlA.addActionListener(listener3)
        
        listener4 = Export_Button_Listener(self.mb)
        controlB.addActionListener(listener4)
        
        listener5 = A_Trenner_Button_Listener(self.mb,self,fenster)
        controlTr.addActionListener(listener5)
        

def decode_utf(term):
    if isinstance(term, str):
        return term
    else:
        return term.decode('utf8')   

        
def erzeuge_Dialog_Container(smgr,ctx,posSize):
    
    X,Y,Width,Height = posSize

    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
    oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)

    # Create Uno Struct
    oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
    oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
    # global oWindow
    oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
    oWindowDesc.WindowServiceName = ""
    oWindowDesc.Parent = toolkit.getDesktopWindow()
    oWindowDesc.ParentIndex = -1
    oWindowDesc.WindowAttributes = 1  +32 +64 + 128 # Flags fuer com.sun.star.awt.WindowAttribute

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
    #oFrame.setCreator(self.mb.desktop)
    oFrame.activate()

    # create new control container
    cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
    cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
    cont_model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE  # 9225984
    cont.setModel(cont_model)
    # need createPeer just only the container
    cont.createPeer(toolkit, oWindow)
    cont.setPosSize(0, 0, 0, 0, 15)

    oFrame.setComponent(cont, None)
    return oWindow,cont
 

from com.sun.star.awt import XItemListener, XActionListener, XFocusListener    

class Fenster_Export_Listener1(unohelper.Base, XItemListener):
    def __init__(self,mb,buttons,but_Auswahl):
        self.mb = mb
        self.buttons = buttons
        self.but_Auswahl = but_Auswahl
        
    # XItemListener    
    def itemStateChanged(self, ev):        
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
        # alle anderen CheckBoxen auf 0 setzen
        for cont in self.buttons:
            if cont != ev.Source:
                cont.State = 0
                
        if ev.Source.Model.Label == lang.AUSWAHL:
            self.but_Auswahl.Enable = True
        else:
            self.but_Auswahl.Enable = False
        
        # settings_exp neu setzen
        self.mb.settings_exp['alles'] = 0
        self.mb.settings_exp['sichtbar'] = 0
        self.mb.settings_exp['eigene_ausw'] = 0
        
        if ev.Source.Model.Label == lang.ALLES:
            self.mb.settings_exp['alles'] = 1
        elif ev.Source.Model.Label == lang.SICHTBARE:
            self.mb.settings_exp['sichtbar'] = 1
        elif ev.Source.Model.Label == lang.AUSWAHL:
            self.mb.settings_exp['eigene_ausw'] = 1
            
        
class Fenster_Export_Listener2(unohelper.Base, XItemListener):
    
    def __init__(self,mb,buttons2,trenner):
        self.mb = mb
        self.buttons2 = buttons2
        self.trenner = trenner
        

    def itemStateChanged(self, ev):        

        if ev.Source.Model.Label in (lang.EIN_DOKUMENT,lang.EINZ_DATEIEN):
            # um sich nicht selbst abzuwaehlen
            if ev.Source.State == 0:
                ev.Source.State = 1
            #pd()
            if ev.Source.Model.Label == lang.EIN_DOKUMENT:
                self.buttons2[1].Enable = True
                self.buttons2[2].State = False
                self.buttons2[3].Enable = False
                if self.mb.settings_exp['trenner']:
                    self.trenner.Enable = True
            else:
                self.buttons2[0].State = False
                self.buttons2[1].Enable = False
                self.buttons2[3].Enable = True
                self.trenner.Enable = False
            
            self.mb.settings_exp['einz_dat'] = 0
            self.mb.settings_exp['einz_dok'] = 0
            
            if ev.Source.Model.Label == lang.EIN_DOKUMENT:
                self.mb.settings_exp['einz_dok'] = 1
            elif ev.Source.Model.Label == lang.EINZ_DATEIEN:
                self.mb.settings_exp['einz_dat'] = 1   
                
        elif ev.Source.Model.Label == lang.ORDNERSTRUKTUR:
            if self.mb.settings_exp['ordner_strukt']:
                self.mb.settings_exp['ordner_strukt'] = 0
            else:
                self.mb.settings_exp['ordner_strukt'] = 1
        
        elif ev.Source.Model.Label == lang.TRENNER: 
            if self.mb.settings_exp['trenner']:
                self.mb.settings_exp['trenner'] = 0
                self.trenner.Enable = False
            else:
                self.mb.settings_exp['trenner'] = 1
                self.trenner.Enable = True

class Export_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
        
    def actionPerformed(self,ev):
        set = self.mb.settings_exp
        
        if set['einz_dok']:
            self.exp_in_ein_dokument()
        elif set['einz_dat']:
            self.exp_in_einzel_dat()
            
            
    def exp_in_ein_dokument(self):
        #pd()
        st_ind = self.mb.current_Contr.Frame.createStatusIndicator()        
        
        set = self.mb.settings_exp
        self.mb.class_Bereiche.starte_oOO()
        
        oOO = self.mb.class_Bereiche.oOO
        cur = oOO.Text.createTextCursor()
        text = oOO.Text
        cur.gotoEnd(False)
        
        sections = self.mb.doc.TextSections.ElementNames

        if set['sichtbar']:
            sections = []
            for sec_name in self.mb.sichtbare_bereiche:
                sections.append(sec_name)
                
        if set['eigene_ausw']:
            sects = [] 
            for sec_name in sections:
                if 'OrganonSec' in sec_name:  
                    sec_ordinal = self.mb.dict_bereiche['Bereichsname-ordinal'][sec_name]
                    if sec_ordinal in set['ausgewaehlte']:
                        if set['ausgewaehlte'][sec_ordinal][1] == 1:
                            sects.append(sec_name)
            sections = sects
        
        anz_sections = len(sections)
        st_ind.start('exportiere, bitte warten',anz_sections)
        st_ind.setValue(anz_sections/2)
        zaehler = 1
            
        for sec_name in sections:
            if 'OrganonSec' in sec_name:         
                
               # st_ind.setValue(zaehler)
                zaehler += 1
                 
                cur.gotoEnd(False)
                
                sec_ordinal = self.mb.dict_bereiche['Bereichsname-ordinal'][sec_name]
                
                if sec_ordinal == self.mb.Papierkorb:
                    break  
                
                
                if set['trenner']:
                
                    if set['seitenumbruch_ord']:
                        if sec_ordinal in self.mb.dict_ordner:
                            from com.sun.star.style.BreakType import PAGE_BEFORE
                            cur.BreakType = PAGE_BEFORE
                            text.insertControlCharacter(cur, 0, True)
                            
                    if set['seitenumbruch_dat']:
                        if sec_ordinal not in self.mb.dict_ordner:
                            from com.sun.star.style.BreakType import PAGE_BEFORE
                            cur.BreakType = PAGE_BEFORE
                            text.insertControlCharacter(cur, 0, True)   
                    
                    if set['ordnertitel']:
                        if sec_ordinal in self.mb.dict_ordner:
                            
                            contr = self.mb.Hauptfeld.getControl(sec_ordinal)
                            tf = contr.getControl('textfeld')
                            titel = tf.Model.Text
                            
                            if set['format_ord']:
                                oldStyle = cur.ParaStyleName
                                cur.ParaStyleName = self.mb.settings_exp['style_ord'] 
                                
                            cur.setString(titel)
                            cur.gotoEnd(False)
                            
                            if set['format_ord']:
                                text.insertControlCharacter(cur,0,False)
                                cur.ParaStyleName = oldStyle
                            
                            cur.gotoEnd(False)
                            
                    
                    if set['dateititel']:
                        if sec_ordinal not in self.mb.dict_ordner:
                            
                            contr = self.mb.Hauptfeld.getControl(sec_ordinal)
                            tf = contr.getControl('textfeld')
                            titel = tf.Model.Text
                            
                            if set['format_dat']:
                                oldStyle = cur.ParaStyleName
                                cur.ParaStyleName = self.mb.settings_exp['style_dat'] 
                                
                            cur.setString(titel)
                            cur.gotoEnd(False)
                            
                            if set['format_dat']:
                                text.insertControlCharacter(cur,0,False)
                                cur.ParaStyleName = oldStyle
                            
                            cur.gotoEnd(False)

                                
                cur.gotoEnd(False)
                cur.gotoEndOfParagraph(False)
                sec = self.mb.doc.TextSections.getByName(sec_name)
                SFLink = sec.FileLink
                
                newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                newSection.setPropertyValue('FileLink',SFLink)
                newSection.setName(sec.Name)
                
                oOO.Text.insertTextContent(cur,newSection,False)
                oOO.Text.removeTextContent(newSection)
                
                if set['trenner']:
                
                    if set['leerzeilen_drunter']:
                        
                        anz = int(set['anz_drunter'])
                        for i in range(anz):
                            cur.gotoEnd(False)
                            text.insertControlCharacter(cur,0,False)
                            cur.gotoEnd(False)
                  
                    
                    if set['dok_einfuegen']:
                        cur.gotoEnd(False)
                        URL = set['url']
                        
                        SFLink2 = sec.FileLink
                        SFLink2.FileURL = URL
                        newSection2 = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                        newSection2.setPropertyValue('FileLink',SFLink2)
                        newSection2.setName(sec.Name)
                        
                        oOO.Text.insertTextContent(cur,newSection2,False)
                        oOO.Text.removeTextContent(newSection2)
                        
                        cur.gotoEnd(False)
       
    
        path = uno.fileUrlToSystemPath(decode_utf(self.mb.settings_exp['speicherort']))
        Path2 = os.path.join(path, self.mb.projekt_name)
        
        if os.path.exists(Path2+set['typ']):
            Path2 = self.pruefe_dateiexistenz(Path2,set['typ'])
        
        Path1 = Path2+set['typ']
        Path3 = uno.systemPathToFileUrl(Path1)   
        oOO.storeToURL(Path3,())
        self.mb.class_Bereiche.schliesse_oOO()   
         
        st_ind.end() 

    def exp_in_einzel_dat(self):
        
        set = self.mb.settings_exp
        st_ind = self.mb.current_Contr.Frame.createStatusIndicator()    
        
        
        # selektiert alle Bereiche
        sections = self.mb.doc.TextSections.ElementNames
        
        # selektiert nur die sichtbaren Bereiche
        if set['sichtbar']:
            sections = []
            for sec_name in self.mb.sichtbare_bereiche:
                sections.append(sec_name)
        
        if set['eigene_ausw']:
            sects = [] 
            for sec_name in sections:
                if 'OrganonSec' in sec_name:  
                    sec_ordinal = self.mb.dict_bereiche['Bereichsname-ordinal'][sec_name]
                    if sec_ordinal in set['ausgewaehlte']:
                        if set['ausgewaehlte'][sec_ordinal][1] == 1:
                            sects.append(sec_name)
            sections = sects
        
        if len(sections) < 1:
            self.mb.Mitteilungen.nachricht(lang.NICHTS_AUSGEWAEHLT,"infobox")

        
        # pruefen, ob speicherordner existiert; Namen aendern
        speicherordner = os.path.join(uno.fileUrlToSystemPath(self.mb.settings_exp['speicherort']), self.mb.projekt_name)
        if os.path.exists(speicherordner):
            speicherordner = self.pruefe_dateiexistenz(speicherordner)
        
        
        def berechne_tree():
             
            tree2 = copy.deepcopy(self.mb.xml_tree)
            root2 = tree2.getroot()
            
            all = root2.findall('.//')
            
            for el in all:
                parent = root2.find('.//'+el.tag+'/..')
                childs = list(parent)                
                
                childnames = []
                for c in childs:
                    childnames.append(c.attrib['Name'])
                
                if childnames.count(el.attrib['Name']) > 1:
                    anzahl = childnames.count(el.attrib['Name'])
                    
                    for an in range(anzahl):
                        el2 = root2.find('.//'+childs[an].tag)
                        el2.attrib['Name'] = el2.attrib['Name'] + '(%s)'  %an 
            
            return tree2
         
        # Pfade zum Speichern in Ordnerstruktur berechnen
        # Die Berechnung durchlaeuft mehrfach den Baum und koennte vereinfacht werden
        # Die Pfade koennten schon in berechne_tree angelegt werden
        
        if set['ordner_strukt']:

            tree = berechne_tree()
            root = tree.getroot()
            
            baum = []
            self.mb.class_XML.get_tree_info(root,baum)

            pfade = {}
            dict_baum = {}

            for eintrag in baum:
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag  
                dict_baum.update({ordinal:(parent,name,lvl,art,zustand,sicht)})
            
            def suche_parent(ord_kind):
                ord_parent = dict_baum[ord_kind][0]
                return ord_parent
            
            for eintrag in baum:
                pfad2 = speicherordner
                ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
                
                ordner = []
                ord = ordinal
                for i in range(int(lvl)):
                    ord = suche_parent(ord)
                    ordner.append(dict_baum[ord][1])
                
                ordner.reverse()  
                for ordn in ordner:
                    pfad2 = os.path.join(pfad2, ordn)
                    
                pfad2 = os.path.join(pfad2, dict_baum[ordinal][1])
                if art in ('dir','prj'):
                    pfad2 = os.path.join(pfad2, name)
                    
                    
                pfade.update({ordinal:(name,pfad2,art)})
        
        
        
        
        # Statusindicator
        anz_sections = len(sections)
        st_ind.start('exportiere, bitte warten',anz_sections)
        zaehler = 0
        
        
        
        
        

        while zaehler < anz_sections - 1:
             
            self.mb.class_Bereiche.starte_oOO()
            oOO = self.mb.class_Bereiche.oOO
            cur = oOO.Text.createTextCursor()
            text = oOO.Text
             
            # Speichern     
            for i in range(3):
 
                if zaehler  > len(sections) - 1:
                    break
                 
                 
                sec_name = sections[zaehler]
                 
                if 'OrganonSec' in sec_name:  
                     
                    zaehler += 1       
                     
                    sec_ordinal = self.mb.dict_bereiche['Bereichsname-ordinal'][sec_name]
                    if sec_ordinal == self.mb.Papierkorb:
                        break  
                     
                    cur.gotoStart(False)
                    cur.gotoEnd(True)
                    cur.setString('')
                     
                    sec = self.mb.doc.TextSections.getByName(sec_name)
                    SFLink = sec.FileLink
                     
                    newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                    newSection.setPropertyValue('FileLink',SFLink)
                    newSection.setName(sec.Name)
                     
                    oOO.Text.insertTextContent(cur,newSection,False)
                    oOO.Text.removeTextContent(newSection)
                     
                    contr = self.mb.Hauptfeld.getControl(sec_ordinal)
                    tf = contr.getControl('textfeld')
                    titel = tf.Model.Text
                     
                    if not set['ordner_strukt']:
                        pfad = os.path.join(speicherordner, titel)
                    else:
                        pfad = pfade[sec_ordinal][1]
                         
                    # pruefen, ob datei existiert; Namen aendern       
                    if os.path.exists(pfad+set['typ']):
                        pfad = self.pruefe_dateiexistenz(pfad,set['typ'])
                    print(pfad)
                    path = pfad + set['typ'] 
                    path2 = uno.systemPathToFileUrl(path)  
                    oOO.storeToURL(path2,())
                 
                # unterbricht while Schleife, wenn nur Trenner und keine keine OrganonSec mehr uebrig sind 
                if self.teste_auf_verbliebene_bereiche(sections[zaehler::]):
                    zaehler = anz_sections
                     
            self.mb.class_Bereiche.schliesse_oOO()    
            st_ind.setValue(zaehler)

        st_ind.end()  

    def teste_auf_verbliebene_bereiche(self,sections):
        
        regex = re.compile('OrganonSec')
        matches = [string for string in sections if re.match(regex, string)]
        
        if len(matches) == 0:
            return True
        else:
            return False


    def pruefe_dateiexistenz(self,pfad,dateierweiterung = None):
        
        if dateierweiterung != None:
            i = 0
            while os.path.exists(pfad+dateierweiterung):
                pfadX = pfad.split(dateierweiterung)[0]
                if pfadX[-1] == ')':                                    
                    sub = re.sub(r"\([0-9]+\)",'('+str(i)+')', pfadX) 
                    if sub == pfad:
                        pfad = pfad + '(%s)' %i
                    else:
                        pfad = sub
                else:
                    pfad = pfad + '(%s)' %i
                i += 1
            return pfad
        
        else:
            i = 0
            while os.path.exists(pfad):
                if pfad[-1] == ')':                                    
                    sub = re.sub(r"\([0-9]*\)",'('+str(i)+')', pfad) 
                    if sub == pfad:
                        pfad = pfad + '(%s)' %i
                    else:
                        pfad = sub
                else:
                    pfad = pfad + '(%s)' %i
                i += 1
            return pfad
                        
from com.sun.star.lang import XEventListener
class AB_Fenster_Dispose_Listener(unohelper.Base, XEventListener):
    # Listener um Position zu bestimmen
    def __init__(self,mb,cl_exp):
        self.mb = mb
        self.cl_exp = cl_exp
        
    def disposing(self,ev):

        if ev.Source.Model.Text == lang.TRENNER_TIT:
            self.cl_exp.trenner_fenster = None
        if ev.Source.Model.Text == lang.AUSWAHL:
            self.cl_exp.auswahl_fenster = None
            
        if ev.Source.Model.Text == lang.EXPORT:
            if self.cl_exp.auswahl_fenster != None:
                self.cl_exp.auswahl_fenster.dispose()
                self.cl_exp.auswahl_fenster = None
            if self.cl_exp.trenner_fenster != None:
                self.cl_exp.trenner_fenster.dispose()
                self.cl_exp.trenner_fenster = None
               
            # Settings speichern
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
            
     
class A_Trenner_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,cl_exp,exp_fenster):
        self.mb = mb
        self.exp_fenster = exp_fenster
        self.cl_exp = cl_exp
        
    def actionPerformed(self,ev):      
        
        if self.cl_exp.trenner_fenster != None:
            return

        posSize = berechne_pos(self.mb,self.cl_exp,self.exp_fenster,'Trenner')
        
        set = self.mb.settings_exp
        cb_listener = A_Trenner_CheckBox_Listener(self.mb)        

        posSize = posSize[0],posSize[1],320,360
        fenster,fenster_cont = erzeuge_Dialog_Container(self.mb.smgr,self.mb.ctx,posSize)
        fenster_cont.Model.Text = lang.TRENNER_TIT
        listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
        fenster_cont.addEventListener(listenerF)
        self.cl_exp.trenner_fenster = fenster

        y = 10
        
        # Titel
        controlE, modelE = createControl(self.mb.ctx,"FixedText",20,y ,80,20,(),() )  
        controlE.Text = lang.TRENNER_TIT
        modelE.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlE)
        
        y += 40
        
        # Ordner
        controlO, modelO = createControl(self.mb.ctx,"CheckBox",20 ,y,80,22,(),() )  
        modelO.Label = lang.ORDNERTITEL
        modelO.State = set['ordnertitel']
        controlO.ActionCommand = 'ordnertitel'
        controlO.addActionListener(cb_listener)
        fenster_cont.addControl('Ordnertitel', controlO)
        
        controlF, modelF = createControl(self.mb.ctx,"CheckBox",20 + 100 ,y,160,22,(),() )  
        modelF.Label = lang.FORMAT
        modelF.State = set['format_ord']
        controlF.ActionCommand = 'format_ord'
        controlF.addActionListener(cb_listener)
        fenster_cont.addControl('Format', controlF)
        
            # Liste der Formate
        controlL, modelL = createControl(self.mb.ctx,"ListBox",20 + 180,y -3 ,100,20,(),() )  
        #controlL.setMultipleMode(False)
        pStyles = self.mb.doc.StyleFamilies.ParagraphStyles
        style_names = pStyles.ElementNames
        controlL.addItems(style_names,0)
        modelL.Dropdown = True
        index = style_names.index(set['style_ord'])
        modelL.SelectedItems = index,
        fenster_cont.addControl('Liste_Ord', controlL)
        
        y += 30
        
        # Datei
        controlD, modelD = createControl(self.mb.ctx,"CheckBox",20 ,y,80,22,(),() )  
        modelD.Label = lang.DATEITITEL
        modelD.State = set['dateititel']
        controlD.ActionCommand = 'dateititel'
        controlD.addActionListener(cb_listener)
        fenster_cont.addControl('Dateititel', controlD)
        
        controlF2, modelF2 = createControl(self.mb.ctx,"CheckBox",20 + 100 ,y,160,22,(),() )  
        modelF2.Label = lang.FORMAT
        modelF2.State = set['format_dat']
        controlF2.ActionCommand = 'format_dat'
        controlF2.addActionListener(cb_listener)
        fenster_cont.addControl('Format2', controlF2)
        
            # Liste der Formate
        controlL2, modelL2 = createControl(self.mb.ctx,"ListBox",20 + 180,y -3 ,100,20,(),() )  
        #controlL.setMultipleMode(False)
        controlL2.addItems(style_names,0)
        modelL2.Dropdown = True
        index = style_names.index(set['style_dat'])
        modelL2.SelectedItems = index,
        fenster_cont.addControl('Liste_Dat', controlL2)
            # Listener fuer beide Stylelisten
        listenerLB = A_ParaStyle_Item_Listener(self.mb,controlL,controlL2)
        controlL.addItemListener(listenerLB)
        controlL2.addItemListener(listenerLB)
        
        y += 50
        
        # DOKUMENT
        controlD, modelD = createControl(self.mb.ctx,"FixedText",100,y ,200,20,(),() )  
        controlD.Text = lang.ORT_DES_DOKUMENTS
        modelD.FontWeight = 200.0
        fenster_cont.addControl('Titel2', controlD)
        
        y += 50
        
        controlL2, modelL2 = createControl(self.mb.ctx,"CheckBox",20 ,y,160,22,(),() )  
        modelL2.Label = lang.LEERZEILEN
        modelL2.State = set['leerzeilen_drunter']
        controlL2.ActionCommand = 'leerzeilen_drunter'
        controlL2.addActionListener(cb_listener)
        fenster_cont.addControl('Leerzeilen2', controlL2)     
        
        controlA2, modelA2 = createControl(self.mb.ctx,"Edit",120 ,y,20,30,(),() )  
        modelA2.HelpText = lang.ANZAHL_LEERZEILEN
        modelA2.Text = set['anz_drunter']
        listenerLZ = A_Anz_Leerzeilen_Focus_Listener(self.mb)
        controlA2.addFocusListener(listenerLZ)
        fenster_cont.addControl('Anzahl', controlA2) 
        
        y += 50
        
        controlDo, modelDo = createControl(self.mb.ctx,"CheckBox",20 ,y,160,22,(),() )  
        modelDo.Label = lang.DOK_EINFUEGEN
        modelDo.State = set['dok_einfuegen']
        controlDo.ActionCommand = 'dok_einfuegen'
        controlDo.addActionListener(cb_listener)
        fenster_cont.addControl('Dokument', controlDo)
        
        # Button
        controlD, modelD = createControl(self.mb.ctx,"Button",200,y-3,100,22,(),() )  
        controlD.Label = lang.WAEHLEN
        fenster_cont.addControl('Dok', controlD)
        
        y += 20
        
        controlF, modelF = createControl(self.mb.ctx,"FixedText",40 ,y,500,22,(),() )  
        modelF.HelpText = 'URL'
        #modelF.Border = True
        if self.mb.settings_exp['url'] != '':
            modelF.Label = uno.fileUrlToSystemPath(decode_utf(set['url']))#.decode("utf-8"))
        fenster_cont.addControl('Anzahl', controlF) 
        
        listener = A_TrennDatei_Button_Listener(self.mb,modelF)
        controlD.addActionListener(listener)
        
        y += 40
        
        controlSB, modelSB = createControl(self.mb.ctx,"CheckBox",20 ,y,200,22,(),() )  
        modelSB.Label = lang.SEITENUMBRUCH_ORD
        modelSB.State = set['seitenumbruch_ord']
        controlSB.ActionCommand = 'seitenumbruch_ord'
        controlSB.addActionListener(cb_listener)
        fenster_cont.addControl('seitenumbruch_ord', controlSB) 
        
        y += 20
        
        controlSb2, modelSb2 = createControl(self.mb.ctx,"CheckBox",20 ,y,200,22,(),() )  
        modelSb2.Label = lang.SEITENUMBRUCH_DAT
        modelSb2.State = set['seitenumbruch_dat']
        controlSb2.ActionCommand = 'seitenumbruch_dat'
        controlSb2.addActionListener(cb_listener)
        fenster_cont.addControl('seitenumbruch_dat', controlSb2) 
        
     
         

class A_Trenner_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb):
        self.mb = mb
        
    def actionPerformed(self,ev):
        set = self.mb.settings_exp
        set[ev.ActionCommand] = self.toggle(set[ev.ActionCommand])

    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:
            return 1


class A_TrennDatei_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,model):
        self.mb = mb
        self.model = model
        
    def actionPerformed(self,ev):

        Filepicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
        if self.mb.settings_exp['url'] != '':
            Filepicker.setDisplayDirectory(self.mb.settings_exp['url'])
        Filepicker.execute()
    
        if Filepicker.Files == '':
            return
        
        filepath = Filepicker.Files[0]
        self.mb.settings_exp['url'] = filepath
        self.model.Label = uno.fileUrlToSystemPath(filepath)

class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,model):
        self.mb = mb
        self.model = model
        
    def actionPerformed(self,ev):
        
        Filepicker = createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        Filepicker.setDisplayDirectory(self.mb.settings_exp['speicherort'])
        Filepicker.execute()
        
        if Filepicker.Directory == '':
            return
        
        filepath = Filepicker.getDirectory()
        
        self.mb.settings_exp['speicherort'] = filepath
        self.model.Label = uno.fileUrlToSystemPath(filepath)

            

       
    
         
class A_ParaStyle_Item_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,cont_ord,cont_dat):
        self.mb = mb
        self.cont_ord = cont_ord
        self.cont_dat = cont_dat
        
    # XItemListener    
    def itemStateChanged(self, ev):  
        if ev.Source == self.cont_ord:
            self.mb.settings_exp['style_ord'] = ev.Source.Items[ev.Selected] 
        elif ev.Source == self.cont_dat:
            self.mb.settings_exp['style_dat'] = ev.Source.Items[ev.Selected]           
       
    
class A_Anz_Leerzeilen_Focus_Listener(unohelper.Base, XFocusListener):
    def __init__(self,mb):
        self.mb = mb
        
    # XItemListener    
    def focusLost(self, ev): 

        if ev.Source.Model.Text.isdigit():
            self.mb.settings_exp['anz_drunter'] = int(ev.Source.Model.Text)
        else:
            ev.Source.Model.Text = self.mb.settings_exp['anz_drunter']
                
    def focusGained(self,ev):
        return False  


def berechne_pos(mb,cl_exp,exp_fenster,Rufer):

    anderes_fenster = None
    
    if Rufer == 'Auswahl':
        if cl_exp.trenner_fenster != None:
            anderes_fenster = cl_exp.trenner_fenster
    elif Rufer == 'Trenner':
        if cl_exp.auswahl_fenster != None:
            anderes_fenster = cl_exp.auswahl_fenster
         
    if anderes_fenster != None:
        posSizePlus = anderes_fenster.PosSize
        XPlus = posSizePlus.Width + 20
    else:
        XPlus = 0
    
    posSize_main = mb.desktop.ActiveFrame.ContainerWindow.PosSize
    posSize_expWin = exp_fenster.PosSize
    
    X = posSize_expWin.X + XPlus
    Y = posSize_expWin.Y
    Width = posSize_expWin.Width
    Height = posSize_main.Height - 40

    posSize = X + Width + 20,Y,Width,Height
    
    return posSize


class B_Auswahl_Button_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,cl_exp,exp_fenster):
        self.mb = mb
        self.exp_fenster = exp_fenster
        self.cl_exp = cl_exp
        
        
    def actionPerformed(self,ev):
        
    
        if self.cl_exp.auswahl_fenster != None:
            return
        
        posSize = berechne_pos(self.mb,self.cl_exp,self.exp_fenster,'Auswahl')
        posSize = posSize[0],posSize[1],400,posSize[3]

        set = self.mb.settings_exp

        # Dict von alten Eintraegen bereinigen
        eintr = []
        for ordinal in set['ausgewaehlte']:
            if ordinal not in self.mb.dict_bereiche['ordinal']:
                eintr.append(ordinal)
        for ord in eintr:
            del set['ausgewaehlte'][ord]

        fenster,fenster_cont = erzeuge_Dialog_Container(self.mb.smgr,self.mb.ctx,posSize)
        # Listener um Position zu bestimmen
        fenster_cont.Model.Text = lang.AUSWAHL
        listenerF = AB_Fenster_Dispose_Listener(self.mb,self.cl_exp)
        fenster_cont.addEventListener(listenerF)
        self.cl_exp.auswahl_fenster = fenster
        
        
        control_innen, model = createControl(self.mb.ctx,"Container",20,0,posSize[2],posSize[3],(),() )  
        model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
        fenster_cont.addControl('Huelle', control_innen)
        
        
        y = self.erzeuge_auswahl(control_innen)
        control_innen.setPosSize(0, 0,0,y + 20,8)

        self.setze_hoehe_und_scrollbalken(y,posSize[3],fenster,fenster_cont,control_innen)
        

    def setze_hoehe_und_scrollbalken(self,y,y_desk,fenster,fenster_cont,control_innen):  
        
        
        if y < y_desk-20:
            fenster.setPosSize(0,0,0,y + 20,8) 
        else:

            Attr = (0,0,20,y_desk,'scrollbar', None)    
            PosX,PosY,Width,Height,Name,Color = Attr
            
            # SCROLLBAR
            control, model = createControl(self.mb.ctx,"ScrollBar",PosX,PosY,Width,Height,(),() )  
            model.BackgroundColor = KONST.EXPORT_DIALOG_FARBE
            model.Orientation = 1
            model.BorderColor = KONST.EXPORT_DIALOG_FARBE
            model.LiveScroll = True        
            model.ScrollValueMax = y/2
            
            listener = B_Auswahl_ScrollBar_Listener(self.mb,control_innen)
            control.addAdjustmentListener(listener) 
            
            fenster_cont.addControl('ScrollBar',control)  

        
        
    def erzeuge_auswahl(self,fenster_cont):
        
        set = self.mb.settings_exp
        
        tree = self.mb.xml_tree
        root = tree.getroot()
        
        baum = []
        self.mb.class_XML.get_tree_info(root,baum)
        
        y = 10
        x = 10

        listener = B_Auswahl_CheckBox_Listener(self.mb,fenster_cont)
        
        #Titel
        control, model = createControl(self.mb.ctx,"FixedText",x,y ,300,20,(),() )  
        control.Text = lang.AUSWAHL_TIT
        model.FontWeight = 200.0
        fenster_cont.addControl('Titel', control)
        
        y += 30
        
        # Untereintraege auswaehlen
        control, model = createControl(self.mb.ctx,"FixedText",x + 40,y ,300,20,(),() )  
        control.Text = lang.ORDNER_CLICK
        model.FontWeight = 150.0
        fenster_cont.addControl('ausw', control)
        
        control, model = createControl(self.mb.ctx,"CheckBox",x+20,y ,20,20,(),() )  
        control.State = set['auswahl']
        control.ActionCommand = 'untereintraege_auswaehlen'
        control.addActionListener(listener)
        fenster_cont.addControl('Titel', control)

        y += 30
        
        for eintrag in baum:

            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            
            if art == 'waste':
                break
            
            control, model = createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,200,20,(),() )  
            control.Text = name
            fenster_cont.addControl('Titel', control)
            
            
            control, model = createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
            model.Border = False
            if art in ('dir','prj'):
                model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
            else:
                model.ImageURL = 'private:graphicrepository/res/sx03150.png' 
            fenster_cont.addControl('Titel', control)   
              
                
            control, model = createControl(self.mb.ctx,"CheckBox",x+20*int(lvl),y ,20,20,(),() )  
            control.addActionListener(listener)
            control.ActionCommand = ordinal+'xxx'+name
            if ordinal in set['ausgewaehlte']:
                model.State = set['ausgewaehlte'][ordinal][1]
            fenster_cont.addControl(ordinal, control)
            
            y += 20 
            
        return y   


from com.sun.star.awt import XAdjustmentListener
class B_Auswahl_ScrollBar_Listener (unohelper.Base,XAdjustmentListener):
    
    def __init__(self,mb,fenster_cont):        
        self.mb = mb
        self.fenster_cont = fenster_cont
        
    def adjustmentValueChanged(self,ev):
        self.fenster_cont.setPosSize(0, -ev.value.Value,0,0,2)
        
    def disposing(self,ev):
        return False

            
class B_Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,fenster_cont):
        self.mb = mb
        self.fenster_cont = fenster_cont
        
    def actionPerformed(self,ev):

        set = self.mb.settings_exp
        if ev.ActionCommand == 'untereintraege_auswaehlen':
            set['auswahl'] = self.toggle(set['auswahl'])
            self.mb.speicher_settings("export_settings.txt", self.mb.settings_exp) 
        else:
            ordinal,titel = ev.ActionCommand.split('xxx')
            state = ev.Source.Model.State
            set['ausgewaehlte'].update({ordinal:(titel,state)})

            if set['auswahl']:
                if ordinal in self.mb.dict_ordner:
                    
                    tree = self.mb.xml_tree
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
                    
                    for ord in ordinale:
                        if ord != self.mb.Papierkorb:
                            control = self.fenster_cont.getControl(ord)
                            control.Model.State = state
                            zeile = self.mb.Hauptfeld.getControl(ord)
                            titel = zeile.getControl('textfeld').Text
                            set['ausgewaehlte'].update({ord:(titel,state)}) 


    def toggle(self,wert):   
        if wert == 1:
            return 0
        else:              
            return 1           
        
            
            
             
################ TOOLS ################################################################

# Handy function provided by hanya (from the OOo forums) to create a control, model.
def createControl(ctx,type,x,y,width,height,names,values):
   smgr = ctx.getServiceManager()
   ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
   ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
   ctrl_model.setPropertyValues(names,values)
   ctrl.setModel(ctrl_model)
   ctrl.setPosSize(x,y,width,height,15)
   return (ctrl, ctrl_model)
def createUnoService(serviceName):
  sm = uno.getComponentContext().ServiceManager
  return sm.createInstanceWithContext(serviceName, uno.getComponentContext())

