# -*- coding: utf-8 -*-

import unohelper
from string import punctuation as Punctuation


class WListe():
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.win = None
        self.ctx = self.mb.ctx
        
        self.p = {}
        self.p['text1'] = LANG.DATEI_AUSSUCHEN
        self.p['text1_intern'] = 0
        self.p['text1_extern'] = 1

        self.p['ordner'] = LANG.ORDNER_AUSSUCHEN
        self.p['chronologisch'] = 1
        
        
        self.ord = {}
        self.ord['text1'] = None
        
        self.name = {}
        self.name['text1'] = None
        
        
        
    def start(self):
        if self.mb.debug: log(inspect.stack)
                
        X = self.mb.dialog.Size.Width
        Y = 30
        posSize = (X,Y,0,0)
        
        oWindow,cont = self.mb.erzeuge_Dialog_Container(posSize)
        self.erzeuge_Menu(cont,oWindow)

        
    def erzeuge_Menu(self,cont,oWindow):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            self.win = cont
            
            y = 20
            
            tab0 = tab0x = 10
            tab1 = tab1x = 150
            tab2 = tab2x = 120
            tab3 = tab3x = 120
            tabs = [tab0,tab1,tab2,tab3]
            
            listener = Speicherordner_Button_Listener(self.mb,self,oWindow)
                        
            design = self.mb.class_Design
            design.set_default(tabs)

            controls = (
            10,
            ('titel',"FixedText",               'tab0',y,250,20,  ('Label','FontWeight'),(LANG.WOERTERLISTE ,150),                 {} ),
            30,
            
            ('control0',"Button",               'tab0',y,80,22,  ('Label',),('Text 1',),                                            {'setActionCommand':'text1','addActionListener':(listener)} ),
            0,
            ('rb1_txt1',"RadioButton",          'tab1x',y,80,22,  ('Label','State'),(LANG.INTERN,self.p['text1_intern']),           {'setActionCommand':'text1 intern','addActionListener':(listener)} ),
            0,
            ('rb2_txt1',"RadioButton",          'tab2x',y,80,22,  ('Label','State'),(LANG.EXTERN,self.p['text1_extern']),           {'setActionCommand':'text1 extern','addActionListener':(listener)} ),
            30,
            ('text1',"FixedText",               'tab0',y,250,20,  ('Label',),(self.p['text1'],),                                    {} ),
            30,
            
            ('control4',"Button",               'tab0',y,100,22,  ('Label',),(LANG.SPEICHERORT,),                                   {'setActionCommand':'speicherordner','addActionListener':(listener)} ),
            30,
            ('ordner',"FixedText",              'tab0',y,250,20,  ('Label',),(self.p['ordner'],),                                   {} ),
            30,
            
            ('control5',"FixedLine",              'tab0',y,250,20,  (),(),                                                          {} ),
            30,
            
            ('text3',"FixedText",               'tab0',y,250,20,  ('Label',),(LANG.SORTIERUNG,),                                    {} ),
            0,
            ('rb_sort1',"RadioButton",          'tab1x',y,80,22,  ('Label','State'),(LANG.CHRONOLOGISCH,self.p['chronologisch']),   {'setActionCommand':'chronologisch','addActionListener':(listener)} ),
            20,
            ('rb_sort2',"RadioButton",          'tab1x',y,80,22,  ('Label',),(LANG.ALPHABETISCH,),                                  {} ),
            40,
            
            ('control18',"Button",              'tab2x',y,140,30,  ('Label',),(LANG.START,),                                        {'setActionCommand':'search','addActionListener':(listener)} ),
            30,)
            
            ctrls = {}
            pos_y = 0
            
            for ctrl in controls:
                
                if isinstance(ctrl,int):
                    pos_y += ctrl
                    
                else:

                    name,unoCtrl,X,Y,width,height,prop_names,prop_values,extras = ctrl
                    pos_x = locals()[X]
                    
                    locals()[name],locals()[name+'_model'] = self.mb.createControl(self.ctx,unoCtrl,pos_x,pos_y,width,height,prop_names,prop_values)
                    
                    try:
                        w,h = self.mb.kalkuliere_und_setze_Control(locals()[name],'w')
                    except:
                        pass
                    
                    if 'x' in X:
                        design.setze_tab(X,w)        
                     
                    if 'setActionCommand' in extras:
                        locals()[name].setActionCommand(extras['setActionCommand'])
                        
                    if 'addActionListener' in extras:
                        #for l in extras['addActionListener']:
                        locals()[name].addActionListener(extras['addActionListener'])

                    self.win.addControl(name,locals()[name])
                    ctrls[name] = locals()[name]

                    
                    
            # Tabs x-Position neu berechnen
            design.kalkuliere_tabs()

            for i in range(len(tabs)):
                locals()['tab%sx'%i] = design.new_tabs['tab%sx'%(i)]                
            
            breite = 0
            
            for ctrl in controls:
                if isinstance(ctrl,int):
                    pass
                else:
                    name,unoCtrl,X,Y,width,height,prop_names,prop_values,extras = ctrl                    
                    pos_x = design.new_tabs[X]
                    
                    # Sonderregeln
                    if X == 'tab2x':
                        pos_x -= 50
                    
                    locals()[name].setPosSize(pos_x,0,0,0,1)
                    
                    if pos_x > breite:
                        breite = pos_x
                        
            breite += 60
            
            listener.controls = ctrls
            
            locals()['text1'].setPosSize(0,0,breite-20,0,4)
            locals()['ordner'].setPosSize(0,0,breite-20,0,4)
            
            locals()['control5'].setPosSize(0,0,breite-20,0,4)
            
            locals()['control18'].setPosSize(breite - locals()['control18'].PosSize.Width-10,0,0,0,1)

            self.win.setPosSize(0,0,breite,pos_y + 10,12)
            oWindow.setPosSize(0,0,breite,pos_y + 10,12)

        except:
            log(inspect.stack,tb())
            
            

from com.sun.star.awt import XActionListener
class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,class_Zitate,oWindow):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.controls = None
        self.nachricht = mb.nachricht
        self.class_Zitate = class_Zitate
        self.oWindow = oWindow
        
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        try:
            if ev.ActionCommand == 'text1':
                if self.class_Zitate.p['text1_extern'] == 1:
                    self.filepicker('text1')
                else:
                    self.get_interne_auswahl('text1')
                    
            elif ev.ActionCommand == 'speicherordner':
                self.folderpicker()
                
            elif ev.ActionCommand == 'search':
                args = self.suchbefehle_erstellen()

                (pfad1,
                text1_intern,
                benutze_txt1_intern,
                speicherordner,
                chronologisch) = args
                
                if not os.path.exists(speicherordner):
                    self.nachricht(LANG.KEIN_SPEICHERORT)
                    return
                                                
                SI = self.mb.desktop.getCurrentFrame().createStatusIndicator()
                SI.start('',3)
                
                if benutze_txt1_intern == 0:             
                    if self.pruefe_pfad(pfad1,'Text 1'):           
                        text1,name1 = self.oeffne_text(pfad1)
                        SI.setValue(1)
                    else:
                        SI.end()
                        return
                else:
                    text1 = self.get_internal_text(text1_intern)
                    name1 = self.class_Zitate.name['text1']
                    SI.setValue(1)
                    
                    
                args = text1,name1,speicherordner,chronologisch
                s = Liste_Erstellen(args,self.mb) 
                
                SI.setValue(2)
                
                s.run()
                
                SI.end()

            elif ev.ActionCommand == 'text1 intern':
                self.class_Zitate.p['text1_intern'] = 1
                self.class_Zitate.p['text1_extern'] = 0
            elif ev.ActionCommand == 'text1 extern':
                self.class_Zitate.p['text1_intern'] = 0
                self.class_Zitate.p['text1_extern'] = 1
                
        except:
            log(inspect.stack,tb())
    
    
    def pruefe_pfad(self,pfad,text):
        if self.mb.debug: log(inspect.stack)
        
        if not os.path.exists(pfad):
            ntext = LANG.NOCH_NICHT_AUSGESUCHT %text
            self.mb.nachricht(ntext)
            return False
        
        return True
    
     
    def get_internal_text(self,ordinal):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props[T.AB]
                    
        if ordinal in props.dict_ordner:
            ordinale = props.dict_ordner[ordinal]
        else:
            ordinale = [ordinal]
        
        text = []
#         for o in ordinale:
#             sec_name = props.dict_bereiche['ordinal'][o]
#             sec = self.mb.doc.TextSections.getByName(sec_name)
#             t = sec.Anchor.String.splitlines()
#             text.extend(t)

        pfade = []
        
        for o in ordinale:
            section = props.dict_bereiche['ordinal'][o]
            pfade.append( props.dict_bereiche['Bereichsname'][section])

        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        
        URL="private:factory/swriter"
        doc = self.mb.desktop.loadComponentFromURL(URL,'_blank',8+32,(prop,))
        
        try:
            newSection = self.mb.doc.createInstance("com.sun.star.text.TextSection")
            SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")

            text = doc.Text
            textSectionCursor = text.createTextCursor()
            text.insertTextContent(textSectionCursor, newSection, False)
            
            lines = []

            for p in pfade: 
                SFLink.FileURL = p
                SFLink.FilterName = 'writer8'
                newSection.setPropertyValue('FileLink',SFLink)
                
                lines.extend( doc.Text.String.splitlines() )
        except:
            log(inspect.stack,tb())
            doc.close(False)
            return []
        
        doc.close(False)

        return lines
    
    def get_interne_auswahl(self,text):  
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            ps = self.oWindow.PosSize
            X = ps.X + ps.Width +20
            Y = ps.Y 
            
            posSize = (X,Y,400,0)

            container,Y,listener = self.erzeuge_auswahl(text)
            oWindow,cont = self.mb.erzeuge_Dialog_Container(posSize)
            listener.window = oWindow
            
            y_desk = self.mb.current_Contr.ComponentWindow.PosSize.Height
            oWindow.setPosSize(0,0,0,y_desk,8) 
            self.setze_hoehe_und_scrollbalken(Y,y_desk,oWindow,cont,container)
            
            cont.addControl('Container',container)

        except:
            log(inspect.stack,tb())
        
    
    def erzeuge_auswahl(self,text):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props['ORGANON'].xml_tree
        root = tree.getroot()
        
        baum = []
        self.mb.class_XML.get_tree_info(root,baum)
        
        y = 10
        x = 10
        
        #Inneres Fenster
        controlContainer, modelContainer = self.mb.createControl(self.mb.ctx,"Container",22,y ,400,2000,(),() )  
        modelContainer.BackgroundColor = KONST.FARBE_ORGANON_FENSTER
        
        listener = Text_Intern_Listener(self.mb,self.controls[text],text)
        
        #Titel
        control, model = self.mb.createControl(self.mb.ctx,"FixedText",x,y ,300,20,(),() )  
        control.Text = LANG.AUSGEWAEHLTER_ORDNER_WAEHLT
        model.FontWeight = 200.0
        controlContainer.addControl('Titel', control)
        
        y += 30

        y2 = y
        for eintrag in baum:

            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            
            if art == 'waste':
                break
            
            control, model = self.mb.createControl(self.mb.ctx,"RadioButton",x+20*int(lvl),y ,20,20,(),() )  
            control.addActionListener(listener)
            control.ActionCommand = ordinal+'xxx'+name
            
            controlContainer.addControl(ordinal, control)
            
            y += 20 
            
        y = y2    
        for eintrag in baum:

            ordinal,parent,name,lvl,art,zustand,sicht,tag1,tag2,tag3 = eintrag
            
            if art == 'waste':
                break
            
            control, model = self.mb.createControl(self.mb.ctx,"FixedText",x + 40+20*int(lvl),y ,200,20,(),() )  
            control.Text = name
            controlContainer.addControl('Titel', control)
            
            
            control, model = self.mb.createControl(self.mb.ctx,"ImageControl",x + 20+20*int(lvl),y ,16,16,(),() )  
            model.Border = False
            if art in ('dir','prj'):
                model.ImageURL = 'vnd.sun.star.extension://xaver.roemers.organon/img/Ordner_16.png' 
            else:
                model.ImageURL = 'private:graphicrepository/res/sx03150.png' 
            controlContainer.addControl('Titel', control)               
            
            y += 20 
            
        controlContainer.setPosSize(0,0,400,y,8)    
            
        return controlContainer,y,listener


    def setze_hoehe_und_scrollbalken(self,y,y_desk,fenster,fenster_cont,control_innen):  
        if self.mb.debug: log(inspect.stack)
        
        if y < y_desk-20:
            fenster.setPosSize(0,0,0,y + 20,8) 
            fenster_cont.setPosSize(0,0,0,y + 20,8) 
        else:
            try:
                PosSize = 0,0,0,y_desk 
                control = self.mb.erzeuge_Scrollbar(fenster_cont,PosSize,control_innen)
            except:
                log(inspect.stack,tb())
            

    def filepicker(self,ctrl):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FilePicker")
        Filepicker.appendFilter('Find Quotations','*.txt')
        Filepicker.execute()

        if Filepicker.Files == '':
            return
        filepath =  uno.fileUrlToSystemPath(Filepicker.Files[0])
        self.controls[ctrl].Model.Label = filepath
        
        
    def folderpicker(self):
        if self.mb.debug: log(inspect.stack)
        
        Filepicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        Filepicker.execute()
        
        if Filepicker.Directory == '':
            return
        
        filepath = uno.fileUrlToSystemPath(Filepicker.getDirectory())
        self.controls['ordner'].Model.Label = filepath
        
        
    def disposing(self,ev):
        return False


    def suchbefehle_erstellen(self):
        if self.mb.debug: log(inspect.stack)

        pfad1 =         self.class_Zitate.p['text1'] =   self.controls['text1'].Model.Label
        text1_intern =  self.class_Zitate.ord['text1']  
        
        benutze_txt1_intern = self.class_Zitate.p['text1_intern']
        
        speicherordner = self.class_Zitate.p['ordner'] = self.controls['ordner'].Model.Label
        
        if int(self.controls['rb_sort1'].State) == 1:
            
            chronologisch =  self.class_Zitate.p['chronologisch'] =  1
        else:
            chronologisch =  self.class_Zitate.p['chronologisch'] =  0

        args = (pfad1,
                text1_intern,
                benutze_txt1_intern,
                speicherordner,chronologisch)
        
        return args
    
    
    def oeffne_text(self,pfad):   
        if self.mb.debug: log(inspect.stack)
        
        extension = os.path.splitext(pfad)[1]
        name = os.path.basename(pfad).split('.')[0]
        
        if extension == '.txt':
 
            with codecs_open( pfad, "r",'utf-8') as file:
                text = file.readlines()
            
        else:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
    
            doc = self.mb.desktop.loadComponentFromURL(uno.systemPathToFileUrl(pfad),'_blank',8+32,(prop,))
            
            text = doc.Text.String.splitlines()
            doc.close(False)
        
        return text,name
    


class Text_Intern_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,ctrl,text):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.ctrl = ctrl
        self.text = text
        self.window = None
    
    def disposing(self,ev):
        return False
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)

        ordinal,name = ev.ActionCommand.split('xxx')
        
        self.ctrl.Model.Label = name
        self.mb.class_werkzeug_wListe.ord[self.text] = ordinal
        self.mb.class_werkzeug_wListe.name[self.text] = name
        
        self.window.dispose()


class Liste_Erstellen():
    
    def __init__(self,args,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.anzSW = 1
        
        (self.text1,
         self.titel_txt1,
         self.speicherordner,
         self.chronologisch
         ) = args
                 
        self.nachricht = mb.nachricht
        
        
    
    def run(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
                        
            posDict1, posDict_O1, WoerterDict1, WoerterListe1 = self.listen_und_dicts_erstellen(self.text1)

            if self.chronologisch:
                woerter_set = set(WoerterListe1)
                woerter = sorted(woerter_set, key= lambda elem : WoerterListe1.index(elem))
            else:
                woerter = sorted(set(WoerterListe1))

            if len(woerter) == 0:
                self.nachricht(LANG.KEINE_UEBEREINSTIMMUNGEN) 
                return
            
            self.oeffne_calc()
            
            sheet = self.calc.Sheets.getByIndex(0)            
            
            for i in range(len(woerter)):
                cell = sheet.getCellByPosition(0, i)
                cell.setString(woerter[i])
             
            path = os.path.join(self.speicherordner,'list_of_words.odt')
             
            self.calc.storeToURL(uno.systemPathToFileUrl(path),())    
            self.calc.close(False)
            
            self.nachricht(LANG.LISTE_GESPEICHERT.format(path),'infobox')
                                                
        except:
            log(inspect.stack,tb())
        
        
    def oeffne_calc(self):
        if self.mb.debug: log(inspect.stack)
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True
        
        URL="private:factory/scalc"
                
        self.calc = self.mb.doc.CurrentController.Frame.loadComponentFromURL(URL,'_blank',0,(prop,))


    def listen_und_dicts_erstellen(self,lines):     
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            zaehlung = ''
            
            posDict = {}
            posDict_O = {}
            WoerterDict = {}
            WoerterListe = []
            
            getrenntesWort = ''
            zaehler = 0
            
            anz_range = range(self.anzSW)

            regex = re.compile('[%s]' % re.escape(Punctuation))

            for l in lines:                
            
                if 'XXXXX' in l:
                    zaehlung = l.split('XXXXX')[0]
                else:
                    woerter = l.split()
                    woerter2 = l.split()
                    
                    if len(woerter) == 0:
                        w = posDict_O[zaehler-1][0]+'\n'
                        posDict_O.update({zaehler-1:(w,zaehlung)})
                        continue
                                        
                    woerter2[-1] = woerter2[-1]+'\n'
                    
                    # getrennte Woerter wieder zusammenfuegen
                    if getrenntesWort != '':
                        woerter[0] = getrenntesWort + woerter[0]
                        woerter2[0] = getrenntesWort2 + woerter2[0]+'\n'
                        getrenntesWort = ''
                     
                    if woerter[-1][-1] == '-':
                        getrenntesWort = woerter[-1][:-1]
                        getrenntesWort2 = woerter2[-1][:-1]
                        del woerter[-1] 
                        del woerter2[-1]                       
                    
                    if len(woerter) != len(woerter2):
                        a = len(woerter)
                        b = len(woerter2)
                    
                    # Suchwoerter zusammenfuegen
                    for u in range(len(woerter)):
                        
                        w = woerter[u]
                        w2 = woerter2[u]
                        
                        posDict_O.update({zaehler:(w2,zaehlung)})
                        w = w.lower()
                        w = regex.sub('', w)
                        posDict.update({zaehler:(w,zaehlung)})
            
                        # try/except statt if fuer schnellere Performance, da 
                        # zaehler < anzSW uebersprungen werden muss
                        try:                            
                            suchwoerter = ' '.join( [ ( posDict[ len(posDict) - (self.anzSW+1) + i ][0] ) for i in anz_range ] )
                                                    
                            WoerterListe.append(suchwoerter)
            
                            if suchwoerter in WoerterDict:
                                WoerterDict[suchwoerter].extend([zaehler-self.anzSW])
                            else:
                                liste_VK = [zaehler-self.anzSW]
                                WoerterDict.update({suchwoerter:liste_VK})
                            
                        except:
                            pass
                            
            
                        zaehler += 1
                     
            
            return posDict,posDict_O, WoerterDict, WoerterListe
        except:
            log(inspect.stack,tb())
            return {},{},{},[],{}
     

