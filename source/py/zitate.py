# -*- coding: utf-8 -*-

import unohelper
from string import punctuation as Punctuation


class Zitate():
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.win = None
        self.ctx = self.mb.ctx
        
        self.p = {}
        self.p['text1_intern'] = 0
        self.p['text1_extern'] = 1

        self.p['text2_intern'] = 0
        self.p['text2_extern'] = 1

        self.p['anz_SW'] = 5
        self.p['chronologisch'] = 1
        
        
        self.ord = {}
        self.ord['text1'] = None
        self.ord['text2'] = None
        
        self.name = {}
        self.name['text1'] = None
        self.name['text2'] = None
        
        
        
    def start(self):
        if self.mb.debug: log(inspect.stack)
        
        self.erzeuge_menu_zitate()

    
    def erzeuge_menu_zitate_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        y = 0
        
        controls = (
            10,
            ('titel',"FixedText",1,               
                                'tab0x',y,250,20,  
                                ('Label','FontWeight'),
                                (LANG.TEXTVERGLEICH ,150),                 
                                {} 
                                ),
            30,
            
            ('control0',"Button",1,               
                                'tab0',y,80,22,  
                                ('Label',),
                                ('Text 1',),                                            
                                {'setActionCommand':'text1','addActionListener':(listener)} 
                                ),
            0,
            ('radio_txt1a',"RadioButton", 1,         
                                'tab1',y,80,22,  
                                ('Label','State'),
                                (LANG.INTERN,self.p['text1_intern']),           
                                {'setActionCommand':'text1 intern','addActionListener':(listener)} 
                                ),
            20,
            ('radio_txt1b',"RadioButton",1,          
                                'tab1',y,80,22,  
                                ('Label','State'),
                                (LANG.EXTERN,self.p['text1_extern']),           
                                {'setActionCommand':'text1 extern','addActionListener':(listener)} 
                                ),
            20,
            ('text1',"FixedText",0,              
                                'tab0x-max',y,250,45,  
                                ('Label','MultiLine'),
                                (LANG.DATEI_AUSSUCHEN,True),                                    
                                {} 
                                ),
            50,
            
            ('control2',"Button",1,               
                                'tab0',y,80,22,  
                                ('Label',),
                                ('Text 2',),                                            
                                {'setActionCommand':'text2','addActionListener':(listener)} 
                                ),
            0,
            ('xradio_txt2a',"RadioButton",1,          
                                'tab1',y,80,22,  
                                ('Label','State'),
                                (LANG.INTERN,self.p['text2_intern']),           
                                {'setActionCommand':'text2 intern','addActionListener':(listener)} 
                                ),
            20,
            ('xradio_txt2b',"RadioButton",1,          
                                'tab1',y,80,22,  
                                ('Label','State'),
                                (LANG.EXTERN,self.p['text2_extern']),           
                                {'setActionCommand':'text2 extern','addActionListener':(listener)} 
                                ),
            20,
            ('text2',"FixedText",0,               
                                'tab0x-max',y,250,45,  
                                ('Label','MultiLine'),
                                (LANG.DATEI_AUSSUCHEN,True),                                    
                                {} 
                                ),
            50,
            
            ('control4',"Button",1,               
                                'tab0',y,100,22,  
                                ('Label',),
                                (LANG.SPEICHERORT,),                                   
                                {'setActionCommand':'speicherordner','addActionListener':(listener)} 
                                ),
            30,
            ('ordner',"FixedText",0,              
                                'tab0x-max',y,250,45,  
                                ('Label','MultiLine'),
                                (LANG.ORDNER_AUSSUCHEN,True),                                   
                                {} 
                                ),
            50,
            
            ('control5',"FixedLine",0,              
                                'tab0x-max',y,250,20,  
                                (),
                                (),                                                          
                                {} 
                                ),
            30,
            ('control6',"FixedText",1,            
                                'tab0x',y,130,20,  
                                ('Label',),
                                (LANG.ANZ_SUCHWOERTER,),                              
                                {} 
                                ),
            0,
            ('NumericField1',"NumericField",1,    
                                'tab1x',y,40,20,   
                                ('ValueMin','DecimalAccuracy','Value'),
                                (3,0,self.p['anz_SW']),   
                                {} 
                                ),
            40,
            ('text3',"FixedText",1,               
                                'tab0',y,250,20,  
                                ('Label',),
                                (LANG.SORTIERUNG,),                                    
                                {} 
                                ),
            0,
            ('arb_sort1',"RadioButton",1,          
                                'tab1',y,80,22,  
                                ('Label','State'),
                                (LANG.CHRONOLOGISCH,self.p['chronologisch']),   
                                {'setActionCommand':'chronologisch','addActionListener':(listener)} 
                                ),
            20,
            ('arb_sort2',"RadioButton",1,          
                                'tab1',y,80,22,  
                                ('Label',),
                                (LANG.ALPHABETISCH,),                                  
                                {} 
                                ),
            40,
            
            ('control18',"Button",1,              
                                'tab1-max',y,140,25,  
                                ('Label',),
                                (LANG.START,),                                        
                                {'setActionCommand':'search','addActionListener':(listener)} 
                                ),
            20,
            )
        # feste Breite, Mindestabstand
        tabs = {
                 0 : (None, 30),
                 1 : (None, 5),
                 
                 }
        
        abstand_links = 10
        controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                
        return controls2,max_breite
    
        
    def erzeuge_menu_zitate(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Speicherordner_Button_Listener(self.mb,self)
            
            controls,max_breite = self.erzeuge_menu_zitate_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls) 
            
            
            # Hauptfenster erzeugen
            posSize = None,None,max_breite,max_hoehe
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            #fenster_cont.Model.Text = LANG.EXPORT
            
            # Controls in Hauptfenster eintragen
            for name,c in sorted(ctrls.items()):
                fenster_cont.addControl(name,c)
            
            
            listener.controls = ctrls
            listener.oWindow = fenster

        except:
            log(inspect.stack,tb())
            

from com.sun.star.awt import XActionListener
class Speicherordner_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,class_Zitate):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.controls = None
        self.class_Zitate = class_Zitate
        self.oWindow = None
        
        
    def actionPerformed(self,ev):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if ev.ActionCommand == 'text1':
                if self.class_Zitate.p['text1_extern'] == 1:
                    self.filepicker('text1')
                else:
                    self.get_interne_auswahl('text1')
                    
            elif ev.ActionCommand == 'text2':
                if self.class_Zitate.p['text2_extern'] == 1:
                    self.filepicker('text2')
                else:
                    self.get_interne_auswahl('text2')
                    
            elif ev.ActionCommand == 'speicherordner':
                self.folderpicker()
                
            elif ev.ActionCommand == 'search':
                args = self.suchbefehle_erstellen()

                (pfad1,pfad2,
                text1_intern,text2_intern,
                benutze_txt1_intern,benutze_txt2_intern,
                speicherordner,anz_SW,chronologisch) = args
                
                
                if benutze_txt1_intern == 0:             
                    if self.pruefe_pfad(pfad1,'Text 1'):           
                        text1,name1 = self.oeffne_text(pfad1)
                    else:
                        return
                else:
                    text1 = self.get_internal_text(text1_intern)
                    name1 = self.class_Zitate.name['text1']
                    
                    
                if benutze_txt2_intern == 0:  
                    if self.pruefe_pfad(pfad2,'Text 2'):           
                        text2,name2 = self.oeffne_text(pfad2)
                    else:
                        return
                else:
                    text2 = self.get_internal_text(text2_intern)
                    name2 = self.class_Zitate.name['text2']
                                        
                if not os.path.exists(speicherordner):
                    ntext = LANG.KEIN_SPEICHERORT
                    self.mb.nachricht(ntext)
                    return
                
                args = text1,text2,name1,name2,speicherordner,anz_SW,chronologisch
                s = Suche(args,self.mb) 
                s.run()

                    

            elif ev.ActionCommand == 'text1 intern':
                self.class_Zitate.p['text1_intern'] = 1
                self.class_Zitate.p['text1_extern'] = 0
            elif ev.ActionCommand == 'text1 extern':
                self.class_Zitate.p['text1_intern'] = 0
                self.class_Zitate.p['text1_extern'] = 1
            elif ev.ActionCommand == 'text2 intern':
                self.class_Zitate.p['text2_intern'] = 1
                self.class_Zitate.p['text2_extern'] = 0
            elif ev.ActionCommand == 'text2 extern':
                self.class_Zitate.p['text2_intern'] = 0
                self.class_Zitate.p['text2_extern'] = 1
                
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
        for o in ordinale:
            sec_name = props.dict_bereiche['ordinal'][o]
            sec = self.mb.doc.TextSections.getByName(sec_name)
            
            if sec.FileLink.FileURL == '':
                pfad = os.path.join(self.mb.pfade['odts'] , '%s.odt' % ordinal)
                fl = sec.FileLink
                fl.FileURL = pfad
                sec.setPropertyValue('FileLink',fl)
                
            t = sec.Anchor.String.splitlines()
            text.extend(t)
        
        return text
    
    def get_interne_auswahl(self,text):  
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            ps = self.oWindow.PosSize
            X = ps.X + ps.Width +20
            Y = ps.Y 
            
            posSize = (X,Y,400,0)

            container,Y,listener = self.erzeuge_auswahl(text)
            oWindow,cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            listener.window = oWindow
            
            y_desk = self.mb.current_Contr.ComponentWindow.PosSize.Height
            oWindow.setPosSize(0,0,0,y_desk,8) 
            self.setze_hoehe_und_scrollbalken(Y,y_desk,oWindow,cont,container)
            
            cont.addControl('Container',container)

        except:
            log(inspect.stack,tb())
        
    
    def erzeuge_auswahl(self,text):
        if self.mb.debug: log(inspect.stack)
        
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()
        
        baum = []
        self.mb.class_XML.get_tree_info(root,baum)
        
        y = 10
        x = 10
        
        #Inneres Fenster
        controlContainer, modelContainer = self.mb.createControl(self.mb.ctx,"Container",22,y ,400,2000,(),() )  
        
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
                control = self.mb.class_Fenster.erzeuge_Scrollbar(fenster_cont,PosSize,control_innen)
            except:
                log(inspect.stack,tb())
            

    def filepicker(self,ctrl):
        if self.mb.debug: log(inspect.stack)

        ofilter = ('Find Quotations','*.txt')
        filepath,ok = self.mb.class_Funktionen.filepicker2(ofilter=ofilter,url_to_sys=True)
        
        if not ok:
            return
        
        self.controls[ctrl].Model.Label = filepath
        
        
    def folderpicker(self):
        if self.mb.debug: log(inspect.stack)
        
        folderpicker = self.mb.createUnoService("com.sun.star.ui.dialogs.FolderPicker")
        folderpicker.execute()
        
        if folderpicker.Directory == '':
            return
        
        filepath = uno.fileUrlToSystemPath(folderpicker.getDirectory())
        self.controls['ordner'].Model.Label = filepath
        
        
    def disposing(self,ev):
        return False


    def suchbefehle_erstellen(self):
        if self.mb.debug: log(inspect.stack)

        pfad1 = self.controls['text1'].Model.Label
        pfad2 = self.controls['text2'].Model.Label
        
        text1_intern = self.class_Zitate.ord['text1']  
        text2_intern = self.class_Zitate.ord['text2']
        
        benutze_txt1_intern = self.class_Zitate.p['text1_intern']
        benutze_txt2_intern = self.class_Zitate.p['text2_intern']
        
        speicherordner = self.controls['ordner'].Model.Label
        
        anz_SW =        self.class_Zitate.p['anz_SW'] =  int(self.controls['NumericField1'].Value)
        
        chronologisch =  int(self.controls['arb_sort1'].State)
        self.class_Zitate.p['chronologisch'] = int(chronologisch)

        args = (pfad1,pfad2,
                text1_intern,text2_intern,
                benutze_txt1_intern,benutze_txt2_intern,
                speicherordner,anz_SW,chronologisch)

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
    
        
    def vergleiche_ordner(self,ordner1,ordner2,speicherordner):
        '''
        Bsp.
        ordner1 = 'C:\\Users\\Homer\\Desktop\\Texte1'
        ordner2 = 'C:\\Users\\Homer\\Desktop\\Texte2'
        speicherordner = 'C:\\Users\\Homer\\Desktop\\Ergebnis'
        self.vergleiche_ordner(ordner1,ordner2,speicherordner)
        '''
        
        try:
                    
            
            from zitate import Suche
            
            def oeffne_text(pfad):
                extension = os.path.splitext(pfad)[1]
                name = os.path.basename(pfad).split('.')[0]
                
                if extension == '.txt':
         
                    with codecs_open( pfad, "r",'utf-8') as file:
                        text = file.readlines()
                
                return name,text
            
            
            pfade_ordner1 = os.listdir(ordner1)
            pfade_ordner2 = os.listdir(ordner2)
            
            for pfad1 in pfade_ordner1:
                name1,text1 = oeffne_text(os.path.join(ordner1,pfad1))
                
                for pfad2 in pfade_ordner2:
                    
                    name2,text2  = oeffne_text(os.path.join(ordner2,pfad2))
                    anz_SW = 6
                    chronologisch = 1
                    
                    
                    args = text1,text2,name1,name2,speicherordner,anz_SW,chronologisch
                    s = Suche(args,self.mb) 
                    s.run()

        except:
            log(inspect.stack,tb())

    


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
        self.mb.class_Zitate.ord[self.text] = ordinal
        self.mb.class_Zitate.name[self.text] = name
        
        self.window.dispose()


class Suche():
    
    def __init__(self,args,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.ctx = self.mb.ctx
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext( "com.sun.star.frame.Desktop",self.ctx)
        self.doc = self.desktop.getCurrentComponent() 
                
        (self.text1,
         self.text2,
         name1,name2,
         self.speicherordner,
         self.anzSW,
         self.chronologisch
         ) = args
         
        
        self.titel = LANG.VERGLEICH_VON_MIT %(name1,name2)
        self.titel_txt1 = name1
        self.titel_txt2 = name2
                
        
    
    def run(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            SI = self.doc.CurrentController.Frame.createStatusIndicator()
            SI.start(LANG.ZITATE_FINDEN +' 1/4',5)
            SI.setValue(1)
            
            lines1,lines2 = self.text1,self.text2 
            
            posDict1, posDict_O1, WoerterDict1, WoerterListe1 = self.listen_und_dicts_erstellen(lines1)
            posDict2, posDict_O2, WoerterDict2, WoerterListe2 = self.listen_und_dicts_erstellen(lines2)
            
            SI.setText(LANG.ZITATE_FINDEN +' 2/4')
            SI.setValue(2)
            
            set_t1 = set(WoerterListe1)
            set_t2 = set(WoerterListe2)
            
            uebereinstimmungen = set_t1.intersection(set_t2)  
            
            if len(uebereinstimmungen) == 0:
                SI.end()
                self.mb.nachricht(LANG.KEINE_UEBEREINSTIMMUNGEN) 
                return
            
            vks = {}     
            for u in uebereinstimmungen:
                vks.update( {u: (WoerterDict1[u],WoerterDict2[u]) } )
            
            
            SI.setText(LANG.ZITATE_FINDEN +' 3/4')
            SI.setValue(3)
            
            FS = self.erstelle_fundstellen(uebereinstimmungen,WoerterDict1,posDict1,posDict_O1,WoerterListe1,vks)
                        
            SI.setText(LANG.ZITATE_FINDEN +' 4/4')
            SI.setValue(4)

            zaehler, links_u_farbe, txt_nav = self.erzeuge_navigation(FS,posDict_O1,posDict_O2)
            txt1, txt2 = self.erzeuge_text(FS,posDict_O1,posDict_O2,links_u_farbe)
            
            txt_main = self.erzeuge_main_html(zaehler)

            self.speichern_und_oeffnen(txt_main,txt_nav,txt1,txt2)
                        
        except:
            log(inspect.stack,tb())
            
        SI.end()
        
    
    
    def erzeuge_main_html(self,zaehler):
        if self.mb.debug: log(inspect.stack)
        
        html = u'''
<!DOCTYPE html>

<head>
    <title>{0}</title>

    <meta charset="utf-8">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="keywords" content="">
     <meta charset="UTF-8">

    <style type="text/css" media="screen">
        a {{ text-decoration: none; color: #FF0000; }}
        p {{ font-size:10pt; }}
        b {{ font-size: 9; }}
        .col {{ background-color: #FFFF00; }}
    </style>


    <script type="text/javascript">

        function getDocHeight(doc) {{
            doc = doc || document;
            var body = doc.body, html = doc.documentElement;
            var height = Math.max( body.scrollHeight, body.offsetHeight, html.clientHeight, html.scrollHeight, html.offsetHeight );
            return height;
        }}

    </script>

</head>

<html>
    <body  style="min-width: 1000px;margin-top: -10px; background-color: #000080;">

        <noscript>
            <p>This site uses JavaScript. You must allow JavaScript in your browser.</p>
        </noscript>

        <div style="background-color: #E9E7E3;height: 90px;width: 100%;border-radius: 10px;">
        
            <h3 style="text-align: center;">{1}</h3>
            
            <img src="content/organon icon_120.png"
            style="height: 60px;
                position: absolute;
                right: 0;
                padding-right: 50px;
                margin-top: -35px; ;
                ">
                
            <p style="text-align: center;">{2}  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; {3}  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; {4} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; {5} </p>

            <div style="float: left;
                        width: 19.9%;
                        text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;
                        border-left:none">
                <b>{6}</b>
            </div>
            <div style="float: right;
                        width: 40%;
                        text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;
                        border-right: none">
                <b>{8}</b>
            </div>
            <div style="text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;">
                <b>{7}</b>
            </div>

            <iframe
                id="navigation"
                src="content/navigation.html"
                width="20%"
                height="500"
                style="float: left;
                    background-color: #FAFAE7;
                    overflow:hidden"
                frameborder="0">
            </iframe>

            <iframe
                id="txt2"
                src="content/txt2.html"
                width="40%"
                height="500"
                style="float: right;
                    background-color: #FAFAE7;
                    overflow:hidden  "
                frameborder="0">
            </iframe>

            <iframe
                id="txt1"
                src="content/txt1.html"
                width="40%"
                height="500"
                style="background-color:  #FAFAE7;
                    overflow:hidden"
                frameborder="0">
            </iframe>

            <script type="text/javascript">
                var frame1 =  document.getElementById('navigation')
                var frame2 =  document.getElementById('txt1')
                var frame3 =  document.getElementById('txt2')

                var height = getDocHeight(document)
                window.height = height   -5

                frame1.height = height-115
                frame2.height = height-115
                frame3.height = height-115

            </script>

        </div>

    </body>
</html>
 '''
        

        text = html.format('Find Quotations',
                           self.titel,
                           LANG.SUCHWOERTER.format(self.anzSW),
                           LANG.FUNDSTELLEN.format(zaehler[0]),
                           LANG.VERSIONEN_TXT1.format(zaehler[1]),
                           LANG.VARIANTEN_TXT2.format(zaehler[2]),
                           LANG.NAVIGATION,
                           self.titel_txt1,
                           self.titel_txt2
                           )

        return text.encode('ascii', 'xmlcharrefreplace')
    

    def erzeuge_text(self,FS,posDict_O1,posDict_O2,links_u_farbe):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            txt_anf = '''<!DOCTYPE html>
<html>

<head>
    <title>navigation</title>

    <meta charset="utf-8">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="keywords" content="">
    

    <style type="text/css" media="screen">
        a { text-decoration: none; color: #FF0000; }
        p { font-size:10pt; }
        b { font-size: 9; }
        .col { background-color: #FFFF00; }
    </style>
    

    <script language="JavaScript">
        <!--

        function go_txt1(x) {
            loc = "content/txt1.html#" + x;
            window.parent.document.getElementById('txt1').src = loc;            
        return
        }
        
        function go_txt2(x) {
            loc = "content/txt2.html#" + x;
            window.parent.document.getElementById('txt2').src = loc;            
        return
        }
        
        function go_nav(x) {
            loc = "content/navigation.html#" + x;
            window.parent.document.getElementById('navigation').src = loc;            
        return
        }

        //-->
    </script>
    
    
</head>

<body>
 '''
            txt_end = '''</body>
</html> '''
            
                        
            
            def get_dict(odict):
                
                a1 = odict['anfang_txt1']
                e1 = odict['ende_txt1']
                stelle_txt1 = odict['stelle_txt1']
                
                vers_txt2 = []
                for l in odict['links_txt2']:
                    vers_txt2.append('{0}X{1}'.format(*l))
                    
                txt = odict['txt_original']
                kinder = odict['kinder']
                
                return a1,e1,stelle_txt1,txt,vers_txt2,kinder

            
            links_txt1,links_txt2,farbe_txt1,farbe_txt2 = links_u_farbe

            # TEXT 1 + 2 zusammensetzen
            def text_zusammensetzen(farbige_stellen,stellen_sprung,posDict,txt_anf,txt_end,txt_vers):
                
                farbige_stellen = set(farbige_stellen)
                txt = [txt_anf]
                txt.append('\n<p>\n')
                
                zaehlung_alt = ''
                
                if txt_vers == 'txt1':
                    link = '\n<a name="{0}" href="javascript:void(0)" onClick="javascript:go_txt2(\'{1}\');javascript:go_nav(\'{2}\')"> \n'
                elif txt_vers == 'txt2':
                    link = '\n<a name="{0}" href="javascript:void(0)" onClick="javascript:go_txt1(\'{1}\');javascript:go_nav(\'{2}\')"> \n'
                
                
                for t in range(len(posDict)):
                    farbe = ''
                    farbe_e = ''
                    sprung = ''
                    sprung_e = ''
                    
                    
                    text_O = posDict[t][0]
                    zaehlung = posDict[t][1]
                    
                    if t in farbige_stellen:
                        farbe = '<span class=\'col\'>'
                        farbe_e = '</span>\n'

                    if t in stellen_sprung:
                        sprung = link.format(t,stellen_sprung[t][0],stellen_sprung[t][1])
                        sprung_e = '</a>'
                            
                    if zaehlung != zaehlung_alt:
                        txt.append('\n</p><p>\n' + zaehlung + '\n</p>\n<p>')
                        zaehlung_alt = zaehlung
                        
                    txt.append(farbe + sprung + text_O + sprung_e + farbe_e )
                    
                    if '\n' in text_O:
                        br = ''
                        for i in range(text_O.count('\n')):
                            br = br + '<br />'
                        txt.append(  br + '\n')
                        
                
                txt.append(txt_end)
                final = ' '.join(txt).encode('ascii', 'xmlcharrefreplace')
                
                return final
            
            final1 = text_zusammensetzen(farbe_txt1,links_txt1,posDict_O1,txt_anf,txt_end,'txt1')
            final2 = text_zusammensetzen(farbe_txt2,links_txt2,posDict_O2,txt_anf,txt_end,'txt2')
            
        except:
            
            log(inspect.stack,tb())
            return None, None
        
        return final1, final2    
    
    def erzeuge_navigation(self,ergebnis,posDict_O1,posDict_O2):
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            nav_prae = '''<!DOCTYPE html>
<html>

<head>
    <title>navigation</title>

    <meta charset="utf-8">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="keywords" content="">
    
    <script language="JavaScript">
        <!--

        function Sprung1(x) {
            loc = "content/txt1.html#" + x;
            window.parent.document.getElementById('txt1').src = loc;
            console.log(window,document)
            /*window.location.href = "#NavStop"; */
            return
        }
        
        function Sprung2(x) {
            loc = "content/txt2.html#" + x;
            window.parent.document.getElementById('txt2').src = loc;

            /*window.location.href = "#NavStop"; */
            return
        }

        //-->
    </script>

    <style type="text/css" media="screen">
        a { text-decoration: none; }
        p { font-size:10pt; }
        b { font-size: 9;}
    </style>
    
    
    
</head>

<body>
'''
            
            nav_end = '''</body>
</html> '''
            
            
            
            name = '<a name="{0}" '
            name1 = '<a style="color:#030393" name="{0}" '
            sprung = 'href="javascript:void(0)"onClick="javascript:Sprung1(\'{0}\');javascript:Sprung2(\'{1}\')">\n'
            link_e ='</a>\n'

            n = '\n'
            div_e = '</div>\n'
            
                  
            
            def get_dict(odict):
                stelle_txt1 = odict['stelle_txt1']
                a1 = odict['anfang_txt1']
                e1 = odict['ende_txt1']
                vers_txt2 = []
                for l in odict['links_txt2']:
                    vers_txt2.append(l)
                txt = odict['txt_original']
                kinder = odict['kinder']
                
                return a1,e1,txt,vers_txt2,kinder 
                             
                
            def tab(x):
                return '\t'*x
            
            def faerben_und_linken(a1,e1,txt2,nav,links_u_farbe):
                
                links_txt1,links_txt2,farbe_txt1,farbe_txt2 = links_u_farbe
                a2,e2 = txt2
                farbe_txt1.extend(range(a1,e1 + self.anzSW))
                farbe_txt2.extend(range(a2,e2 + self.anzSW))
                if a1 not in links_txt1:
                    links_txt1.update({a1:[a2,nav]})
                if a2 not in links_txt2:
                    links_txt2.update({a2:[a1,nav]})
                
                
        
            txt_liste = []
            text_navi = [nav_prae]
            
            
            links_txt1 = {}
            links_txt2 = {}
            farbe_txt1 = []
            farbe_txt2 = []
            
            links_u_farbe = links_txt1,links_txt2,farbe_txt1,farbe_txt2
            
            
            zaehler = 1
            zaehler_vers = 0
            zaehler_kind = 0
            
            zaehler_alle = 1
            
                        
            if self.chronologisch:
                ergebnis_keys = sorted(ergebnis)
            else:
                ergebnis_keys = sorted(ergebnis, key=lambda t: ergebnis[t]['txt'])
            

            for e in ergebnis_keys:  
                
                a1,e1,txt,vers_txt2,kinder = get_dict(ergebnis[e])
                          
                text_navi.append(tab(1) + '<div style="margin-left:0.1em;font-size:11pt;">\n')
                text_navi.append(tab(2) +'<b style="font-style: italic;">#%s</b><br />\n\n' %zaehler)
                text_navi.append(tab(2) + name1.format(zaehler_alle))
                text_navi.append(sprung.format(a1,vers_txt2[0][0]))
                text_navi.append(tab(2) + txt.strip() )
                text_navi.append(link_e + n)
                
                faerben_und_linken(a1,e1,vers_txt2[0],zaehler_alle,links_u_farbe)

                zaehler_alle += 1
                
                if len(vers_txt2) > 1:
                    t = []
                    text_navi.append(tab(3) + '<div style="margin-left:2em;font-size:10pt;">\n')
                    text_navi.append(tab(4) + '<b style="font-style: italic;color:grey">weitere(s) Vorkommen in Text 2</b><br />\n\n')  
                     
                    for v in range(1,len(vers_txt2)):

                        text_navi.append(tab(4) + name.format(zaehler_alle))
                        text_navi.append(sprung.format(a1,vers_txt2[v][0]))
                        text_navi.append(tab(4) + txt.strip())
                        text_navi.append(link_e + n)
                        text_navi.append(tab(4) +'<br />' + n + n)
                        zaehler_vers += 1
                        
                        faerben_und_linken(a1,e1,vers_txt2[v],zaehler_alle,links_u_farbe)
                        
                        zaehler_alle += 1
                        
                    text_navi.append(tab(3) + '</div>' + n)
                 
                if len(kinder) > 0:
                       
                    text_navi.append(tab(3) + '<div style="margin-left:2em;font-size:10pt;">\n')
                    text_navi.append(tab(4) + '<b style="font-style: italic;">Variante(n) in Text 1</b><br />\n\n')
                      
                    for k in kinder:
                          
                        a1,e1,txt,vers_txt2,kinder = get_dict(k)
  
                        text_navi.append(tab(4) + name1.format(zaehler_alle))
                        text_navi.append(sprung.format(a1,vers_txt2[0][0]))
                        text_navi.append(tab(4) + '- ' +txt.strip())
                        text_navi.append(link_e)
                        text_navi.append(tab(4) +'<br />' + n + n)
                        
                        faerben_und_linken(a1,e1,vers_txt2[0],zaehler_alle,links_u_farbe)
                        
                        zaehler_alle += 1
                        
                        if len(vers_txt2) > 1:

                            text_navi.append(tab(4) + '<div style="margin-left:2em;font-size:10pt;">\n')
                            text_navi.append(tab(5) + '<b style="font-style: italic;color:grey">weitere(s) Vorkommen in Text 2</b><br />\n\n')  
                                                     
                            for v in range(1,len(vers_txt2)):

                                text_navi.append(tab(5) + name.format(zaehler_alle))
                                text_navi.append(sprung.format(a1,vers_txt2[v][0]))
                                text_navi.append(tab(5) + txt.strip())
                                text_navi.append(link_e + n)
                                text_navi.append(tab(5) +'<br />' + n + n)
                                zaehler_vers += 1
                                
                                faerben_und_linken(a1,e1,vers_txt2[v],zaehler_alle,links_u_farbe)
                                
                                zaehler_alle += 1
                                
                            text_navi.append(tab(4) + '</div>' + n)
                            
                        
                        zaehler_kind += 1
                           
                    text_navi.append(tab(3) + div_e)   

                text_navi.append(tab(1) + div_e + n)
                text_navi.append(tab(2) +'<br />\n')
                
                zaehler += 1
               

            text_navi.append(nav_end)
            nav = ''.join(text_navi).encode('ascii', 'xmlcharrefreplace')
            
            return [zaehler -1 ,zaehler_kind,zaehler_vers], links_u_farbe, nav
        
        except:
            log(inspect.stack,tb())
            return None, None, None
           
    
    def speichern_und_oeffnen(self,txt_main,txt_nav,txt1,txt2):
        if self.mb.debug: log(inspect.stack)
            
        pfad = os.path.join(self.speicherordner,self.titel + ' %s searchterms' %self.anzSW)
        pfad_content = os.path.join(pfad,'content')
        
        # Python 3 verwandelt mit .encode('ascii', 'xmlcharrefreplace')
        # den String in Bytes. Deaher wird hier wieder dekodiert
        if self.mb.programm == 'LibreOffice':
            txt_main = txt_main.decode('utf-8')
            txt_nav = txt_nav.decode('utf-8')
            txt1 = txt1.decode('utf-8')
            txt2 = txt2.decode('utf-8')

        
        if not os.path.exists(pfad):
            os.makedirs(pfad)
        if not os.path.exists(pfad_content):
            os.makedirs(pfad_content)
        
        with codecs_open( os.path.join(pfad,'main.html'), "w",'utf-8') as file:
            file.write(txt_main)
            
        with codecs_open( os.path.join(pfad_content,'navigation.html'), "w",'utf-8') as file:
            file.write(txt_nav)
        with codecs_open( os.path.join(pfad_content,'txt1.html'), "w",'utf-8') as file:
            file.write(txt1)
        with codecs_open( os.path.join(pfad_content,'txt2.html'), "w",'utf-8') as file:
            file.write(txt2)
        
        from shutil import copyfile 
        
        s = os.path.join(self.mb.path_to_extension,'img')
        src = os.path.join(s,'organon icon_120.png')
        dst = os.path.join(pfad_content,'organon icon_120.png')
        copyfile(src, dst)
        
        import webbrowser
        webbrowser.open(os.path.join(pfad,'main.html'))
        

    def erstelle_fundstellen(self,uebereinstimmungen,WoerterDict1,posDict1,posDict_O1,WoerterListe1,vks):  
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            def vorkommen_sortiert_ausgeben(uebereinstimmungen,WoerterDict1):
                # vk = Vorkommen
                vk_txt1 = []
                vk_txt2 = []
                      
                for u in uebereinstimmungen:
                    
                    for vk in WoerterDict1[u]:
                        vk_txt1.append(vk)
                        
                return sorted(vk_txt1)
            
            sorted_vks1 = vorkommen_sortiert_ausgeben(uebereinstimmungen,WoerterDict1)
            
            
            def liste_mit_allen_vks_erstellen(sorted_vks,WoerterListe):
                
                text = []
                # Liste mit allen Vorkommen: Satz, vks in txt1 + txt2, index im posDict
                for s in sorted_vks:
                    satz = WoerterListe[s]
                    vk1 = vks[WoerterListe[s]][0][:]
                    vk2 = vks[WoerterListe[s]][1]
                        
                    text.append([s,satz,vk1,vk2])
                    
                return text
            
            text1 = liste_mit_allen_vks_erstellen(sorted_vks1,WoerterListe1)

                    
            def text_erstellen(zaehler,ende_txt1,posDict,posDict_O):
                zaehlung = []
                txt = []
                txt_O = []
                for w in range(zaehler,ende_txt1+self.anzSW):
                    txt.append(posDict[w][0])
                    txt_O.append(posDict_O[w][0])
                    if posDict[w][1] != [] and posDict[w][1] not in zaehlung:
                        zaehlung.append(posDict[w][1])
                txt = ' '.join(txt)
                txt_O = ' '.join(txt_O)
                
                return txt,txt_O,zaehlung
            
            
            def link_finden(anf_txt2,end_txt2,diff):
                links = []
                for a in anf_txt2:
                    for b in end_txt2:
                        if b - a == diff:
                            links.append([a,b])
                return links
            
       
            def aenderung_in_vk2(a1,b1):
                ab = [a for a in a1 if a+1 not in b1]
                zu = [b for b in b1 if b-1 not in a1]
                return('          ende',ab,'anfang',zu) 
                        
                    
            
            def texte_und_links_aller_vks_erstellen(text,posDict1,posDict_O1):
                
                    
                def einsortieren(anf,end,helf,eins,txt_alle,vk2_anfangsstellen,z2,laenge2):
                        
                    vs = {a:anf for a in vk2_anfangsstellen}
                    
                    if len(eins) == 0:

                        diff = end-anf
                        liste = [[a,a+diff] for a in vk2_anfangsstellen]
                        txt,txt_O,zaehlung = text_erstellen(anf,end,posDict1,posDict_O1)
                        txt_alle.append([anf,end,txt,txt_O,zaehlung,liste])
                        
                        laenge2.append([end-anf,z2])
                        z2 += 1
                        return z2
                    
                    else:
                        neu = []
                        
                        for ein in sorted(eins):
                            
                            anfaenge = eins[ein][3]
                            enden = eins[ein][1]
                            
                            # Anfang
                            if anfaenge != []:
                                vs.update({a:ein for a in anfaenge})
                                
                            # Ende
                            if enden != []:
                                
                                for ende in enden:
                                    diff = ein-anf
                                    
                                    for anfang in range(ende,ende-diff-1,-1):
                                        if anfang in vs:

                                            neu.append( [ [ vs[anfang], vs[anfang] + ende - anfang ], [[anfang,ende]] ] )
                                            
                                            del(vs[anfang])
                                            break
                                    
                         
                        for anfang in vs:
                            diff = end - vs[anfang] 
                            neu.append( [ [ vs[anfang], end ], [[anfang,anfang+diff]] ] )    
                             
                        for n in sorted(neu):
                            
                            if [ txt_alle[-1][0], txt_alle[-1][1] ]!= n[0]:
                            
                                txt,txt_O,zaehlung = text_erstellen(n[0][0],n[0][1],posDict1,posDict_O1)
                                txt_alle.append([n[0][0],n[0][1],txt,txt_O,zaehlung,n[1]])
                            
                                laenge2.append([n[0][1]-n[0][0],z2])
                                z2 += 1
                                
                            else:
                                txt_alle[-1][5].extend(n[1])
                            
                        return z2
                
                
                
                zaehler,satz,vk1,vk2 = text[0]

                laenge2 = []
                
                helfer4 = {text[0][0]:None}
                txt_alle = [[None,None]]
                helfer4alt = text[0][0]
                vk2_anfangsstellen = vk2
                
                einsortierte = {}
                
                z2 = 0
                
                for t in range (1,len(text)):
                    
                    zaehler1,satz1,vk1_1,vk2_1 = text[t-1]
                    zaehler,satz,vk1,vk2 = text[t]
                    
                    schnipsel1 = satz1.split()[1:]
                    schnipsel2 = satz.split()[:-1]

                    if schnipsel1 != schnipsel2 or zaehler1 +1 != zaehler:
                        
                        txt,txt_O,zaehlung = text_erstellen(helfer4alt,zaehler1,posDict1,posDict_O1)
                        helfer4[helfer4alt] = [zaehler1,txt,zaehlung,vk2_anfangsstellen]
                        
                        z2 = einsortieren(helfer4alt,zaehler1,helfer4,einsortierte,txt_alle,vk2_anfangsstellen,z2,laenge2)
                        helfer4alt = zaehler
                        einsortierte = {}
                        vk2_anfangsstellen = vk2
                        
                    else:
                        quersumme1 = sum(vk2_1) + len(vk2_1)
                        quersumme2 = sum(vk2)
                        
                        if quersumme1 != quersumme2:
                            aenderung = aenderung_in_vk2(vk2_1,vk2)
                            einsortierte.update({zaehler:aenderung})

                
                zaehler1,satz1,vk1_1,vk2_1 = text[-1]
                txt,txt_O,zaehlung = text_erstellen(helfer4alt,zaehler1,posDict1,posDict_O1)
                
                einsortieren(helfer4alt,zaehler1,helfer4,einsortierte,txt_alle,vk2_anfangsstellen,z2,laenge2)                
                
                del(txt_alle[0])

                
                
                return txt_alle,sorted(laenge2) 
            

            t1,l1 = texte_und_links_aller_vks_erstellen(text1,posDict1,posDict_O1)  


            def vorsortieren(txt,laenge2):

                nach_laenge_sortiert = [['None','None']]
                
                  
                laeng = copy.deepcopy(laenge2)
                             
                for l in laeng:
                    index = l[1]
                    nach_laenge_sortiert.append([txt[index][2],index])
                    
                    z2 = 1
                    tausch = False
                    z = len(nach_laenge_sortiert)-1
                    
                    while nach_laenge_sortiert[z-z2][0] == nach_laenge_sortiert[z][0]:
#                         print(nach_laenge_sortiert[z-z2][1], nach_laenge_sortiert[z][1])
#                         print(nach_laenge_sortiert[z-z2][0], nach_laenge_sortiert[z][0])
                        z2 += 1
                        tausch = True
                        #print(nach_laenge_sortiert[z-z2][0])
                        if nach_laenge_sortiert[z-z2][0] != None: 
                            if nach_laenge_sortiert[z-z2][0] in nach_laenge_sortiert[z-z2-1][0]:
                                tausch = False
                            #print(nach_laenge_sortiert[z-z2][0])
                    
                    if tausch:

                        z2 -= 1
                        y = z - 1
                        
                        obj = nach_laenge_sortiert.pop(z)
                        nach_laenge_sortiert.insert(z-z2,obj)
                        obj = laenge2.pop(y)
                        laenge2.insert(y-z2,obj)


                nach_laenge_sortiert.pop(0)    
                
                unikate = {}
                unikate_txt = []
                unikate_stelle = []
                
                einzusortierende = {}
                einz_txt = []
                einz_txt_stelle = []
                
                zuLoeschende = []
                
                gefunden = False
                
                for l in laenge2:
                    index = l[1]
                    kind = txt[index]
                    
                    gefunden = False
                    
                    for i in range(laenge2.index(l),len(nach_laenge_sortiert)): 
                        
                        parent = nach_laenge_sortiert[i]                         
                        p_index = parent[1]  
                                                
                        if kind[2] in parent[0] and index != p_index :
                            
                            if parent[1] in einzusortierende:
                                einzusortierende[parent[1]].extend([index])
                            else:
                                einzusortierende.update({parent[1]:[index]})
                                
                            if index in einzusortierende:
                                einzusortierende[parent[1]].extend(einzusortierende[index])
                                zuLoeschende.append(index)
                            
                            gefunden = True
                            break
                        
                    if not gefunden:
                        unikate.update({index:txt[index]})
                        unikate_txt.append(txt[index][2])
                        unikate_stelle.append(txt[index])
                     
                for z in zuLoeschende:
                    del(einzusortierende[z])
                
                

                # DOPPELTE AUSSORTIEREN
                from operator import itemgetter
                
                for e in einzusortierende:
                    
                    helfer = []
                    for t in einzusortierende[e]:
                        helfer.append([txt[t],t])
                    
                    helfer.sort(key=itemgetter(0))    
                    helfer.sort(key=lambda ein: len(ein[0][2]), reverse=True)
                    
                    
                    helfer2 = []
                    kontrolle1 = []
                    kontrolle2 = []

                    kontrolle1.extend(range(unikate[e][0],unikate[e][1]+1))
                         
                    for vks2 in unikate[e][5]:
                        kontrolle2.extend(range(vks2[0],vks2[1]+1))
                    
                    
                    for s in helfer:
                        range_1 = set(range(s[0][0],s[0][1]+1))
                        
                        VS = copy.deepcopy(s[0][5])
                     
                        for vs in VS:
                            range_2 = set(range(vs[0],vs[1]+1))
                            
                            if range_2.intersection(set(kontrolle2)) == range_2:
                                if range_1 == set(range_1).intersection(set(kontrolle1)):
                                    s[0][5].remove(vs)
                                else:
                                    kontrolle1.extend(range_1)
                            else:        
                                kontrolle2.extend(range_2)
                                kontrolle1.extend(range_1)
                                                                       
                    for s in helfer:
                        if s[0][5] == []:
                            continue
                        einz_txt.append(s[0][2])
                        einz_txt_stelle.append(s[0])
                        helfer2.append(s[1])
                        
                    
                    einzusortierende[e] = helfer2
                
                return [einzusortierende,einz_txt,einz_txt_stelle],[unikate,unikate_txt,unikate_stelle]
            

            einzusortierende1,unikate1 = vorsortieren(t1,l1)

            
            def erstelle_dict(anf,end,txt,txt_o,zae,links_txt2):
                
                ver = {'anfang_txt1': anf,
                       'ende_txt1': end,
                        'txt': txt,
                        'txt_original': txt_o,
                        'zaehlung': zae,
                        'laenge': len(txt),
                        'stelle_txt1':'%sX%s' %(anf,end),
                        'links_txt2' : links_txt2,
                        'kinder' : []
                        }
                return ver
            
            
            def FS_erstellen(t1,einzusortierende,unikate):
            
                FS = {}
                
                farbe_txt1 = []
                farbe_txt2 = []
                link_txt1 = {}
                link_txt2 = {}
                                
                uni,uni_txt,uni_st = unikate
                einz,einz_txt,einz_st = einzusortierende
                
                helfer = []
                
                stellen_txt1 = []
                stellen_txt2 = []
                
                for u in uni:
                    
                    dic = uni[u]
                    
                    anf = dic[0]
                    end = dic[1]
                    txt = dic[2]
                    txt_o = dic[3]
                    zae = dic[4]
                    links_txt2 = dic[5]
                    
                    d = erstelle_dict(anf,end,txt,txt_o,zae,links_txt2)
                    
                    FS.update({u:d})

                for e in einz:
                    for kind in einz[e]:
                        
                        dic = t1[kind]
                        
                        anf = dic[0]
                        end = dic[1]
                        txt = dic[2]
                        txt_o = dic[3]
                        zae = dic[4]
                        links_txt2 = dic[5]
                        
                        d = erstelle_dict(anf,end,txt,txt_o,zae,links_txt2)
                        
                        FS[e]['kinder'].append(d)

                
                return FS
            
            FS1 = FS_erstellen(t1,einzusortierende1,unikate1)
            
        except:
            log(inspect.stack,tb())
            return {}

        return FS1
 
  
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
     

