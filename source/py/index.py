# -*- coding: utf-8 -*-

import unohelper
from string import punctuation as Punctuation
from operator import itemgetter

class Index():
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.win = None
        self.ctx = self.mb.ctx        
        
        
    def start(self):
        if self.mb.debug: log(inspect.stack)

        self.erzeuge_dialog_index()
        

        
    def erzeuge_dialog_index_elemente(self,listener):
        if self.mb.debug: log(inspect.stack)
        
        try:
            y = 0
            
            controls = (
            10,
            ('titel',"FixedText",1,              
                                'tab0',y,250,20,  
                                ('Label','FontWeight'),
                                ('Create Index' ,150),                 
                                {} 
                                ),
            30,
            
            ('control0',"Button",1,               
                                'tab0',y,80,22,  
                                ('Label',),
                                ('Source',),                                            
                                {'setActionCommand':'text1','addActionListener':(listener)} 
                                ),
            30,
            ('text1',"FixedText",1,               
                                'tab0x-max',y,250,45,  
                                ('Label','MultiLine'),
                                (LANG.DATEI_AUSSUCHEN,True),                                    
                                {} 
                                ),
            45,
            ('control1',"Button",1,               
                                'tab0',y,80,22,  
                                ('Label',),
                                ('Translation',),                                            
                                {'setActionCommand':'text2','addActionListener':(listener)} 
                                ),
            30,
            ('text2',"FixedText",1,               
                                'tab0x-max',y,250,45,  
                                ('Label','MultiLine'),
                                (LANG.DATEI_AUSSUCHEN,True),                                    
                                {} 
                                ),
            45,
            ('control4',"Button",1,               
                                'tab0-tab0-E',y,100,22,  
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
            45,
            
            ('control5',"FixedLine",0,             
                                'tab0x-max',y,250,20,  
                                (),
                                (),                                                          
                                {} 
                                ),
            30,
            
            ('control18',"Button",1,              
                                'tab1',y,140,30,  
                                ('Label',),
                                (LANG.START,),                                        
                                {'setActionCommand':'search','addActionListener':(listener)} 
                                ),
            20,
            )
            
            # feste Breite, Mindestabstand
            tabs = {
                     0 : (None, 10),
                     1 : (None, 5),
                     
                     }
            
            abstand_links = 10
            controls2,tabs3,max_breite = self.mb.class_Fenster.berechne_tabs(controls, tabs, abstand_links)
                    
            return controls2,max_breite

        except:
            log(inspect.stack,tb())
            
    def erzeuge_dialog_index(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            listener = Speicherordner_Button_Listener(self.mb,self)
            
            controls,max_breite = self.erzeuge_dialog_index_elemente(listener)
            ctrls,max_hoehe = self.mb.class_Fenster.erzeuge_fensterinhalt(controls) 
            
            # Hauptfenster erzeugen
            posSize = None,None,max_breite,max_hoehe
            fenster,fenster_cont = self.mb.class_Fenster.erzeuge_Dialog_Container(posSize)
            #fenster_cont.Model.Text = LANG.EXPORT
            
            # Controls in Hauptfenster eintragen
            for name,c in sorted(ctrls.items()):
                fenster_cont.addControl(name,c)
            
            listener.oWindow = fenster
            listener.controls = ctrls
        
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
                self.filepicker('text1')
            elif ev.ActionCommand == 'text2':
                self.filepicker('text2')
                
                    
            elif ev.ActionCommand == 'speicherordner':
                self.folderpicker()
                
            elif ev.ActionCommand == 'search':
                args = self.suchbefehle_erstellen()

                (pfad1,pfad2,
                speicherordner,
                ) = args
                
                
                if self.pruefe_pfad(pfad1,'Source'):           
                    text1,name1 = self.oeffne_text(pfad1)
                else:
                    return
                
                if self.pruefe_pfad(pfad2,'Translation'):           
                    pass
                else:
                    return

                    
                                                            
                if not os.path.exists(speicherordner):
                    ntext = LANG.KEIN_SPEICHERORT
                    Popup(self.mb, 'warning').text = ntext
                    return
                
                args = text1,pfad1,speicherordner,pfad2
                s = Liste_Erstellen(args,self.mb) 
                s.run()


            
                
        except:
            log(inspect.stack,tb())
    
    
    def pruefe_pfad(self,pfad,text):
        if self.mb.debug: log(inspect.stack)
        
        if not os.path.exists(pfad):
            ntext = LANG.NOCH_NICHT_AUSGESUCHT %text
            Popup(self.mb, 'warning').text = ntext
            return False
        
        return True
    
    
        
    

            

    def filepicker(self,ctrl):
        if self.mb.debug: log(inspect.stack)

        ofilter = ('Source','*.txt;*.odt')
        filepath,ok = self.mb.class_Funktionen.filepicker2(ofilter=ofilter,url_to_sys=True)
        
        if not ok:
            return

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

        pfad1 = self.controls['text1'].Model.Label
        pfad2 = self.controls['text2'].Model.Label

        speicherordner = self.controls['ordner'].Model.Label

        args = (pfad1,pfad2,
                speicherordner)

        return args
    
    
    def oeffne_text(self,pfad):   
        if self.mb.debug: log(inspect.stack)
        
        extension = os.path.splitext(pfad)[1]
        name = os.path.basename(pfad)
        
        if extension == '.txt':
 
            with codecs_open( pfad, "r",'utf-8') as file:
                text = file.readlines()
            
        else:
            prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            prop.Name = 'Hidden'
            prop.Value = True
    
            doc = self.mb.desktop.loadComponentFromURL(uno.systemPathToFileUrl(pfad),'_blank',8+32,(prop,))
            
            text = doc.Text#.String.splitlines()
            doc.close(False)
        
        return text,name
    


class Liste_Erstellen():
    
    def __init__(self,args,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.anzSW = 1
        
        (self.text1,
         self.pfad_quelldatei,
         self.speicherordner,
         self.pfad_translation
         ) = args        
        
    
    def run(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
                        
            posDict1, posDict_O1, WoerterDict1, WoerterListe1 = self.listen_und_dicts_erstellen(self.text1)
            
            gr = u'αάὰἀἁἂἃἄἅἆᾳᾶᾷβγδεέὲἐἑἔἕζηήὴἠἡἢἣἤἥἧῃῆῇθιίὶϊἰἱἴἵἶἷῖκλμνξοόὸὀὁὃὄὅπρῥςστυύὺὐὑὓὔὕὖὗῦφχψωώὠὡὥὦὧὼᾤῳῴῶῷ'
            
            gr2 = (u'αάὰἀἁἂἃἄἅἆᾳᾶᾷ',
                    u'β',
                    u'δ',
                    u'γ',
                    u'εέὲἐἑἔἕ',
                    u'ηήὴἠἡἢἣἤἥἧῃῆῇ',
                    u'ιίὶϊἰἱἴἵἶἷῖ',
                    u'κ',
                    u'λ',
                    u'μ',
                    u'ν',
                    u'οόὸὀὁὃὄὅ',
                    u'ωώὠὡὥὦὧὼᾤῳῴῶῷ',
                    u'π',
                    u'φ',
                    u'ρῥ',
                    u'ς',
                    u'σ',
                    u'τ',
                    u'υύὺὐὑὓὔὕὖὗῦ',
                    u'θ',
                    u'χ',
                    u'ψ',
                    u'ξ',
                    u'ζ'
                    )
            
            
            
            def sortieren(t):
#                 try:
                    
                    result = []
                    
                    for i in t:
                        gefunden = 500
                        for j in gr2:
                            if i in j:
                                gefunden = gr2.index(j)
                        
                        result.append(gefunden)
                         
                        
                    return  result       
                        
                        
                    
                    #return [gr.index(c) for c in t]
#                 except:
#                     return [500 for c in t]
            
            woerter = sorted(posDict1.items(), key= lambda elem : sortieren(elem[1][0]))

            if len(woerter) == 0:
                Popup(self.mb, 'warning').text = LANG.KEINE_UEBEREINSTIMMUNGEN
                return
            
            self.oeffne_calc()
            
            sheet = self.calc.Sheets.getByIndex(0)    
            
            self.dict_eintraege = {}    
            
            for i in range(len(woerter)):
                cell = sheet.getCellByPosition(0, i)
                cell.setString(woerter[i][1][0])
                
                zaehlung =woerter[i][1][1]
                para = woerter[i][1][1].split('.')[0]
                zeile = woerter[i][1][2]
                 
                cell = sheet.getCellByPosition(1, i)
                cell.setString('{0}.{1}'.format(para,zeile))
                
                self.dict_eintraege.update({ i: (woerter[i][1][0], para, zeile, zaehlung) })
            
            
            path = os.path.join(self.speicherordner,'list_of_words.odt')
             
            self.calc.storeToURL(uno.systemPathToFileUrl(path),())    
            self.calc.close(False)
            
            
            
            ###### ERZEUGE HTMLS  ##################
            
            htmls = Erzeuge_Htmls(self.mb)
            htmls.greek2latex(self.dict_eintraege,self.speicherordner,self.pfad_quelldatei,self.pfad_translation)      
                     
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
            
            zeile = 1

            for l in lines:                
            
                if 'XXXXX' in l:
                    zaehlung = l.split('XXXXX')[0]
                    para,z = zaehlung.split('.')
                    zeile = int(z)
                    
                else:
                    
                    woerter = l.split()
                    woerter2 = l.split()
                    
                    if len(woerter) == 0:
                        w = posDict_O[zaehler-1][0]+'\n'
                        posDict_O.update({zaehler-1:(w,zaehlung,zeile)})
                        continue
                                        
                    woerter2[-1] = woerter2[-1]+'\n'
                    
                    # getrennte Woerter wieder zusammenfuegen
                    if getrenntesWort != '':
                        woerter[0] = getrenntesWort + woerter[0]
                        woerter2[0] = getrenntesWort2 + woerter2[0]+'\n'+'XXXgetrenntXXX'
                        getrenntesWort = ''
                     
                    if woerter[-1][-1] == '-':
                        getrenntesWort = woerter[-1][:-1]
                        getrenntesWort2 = woerter2[-1][:-1]
                        del woerter[-1] 
                        del woerter2[-1]                       
                    

                    
                    # Suchwoerter zusammenfuegen
                    for u in range(len(woerter)):
                        
                        w = woerter[u]
                        w2 = woerter2[u]
                        
                        if 'XXXgetrenntXXX' in w2:
                            w2 = w2.replace('XXXgetrenntXXX','')
                            z = zeile-1
                        else:
                            z = zeile
                        
                        posDict_O.update({zaehler:(w2,zaehlung,z)})
                        w = w.lower()
                        w = regex.sub('', w)
                        posDict.update({zaehler:(w,zaehlung,z)})
            
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
                    
                    zeile += 1
                     
            
            return posDict,posDict_O, WoerterDict, WoerterListe
        except:
            log(inspect.stack,tb())
            return {},{},{},[],{}
     




from codecs import open as codecs_open
from traceback import format_exc as tb
import os
import uno
from string import punctuation as Punctuation



class Erzeuge_Htmls():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.leerzeile = u'<br>'
        self.umbruch = u'<br>'
        self.ausnahmen = []
        self.inhalt = []
        
        self.dir_path = None
        self.footnoteHtml = []
        self.fn_Anzahl = 0
        
    
     
    def greek2latex(self,dict_eintraege,path,quelldatei,pfad_translation):
        if self.mb.debug: log(inspect.stack)
        
        self.dir_path = path
        text_translation = self.oeffne_text(pfad_translation)
        
        try:
            self.erstelle_html_uebersetzung(text_translation)
            self.ooo.close(False)
            self.erstelle_quelltext_html(quelldatei)
            self.erstelle_nav(dict_eintraege)
            self.erstelle_index_html()
        except:
            log(inspect.stack,tb())
           
    
    def oeffne_text(self,pfad_translation):
        if self.mb.debug: log(inspect.stack)
         
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True

        url = uno.systemPathToFileUrl(pfad_translation)       
        self.ooo = self.mb.doc.CurrentController.Frame.loadComponentFromURL(url,'_blank',0,(prop,)) 
        
        return self.ooo.Text
    
    
    def erstelle_html_uebersetzung(self,text):
        if self.mb.debug: log(inspect.stack)
        
        self.fn_Anzahl = 0
        self.footnoteHtml = []
        
        try:

            self.inhalt = []
            self.inhalt.append(self.Praeambel())
            paras = []

            enum = text.createEnumeration()
            
            while enum.hasMoreElements():
                paras.append(enum.nextElement())
            
            StatusIndicator = self.mb.doc.CurrentController.Frame.createStatusIndicator()
            StatusIndicator.start('Export ',len(paras))
            
            
            self.inhalt.append('<div style="padding:3%;>"')
            
            x = 0
            for par in paras:
                 
                x += 1
                StatusIndicator.setValue(x)
                 
                if 'com.sun.star.text.TextContent' not in par.SupportedServiceNames:
                    continue
                 
                teste_kursiv = True
                teste_fett = True
                leerer_para = False
              
                enum2 = par.createEnumeration()
                portions = []
                 
                while enum2.hasMoreElements():
                    portions.append(enum2.nextElement())
                 
                # auf leeren Paragraph testen
                if len(portions) == 1:
                    if portions[0].String == '':
                        self.inhalt.append(u'<br>')
                        leerer_para = True
 
                for portion in portions:
                    open = 0
                    
 
                    # fussnote
                    if portion.Footnote != None:
                        self.fuege_fussnote_ein(portion.Footnote)
                        self.fn_Anzahl += 1
                        continue
                     
                    if portion.CharColor not in(-1,0):
                        col = '%x' %portion.CharColor
                        self.inhalt.append(u'<span style="color: #{};">'.format(col))
                        open += 1
                         
                    if portion.CharBackColor not in(-1,0):
                        col = '%x' %portion.CharBackColor
                        self.inhalt.append(u'<span style="background-color: #{};">'.format(col))
                        open += 1
                     
                    # kursiv
                    if teste_kursiv:
                        if portion.CharPosture.value == 'ITALIC':
                            self.inhalt.append(u'<span style="font-style: italic;">')
                            open += 1
                     
                    # fett
                    if teste_fett:
                        if portion.CharWeight == 150:
                            self.inhalt.append(u'<span style="font-weight: bold;">')
                            open += 1
                     
                     
                    # alle geoeffneten Klammern schliessen
                    self.inhalt.append(portion.String)
                    self.inhalt.append(u'</span>' * open)
 
                 
                 
                self.inhalt.append(self.leerzeile)
                if leerer_para:
                    self.inhalt.append(u'<br>') 
              
              
                 
            self.inhalt.append('</div>')
 
            self.inhalt = '\n'.join(self.inhalt) 
 
            self.zaehlung_suchen_und_anker_setzen()
                 
            t = '\n'.join(self.footnoteHtml)
            self.inhalt = self.inhalt + t
                 
            self.inhalt = self.inhalt + self.ende()
            t = ''.join(self.inhalt)
            
            pfad = os.path.join(self.dir_path,'translation.html')
            self.speicher(t, 'w', pfad)
             
        except:
            log(inspect.stack,tb())

   
        StatusIndicator.end()
        
    
    
    def zaehlung_suchen_und_anker_setzen(self):
        
        import re
        suche = re.findall('\[[0-9]{1,3},[0-9]{1,2}\]',self.inhalt)
        
        self.sprungziele = []
        
        for s in suche:
            new_term = re.sub('[\[\]]','',s)
            new_term2 = '<a name="{0}"><b>{1}</b></a>'.format(new_term,s)
            
            self.inhalt = self.inhalt.replace(s,new_term2)
            self.sprungziele.append(new_term)
        
    
    
    def erstelle_quelltext_html(self,quelldatei):
        
        
        self.quelltext = []
        
        self.quelltext.append(self.Praeambel())
        self.quelltext.append('<div style="padding:3%">')
        
        pfad = os.path.join(self.dir_path,quelldatei)
        with codecs_open( pfad, 'r',"utf-8") as file:
            lines = file.readlines()
        
        for l in lines:
            
            if 'XXXXX' in l:
                t = l.replace('XXXXX','').replace('\n','')
                sprung = '<a name="{0}">{0}</a>'.format(t)
                self.quelltext.append(sprung)
                continue
            
            self.quelltext.append(l.replace('<','&lt;').replace('>','&gt;'))
        
        self.quelltext.append('</div>')   
        self.quelltext.append(self.ende())
        
        quelle = '<br>\n'.join(self.quelltext)
        
        pfad = os.path.join(self.dir_path,'source.html')
        self.speicher(quelle, 'w', pfad)
        

    
    def erstelle_nav(self,dict_eintraege):    
        if self.mb.debug: log(inspect.stack)
        
        self.navigation = []
        self.navigation.append(self.nav_Praeambel())        
        
        text = u'<a style="margin-left:.5em" href="javascript:void(0)"onClick="javascript:Sprung1(\'{0}\');javascript:Sprung2(\'{1}\')">' \
        '{2}</a><span style="position:absolute;left:12em">{3}</span><br><hr noshade size="1" style="margin: 2px">'
    
        
        for k in sorted(dict_eintraege):

            eintrag = dict_eintraege[k]
            sprung = eintrag[3]
            zaehlung = '{0}.{1}'.format(eintrag[1],eintrag[2])
            wort = eintrag[0]

            t = text.format(sprung,sprung.replace('.',','),wort,zaehlung.encode('utf-8'))
            self.navigation.append(t)
            
        
        self.navigation.append('<div style="height:50px;width:100%"></div>')
        self.navigation.append('     </div>')  
        self.navigation.append(self.ende())  
        
        t = '\n'.join(self.navigation)
        
        pfad = os.path.join(self.dir_path,'navigation.html')
        self.speicher(t, 'w', pfad)
     
    
    
    def fuege_fussnote_ein(self,footnote):
        if self.mb.debug: log(inspect.stack)
        
        fn = self.erstelle_fussnote(footnote)
        
        self.inhalt.append(u'<sup><a href="#fn{0}" id="ref{0}">{0}</a></sup>'.format(self.fn_Anzahl))
        fussnotentext = u'<p><sup id="fn{0}">{0}. {1}<a href="#ref{0}" title="Jump back">back</a></sup></p>'.format(self.fn_Anzahl,fn)
        self.footnoteHtml.append(fussnotentext)
        
     
            
    def speicher(self,inhalt,mode,pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if not os.path.exists(os.path.dirname(pfad)):
                os.makedirs(os.path.dirname(pfad))            
                
            with codecs_open( pfad, mode,"utf-8") as file:
                file.write(inhalt)
        except:
            log(inspect.stack,tb())
            


    def erstelle_fussnote(self,fn):
        if self.mb.debug: log(inspect.stack)
        
        try:
            text = fn.Text
            
            paras = []
            fussnote = []
            
            enum = text.createEnumeration()
            
            while enum.hasMoreElements():
                paras.append(enum.nextElement())
            

            for par in paras:
                
                enum2 = par.createEnumeration()
                portions = []
                
                while enum2.hasMoreElements():
                    portions.append(enum2.nextElement())
                
                # auf leeren Paragraph testen
                if len(portions) == 1:
                    if portions[0].String == '':
                        fussnote.append(u'<br>')
                

                for portion in portions:
                    open = 0
                    
                    # kursiv
                    if portion.CharPosture.value == 'ITALIC':
                        fussnote.append(u'<span style="font-style: italic;">')
                        open += 1
                    
                    # fett
                    if portion.CharWeight == 150:
                        fussnote.append(u'<span style="font-weight: bold;">')
                        open += 1
                    
                    try:
                        
                        if portion.String != '':
                            fussnote.append(portion.String)  
                    except:
                        #print("FEHLER: ",portion.String) 
                        pass                 
                    fussnote.append(u'</span>' * open)
                
                if par != paras[-1]:    
                    fussnote.append(self.leerzeile)
                
            return u''.join(fussnote)
            
        except:
            log(inspect.stack,tb())
            


          
#############################################################
            
    def Praeambel(self):
        if self.mb.debug: log(inspect.stack)
        
        Praeambel = u'''
<!DOCTYPE html>
<html>

<head>
  <title></title>

  <meta charset="UTF-8">
  
  <script type="text/javascript">
  
  </script>



</head>



<body> 
'''
        
        return Praeambel
        
        

################################################################       
             
    def ende(self):
        if self.mb.debug: log(inspect.stack)
        
        ende = u'''
        </body>
</html>'''
        
        
        return ende    
        
 
 
#####################################################################
        
        
    def nav_Praeambel(self):
        
        prae= '''
<!DOCTYPE html>
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
            loc = "source.html#" + x;
            window.parent.document.getElementById('txt1').src = loc;
            // console.log(window,document)
            /*window.location.href = "#NavStop"; */
            return
        }
        
        function Sprung2(x) {
            loc = "translation.html#" + x;
            window.parent.document.getElementById('txt2').src = loc;
            console.log(window.parent.document.getElementById('txt2'))
            window.location.href = "#NavStop"; 
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
    <div style="margin-left:0.1em;font-size:11pt;">
'''
        return prae 
      
      
###############################################################################      
      
      
    def erstelle_index_html(self):  
        html = '''
<!DOCTYPE html>

<head>
    <title>Indexation</title>

    <meta charset="utf-8">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="keywords" content="">
     <meta charset="UTF-8">

    <style type="text/css" media="screen">
        a { text-decoration: none; color: #FF0000; }
        p { font-size:10pt; }
        b { font-size: 9; }
        .col { background-color: #FFFF00; }
    </style>
    
        
    <script type="text/javascript">
                 
        function getDocHeight(doc) {
            doc = doc || document;
            var body = doc.body, html = doc.documentElement;
            var height = Math.max( body.scrollHeight, body.offsetHeight, html.clientHeight, html.scrollHeight, html.offsetHeight );
            return height;
        }
             
    </script>

</head>

<html>
    <body  style="min-width: 1000px;margin-top: -10px; background-color: #000080;">
    
        <noscript>
            <p>This site uses JavaScript. You must allow JavaScript in your browser.</p>
        </noscript>

        <div style="background-color: #E9E7E3;height: 90px;width: 100%;border-radius: 10px;"> 
            <h3 style="text-align: center;">Indexation</h3>
            <!-- <p style="text-align: center;">Fundstellen: 5  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;   Versionen Text 1: 3 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;     Versionen Text 2: 2 </p> --> 
            
            <div style="float: left;
                        width: 15%;
                        text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;
                        border-left:none">
                <b>Words</b>
            </div>

            <div style="float: right;
                        width: 50%;
                        text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;
                        border-right: none">
                <b>Translation</b>
            </div>

            <div style="text-align: center;
                        background-color: #FAFAE7;
                        border: solid;
                        border-color:#000080;">
                <b>Source</b>
            </div>
    
            <iframe
                id="navigation"
                src="navigation.html"
                
                
                style="float: left;
                    width:15%;

                    background-color: #FAFAE7;
                    overflow:hidden"
                frameborder="0">
            </iframe>
    
            <iframe
                id="txt2"
                src="translation.html"
                width="50%"
                height="500"
                style="float: right;
                    background-color: #FAFAE7;
                    overflow:hidden  "
                frameborder="0">
            </iframe>
    
            <iframe
                id="txt1"
                src="source.html"
                width="35%"
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
                window.height = height +20
    
                frame1.height = height-105
                frame2.height = height-105
                frame3.height = height-105

            </script>
            
        </div>
        
    </body>
</html>
        '''
        
        pfad = os.path.join(self.dir_path,'index.html')
        self.speicher(html, 'w', pfad)


         
