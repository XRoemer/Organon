# -*- coding: utf-8 -*-

from codecs import open as codecs_open
from traceback import format_exc as tb
import os
import uno
from string import punctuation as Punctuation
from unicodedata import name as unicode_name
from bisect import bisect_right

verbotene_Buchstaben = {#u"\u0313":u"\u2018", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                        #u"\u0312":'',
                        # u"\u02BB":u"\u2018", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                         #u"\u02BC":u"\u2019", # RIGHT SINGLE QUOTATION MARK
                        #u"\u0315":'',
                        #u"\u0308":'',
                        #u"\u2028":'',
                         #u"\u03F2":'',
                         u"\u201B":u"\u2019", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                        #u"\u2012":''
                         }




class ExportToLatex():
    
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
        
        
    def greek2latex(self,text,path,dateiname):
        if self.mb.debug: log(inspect.stack)
        
        self.dir_path = os.path.split(path)[0]
        quelldatei = 'galen_529_533.txt'
        
        try:
            self.erstelle_html_uebersetzung(text)
            self.erstelle_quelltext_html(quelldatei)
            self.erstelle_nav()
            self.erstelle_index_html()
        except:
            log(inspect.stack,tb())
           

    
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
            new_term2 = '<a name="{0}">{1}</a>'.format(new_term,s)
            
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
        

    
    def erstelle_nav(self):    
        
        import json
        
        self.navigation = []
        self.navigation.append(self.nav_Praeambel())
        
        pfad = os.path.join(self.dir_path,'navi.json')
        
        with codecs_open( pfad, 'r',"utf-8") as file:
            nav_dict = json.loads(file.read())
        
        
        text = u'<a style="margin-left:.5em" href="javascript:void(0)"onClick="javascript:Sprung1(\'{0}\');javascript:Sprung2(\'{1}\')">' \
        '{2}</a><span style="position:absolute;left:12em">{3}</span><br><hr noshade size="1" style="margin: 2px">'
        
        for k in sorted(nav_dict):
            eintrag = nav_dict[k]
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
                        print("FEHLER: ",portion.String)                  
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
            loc = "galen_529_533.html#" + x;
            window.parent.document.getElementById('txt1').src = loc;
            // console.log(window,document)
            /*window.location.href = "#NavStop"; */
            return
        }
        
        function Sprung2(x) {
            loc = "tlg.html#" + x;
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


         
