# -*- coding: utf-8 -*-


class ExportToHtml():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.leerzeile = u'\t\t\n<br>'
        self.inhalt = []
                
        
    def export2html(self,text,path,dateiname):
        if self.mb.debug: log(inspect.stack)
        
        self.path = path
        
        self.fn_Anzahl = 0
        self.footnoteHtml = []
        
        try:

            self.inhalt = []
            self.inhalt.append(self.Praeambel())
            self.inhalt.append('\t<div>')
            
            t = Text(self.mb)
            
            settings = self.mb.settings_exp['html_export']
            
            html,fussnoten = t.erstelle_html(text, SI=True,**settings)
            self.inhalt.extend(html)
            if len(fussnoten) > 0:
                self.inhalt.append('\n\n\t\t<!-- FOOTNOTES -->\n\n')
                self.inhalt.extend(fussnoten)
                
            self.inhalt.append('\n\t</div>')
 
            self.inhalt = ''.join(self.inhalt) 
 
                 
#             t = '\n'.join(self.footnoteHtml)
#             self.inhalt = self.inhalt + t
                 
            self.inhalt = self.inhalt + self.ende()
            t = ''.join(self.inhalt)
            
            self.speicher(t, 'w', self.path)
             
        except:
            log(inspect.stack,tb())

           
    def set_heading(self,par):
        if self.mb.debug: log(inspect.stack)
        
        
        if par.ParaStyleName == 'Heading':
            self.inhalt.append('\n\t\t<h1>')
            return '</h1>\n'
        
        elif 'Heading' in par.ParaStyleName:
            nr = par.ParaStyleName[-1]
            self.inhalt.append('\n\t\t<h{}>'.format(nr))
            return '</h{}>\n'.format(nr)
        
        return ''
       
         
    def speicher(self,inhalt,mode,pfad):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if not os.path.exists(os.path.dirname(pfad)):
                os.makedirs(os.path.dirname(pfad))            
                
            with codecs_open( pfad, mode,"utf-8") as file:
                file.write(inhalt)
        except:
            log(inspect.stack,tb())

            
    def Praeambel(self):
        if self.mb.debug: log(inspect.stack)
        
        Praeambel = u'''
<!DOCTYPE html>
<html>

<head>
  <title></title>
  <meta charset="UTF-8">
</head>

<body> 

'''
        
        return Praeambel
              
             
    def ende(self):
        if self.mb.debug: log(inspect.stack)
        
        ende = u'''

</body>
</html>'''
        
        return ende    
        





class Text(): 
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)

        
        self.mb = mb
        self.leerzeile = u'\t\t\n<br>'
        self.inhalt = []
        self.align = {
                      0:' style="text-align:left"',
                      1:' style="text-align:right"',
                      2:' style="text-align:justify"',
                      3:' style="text-align:center"',
                      4:' style="text-align:justify"'
                      }
        
        #self.nutze_fussnote = 
    
    
    def erstelle_html(self,text,
                      SI=True,
                      FETT=True,
                      KURSIV=True,
                      UEBERSCHRIFT=True,
                      FUSSNOTE=True,
                      FARBEN=True,
                      PARA=True,
                      AUSRICHTUNG=True,
                      LINKS=True,
                      ZITATE=False,
                      SCHRIFTGROESSE=False,
                      CSS=False,
                      SCHRIFTART=False
                      ): # die untersten 3 KWs werden noch nicht benutzt
        
        if self.mb.debug: log(inspect.stack)
        
        self.fn_Anzahl = 0
        self.fussnoteHtml = []
        
        try:

            self.inhalt = []
            paras = []

            enum = text.createEnumeration()
            
            while enum.hasMoreElements():
                paras.append(enum.nextElement())
            #pd()
            if SI:
                StatusIndicator = self.mb.doc.CurrentController.Frame.createStatusIndicator()
                StatusIndicator.start('Export ',len(paras))
            
                        
            x = 0
            for par in paras:
                
                x += 1
                if SI:
                    StatusIndicator.setValue(x)
                 
                if 'com.sun.star.text.TextContent' not in par.SupportedServiceNames:
                    continue
                 
                teste_kursiv = KURSIV
                teste_fett = FETT
                ist_para = PARA
                
                
                if UEBERSCHRIFT:
                    if 'Heading' in par.ParaStyleName:
                        heading_ende = self.set_heading(par,AUSRICHTUNG)
                        teste_kursiv = False
                        teste_fett = False
                        ist_para = False

#                 if 'Quotations' in par.ParaStyleName:
#                     self.inhalt.append('\\begin{quote}')
#                     teste_kursiv = False
#                     teste_fett = False
                
                
                enum2 = par.createEnumeration()
                portions = []
                 
                while enum2.hasMoreElements():
                    portions.append(enum2.nextElement())
                 
                # auf leeren Paragraph testen
                if len(portions) == 1:
                    if portions[0].String == '':
                        continue
                
                if AUSRICHTUNG and 'Heading' not in par.ParaStyleName:
                    ausr = par.ParaAdjust
                    self.inhalt.append(u'\n\t\t<p{0}>'.format(self.align[ausr]))
                elif ist_para:
                    self.inhalt.append(u'\n\t\t<p>')
                
                for portion in portions:
                    
                    nutze_farben = FARBEN
                    anfang = []
                    ende = []
 
                    if FUSSNOTE:
                        # fussnote
                        if portion.Footnote != None:
                            self.fuege_fussnote_ein(portion.Footnote)
                            self.fn_Anzahl += 1
                            continue
                    
                    if LINKS:
                        if portion.HyperLinkURL != '':
                            anfang.append(u'<a href="{0}" target="_blank">'.format(portion.HyperLinkURL))
                            ende.insert(0,u'</a>')
                            nutze_farben = False

                    if nutze_farben:
                        if portion.CharColor not in(-1,0):
                            col = '%x' %portion.CharColor
                            anfang.append(u'<span style="color: #{};">'.format(col))
                            ende.append(u'</span>')
                             
                        if portion.CharBackColor not in(-1,0):
                            col = '%x' %portion.CharBackColor
                            anfang.append(u'<span style="background-color: #{};">'.format(col))
                            ende.append(u'</span>')
                                                
                    # fett
                    if teste_fett:
                        if portion.CharWeight == 150:
                            anfang.append(u'<b>')
                            ende.insert(0,u'</b>')
                             
                    # kursiv
                    if teste_kursiv:
                        if portion.CharPosture.value == 'ITALIC':
                            anfang.append(u'<em>')
                            ende.insert(0,u'</em>')
                            
                    self.inhalt.extend(anfang)
                    self.inhalt.append(portion.String)
                    self.inhalt.extend(ende)
                
                if UEBERSCHRIFT:
                    # schliesse Klammer fuer Ueberschrift
                    if 'Heading' in par.ParaStyleName:
                        self.inhalt.append(heading_ende)
                        teste_kursiv = True
                        teste_fett = True
#                 if 'Quotations' in par.ParaStyleName:
#                     self.inhalt.append('\\end{quote}')
                    
                
                if ist_para:
                    self.inhalt.append(u'\n\t\t</p>\n')
             
        except:
            log(inspect.stack,tb())
            return [],[]

        if SI:
            StatusIndicator.end()
        
        return self.inhalt,self.fussnoteHtml
        
        
    def set_heading(self,par,AUSRICHTUNG):
        if self.mb.debug: log(inspect.stack)
        
        tex = ''
        
        if AUSRICHTUNG:
            ausr = par.ParaAdjust
            tex = self.align[ausr]
        
        if par.ParaStyleName == 'Heading':
            self.inhalt.append('\n\t\t<h1{0}>'.format(tex))
            return '</h1>\n'
        
        elif 'Heading' in par.ParaStyleName:
            nr = par.ParaStyleName[-1]
            self.inhalt.append('\n\t\t<h{0}{1}>'.format(nr,tex))
            return '</h{}>\n'.format(nr)
        
        return ''
    

    def fuege_fussnote_ein(self,fussnote):
        if self.mb.debug: log(inspect.stack)
        
        t = Text(self.mb)
        settings = copy.deepcopy(self.mb.settings_exp['html_export'])
        settings['AUSRICHTUNG'] = False
        text,fn = t.erstelle_html(fussnote,SI=False,PARA=False,**settings)
        
        self.inhalt.append(u'<sup><a href="#fn{0}" id="ref{0}">{0}</a></sup>'.format(self.fn_Anzahl+1))
        t2 = ''.join(text)
        fussnotentext = u'\n\t\t<sup id="fn{0}">{0}. {1} <a href="#ref{0}">back</a></sup><br>'.format(self.fn_Anzahl+1,t2)
        self.fussnoteHtml.append(fussnotentext)
        













