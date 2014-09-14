# -*- coding: utf-8 -*-

from codecs import open as codecs_open
from traceback import format_exc as tb
import os
import uno

verbotene_Buchstaben = {u"\u0313":u"\u2018", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                         u"\u0312":'',
                         u"\u02BB":u"\u2018", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                         u"\u02BC":u"\u2019", # RIGHT SINGLE QUOTATION MARK
                         u"\u0315":'',
                         u"\u0308":'',
                         u"\u2028":'',
                         u"\u03F2":'',
                         u"\u201B":u"\u2018", # SINGLE HIGH-REVERSED-9 QUOTATION MARK
                         u"\u2012":''}

class ExportToLatex():
    
    def __init__(self):
        
        self.leerzeile = '\r\n\r\n'
        self.umbruch = '\r\n'
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.ausnahmen = []
        
        self.path = self.pfad_zum_desktop_berechnen()
        
        # Datei leeren
        pfad = os.path.join(self.path,'eventuelle_Fehler.txt')
        self.speicher('','w',pfad)
    
    def greek2latex(self):
        
        try:
            
            self.schreibe_Praeambel()
            
            StatusIndicator = self.doc.CurrentController.Frame.createStatusIndicator()

            text = self.doc.Text
            
            paras = []
            self.inhalt = []
            
            enum = text.createEnumeration()
            
            while enum.hasMoreElements():
                paras.append(enum.nextElement())
                            
            StatusIndicator.start('LATEX',len(paras))
            
            x = 0
            for par in paras:
                
                x += 1
                StatusIndicator.setValue(x)
                
                if 'com.sun.star.text.TextContent' not in par.SupportedServiceNames:
                    continue
                
                teste_kursiv = True
                teste_fett = True

                if 'Heading' in par.ParaStyleName:
                    self.set_heading(par)
                    teste_kursiv = False
                    teste_fett = False

                
                enum2 = par.createEnumeration()
                portions = []
                
                while enum2.hasMoreElements():
                    portions.append(enum2.nextElement())
                
                # auf leeren Paragraph testen
                if len(portions) == 1:
                    if portions[0].String == '':
                        self.inhalt.append('\\bigskip')
                

                for portion in portions:
                    open = 0

                    # fussnote
                    if portion.Footnote != None:
                        self.fuege_fussnote_ein(portion.Footnote)
                        continue
                    
                    # kursiv
                    if teste_kursiv:
                        if portion.CharPosture.value == 'ITALIC':
                            self.inhalt.append('\\emph{')
                            open += 1
                    
                    # fett
                    if teste_fett:
                        if portion.CharWeight == 150:
                            self.inhalt.append('\\textbf{')
                            open += 1
                    
                    # griechisch
                    # haengt den text an
                    self.test_auf_griechisch(portion)
                    
                    # alle geoeffneten Klammern schliessen
                    self.inhalt.append('}' * open)
                
                

                # schliesse Klammer fuer Ueberschrift
                if 'Heading' in par.ParaStyleName:
                    self.inhalt.append('}')
                    
                self.inhalt.append(self.leerzeile)
                
            inhalt = self.verbotene_buchstaben_auswechseln(self.inhalt)
            self.speicher(inhalt,'a') 
            
            self.speicher_dok_ende()
            
    
            
        except:
            x = tb()
            print(x)
           
            
        StatusIndicator.end()
        
    
    def test_auf_griechisch(self,portion):
        
        text_len = len(portion.String)
        
        try:
            erster_b = portion.String[0]
        except:
            return
        
        self.is_greek = ((879 < ord(erster_b) < 1023) or (7935 < ord(erster_b) < 8191))
        
        
        inhalt = []
        text = []
        greek = []
        
        try:
            for i in range(text_len):
                buchstabe =  portion.String[i]
                
#                 if buchstabe in verbotene_Buchstaben:
#                     print('fehler',buchstabe)
#                     continue
                
                self.berechne_Ausnahmen(buchstabe)
                
                if buchstabe in (u' ' , u'.' , u',' , u';' , u':' , u'-', '(' , ')', '/'):
                    is_greek = self.is_greek
                else:
                    is_greek = ((879 < ord(buchstabe) < 1023) or (7935 < ord(buchstabe) < 8191))
                
                
                if is_greek != self.is_greek:

                    if is_greek:
                        inhalt.append((''.join(text),'txt'))
                        text = []
                    else:
                        inhalt.append((''.join(greek),'gr'))
                        greek = []
                        
                    self.is_greek = is_greek
                    
                if is_greek:
                    greek.append(buchstabe)
                else:
                    text.append(buchstabe)
                    

            if self.is_greek:
                inhalt.append((''.join(greek),'gr'))
            else:
                inhalt.append((''.join(text),'txt'))
                       
                    
            for inh in inhalt:
                if inh[1] == 'txt':
                    self.inhalt.append(inh[0])
                else:
                    self.inhalt.append('\\textgreek{' + inh[0] + '}')
                    
        except:
            print(tb())

        
    
    def fuege_fussnote_ein(self,footnote):
        
        fn = Fussnote(self)
        inhalt = fn.greek2latex(footnote)
        
        self.inhalt.append('\\footnote{')
        self.inhalt.append(''.join(inhalt))
        self.inhalt.append('}')

    
    def set_heading(self,par):
        
        if par.ParaStyleName == 'Heading 1':
            self.inhalt.append('\\section{')
        elif par.ParaStyleName == 'Heading 2':
            self.inhalt.append('\\subsection{')
        elif par.ParaStyleName == 'Heading 3':
            self.inhalt.append('\\subsubsection{')
        elif par.ParaStyleName == 'Heading 4':
            self.inhalt.append('\\paragraph{')
        elif par.ParaStyleName == 'Heading 5':
            self.inhalt.append('\\subparagraph{')
        elif par.ParaStyleName == 'Heading 6':
            pass
        
            
    def speicher(self,inhalt,mode,pfad = None):
        
        if pfad == None:
            pfad = os.path.join(self.path,'latexttest.tex')
        
        content = ''.join(inhalt)
        
        with codecs_open( pfad, mode,"utf-8") as file:
            file.write(content)
    
    def verbotene_buchstaben_auswechseln(self,content):    
        try:
            ausgewechselte = []
            content = ''.join(content)
            
            for b in verbotene_Buchstaben:
                anz = content.count(b)
                
                if anz > 0:
                    
                     
                    if verbotene_Buchstaben[b] == '':
                        tausch = 'XXX %s XXX'%anz
                    else:
                        tausch = verbotene_Buchstaben[b]
                    content = content.replace(b,tausch)
                    
                    mitteil = b , str(anz) , b.encode("unicode_escape"),tausch
                    ausgewechselte.append(mitteil) 
                
                
            pfad_a = pfad = os.path.join(self.path,'ausgewechselte.txt')
            
            a2 = 10
            b = 15
            c = 20
            
            with codecs_open( pfad_a, 'w',"utf-8") as file:
                    top = 'Symbol'.ljust(a2) + u'Häufigkeit'.ljust(b) + 'Unicode Nummer'.ljust(c)+ 'ersetzt mit:' + '\r\n'
                    file.write(top)
                    
            for aus in ausgewechselte:
                                
                symbol = aus[0].ljust(a2) + aus[1].ljust(b) + aus[2].ljust(c) + aus[3].ljust(c) + '\r\n'
                with codecs_open( pfad_a, 'a',"utf-8") as file:
                    file.write(symbol)
            #pd()
            return content
        except:
            print(tb())   
            
            
    def schreibe_Praeambel(self):
        
        Praeambel = ('\\documentclass[12pt,a4paper]{scrreprt}',
        '\\usepackage[polutonikogreek,ngerman]{babel}',
        '\\usepackage{ucs}',
        '\\usepackage[T1]{fontenc}',
        '\\usepackage[utf8x]{inputenc}',
        '\\newcommand{\\gr}{\\foreignlanguage{polutonikogreek}}',
        '\\newcommand*{\\fn}[1]{\\footnote{#1}}',
        '\\usepackage[paper=a4paper,left=20mm,right=20mm,top=20mm,bottom=20mm] {geometry}',
        '\\usepackage{makeidx}',
        '%\\usepackage[hyperref=true,indexing=cite,style=klassphilbib]{biblatex}',
        '\\usepackage{index}',
        '%\\usepackage[breaklinks]{hyperref}',
        '\\usepackage{remreset}',
        '\\usepackage{graphicx}',
        '\\makeatletter',
        '\\@removefromreset{footnote}{chapter}',
        '\\makeatother',
        '     ',
        '\\begin{document}',
        '     ')
        
        prae = ''
        
        for p in Praeambel:
            prae = prae + p + '\r\n'
            
        self.speicher(prae,'w')
             
    def speicher_dok_ende(self):
        
        ende = ('     ',
        '     ',
        '\\end{document}',
        '     ')
        
        en = ''
        
        for p in ende:
            en = en + p + '\r\n'
            
        self.speicher(en,'a')
        
    
    def speicher_ausnahmen(self,buchst):
        
        if buchst not in self.ausnahmen:
            
            self.ausnahmen.append(buchst)
        
            b = buchst + '  ' + str(ord(buchst)) + ' \r\n'
            
            pfad = pfad = os.path.join(self.path,'eventuelle_Fehler.txt')
            self.speicher(b,'a',pfad)
                
    
    def berechne_Ausnahmen(self,buchstabe):
        
        OB =  ord(buchstabe)
        
        is_greek = ((879 < ord(buchstabe) < 1023) or (7935 < ord(buchstabe) < 8191))
        
        if not is_greek:
            if 256 < OB:
                self.speicher_ausnahmen(buchstabe)
    
    def pfad_zum_desktop_berechnen(self):
                
        user = os.path.expanduser("~")
        path = os.path.join(user,'Desktop','Latex_Export')
        
        if not os.path.exists(path):
            os.makedirs(path)
    
        return path


class Fussnote():  
    
    def __init__(self,cl):
        
        self.leerzeile = '\r\n\r\n'
        self.open = 0
        self.inhalt = []
        self.cl =cl
    
    def greek2latex(self,fn):
        
        try:
            text = fn.Text
            
            paras = []
            self.inhalt = []
            
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
                        self.inhalt.append('\\bigskip')
                

                for portion in portions:
                    open = 0

                    # fussnote
                    if portion.Footnote != None:
                        self.fuege_fussnote_ein(portion)
                        continue
                    
                    # kursiv
                    if portion.CharPosture.value == 'ITALIC':
                        self.inhalt.append('\\emph{')
                        open += 1
                    
                    # fett
                    if portion.CharWeight == 150:
                        self.inhalt.append('\\textbf{')
                        open += 1
                    
                    # griechisch
                    # haengt den text an
                    self.test_auf_griechisch(portion)
                    
                    self.inhalt.append('}' * open)
                
                if par != paras[-1]:    
                    self.inhalt.append(self.leerzeile)
                
            return self.inhalt
            
        except:
            x = tb()
            print(x)
        
    
    def test_auf_griechisch(self,portion):
        
        text_len = len(portion.String)
        
        try:
            erster_b = portion.String[0]
        except:
            return
        
        self.is_greek = ((879 < ord(erster_b) < 1023) or (7935 < ord(erster_b) < 8191))
        
        
        inhalt = []
        text = []
        greek = []
        
        try:
            for i in range(text_len):
                buchstabe =  portion.String[i]
                
#                 if buchstabe in verbotene_Buchstaben:
#                     print('fehler',buchstabe)
#                     continue
                
                self.cl.berechne_Ausnahmen(buchstabe)
                
                if buchstabe in (u' ' , u'.' , u',' , u';' , u':' , u'-' , '(' , ')', '/'):
                    is_greek = self.is_greek
                else:
                    is_greek = ((879 < ord(buchstabe) < 1023) or (7935 < ord(buchstabe) < 8191))
                
                
                if is_greek != self.is_greek:

                    if is_greek:
                        inhalt.append((''.join(text),'txt'))
                        text = []
                    else:
                        inhalt.append((''.join(greek),'gr'))
                        greek = []
                        
                    self.is_greek = is_greek
                    
                if is_greek:
                    greek.append(buchstabe)
                else:
                    text.append(buchstabe)
                    

            if self.is_greek:
                inhalt.append((''.join(greek),'gr'))
            else:
                inhalt.append((''.join(text),'txt'))
                       
                    
            for inh in inhalt:
                if inh[1] == 'txt':
                    self.inhalt.append(inh[0])
                else:
                    self.inhalt.append('\\textgreek{' + inh[0] + '}')
                    
        except:
            print(tb())
            
            
            
def Latex_Export(args):
        
    print(args)
    #exL = ExportToLatex()
    ExportToLatex().greek2latex()
    
    
    
    
    
    