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
        self.leerzeile = '\n\n'
        self.umbruch = '\n'
        self.ausnahmen = []
        self.inhalt = []
        self.blocks = unicode_blocks
        
        self.path = None
        self.dateiname = None
        
        
    def greek2latex(self,text,path,dateiname):
        if self.mb.debug: log(inspect.stack)
        
        try:
            self.path = path
            self.dateiname = dateiname
            
            self.inhalt = []
            self.Praeambel()
            paras = []
            

            enum = text.createEnumeration()
            
            while enum.hasMoreElements():
                paras.append(enum.nextElement())
            
            StatusIndicator = self.mb.doc.CurrentController.Frame.createStatusIndicator()
            StatusIndicator.start('Export ' + os.path.split(path)[1],len(paras))

            x = 0
            for par in paras:
                
                x += 1
                StatusIndicator.setValue(x)
                
                if 'com.sun.star.text.TextContent' not in par.SupportedServiceNames:
                    continue
                
                teste_kursiv = True
                teste_fett = True
                leerer_para = False

                if 'Heading' in par.ParaStyleName:
                    self.set_heading(par)
                    teste_kursiv = False
                    teste_fett = False

                if 'Quotations' in par.ParaStyleName:
                    self.inhalt.append('\\begin{quote}')
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
                        leerer_para = True

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
                    
                    
                    # alle geoeffneten Klammern schliessen
                    self.inhalt.append(portion.String)
                    self.inhalt.append('}' * open)
                
                

                # schliesse Klammer fuer Ueberschrift
                if 'Heading' in par.ParaStyleName:
                    self.inhalt.append('}')
                if 'Quotations' in par.ParaStyleName:
                    self.inhalt.append('\\end{quote}')
                
                self.inhalt.append(self.leerzeile)
                if leerer_para:
                    self.inhalt.append('\\noindent\n') 
                
                
            #self.inhalt = self.verbotene_buchstaben_auswechseln(self.inhalt)
            #self.speicher(self.inhalt,'a') 
            
            self.ende()
            
            self.ausgabe(''.join(self.inhalt))
            
        except:
            log(inspect.stack,tb())
           
            
        StatusIndicator.end()
        
    
  
    
    def fuege_fussnote_ein(self,footnote):
        if self.mb.debug: log(inspect.stack)
        
        fn = self.erstelle_fussnote(footnote)

        self.inhalt.append('\\footnote{')
        self.inhalt.append(''.join(fn))
        self.inhalt.append('}')

    
    def set_heading(self,par):
        if self.mb.debug: log(inspect.stack)
        
        
        if par.ParaStyleName == 'Heading':
            self.inhalt.append('{')
            pass#self.inhalt.append('\\chapter{')
        elif par.ParaStyleName == 'Heading 1':
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
        if self.mb.debug: log(inspect.stack)
        
        try:
        
            if pfad == None:
                pfad = os.path.join(self.path,self.dateiname+'.tex')
            
            if not os.path.exists(os.path.dirname(pfad)):
                os.makedirs(os.path.dirname(pfad))
                
            
            content = ''.join(inhalt)
            
                
            with codecs_open( pfad, mode,"utf-8") as file:
                file.write(content)
        except:
            log(inspect.stack,tb())
    
    def verbotene_buchstaben_auswechseln(self,content):    
        if self.mb.debug: log(inspect.stack)
         
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
                 
                 
            pfad_a = os.path.join(self.path,'exchanged_letters.txt')
             
            a2 = 10
            b = 15
            c = 20
             
            with codecs_open( pfad_a, 'w',"utf-8") as file:
                    top = 'Symbol'.ljust(a2) + u'Amount'.ljust(b) + 'Unicode Number'.ljust(c)+ 'exchanged with:' + '\r\n'
                    file.write(top)
                     
            for aus in ausgewechselte:
                                 
                symbol = aus[0].ljust(a2) + aus[1].ljust(b) + aus[2].ljust(c) + aus[3].ljust(c) + '\r\n'
                with codecs_open( pfad_a, 'a',"utf-8") as file:
                    file.write(symbol)

            return content
        except:
            log(inspect.stack,tb())
            
            
    def Praeambel(self):
        if self.mb.debug: log(inspect.stack)
        
        Praeambel = ('\\documentclass[12pt,a4paper]{scrreprt}',
        '\\usepackage[polutonikogreek,ngerman]{babel}',
        '\\usepackage{ucs}',
        '\\usepackage[T1]{fontenc}',
        '\\usepackage[utf8x]{inputenc}',
        '\\newcommand*{\\fn}[1]{\\footnote{#1}}',
        '\\usepackage[paper=a4paper,left=20mm,right=20mm,top=20mm,bottom=20mm] {geometry}',
        '\\usepackage{makeidx}',
        '\\usepackage{index}',
        '\\usepackage{remreset}',
        '\\usepackage{graphicx}',
        '\\makeatletter',
        '\\@removefromreset{footnote}{chapter}',
        '\\makeatother',
        '     ',
        '%\setcounter{secnumdepth}{0}',
        '%\setcounter{tocdepth}{0}',
        '\\begin{document}',
        '\\shorthandoff{"}',
        '     ')
        
        prae = ''
        
        for p in Praeambel:
            prae = prae + p + '\r\n'
        
        self.inhalt.append(prae)
        #self.speicher(prae,'w')
             
    def ende(self):
        if self.mb.debug: log(inspect.stack)
        
        ende = ('     ',
        '     ',
        '\\end{document}',
        '     ')
        
        en = ''
        
        for p in ende:
            en = en + p + '\r\n'
            
        self.inhalt.append(en)
        
    
#     def speicher_ausnahmen(self,buchst):
#         if self.mb.debug: log(inspect.stack)
#         
#         if buchst not in self.ausnahmen:
#             
#             self.ausnahmen.append(buchst)
#         
#             b = buchst + '  ' + str(ord(buchst)) + ' \r\n'
#             
#             pfad = pfad = os.path.join(self.path,'potential_errors.txt')
#             self.speicher(b,'a',pfad)
#                 
#     
#     def berechne_Ausnahmen(self,buchstabe):
#         return
#         if self.mb.debug: log(inspect.stack)
#         
#         OB =  ord(buchstabe)
#         
#         is_greek = ((879 < ord(buchstabe) < 1023) or (7935 < ord(buchstabe) < 8191))
#         
#         if not is_greek:
#             if 256 < OB:
#                 self.speicher_ausnahmen(buchstabe)
                
                
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
                        fussnote.append('\\bigskip')
                

                for portion in portions:
                    open = 0
                    
                    # kursiv
                    if portion.CharPosture.value == 'ITALIC':
                        fussnote.append('\\emph{')
                        open += 1
                    
                    # fett
                    if portion.CharWeight == 150:
                        fussnote.append('\\textbf{')
                        open += 1

                    fussnote.append(portion.String)                    
                    fussnote.append('}' * open)
                
                if par != paras[-1]:    
                    fussnote.append(self.leerzeile)
                
            return fussnote
            
        except:
            log(inspect.stack,tb())

    def ausgabe(self,text):
        
        try:
            
            blocks_sortiert = sorted(int(a) for a in list(self.blocks))
                
            buchstaben_set = set(text)
            buchstaben_liste = list(text)
            
    
            codepoints = {}
            codepoints_b = {}
            for a in buchstaben_set:
                codepoints.update({ord(a):a})
                codepoints_b.update({a:ord(a)})
                 
            codepoints_sortiert = sorted(list(codepoints))
            
            verwendete_Buchstaben = []
            for s in codepoints_sortiert:
                x = bisect_right(blocks_sortiert,s)
                unr = blocks_sortiert[x-1]
                verwendete_Buchstaben.append((unr,self.blocks[str(unr)],unichr(s)))
            
            
            dict_gr_BS = {}
            dict_auszuschliessende = {}
             
            for z in verwendete_Buchstaben:
                if u'Greek' in z[1]:
                    dict_gr_BS.update({z[2]:(z[1],z[0])})
                elif z[0] > 256:
                    dict_auszuschliessende.update({z[2]:(z[1],z[0])})
                
            
            verwendete = [u'A searchable unicode table can be found at: http://unicode-table.com\n',
                          u'An editor which is able to view all unicode signs is for example: Notepad++ \n',
                          u'\n',
                          u'%-5s %-8s %-30s %-20s \n' %(u'Sign',u'Dec',u'Unicode Block',u'Name'),
                          u'\n',]
            
            zu_tauschende = {}
            zu_t = []
            for vw in verwendete_Buchstaben:
                v1 = vw[1]
                v2 = vw[2]
                v3 = codepoints_b[v2]
                v4 = unicode_name(v2,'not defined, search in unicode table')
                
                if v1 == 'Control Character':
                    v2 = ' '
                    
                if v1 in ["Spacing Modifier Letters","Combining Diacritical Marks"]:
                    zu_tauschende.update({v2:v3})
                    zu_t.append(v2)
                    
                ausg = '%-5s %-8s %-30s %-20s \n' %(v2,v3,v1,v4) 
                verwendete.append(ausg)
            
 
            
            neuer_text = []
            greek = False
    
            pu = Punctuation + ' '
            pu = pu.replace('{','').replace('}','').replace('\\','')
            
            for b in buchstaben_liste:
                if b in zu_t:
                    b = 'XXX'+str(zu_tauschende[b])+'XXX'
                if b in verbotene_Buchstaben:
                    b = verbotene_Buchstaben[b]
                elif b == '_':
                    b = ' '
                if b in pu:
                    neuer_text.append(b)
                    continue
                    
                elif b not in dict_gr_BS:
                    if greek == True:
                        neuer_text.append('}')
                        greek = False
                    neuer_text.append(b)
                else:
                    if greek == False:
                        greek = True
                        neuer_text.append('\\textgreek{')
                        neuer_text.append(b)
                    else:
                        neuer_text.append(b)
                        

            pfad1 = self.path+'.tex'
            pfad2 = self.path +'_used_letters.txt'
            
            
            text_ver = ''.join(verwendete)
            self.speicher(text_ver,'w',pfad2)
            
            
            self.speicher(neuer_text,'w',pfad1)
        except:
            log(inspect.stack,tb())
        
        
        

unicode_blocks = {
 "0": "Control Character",
 "32": "Basic Latin",
 "128": "Latin-1 Supplement",
 "256": "Latin Extended-A",
 "384": "Latin Extended-B",
 "592": "IPA Extensions",
 "688": "Spacing Modifier Letters",
 "768": "Combining Diacritical Marks",
 "880": "Greek and Coptic",
 "1024": "Cyrillic",
 "1280": "Cyrillic Supplement",
 "1328": "Armenian",
 "1424": "Hebrew",
 "1536": "Arabic",
 "1792": "Syriac",
 "1872": "Arabic Supplement",
 "1920": "Thaana",
 "1984": "NKo",
 "2048": "Samaritan",
 "2112": "Mandaic",
 "2208": "Arabic Extended-A",
 "2304": "Devanagari",
 "2432": "Bengali",
 "2560": "Gurmukhi",
 "2688": "Gujarati",
 "2816": "Oriya",
 "2944": "Tamil",
 "3072": "Telugu",
 "3200": "Kannada",
 "3328": "Malayalam",
 "3456": "Sinhala",
 "3584": "Thai",
 "3712": "Lao",
 "3840": "Tibetan",
 "4096": "Myanmar",
 "4256": "Georgian",
 "4352": "Hangul Jamo",
 "4608": "Ethiopic",
 "4992": "Ethiopic Supplement",
 "5024": "Cherokee",
 "5120": "Unified Canadian Aboriginal Syllabics",
 "5760": "Ogham",
 "5792": "Runic",
 "5888": "Tagalog",
 "5920": "Hanunoo",
 "5952": "Buhid",
 "5984": "Tagbanwa",
 "6016": "Khmer",
 "6144": "Mongolian",
 "6320": "Unified Canadian Aboriginal Syllabics Extended",
 "6400": "Limbu",
 "6480": "Tai Le",
 "6528": "New Tai Lue",
 "6624": "Khmer Symbols",
 "6656": "Buginese",
 "6688": "Tai Tham",
 "6832": "Combining Diacritical Marks Extended",
 "6912": "Balinese",
 "7040": "Sundanese",
 "7104": "Batak",
 "7168": "Lepcha",
 "7248": "Ol Chiki",
 "7360": "Sundanese Supplement",
 "7376": "Vedic Extensions",
 "7424": "Phonetic Extensions",
 "7552": "Phonetic Extensions Supplement",
 "7616": "Combining Diacritical Marks Supplement",
 "7680": "Latin Extended Additional",
 "7936": "Greek Extended",
 "8192": "General Punctuation",
 "8304": "Superscripts and Subscripts",
 "8352": "Currency Symbols",
 "8400": "Combining Diacritical Marks for Symbols",
 "8448": "Letterlike Symbols",
 "8528": "Number Forms",
 "8592": "Arrows",
 "8704": "Mathematical Operators",
 "8960": "Miscellaneous Technical",
 "9216": "Control Pictures",
 "9280": "Optical Character Recognition",
 "9312": "Enclosed Alphanumerics",
 "9472": "Box Drawing",
 "9600": "Block Elements",
 "9632": "Geometric Shapes",
 "9728": "Miscellaneous Symbols",
 "9984": "Dingbats",
 "10176": "Miscellaneous Mathematical Symbols-A",
 "10224": "Supplemental Arrows-A",
 "10240": "Braille Patterns",
 "10496": "Supplemental Arrows-B",
 "10624": "Miscellaneous Mathematical Symbols-B",
 "10752": "Supplemental Mathematical Operators",
 "11008": "Miscellaneous Symbols and Arrows",
 "11264": "Glagolitic",
 "11360": "Latin Extended-C",
 "11392": "Coptic",
 "11520": "Georgian Supplement",
 "11568": "Tifinagh",
 "11648": "Ethiopic Extended",
 "11744": "Cyrillic Extended-A",
 "11776": "Supplemental Punctuation",
 "11904": "CJK Radicals Supplement",
 "12032": "Kangxi Radicals",
 "12272": "Ideographic Description Characters",
 "12288": "CJK Symbols and Punctuation",
 "12352": "Hiragana",
 "12448": "Katakana",
 "12544": "Bopomofo",
 "12592": "Hangul Compatibility Jamo",
 "12688": "Kanbun",
 "12704": "Bopomofo Extended",
 "12736": "CJK Strokes",
 "12784": "Katakana Phonetic Extensions",
 "12800": "Enclosed CJK Letters and Months",
 "13056": "CJK Compatibility",
 "13312": "CJK Unified Ideographs Extension A",
 "19904": "Yijing Hexagram Symbols",
 "19968": "CJK Unified Ideographs",
 "40960": "Yi Syllables",
 "42128": "Yi Radicals",
 "42192": "Lisu",
 "42240": "Vai",
 "42560": "Cyrillic Extended-B",
 "42656": "Bamum",
 "42752": "Modifier Tone Letters",
 "42784": "Latin Extended-D",
 "43008": "Syloti Nagri",
 "43056": "Common Indic Number Forms",
 "43072": "Phags-pa",
 "43136": "Saurashtra",
 "43232": "Devanagari Extended",
 "43264": "Kayah Li",
 "43312": "Rejang",
 "43360": "Hangul Jamo Extended-A",
 "43392": "Javanese",
 "43488": "Myanmar Extended-B",
 "43520": "Cham",
 "43616": "Myanmar Extended-A",
 "43648": "Tai Viet",
 "43744": "Meetei Mayek Extensions",
 "43776": "Ethiopic Extended-A",
 "43824": "Latin Extended-E",
 "43968": "Meetei Mayek",
 "44032": "Hangul Syllables",
 "55216": "Hangul Jamo Extended-B",
 "55296": "High Surrogates",
 "56192": "High Private Use Surrogates",
 "56320": "Low Surrogates",
 "57344": "Private Use Area",
 "63744": "CJK Compatibility Ideographs",
 "64256": "Alphabetic Presentation Forms",
 "64336": "Arabic Presentation Forms-A",
 "65024": "Variation Selectors",
 "65040": "Vertical Forms",
 "65056": "Combining Half Marks",
 "65072": "CJK Compatibility Forms",
 "65104": "Small Form Variants",
 "65136": "Arabic Presentation Forms-B",
 "65280": "Halfwidth and Fullwidth Forms",
 "65520": "Specials",
 "65536": "Linear B Syllabary",
 "65664": "Linear B Ideograms",
 "65792": "Aegean Numbers",
 "65856": "Ancient Greek Numbers",
 "65936": "Ancient Symbols",
 "66000": "Phaistos Disc",
 "66176": "Lycian",
 "66208": "Carian",
 "66272": "Coptic Epact Numbers",
 "66304": "Old Italic",
 "66352": "Gothic",
 "66384": "Old Permic",
 "66432": "Ugaritic",
 "66464": "Old Persian",
 "66560": "Deseret",
 "66640": "Shavian",
 "66688": "Osmanya",
 "66816": "Elbasan",
 "66864": "Caucasian Albanian",
 "67072": "Linear A",
 "67584": "Cypriot Syllabary",
 "67648": "Imperial Aramaic",
 "67680": "Palmyrene",
 "67712": "Nabataean",
 "67840": "Phoenician",
 "67872": "Lydian",
 "67968": "Meroitic Hieroglyphs",
 "68000": "Meroitic Cursive",
 "68096": "Kharoshthi",
 "68192": "Old South Arabian",
 "68224": "Old North Arabian",
 "68288": "Manichaean",
 "68352": "Avestan",
 "68416": "Inscriptional Parthian",
 "68448": "Inscriptional Pahlavi",
 "68480": "Psalter Pahlavi",
 "68608": "Old Turkic",
 "69216": "Rumi Numeral Symbols",
 "69632": "Brahmi",
 "69760": "Kaithi",
 "69840": "Sora Sompeng",
 "69888": "Chakma",
 "69968": "Mahajani",
 "70016": "Sharada",
 "70112": "Sinhala Archaic Numbers",
 "70144": "Khojki",
 "70320": "Khudawadi",
 "70400": "Grantha",
 "70784": "Tirhuta",
 "71040": "Siddham",
 "71168": "Modi",
 "71296": "Takri",
 "71840": "Warang Citi",
 "72384": "Pau Cin Hau",
 "73728": "Cuneiform",
 "74752": "Cuneiform Numbers and Punctuation",
 "77824": "Egyptian Hieroglyphs",
 "92160": "Bamum Supplement",
 "92736": "Mro",
 "92880": "Bassa Vah",
 "92928": "Pahawh Hmong",
 "93952": "Miao",
 "110592": "Kana Supplement",
 "113664": "Duployan",
 "113824": "Shorthand Format Controls",
 "118784": "Byzantine Musical Symbols",
 "119040": "Musical Symbols",
 "119296": "Ancient Greek Musical Notation",
 "119552": "Tai Xuan Jing Symbols",
 "119648": "Counting Rod Numerals",
 "119808": "Mathematical Alphanumeric Symbols",
 "124928": "Mende Kikakui",
 "126464": "Arabic Mathematical Alphabetic Symbols",
 "126976": "Mahjong Tiles",
 "127024": "Domino Tiles",
 "127136": "Playing Cards",
 "127232": "Enclosed Alphanumeric Supplement",
 "127488": "Enclosed Ideographic Supplement",
 "127744": "Miscellaneous Symbols and Pictographs",
 "128512": "Emoticons",
 "128592": "Ornamental Dingbats",
 "128640": "Transport and Map Symbols",
 "128768": "Alchemical Symbols",
 "128896": "Geometric Shapes Extended",
 "129024": "Supplemental Arrows-C",
 "131072": "CJK Unified Ideographs Extension B",
 "173824": "CJK Unified Ideographs Extension C",
 "177984": "CJK Unified Ideographs Extension D",
 "194560": "CJK Compatibility Ideographs Supplement",
 "917504": "Tags",
 "917760": "Variation Selectors Supplement",
 "983040": "Supplementary Private Use Area-A",
 "1048576": "Supplementary Private Use Area-B"}
          
