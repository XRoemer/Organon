# -*- coding: utf-8 -*-

from datetime import datetime
from zipfile import ZipFile
from math import modf as math_modf
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

class Querverweise():
    '''
    Organon verwaltet folgende Querverweise:
    - Sequenzen:
        - Tabellen
        - Textrahmen
        - Bildrahmen
        - Zeichnungsrahmen
    - Fuss- und Endnoten
    
    Referenzen auf Ueberschriften, Listen und Nutzerreferenzen werden belassen
    und auch durch Writer selbst aktualisiert.
    
    Wird ein Bereich via Teile_Text_Batch geteilt, werden alle Referenzen,
    auch Ueberschriften, Listen und Nutzerreferenzen, in den neuen Dateien umbenannt,
    damit sie auch auf die neuen Dateien verweisen.
    
    '''
    
    
    
    def __init__(self,mb):
                
        self.mb = mb  
        self.neue_bms = []
        
        # als: xml Name, OO Konst Wert, OO Konst Name   
        self.dict_ref_field_part = {
                        'page' : 0,                 # PAGE
                        'chapter' : 1,              # CHAPTER  ## Number
                        'text' : 2,                 # TEXT
                        'direction' : 3,            # UP_DOWN
                        '' : 4,                     # PAGE_DESC
                        'category-and-value' : 5,   # CATEGORY_AND_NUMBER
                        'caption' : 6,              # ONLY_CAPTION
                        'value' : 7,                # ONLY_SEQUENCE_NUMBER
                        'number' : 8,               # NUMBER
                        'number-no-superior' : 9,   # NUMBER_NO_CONTEXT
                        'number-all-superior' : 10, # NUMBER_FULL_CONTEXT
                        }
                    
        self.ref_field_part_dict = {v:k for k,v in self.dict_ref_field_part.items()}
        
        # als: xml Name, OO Konst Wert, OO Konst Name
        self.reference_field_source = {
                                 0 : 'reference',   # benutzer referenz  
                                 1 : 'sequence',    # Beschriftungsfelder bei Tabellen, Illustrationen, Text, Zeichnungen 
                                 2 : 'bookmark',    # Lesezeichen, Ueberschriften, Num. Absaetze
                                 3 : 'footnote',  
                                 4 : 'endnote', 
                                 }    
            
        self.prae = u'''<orga xmlns:office="orgax_nsp123"
xmlns:style="orgax_nsp123"
xmlns:text="orgax_nsp123"
xmlns:table="orgax_nsp123"
xmlns:draw="orgax_nsp123"
xmlns:fo="orgax_nsp123"
xmlns:xlink="orgax_nsp123"
xmlns:dc="orgax_nsp123"
xmlns:meta="orgax_nsp123"
xmlns:number="orgax_nsp123"
xmlns:svg="orgax_nsp123"
xmlns:chart="orgax_nsp123"
xmlns:dr3d="orgax_nsp123"
xmlns:math="orgax_nsp123"
xmlns:form="orgax_nsp123"
xmlns:script="orgax_nsp123"
xmlns:ooo="orgax_nsp123"
xmlns:ooow="orgax_nsp123"
xmlns:oooc="orgax_nsp123"
xmlns:dom="orgax_nsp123"
xmlns:xforms="orgax_nsp123"
xmlns:xsd="orgax_nsp123"
xmlns:xsi="orgax_nsp123"
xmlns:rpt="orgax_nsp123"
xmlns:of="orgax_nsp123"
xmlns:xhtml="orgax_nsp123"
xmlns:grddl="orgax_nsp123"
xmlns:officeooo="orgax_nsp123"
xmlns:tableooo="orgax_nsp123"
xmlns:drawooo="orgax_nsp123"
xmlns:calcext="orgax_nsp123"
xmlns:loext="orgax_nsp123"
xmlns:field="orgax_nsp123"
xmlns:formx="orgax_nsp123"
xmlns:css3t="orgax_nsp123">'''
                    
        self.post = u'''</orga>'''
        
    
    def get_root(self,txt):
        #import xml.etree.ElementTree as ElementTree
        try:
            if ':' in txt:
                txt = self.prae + txt + self.post
                t = txt.encode('utf-8')
                root = ElementTree.fromstring( t )
                
                for el in root.iter():
                    el.tag = el.tag.replace('{orgax_nsp123}','')
                    attribs = el.attrib
                    new_attr = { k.replace('{orgax_nsp123}','') : v for k,v in attribs.items()}
                    el.attrib = new_attr
                    
                return root.find('.//')
            
            else:
                return ElementTree.fromstring( txt.encode('utf-8') )
            
        except ElementTree.ParseError as e:
            #print('#',txt[e.position[1]-1:])
            log(inspect.stack,tb(),extras=txt[e.position[1]-1:])
#             with codecs_open(pfad_plain_txt , "w","utf-8") as f:
#                 f.write(t.decode('utf-8'))
            return None
        except:
            log(inspect.stack,tb())
    
    
    def get_ueberschrift(self, ordinal, text_range, is_frame=False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            ueberschriften = self.mb.class_Tools.get_ueberschriften()  
            if is_frame:
                
                try:
                    text_range = text_range.TextFrame.Anchor
                except:
                    text_range = text_range
                
            for i,u in enumerate(reversed(ueberschriften)):
                sec = u['para']
                try:
                    if self.mb.doc.Text.compareRegionStarts(sec,text_range) == 1:
                        return str(len(ueberschriften) - i)
                except:
                    pass
        except:
            log(inspect.stack,tb())
            
        return '???'
    
    
    def get_notes(self,doc):
        if self.mb.debug: log(inspect.stack)

        fns = [ doc.Footnotes.getByIndex(c) for c in range(doc.Footnotes.Count) ]
        self.footnotes = {f.ReferenceId : f for f in fns}

        ens = [ doc.Endnotes.getByIndex(c) for c in range(doc.Endnotes.Count) ]
        self.endnotes = {e.ReferenceId : e for e in ens}
    
    
    def get_text_fields_seq(self,doc):
        if self.mb.debug: log(inspect.stack)

        enum = doc.TextFields.createEnumeration()
        tfs = []
        while enum.hasMoreElements():
            tfs.append(enum.nextElement())

        ziel_felder = [t for t in tfs if 'com.sun.star.text.TextField.SetExpression' in t.SupportedServiceNames]
        
        self.ziel_felder_seq = {
                                'Table' : {},
                                'Drawing' : {},
                                'Text' : {},
                                'Illustration' : {}
                                }
        
        for z in ziel_felder:
            self.ziel_felder_seq[z.VariableName][z.SequenceValue] = z, 'com.sun.star.text.TextFrame' in z.Anchor.Text.SupportedServiceNames

 
    def fuege_querverweis_ein_import(self, tf, doc, bms):
        if self.mb.debug: log(inspect.stack)
        
        try:
            cursor = tf.Anchor

            source_name = tf.SourceName
            sequence_number = tf.SequenceNumber
            ref_field_source = tf.ReferenceFieldSource
            ref_field_part = tf.ReferenceFieldPart
            

            def get_ziel_feld_text(text_instance,cur):
                
                cur2 = text_instance.createTextCursorByRange(cur)
                cur3 = text_instance.createTextCursorByRange(cur)
                cur4 = text_instance.createTextCursorByRange(cur)
                nummer = cur.String
                cur2.gotoNextWord(False)
                cur2.gotoEndOfParagraph(True)
                beschriftung = cur2.String
                
                cur3.collapseToStart()
                cur3.gotoStartOfParagraph(True)
                ref_txt = cur3.String
                
                cur4.collapseToEnd()
                cur4.gotoNextWord(True)
                sep = cur4.String.strip()
                
                return ref_txt, nummer, beschriftung, sep
            
            
            def referenzfeld_anlegen(bm_name):
                # page, text, direction, page_desc
                # Es wird ein Referenzfeld gesetzt, das sich via OO Mechanismus
                # selbst updated (oder F9)
                ref = doc.createInstance("com.sun.star.text.textfield.GetReference")   
                ref.SourceName = bm_name
                ref.ReferenceFieldPart = ref_field_part
                ref.ReferenceFieldSource = 2 
                cursor.Text.insertTextContent(cursor,ref,False)
                ref.update()
                
                
            if self.reference_field_source[ref_field_source] == 'reference':
                return False   
            
            elif self.reference_field_source[ref_field_source] == 'bookmark':
                return False
            
            elif self.reference_field_source[ref_field_source] == 'sequence':
                
                ziel_feld, is_frame = self.ziel_felder_seq[source_name][sequence_number]
                
                if is_frame:
                    cur = ziel_feld.Anchor.Text.createTextCursorByRange(ziel_feld.Anchor)
                    bm_name = self.lesezeichen_im_zielfeld_anlegen_import(bms, doc, ziel_feld, is_frame)
                else:
                    cur = doc.Text.createTextCursorByRange(ziel_feld.Anchor)
                    bm_name = self.lesezeichen_im_zielfeld_anlegen_import(bms, doc, ziel_feld)     
                
                if self.ref_field_part_dict[ref_field_part] in ( 'page', 'text', 'direction', ''):
                    referenzfeld_anlegen(bm_name)
                    return True
                    
                else:
                    if self.ref_field_part_dict[ref_field_part] in ('category-and-value', 'caption', 'value'):
                        
                        text_instance = doc.Text
                        if is_frame:
                            text_instance = ziel_feld.Anchor.Text
                        
                        ref_txt, nummer, beschriftung, sep = get_ziel_feld_text(text_instance, cur)

                        if self.ref_field_part_dict[ref_field_part] == 'category-and-value':
                            text = ref_txt + nummer
                        elif self.ref_field_part_dict[ref_field_part] == 'caption':
                            text = beschriftung
                        else:
                            text = nummer
                          
                    elif self.ref_field_part_dict[ref_field_part] == 'chapter':
                        text = '???'
                        sep = '' 
                        
                    orga_ref = 'zzOrganonField_{0}_{1}_{2}_'.format(ref_field_source,ref_field_part, sep) + bm_name.replace('zzOrganonBM_', '')
                    
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)
                    return True
                    
                
            elif self.reference_field_source[ref_field_source] in ('footnote', 'endnote'): 
                
                if self.reference_field_source[ref_field_source] == 'footnote':
                    fn = self.footnotes[sequence_number]
                else:
                    fn = self.endnotes[sequence_number]

                bm_name = self.lesezeichen_im_zielfeld_anlegen_fussnote(fn,doc)
                
                if self.ref_field_part_dict[ref_field_part] in ('page', 'direction', ''):
                    referenzfeld_anlegen(bm_name)
                    
                elif self.ref_field_part_dict[ref_field_part] == 'text':
                    text = fn.Anchor.String
                    orga_ref = 'zzOrganonField_{0}_{1}_'.format(ref_field_source, ref_field_part) + bm_name.replace('zzOrganonBM_', '')
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)
                    
                else:
                    text = '???'
                    orga_ref = 'zzOrganonField_{0}_{1}_'.format(ref_field_source, ref_field_part) + bm_name.replace('zzOrganonBM_', '')
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)
                
                return True
                
        except:
            log(inspect.stack,tb())
            return False
    
    
    def lesezeichen_im_zielfeld_anlegen_import(self, bms, doc, ziel_feld, is_frame=False):   
        if self.mb.debug: log(inspect.stack) 

        try:
            if is_frame:
                cur = ziel_feld.Anchor.Text.createTextCursorByRange(ziel_feld.Anchor)
                instance = ziel_feld.Anchor
            else:
                cur = doc.Text.createTextCursorByRange(ziel_feld.Anchor)
                instance = doc

            cur2 = instance.Text.createTextCursorByRange(cur)
            cur2.gotoEndOfParagraph(False)
            cur2.gotoStartOfParagraph(True)
            
            bookmarks = doc.Bookmarks
            bms.extend([b.Name for b in self.neue_bms])
            
            bms = set(bms)
            
            for name in bms:
                bm = bookmarks.getByName(name)

                try:
                    if bm.Anchor.Text.compareRegionStarts(bm.Anchor,cur2) == 0:
                        
                        if 'zzOrganonBM_' in bm.Name:
                            return bm.Name
                except:
                    # wenn Textinstanzen von doc und einem frame verglichen
                    # werden, erzeugt das Fehler
                    pass
                        
            cur3 = instance.Text.createTextCursorByRange(cur)
            cur3.gotoEndOfParagraph(False)
            cur3.gotoStartOfParagraph(True)
                
            bm = doc.createInstance("com.sun.star.text.Bookmark")
            bm_name = 'zzOrganonBM_{0}'.format(self.time_stamp())
            
            while bm_name in doc.Bookmarks.ElementNames:
                bm_name = bm_name + '1'
            
            bm.Name = bm_name
            instance.Text.insertTextContent(cur3,bm,True)
            self.neue_bms.append(bm)
            
            return bm.Name
    
        except:
            log(inspect.stack,tb())

    
    def fuege_querverweis_ein(self,tf):
        if self.mb.debug: log(inspect.stack)
        
        try:
            doc = self.mb.doc          
            cursor = tf.Anchor

            source_name = tf.SourceName
            sequence_number = tf.SequenceNumber
            ref_field_source = tf.ReferenceFieldSource
            ref_field_part = tf.ReferenceFieldPart
            
            
            def get_ziel_feld_seq():
            
                enum = doc.TextFields.createEnumeration()
                tfs = []
                while enum.hasMoreElements():
                    tfs.append(enum.nextElement())
                    
                ziel_feld = [t for t in tfs if 'com.sun.star.text.TextField.SetExpression' 
                           in t.SupportedServiceNames and t.VariableName == source_name
                           and int(t.SequenceValue) == sequence_number][0]
                
                ordinal, is_frame = self.mb.class_Bereiche.get_ordinal(ziel_feld.Anchor, True)
                
                return ziel_feld, ordinal, is_frame
            
            
            def get_ziel_feld_text(text_instance,cur):
                
                cur2 = text_instance.createTextCursorByRange(cur)
                cur3 = text_instance.createTextCursorByRange(cur)
                cur4 = text_instance.createTextCursorByRange(cur)
                nummer = cur.String
                cur2.gotoNextWord(False)
                cur2.gotoEndOfParagraph(True)
                beschriftung = cur2.String
                
                cur3.collapseToStart()
                cur3.gotoStartOfParagraph(True)
                ref_txt = cur3.String
                
                cur4.collapseToEnd()
                cur4.gotoNextWord(True)
                sep = cur4.String.strip()
                
                return ref_txt, nummer, beschriftung, sep
                        
            def referenzfeld_anlegen(bm_name):
                # page, text, direction, page_desc
                # Es wird ein Referenzfeld gesetzt, das sich via OO Mechanismus
                # selbst updated (oder F9)
                ref = doc.createInstance("com.sun.star.text.textfield.GetReference")   
                ref.SourceName = bm_name
                ref.ReferenceFieldPart = ref_field_part
                ref.ReferenceFieldSource = 2 
                cursor.Text.insertTextContent(cursor,ref,False)
                ref.update()
            
            # Wenn der Benutzer eine Referenz einfuegt,
            # wird die von Organon nicht beachtet
            if self.reference_field_source[ref_field_source] == 'reference':
                return
            
            ordinal = self.mb.class_Bereiche.get_ordinal(cursor)
            
            if self.reference_field_source[ref_field_source] == 'bookmark':
                # nummerierte Absaetze (Listen) und Ueberschriften
                # Dateien, die das Ziel enthalten, muessen nur gespeichert werden
                
                # Die neue Referenz wurde als bookmark-ref von Writer angelegt
                bms = doc.Bookmarks
                bm = bms.getByName(source_name)
                
                ordinal2 = self.mb.class_Bereiche.get_ordinal(bm.Anchor)
                self.mb.class_Bereiche.datei_speichern(ordinal2)
                self.mb.class_Bereiche.datei_speichern(ordinal)
                return


            elif self.reference_field_source[ref_field_source] == 'sequence':
                
                ziel_feld, zf_ordinal, is_frame = get_ziel_feld_seq()
                
                if is_frame:
                    cur = ziel_feld.Anchor.Text.createTextCursorByRange(ziel_feld.Anchor)
                    self.mb.undo_mgr.undo()
                    bm_name = self.lesezeichen_im_zielfeld_anlegen(cur, zf_ordinal, ziel_feld)

                else:
                    cur = ziel_feld.Anchor.Text.createTextCursorByRange(ziel_feld.Anchor)
                    self.mb.undo_mgr.undo()
                    bm_name = self.lesezeichen_im_zielfeld_anlegen(cur, zf_ordinal)                    
                    
                    
                if self.ref_field_part_dict[ref_field_part] in ( 'page', 'text', 'direction', ''):
                    referenzfeld_anlegen(bm_name)
                    
                else:
                    if self.ref_field_part_dict[ref_field_part] in ('category-and-value', 'caption', 'value'):

                        text_instance = ziel_feld.Anchor.Text
                        
                        ref_txt, nummer, beschriftung, sep = get_ziel_feld_text(text_instance, cur)

                        if self.ref_field_part_dict[ref_field_part] == 'category-and-value':
                            text = ref_txt + nummer
                        elif self.ref_field_part_dict[ref_field_part] == 'caption':
                            text = beschriftung
                        else:
                            text = nummer
                          
                    elif self.ref_field_part_dict[ref_field_part] == 'chapter':
                        text = self.get_ueberschrift(ordinal, cur, is_frame) 
                        sep = '' 
                        
                    orga_ref = 'zzOrganonField_{0}_{1}_{2}_'.format(ref_field_source,ref_field_part, sep) + bm_name.replace('zzOrganonBM_', '')
                    
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)    
                
                
            elif self.reference_field_source[ref_field_source] in ('footnote', 'endnote'): 
                
                if self.reference_field_source[ref_field_source] == 'footnote':
                    notes = doc.Footnotes
                else:
                    notes = doc.Endnotes
                    
                fn = None
                for c in range(notes.Count):
                    f = notes.getByIndex(c)
                    if f.ReferenceId == sequence_number:
                        fn = f
                        break 
              
                zf_ordinal = self.mb.class_Bereiche.get_ordinal(fn.Anchor)
                self.mb.undo_mgr.undo()
                
                bm_name = self.lesezeichen_im_zielfeld_anlegen_fussnote(fn)
                
                if self.ref_field_part_dict[ref_field_part] in ('page', 'direction', ''):
                    referenzfeld_anlegen(bm_name)
                    
                elif self.ref_field_part_dict[ref_field_part] == 'text':
                    text = fn.Anchor.String
                    orga_ref = 'zzOrganonField_{0}_{1}_'.format(ref_field_source, ref_field_part) + bm_name.replace('zzOrganonBM_', '')
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)
                    
                else:
                    text = self.get_ueberschrift(ordinal, fn.Anchor)
                    orga_ref = 'zzOrganonField_{0}_{1}_'.format(ref_field_source, ref_field_part) + bm_name.replace('zzOrganonBM_', '')
                    self.benutzerfeld_anlegen(doc, cursor, orga_ref, text)
              
            if ordinal == zf_ordinal:
                self.mb.class_Bereiche.datei_speichern(ordinal) 
            else:
                self.mb.class_Bereiche.datei_speichern(ordinal)   
                self.mb.class_Bereiche.datei_speichern(zf_ordinal)
            
            self.mb.undo_mgr.reset()
        except:
            log(inspect.stack,tb())
    
    
    def get_lesezeichen_in_content_xml(self, ordinal='', pfad=None):
        if self.mb.debug: log(inspect.stack) 
        
        try:
            if not pfad:
                pfad = os.path.join(self.mb.pfade['odts'], ordinal + '.odt')
            
            zipped = ZipFile(pfad)
            text_xml = zipped.read('content.xml')
            zipped.close()
            
            text_part = text_xml.split(b'<office:text',1)[1]
            
            bms = [t.split(b'"/>')[0].decode('utf-8') for t in text_part.split(b'<text:bookmark-start text:name="')][1:]            
            return bms

        except:
            log(inspect.stack, tb())
            return []


    def lesezeichen_im_zielfeld_anlegen(self, cur, zf_ordinal, zielfeld=None, doc=None):   
        if self.mb.debug: log(inspect.stack) 
        
        try:
            if doc == None:
                doc = self.mb.doc
            
            if zielfeld == None:
                text_range = self.mb.doc   
            else:
                text_range = zielfeld.Anchor
                
            cur2 = cur.Text.createTextCursorByRange(cur)
            cur2.gotoEndOfParagraph(False)
            cur2.gotoStartOfParagraph(True)
            
            bookmarks = doc.Bookmarks
            bms = self.get_lesezeichen_in_content_xml(zf_ordinal)
            
            try:
                for name in bms:
                    bm = bookmarks.getByName(name)
                    try:
                        if bm.Anchor.Text.compareRegionStarts(bm.Anchor,cur2) == 0:
                            if 'zzOrganonBM_' in bm.Name:
                                return bm.Name
                    except:
                        # wenn Textinstanzen von doc und einem textfeld verglichen
                        # werden, erzeugt das Fehler
                        pass
            except:
                # for schleife erzeugt fehler, wenn ElementNames
                # eine leere byte sequenz sind.
                pass
            
            cur3 = cur.Text.createTextCursorByRange(cur)
            cur3.gotoEndOfParagraph(False)
            cur3.gotoStartOfParagraph(True)
                
            bm = doc.createInstance("com.sun.star.text.Bookmark")
            bm_name = 'zzOrganonBM_{0}'.format(self.time_stamp())
            bm.Name = bm_name
            text_range.Text.insertTextContent(cur3,bm,True)
            
            return bm.Name
    
        except:
            log(inspect.stack,tb())
            
    
    def lesezeichen_im_zielfeld_anlegen_fussnote(self, fn, doc=None):   
        if self.mb.debug: log(inspect.stack) 
        
        try:
            if doc == None:
                doc = self.mb.doc
                bms = self.mb.doc.Bookmarks
            
            else:
                bms = doc.Bookmarks
            
            cur = fn.Anchor.Text.createTextCursorByRange(fn.Anchor)
            
            try:
                for name in bms.ElementNames:
                    bm = bms.getByName(name)
                    try:
                        if doc.Text.compareRegionStarts(bm.Anchor,cur) == 0:
                            if 'zzOrganonBM_' in bm.Name:
                                return bm.Name
                    except:
                        # wenn Textinstanzen von doc und einem frame verglichen
                        # werden, erzeugt das Fehler
                        pass
            except:
                # for schleife erzeugt fehler, wenn ElementNames
                # eine leere byte sequenz sind.
                pass
                            
            bm = self.mb.doc.createInstance("com.sun.star.text.Bookmark")
            bm_name = 'zzOrganonBM_{0}'.format(self.time_stamp())
            bm.Name = bm_name
            cur.Text.insertTextContent(cur,bm,True)

            return bm.Name
        
        except:
            log(inspect.stack,tb())
    
    
    def update_querverweise(self):
        if self.mb.debug: log(inspect.stack)
             
        tfm = self.mb.doc.TextFieldMasters
        
        ref_names = [ [e.split('_')[1:], tfm.getByName(e) ] for e in tfm.ElementNames if 'zzOrganonField' in e]
        sequences = [[r,m] for r,m in ref_names if r[0] == '1']
        notes     = [[r,m] for r,m in ref_names if r[0] in ('3', '4')]
                    
        bookmarks = self.mb.doc.Bookmarks
        
        ueberschriften = self.mb.class_Tools.get_ueberschriften()  
        
        for s,master in sequences:
            try:
                ref_field_source, ref_field_part = int(s[0]), int(s[1])
                sep, stamp = s[2], s[3]
                name = 'zzOrganonBM_' + stamp

                if bookmarks.hasByName(name):
                    
                    bm = bookmarks.getByName(name)
                    cur = bm.Anchor.Text.createTextCursorByRange(bm.Anchor)
                    cur.gotoEndOfParagraph(False)
                    cur.gotoStartOfParagraph(True)
                    text_bm = cur.String
                                        
                    x1 = 0
                    while not text_bm[x1].isdigit():
                        x1 += 1
                    x2 = x1
                    while text_bm[x2].isdigit():
                        x2 += 1
                    
                    nummer = text_bm[x1:x2]
                    separator = nummer + sep
                    
                    ref_txt = text_bm.split(separator)[0].strip()
                    beschriftung = text_bm.split(separator)[1].strip()
                    
                    if self.ref_field_part_dict[ref_field_part] == 'category-and-value':
                        text = ref_txt + ' ' + nummer
                        
                    elif self.ref_field_part_dict[ref_field_part] == 'caption':
                        text = beschriftung
                        
                    elif self.ref_field_part_dict[ref_field_part] == 'chapter':
                        ordinal, is_frame = self.mb.class_Bereiche.get_ordinal(bm.Anchor, True)
                        text = self.mb.class_Querverweise.get_ueberschrift2(ordinal, bm.Anchor, ueberschriften, is_frame) 
                    else:
                        text = nummer
                    
                    if master.Content != text:
                        master.Content = text
                        
                else:
                    master.Content = self.mb.settings_orga['CMDs'][self.mb.language]['REFERENZ_NICHT_GEFUNDEN']
            except:
                log(inspect.stack,tb())

                
        for s,master in notes:
            try:
                ref_field_source, ref_field_part = int(s[0]), int(s[1])
                stamp = s[2]
                name = 'zzOrganonBM_' + stamp
                
                if name in bookmarks.ElementNames:
                    
                    bm = bookmarks.getByName(name)
                    
                    if self.ref_field_part_dict[ref_field_part] == 'text':
                        text = bm.Anchor.String
                        
                    elif self.ref_field_part_dict[ref_field_part] == 'chapter':
                        ordinal, is_frame = self.mb.class_Bereiche.get_ordinal(bm.Anchor, True)
                        text = self.mb.class_Querverweise.get_ueberschrift2(ordinal, bm.Anchor, ueberschriften, is_frame)

                    if master.Content != text:
                        master.Content = text
                        
                else:
                    master.Content = self.mb.settings_orga['CMDs'][self.mb.language]['REFERENZ_NICHT_GEFUNDEN']

            except:
                log(inspect.stack,tb())
        
    
    def get_ueberschrift2(self, ordinal, text_range, ueberschriften, is_frame=False, ):
        # Ein Teile und Herrsche Algorithmus
        # fuer update_querverweise()
        # Eine Schleife waere zu langsam
        if self.mb.debug: log(inspect.stack)
        
        try:
            
            if is_frame:
                
                try:
                    text_range = text_range.TextFrame.Anchor
                except:
                    text_range = text_range
                            
            anzahl = len(ueberschriften)
            frac, index = math_modf(anzahl / 2)
            index = int(index)
            
            if frac != 0:
                index += 1

            
            while index > 1:
                
                erg = self.mb.doc.Text.compareRegionStarts( text_range, ueberschriften[index]['para'] ) 
                #  1 = Abschnitt vor Ueberschrift
                # -1 = danach
                if erg == 1:
                    ueberschriften = ueberschriften[:index]
                else:
                    ueberschriften = ueberschriften[index:]
                
                frac, index = math_modf(len(ueberschriften) / 2)
                index = int(index)
                if frac != 0:
                    index += 1
            
            if not ueberschriften:
                return '???'
                
            erg = self.mb.doc.Text.compareRegionStarts( text_range, ueberschriften[-1]['para'] ) 
            
            if erg == 1:
                ziel_index = ueberschriften[-1]['index'] - 1
            else:
                ziel_index = ueberschriften[-1]['index'] 
            
            
            if ziel_index == 0:
                # Ob sich der Textabschnitt vor der ersten Ueberschrift befindet,
                # wird von der Schleife nicht getestet, daher hier eine spezielle Abfrage.
                erg = self.mb.doc.Text.compareRegionStarts( text_range, ueberschriften[0]['para'])
                if erg == 1:
                    return '-'
                
            return ziel_index
        except:
            log(inspect.stack,tb())
            
        return '???'
    
    
                  
    def springe_zum_zielfeld(self,tf):                   
        if self.mb.debug: log(inspect.stack)
        
        try:
            vc = self.mb.viewcursor
            bookmarks = self.mb.doc.Bookmarks
            props = self.mb.props[T.AB]

            zielfeld = None
            
            if tf.TextFieldMaster.Name != '':
                ref = tf.TextFieldMaster.Name
            else:
                ref = tf.Anchor.TextField.SourceName
            
            if ref == '':
                return
            
            if 'zzOrganonField' in ref:
                ref = 'zzOrganonBM_' + ref.split('_')[-1]
            
            if bookmarks.Count != 0: 
                if bookmarks.hasByName(ref):
                    zielfeld = bookmarks.getByName(ref)
            
            if zielfeld:
                ordinal_zf = self.mb.class_Bereiche.get_ordinal(zielfeld.Anchor)
                sichtbare_ordinale = [props.dict_bereiche['Bereichsname-ordinal'][s] for s in self.mb.sichtbare_bereiche]
            
                if ordinal_zf in sichtbare_ordinale:
                    vc.gotoRange(zielfeld.Anchor,False)
                else:
                    self.mb.class_Baumansicht.selektiere_zeile(ordinal_zf)
                    vc.gotoRange(zielfeld.Anchor,False)
                  
            else:
                ref_xml = '<text:bookmark-start text:name="{0}"/>'.format(ref)
                ziel_ordinal = None
                
                for ordinal in props.dict_bereiche['ordinal']:
        
                    pf = os.path.join(self.mb.pfade['odts'], ordinal + '.odt')
                    
                    zipped = ZipFile(pf)
                    text_xml = zipped.read('content.xml').decode('utf-8')
                    zipped.close()
                    
                    if ref_xml in text_xml:
                        ziel_ordinal = ordinal
                        break
                
                if ziel_ordinal:
                    sichtbare_ordinale = [props.dict_bereiche['Bereichsname-ordinal'][s] for s in self.mb.sichtbare_bereiche]
                            
                    if ziel_ordinal not in sichtbare_ordinale:
                        self.mb.class_Baumansicht.selektiere_zeile(ziel_ordinal)
                    
                    zielfeld = None
                    if bookmarks.hasByName(ref):
                        zielfeld = bookmarks.getByName(ref)
                    
                    if zielfeld:
                        vc.gotoRange(zielfeld.Anchor,False)
                    
                    
                    else:
                        # Fallback
                        
                        # Sollte nicht mehr ausgeführt werden, da das Zielfeld
                        # gefunden sein sollte. 
                        # Springt nur ungenau an die Stelle des Zielfeldes
                        
                        text_part = '<office:text' + text_xml.split('<office:text',1)[1]
                        text_part2 = text_part.split('</office:body>')[0]
                        
                        root = self.get_root(text_part2)
                        section_xml = root.find('.//section')
                        
                        childs = list(section_xml)
                        
                        nr = None
                        for i,p in enumerate(childs):
                            if p.find(".//*[@name='{0}']".format(tf.SourceName)) != None:
                                nr = i
                                break
                        if nr:
                            sec = self.mb.doc.TextSections.getByName( props.dict_bereiche['ordinal'][ordinal] )
                            enum = sec.Anchor.createEnumeration()
                            paras = []
                            while enum.hasMoreElements():
                                paras.append(enum.nextElement())
                            
                            vc.gotoRange(paras[i],False)

                
        except:
            log(inspect.stack,tb())                        

   
    def orga_refs_in_writer_refs_umwandeln(self, doc, sections, popup):
        ''' Headings und numerierte Listen werden uebernommen            
        
        Organon spezifische Referenzen sind Bookmarks
        Referenzen sind:
        
        a) bookmark-ref
           = TextFields mit ReferenceFieldSource und ReferenceFieldPart, CurrentPresentation
            mit com.sun.star.text.TextField.GetReference
    
        b) user-field-get
           = TextFields mit TextFieldMaster
           -> Name und Content
           mit com.sun.star.text.TextField.User
        
        werden beim Export umgewandelt in Writer Referenzen 
        
        bei Nutzung eines Filters (z.B. pdf) werden
        die Referenzen automatisch in Text umgewandelt
        '''
        
        if self.mb.debug: log(inspect.stack)
        
        try:
            popup.text = LANG.WANDLE_QUERVERWEISE + str(1)  
            
            enum = doc.TextFields.createEnumeration()
            tfs = []
            while enum.hasMoreElements():
                tfs.append(enum.nextElement())
            
            
            popup.text = LANG.WANDLE_QUERVERWEISE + str(2)  
            
            bookmarks = doc.Bookmarks
            referenzierte = [bookmarks.getByName(n) for n in bookmarks.ElementNames if 'zzOrganonBM' in n]
            
            bm_refs = [[t.SourceName, t] 
                       for t in tfs 
                       if ('com.sun.star.text.TextField.GetReference' in t.SupportedServiceNames
                            and 'zzOrganonBM' in t.SourceName)]
            user_f_refs = [
                           ['zzOrganonBM_' + t.TextFieldMaster.Name.split('_')[-1],
                            t.TextFieldMaster.Name, 
                            t] 
                           
                           for t in tfs 
                           
                           if ('com.sun.star.text.TextField.User' in t.SupportedServiceNames
                                and 'zzOrganonField' in t.TextFieldMaster.Name)]
                        
            vorkommende = [b[0] for b in bm_refs]
            vorkommende.extend([b[0] for b in user_f_refs]) 
            vorkommende = set(vorkommende)
            
            referenzierte_sequenzen = [t for t in tfs if 'com.sun.star.text.TextField.SetExpression' in t.SupportedServiceNames ]
            
                        
            def get_bereiche():
                bereiche = {}
                
                for n in bookmarks.ElementNames:
                    
                    if 'zzOrganonBM' not in n:
                        continue
                    
                    bm = bookmarks.getByName(n)
                    ordinal = self.mb.class_Bereiche.get_ordinal(bm.Anchor)
                    
                    if ordinal not in bereiche:
                        bereiche.update( {ordinal : [n]})
                    else:
                        bereiche[ordinal].append(n)
                return bereiche
            
            
            def get_lesezeichen_typen(ordinal, bereiche):
                # Es waere einfacher, die Typen ueber den xml-tree des
                # neu angelegten Dokuments auszulesen. Leider sehe ich keine
                # Moeglichkeit, an den tree zu kommen.
                
                pfad = os.path.join(self.mb.pfade['odts'], ordinal + '.odt')
                zipped = ZipFile(pfad)
                text_xml = zipped.read('content.xml').decode('utf-8')
                zipped.close()
                
                fundstellen = sorted( 
                                     [ 
                                     [text_xml.find( '<text:bookmark-start text:name="' + t), t ] 
                                     for t in bereiche[ordinal] 
                                     ] 
                                     )
                fundstellen.append( [ len(text_xml) -10, ''] )
                
                # Die Listcomprehension ist etwas unuebersichtlich,
                # dafuer muss die Schleife aber nur einmal durchlaufen werden
                splittext = [ 
                             [text_xml[ v[0] : fundstellen[i+1][0] ].split( '<text:bookmark-end text:name="' + v[1] + '"/>')[0] 
                             + '<text:bookmark-end text:name="' + v[1] + '"/>',
                             v[1] ]
                             for i,v in enumerate(fundstellen[:-1]) 
                             ]
                
                typen = {}
                
                for txt, bm in splittext:
                    try:
                        bm_xml = list( self.get_root('<egal>' + txt + '</egal>') )
                        
                        typen[bm] = ( [ 'sequence',
                                        bm_xml[1].attrib['name'], 
                                        bm_xml[0].tail, 
                                        bm_xml[1].text, 
                                        bm_xml[1].tail ] 
                                      if bm_xml[1].tag == 'sequence' 
                                      
                                      else ['note', 
                                            bm_xml[1].attrib['note-class'] ])
                    except:
                        log(inspect.stack,tb())
                return typen
            
            
            def get_referenzierte(art):
                
                refs = {}
                secs = []
                
                for i,r in enumerate(referenzierte):
                    cur = r.Anchor.Text.createTextCursorByRange(r.Anchor)
                    sec = self.mb.doc.createInstance("com.sun.star.text.TextSection")
                    sec.setName('Organon_TempXYZ' + str(i))
                    r.Anchor.Text.insertTextContent(cur, sec, True)  
                                       
                    refs[sec.LocalName] = r
                    secs.append(sec)
                
                nested_refs = {r.Anchor.TextSection.LocalName : r for r in art if (r.Anchor.TextSection 
                                                                                    and 'Organon_Temp' in r.Anchor.TextSection.Name) }
                                
                refs_art = {refs[k].Name : v for k,v in nested_refs.items() }
                
                for s in secs:
                    s.dispose()
       
                return refs_art
                        
            
            bereiche = get_bereiche()
            alle_referenzierten = {}
                        
            for ordinal in bereiche:
                alle_referenzierten.update(get_lesezeichen_typen(ordinal, bereiche))
            
            popup.text = LANG.WANDLE_QUERVERWEISE + str(3)  
            
            seq_refs = get_referenzierte(referenzierte_sequenzen)
            
            popup.text = LANG.WANDLE_QUERVERWEISE + str(4)  
            
            footnotes = [doc.Footnotes.getByIndex(i) for i in range(doc.Footnotes.Count)]   
            endnotes = [doc.Endnotes.getByIndex(i) for i in range(doc.Endnotes.Count)]      
            notes = footnotes + endnotes
            
            note_refs = get_referenzierte(notes)
            
            popup.text = LANG.WANDLE_QUERVERWEISE + str(5)          
            
            
            total = len(bm_refs)
            
            for i,b in enumerate(bm_refs):
                
                if i % 50 == 0:
                    popup.text = '{0}\n{1}: {2}/{3}'.format(LANG.QUERVERWEISE_UMWANDELN, LANG.ORGA_LESEZEICHEN, i, total )
                
                name, field = b
                ref = alle_referenzierten[name]
                
                if ref[0] == 'sequence':
                    ReferenceFieldSource = 1
                    ReferenceFieldPart = field.ReferenceFieldPart
                    SourceName = ref[1]
                    
                    if ref[1] == 'Drawing':                    
                        SequenceNumber = int(seq_refs[name].Value) - 1
                    else:
                        SequenceNumber = seq_refs[name].SequenceValue
                
                elif ref[0] == 'note':

                    if ref[1] == 'footnote':
                        ReferenceFieldSource = 3
                    else:
                        ReferenceFieldSource = 4
                        
                    ReferenceFieldPart = field.ReferenceFieldPart
                    SourceName = ref[1]
                    SequenceNumber = note_refs[name].ReferenceId
                
                else:
                    Popup(self.mb).text = u"Couldn't convert Organon reference {0}".format(name)
                    continue
                
                txt_fd = doc.createInstance('com.sun.star.text.TextField.GetReference')
                txt_fd.ReferenceFieldPart = ReferenceFieldPart
                txt_fd.ReferenceFieldSource = ReferenceFieldSource
                txt_fd.SequenceNumber = SequenceNumber
                txt_fd.SourceName = SourceName
                
                cur = field.Anchor.Text.createTextCursorByRange(field.Anchor)
                
                cur.Text.insertTextContent(cur,txt_fd,True)
                field.dispose()
            
            total = len(user_f_refs)
            
            for i,b in enumerate(user_f_refs):
                
                if i % 50 == 0:
                    popup.text = '{0}\n{1}: {2}/{3}'.format(LANG.QUERVERWEISE_UMWANDELN, LANG.ORGA_FELDER, i, total )
                
                name, full_name, field = b
                ref = alle_referenzierten[name]
                
                ReferenceFieldSource, ReferenceFieldPart = full_name.split('_')[1:3]
                SourceName = ref[1]
                
                if self.reference_field_source[int(ReferenceFieldSource)] == 'sequence':
                     
                    if ref[1] == 'Drawing':                    
                        SequenceNumber = int(seq_refs[name].Value) - 1
                    else:
                        SequenceNumber = seq_refs[name].SequenceValue
                 
                elif self.reference_field_source[int(ReferenceFieldSource)] in ('footnote', 'endnote'):
                    SequenceNumber = note_refs[name].ReferenceId
                else:
                    Popup(self.mb).text = u"Couldn't convert Organon reference {0}".format(name)
                    continue
                 
                txt_fd = doc.createInstance('com.sun.star.text.TextField.GetReference')
                txt_fd.ReferenceFieldPart = ReferenceFieldPart
                txt_fd.ReferenceFieldSource = ReferenceFieldSource
                txt_fd.SequenceNumber = SequenceNumber
                txt_fd.SourceName = SourceName
                 
                cur = field.Anchor.Text.createTextCursorByRange(field.Anchor)
                 
                cur.Text.insertTextContent(cur,txt_fd,True)
                field.dispose()
            
            
            
            for n in bookmarks.ElementNames:
                    
                if 'zzOrganonBM' in n:
                    bm = bookmarks.getByName(n)
                    bm.dispose()
    
        except:
            log(inspect.stack,tb())

 
    def time_stamp(self):
        time.sleep(0.001)
        return datetime.utcnow().strftime('%H%M%S%f')[:-3]           
        
    
    def benutzerfeld_anlegen(self, doc, cur, referenz, inhalt):
            
        txt_fd = doc.createInstance("com.sun.star.text.textfield.User")
        txt_fd.NumberFormat = -1
        
        if 'com.sun.star.text.fieldmaster.User.' + referenz in doc.TextFieldMasters.ElementNames:
            master = doc.TextFieldMasters.getByName('com.sun.star.text.fieldmaster.User.' + referenz)
        else:
            master = doc.createInstance("com.sun.star.text.fieldmaster.User")
            master.setPropertyValue ("Name", referenz)
            master.setPropertyValue ("Content", inhalt)
        
        txt_fd.attachTextFieldMaster(master)
        master.Content = inhalt
        cur.Text.insertTextContent(cur,txt_fd,False)    
        
        
    def get_verweise(self,doc): 
        
        try:
            enum = doc.TextFields.createEnumeration()
            
            verweise = {
                        'organon_bm' : {},
                        'listen_und_ueberschriften' : {},
                        'benutzer': {},
                        'organon_felder' : {},
                        'benutzer_feld' : {}
                        }
            
            while enum.hasMoreElements():
                
                tf = enum.nextElement()
                
                if 'com.sun.star.text.TextField.GetReference' in tf.SupportedServiceNames:
                    source_name = tf.SourceName
                    
                    if 'zzOrganonBM' in source_name:
                        key = 'organon_bm'
                    elif '__Ref' in source_name:
                        key = 'listen_und_ueberschriften'
                    else:
                        key = 'benutzer'
                        
                    if source_name in verweise[key]:
                        verweise[key][source_name].append(tf)
                    else:
                        verweise[key][source_name] = [tf]
                        
                elif 'com.sun.star.text.TextField.User' in tf.SupportedServiceNames:
                    
                    source_name = 'zzOrganonBM_' + tf.TextFieldMaster.Name.split('_')[-1]
                    
                    if 'zzOrganonField' in tf.TextFieldMaster.Name:
                        key = 'organon_felder'
                    else:
                        key = 'benutzer_feld'
                
                    if source_name in verweise[key]:
                        verweise[key][source_name].append(tf)
                    else:
                        verweise[key][source_name] = [tf]
            
            return verweise
        
        except:
            log(inspect.stack,tb())
            return verweise   
    
    
    def get_xml(self, ordinal, art='content.xml', split_text=True):
        if self.mb.debug: log(inspect.stack)
        try:
            pfad = os.path.join(self.mb.pfade['odts'], ordinal + '.odt')
                
            zipped = ZipFile(pfad)
            text_xml = zipped.read(art)
            zipped.close()
            
            if split_text:
                return text_xml.split(b'<office:text',1)[1]
            else:
                return text_xml
        except:
            log(inpsect.stack,tb())
            return b''
    
    def get_referenzierte_listen_und_ueberschriften(self, ordinal, doc ):
        if self.mb.debug: log(inspect.stack)
        try:                   
            text_part = self.get_xml(ordinal)
            
            listen_bms = [t.split(b'"/>')[0] for t in text_part.split(b'<text:bookmark-start text:name="__RefNumPara__')][1:]
            listen_namen = ['__RefNumPara__' + b.decode('utf-8') for b in listen_bms]
            refs_listen = {n : doc.Bookmarks.getByName(n) for n in listen_namen }
            
            ueber_bms = [t.split(b'"/>')[0] for t in text_part.split(b'<text:bookmark-start text:name="__RefHeading__')][1:]
            ueber_namen = ['__RefHeading__' + b.decode('utf-8') for b in ueber_bms]
            refs_ueberschriften = { n : doc.Bookmarks.getByName(n) for n in ueber_namen }
            
            ziele = {
                     'listen' : refs_listen,
                     'ueberschriften' : refs_ueberschriften
                     }
            namen = listen_namen + ueber_namen
            
            return ziele, namen
        except:
            log(inspect.stack,tb())
            return {}
                
        
    def querverweise_umbenennen(self, doc, ordinal, verlinken):
        if self.mb.debug: log(inspect.stack)
        
        try:
            bookmarks = doc.Bookmarks
            verweise = self.get_verweise(doc)
                    
            ziele_namen = self.get_lesezeichen_in_content_xml(ordinal)
            ziele_namen = [ b for b in ziele_namen if 'zzOrganonBM' in b ]
            ziele_li_ue, ziele_li_ue_namen = self.get_referenzierte_listen_und_ueberschriften(ordinal, doc)
            
                    
            def get_neue_verweisnamen():
                
                neue_refs = { b : self.time_stamp()  for b in ziele_namen }
                
                for alter_name in ziele_li_ue_namen:
                    gesplittet = alter_name.split('_')
                    art = gesplittet[2]
                    nummern = gesplittet[-2:]
                    neuer_name = '__{0}__{1}_{2}'.format( art, nummern[0], self.time_stamp() )
                    neue_refs.update({alter_name : neuer_name})
                    
                return neue_refs
            
            
            # VERWEISE LISTEN UND UEBERSCHRIFTEN umbenennen            
            def benenne_listen_und_ueberschriften_um():
                                
                def umbenennen(ziel, art, key, tfs_verweise):
                    
                    neuer_name = neue_refs[ziel.Name]
                    ziel.setName(neuer_name)
                    
                    if key in tfs_verweise['listen_und_ueberschriften']:
                        for v in tfs_verweise['listen_und_ueberschriften'][key]:
                            v.SourceName = neuer_name
                                        
                    if art == 'RefHeading':
                        # Wegen eines Bugs in Writer, der bei einem Link zu einer 
                        # Region eines Bereiches  __RefHeading__ loescht, wird
                        # hier ein Paragraph eingefuegt, der in funktionen BatchImport
                        # neue_dateien_einfuegen() wieder geloescht wird.
                        cur = ziel.Anchor.Text.createTextCursorByRange(ziel.Anchor)
                        cur.gotoStartOfParagraph(False)
                        ziel.Anchor.Text.insertControlCharacter( cur, PARAGRAPH_BREAK, 0 )
                        cur.goLeft(1,False)
                        cur.setString('***Inserted By Organon***')
                        
                
                for key,ziel in ziele_li_ue['listen'].items():
                    umbenennen(ziel, 'RefNumPara', key, verweise)
                            
                for key,ziel in ziele_li_ue['ueberschriften'].items():
                    umbenennen(ziel, 'RefHeading', key, verweise)
                
            
            neue_refs = get_neue_verweisnamen()   
            
            benenne_listen_und_ueberschriften_um()    

            ziele = [ bookmarks.getByName(b) for b in ziele_namen ]
            
            # ZIELE umbenennen
            for b in ziele:
                b.setName( 'zzOrganonBM_' + neue_refs[b.Name] )
                
            # VERWEISE umbenennen
            def verweise_umbenennen(verweise_dict,odoc,ziele_namen2=None):
                # BOOKMARKS umbenennen
                for key,verw in verweise_dict['organon_bm'].items():
                    if ziele_namen2:
                        if key not in ziele_namen2:
                            continue
                    for b in verw:
                        b.SourceName = 'zzOrganonBM_' + neue_refs[key]
                # FELDER umbenennen
                for key,verw in verweise_dict['organon_felder'].items():
                    if ziele_namen2:
                        if key not in ziele_namen2:
                            continue
                    for b in verw:
                        splittext = b.TextFieldMaster.Name.split('_')
                        name = '_'.join(splittext[:-1])
                        stamp = splittext[-1]
                        referenz = name + '_' + neue_refs['zzOrganonBM_' + stamp]
                        
                        text_range = b.Anchor
                        inhalt = b.TextFieldMaster.Content
            
                        self.benutzerfeld_anlegen(odoc, text_range, referenz, inhalt)
                        b.dispose()  
            
            verweise_umbenennen(verweise,doc)
            
             
            def verlinkungen_umbenennen():
                # VERLINKTE VERWEISE umbenennen

                # Verweise in anderen Dateien 
                # auf die zu trennende Datei suchen
                props = self.mb.props['ORGANON']                
                ordinale = [ o for o in props.dict_bereiche['ordinal'] if o != ordinal ]
                
                try:
                    for o in ordinale:
                        
                        text_part = self.get_xml(o)
                                            
                        for name in neue_refs:
                            if bytes(name,'utf-8') in text_part:
                                
                                url = props.dict_bereiche['Bereichsname'][ props.dict_bereiche['ordinal'][o] ]
                                url = uno.systemPathToFileUrl( url )
                                doc2 = self.mb.doc.CurrentController.Frame.loadComponentFromURL(url,'_blank',0,(PROP_HIDDEN,))                            
                                verlinkte = self.get_verweise(doc2)
                                
                                verweise_umbenennen( verlinkte, doc2, list(neue_refs) )
                                
                                for key,ver in verlinkte['listen_und_ueberschriften'].items():
                                    for v in ver:
                                        if v.SourceName in neue_refs:
                                            v.SourceName = neue_refs[v.SourceName]
                                
                                doc2.store()
                                doc2.close(False)
                                break
                            
                except:
                    log(inspect.stack,tb())
            
            if verlinken:
                verlinkungen_umbenennen()    

        except:
            log(inspect.stack,tb())
            
        








            
            

        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
            
            
            
            