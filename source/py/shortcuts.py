# -*- coding: utf-8 -*-

import unohelper


class Shortcuts():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
    
        AZ = [chr(i).upper() for i in range(ord('a'), ord('z')+1)]
        numbers = [i for i in range(10)]
        
        self.keycodes = {}
        
        for a in AZ:
            co = uno.getConstantByName( "com.sun.star.awt.Key.{}".format(a) )
            self.keycodes.update({co:a.lower()})
        for a in numbers:
            co = uno.getConstantByName( "com.sun.star.awt.Key.NUM{}".format(a) )
            self.keycodes.update({co:a})
        for a in ['DOWN','UP','LEFT','RIGHT']:
            co = uno.getConstantByName( "com.sun.star.awt.Key.{}".format(a) )
            self.keycodes.update({co:a})
        
        self.nutze_shortcuts = True
        self.shortcuts = {
                         'd' : self.teile_text,
                         'n' : self.erzeuge_neue_Datei,
                         'm' : self.erzeuge_neuen_Ordner,
                         'x' : self.in_Papierkorb_einfuegen,
                         'y' : self.leere_Papierkorb,
                         's' : self.datei_nach_aenderung_speichern,
                         't' : self.starte_neuen_Tab,
                         'w' : self.schliesse_Tab,
                         'r' : self.erzeuge_Backup,
                         'o' : self.oeffne_Organizer,
                         'j' : self.toggle_tag1,
                         'k' : self.toggle_tag2,
                         'l' : self.toggle_tag3,
                         'UP': self.tv_up,
                         'DOWN':self.tv_down
                        }  


        
    def shortcut_ausfuehren(self,code):
        if self.mb.debug: log(inspect.stack)
        print(code)
        if code in self.keycodes:
            keychar = self.keycodes[code]
        else:
            keychar = -1

        if keychar in self.shortcuts:
            self.shortcuts[keychar]()
        print(keychar) 
    
    def teile_text(self):
        if self.mb.debug: log(inspect.stack)

        if T.AB != 'Projekt': return
        self.mb.class_Funktionen.teile_text()
        
    def erzeuge_neue_Datei(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB != 'Projekt': return
        self.mb.class_Baumansicht.erzeuge_neue_Zeile('dokument')
    
    def erzeuge_neuen_Ordner(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB != 'Projekt': return
        self.mb.class_Baumansicht.erzeuge_neue_Zeile('ordner')
        
    def in_Papierkorb_einfuegen(self):
        if self.mb.debug: log(inspect.stack)
        
        ordinal = self.mb.props[T.AB].selektierte_zeile
        papierkorb = self.mb.props[T.AB].Papierkorb
        self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(ordinal,papierkorb,'inPapierkorbEinfuegen')
        
    def leere_Papierkorb(self):
        if self.mb.debug: log(inspect.stack)
        self.mb.class_Baumansicht.leere_Papierkorb()  
        
    def datei_nach_aenderung_speichern(self):
        if self.mb.debug: log(inspect.stack)
        
        props = self.mb.props[T.AB]
        zuletzt = props.selektierte_zeile_alt
        bereichsname = props.dict_bereiche['ordinal'][zuletzt]
        path = props.dict_bereiche['Bereichsname'][bereichsname]
        self.mb.props[T.AB].tastatureingabe = True

        self.mb.class_Bereiche.datei_nach_aenderung_speichern(uno.systemPathToFileUrl(path),bereichsname)
            
    def starte_neuen_Tab(self):
        if self.mb.debug: log(inspect.stack)
        self.mb.class_Tabs.start(False)
        
    def schliesse_Tab(self):
        if self.mb.debug: log(inspect.stack)
        self.mb.class_Tabs.schliesse_Tab()
        
    def erzeuge_Backup(self):
        if self.mb.debug: log(inspect.stack)
        self.mb.erzeuge_Backup()
        
    def oeffne_Organizer(self):
        if self.mb.debug: log(inspect.stack)
        self.mb.class_Organizer.run()
        
    def toggle_tag1(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_proj
        nummer = 1
        tag = 'tag%s'%nummer
        sett[tag] = not sett[tag]
        
        self.mb.class_Funktionen.mache_tag_sichtbar(sett[tag],tag)
        
        self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        
    def toggle_tag2(self):
        if self.mb.debug: log(inspect.stack)
        
        if not self.mb.class_Funktionen.pruefe_galerie_eintrag():
            return
        
        sett = self.mb.settings_proj
        nummer = 2
        tag = 'tag%s'%nummer
        sett[tag] = not sett[tag]
        
        self.mb.class_Funktionen.mache_tag_sichtbar(sett[tag],tag)
        
        self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        
    def toggle_tag3(self):
        if self.mb.debug: log(inspect.stack)
        
        sett = self.mb.settings_proj
        nummer = 3
        tag = 'tag%s'%nummer
        sett[tag] = not sett[tag]
        
        self.mb.class_Funktionen.mache_tag_sichtbar(sett[tag],tag)
        
        self.mb.speicher_settings("project_settings.txt", self.mb.settings_proj)  
        
    def tv_up(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            selektierte_zeile = self.mb.props[T.AB].selektierte_zeile_alt
            props = self.mb.props[T.AB]
            
            tree = self.mb.props[T.AB].xml_tree
            root = tree.getroot()        
            source = root.find('.//'+selektierte_zeile)            
            
            eintr = []
            self.mb.class_XML.get_tree_info(root,eintr)
            
            vorgaenger = None
            
            for e in eintr:
                if e[0] == selektierte_zeile:
                    index = eintr.index(e)
                    if index > 0:
                        vorgaenger = eintr[index-1][0]
                    
            if vorgaenger != None:
                self.mb.class_Baumansicht.selektiere_zeile(vorgaenger)
            
        except:
            print(tb())
        
    def tv_down(self):
        if self.mb.debug: log(inspect.stack)
        
        try:
            selektierte_zeile = self.mb.props[T.AB].selektierte_zeile_alt
            props = self.mb.props[T.AB]
            
            tree = self.mb.props[T.AB].xml_tree
            root = tree.getroot()        
            source = root.find('.//'+selektierte_zeile)            
            
            eintr = []
            self.mb.class_XML.get_tree_info(root,eintr)
            
            nachfolger = None
            
            for e in eintr:
                if e[0] == selektierte_zeile:
                    index = eintr.index(e)
    
                    if index < len(eintr)-1:
                        nachfolger = eintr[index+1][0]
                        
            if nachfolger != None:
                self.mb.class_Baumansicht.selektiere_zeile(nachfolger)

        except:
            print(tb())
        
        
        
        
        
        
        
        
        
        
#         elif code == 532 and mods == 6: # ctrl alt r
#             ordinal = self.mb.props[T.AB].selektierte_zeile
#             print(ordinal)
#             print(self.mb.props['Projekt'].dict_ordner)
#             if ordinal not in self.mb.props['Projekt'].dict_ordner:
#                 return
#             self.mb.class_Funktionen.projektordner_ausklappen(ordinal)
