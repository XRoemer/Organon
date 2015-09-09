# -*- coding: utf-8 -*-

import unohelper


class Shortcuts():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        
    
        AZ = [chr(i).upper() for i in range(ord('a'), ord('z')+1)]
        F1F12 = ['F'+str(i) for i in range(1,13)]
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
        for a in F1F12:
            co = uno.getConstantByName( "com.sun.star.awt.Key.{}".format(a) )
            self.keycodes.update({co:a})
        
        
        self.moegliche_shortcuts = AZ + numbers + F1F12 + ['DOWN','UP','LEFT','RIGHT']
        
        self.shortcuts = self.mb.settings_orga['shortcuts']
        self.writer_shortcuts = self.mb.class_Funktionen.get_writer_shortcuts()
        
        
        self.shortcuts_befehle = {
                         'TRENNE_TEXT' : self.teile_text,
                         'INSERT_DOC' : self.erzeuge_neue_Datei,
                         'INSERT_DIR' : self.erzeuge_neuen_Ordner,
                         'IN_PAPIERKORB_VERSCHIEBEN' : self.in_Papierkorb_einfuegen,
                         'CLEAR_RECYCLE_BIN' : self.leere_Papierkorb,
                         'FORMATIERUNG_SPEICHERN2' : self.datei_nach_aenderung_speichern,
                         'NEUER_TAB' : self.starte_neuen_Tab,
                         'SCHLIESSE_TAB' : self.schliesse_Tab,
                         'BACKUP' : self.erzeuge_Backup,
                         'OEFFNE_ORGANIZER' : self.oeffne_Organizer,
                         'SHOW_TAG1' : self.toggle_tag1,
                         'SHOW_TAG2' : self.toggle_tag2,
                         'GLIEDERUNG' : self.toggle_tag3,
                         'BAUMANSICHT_HOCH' : self.tv_up,
                         'BAUMANSICHT_RUNTER' : self.tv_down                                  
                          }


        
    def shortcut_ausfuehren(self,code,mods):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if code in self.keycodes:
                keychar = self.keycodes[code]
            else:
                keychar = -1
            
            if keychar.upper() in self.shortcuts[str(mods)]:
                cmd = self.shortcuts[str(mods)][keychar.upper()]
                self.shortcuts_befehle[cmd]()
        except:
            log(inspect.stack,tb())
            
    
    def teile_text(self):
        if self.mb.debug: log(inspect.stack)

        if T.AB != 'Projekt': return
        self.mb.class_Funktionen.teile_text()
        
    def erzeuge_neue_Datei(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB != 'Projekt': return
        self.mb.class_Baumansicht.erzeuge_neue_Zeile('Dokument')
    
    def erzeuge_neuen_Ordner(self):
        if self.mb.debug: log(inspect.stack)
        
        if T.AB != 'Projekt': return
        self.mb.class_Baumansicht.erzeuge_neue_Zeile('Ordner')
        
    def in_Papierkorb_einfuegen(self):
        if self.mb.debug: log(inspect.stack)
        
        nachfolger = self.mb.class_XML.finde_nachfolger_oder_vorgaenger('nachfolger')    
        vorgaenger = self.mb.class_XML.finde_nachfolger_oder_vorgaenger('vorgaenger') 
        
        ordinal = self.mb.props[T.AB].selektierte_zeile
        papierkorb = self.mb.props[T.AB].Papierkorb
        self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(ordinal,papierkorb,'inPapierkorbEinfuegen')
        
        if nachfolger != None:
            self.mb.class_Baumansicht.selektiere_zeile(nachfolger)
        else:
            if vorgaenger != None:
                self.mb.class_Baumansicht.selektiere_zeile(vorgaenger)

        
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
        
        # Bestaetigung ausgeben
        tree = self.mb.props[T.AB].xml_tree
        root = tree.getroot()        
        source = root.find('.//'+zuletzt)  
        name = source.attrib['Name']
        
        self.mb.nachricht_si(LANG.FORMATIERUNG_SPEICHERN.format(name),1)
            
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
        
        vorgaenger = self.mb.class_XML.finde_nachfolger_oder_vorgaenger('vorgaenger') 
        if vorgaenger != None:
            self.mb.class_Baumansicht.selektiere_zeile(vorgaenger)
        
    def tv_down(self):
        if self.mb.debug: log(inspect.stack)

        nachfolger = self.mb.class_XML.finde_nachfolger_oder_vorgaenger('nachfolger')    
        if nachfolger != None:
            self.mb.class_Baumansicht.selektiere_zeile(nachfolger)
 

    def get_mods(self,cmd,ctrls):
        if self.mb.debug: log(inspect.stack)
        # 0 = keine Modifikation
        # 1 = Shift
        # 2 = Strg
        # 3 = Shift + Strg
        # 4 = Alt
        # 5 = Shift + Alt
        # 6 = Strg + Alt
        # 7 = Shift + Strg + Alt
                        
        shift = ctrls['_Shift'+cmd].State
        alt = ctrls['_Alt'+cmd].State
        ctrl = ctrls['_Ctrl'+cmd].State
        
        mods = 0
        
        if shift: mods += 1
        if alt: mods += 4
        if ctrl: mods += 2
        mods = shift*1 + ctrl*2 + alt*4
        return mods
    
    def get_moegliche_shortcuts(self,mods,use_settings = True):
        if self.mb.debug: log(inspect.stack)
        
        moegliche = self.mb.class_Shortcuts.moegliche_shortcuts
        
        if mods < 2:
            return ('-',)
        
        if use_settings:
            in_settings = self.mb.settings_orga['shortcuts'][str(mods)]
        else:
            in_settings = []
                                                                
        uebrige = [str(m) for m in moegliche if str(m) not in self.writer_shortcuts[mods] and
                                           str(m) not in in_settings
                   ]
        uebrige.insert(0, '-')
        uebrige = tuple(uebrige)
        
        return uebrige
        
    
        
#         elif code == 532 and mods == 6: # ctrl alt r
#             ordinal = self.mb.props[T.AB].selektierte_zeile
#             print(ordinal)
#             print(self.mb.props['Projekt'].dict_ordner)
#             if ordinal not in self.mb.props['Projekt'].dict_ordner:
#                 return
#             self.mb.class_Funktionen.projektordner_ausklappen(ordinal)
