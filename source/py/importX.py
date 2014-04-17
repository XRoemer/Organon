# -*- coding: utf-8 -*-

#print('Export')
# import traceback
# import uno
import unohelper
# import os

# tb = traceback.print_exc

EXPORT_DIALOG_FARBE = 305099

class ImportX():
    
    def __init__(self,mb,pdk):
        self.mb = mb
        
        global pd
        pd = pdk
        
        global lang
        lang = self.mb.lang
        
        
    def importX(self):
        self.erzeuge_importfenster()
 
    def erzeuge_importfenster(self): 
        
        imp_set = self.mb.settings_imp

        posSize_main = self.mb.desktop.ActiveFrame.ContainerWindow.PosSize
        X = posSize_main.X +20
        Y = posSize_main.Y +20
        Width = 360
        Height = 340
        
        posSize = X,Y,Width,Height
        fenster,fenster_cont = erzeuge_Dialog_Container(self.mb.smgr,self.mb.ctx,posSize)
        fenster_cont.Model.Text = lang.IMPORT
        
        y = 10
        
        # Titel
        controlE, modelE = createControl(self.mb.ctx,"FixedText",20,y ,50,20,(),() )  
        controlE.Text = lang.IMPORT
        modelE.FontWeight = 200.0
        fenster_cont.addControl('Titel', controlE)
        
        y += 30
        
        control1, model1 = createControl(self.mb.ctx,"CheckBox",20,y,120,22,(),() )  
        model1.Label = lang.IMPORT_DATEI
        model1.State = int(imp_set['imp_dat'])
        control1.ActionCommand = 'model1'
        fenster_cont.addControl('ImportD', control1)
        
        
        
        control2, model2 = createControl(self.mb.ctx,"CheckBox",160,y,120,22,(),() )  
        model2.Label = lang.IMPORT_ORDNER
        model2.State = not int(imp_set['imp_dat'])
        control2.ActionCommand = 'model2'
        fenster_cont.addControl('ImportO', control2)

        y += 20
        
        control3, model3 = createControl(self.mb.ctx,"CheckBox",180,y,180,22,(),() )  
        model3.Label = lang.ORDNERSTRUKTUR
        model3.State =  int(imp_set['ord_strukt'])
        control3.Enable = not int(imp_set['imp_dat'])
        control3.ActionCommand = 'struktur'
        fenster_cont.addControl('struktur', control3)
        
        
        # CheckBox Listener
        listenerA = Auswahl_CheckBox_Listener(self.mb,model1,model2,control3)
        control1.addActionListener(listenerA)
        control2.addActionListener(listenerA)
        control3.addActionListener(listenerA)
                    
        y += 40
        
        controlE, modelE = createControl(self.mb.ctx,"FixedText",20,y ,50,20,(),() )  
        controlE.Text = lang.DATEIFILTER
        fenster_cont.addControl('Filter', controlE)
        
        buttons = []
        labels = ('odt','doc','docx','rtf','txt')
         
        for label in labels:
            control, model = createControl(self.mb.ctx,"CheckBox",100,y,120,22,(),() )  
            model.Label = label
            model.State = imp_set[label]
            
            if label != 'odt':
                control.Enable = False
                 
            fenster_cont.addControl(label, control)
            buttons.append(control)
            y += 16
            
        y += 20   
        
        controlA, modelA = createControl(self.mb.ctx,"Button",20,y  ,110,20,(),() )  ###
        controlA.Label = lang.DATEIAUSWAHL
        controlA.ActionCommand = 'Datei'
        fenster_cont.addControl('Auswahl', controlA) 
        
        y += 25
        
        controlE, modelE = createControl(self.mb.ctx,"FixedText",20,y ,600,20,(),() )  
        text = imp_set['url_dat']
        if text == None or text == '': 
            text = '-'
        else:
            text = uno.fileUrlToSystemPath(decode_utf(text))
        controlE.Text = text
        fenster_cont.addControl('Filter', controlE)    
        
        y += 20   
        
        controlA2, modelA2 = createControl(self.mb.ctx,"Button",20,y  ,110,20,(),() )  ###
        controlA2.Label = lang.ORDNERAUSWAHL
        controlA2.ActionCommand = 'Ordner'
        fenster_cont.addControl('Auswahl2', controlA2)             
        
        y += 25
        
        controlE2, modelE2 = createControl(self.mb.ctx,"FixedText",20,y ,600,20,(),() )  
        text = imp_set['url_ord']
        if text == None or text == '': 
            text = '-'
        else:
            text = uno.fileUrlToSystemPath(decode_utf(text))
        controlE2.Text = text
        fenster_cont.addControl('Filter2', controlE2) 
        
        listenerA2 = Auswahl_Button_Listener(self.mb,modelE,modelE2)
        controlA.addActionListener(listenerA2)
        controlA2.addActionListener(listenerA2)   
        
        y += 30
        
        controlI, modelI = createControl(self.mb.ctx,"Button",280,y  ,70,30,(),() )  ###
        controlI.Label = lang.IMPORTIEREN
        listener_imp = Import_Button_Listener(self.mb)
        controlI.addActionListener(listener_imp)
        fenster_cont.addControl('importieren', controlI) 



def encode_utf(term):
    if isinstance(term, str):
        return term
    else:
        return term.encode('utf8')

def decode_utf(term):
    if isinstance(term, str):
        return term
    else:
        return term.decode('utf8')   

def escape_xml(term):
    Zeichen = {
               ' ' : '_Leerzeichen_',
               '(' : '_KlammerAuf_',
               ')' : '_KlammerZu_',
               '.' : '_PuNkt_'}
    
    for z in Zeichen:        
        term = term.replace(z, Zeichen[z])
    return term

def unescape_xml(term):
    Zeichen = {
               '_Leerzeichen_' : ' ',
               '_KlammerAuf_' : '(',
               '_KlammerZu_' : ')',
               '_PuNkt_' : '.'}
    
    for z in Zeichen:        
        term = term.replace(z, Zeichen[z])
    return term

def erzeuge_Dialog_Container(smgr,ctx,posSize):
    
    X,Y,Width,Height = posSize

    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)    
    oCoreReflection = smgr.createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)

    # Create Uno Struct
    oXIdlClass = oCoreReflection.forName("com.sun.star.awt.WindowDescriptor")
    oReturnValue, oWindowDesc = oXIdlClass.createObject(None)
    # global oWindow
    oWindowDesc.Type = uno.Enum("com.sun.star.awt.WindowClass", "TOP")
    oWindowDesc.WindowServiceName = ""
    oWindowDesc.Parent = toolkit.getDesktopWindow()
    oWindowDesc.ParentIndex = -1
    oWindowDesc.WindowAttributes = 1  +32 +64 + 128 # Flags fuer com.sun.star.awt.WindowAttribute

    oXIdlClass = oCoreReflection.forName("com.sun.star.awt.Rectangle")
    oReturnValue, oRect = oXIdlClass.createObject(None)
    oRect.X = X
    oRect.Y = Y
    oRect.Width = Width 
    oRect.Height = Height 
    
    oWindowDesc.Bounds = oRect

    # create window
    oWindow = toolkit.createWindow(oWindowDesc)
     
    # create frame for window
    oFrame = smgr.createInstanceWithContext("com.sun.star.frame.Frame",ctx)
    oFrame.initialize(oWindow)
    #oFrame.setCreator(self.mb.desktop)
    oFrame.activate()

    # create new control container
    cont = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", ctx)
    cont_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", ctx)
    cont_model.BackgroundColor = EXPORT_DIALOG_FARBE  # 9225984
    cont.setModel(cont_model)
    # need createPeer just only the container
    cont.createPeer(toolkit, oWindow)
    cont.setPosSize(0, 0, 0, 0, 15)

    oFrame.setComponent(cont, None)
    return oWindow,cont
 

from com.sun.star.awt import XItemListener, XActionListener, XFocusListener    

class Import_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb):
        self.mb = mb

        
    def actionPerformed(self,ev):

        imp_set = self.mb.settings_imp
        
        if imp_set['imp_dat'] == '1':
            url_dat = imp_set['url_dat']
            typ = 'dokument'
            self.datei_importieren(typ,url_dat)
        else:
            self.ordner_importieren()
        
        
        
    def datei_importieren(self,typ,url_dat=None,name2=None):

        zeile_nr,zeile_pfad = self.mb.class_Hauptfeld.erzeuge_neue_Zeile(typ,False)
        
        sections = self.mb.doc.TextSections

        ordinal_neuer_Eintrag = 'nr'+str(zeile_nr)
        bereichsname = self.mb.dict_bereiche['ordinal'][ordinal_neuer_Eintrag]
        neuer_Bereich = sections.getByName(bereichsname)            
                            
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = 'Hidden'
        prop.Value = True

        if url_dat != "private:factory/swriter":
            link_quelle = url_dat
        else:
            link_quelle = "private:factory/swriter"
          
        self.oOO = self.mb.desktop.loadComponentFromURL(link_quelle,'_blank',8+32,(prop,))
        
        Path2 = uno.systemPathToFileUrl(zeile_pfad)
        self.oOO.storeToURL(Path2,())
        self.oOO.close(False)
        
        SFLink = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLinkL = uno.createUnoStruct("com.sun.star.text.SectionFileLink")
        SFLink = neuer_Bereich.FileLink
         
        neuer_Bereich.setPropertyValue('FileLink',SFLinkL)
        neuer_Bereich.setPropertyValue('FileLink',SFLink)
        
        # zeilenname der neuen Datei setzen
        zeile_hf = self.mb.Hauptfeld.getControl('nr'+str(zeile_nr))
        zeile_textfeld = zeile_hf.getControl('textfeld')

        if url_dat != "private:factory/swriter":
            link_quelle1 = uno.fileUrlToSystemPath(link_quelle)
            endung = os.path.splitext(link_quelle1)[1]
            name1 = os.path.basename(link_quelle1)
            name2 = name1.split(endung)[0]

        zeile_textfeld.Model.Text = name2
        
        tree = self.mb.xml_tree
        root = tree.getroot() 
        
        xml_elem = root.find('.//'+ ordinal_neuer_Eintrag)
        xml_elem.attrib['Name'] = name2
        
        pfad = os.path.join(self.mb.pfade['settings'],'ElementTree.xml')
        self.mb.xml_tree.write(pfad)

        return ordinal_neuer_Eintrag,bereichsname
    
    
    def ordner_importieren(self):
        
        imp_set = self.mb.settings_imp
        try:
            path = uno.fileUrlToSystemPath(imp_set['url_ord'])
            Verzeichnis,AusgangsOrdner = self.durchlaufe_ordner(path)
            
            # Elementtree ist nur ein Mittel,
            # die Ordnereintraege richtig zu setzen
            et = self.mb.ET
            root_xml = et.Element(AusgangsOrdner)
            tree_xml = et.ElementTree(root_xml)
            
            # Der Hauptordner
            ordinal_neuer_Eintrag,bereichsname = self.datei_importieren('Ordner',"private:factory/swriter",AusgangsOrdner)
            root_xml.text = ordinal_neuer_Eintrag
            self.setze_selektierte_zeile(ordinal_neuer_Eintrag)
            
                     
            if imp_set['ord_strukt'] == '0':

                for eintr in Verzeichnis:
                    File,Pfad,Ordner = eintr
                    elem = root_xml
                    
                    for o in Ordner[::-1]:
                        o = escape_xml(o)
                        if o != AusgangsOrdner:
                            x = elem.find(o)
                            if x == None:
                                #print('schreibe Ordner',o, 'in Ordner',elem.tag)
                                et.SubElement(elem,o)
                            elem = elem.find(o)
                             
                    print('schreibe Datei',File,'in Ordner',o) 
                    self.text = uno.systemPathToFileUrl(Pfad)
                    ordinal_neuer_Eintrag,bereichsname = self.datei_importieren('Datei',self.text,File)
                    self.setze_selektierte_zeile(ordinal_neuer_Eintrag)   
                    
            else:
                Verzeichnis = Verzeichnis[::-1]
                for eintr in Verzeichnis:
                    File,Pfad,Ordner = eintr
                    elem = root_xml
                    
                    for o in Ordner[::-1]:
                        o = escape_xml(o)
           
                        if o != AusgangsOrdner:

                            x = elem.find(o)
                            if x == None:
                                self.setze_selektierte_zeile(elem.text) 
                                
                                name = unescape_xml(o)
                                
                                ordinal_neuer_Eintrag,bereichsname = self.datei_importieren('Ordner',"private:factory/swriter",name)
                                et.SubElement(elem,o)
                                element = elem.find(o)
                                element.text = ordinal_neuer_Eintrag
                                
                                source = ordinal_neuer_Eintrag
                                target = elem.text#self.mb.selektierte_zeile.AccessibleName
                                action = 'inOrdnerEinfuegen'
                                self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(source,target,action)
                                #print('schreibe Ordner',o, 'in Ordner',elem.tag)
                                
                            elem = elem.find(o)
                    
                    self.setze_selektierte_zeile(elem.text)         
                    #print('schreibe Datei',File,'in Ordner',o) 
                    self.text = uno.systemPathToFileUrl(Pfad)
                    ordinal_neuer_Eintrag,bereichsname = self.datei_importieren('Datei',self.text)
                    #ordner = elem.find(o)
                    et.SubElement(elem,File)
                    
                    
                    source = ordinal_neuer_Eintrag
                    target = elem.text 
                    action = 'inOrdnerEinfuegen'
                    self.mb.class_Zeilen_Listener.zeilen_neu_ordnen(source,target,action)
 
        except Exception as e:
            self.mb.Mitteilungen.nachricht(str(e),"warningbox")
            tb()
    
    
    def durchlaufe_ordner(self,path):
    
    
        def find_parents(ord,AusgangsOrdner,Liste):
            drueber = os.path.dirname(ord)
            uebergeordneter = os.path.split(drueber)[1]
            
            if uebergeordneter != AusgangsOrdner:
                Liste.append(uebergeordneter)
                find_parents(drueber,AusgangsOrdner,Liste)
        
        
        Verzeichnis = []
        first_time = True
        
        for root, dirs, files in os.walk(path):

            if first_time:
                AusgangsOrdner = os.path.split(root)[1]
                first_time = False
                drueber = os.path.dirname(root)            
            
            Ordner = os.path.split(root)[1]
    
            for file in files:
                if file.endswith(".odt"):
                    if Ordner != AusgangsOrdner:
                        Liste = []
                        find_parents(root,AusgangsOrdner,Liste)
                        
                        pfad = os.path.join(root,file)
                        Liste.insert(0,Ordner)
                        Verzeichnis.append((file,pfad,Liste)) 
                    else:
                        Liste = [AusgangsOrdner]
                        pfad = os.path.join(root,file)
                        Verzeichnis.append((file,pfad,Liste)) 
        
        return Verzeichnis,AusgangsOrdner
    
    def setze_selektierte_zeile(self,ord):
        zeile = self.mb.Hauptfeld.getControl(ord)
        icon = zeile.getControl('icon')
        self.mb.selektierte_zeile = icon.Context.AccessibleContext


class Import_CheckBox_Listener(unohelper.Base, XItemListener):
    def __init__(self,mb,buttons,but_Auswahl):
        self.mb = mb
        self.buttons = buttons
        self.but_Auswahl = but_Auswahl
        
    # XItemListener    
    def itemStateChanged(self, ev):        
        # um sich nicht selbst abzuwaehlen
        if ev.Source.State == 0:
            ev.Source.State = 1
       
     
         

class Auswahl_CheckBox_Listener(unohelper.Base, XActionListener):
    def __init__(self,mb,model1,model2,contr_strukt):
        self.mb = mb
        self.model1 = model1
        self.model2 = model2
        self.contr_strukt = contr_strukt
        
    def actionPerformed(self,ev):
        imp_set = self.mb.settings_imp

        if ev.ActionCommand == 'struktur':
            if imp_set['ord_strukt'] == 0:
                imp_set['ord_strukt'] = 1
            else:
                imp_set['ord_strukt'] = 0                
        elif ev.Source.State == 0:
            ev.Source.State = 1
            return

        if ev.ActionCommand == 'model1':
            self.model2.State = False
            imp_set['imp_dat'] = '1'
            self.contr_strukt.Enable = False
        elif ev.ActionCommand == 'model2':
            self.model1.State = False
            imp_set['imp_dat'] = '0'
            self.contr_strukt.Enable = True
        else:
            pass
            #self.model1.State = False
            #import_datei.text = '0'
        self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)     
            

class Auswahl_Button_Listener(unohelper.Base, XActionListener):
    
    def __init__(self,mb,model1,model2):
        self.mb = mb
        self.model1 = model1
        self.model2 = model2
        
    def actionPerformed(self,ev):
        imp_set = self.mb.settings_imp 

        if ev.ActionCommand == 'Datei':
        
            Filepicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
            if imp_set['url_dat'] != None:
                Filepicker.setDisplayDirectory(imp_set['url_dat'])
            Filepicker.execute()
         
            if Filepicker.Files == '':
                return
             
            filepath = Filepicker.Files[0]
            self.model1.Label = uno.fileUrlToSystemPath(filepath)
            imp_set['url_dat'] = filepath
            
            self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  

        
        elif ev.ActionCommand == 'Ordner':
        
            Filepicker = createUnoService("com.sun.star.ui.dialogs.FolderPicker")
            if imp_set['url_ord'] != None:
                Filepicker.setDisplayDirectory(imp_set['url_ord'])
            Filepicker.execute()
         
            if Filepicker.Directory == '':
                return
             
            filepath = Filepicker.Directory
            self.model2.Label = uno.fileUrlToSystemPath(filepath)
            imp_set['url_ord'] = filepath
            
            self.mb.speicher_settings("import_settings.txt", self.mb.settings_imp)  


           
             
################ TOOLS ################################################################

# Handy function provided by hanya (from the OOo forums) to create a control, model.
def createControl(ctx,type,x,y,width,height,names,values):
   smgr = ctx.getServiceManager()
   ctrl = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % type,ctx)
   ctrl_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % type,ctx)
   ctrl_model.setPropertyValues(names,values)
   ctrl.setModel(ctrl_model)
   ctrl.setPosSize(x,y,width,height,15)
   return (ctrl, ctrl_model)
def createUnoService(serviceName):
  sm = uno.getComponentContext().ServiceManager
  return sm.createInstanceWithContext(serviceName, uno.getComponentContext())

