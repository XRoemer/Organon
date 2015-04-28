# -*- coding: utf-8 -*-

import unohelper
from threading import Thread

class Mausrad():
    
    def __init__(self,mb,pydevBrk):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        
        global pd
        pd = pydevBrk
        
        # Der Code fuer das Mausrad funktioniert nur unter Windows.
        if sys.platform != 'win32':
            self.mb.settings_proj['nutze_mausrad'] = False
 
        
    def starte_mausrad(self,treeview = False):
        if self.mb.debug: log(inspect.stack)
        
        try:
            try:               
                if not self.mb.settings_proj['nutze_mausrad']:
                    return
            except:
                # Die Property existiert noch nicht, deswegen setzen
                self.mb.settings_proj['nutze_mausrad'] = False
                return
            try:
                if treeview:
                    hauptfeld = self.mb.props[T.AB].Hauptfeld
                    container = hauptfeld.Context.Context
                    scrollLeiste = container.getControl('ScrollBar')
                else:
                    hauptfeld = self.mb.maus_fenster.getControl('Container_innen')
                    scrollLeiste = self.mb.maus_fenster.getControl('ScrollBar')                
                
            except Exception as e:
                log(inspect.stack,tb())
                return
            
            # scrollLeiste.Visible ist immer None (OO Bug?), daher wird hier das Maximum abgefragt
            if scrollLeiste.Maximum == 1:
                return
              
            RawInputReader = self.mb.class_RawInputReader(self.mb,pd,tb,log,inspect,hauptfeld,scrollLeiste)
    
            
            def sleeper(rir):     
                rir.start()
                 
                while self.mb.mausrad_an:
                    time.sleep(.1)
                    rir.pollEvents()
    
                rir.stop()
            
            
            rir = RawInputReader
            
            t = Thread(target=sleeper,args=(rir,))
            t.start()
        except:
            log(inspect.stack,tb())
        
    
    def registriere_Maus_Focus_Listener(self,cont):
        if self.mb.debug: log(inspect.stack)
        
        for c in cont.Controls:

            if 'UnoControlScrollBar' in str(c):
                
                test_listener = Window_Focus_Listener(self.mb,cont)
                listener_disposing = Window_Disposing_Listener(self.mb,cont,test_listener)
                
                cont.addFocusListener(test_listener)
                cont.addEventListener(listener_disposing)
        
                for c in cont.Controls:
                    c.addFocusListener(test_listener)
                    try:
                        for cb in c.Controls:
                            cb.addFocusListener(test_listener)
                    except:
                        pass
                    
                    

def formatiere(cont):
    return str(cont).split('XInterface)')[1].split('{implementationName')[0]



from com.sun.star.awt import XFocusListener
class Window_Focus_Listener (unohelper.Base,XFocusListener):
    
    def __init__(self,mb,cont):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.cont = cont
        self.mb.registrierte_maus_listener.append(formatiere(cont))
        self.first = True
    
    def focusGained(self,ev):
        
        cont = []
        self.get_parent(cont,ev.Source)
        
        if formatiere(cont[0]) in self.mb.registrierte_maus_listener:
            #print('focus gained',formatiere(cont[0]))
            self.mb.maus_fenster = cont[0]
            self.mb.mausrad_an = True
            self.mb.class_Mausrad.starte_mausrad()
            
    def focusLost(self, ev):
        
        if self.first:
            self.first = False
            return
        
        if formatiere(self.cont) in self.mb.registrierte_maus_listener:
            #print('focus lost',formatiere(self.cont))
            self.mb.mausrad_an = False
            self.mb.maus_fenster = None
            
    def get_parent(self,cont,cont2):
        
        try:
            if cont2.Context == None:
                cont.append(cont2)
            else:
                self.get_parent(cont,cont2.Context)
        except Exception as e:
            print(e)
      
    def disposing(self,arg):
        return
  
        
from com.sun.star.lang import XEventListener       
class Window_Disposing_Listener (unohelper.Base,XEventListener):
    
    def __init__(self,mb,cont,focus_listener):
        if mb.debug: log(inspect.stack)
        
        self.mb = mb
        self.cont = cont
        self.focus_listener = focus_listener
        
    def disposing(self,arg):
        if self.mb.debug: log(inspect.stack)
                
        self.mb.registrierte_maus_listener.remove(formatiere(self.cont))
        self.mb.mausrad_an = False
        self.mb.maus_fenster = None

