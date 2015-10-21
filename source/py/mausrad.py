# -*- coding: utf-8 -*-

import unohelper
from ctypes import cdll,c_int
from threading import Thread

class Mausrad():
    
    def __init__(self,mb):
        if mb.debug: log(inspect.stack)
        self.mb = mb
        self.is_running = False
        
        # Der Code fuer das Mausrad funktioniert nur unter Windows.
        if sys.platform.lower() not in['win32','linux','linux2']:
            self.mb.settings_orga['mausrad'] = False
 
        
    def starte_mausrad(self,called_from_treeview = False):
        if self.mb.debug: log(inspect.stack)

        try:
            try:               
                if not self.mb.settings_orga['mausrad']:
                    return
            except:
                # Die Property existiert noch nicht, deswegen setzen
                self.mb.settings_orga['mausrad'] = False
                return
            try:
                if called_from_treeview:
                    self.hauptfeld = self.mb.props[T.AB].Hauptfeld
                    container = self.hauptfeld.Context.Context
                    self.scrollLeiste = container.getControl('ScrollBar')
                else:
                    self.hauptfeld = self.mb.maus_fenster.getControl('Container_innen')
                    self.scrollLeiste = self.mb.maus_fenster.getControl('ScrollBar')                
                
            except Exception as e:
                log(inspect.stack,tb())
                return
            
            # scrollLeiste.Visible ist immer None (Office Bug?), daher wird hier das Maximum abgefragt
            if self.scrollLeiste.Maximum == 1:
                return
            
            
            if sys.platform == 'win32':
                self.get_mausrad_windows()
            elif sys.platform.lower() in['linux','linux2']:
                self.get_mausrad_linux()
            
        except:
            log(inspect.stack,tb())
            return
            
            
            
    def get_mausrad_windows(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:            
            def sleeper(rir):     
                rir.start()
                                
                while self.mb.mausrad_an:
                    ev = rir.pollEvents()
                    time.sleep(.01)
                    
                rir.stop()
            
            rir = self.mb.class_RawInputReader

            t = Thread(target=sleeper,args=(rir,))
            t.start()
        except:
            log(inspect.stack,tb())
            
    
    def get_mausrad_linux(self): 
        if self.mb.debug: log(inspect.stack)
        
        try:
            # abfangen, falls schon ein Thread beschaeftigt ist
            if self.is_running:
                return
            
            path = os.path.join(self.mb.path_to_extension,'libs','libGetWheel.so')
            mydll = cdll.LoadLibrary(path)
    
            f = mydll.grab_wheel
            f.restype = c_int
            
            
            def sleeper(f,mb):     

                    while mb.mausrad_an:
                        button = f()
                        # da libGetWheel eine blockierende Funktion ist, kann sich mausrad_an geaendert haben,
                        # waehrend libGetWheel noch auf ein Scrollevent wartet.
                        # Das erste Scrollevent nach Focuswechsel wird dadurch unterschlagen.
                        if mb.mausrad_an:
                            if button == 4:
                                self.bewege_scrollrad(1)
                            else:
                                self.bewege_scrollrad(-1)

                    self.is_running = False         
                    
            t = Thread(target=sleeper,args=(f,self.mb))
            t.start()
            
            self.is_running = True
            
        except:
            log(inspect.stack,tb())
    
    
    def bewege_scrollrad(self,richtung):
       
        v = richtung * 20
        
        hoehe_leiste = self.scrollLeiste.Model.VisibleSize
        maximum = self.scrollLeiste.Maximum
           
        y = self.hauptfeld.PosSize.Y + v
        if y > 0:
            y = 0
        
        if -y > maximum - hoehe_leiste:
            y = -(maximum - hoehe_leiste - 1)
        
        
        self.hauptfeld.setPosSize(0, y ,0,0,2)
        self.scrollLeiste.Model.ScrollValue = -y
        
    
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
        
        if self.mb.mausrad_an:
            return
        
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

