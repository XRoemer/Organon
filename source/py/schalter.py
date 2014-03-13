# -*- coding: utf-8 -*-
 
import unohelper
import sys

class Test(unohelper.Base):

    IMPLE_NAME = "com.open.office.roemers.schalter"
    SERVICE_NAMES = IMPLE_NAME,

    @classmethod
    def get_imple(klass):
        return klass, klass.IMPLE_NAME, klass.SERVICE_NAMES
    
    def __init__(self, ctx):
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        doc = desktop.getCurrentComponent() 
        RESOURCE_URL = "private:resource/dockingwindow/9809"
        
        layoutmgr = doc.getCurrentController().getFrame().LayoutManager
        if layoutmgr.isElementVisible(RESOURCE_URL):
            layoutmgr.hideElement(RESOURCE_URL)
        else:
            layoutmgr.showElement(RESOURCE_URL)
            

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(*Test.get_imple())



      
        
        
        

