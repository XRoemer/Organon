# -*- coding: utf-8 -*-

from pprint import pformat

class Tools():
    
    def __init__(self,mb):
        self.mb = mb  
             

    def zeige_content_xml(self, ordinal=None, art='content.xml', pfad=None):
        try:
            import zipfile
            if pfad == None:
                pfad = os.path.join(self.mb.pfade['odts'], ordinal + '.odt')
                pfad_plain_txt = u'C:\\Users\\Homer\\Desktop\\Neuer Ordner\\{0}_{1}'.format(ordinal,art)
            else:
                pf, base = os.path.split(pfad) 
                pfad_plain_txt = os.path.join(pf , base.replace('odt','xml') )
                
            zipped = zipfile.ZipFile(pfad)
            content_xml = zipped.read(art).decode('utf-8')
            
            
            with codecs_open(pfad_plain_txt , "w","utf-8") as f:
                f.write(content_xml)
            
            import webbrowser
            webbrowser.open(pfad_plain_txt)
        except:
            log(inspect.stack,tb())

        
    def odt_ordner_entpacken(self,name):
        import zipfile
        pf = os.path.join(self.mb.pfade['odts'], name + '.odt')
        pf2 = os.path.join(self.mb.pfade['odts'], name + '2.odt')
        pf3 = os.path.join(self.mb.pfade['odts'], name + '2')
        
        if os.path.exists(pf2):
            os.remove(pf2)
        if os.path.exists(pf3):
            os.remove(pf3)
        
        import shutil
        shutil.copy( pf , pf2 )
        
        if not os.path.exists(pf3):
            os.makedirs(pf3)
        
        with zipfile.ZipFile(pf2, "r") as z:
            z.extractall(pf3)    
        
    def dict_or_OrdDict_to_formatted_str(self,OD, mode='dict', s="", indent=' '*4, level=0):
        # method taken from:
        # http://stackoverflow.com/questions/4301069/any-way-to-properly-pretty-print-ordered-dictionaries-in-python
        def is_number(s):
            try:
                float(s)
                return True
            except ValueError:
                return False
        def fstr(s):
            return s if is_number(s) else '"%s"'%s
        if mode != 'dict':
            kv_tpl = '("%s", %s)'
            ST = 'OrderedDict([\n'; END = '])'
        else:
            kv_tpl = '"%s": %s'
            ST = '{\n'; END = '}'
        for i,k in enumerate(OD.keys()):
            if type(OD[k]) in [dict, OrderedDict]:
                level += 1
                s += (level-1)*indent+kv_tpl%(k,ST+dict_or_OrdDict_to_formatted_str(OD[k], mode=mode, indent=indent, level=level)+(level-1)*indent+END)
                level -= 1
            else:
                s += level*indent+kv_tpl%(k,fstr(OD[k]))
            if i!=len(OD)-1:
                s += ","
            s += "\n"
        return s



    def xml2flat_dict(self,el):
        
        odict = {}
        
        for el in root.iter():
#             print('tag:{0: <20.20} text:{1: <24.24} tail:{2: <36.36}'.format(el.tag,el.text,el.tail))
#             print(el.attrib)
#             time.sleep(.2)
            tag = el.tag
            while tag in odict:
                tag = tag + '_1'
            odict[tag] = {}
            if el.text:
                odict[tag]['text'] = el.text
            if el.tail:
                odict[tag]['tail'] = el.tail
            if el.tag:
                odict[tag]['tag'] = el.tag
            if el.attrib:
                odict[tag]['attrib'] = dict(el.attrib)
                
        return odict
    
    
    def xml2dict(self, node, ordered_dict=False):
        
        if ordered_dict:
            from collections import OrderedDict
            d = OrderedDict()
        else:
            d = {}
        
        if node.text:
            d['text'] = node.text
        if node.tail:
            d['tail'] = node.tail
        if node.attrib:
            d['attrib'] = node.attrib
                        
        for e in node:
            key = e.tag
            
            while key in d:
                key = key + '_'
            d[key] = xml2dict(e,ordered_dict)

        return d
    
    
    def get_screen_size(self):
        import ctypes
        if sys.platform == 'win32':
            user32 = ctypes.windll.user32
            res = '{}x{}'.format(
                user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
        else:
            args = 'xrandr | grep "\*" | cut -d" " -f4'
            res = subprocess.check_output(args, shell=True).decode()
         
        return res.strip()
     
       
    def prettyprint(self,pfad,oObject,w=600):
        if self.mb.debug: log(inspect.stack) 
        
        imp = pformat(oObject,width=w)
        with codecs_open(pfad , "w",'utf-8') as f:
            f.write(imp)
            
            
    def zeitmesser(self,fkt,args=None):
        z = time.clock()
        
        if args:
            result = fkt(*args)
        else:
            result = fkt()
        print( round(time.clock()-z,3))
        return result 
        
        

    def get_attribs(self,obj,max_lvl,lvl=0):
        results = {}
        for key in dir(obj):
    
            try:
                value = getattr(obj, key)
                if 'callable' in str(type(value)):
                    continue
            except :
                #print(key)
                continue
    
            if key not in results:
                if type(value) in (
                                   type(None),
                                   type(True),
                                   type(1),
                                   type(.1),
                                   type('string'),
                                   type(()),
                                   type([]),
                                   type(b''),
                                   type(r''),
                                   type(u'')
                                   ):
                    results.update({key: value})
    
                elif lvl < max_lvl:
                    try:
                        results.update({key: get_attribs(value,max_lvl,lvl+1)})
                    except:
                        pass
    
        return results
    
        
    def find_differences(self,dict1,dict2):
        diff = []     
                            
        def findDiff2(d1, d2, path = []):
            for k in d1:
                try:
                    if k not in d2:
                        continue
#                         print (path, ":")
#                         print (k + " as key not in d2", "\n")
                    else:
                        if type(d1[k]) is dict and type(d2[k]) is dict:
                            #print(k,path)
                            if path == "":
                                #path = k
                                findDiff2(d1[k],d2[k],[k])
                            else:
                                #path = path + "->" + k
                                #print( path + "->" + k)
                                findDiff2(d1[k],d2[k], path + [k])
                            
                        else:
                            if d1[k] != d2[k]:
                                if path == '':
                                    path = [k]
                                #print('path:'+path+'#')
                                diff.append((path,k,d1[k],d2[k]))
                                #path = []
                except:
                    print(tb())
                    # Fehler produzieren, damit Schleife nur einmal
                    # durchlaufen wird
                    wer = wer1
                            
        findDiff2(dict1, dict2)
        
        return diff
    
    
    def diffenrences_als_dict(self,odiff):
        
        def update_odict(odict, v, d):
                
            try:
                if v[0] not in odict:
                    odict[v[0]] = {}
                
                if len(v) > 1:
                    update_odict(odict[v[0]], v[1:], d)
                    
                elif len(v) == 1:
                    odict[ v[0] ][ d[1] ] = [ d[2], d[3] ]
            
            except:
                print(tb())
                # Fehler produzieren, damit Schleife nur einmal
                # durchlaufen wird
                wer = wer
        
        
        def als_dict(diff):
        
            untersch = {}
            
            for d in diff:
                if len(d[0]) == 0:
                    untersch[d[1]] = [d[2], d[3]]
                else:
                    update_odict(untersch, d[0], d )
                    
            return untersch
        
        return als_dict(odiff)
    
    
    def find_diffs(self,ob1,ob2,lvl):
        attr1 = self.get_attribs(ob1,lvl)
        attr2 = self.get_attribs(ob2,lvl)
        diffs = self.find_differences(attr1,attr2)
        diff2 = self.diffenrences_als_dict(diffs)
        return diff2
        
        
    
    def get_config(self, key, node_name):
        if self.mb.debug: log(inspect.stack)
        
        provider = self.mb.ctx.ServiceManager.createInstanceWithContext(
                    'com.sun.star.configuration.ConfigurationProvider', self.mb.ctx)
        
        from com.sun.star.beans import PropertyValue
        
        prop = PropertyValue()
        prop.Name = 'nodepath'
        prop.Value = node_name
        
        try:
            access = provider.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (prop,))
            
            if access and (access.hasByName(key)):
                data = access.getPropertyValue(key)
                return data

        except Exception as e:
            log(inspect.stack,tb())
            return ''
    
    def get_ueberschriften(self, doc=None, ordinal=None, get_ordinal=True):
        if self.mb.debug: log(inspect.stack)
        
        try:
            if doc == None:
                doc = self.mb.doc
                
            sd = doc.createSearchDescriptor()
            sd.SearchStyles = True
            cur = doc.Text.createTextCursor()
             
            StyleFamilies = doc.StyleFamilies
            ParagraphStyles = StyleFamilies.getByName("ParagraphStyles")
             
            for p in ParagraphStyles.ElementNames:
                if p == 'Heading 1':
                    elem = ParagraphStyles.getByName(p)
                    display_name = elem.DisplayName
                    break
             
            ergebnisse = []
            
            sd.SearchString = display_name
            ergebnisse = doc.findAll(sd)
             
            ordnung = []
            
            if get_ordinal:
    
                for e in range(ergebnisse.Count):
                        
                    erg = ergebnisse.getByIndex(e)
                    cur.gotoRange(erg.Start,False)
                    ordinal_par = self.mb.class_Bereiche.get_ordinal(cur)
                    
                    if ordinal == None:
                        ordnung.append({
                                        'typ' : erg.ParaStyleName,
                                        'para' : erg,
                                        'ueberschrift' : erg.String,
                                        'ordinal' : ordinal_par,
                                        'index' : e
                                        })

                    else:
                        if ordinal_par == ordinal:
                            ordnung.append({
                                            'typ' : erg.ParaStyleName,
                                            'para' : erg,
                                            'ueberschrift' : erg.String,
                                            'ordinal' : ordinal_par,
                                            'index' : e
                                            })
            else:
                for e in range(ergebnisse.Count):
                    
                    erg = ergebnisse.getByIndex(e)
                    cur.gotoRange(erg.Start,False)
                    
                    ordnung.append({
                                    'typ' : erg.ParaStyleName,
                                    'para' : erg,
                                    'ueberschrift' : erg.String,
                                    'index' : e
                                    })

                        
            
        except:
            log(inspect.stack,tb())
            ordnung = []
            
        return ordnung    
    
    
    
    
    
    
    
    
    
    