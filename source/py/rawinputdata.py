# -*- coding: utf-8 -*-

# This module is an adjusted version of pymultimouse/rawinputreader.py
# http://pymultimouse.googlecode.com/svn-history/r2/trunk/rawinputreader.py

# http://aspn.activestate.com/ASPN/Cookbook/Python/Recipe/208699
# http://msdn.microsoft.com/en-us/library/ms645565(VS.85).aspx
# http://www.eventghost.org (the source code)

import sys
from ctypes import c_long,c_int,c_uint,c_char_p,c_ushort,Structure,Union
from ctypes import WINFUNCTYPE,windll,byref,sizeof,pointer,WinError
from ctypes.wintypes import DWORD,HWND,HANDLE,WPARAM,ULONG,LONG,UINT,BYTE


WNDPROC = WINFUNCTYPE(c_long, c_int, c_uint, c_int, c_int)

class WNDCLASS(Structure):
    _fields_ = [('style', c_uint),
                ('lpfnWndProc', WNDPROC),
                ('cbClsExtra', c_int),
                ('cbWndExtra', c_int),
                ('hInstance', c_int),
                ('hIcon', c_int),
                ('hCursor', c_int),
                ('hbrBackground', c_int),
                ('lpszMenuName', c_char_p),
                ('lpszClassName', c_char_p)]

class POINT(Structure):
    _fields_ = [('x', c_long),
                ('y', c_long)]
    
class MSG(Structure):
    _fields_ = [('hwnd', c_int),
                ('message', c_uint),
                ('wParam', c_int),
                ('lParam', c_int),
                ('time', c_int),
                ('pt', POINT)]

class RAWINPUTDEVICE(Structure):
    _fields_ = [
        ("usUsagePage", c_ushort),
        ("usUsage", c_ushort),
        ("dwFlags", DWORD),
        ("hwndTarget", HWND),
    ]

class RAWINPUTHEADER(Structure):
    _fields_ = [
        ("dwType", DWORD),
        ("dwSize", DWORD),
        ("hDevice", HANDLE),
        ("wParam", WPARAM),
    ] 
    
class RAWMOUSE(Structure):
    class _U1(Union):
        class _S2(Structure):
            _fields_ = [
                ("usButtonFlags", c_ushort),
                ("usButtonData", c_ushort),
            ]
        _fields_ = [
            ("ulButtons", ULONG),
            ("_s2", _S2),
        ]

    _fields_ = [
        ("usFlags", c_ushort),
        ("_u1", _U1),
        ("ulRawButtons", ULONG),
        ("lLastX", LONG),
        ("lLastY", LONG),
        ("ulExtraInformation", ULONG),
    ]
    _anonymous_ = ("_u1", )   

class RAWKEYBOARD(Structure):
    _fields_ = [
        ("MakeCode", c_ushort),
        ("Flags", c_ushort),
        ("Reserved", c_ushort),
        ("VKey", c_ushort),
        ("Message", UINT),
        ("ExtraInformation", ULONG),
    ]


class RAWHID(Structure):
    _fields_ = [
        ("dwSizeHid", DWORD),
        ("dwCount", DWORD),
        ("bRawData", BYTE),
    ]


class RAWINPUT(Structure):
    class _U1(Union):
        _fields_ = [
            ("mouse", RAWMOUSE),
            ("keyboard", RAWKEYBOARD),
            ("hid", RAWHID),
        ]

    _fields_ = [
        ("header", RAWINPUTHEADER),
        ("_u1", _U1),
        ("hDevice", HANDLE),
        ("wParam", WPARAM),
    ]
    _anonymous_ = ("_u1", )


def ErrorIfZero(HANDLE):
    if HANDLE == 0:
        # raise WinError hat einen Fehler erzeugt. 
        # raise WinError
        log(inspect.stack,extras='No Handle for Window // rawinputdata.py')
    else:
        return HANDLE


class RawInputReader():
 
    def __init__(self,mbx):
        if mbx.debug: log(inspect.stack)
        self.mb = mbx
        
        global mb
        mb = mbx
        
    def start(self): 
        if self.mb.debug: log(inspect.stack)
        
        WS_OVERLAPPEDWINDOW = (0     | \
                             12582912  | \
                             524288    | \
                             262144    | \
                             131072    | \
                             65536)
        CW_USEDEFAULT = -2147483648
         
        try:       
            CreateWindowEx = windll.user32.CreateWindowExA
            CreateWindowEx.argtypes = [c_int, c_char_p, c_char_p, c_int, 
                                       c_int, c_int, c_int, c_int, c_int, 
                                       c_int, c_int, c_int]
            CreateWindowEx.restype = ErrorIfZero
    
            wndclass = self.get_window()
            
            # Create Window
            HWND_MESSAGE = -3
            
            hwnd = CreateWindowEx(0,
                              wndclass.lpszClassName,
                              b"Python Window",
                              WS_OVERLAPPEDWINDOW,
                              CW_USEDEFAULT,
                              CW_USEDEFAULT,
                              CW_USEDEFAULT,
                              CW_USEDEFAULT,
                              0,
                              0,
                              wndclass.hInstance,
                              0)
            
            # Register for raw input
            Rid = (1 * RAWINPUTDEVICE)()
            self.Rid = Rid
            Rid[0].usUsagePage = 0x01
            Rid[0].usUsage = 0x02
            RIDEV_INPUTSINK = 0x00000100 # Get events even when not focused
            Rid[0].dwFlags = RIDEV_INPUTSINK
            Rid[0].hwndTarget = hwnd
     
            RegisterRawInputDevices = windll.user32.RegisterRawInputDevices
            RegisterRawInputDevices(Rid, 1, sizeof(RAWINPUTDEVICE))
            self.hwnd = hwnd
        
        except:
            log(inspect.stack,tb())


    def get_window(self):
        if self.mb.debug: log(inspect.stack)
        
        CS_VREDRAW = 1
        CS_HREDRAW = 2
        IDI_APPLICATION = 32512
        IDC_ARROW = 32512
        WHITE_BRUSH = 0
        
        # Define Window Class
        wndclass = WNDCLASS()
        self.wndclass = wndclass
        wndclass.style = CS_HREDRAW | CS_VREDRAW            
        wndclass.lpfnWndProc = WNDPROC(lambda h, m, w, l: self.WndProc(h, m, w, l))
        wndclass.cbClsExtra = wndclass.cbWndExtra = 0
        wndclass.hInstance = windll.kernel32.GetModuleHandleA(c_int(0))
        wndclass.hIcon = windll.user32.LoadIconA(c_int(0), c_int(IDI_APPLICATION))
        wndclass.hCursor = windll.user32.LoadCursorA(c_int(0), c_int(IDC_ARROW))
        wndclass.hbrBackground = windll.gdi32.GetStockObject(c_int(WHITE_BRUSH))
        wndclass.lpszMenuName = None
        wndclass.lpszClassName = b"MainWin"
         
        if not windll.user32.RegisterClassA(byref(wndclass)):
            raise WinError()
         
        return wndclass


    def pollEvents(self):
        #if self.mb.debug: self.log(inspect.stack)
                
        # Pump Messages
        msg = MSG()
        pMsg = pointer(msg)
        NULL = c_int(0)
     
        PM_REMOVE = 1
                
        while windll.user32.PeekMessageA(pMsg, self.hwnd, 0, 0, PM_REMOVE) != 0:
            windll.user32.DispatchMessageA(pMsg)
     
     
    def __del__(self):
        pass
 
 
    def stop(self):
        if self.mb.debug: log(inspect.stack)
         
        self.Rid[0].dwFlags = 0x00000001
        windll.user32.DestroyWindow(self.hwnd)

 
    def WndProc( self, hwnd, message, wParam, lParam):
        #if self.mb.debug: log(inspect.stack)
        
        try:            
            WM_INPUT = 255
            RI_MOUSE_WHEEL = 0x0400
            WM_DESTROY = 2
             
            if message == WM_DESTROY:
                windll.user32.PostQuitMessage(0)
                return 0
             
            elif message == WM_INPUT:
                GetRawInputData = windll.user32.GetRawInputData
                NULL = c_int(0)
                dwSize = c_uint()
                RID_INPUT = 0x10000003
                GetRawInputData(lParam, RID_INPUT, NULL, byref(dwSize), sizeof(RAWINPUTHEADER))
                 
                if dwSize.value == 40:
                    # Mouse
                    raw = RAWINPUT()
                     
                    if GetRawInputData(lParam, RID_INPUT, byref(raw), byref(dwSize), sizeof(RAWINPUTHEADER)) == dwSize.value:
                        RIM_TYPEMOUSE = 0x00000000
                        #RIM_TYPEKEYBOARD = 0x00000001
                         
                        if raw.header.dwType == RIM_TYPEMOUSE:
                             
                            if raw.mouse._u1._s2.usButtonFlags != RI_MOUSE_WHEEL:
                                return 0
                             
                            direction = (raw.mouse._u1._s2.usButtonData == 120)
                            
                            if direction:
                                v = 1
                            else:
                                v = -1
                            
                            # mb wird als globale Variable genutzt, da bei erneutem Oeffnen von Organon
                            # ansonsten die alte Instanz von self.mb benutzt wird. WARUM?
                            mb.class_Mausrad.bewege_scrollrad(v)
                                                          
            return windll.user32.DefWindowProcA(c_int(hwnd), c_int(message), c_int(wParam), c_int(lParam))
    
        except Exception as e:
            pass
            #print(e)
            #log(inspect.stack,tb())





