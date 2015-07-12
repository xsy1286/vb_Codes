Attribute VB_Name = "WM_Msg"
Option Explicit
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'send的Message被取掉后才返回  sendmessage发送的消息，不再经过消息队列
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
'有Get消息才返回,没消息一直等下去
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
'send的Key的Message，得是获得焦距的。
Private Type POINTAPI
 x As Long
 y As Long
End Type
Public Type Msg
        hwnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
End Type

'******************************************************
'公共函数
'sendMouse(hwd As Long, Msg As Long, x1 As Long, y1 As Long) As Long  '竟对本进程外的窗口无效
'******************************************************
  Public Const WM_MOUSEFIRST = &H200
  Public Const WM_MOUSEMOVE = &H200
  Public Const WM_LBUTTONDOWN = &H201
  Public Const WM_LBUTTONUP = &H202
  Public Const WM_LBUTTONDBLCLK = &H203
  Public Const WM_RBUTTONDOWN = &H204
  Public Const WM_RBUTTONUP = &H205
  Public Const WM_RBUTTONDBLCLK = &H206
  Public Const WM_MBUTTONDOWN = &H207
  Public Const WM_MBUTTONUP = &H208
  Public Const WM_MBUTTONDBLCLK = &H209
  Public Const WM_MOUSELAST = &H209
      
  
  '   Window   Messages
  Public Const WM_NULL = &H0
  Public Const WM_CREATE = &H1
  Public Const WM_DESTROY = &H2
  Public Const WM_MOVE = &H3
  Public Const WM_SIZE = &H5
    
  Public Const WM_ACTIVATE = &H6
  '
  '     WM_ACTIVATE   state   values
    
  Public Const WA_INACTIVE = 0
  Public Const WA_ACTIVE = 1
  Public Const WA_CLICKACTIVE = 2
    
  Public Const WM_SETFOCUS = &H7
  Public Const WM_KILLFOCUS = &H8
  Public Const WM_ENABLE = &HA
  Public Const WM_SETREDRAW = &HB
  Public Const WM_SETTEXT = &HC
  Public Const WM_GETTEXT = &HD
  Public Const WM_GETTEXTLENGTH = &HE
  Public Const WM_PAINT = &HF
  Public Const WM_CLOSE = &H10
  Public Const WM_QUERYENDSESSION = &H11
  Public Const WM_QUIT = &H12
  Public Const WM_QUERYOPEN = &H13
  Public Const WM_ERASEBKGND = &H14
  Public Const WM_SYSCOLORCHANGE = &H15
  Public Const WM_ENDSESSION = &H16
  Public Const WM_SHOWWINDOW = &H18
  Public Const WM_WININICHANGE = &H1A
  Public Const WM_DEVMODECHANGE = &H1B
  Public Const WM_ACTIVATEAPP = &H1C
  Public Const WM_FONTCHANGE = &H1D
  Public Const WM_TIMECHANGE = &H1E
  Public Const WM_CANCELMODE = &H1F
  Public Const WM_SETCURSOR = &H20
  Public Const WM_MOUSEACTIVATE = &H21
  Public Const WM_CHILDACTIVATE = &H22
  Public Const WM_QUEUESYNC = &H23
    
  Public Const WM_GETMINMAXINFO = &H24
    
  Type MINMAXINFO
                  ptReserved   As POINTAPI
                  ptMaxSize   As POINTAPI
                  ptMaxPosition   As POINTAPI
                  ptMinTrackSize   As POINTAPI
                  ptMaxTrackSize   As POINTAPI
  End Type
    
  Public Const WM_PAINTICON = &H26
  Public Const WM_ICONERASEBKGND = &H27
  Public Const WM_NEXTDLGCTL = &H28
  Public Const WM_SPOOLERSTATUS = &H2A
  Public Const WM_DRAWITEM = &H2B
  Public Const WM_MEASUREITEM = &H2C
  Public Const WM_DELETEITEM = &H2D
  Public Const WM_VKEYTOITEM = &H2E
  Public Const WM_CHARTOITEM = &H2F
  Public Const WM_SETFONT = &H30
  Public Const WM_GETFONT = &H31
  Public Const WM_SETHOTKEY = &H32
  Public Const WM_GETHOTKEY = &H33
  Public Const WM_QUERYDRAGICON = &H37
  Public Const WM_COMPAREITEM = &H39
  Public Const WM_COMPACTING = &H41
  Public Const WM_OTHERWINDOWCREATED = &H42                                     '     no   longer   suported
  Public Const WM_OTHERWINDOWDESTROYED = &H43                                 '     no   longer   suported
  Public Const WM_COMMNOTIFY = &H44                                                     '     no   longer   suported
    
  '   notifications   passed   in   low   word   of   lParam   on   WM_COMMNOTIFY   messages
  Public Const CN_RECEIVE = &H1
  Public Const CN_TRANSMIT = &H2
  Public Const CN_EVENT = &H4
    
  Public Const WM_WINDOWPOSCHANGING = &H46
  Public Const WM_WINDOWPOSCHANGED = &H47
    
  Public Const WM_POWER = &H48
  '
  '     wParam   for   WM_POWER   window   message   and   DRV_POWER   driver   notification
    
  Public Const PWR_OK = 1
  Public Const PWR_FAIL = (-1)
  Public Const PWR_SUSPENDREQUEST = 1
  Public Const PWR_SUSPENDRESUME = 2
  Public Const PWR_CRITICALRESUME = 3
    
  Public Const WM_COPYDATA = &H4A
  Public Const WM_CANCELJOURNAL = &H4B
    
  Type COPYDATASTRUCT
                  dwData   As Long
                  cbData   As Long
                  lpData   As Long
  End Type
    
  Public Const WM_NCCREATE = &H81
  Public Const WM_NCDESTROY = &H82
  Public Const WM_NCCALCSIZE = &H83
  Public Const WM_NCHITTEST = &H84
  Public Const WM_NCPAINT = &H85
  Public Const WM_NCACTIVATE = &H86
  Public Const WM_GETDLGCODE = &H87
  Public Const WM_NCMOUSEMOVE = &HA0
  Public Const WM_NCLBUTTONDOWN = &HA1
  Public Const WM_NCLBUTTONUP = &HA2
  Public Const WM_NCLBUTTONDBLCLK = &HA3
  Public Const WM_NCRBUTTONDOWN = &HA4
  Public Const WM_NCRBUTTONUP = &HA5
  Public Const WM_NCRBUTTONDBLCLK = &HA6
  Public Const WM_NCMBUTTONDOWN = &HA7
  Public Const WM_NCMBUTTONUP = &HA8
  Public Const WM_NCMBUTTONDBLCLK = &HA9
    
  Public Const WM_KEYFIRST = &H100
  Public Const WM_KEYDOWN = &H100
  Public Const WM_KEYUP = &H101
  Public Const WM_CHAR = &H102
  Public Const WM_DEADCHAR = &H103
  Public Const WM_SYSKEYDOWN = &H104
  Public Const WM_SYSKEYUP = &H105
  Public Const WM_SYSCHAR = &H106
  Public Const WM_SYSDEADCHAR = &H107
  Public Const WM_KEYLAST = &H108
  Public Const WM_INITDIALOG = &H110
  Public Const WM_COMMAND = &H111
  Public Const WM_SYSCOMMAND = &H112
  Public Const WM_TIMER = &H113
  Public Const WM_HSCROLL = &H114
  Public Const WM_VSCROLL = &H115
  Public Const WM_INITMENU = &H116
  Public Const WM_INITMENUPOPUP = &H117
  Public Const WM_MENUSELECT = &H11F
  Public Const WM_MENUCHAR = &H120
  Public Const WM_ENTERIDLE = &H121
    
  Public Const WM_CTLCOLORMSGBOX = &H132
  Public Const WM_CTLCOLOREDIT = &H133
  Public Const WM_CTLCOLORLISTBOX = &H134
  Public Const WM_CTLCOLORBTN = &H135
  Public Const WM_CTLCOLORDLG = &H136
  Public Const WM_CTLCOLORSCROLLBAR = &H137
  Public Const WM_CTLCOLORSTATIC = &H138
    

  Public Const WM_PARENTNOTIFY = &H210
  Public Const WM_ENTERMENULOOP = &H211
  Public Const WM_EXITMENULOOP = &H212
  Public Const WM_MDICREATE = &H220
  Public Const WM_MDIDESTROY = &H221
  Public Const WM_MDIACTIVATE = &H222
  Public Const WM_MDIRESTORE = &H223
  Public Const WM_MDINEXT = &H224
  Public Const WM_MDIMAXIMIZE = &H225
  Public Const WM_MDITILE = &H226
  Public Const WM_MDICASCADE = &H227
  Public Const WM_MDIICONARRANGE = &H228
  Public Const WM_MDIGETACTIVE = &H229
  Public Const WM_MDISETMENU = &H230
  Public Const WM_DROPFILES = &H233
  Public Const WM_MDIREFRESHMENU = &H234
    
    
  Public Const WM_CUT = &H300
  Public Const WM_COPY = &H301
  Public Const WM_PASTE = &H302
  Public Const WM_CLEAR = &H303
  Public Const WM_UNDO = &H304
  Public Const WM_RENDERFORMAT = &H305
  Public Const WM_RENDERALLFORMATS = &H306
  Public Const WM_DESTROYCLIPBOARD = &H307
  Public Const WM_DRAWCLIPBOARD = &H308
  Public Const WM_PAINTCLIPBOARD = &H309
  Public Const WM_VSCROLLCLIPBOARD = &H30A
  Public Const WM_SIZECLIPBOARD = &H30B
  Public Const WM_ASKCBFORMATNAME = &H30C
  Public Const WM_CHANGECBCHAIN = &H30D
  Public Const WM_HSCROLLCLIPBOARD = &H30E
  Public Const WM_QUERYNEWPALETTE = &H30F
  Public Const WM_PALETTEISCHANGING = &H310
  Public Const WM_PALETTECHANGED = &H311
  Public Const WM_HOTKEY = &H312
    
  Public Const WM_PENWINFIRST = &H380
  Public Const WM_PENWINLAST = &H38F
    
  '   NOTE:   All   Message   Numbers   below   0x0400   are   RESERVED.
  '   Private   Window   Messages   Start   Here:
  Public Const WM_USER = &H400

Public Function sendMouse(hwd As Long, Msg As Long, x1 As Long, y1 As Long) As Long
  Dim r As Long
  r = PostMessage(hwd, Msg, Msg, ((x1 - 1) And &HFFFF) + ((y1 - 1) And &HFFFF) * &H10000) '坐标只是相对此hwd的客户区
  Debug.Print "PostMessage Result:" & CStr(r)
  sendMouse = r
End Function

