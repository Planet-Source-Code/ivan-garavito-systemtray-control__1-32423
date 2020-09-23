Attribute VB_Name = "mVKeys"
Public Const VK_ADD = &H6B
Public Const VK_ATTN = &HF6
Public Const VK_BACK = &H8
Public Const VK_CANCEL = &H3
Public Const VK_CAPITAL = &H14
Public Const VK_CLEAR = &HC
Public Const VK_CONTROL = &H11
Public Const VK_CRSEL = &HF7
Public Const VK_DECIMAL = &H6E
Public Const VK_DELETE = &H2E
Public Const VK_DIVIDE = &H6F
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
Public Const VK_EREOF = &HF9
Public Const VK_ESCAPE = &H1B
Public Const VK_EXECUTE = &H2B
Public Const VK_EXSEL = &HF8
Public Const VK_F1 = &H70
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F2 = &H71
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_F3 = &H72
Public Const VK_F5 = &H74
Public Const VK_F4 = &H73
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_HELP = &H2F
Public Const VK_HOME = &H24
Public Const VK_INSERT = &H2D
Public Const VK_LBUTTON = &H1
Public Const VK_LCONTROL = &HA2
Public Const VK_LEFT = &H25
Public Const VK_LMENU = &HA4
Public Const VK_LSHIFT = &HA0
Public Const VK_MBUTTON = &H4
Public Const VK_MENU = &H12
Public Const VK_MULTIPLY = &H6A
Public Const VK_NEXT = &H22
Public Const VK_NONAME = &HFC
Public Const VK_NUMLOCK = &H90
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_OEM_CLEAR = &HFE
Public Const VK_PA1 = &HFD
Public Const VK_PAUSE = &H13
Public Const VK_PLAY = &HFA
Public Const VK_PRINT = &H2A
Public Const VK_PRIOR = &H21
Public Const VK_PROCESSKEY = &HE5
Public Const VK_RBUTTON = &H2
Public Const VK_RCONTROL = &HA3
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_RMENU = &HA5
Public Const VK_RSHIFT = &HA1
Public Const VK_SCROLL = &H91
Public Const VK_SELECT = &H29
Public Const VK_SEPARATOR = &H6C
Public Const VK_SHIFT = &H10
Public Const VK_SNAPSHOT = &H2C
Public Const VK_SPACE = &H20
Public Const VK_SUBTRACT = &H6D
Public Const VK_TAB = &H9
Public Const VK_UP = &H26
Public Const VK_ZOOM = &HFB


Public Enum VirtualKeysConstants
  'Arrows keys
  ArrowLeft = VK_LEFT
  ArrowRight = VK_RIGHT
  ArrowUp = VK_UP
  ArrowDown = VK_DOWN

  ATTN = VK_ATTN
  
  Back = VK_BACK
  
  Cancel = VK_CANCEL
  
  Clear = VK_CLEAR
  
  Capital = VK_CAPITAL
  
  AddKey = VK_ADD
  DecimalKey = VK_DECIMAL
  DivideKey = VK_DIVIDE
  MultiplyKey = VK_MULTIPLY
  SubtractKey = VK_SUBTRACT

  DeleteKey = VK_DELETE
  EndKey = VK_END
  HomeKey = VK_HOME
  InsertKey = VK_INSERT
  
  EREof = VK_EREOF
  
  Execute = VK_EXECUTE
  
  CrSel = VK_CRSEL
  ExSel = VK_EXSEL
  
  'Series Fx
  F1 = VK_F1
  F2 = VK_F2
  F3 = VK_F3
  F4 = VK_F4
  F5 = VK_F5
  F6 = VK_F6
  F7 = VK_F7
  F8 = VK_F8
  F9 = VK_F9
  F10 = VK_F10
  F11 = VK_F11
  F12 = VK_F12
  F13 = VK_F13
  F14 = VK_F14
  F15 = VK_F15
  F16 = VK_F16
  F17 = VK_F17
  F18 = VK_F18
  F19 = VK_F19
  F20 = VK_F20
  F21 = VK_F21
  F22 = VK_F22
  F23 = VK_F23
  F24 = VK_F24
  
  Help = VK_HELP
  
  'Diferenciación de teclas
  LeftButton = VK_LBUTTON   'Mouse
  MidButton = VK_MBUTTON    'Mouse
  RightButton = VK_RBUTTON  'Mouse
  
  LeftControl = VK_LCONTROL
  LeftALT = VK_LMENU
  LeftMenu = VK_LMENU
  LeftShift = VK_LSHIFT
  
  RightControl = VK_RCONTROL
  RightMenu = VK_RMENU
  RightALT = VK_RMENU
  RightShift = VK_RSHIFT
  
  ALT = VK_MENU
  Control = VK_CONTROL
  Menu = VK_MENU
  Shift = VK_SHIFT
  
  NextKey = VK_NEXT
  NoName = VK_NONAME
  'NumPad
  NumLock = VK_NUMLOCK
  NumPad0 = VK_NUMPAD0
  NumPad1 = VK_NUMPAD1
  NumPad2 = VK_NUMPAD2
  NumPad3 = VK_NUMPAD3
  NumPad4 = VK_NUMPAD4
  NumPad5 = VK_NUMPAD5
  NumPad6 = VK_NUMPAD6
  NumPad7 = VK_NUMPAD7
  NumPad8 = VK_NUMPAD8
  NumPad9 = VK_NUMPAD9
  
  OEM_Clear = VK_OEM_CLEAR
  PA1 = VK_PA1
  
  Pause = VK_PAUSE
  Play = VK_PLAY
  PrintKey = VK_PRINT
  
  Prior = VK_PRIOR
  ProcessKey = VK_PROCESSKEY
  
  'Especial common keys
  Escape = VK_ESCAPE
  Enter = VK_RETURN
  ReturnKey = VK_RETURN
  SpaceKey = VK_SPACE
  TabKey = VK_TAB
  
  Scroll = VK_SCROLL
  SelectKey = VK_SELECT
  Separator = VK_SEPARATOR

  SnapShot = VK_SNAPSHOT
  
  Zoom = VK_ZOOM
End Enum


Public Declare Function GetAsyncKeyState Lib "user32" (vKey As Long) As Long

Public Declare Function GetKeyState Lib "user32" (nVirtKey As Long) As Long


Private Const TrueFunction = -32767



Public Function GetKeyAsync(ByVal vKey As VirtualKeysConstants) As Boolean
  GetKeyAsync = False
  If GetAsyncKeyState(vKey) = TrueFunction Then
    GetKeyAsync = True
  End If
End Function


Public Function GetKeyStatus(ByVal vKey As VirtualKeysConstants) As Boolean

  GetKeyStatus = CBool(GetKeyState(vKey))
  
End Function
