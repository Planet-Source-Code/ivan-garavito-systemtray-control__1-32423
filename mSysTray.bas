Attribute VB_Name = "mSysTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const WM_USER = &H400

'Mensaje que envía Windows a la ventana, indicando el evento
'del mouse sobre el icono en el System Tray. Definido en:
'  uCallBackMessage
'Message to receive the system tray notificacions
Public Const TRAY_CALLBACK = (WM_USER + 101)
'Mensajes - messages
Public Const NIM_ADD = &H0     'Agrega un icono
Public Const NIM_MODIFY = &H1  'Modifica un icono
Public Const NIM_DELETE = &H2  'Elimina un icono
'Banderas - Flags
Public Const NIF_MESSAGE = &H1 '
Public Const NIF_ICON = &H2    'Establece que se presentará un icono
Public Const NIF_TIP = &H4     'Establece que el icono muestra ToolTip
'Establece todas las banderas - Every style
Public Const NIF_ALL = NIF_ICON Or NIF_TIP Or NIF_MESSAGE

'Mensajes que envía Windows para indicar el evento que se
'activó
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200     'El mouse se mueve
Public Const WM_LBUTTONDOWN = &H201   'BI presionado
Public Const WM_LBUTTONUP = &H202     'BI soltado
Public Const WM_LBUTTONDBLCLK = &H203 'Doble clic con BI
Public Const WM_RBUTTONDOWN = &H204   'BD presionado
Public Const WM_RBUTTONUP = &H205     'BD soltado
Public Const WM_RBUTTONDBLCLK = &H206 'Doble clic con BD
Public Const WM_MBUTTONDOWN = &H207   'BM presionado
Public Const WM_MBUTTONUP = &H208     'BM soltado
Public Const WM_MBUTTONDBLCLK = &H209 'Doble clic con BM
Public Const WM_MOUSELAST = &H209

Public Enum WM_MOUSEEVENT
  MouseFirst = WM_MOUSEFIRST
  MouseMove = WM_MOUSEMOVE
  LeftButtonDown = WM_LBUTTONDOWN
  LeftButtonUp = WM_LBUTTONUP
  LeftButtonDblClic = WM_LBUTTONDBLCLK
  RightButtonDown = WM_RBUTTONDOWN
  RightButtonUp = WM_RBUTTONUP
  RightButtonDblClic = WM_RBUTTONDBLCLK
  MidButtonDown = WM_MBUTTONDOWN
  MidButtonUp = WM_MBUTTONUP
  MidButtonDblClic = WM_MBUTTONDBLCLK
  MouseLast = WM_MOUSELAST
End Enum

Public Type NOTIFYICONDATA
  cbSize As Long  'Tamaño de la estructura - Size
  hwnd As Long    'Ventana que recibe los mensajes - Window
  uID As Long     'ID del icono - Icon ID
  uFlags As Long  'Banderas - Flags
  uCallbackMessage As Long  'No. de mensaje que se utilizará - Windows message
  hIcon As Long   'Controlador del icono  - Handle icon
  szTip As String * 64  'Texto que se mostrará - Tooltip Text
End Type


Public Function AddIcon(hwnd As Long, ID As Long, hIcon As Long, ToolTip As String) As Long
Dim Tray As NOTIFYICONDATA

  'Datos - Data
  Tray.hwnd = hwnd
  Tray.uID = ID
  Tray.hIcon = hIcon
  Tray.szTip = ToolTip & vbNullChar
  Tray.uCallbackMessage = TRAY_CALLBACK
  Tray.uFlags = NIF_ALL
  Tray.cbSize = Len(Tray)
  
  AddIcon = Shell_NotifyIcon(NIM_ADD, Tray)
End Function

Public Function DeleteIcon(hwnd As Long, ID As Long) As Long
Dim Tray As NOTIFYICONDATA

  Tray.hwnd = hwnd
  Tray.uID = ID
  Tray.uFlags = 0&
  'Devuelve el tamaño de la estructura a la función
  'Return the size of Tray
  Tray.cbSize = Len(Tray)
  
  DeleteIcon = Shell_NotifyIcon(NIM_DELETE, Tray)
End Function

Public Function ModifyIcon(hwnd As Long, ID As Long, hIcon As Long, ToolTip As String) As Long
Dim Tray As NOTIFYICONDATA

  'Establece los cambios
  'Set the changes
  Tray.hwnd = hwnd
  Tray.uID = ID
  Tray.hIcon = hIcon
  Tray.szTip = ToolTip & vbNullChar
  Tray.uCallbackMessage = TRAY_CALLBACK
  Tray.uFlags = NIF_ALL
  Tray.cbSize = Len(Tray)
  
  ModifyIcon = Shell_NotifyIcon(NIM_MODIFY, Tray)
End Function

