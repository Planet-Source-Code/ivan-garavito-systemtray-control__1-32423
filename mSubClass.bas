Attribute VB_Name = "mSubC"
Option Explicit

Public Const GWL_WNDPROC = (-4)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public cST As SysTray
Public hWndOld As Long


'****************************************************
'NAME     : SubClass()
'FUNCTION : Subclass a window.
Public Function SubClass(ByVal hwnd As Long) As Long
  'Si hWnd es 0, no tiene caso seguir.
  If hwnd = 0 Then Exit Function
  
  hWndOld = GetWindowLong(hwnd, GWL_WNDPROC)
  SubClass = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
End Function


Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim uEvent As WM_MOUSEEVENT
  If Not (uMsg = mSysTray.TRAY_CALLBACK) Then
    'Como el mensage no corresponde, dejamos que fluya
    'An not important message for us
    GoTo LetMessage
  End If

  'Guardamos el mensage en uEvent
  'Save the data
  CopyMemory uEvent, lParam, Len(uEvent)
  
  cST.RaiseEvents ByVal uEvent
  
  Exit Function

LetMessage:
  'Enviamos el mensaje al procedimiento original de la ventana.
  'Call the true window proc.
  WndProc = CallWindowProc(hWndOld, hwnd, uMsg, wParam, lParam)
End Function


Public Function UnSubClass(hwnd As Long) As Long
  UnSubClass = SetWindowLong(hwnd, GWL_WNDPROC, hWndOld)
End Function
