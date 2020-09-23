Attribute VB_Name = "Otros"
Private Const MCI_CDADUIO_DOOR_OPEN = "CDAudio Door Open"
Private Const MCI_CDADUIO_DOOR_CLOSE = "CDAudio Door Close"


'CONSTANTS FOR SEE2BOO
Public Const s2bTT = 1
Public Const s2bTF = 2
Public Const s2bFT = 3
Public Const s2bFF = 4

Public Enum See2BooConstants
  sTwobTT = s2bTT
  sTwobTF = s2bTF
  sTwobFT = s2bFT
  sTwobFF = s2bFF
End Enum

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lDest As Long, lSource As Long, ByVal nBytes As Long)
Declare Sub CopyMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (lDest As Any, lSource As Any, ByVal nBytes As Long)

'More implementations for RtlMoveMemory
Declare Sub CopyToByteFromByte Lib "kernel32" Alias "RtlMoveMemory" (dest As Byte, Src As Byte, ByVal length&)
'Declare Sub CopyToStrFromInt Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest$, Src%, ByVal length&)
'Declare Sub CopyToStrFromLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest$, Src&, ByVal length&)
'Declare Sub CopyToIntFromStr Lib "kernel32" Alias "RtlMoveMemory" (Src%, ByVal dest$, ByVal Length&)
'Declare Sub CopyToLongFromStr Lib "kernel32" Alias "RtlMoveMemory" (Src&, ByVal dest$, ByVal Length&)
Declare Sub CopyToPtrFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal DestAddr&, ByVal SrcAddr&, ByVal length&)
Declare Sub CopyToByteFromPtr Lib "kernel32" Alias "RtlMoveMemory" (dest As Byte, ByVal SrcAddr&, ByVal length&)
Declare Sub CopyToLongFromPtr Lib "kernel32" Alias "RtlMoveMemory" (dest&, ByVal SrcAddr&, ByVal length&)
Declare Sub CopyToTypeFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByVal SrcAddr&, ByVal length&)
Declare Sub CopyToTypeFromType Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef Src As Any, ByVal length&)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Decreases a value
Public Function Dec(n As Long, Optional Step As Long = 1) As Long
  n = n - Step
  Dec = n
End Function

'Closes the cd tray
Public Function CDCloseDoor() As Long
  CDCloseDoor = mciSendString("Set " & MCI_CDADUIO_DOOR_CLOSE, vbNullString, 0&, 0&)
End Function

'Open the cd tray
Public Function CDOpenDoor() As Long
  CDOpenDoor = mciSendString("Set " & MCI_CDADUIO_DOOR_OPEN, vbNullString, 0&, 0&)
End Function

Public Function GetRedValue(ByVal xVal As Long) As Byte
  GetRedValue = (xVal \ &H10000) And &HFF
End Function


Public Function GetGreenValue(ByVal xVal As Long) As Byte
  GetGreenValue = (xVal \ &H100) And &HFF
End Function


Public Function GetBlueValue(ByVal xVal As Long) As Byte
  GetBlueValue = xVal And &HFF
End Function


Public Function HiByte(WordIn As Integer) As Byte
  CopyMemory ByVal HiByte, ByVal VarPtr(WordIn) + 1, 2
End Function


Public Function HiWord(LongIn As Long) As Integer
  CopyMemory ByVal HiWord, ByVal VarPtr(LongIn) + 2, 2
End Function


Public Function HiParam(ByVal lParam As Long) As Long
  HiParam = (lParam And &H7FFF0000) \ &H10000
End Function


Public Function Inc(n As Long, Optional Step As Long = 1) As Long
  n = n + Step
  Inc = n
End Function


Public Function IsNothing(xObject As Object) As Boolean
  On Error GoTo IsNothingObj
  
  IsNothing = (xObject Is Nothing)
'  IsNothing = False
'  p = xObject
  Exit Function
  
IsNothingObj:
  If Err.Number = 91 Then IsNothing = True
End Function


Public Function IsPicture(xPicture As Picture) As Boolean
  On Error GoTo ExistError
  
  IsPicture = True
  P = xPicture
  Exit Function

ExistError:
  If Err.Number = 91 Then IsPicture = False
End Function


Public Function LoByte(WordIn As Integer) As Byte
  CopyMemory ByVal LoByte, ByVal WordIn, 2
End Function


Public Function LoWord(LongIn As Long) As Integer
  CopyMemory ByVal LoWord, ByVal LongIn, 2
End Function


Public Function LoParam(ByVal lParam As Long) As Long
  LoParam = lParam And &H7FFF
End Function


Public Function MakeLong(ByVal LoW As Integer, ByVal HiW As Integer) As Long
  MakeLong = CLng(LoW)
  CopyMemoryAny ByVal VarPtr(MakeLong) + 2, HiW, 2
End Function


Public Function MakeWord(ByVal LoB As Byte, ByVal HiB As Byte) As Integer
  MakeWord = CInt(LoB)
  CopyMemoryAny ByVal VarPtr(MakeWord) + 1, HiB, 1
End Function


'Devuelve 4 posible valores
'Returns 1 of 4 values
Public Function See2Boo(ByVal b1 As Boolean, ByVal b2 As Boolean) As See2BooConstants
  If b1 And b2 Then
    See2Boo = sTwobTT
  ElseIf b1 And (Not b2) Then
    See2Boo = sTwobTF
  ElseIf (Not b1) And b2 Then
    See2Boo = sTwobFT
  ElseIf (Not b1) And (Not b2) Then
    See2Boo = sTwobFF
  End If
End Function


'Función creada el 11 de Mayo del 2001 para trabajar
'con dos condiciones y regresar tres posibles valores
Public Function TwinIf(ByVal Expression1 As Boolean, ByVal Expression2 As Boolean, _
                       True1, True2, Other)
  If Expression1 Then
    TwinIf = True1
  ElseIf Expression2 Then
    TwinIf = True2
  Else
    TwinIf = Other
  End If
End Function

