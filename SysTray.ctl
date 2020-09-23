VERSION 5.00
Begin VB.UserControl SysTray 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "SysTray.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   135
   ToolboxBitmap   =   "SysTray.ctx":0471
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "06:44 p.m."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      X1              =   3
      X2              =   3
      Y1              =   2
      Y2              =   25
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000010&
      X1              =   2
      X2              =   2
      Y1              =   3
      Y2              =   24
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   131
      X2              =   131
      Y1              =   3
      Y2              =   25
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   6
      X2              =   131
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   6
      X2              =   6
      Y1              =   3
      Y2              =   24
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   6
      X2              =   131
      Y1              =   3
      Y2              =   3
   End
   Begin VB.Image imgPrev 
      Height          =   240
      Left            =   645
      Stretch         =   -1  'True
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Constants de tamaño en tiempo de diseño
'Size constants at design time
Private Const d_Height = 375
Private Const d_Width = 1980

'My events
Event Click(Button As MouseButtonConstants)
Attribute Click.VB_Description = "Se activa cuando el usuario hace clic sobre el icono del SystemTray."
Event DblClick(Button As MouseButtonConstants)
Attribute DblClick.VB_Description = "Se activa cuando el usuario hace doble clic en el icono del SystemTray."
Event MouseDown(Button As MouseButtonConstants)
Attribute MouseDown.VB_Description = "Se activa cuando el usuario mantiene presionado alguno de los botones del mouse sobre el icono del SystemTray."
Event MouseMove(Button As MouseButtonConstants)
Attribute MouseMove.VB_Description = "Se activa cuando el usuario mueve el mouse sobre el icono en el SystemTray."
Event MouseUp(Button As MouseButtonConstants)
Attribute MouseUp.VB_Description = "Se activa cuando el usuario libera uno de los botones sobre el icono en el SystemTray."

Dim LastButton As MouseButtonConstants
Dim MouseDown As Boolean

Private Running As Boolean

Dim p_Form As Form

Private Const cID = vbNull

Dim p_Icon As Picture
Dim p_TrayToolTip As String


Friend Property Get FormParent() As Form
  Set FormParent = p_Form
End Property

Friend Property Let FormParent(xFrm As Form)
  Set p_Form = xFrm
End Property

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Devuelve o establece el icono que se presentará en el SystemTray."
  Set Icon = p_Icon
End Property

Public Property Set Icon(xIcon As Picture)
  Set p_Icon = xIcon
  
  Set imgPrev.Picture = p_Icon
  
  If Running Then
    Me.ModifyTrayIcon
  End If
  
  PropertyChanged "Icon"
End Property

Public Property Get TrayToolTip() As String
Attribute TrayToolTip.VB_Description = "Devuelve o establece el mensaje que se muestra al colocar el mouse sobre el icono del SystemTray."
  TrayToolTip = p_TrayToolTip
End Property

Public Property Let TrayToolTip(xToolTip As String)
  p_TrayToolTip = xToolTip
  
  If Running Then
    Me.ModifyTrayIcon
  End If
  
  PropertyChanged "TrayToolTip"
End Property

Public Function About() As Long
Attribute About.VB_Description = "Muestra una ventana Acerca de..."
Attribute About.VB_UserMemId = -552
  frmAbout.Show vbModal
End Function

Public Function AddTrayIcon() As Long
Attribute AddTrayIcon.VB_Description = "Agrega un icono al SystemTray de Windows"
  AddTrayIcon = mSysTray.AddIcon(Me.FormParent.hwnd, cID, Icon.Handle, TrayToolTip)
  mSubC.SubClass Me.FormParent.hwnd
End Function

Public Function DeleteTrayIcon() As Long
Attribute DeleteTrayIcon.VB_Description = "Elimina el icono del SystemTray de Windows."
  mSubC.UnSubClass Me.FormParent.hwnd
  DeleteTrayIcon = mSysTray.DeleteIcon(Me.FormParent.hwnd, cID)
End Function

Public Function ModifyTrayIcon() As Long
Attribute ModifyTrayIcon.VB_Description = "Modifica el icono del SystemTray."
  ModifyTrayIcon = mSysTray.ModifyIcon(FormParent.hwnd, cID, Icon.Handle, TrayToolTip)
End Function

'Finds the form parent
Private Function FindFormParent() As Form
Dim xObj As Object

  'Asignamos el objeto contenedor actual
  Set xObj = UserControl.Parent

  'Mientras el tipo de objeto contenedor no sea
  'una ventana, asignamos el contenedor de este
  'objeto contenedor
  Do While Not (TypeOf xObj Is Form)
    Set xObj = xObj.Parent
  Loop
  
  'Copiamos el objeto
  Set FindFormParent = xObj
End Function

'Design time or Run time
Private Function ModoDR(ByVal b As Boolean) As Long
  'Checa el modo en que se encuentra el proyecto
  'Valores:
  ' True : Modo de ejecución.
  ' False: Modo de diseño.
  If b Then
    Set mSubC.cST = Me
  Else
    Set mSubC.cST = Nothing
  End If
  Running = b
End Function


'This function raises the mouse events
Friend Function RaiseEvents(ByVal uEvent As Long) As Long
Dim uE As WM_MOUSEEVENT
Dim CurrentButton As MouseButtonConstants
Dim CurrentShift As ShiftConstants

  uE = uEvent
  
  'Mouse button
  If (uE = LeftButtonDblClic) Or _
     (uE = LeftButtonDown) Or _
     (uE = LeftButtonUp) Then
    CurrentButton = vbLeftButton
  ElseIf (uE = RightButtonDown) Or _
         (uE = RightButtonUp) Or _
         (uE = RightButtonDblClic) Then
    CurrentButton = vbRightButton
  ElseIf (uE = MidButtonDown) Or _
         (uE = MidButtonUp) Or _
         (uE = MidButtonDblClic) Then
    CurrentButton = vbMiddleButton
  Else: CurrentButton = 0
  End If
    
  'Mouse event
  If uE = MouseMove Then
    RaiseEvent MouseMove(CurrentButton)
  ElseIf (uE = LeftButtonDown) Or _
         (uE = RightButtonDown) Or _
         (uE = MidButtonDown) Then
    RaiseEvent MouseDown(CurrentButton)
  ElseIf (uE = LeftButtonUp) Or _
         (uE = MidButtonUp) Or _
         (uE = RightButtonUp) Then
    RaiseEvent MouseUp(CurrentButton)
  ElseIf (uE = LeftButtonDblClic) Or _
         (uE = MidButtonDblClic) Or _
         (uE = RightButtonDblClic) Then
    RaiseEvent DblClick(CurrentButton)
  End If
  
  If (LastButton And CurrentButton) And MouseDown Then
    'Activamos el evento al usuario
    'Raises click event
    RaiseEvent Click(CurrentButton)
    'Desactivamos la bandera
    'Deactive the mousedown flag
    MouseDown = False
  End If
  
  'Si el evento corresponde a que algún botón este
  'presionado, validamos lo siguiente
  If (uEvent = LeftButtonDown) Or (uEvent = MidButtonDown) _
  Or (uEvent = RightButtonDown) Then
    'Se activa la bandera del evento MouseDown
    MouseDown = True
  ElseIf Not (uEvent = MouseMove) Then
    'Si no, mientras el evento no sea el movimiento
    'del mouse, desactivamos la bandera del evento
    'MouseDown
    MouseDown = False
  End If
  
  'Guardamos el botón actual
  'Save the current button
  LastButton = CurrentButton
End Function

Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_InitProperties()
  'Initial properties
  FormParent = FindFormParent
  Set Me.Icon = FormParent.Icon
  Me.TrayToolTip = FormParent.Caption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ModoDR Ambient.UserMode
  
  FormParent = FindFormParent
  
  Set Me.Icon = PropBag.ReadProperty("Icon", Me.FormParent.Icon)
  Me.TrayToolTip = PropBag.ReadProperty("TrayToolTip")
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = d_Height
  UserControl.Width = d_Width
End Sub

Private Sub UserControl_Terminate()
  Running = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Icon", Me.Icon
  PropBag.WriteProperty "TrayToolTip", Me.TrayToolTip
End Sub
