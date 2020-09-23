VERSION 5.00
Object = "*\ACtl_SysTray.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejemplo de SysTray"
   ClientHeight    =   3360
   ClientLeft      =   5385
   ClientTop       =   3195
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   119
   Begin Ctl_SysTray.SysTray SysTray1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   3493
      _ExtentY        =   661
      Icon            =   "Form1.frx":0000
      TrayToolTip     =   "Ejemplo de SysTray"
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   3060
      Width           =   1455
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   825
      Picture         =   "Form1.frx":031A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   825
      Picture         =   "Form1.frx":0464
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   825
      Picture         =   "Form1.frx":05AE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   825
      Picture         =   "Form1.frx":06F8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   825
      Picture         =   "Form1.frx":0842
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   825
      Picture         =   "Form1.frx":098C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   825
      Picture         =   "Form1.frx":0AD6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   495
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2535
      Left            =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione un ICONO:"
      Height          =   255
      Left            =   98
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu mnST 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnPic 
         Caption         =   "Rojo"
         Index           =   1
      End
      Begin VB.Menu mnPic 
         Caption         =   "Verde"
         Index           =   2
      End
      Begin VB.Menu mnPic 
         Caption         =   "Azul"
         Index           =   3
      End
      Begin VB.Menu mnPic 
         Caption         =   "Amarillo"
         Index           =   4
      End
      Begin VB.Menu mnPic 
         Caption         =   "Magenta"
         Index           =   5
      End
      Begin VB.Menu mnPic 
         Caption         =   "Fucia"
         Index           =   6
      End
      Begin VB.Menu mnPic 
         Caption         =   "Anaranjado"
         Index           =   7
      End
      Begin VB.Menu mnsST1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  SysTray1.AddTrayIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SysTray1.DeleteTrayIcon
End Sub

Private Sub mnPic_Click(Index As Integer)
  Set SysTray1.Icon = P(Index).Picture
  SysTray1.ModifyTrayIcon
End Sub

Private Sub mnSalir_Click()
  Unload Me
End Sub

Private Sub P_Click(Index As Integer)
  Set SysTray1.Icon = P(Index).Picture
  SysTray1.ModifyTrayIcon
End Sub

Private Sub SysTray1_DblClick(Button As MouseButtonConstants)
  If Button = vbLeftButton Then Me.Show
End Sub

Private Sub SysTray1_MouseDown(Button As MouseButtonConstants)
  If Button = vbRightButton Then
    PopupMenu mnST
  End If
End Sub
