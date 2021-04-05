VERSION 5.00
Begin VB.Form frmGeneradorMinimap 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generador de Minimapa"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5640
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   5655
   End
   Begin Indexador.lvButtons_H LvBCerrar 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Cerrar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Indexador.ucImage ucImage 
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      bData           =   0   'False
      Filename        =   ""
      eScale          =   2
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
End
Attribute VB_Name = "frmGeneradorMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBCerrar_Click()
    'Si esta trabajando, no salimos
    If Trabajando Then
        MsgBox "Trabajando, espera"
        Exit Sub
    End If
    
    Unload Me
End Sub
