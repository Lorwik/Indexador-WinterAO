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
      TabIndex        =   2
      Top             =   3600
      Width           =   5655
   End
   Begin Indexador.lvButtons_H LvBCerrar 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      caption         =   "Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmGeneradorMinimap.frx":0000
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.PictureBox Previewer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   5625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5655
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
