VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmCargando 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6105
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   1560
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   4560
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2752
      _Version        =   393217
      BackColor       =   4210752
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   0
      Left            =   9645
      ScaleHeight     =   0
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   7260
      Width           =   0
   End
   Begin VB.Image picload 
      Height          =   4500
      Left            =   0
      Picture         =   "frmCargando.frx":007E
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picload.Picture = LoadPicture(App.Path & "\Indexador.jpg")
End Sub
