VERSION 5.00
Begin VB.Form frmAcercade 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acercade"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblBasadoEn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basado en el editor de particulas de ORE"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2925
   End
   Begin VB.Label lblDesarrolladoPor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desarrollado por LwKStudios"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2040
   End
   Begin VB.Label lblIndexadorWinterAO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indexador WinterAO 2020"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
