VERSION 5.00
Begin VB.Form frmAdaptador 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adaptador"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370
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
   ScaleHeight     =   5985
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtAdaptado 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Width           =   7815
   End
   Begin VB.CommandButton cmdAdaptar 
      Caption         =   "Adaptar"
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtOriginal 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label lblPrimeraPosición 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primera posición:"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblAdaptado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adaptado"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label lblGrhOriginal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grh Original"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmAdaptador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdaptar_Click()
    Dim Lineas As Variant, i As Long
    Dim resultado As String
    Dim Contador As Long
    
    'Debe ingresar almenos un numero
    If txtPos.Text = "" Then
        MsgBox ("El valor de la posición es nulo.")
        Exit Sub
    End If
        
    Contador = txtPos.Text
        
    'Separamos todas las lineas
    Lineas = Split(txtOriginal.Text, vbCrLf)
    
    For i = LBound(Lineas) To UBound(Lineas)
        'Recortamos y reemplazamos
        resultado = resultado & "Grh" & Contador & "=" & General_Field_Read(2, Lineas(i), 61) & vbCrLf
        
        'Aumentamos en 1 el contador
        Contador = Contador + 1
    Next i
    
    txtAdaptado.Text = resultado
End Sub

