VERSION 5.00
Begin VB.Form frmAdaptador 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adaptador"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
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
   ScaleHeight     =   5985
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Indexador.lvButtons_H cmdAdaptar 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Adaptar"
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
   Begin VB.TextBox txtPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtAdaptado 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3600
      Width           =   7815
   End
   Begin VB.TextBox txtOriginal 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
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
      ForeColor       =   &H00FFFFFF&
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
    Dim Resultado As String
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
        Resultado = Resultado & "Grh" & Contador & "=" & General_Field_Read(2, Lineas(i), 61) & vbCrLf
        
        'Aumentamos en 1 el contador
        Contador = Contador + 1
    Next i
    
    txtAdaptado.Text = Resultado
End Sub

