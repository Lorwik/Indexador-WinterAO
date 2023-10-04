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
   Begin VB.TextBox TxtNumAnimaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   5280
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtGrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin Indexador.lvButtons_H cmdAdaptar 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      BackColor       =   &H00C0C0C0&
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
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
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
   Begin VB.Label lblAnimaciones 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Animaciones"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grafico"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
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
    Dim Lineas As Variant, i As Long, j As Long
    Dim resultado As String
    Dim Contador As Long
    Dim lineaGrh As String
    Dim Fr As Integer
    Dim tmp As String
    
    'Debe ingresar almenos un numero
    If txtPos.Text = "" Then
        MsgBox ("El valor de la posición es nulo.")
        Exit Sub
    End If
    
    'Debe ingresar almenos un numero
    If txtGrafico.Text = "" Then
        MsgBox ("El valor del grafico es nulo.")
        Exit Sub
    End If
    

    Contador = txtPos.Text
        
    'Separamos todas las lineas
    Lineas = Split(txtOriginal.Text, vbCrLf)
    
    For i = LBound(Lineas) To UBound(Lineas)
    
        lineaGrh = General_Field_Read(2, Lineas(i), 61) 'Linea apartir del GrhNº
        Fr = General_Field_Read(1, lineaGrh, 45) 'Numero del frame
        
        If Fr = 1 Then
            'Recortamos y reemplazamos
            resultado = resultado & "Grh" & Contador & "=" & Fr & "-" & txtGrafico.Text & "-" & General_Field_Read(3, lineaGrh, 45) _
            & "-" & General_Field_Read(4, lineaGrh, 45) & "-" & General_Field_Read(5, lineaGrh, 45) & "-" & General_Field_Read(6, lineaGrh, 45) & vbCrLf
            
        Else '¿Es una animacion?
            
            If TxtNumAnimaciones.Text = "" Then
                MsgBox ("Hay animaciones y no se especifico el numero de estas.")
                Exit Sub
            End If
            
            tmp = "Grh" & Contador & "=" & Fr

            For j = LBound(Lineas) To UBound(Lineas) - TxtNumAnimaciones.Text
                tmp = tmp + "-" & j
            Next j
            
            resultado = resultado + tmp & "-" & General_Field_Read(Fr + 2, lineaGrh, 45)
            
        End If
        
        'Aumentamos en 1 el contador
        Contador = Contador + 1
    Next i
    
    txtAdaptado.Text = resultado
End Sub

