VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Indexador - Editor de Particulas - WinterAO"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox invpic 
      BackColor       =   &H00000000&
      Height          =   9285
      Left            =   2520
      ScaleHeight     =   615
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   7155
   End
   Begin VB.ListBox Grhs 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9300
      ItemData        =   "frmMain.frx":000C
      Left            =   120
      List            =   "frmMain.frx":000E
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame FraFX 
      BackColor       =   &H00000000&
      Caption         =   "Lista de FX"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   11
      Top             =   7800
      Width           =   2535
      Begin VB.ListBox lstFx 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":0010
         Left            =   120
         List            =   "frmMain.frx":0012
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FraShield 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Escudos"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   9
      Top             =   6240
      Width           =   2535
      Begin VB.ListBox lstEscudos 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":0014
         Left            =   120
         List            =   "frmMain.frx":0016
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FraWeapons 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Armas"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   7
      Top             =   4680
      Width           =   2535
      Begin VB.ListBox lstArmas 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":0018
         Left            =   120
         List            =   "frmMain.frx":001A
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraBodys 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Cuerpos"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
      Begin VB.ListBox lstBodys 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":001C
         Left            =   120
         List            =   "frmMain.frx":001E
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FraHelmets 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Cascos"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
      Begin VB.ListBox lstHelmets 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":0020
         Left            =   120
         List            =   "frmMain.frx":0022
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FraHead 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Cabezas"
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
      Height          =   1575
      Left            =   9800
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.ListBox lstHead 
         BackColor       =   &H00808080&
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
         Height          =   1230
         ItemData        =   "frmMain.frx":0024
         Left            =   120
         List            =   "frmMain.frx":0026
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iniciado!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   9480
      Width           =   12120
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuRGraficos 
         Caption         =   "Recargar"
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Graficos"
            Index           =   0
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Cuerpos"
            Index           =   1
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Cabezas"
            Index           =   2
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Cascos"
            Index           =   3
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Armas"
            Index           =   4
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar TODO"
            Index           =   9
         End
      End
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuIndex 
      Caption         =   "&Index"
      Begin VB.Menu mnuoptIndexar 
         Caption         =   "Indexar"
         Begin VB.Menu mnuIndexar 
            Caption         =   "Graficos.ind"
            Index           =   0
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Heads.ind"
            Index           =   1
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Helmets.ind"
            Index           =   2
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Body.dat"
            Index           =   3
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Weapons.dat"
            Index           =   4
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Shields.dat"
            Index           =   5
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Fx.dat"
            Index           =   6
         End
      End
      Begin VB.Menu mnuDesindexar 
         Caption         =   "Desindexar"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Index           =   0
      Begin VB.Menu mnuBuscarGrh 
         Caption         =   "Buscar Grh's Libres"
         Index           =   0
      End
      Begin VB.Menu mnuBuscarGrhConsecutivo 
         Caption         =   "Buscar Grh's Consecutivos"
         Index           =   1
      End
      Begin VB.Menu mnuAdaptador 
         Caption         =   "Adaptador de Grh"
      End
      Begin VB.Menu csmbuscarnoindex 
         Caption         =   "Buscar graficos NO indexados"
      End
   End
   Begin VB.Menu mnuParticles 
      Caption         =   "&Particulas"
      Begin VB.Menu mnuEditorParticulas 
         Caption         =   "&Abrir Editor"
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnusobre 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub csmbuscarnoindex_Click()
    Dim i As Long

    Dim grafico As Long
    grafico = InputBox("Numero del grafico")
    
    For i = 1 To grhCount
        If GrhData(i).FileNum = grafico Then
            MsgBox "El grafico " & grafico & " esta indexado."
            Exit Sub
        End If
    Next i
    
    'Si sale del For es por que esta indexado
    MsgBox "El grafico " & grafico & " no esta indexado."
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label2_Click()
    End
End Sub

Private Sub Grhs_Click()
    GrhSeleccionado = frmMain.Grhs.List(frmMain.Grhs.ListIndex)
End Sub

Private Sub mnuAdaptador_Click()
    frmAdaptador.Show
End Sub

Private Sub mnuBuscarGrh_Click(Index As Integer)
    Dim i As Long
    Dim txtgrh As String
    
    For i = 1 To grhCount
        txtgrh = Left(frmMain.Grhs.List(i - 1), Len(CStr(i)))
        If Not i = txtgrh Then
            MsgBox "Grh " & i & " esta libre."
            Exit Sub
        End If
    Next i
End Sub

Private Sub mnuBuscarGrhConsecutivo_Click(Index As Integer)
    On Error Resume Next
    Dim libres As Integer
    Dim i As Integer
    Dim Conta As Integer
    libres = InputBox("Grh Libres Consecutivos")
    If IsNumeric(libres) = False Then Exit Sub
    For i = 1 To grhCount
        If GrhData(i).NumFrames = 0 Then
            Conta = Conta + 1
            If Conta = libres Then
                MsgBox "Desde Grh" & i - (Conta - 1) & " hasta Grh" & i & " se encuentran libres."
                Exit Sub
            End If
        ElseIf Conta > 0 Then
            Conta = 0
        End If
    Next
    MsgBox "No se encontraron " & libres & " GRH Libres Consecutivos"
End Sub

Private Sub mnuEditorParticulas_Click()
    frmParticleEditor.Show
End Sub

Private Sub mnuIndexar_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Graficos.ind
            If GrhIniToGrhDataNew Then
                frmMain.lblstatus.Caption = "Graficos.ind Generado."
                LoadGrhData 'Recargamos los graficos
            Else
                frmMain.lblstatus.Caption = "Error en la indexación."
            End If
            
        Case 1 'Head.ind
            Call IndexarCabezas
            
        Case 2 'Helmets.ind
            Call IndexarCascos
            
        Case 3 'Body.ind
            Call IndexarCuerpos
            
        Case 4 'Weapons.ind
            Call IndexarArmas
        
        Case 5 'Shields.ind
            Call IndexarEscudos
            
        Case 6 'Fx.ind
            Call IndexarFx
            
    End Select
End Sub

Private Sub mnuRecargar_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Graficos.ind
            Call LoadGrhData
            
        Case 1  'Cuerpos
            Call CargarBodys
            
        Case 2 'Cabezas
            Call CargarCabezas
            
        Case 3 'Cascos
            Call CargarHelmets
            
        Case 4 ' Armas
            Call CargarArmas
            
        Case 5 ' Escudos
            Call CargarEscudos
            
        Case 6 'Cargar Fxs
            Call CargarFX
            
        Case 9 'Recargar todo
            Call CargarIndex
    
    End Select
End Sub

Private Sub mnusalir_Click()
    End
End Sub

Private Sub mnusobre_Click()
    frmAcercade.Show
End Sub
