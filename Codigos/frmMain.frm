VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Indexador - Editor de Particulas - WinterAO"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin Indexador.lvButtons_H LvBEditorDe 
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   9600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Editor de Particulas"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10800
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   1455
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   9400
      Width           =   1455
   End
   Begin VB.Frame grhFrame 
      BackColor       =   &H00000000&
      Caption         =   "Grh"
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
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   9480
      Width           =   7095
      Begin VB.TextBox grhXTxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhYTxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhHeightTxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhWidthTxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox bmpTxt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6240
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
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
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
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
         Height          =   195
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alto:"
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
         Height          =   195
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
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
         Height          =   195
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bmp:"
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
         Height          =   195
         Left            =   5760
         TabIndex        =   21
         Top             =   240
         Width           =   360
      End
   End
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
      ItemData        =   "frmMain.frx":4154
      Left            =   120
      List            =   "frmMain.frx":4156
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
         ItemData        =   "frmMain.frx":4158
         Left            =   120
         List            =   "frmMain.frx":415A
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
         ItemData        =   "frmMain.frx":415C
         Left            =   120
         List            =   "frmMain.frx":415E
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
         ItemData        =   "frmMain.frx":4160
         Left            =   120
         List            =   "frmMain.frx":4162
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
         ItemData        =   "frmMain.frx":4164
         Left            =   120
         List            =   "frmMain.frx":4166
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
         ItemData        =   "frmMain.frx":4168
         Left            =   120
         List            =   "frmMain.frx":416A
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
         ItemData        =   "frmMain.frx":416C
         Left            =   120
         List            =   "frmMain.frx":416E
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
      Top             =   10320
      Width           =   12120
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuMinimap 
         Caption         =   "Generar Minimapa.dat"
      End
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
      Begin VB.Menu mnubuscarerrores 
         Caption         =   "Buscar errores de indexacion"
      End
   End
   Begin VB.Menu mnuParticles 
      Caption         =   "&Particulas"
      Begin VB.Menu mnuEditorParticulas 
         Caption         =   "&Abrir Editor"
      End
   End
   Begin VB.Menu mnuAuto 
      Caption         =   "&Index. Automatica"
      Begin VB.Menu mnuAutoIndex 
         Caption         =   "Armaduras, Tunicas y Ropajes TIPO (6,6,5,5)"
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnusobre 
         Caption         =   "&Acerca de..."
      End
      Begin VB.Menu mnuComoIndexar 
         Caption         =   "&¿Como indexar?"
      End
      Begin VB.Menu mnusobreindexauto 
         Caption         =   "&Sobre la indexación automatica..."
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

    Dim Grafico As Long
    Grafico = InputBox("Numero del grafico")
    
    For i = 1 To grhCount
        If GrhData(i).FileNum = Grafico Then
            MsgBox "El grafico " & Grafico & " esta indexado en el Grh " & i & "."
            Exit Sub
        End If
    Next i
    
    'Si sale del For es por que esta indexado
    MsgBox "El grafico " & Grafico & " no esta indexado."
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
    Call GrhInfo
End Sub

Private Sub lstBodys_Click()
    GrhSeleccionado = BodyData(frmMain.lstBodys.List(frmMain.lstBodys.ListIndex)).Walk(1).GrhIndex
    Debug.Print GrhSeleccionado
End Sub

Private Sub LvBEditorDe_Click()
    frmParticleEditor.Show
End Sub

Private Sub mnuAdaptador_Click()
    frmAdaptador.Show
End Sub

Private Sub mnuAutoIndex_Click()
    Call AutoIndex_Cuerpos
End Sub

Private Function GrhInfo()
'*********************************
'Autor: Lorwik
'Fecha: 10/05/2020
'Descripcion: Muestra la informacion del GRH seleccionado
'*********************************

    grhXTxt.Text = GrhData(GrhSeleccionado).sX
    grhYTxt.Text = GrhData(GrhSeleccionado).sY
    grhWidthTxt.Text = GrhData(GrhSeleccionado).pixelWidth
    grhHeightTxt.Text = GrhData(GrhSeleccionado).pixelHeight
    bmpTxt.Text = GrhData(GrhSeleccionado).FileNum
    
End Function

Private Sub mnubuscarerrores_Click()
    Dim Datos As String
    Dim i As Long
    Dim j As Integer
    Dim Tim As Byte
    Tim = 0
    For i = 1 To grhCount
        If GrhData(i).NumFrames > 1 Then
            Tim = Tim + 1
            If Tim >= 150 Then
                Tim = 0
                lblstatus.Caption = "Procesando " & i & " grh"
                DoEvents
            End If
            
            For j = 1 To GrhData(i).NumFrames
                If GrhData(GrhData(i).Frames(j)).FileNum = 0 Then
                    Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Le falta el PNG " & GrhData(i).Frames(j) & vbCrLf
                ElseIf LenB(Dir(GraphicsDir & "\" & GrhData(GrhData(i).Frames(j)).FileNum & ".png", vbArchive)) = 0 Then
                    Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Le falta el PNG " & GrhData(GrhData(i).Frames(j)).FileNum & " (GRH" & GrhData(i).Frames(j) & ")" & vbCrLf
                End If
            Next
            
        ElseIf GrhData(i).NumFrames = 1 Then
        
            Tim = Tim + 1
            If Tim >= 150 Then
                Tim = 0
                lblstatus.Caption = "Procesando " & i & " grh"
                DoEvents
            End If
            If LenB(Dir(GraphicsDir & "\" & GrhData(i).FileNum & ".png", vbArchive)) = 0 Then
                Datos = Datos & "Grh" & i & " - Le falta el PNG " & GrhData(i).FileNum & vbCrLf
            End If
        End If
    Next
    frmresultado.txtResultado.Text = Datos
    frmresultado.Show
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
'****************************************
'Autor: Lorwik
'Fecha: 07/05/2020
'****************************************

    Dim Libres As Integer
    Libres = InputBox("Grh Libres Consecutivos")
    
    MsgBox BuscarConsecutivo(Libres)
End Sub

Private Sub mnuComoIndexar_Click()
    frmIndexHelp.Show
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
            Call ImportarDAT("Weapons")

        Case 5 'Shields.ind
            Call ImportarDAT("Shields")
            
        Case 6 'Fx.ind
            Call IndexarFx
            
    End Select
End Sub

Private Sub mnuMinimap_Click()
    Call GenerarMinimapa
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

Private Sub mnusobreindexauto_Click()
    MsgBox "La indexación automatica no es perfecta, por lo que puede generar errores. Asi que porfavor, no abuses y utilizala solo en graficos sencillos."
End Sub
