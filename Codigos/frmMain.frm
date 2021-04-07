VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Indexador - Editor de Particulas - WinterAO"
   ClientHeight    =   11865
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   791
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   824
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FraListaDe 
      BackColor       =   &H00000000&
      Caption         =   "Lista de Ataques"
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
      Left            =   9795
      TabIndex        =   27
      Top             =   9480
      Width           =   2535
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   7
         ItemData        =   "frmMain.frx":000C
         Left            =   120
         List            =   "frmMain.frx":000E
         TabIndex        =   28
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   10425
      Left            =   2520
      MousePointer    =   99  'Custom
      ScaleHeight     =   693
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   7200
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7560
      Picture         =   "frmMain.frx":0010
      ScaleHeight     =   855
      ScaleWidth      =   1455
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   14
      Top             =   10560
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.ListBox lstGrh 
      Appearance      =   0  'Flat
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
      Height          =   10320
      Index           =   0
      ItemData        =   "frmMain.frx":4158
      Left            =   120
      List            =   "frmMain.frx":415A
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   6
         ItemData        =   "frmMain.frx":415C
         Left            =   120
         List            =   "frmMain.frx":415E
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   5
         ItemData        =   "frmMain.frx":4160
         Left            =   120
         List            =   "frmMain.frx":4162
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   4
         ItemData        =   "frmMain.frx":4164
         Left            =   120
         List            =   "frmMain.frx":4166
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   3
         ItemData        =   "frmMain.frx":4168
         Left            =   120
         List            =   "frmMain.frx":416A
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   2
         ItemData        =   "frmMain.frx":416C
         Left            =   120
         List            =   "frmMain.frx":416E
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
      Begin VB.ListBox lstGrh 
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
         Height          =   1200
         Index           =   1
         ItemData        =   "frmMain.frx":4170
         Left            =   120
         List            =   "frmMain.frx":4172
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   11400
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
            Caption         =   "Recargar Escudos"
            Index           =   5
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar FX's"
            Index           =   6
         End
         Begin VB.Menu mnuRecargar 
            Caption         =   "Recargar Ataques"
            Index           =   7
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
            Caption         =   "...Graficos"
            Index           =   0
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Cabezas"
            Index           =   1
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Cascos"
            Index           =   2
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Personajes"
            Index           =   3
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Armas"
            Index           =   4
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Escudos"
            Index           =   5
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Fxs"
            Index           =   6
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Particulas"
            Index           =   7
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Colores"
            Index           =   8
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...GUI"
            Index           =   9
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "...Ataques"
            Index           =   10
         End
      End
      Begin VB.Menu mnuopDesindexar 
         Caption         =   "Desindexar"
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Graficos"
            Index           =   0
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Cabezas"
            Index           =   1
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Cascos"
            Index           =   2
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Cuerpos"
            Index           =   3
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Armas"
            Index           =   4
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Escudos"
            Index           =   5
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Ataques"
            Index           =   6
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Fxs"
            Index           =   7
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...GUI"
            Index           =   8
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Colores"
            Index           =   9
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "...Particulas"
            Index           =   10
         End
         Begin VB.Menu lin0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDesindexarTODO 
            Caption         =   "...TODO"
         End
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Index           =   0
      Begin VB.Menu mnuGrhSearch 
         Caption         =   "Buscar Grh"
      End
      Begin VB.Menu mnuBuscarGrh 
         Caption         =   "Buscar Grh's Libres"
         Index           =   0
      End
      Begin VB.Menu mnuBuscarGrhConsecutivo 
         Caption         =   "Buscar Grh's Consecutivos"
         Index           =   1
      End
      Begin VB.Menu lin1 
         Caption         =   "-"
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
         Index           =   0
      End
      Begin VB.Menu mnuAutoIndex 
         Caption         =   "Cabezas en fila"
         Index           =   1
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub lstGrh_Click(Index As Integer)
    Dim i As Byte
    
    Select Case Index
    
        Case 0 'Grh normal
            GrhSelect(0) = frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)
            For i = 1 To 3
                GrhSelect(i) = 0
            Next i
            Call GrhInfo
    
        Case 1 'Cabezas
            For i = 0 To 3
                GrhSelect(i) = 0
            Next i
        
        Case 2 'Cascos
            For i = 0 To 3
                GrhSelect(i) = 0
            Next i
        
        Case 3 'Cuerpos
            For i = 0 To 3
                GrhSelect(i) = BodyData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).Walk(i + 1).GrhIndex
            Next i
            
        Case 4 'Armas
            For i = 0 To 3
                GrhSelect(i) = WeaponAnimData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).WeaponWalk(i + 1).GrhIndex
            Next i
            
        Case 5 'Escudos
            For i = 0 To 3
                GrhSelect(i) = ShieldAnimData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).ShieldWalk(i + 1).GrhIndex
            Next i
            
        Case 6 'FX
            GrhSelect(0) = FxData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).Animacion
            For i = 1 To 3
                GrhSelect(i) = 0
            Next i
            
    End Select
    
    For i = 0 To 3
        Call InitGrh(GrhSelectInit(i), GrhSelect(i))
    Next i
End Sub

Private Sub mnuAdaptador_Click()
    frmAdaptador.Show
End Sub

Private Function GrhInfo()
'*********************************
'Autor: Lorwik
'Fecha: 10/05/2020
'Descripcion: Muestra la informacion del GRH seleccionado
'*********************************

    grhXTxt.Text = GrhData(GrhSelect(0)).sX
    grhYTxt.Text = GrhData(GrhSelect(0)).sY
    grhWidthTxt.Text = GrhData(GrhSelect(0)).pixelWidth
    grhHeightTxt.Text = GrhData(GrhSelect(0)).pixelHeight
    bmpTxt.Text = GrhData(GrhSelect(0)).FileNum
    
End Function

Private Sub mnuAutoIndex_Click(Index As Integer)
#If ModoVisor = 0 Then

    Select Case Index
    
        Case 0
            Call AutoIndex_Cuerpos

    End Select
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

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
        txtgrh = Left(frmMain.lstGrh(0).List(i - 1), Len(CStr(i)))
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

Private Sub mnuDesindexar_Click(Index As Integer)
#If ModoVisor = 0 Then

    Select Case Index

        Case 0 'Graficos
            Call DesindexarGraficos
        
        Case 1 'Cabezas
            Call DesindexarCabezas
         
        Case 2 'Cascos
            Call DesindexarCascos
         
        Case 3 'Cuerpos
            Call DesindexarCuerpos
        
        Case 4 'Armas
            Call DesindexarArmas
            
        Case 5 'Escudos
            Call DesindexarEscudos
            
        Case 6 'Ataques
            Call DesindexarAtaques
            
        Case 7 'Fxs
            Call DesindexarFxs
            
        Case 8 'GUI
        
        Case 9 'Colores
            Call DesindexarColores
        
        Case 10 'Particulas
    End Select
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

Private Sub mnuDesindexarTODO_Click()
    lblstatus.Caption = "Desindexando TODO..."

    Call DesindexarGraficos
    DoEvents
    
    Call DesindexarCuerpos
    DoEvents
    
    Call DesindexarCabezas
    DoEvents
    
    Call DesindexarCascos
    DoEvents
    
    Call DesindexarArmas
    DoEvents
    
    Call DesindexarEscudos
    DoEvents
    
    Call DesindexarAtaques
    DoEvents
    
    Call DesindexarFxs
    DoEvents
    
    lblstatus.Caption = "TODOS los Inits fueron desindexados."
End Sub

Private Sub mnuEditorParticulas_Click()
    frmParticleEditor.Show
End Sub

Private Sub mnuGrhSearch_Click()
On Error Resume Next

    Dim GrhSearch As Long
    Dim i As Long
    Dim j As Long
    
    'GrhSearch = InputBox("Ingrese el numero de GRH:")
    
   ' If IsNumeric(GrhSearch) = False Then Exit Sub
    
    'If GrhSearch < 0 Or GrhSearch > grhCount Then
        
    '    For i = 1 To grhCount
            
    '        If GrhData(i).NumFrames >= 1 And i = Archivo Then
                
    '            For j = 0 To lstGrh(0).ListCount - 1
    '                MsgBox "GRH encontrado."
    '                lstGrh(0).ListIndex = j
    '            Next j
                    
    '        Else
                
    '            MsgBox "GRH NO ENCONTRADO"
                
    '        End If
            
    '    Next i
            
    '    MsgBox "NO SE ENCONTRO EL GRH"
            
    'Else
        
    '    MsgBox "GRH INVALIDO"
        
    'End If
    
    MsgBox "Caracteristica en desarrollo"

End Sub

Private Sub mnuIndexar_Click(Index As Integer)
#If ModoVisor = 0 Then

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
            Call CargarBodys
            
        Case 4 'Weapons.ind
            Call IndexarArmas
            Call CargarArmas

        Case 5 'Shields.ind
            Call IndexarEscudos
            Call CargarEscudos
            
        Case 6 'Fx.ind
            Call IndexarFx
            Call CargarFX
            
        Case 7 'Particulas.ind
            Call IndexarParticulas
            
        Case 8 'Colores.ind
            Call IndexarColores
         
        Case 9 'GUI.ind
            Call IndexarGUI
            
        Case 10 'Ataques.ind
            Call IndexarAtaques
            Call CargarAtaques
    End Select
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

Private Sub mnuMinimap_Click()
#If ModoVisor = 0 Then
    Call GenerarMinimapa
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
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
            
        Case 7 'Cargar Ataques
            Call CargarAtaques
            
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
