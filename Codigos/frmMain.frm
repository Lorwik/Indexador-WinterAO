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
   ScaleHeight     =   718
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   824
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   9345
      Left            =   2520
      MousePointer    =   99  'Custom
      ScaleHeight     =   621
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
      Left            =   10800
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   1455
      TabIndex        =   25
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
      TabIndex        =   14
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
      Index           =   0
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
      Begin VB.ListBox lstGrh 
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
         Index           =   6
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
      Begin VB.ListBox lstGrh 
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
         Index           =   5
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
      Begin VB.ListBox lstGrh 
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
         Index           =   4
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
      Begin VB.ListBox lstGrh 
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
         Index           =   3
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
      Begin VB.ListBox lstGrh 
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
         Index           =   2
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
      Begin VB.ListBox lstGrh 
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
         Index           =   1
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
            Caption         =   "Cabezas.ind"
            Index           =   1
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Cascos.ind"
            Index           =   2
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Personajes.ini"
            Index           =   3
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Armas.ind"
            Index           =   4
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Escudos.ind"
            Index           =   5
         End
         Begin VB.Menu mnuIndexar 
            Caption         =   "Fxs.ind"
            Index           =   6
         End
      End
      Begin VB.Menu mnuopDesindexar 
         Caption         =   "Desindexar"
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Graficos.ini"
            Index           =   0
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Head.ind"
            Index           =   1
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Helmet.ind"
            Index           =   2
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Cuerpos.ini"
            Index           =   3
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Armas.ind"
            Index           =   4
         End
         Begin VB.Menu mnuDesindexar 
            Caption         =   "Escudos.dat"
            Index           =   5
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

Private Sub lstGrh_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Grh normal
            GrhSelect = frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)
            Call GrhInfo
    
        Case 1 'Cabezas
        
        Case 2 'Cascos
        
        Case 3 'Cuerpos
            GrhSelect = BodyData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).Walk(1).GrhIndex
            
        Case 4 'Armas
            GrhSelect = WeaponAnimData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).WeaponWalk(1).GrhIndex
            
        Case 5 'Escudos
            GrhSelect = ShieldAnimData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).ShieldWalk(1).GrhIndex
            
        Case 6 'FX
            GrhSelect = FxData(frmMain.lstGrh(Index).List(frmMain.lstGrh(Index).ListIndex)).Animacion
            
    End Select
    
    If GrhSelect > 0 Then Call InitGrh(GrhSelectInit, GrhSelect)
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

    grhXTxt.Text = GrhData(GrhSelect).sX
    grhYTxt.Text = GrhData(GrhSelect).sY
    grhWidthTxt.Text = GrhData(GrhSelect).pixelWidth
    grhHeightTxt.Text = GrhData(GrhSelect).pixelHeight
    bmpTxt.Text = GrhData(GrhSelect).FileNum
    
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

        Case 0 'Graficos.ind
        
        
        Case 1 'Head.ind
         
         
        Case 3 'Body.ind
            Call DesindexarCuerpos
        
        Case 4 'Armas.ind
            Call DesindexarArmas
            
        Case 5 'Escudos.ind
            Call DesindexarEscudos
    End Select
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

Private Sub mnuEditorParticulas_Click()
    frmParticleEditor.Show
End Sub

Private Sub mnuGrhSearch_Click()
On Error Resume Next

    Dim GrhSearch As Long
    Dim i As Long
    Dim j As Long
    
    GrhSearch = InputBox("Ingrese el numero de GRH:")
    
    If IsNumeric(GrhSearch) = False Then Exit Sub
    
    If GrhSearh < 0 Or grhseach > grhCount Then
        
        For i = 1 To grhCount
            
            If GrhData(i).NumFrames >= 1 And i = Archivo Then
                
                For j = 0 To lstGrh(0).ListCount - 1
                    MsgBox "GRH encontrado."
                    lstGrh(0).ListIndex = j
                Next j
                    
            Else
                
                MsgBox "GRH NO ENCONTRADO"
                
            End If
            
        Next i
            
        MsgBox "NO SE ENCONTRO EL GRH"
            
    Else
        
        MsgBox "GRH INVALIDO"
        
    End If

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
