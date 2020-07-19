Attribute VB_Name = "modCarga"
Option Explicit

Public grhCount As Long
Public NumCuerpos As Integer
Public NumWeaponAnims As Integer
Public NumEscudosAnims As Integer

Public fileVersion As Long

'RUTAS:
Public InitDir As String
Public ExporDir As String
Public GraphicsDir As String

'Indica si se esta trabajando en algo
Public Trabajando As Boolean

Public Function CargarRutas()
'************************************
'Autor: Lorwik
'Fecha: 02/05/2020
'Descripcion: Carga las rutas de los directorios desde un archivo de configuracion
'************************************

    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(App.Path & "\Config.ini")
    
    InitDir = FileManager.GetValue("INIT", "InitDir")
    ExporDir = FileManager.GetValue("INIT", "ExporDir")
    GraphicsDir = FileManager.GetValue("INIT", "GraphicsDir")
    
End Function

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

'************************************************
'LEE ARCHIVOS YA INDEXADOS
'************************************************
Sub ReCargarParticulas()

    Dim StreamFile As String
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
        
    StreamFile = InitDir & "Particulas.ini"
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from particle.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).Y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).X2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).Y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).grh_resize = Val(General_Var_Get(StreamFile, Val(loopc), "resize"))
        StreamData(loopc).grh_resizex = Val(General_Var_Get(StreamFile, Val(loopc), "rx"))
        StreamData(loopc).grh_resizey = Val(General_Var_Get(StreamFile, Val(loopc), "ry"))
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = General_Field_Read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            frmParticleEditor.List2.AddItem loopc & " - " & StreamData(loopc).name
    Next loopc

End Sub

Public Sub LoadGrhData()
On Error GoTo ErrorHandler:

    Dim Grh As Long
    Dim Frame As Long
    Dim Handle As Integer
    Dim LaCabecera As tCabecera
    
    'Open files
    Handle = FreeFile()
    Open InitDir & "Graficos.ind" For Binary Access Read As Handle
    
        'Primero limpiamos el listbox por si es una recarga
        frmMain.lstGrh(0).Clear
    
        Get Handle, , LaCabecera
    
        Get Handle, , fileVersion
        
        Get Handle, , grhCount
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Not EOF(Handle)
            Get Handle, , Grh
            
            With GrhData(Grh)
            
                '.active = True
                Get Handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                'Minimapa
                .active = True
                
                If Not Grh <= 0 Then
                    If .NumFrames > 1 Then
                        frmMain.lstGrh(0).AddItem Grh '& " <ANIMACION>"
                    Else
                        frmMain.lstGrh(0).AddItem Grh
                    End If
                End If
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        Get Handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    Get Handle, , .speed
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    Get Handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = 32
                    .TileHeight = 32
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
    
    Close Handle
    
Exit Sub

ErrorHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
    
End Sub

Public Sub CargarCabezas()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Integer
    Dim NumHeads As Integer
    Dim LaCabecera As tCabecera
    
    N = FreeFile()
    Open InitDir & "Head.ind" For Binary Access Read As #N
    
        Get #N, , LaCabecera
    
        Get #N, , NumHeads   'cantidad de cabezas

        ReDim heads(0 To NumHeads) As tHead
            
        frmMain.lstGrh(1).Clear
            
        For i = 1 To NumHeads
            Get #N, , heads(i).Std
            Get #N, , heads(i).texture
            Get #N, , heads(i).startX
            Get #N, , heads(i).startY
            
            frmMain.lstGrh(1).AddItem i
        Next i

    Close #N
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        End If
        
    End If
    
End Sub

Sub CargarHelmets()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Integer
    Dim NumCascos As Integer
    Dim LaCabecera As tCabecera
    
    N = FreeFile()
    Open InitDir & "Helmet.ind" For Binary Access Read As #N
    
        Get #N, , LaCabecera
    
        Get #N, , NumCascos   'cantidad de cascos
             
        ReDim Cascos(0 To NumCascos) As tHead
            
        frmMain.lstGrh(2).Clear
            
        For i = 1 To NumCascos
            Get #N, , Cascos(i).Std
            Get #N, , Cascos(i).texture
            Get #N, , Cascos(i).startX
            Get #N, , Cascos(i).startY
            
            frmMain.lstGrh(2).AddItem i
        Next i
         
    Close #N
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        End If
        
    End If
End Sub

Public Sub CargarBodys()

On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open InitDir & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    frmMain.lstGrh(3).Clear
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            
            frmMain.lstGrh(3).AddItem i
        End If
        
    Next i
    
    Close #N
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            End
        End If
        
    End If
End Sub

Public Sub CargarArmas()

On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    N = FreeFile
    Open InitDir & "Armas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumWeaponAnims
    
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
    
    frmMain.lstGrh(4).Clear
    
    For i = 1 To NumWeaponAnims
        Get #N, , Weapons(i)
        
        If Weapons(i).Weapon(1) Then
        
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).Weapon(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).Weapon(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).Weapon(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).Weapon(4), 0)
            
            frmMain.lstGrh(4).AddItem i
        End If
    Next i
    
    Close #N

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Armas.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            End
        End If
        
    End If

End Sub

Public Sub CargarEscudos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    N = FreeFile
    Open InitDir & "Escudos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumEscudosAnims
    
    'Resize array
    ReDim ShieldAnimData(1 To NumWeaponAnims) As ShieldAnimData
    ReDim Shields(1 To NumWeaponAnims) As tIndiceEscudos
    
    frmMain.lstGrh(5).Clear
    
    For i = 1 To NumEscudosAnims
        Get #N, , Shields(i)
        
        If Shields(i).Shield(1) Then
        
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).Shield(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).Shield(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).Shield(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).Shield(4), 0)
        
            frmMain.lstGrh(5).AddItem i
        End If
    Next i
    
    Close #N

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Escudos.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            End
        End If
        
    End If
End Sub

Public Sub CargarFX()

On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile
    Open InitDir & "FXs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    frmMain.lstGrh(6).Clear
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
        
        frmMain.lstGrh(6).AddItem i
    Next i
    
    Close #N
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Fxs.ini no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Sub


'************************************************
'LEE DESDE INI PARA INDEXAR
'************************************************

Function GrhIniToGrhDataNew() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Indexa los Graficos.ini
'*************************************

    Dim Grh As Long
    Dim Frame As Long
    Dim Datos As New clsIniReader
    Dim Fr As Integer
    Dim i As Integer
    Dim sTmp As String
    Dim bTmp As Byte
    Dim nF As Integer
    Dim TotalGrh As Long
    
    GrhIniToGrhDataNew = False
    
    'If Dir(Config.initPath & "\Graficos.ind", vbArchive) <> "" Then Call Kill(Config.initPath & "\Graficos.ind")
    
    nF = FreeFile
    
    Call Datos.Initialize(ExporDir & "Graficos.ini")
    
    Open InitDir & "Graficos.ind" For Binary Access Write As #nF
    
    TotalGrh = Datos.GetValue("INIT", "NumGrh")
    
    Seek #nF, 1
    
    Put #nF, , MiCabecera
    
    Put #nF, , fileVersion
    
    Put #nF, , TotalGrh
    
    For Grh = 1 To TotalGrh
        sTmp = Datos.GetValue("Graphics", "Grh" & Grh)
        
        If Len(sTmp) > 0 Then
        
            Fr = General_Field_Read(1, sTmp, 45)
            Put #nF, , Grh
            Put #nF, , Fr 'NumFrames
            
            If Fr > 1 Then
            
                ' ***************** ES UN FRAME **************
                For i = 1 To Fr
                    Put #nF, , CLng(General_Field_Read(i + 1, sTmp, 45))
                Next

                Put #nF, , CSng(General_Field_Read(Fr + 2, sTmp, 45))
                
            ElseIf Fr = 1 Then
            
                ' ***************** ES UN GRH **************
                Put #nF, , CLng(General_Field_Read(2, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(5, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(6, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(3, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(4, sTmp, 45))
                
            End If
            
            frmMain.lblstatus.Caption = "Indexado... Grh: " & Grh & " (" & Format((Grh / TotalGrh * 100), "##") & "%)"
            DoEvents
        End If
    Next
    
    Close #nF

GrhIniToGrhDataNew = True
End Function

Public Function IndexarCabezas()

On Error GoTo fallo

    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumHeads As Integer
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "Head.dat")
    
    NumHeads = CInt(LeerINI.GetValue("INIT", "NumHeads"))
    
    ReDim HeadsT(0 To NumHeads) As tHead
    
    For i = 1 To NumHeads
        HeadsT(i).Std = Val(LeerINI.GetValue("HEAD" & i, "Std"))
        HeadsT(i).texture = Val(LeerINI.GetValue("HEAD" & i, "FileNum"))
        HeadsT(i).startX = Val(LeerINI.GetValue("HEAD" & i, "OffSetX"))
        HeadsT(i).startY = Val(LeerINI.GetValue("HEAD" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Head.ind" For Binary Access Write As #nF
    
    Put #nF, , MiCabecera
    
    Put #nF, , NumHeads
    
    For i = 1 To NumHeads
        Put #nF, , HeadsT(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Cabezas.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Cabezas.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Cabezas.ini"
End Function

Public Function IndexarCascos()
On Error GoTo fallo

    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumCascos As Integer
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "Helmet.dat")
    
    NumCascos = CInt(LeerINI.GetValue("INIT", "NumCascos"))
    
    ReDim HelmesT(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        HelmesT(i).Std = Val(LeerINI.GetValue("CASCO" & i, "Std"))
        HelmesT(i).texture = Val(LeerINI.GetValue("CASCO" & i, "FileNum"))
        HelmesT(i).startX = Val(LeerINI.GetValue("CASCO" & i, "OffSetX"))
        HelmesT(i).startY = Val(LeerINI.GetValue("CASCO" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Helmet.ind" For Binary Access Write As #nF
    
    Put #nF, , MiCabecera
    
    Put #nF, , NumCascos
    
    For i = 1 To NumCascos
        Put #nF, , HelmesT(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Cascos.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Cascos.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Cabezas.ini"
End Function

Public Function IndexarCuerpos()
'******************************
'Autor: Lorwik
'Fecha: 10/05/2020
'Descripcion: Indexa cuerpos.
'********************************

    Dim i As Integer, j, N, K As Integer
    Dim LeerINI As New clsIniReader
    
    'Notificamos que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Call LeerINI.Initialize(ExporDir & "Personajes.ini")
    
    'Total de cuerpos
    NumCuerpos = Val(LeerINI.GetValue("INIT", "NumBodies"))
    
    ReDim CuerpoData(0 To NumCuerpos + 1) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        CuerpoData(i).Body(1) = Val(LeerINI.GetValue("Body" & (i), "WALK1"))
        CuerpoData(i).Body(2) = Val(LeerINI.GetValue("Body" & (i), "WALK2"))
        CuerpoData(i).Body(3) = Val(LeerINI.GetValue("Body" & (i), "WALK3"))
        CuerpoData(i).Body(4) = Val(LeerINI.GetValue("Body" & (i), "WALK4"))
        CuerpoData(i).HeadOffsetX = Val(LeerINI.GetValue("Body" & (i), "HeadOffsetX"))
        CuerpoData(i).HeadOffsetY = Val(LeerINI.GetValue("Body" & (i), "HeadOffsety"))
    Next i
    
    N = FreeFile
    Open InitDir & "Personajes.ind" For Binary Access Write As #N
    
    'Escribimos la cabecera
    Put #N, , MiCabecera
    
    'Guardamos las cabezas
    Put #N, , NumCuerpos
    
    For i = 1 To NumCuerpos
        Put #N, , CuerpoData(i)
    Next i
    
    Close #N
    
    frmMain.lblstatus.Caption = "Compilado...Personajes.ind"

End Function

Public Function IndexarFx()

    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\FXs.ini")
    
    N = FreeFile
    Open InitDir & "\Fxs.ind" For Binary Access Write As #N
    
    Put #N, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumFxs"))
    
    Put #N, , K
    
    Dim EjFx(1) As tIndiceFx
    
    For i = 1 To K
        EjFx(1).OffsetY = LeerINI.GetValue("FX" & i, "OffsetY")
        EjFx(1).OffsetX = LeerINI.GetValue("FX" & i, "OffsetX")
        EjFx(1).Animacion = LeerINI.GetValue("FX" & i, "Animacion")
        Put #N, , EjFx(1)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...FXs.ind"
    DoEvents
    Close #N
    
    frmMain.lblstatus.Caption = "Compilado...FXs.ind"
End Function

Public Function IndexarArmas()

    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\Armas.dat")
    
    N = FreeFile
    Open InitDir & "\Armas.ind" For Binary Access Write As #N
    
    Put #N, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumArmas"))
    
    Put #N, , K
    
    ReDim Weapons(1 To K) As tIndiceArmas
    
    For i = 1 To K
        Weapons(i).Weapon(1) = Val(LeerINI.GetValue("Arma" & i, "Dir1"))
        Weapons(i).Weapon(2) = Val(LeerINI.GetValue("Arma" & i, "Dir2"))
        Weapons(i).Weapon(3) = Val(LeerINI.GetValue("Arma" & i, "Dir3"))
        Weapons(i).Weapon(4) = Val(LeerINI.GetValue("Arma" & i, "Dir4"))
    Next
    
    Put #N, , Weapons()
    
    frmMain.lblstatus.Caption = "Guardando...Armas.ind"
    DoEvents
    Close #N
    
    frmMain.lblstatus.Caption = "Compilado...Armas.ind"
End Function

Public Function IndexarEscudos()

    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\Escudos.dat")
    
    N = FreeFile
    Open InitDir & "\Escudos.ind" For Binary Access Write As #N
    
    Put #N, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumEscudos"))
    
    Put #N, , K
    
    ReDim Shields(1 To K) As tIndiceEscudos
    
    For i = 1 To K
        Shields(i).Shield(1) = Val(LeerINI.GetValue("ESC" & i, "Dir1"))
        Shields(i).Shield(2) = Val(LeerINI.GetValue("ESC" & i, "Dir2"))
        Shields(i).Shield(3) = Val(LeerINI.GetValue("ESC" & i, "Dir3"))
        Shields(i).Shield(4) = Val(LeerINI.GetValue("ESC" & i, "Dir4"))
    Next
    
    Put #N, , Shields()
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.ind"
    DoEvents
    Close #N
    
    frmMain.lblstatus.Caption = "Compilado...Escudos.ind"
End Function

Public Function CargarIndex()
'*************************************
'Autor: Lorwik
'Fecha: 02/05/2020
'Descripción: Carga todos los index
'*************************************
    
    LoadGrhData
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Graficos.ind Recargados!"
    Call CargarBodys
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Personajes.ind Recargados!"
    Call CargarCabezas
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Cabezas.ind Recargadas!"
    Call CargarHelmets
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Cascos.ind Recargados!"
    Call CargarArmas
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Armas.dat Recargadas!"
    Call CargarEscudos
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Escudos.ind Recargados!"
    Call CargarFX
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Fxs.ind Recargados!"
    
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Todos los index fueron recargados"
    
End Function

Public Function DesindexarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: 26/05/2020
'Descripción: Desindexa los cuerpos
'*************************************
On Error Resume Next
    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Personajes.ini", vbArchive) = True Then Call Kill(ExporDir & "Personajes.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumBodies=" & NumCuerpos & vbCrLf & vbCrLf
    
    For i = 1 To NumCuerpos
        Datos = Datos & "[BODY" & (i) & "]" & vbCrLf
        Debug.Print BodyData(i).Walk(N).GrhIndex
        For N = 1 To 4
            Datos = Datos & "WALK" & (N) & "=" & BodyData(i).Walk(N).GrhIndex & vbCrLf & IIf(N = 1, Chr(9) & " ' abajo", "") & IIf(N = 2, Chr(9) & " ' arriba", "") & IIf(N = 3, Chr(9) & " ' izquierda", "") & IIf(N = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
        
        Datos = Datos & "HeadOffsetX=" & BodyData(i).HeadOffset.X & vbCrLf & "HeadOffsetY=" & BodyData(i).HeadOffset.Y & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Personajes.ini"
    DoEvents
    
    Open (ExporDir & "Personajes.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Personajes.ini"
End Function

Public Function DesindexarArmas()
'*************************************
'Autor: Lorwik
'Fecha: 11/06/2020
'Descripción: Desindexa las armas
'*************************************
On Error Resume Next
    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Armas.dat", vbArchive) = True Then Call Kill(ExporDir & "Armas.dat")
    
    Datos = "[INIT]" & vbCrLf & "NumArmas=" & NumWeaponAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumWeaponAnims
        Datos = Datos & "[Arma" & (i) & "]" & vbCrLf
        For N = 1 To 4
            Datos = Datos & "Dir" & (N) & "=" & WeaponAnimData(i).WeaponWalk(N).GrhIndex & vbCrLf & IIf(N = 1, Chr(9) & " ' abajo", "") & IIf(N = 2, Chr(9) & " ' arriba", "") & IIf(N = 3, Chr(9) & " ' izquierda", "") & IIf(N = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Armas.dat"
    DoEvents
    
    Open (ExporDir & "Armas.dat") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Armas.dat"
End Function

Public Function DesindexarEscudos()
'*************************************
'Autor: Lorwik
'Fecha: 11/06/2020
'Descripción: Desindexa las armas
'*************************************
On Error Resume Next
    Dim i As Integer, j, N, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Escudos.dat", vbArchive) = True Then Call Kill(ExporDir & "Armas.dat")
    
    Datos = "[INIT]" & vbCrLf & "NumEscudos=" & NumEscudosAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumEscudosAnims
        Datos = Datos & "[ESC" & (i) & "]" & vbCrLf
        For N = 1 To 4
            Datos = Datos & "Dir" & (N) & "=" & ShieldAnimData(i).ShieldWalk(N).GrhIndex & vbCrLf & IIf(N = 1, Chr(9) & " ' abajo", "") & IIf(N = 2, Chr(9) & " ' arriba", "") & IIf(N = 3, Chr(9) & " ' izquierda", "") & IIf(N = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.dat"
    DoEvents
    
    Open (ExporDir & "Escudos.dat") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Escudos.dat"
End Function
