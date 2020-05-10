Attribute VB_Name = "modCarga"
Option Explicit

Public grhCount As Long
Public NumCuerpos As Integer

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
    
    
    'Open files
    Handle = FreeFile()
    Open InitDir & "Graficos.ind" For Binary Access Read As Handle
    
        'Primero limpiamos el listbox por si es una recarga
        frmMain.Grhs.Clear
    
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
                        frmMain.Grhs.AddItem Grh & " <ANIMACION>"
                    Else
                        frmMain.Grhs.AddItem Grh
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
                    
                    Get Handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
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

    Dim n As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    If Not FileExist(InitDir & "Heads.ind", vbArchive) Then GoTo errhandler
    
    n = FreeFile()
    Open InitDir & "Heads.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Cabezas.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
    
End Sub

Sub CargarHelmets()
On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    If Not FileExist(InitDir & "Helmets.ind", vbArchive) Then GoTo errhandler
    
    n = FreeFile()
    Open InitDir & "Helmets.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Cascos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
    
End Sub

Public Sub CargarBodys()

On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    If Not FileExist(InitDir & "Personajes.ind", vbArchive) Then GoTo errhandler
    
    n = FreeFile()
    Open InitDir & "Personajes.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            
            frmMain.lstBodys.AddItem i
        End If
    Next i
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Sub

Public Sub CargarArmas()

On Error GoTo errhandler:

    Dim loopc As Long
    Dim NumWeaponAnims As Integer
    
    If Not FileExist(InitDir & "Weapons.ind", vbArchive) Then GoTo errhandler
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "Weapons.ind")
    
    NumWeaponAnims = Val(FileManager.GetValue("INIT", "NumArmas"))
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        Call InitGrh(WeaponAnimData(loopc).WeaponWalk(1), Val(FileManager.GetValue("ARMA" & loopc, "Dir1")), 0)
        Call InitGrh(WeaponAnimData(loopc).WeaponWalk(2), Val(FileManager.GetValue("ARMA" & loopc, "Dir2")), 0)
        Call InitGrh(WeaponAnimData(loopc).WeaponWalk(3), Val(FileManager.GetValue("ARMA" & loopc, "Dir3")), 0)
        Call InitGrh(WeaponAnimData(loopc).WeaponWalk(4), Val(FileManager.GetValue("ARMA" & loopc, "Dir4")), 0)
    Next loopc
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo armas.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Sub

Public Sub CargarEscudos()
On Error GoTo errhandler:

    Dim loopc As Long
    Dim NumEscudosAnims As Integer
    
    If Not FileExist(InitDir & "Shields.ind", vbArchive) Then GoTo errhandler
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "Shields.ind")
    
    NumEscudosAnims = Val(FileManager.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        Call InitGrh(ShieldAnimData(loopc).ShieldWalk(1), Val(FileManager.GetValue("ESC" & loopc, "Dir1")), 0)
        Call InitGrh(ShieldAnimData(loopc).ShieldWalk(2), Val(FileManager.GetValue("ESC" & loopc, "Dir2")), 0)
        Call InitGrh(ShieldAnimData(loopc).ShieldWalk(3), Val(FileManager.GetValue("ESC" & loopc, "Dir3")), 0)
        Call InitGrh(ShieldAnimData(loopc).ShieldWalk(4), Val(FileManager.GetValue("ESC" & loopc, "Dir4")), 0)
    Next loopc
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo escudos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Sub

Public Sub CargarFX()

On Error GoTo errhandler:

    Dim i As Long
    
    If Not FileExist(InitDir & "FXs.ind", vbArchive) Then Exit Sub
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "FXs.ind")
    
    'Resize array
    ReDim FxData(0 To FileManager.GetValue("INIT", "NumFxs")) As tIndiceFx
    
    For i = 1 To UBound(FxData())
        
        With FxData(i)
            .Animacion = Val(FileManager.GetValue("FX" & CStr(i), "Animacion"))
            .OffsetX = Val(FileManager.GetValue("FX" & CStr(i), "OffsetX"))
            .OffsetY = Val(FileManager.GetValue("FX" & CStr(i), "OffsetY"))
        End With
    
    Next
    
    Set FileManager = Nothing
    
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
    
    Put #nF, , fileVersion
    
    Put #nF, , TotalGrh
    
    For Grh = 1 To TotalGrh
        sTmp = Datos.GetValue("Graphics", "Grh" & Grh)
        
        If Len(sTmp) > 0 Then
        
            Fr = General_Field_Read(1, sTmp, 45)
            Put #nF, , Grh
            Put #nF, , Fr
            
            If Fr > 1 Then
            
                ' ***************** ES UN FRAME **************
                For i = 1 To Fr
                    Put #nF, , CLng(General_Field_Read(i + 1, sTmp, 45))
                Next
                
                Put #nF, , CSng(General_Field_Read(Fr + 2, sTmp, 45))
                
            ElseIf Fr = 1 Then
            
                ' ***************** ES UN GRH **************
                Put #nF, , CLng(General_Field_Read(2, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(3, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(4, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(5, sTmp, 45))
                Put #nF, , CInt(General_Field_Read(6, sTmp, 45))
                
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
    Dim Numheads As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "Heads.dat")
    
    Numheads = CInt(LeerINI.GetValue("INIT", "NumHeads"))
    
    ReDim HeadsT(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        'HeadsT(i).Std = Val(LeerINI.GetValue("HEAD" & i, "Std"))
        'HeadsT(i).texture = Val(LeerINI.GetValue("HEAD" & i, "FileNum"))
        'HeadsT(i).startX = Val(LeerINI.GetValue("HEAD" & i, "OffSetX"))
        'HeadsT(i).startY = Val(LeerINI.GetValue("HEAD" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Heads.ind" For Binary Access Write As #nF
    
    For i = 0 To 17
        Put #nF, , bTmp
    Next i
    
    Put #nF, , Numheads
    
    For i = 1 To Numheads
        Put #nF, , HeadsT(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Heads.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Heads.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Heads.ini"
End Function

Public Function IndexarCascos()

On Error GoTo fallo
    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumCascos As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "Helmets.dat")
    
    NumCascos = CInt(LeerINI.GetValue("INIT", "NumCascos"))
    
    'ReDim HelmesT(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        'HelmesT(i).Std = Val(LeerINI.GetValue("CASCO" & i, "Std"))
        'HelmesT(i).texture = Val(LeerINI.GetValue("CASCO" & i, "FileNum"))
        'HelmesT(i).startX = Val(LeerINI.GetValue("CASCO" & i, "OffSetX"))
        'HelmesT(i).startY = Val(LeerINI.GetValue("CASCO" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Helmets.ind" For Binary Access Write As #nF
    
    For i = 0 To 17
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumCascos
    
    For i = 1 To NumCascos
        'Put #nF, , HelmesT(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Helmets.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Helmets.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Helmets.dat"
End Function

Public Function IndexarCuerpos()
'******************************
'Autor: Lorwik
'Fecha: 10/05/2020
'Descripcion: Indexa cuerpos.
'********************************

    Dim i As Integer, j, n, K As Integer
    Dim LeerINI As New clsIniReader
    
    'Notificamos que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Call LeerINI.Initialize(ExporDir & "\Cuerpos.ini")
    
    'Total de cuerpos
    NumCuerpos = Val(LeerINI.GetValue("INIT", "NumBodies"))
    
    ReDim CuerpoData(0 To NumCuerpos + 1) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        CuerpoData(i).Body(1) = LeerINI.GetValue("Body" & (i), "WALK1")
        CuerpoData(i).Body(2) = LeerINI.GetValue("Body" & (i), "WALK2")
        CuerpoData(i).Body(3) = LeerINI.GetValue("Body" & (i), "WALK3")
        CuerpoData(i).Body(4) = LeerINI.GetValue("Body" & (i), "WALK4")
        CuerpoData(i).HeadOffsetX = LeerINI.GetValue("Body" & (i), "HeadOffsetX")
        CuerpoData(i).HeadOffsetY = LeerINI.GetValue("Body" & (i), "HeadOffsety")
    Next i
    
    n = FreeFile
    Open InitDir & "\Personajes.ind" For Binary Access Write As #n
    
    'Escribimos la cabecera
    Put #n, , MiCabecera
    'Guardamos las cabezas
    Put #n, , NumCuerpos
    
    For i = 1 To NumCuerpos
        Put #n, , CuerpoData(i)
    Next i
    
    Close #n
    
    frmMain.lblstatus.Caption = "Compilado...Personajes.ind"

End Function

Public Function IndexarFx()

    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\FXs.ini")
    
    n = FreeFile
    Open InitDir & "\Fxs.ind" For Binary Access Write As #n
    
    Put #n, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumFxs"))
    
    Put #n, , K
    
    Dim EjFx(1) As tIndiceFx
    
    For i = 1 To K
        EjFx(1).OffsetY = LeerINI.GetValue("FX" & i, "OffsetY")
        EjFx(1).OffsetX = LeerINI.GetValue("FX" & i, "OffsetX")
        EjFx(1).Animacion = LeerINI.GetValue("FX" & i, "Animacion")
        Put #n, , EjFx(1)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...FXs.ind"
    DoEvents
    Close #n
    
    frmMain.lblstatus.Caption = "Compilado...FXs.ind"
End Function

Sub ImportarDAT(ByVal Nombre As String)
'***************************************
'Autor: ???
'Fecha: ???
'Descripcion: Indexa un archivo .dat
'*****************************************
    On Error Resume Next
    
    frmMain.lblstatus.Caption = "Indexando..."
    DoEvents
    
    '¿Existe el archivo desindexado?
    If LenB(Dir(ExporDir & "\" & Nombre & ".ini", vbArchive)) = 0 Then
        frmMain.lblstatus.Caption = "ERROR: No existe " & Nombre & ".ini"
        Exit Sub
    End If
    
    'Si ya existe un archivo indexado lo eliminamos
    Call Kill(InitDir & "\" & Nombre & ".dat")
    
    Call FileCopy(ExporDir & "\" & Nombre & ".ini", InitDir & "\" & Nombre & ".dat")
    
    If LenB(Dir(InitDir & "\" & Nombre & ".dat", vbArchive)) = 0 Then
        frmMain.lblstatus.Caption = "ERROR: No se ha podido indexar " & Nombre & ".dat"
    Else
        frmMain.lblstatus.Caption = "Indexando..." & Nombre & ".dat"
    End If
    
End Sub

Sub ExportarDAT(ByVal Nombre As String)
'***************************************
'Autor: ???
'Fecha: ???
'Descripcion: desindexa un archivo .dat
'*****************************************

On Error Resume Next

    frmMain.lblstatus.Caption = "Desindexando..."
    DoEvents
    
    '¿Existe el archivo indexado?
    If LenB(Dir(InitDir & "\" & Nombre & ".dat", vbArchive)) = 0 Then
        frmMain.lblstatus.Caption = "ERROR: No existe " & Nombre & ".dat."
        Exit Sub
    End If
    
    'Si ya existe un archivo indexado lo eliminamos
    Call Kill(ExporDir & "\" & Nombre & ".ini")
    
    Call FileCopy(InitDir & "\" & Nombre & ".dat", ExporDir & "\" & Nombre & ".ini")
    
    If LenB(Dir(ExporDir & "\" & Nombre & ".ini", vbArchive)) = 0 Then
        frmMain.lblstatus.Caption = "ERROR: No se ha podido desindexar " & Nombre & ".ini"
    Else
        frmMain.lblstatus.Caption = "Desindexando..." & Nombre & ".ini"
    End If
    
End Sub

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

