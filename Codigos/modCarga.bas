Attribute VB_Name = "modCarga"
Option Explicit

Public grhCount As Long

'DIRECTORIOS
Public Function InitDir() As String
    InitDir = App.Path & "\Init\"
End Function

Public Function ExporDir() As String
    ExporDir = App.Path & "\Exportados\"
End Function

'************************************************
'LEE ARCHIVOS YA INDEXADOS
'************************************************
Sub ReCargarParticulas()

    Dim StreamFile As String
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
        
    StreamFile = InitDir & "Particle.ini"
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from particle.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).name = General_Var_Get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = General_Var_Get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = General_Var_Get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = General_Var_Get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = General_Var_Get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = General_Var_Get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = General_Var_Get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = General_Var_Get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = General_Var_Get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = General_Var_Get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = General_Var_Get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = General_Var_Get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = General_Var_Get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = General_Var_Get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = General_Var_Get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = General_Var_Get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = General_Var_Get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = General_Var_Get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = General_Var_Get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = General_Var_Get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = General_Var_Get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = General_Var_Get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(General_Var_Get(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).grh_resize = Val(General_Var_Get(StreamFile, Val(LoopC), "resize"))
        StreamData(LoopC).grh_resizex = Val(General_Var_Get(StreamFile, Val(LoopC), "rx"))
        StreamData(LoopC).grh_resizey = Val(General_Var_Get(StreamFile, Val(LoopC), "ry"))
        StreamData(LoopC).NumGrhs = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = General_Field_Read(Str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            frmParticleEditor.List2.AddItem LoopC & " - " & StreamData(LoopC).name
    Next LoopC

End Sub

Public Sub LoadGrhData()
On Error GoTo ErrorHandler:

    Dim Grh As Long
    Dim Frame As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    
    'Open files
    Handle = FreeFile()
    Open InitDir & "Graficos.ind" For Binary Access Read As Handle
    
        Get Handle, , fileVersion
        
        Get Handle, , grhCount
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Not EOF(Handle)
            Get Handle, , Grh
            
            With GrhData(Grh)
            
                '.active = True
                Get Handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
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

    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    If Not FileExist(InitDir & "Heads.ind", vbArchive) Then GoTo errhandler
    
    N = FreeFile()
    Open InitDir & "Heads.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
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

    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    If Not FileExist(InitDir & "Helmets.ind", vbArchive) Then GoTo errhandler
    
    N = FreeFile()
    Open InitDir & "Helmets.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
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

    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    If Not FileExist(InitDir & "Body.ind", vbArchive) Then GoTo errhandler
    
    N = FreeFile()
    Open InitDir & "Body.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
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

    Dim LoopC As Long
    Dim NumWeaponAnims As Integer
    
    If Not FileExist(InitDir & "Weapons.ind", vbArchive) Then GoTo errhandler
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "Weapons.ind")
    
    NumWeaponAnims = Val(FileManager.GetValue("INIT", "NumArmas"))
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(1), Val(FileManager.GetValue("ARMA" & LoopC, "Dir1")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(2), Val(FileManager.GetValue("ARMA" & LoopC, "Dir2")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(3), Val(FileManager.GetValue("ARMA" & LoopC, "Dir3")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(4), Val(FileManager.GetValue("ARMA" & LoopC, "Dir4")), 0)
    Next LoopC
    
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

    Dim LoopC As Long
    Dim NumEscudosAnims As Integer
    
    If Not FileExist(InitDir & "Shields.ind", vbArchive) Then GoTo errhandler
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "Shields.ind")
    
    NumEscudosAnims = Val(FileManager.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(1), Val(FileManager.GetValue("ESC" & LoopC, "Dir1")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(2), Val(FileManager.GetValue("ESC" & LoopC, "Dir2")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(3), Val(FileManager.GetValue("ESC" & LoopC, "Dir3")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(4), Val(FileManager.GetValue("ESC" & LoopC, "Dir4")), 0)
    Next LoopC
    
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
    
    Dim FileManager As New clsIniReader
    Call FileManager.Initialize(InitDir & "FX.ind")
    
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
    Dim Grh As Long
    Dim Frame As Long
    Dim Datos As New clsIniReader
    Dim Fr As Integer
    Dim i As Integer
    Dim sTmp As String
    Dim bTmp As Byte
    Dim nF As Integer
    Dim grhCount As Long
    
    GrhIniToGrhDataNew = False
    
    'If Dir(Config.initPath & "\Graficos.ind", vbArchive) <> "" Then Call Kill(Config.initPath & "\Graficos.ind")
    
    nF = FreeFile
    
    Call Datos.Initialize(ExporDir & "Graficos.ini")
    
    Open InitDir & "Graficos.ind" For Binary Access Write As #nF
    
    grhCount = Datos.GetValue("INIT", "grh_count")
    
    Seek #nF, 1
    
    For i = 0 To 13
        Put #nF, , bTmp
    Next i

    Put #nF, , grhCount
    
    For Grh = 1 To grhCount
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
            frmMain.lblstatus.Caption = "Indexado... Grh: " & Grh & " (" & Format((Grh / grhCount * 100), "##") & "%)"
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
    
    ReDim HeadsT(0 To Numheads) As tHead
    
    For i = 1 To Numheads
        HeadsT(i).Std = Val(LeerINI.GetValue("HEAD" & i, "Std"))
        HeadsT(i).Texture = Val(LeerINI.GetValue("HEAD" & i, "FileNum"))
        HeadsT(i).startX = Val(LeerINI.GetValue("HEAD" & i, "OffSetX"))
        HeadsT(i).startY = Val(LeerINI.GetValue("HEAD" & i, "OffSetY"))
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
    
    ReDim HelmesT(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        HelmesT(i).Std = Val(LeerINI.GetValue("CASCO" & i, "Std"))
        HelmesT(i).Texture = Val(LeerINI.GetValue("CASCO" & i, "FileNum"))
        HelmesT(i).startX = Val(LeerINI.GetValue("CASCO" & i, "OffSetX"))
        HelmesT(i).startY = Val(LeerINI.GetValue("CASCO" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Helmets.ind" For Binary Access Write As #nF
    
    For i = 0 To 17
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumCascos
    
    For i = 1 To NumCascos
        Put #nF, , HelmesT(i)
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

On Error GoTo fallo

    Dim NumBodys As Integer
    Dim i As Integer
    Dim j As Byte
    Dim tmpint As Integer
    Dim nF As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando Body.dat..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "Body.dat")
    
    'Obtenemos el numero total de cuerpos
    NumBodys = CInt(LeerINI.GetValue("INIT", "NumBodies"))
    ReDim BodysT(0 To NumBodys) As tIndiceCuerpo
    
    For i = 1 To NumBodys
        'Intentamos leer el Std
        tmpint = Val(LeerINI.GetValue("Body" & i, "Std"))
        
        'Si es 1, se trata del nuevo formato
        If tmpint = 1 Then
            BodysT(i).Std = tmpint
            BodysT(i).Texture = Val(LeerINI.GetValue("Body" & i, "FileNum"))
            BodysT(i).startX = Val(LeerINI.GetValue("Body" & i, "OffSetX"))
            BodysT(i).startY = Val(LeerINI.GetValue("Body" & i, "OffSetY"))
        Else 'Si es 0, es el formato clasico
            BodysT(i).Body(1) = LeerINI.GetValue("Body" & (i), "WALK1")
            BodysT(i).Body(2) = LeerINI.GetValue("Body" & (i), "WALK2")
            BodysT(i).Body(3) = LeerINI.GetValue("Body" & (i), "WALK3")
            BodysT(i).Body(4) = LeerINI.GetValue("Body" & (i), "WALK4")
        End If
        
        'Cosas que siempre va a tener sin importar el formato:
        BodysT(i).HeadOffsetY = Val(LeerINI.GetValue("Body" & i, "HeadOffsetY"))
        BodysT(i).HeadOffsetX = Val(LeerINI.GetValue("Body" & i, "HeadOffsetX"))
        BodysT(i).StaticWalk = Val(LeerINI.GetValue("Body" & i, "StaticWalk"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Body.ind" For Binary Access Write As #nF
    
    For i = 0 To 17
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumBodys
    
    For i = 1 To NumBodys
        Put #nF, , BodysT(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Body.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Body.ind"

    Exit Function
fallo:
    MsgBox "Error en Body.dat"
End Function

Public Function IndexarArmas()

On Error GoTo fallo
    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumArmas As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "Weapons.dat")
    
    NumArmas = CInt(LeerINI.GetValue("INIT", "NumArmas"))
    
    ReDim Armast(0 To NumArmas) As tWeapons
    
    For i = 1 To NumArmas
        Armast(i).Std = Val(LeerINI.GetValue("ARMAS" & i, "Std"))
        Armast(i).Texture = Val(LeerINI.GetValue("ARMAS" & i, "FileNum"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Weapons.ind" For Binary Access Write As #nF
    
    For i = 0 To 23
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumArmas
    
    For i = 1 To NumArmas
        Put #nF, , Armast(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Armas.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Armas.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Armas.dat"
End Function

Public Function IndexarEscudos()

On Error GoTo fallo
    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumEscudos As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "Shields.dat")
    
    NumEscudos = CInt(LeerINI.GetValue("INIT", "NumEscudos"))
    
    ReDim Escudost(0 To NumEscudos) As tShields
    
    For i = 1 To NumEscudos
        Escudost(i).Std = Val(LeerINI.GetValue("ESC" & i, "Std"))
        Escudost(i).Texture = Val(LeerINI.GetValue("ESC" & i, "FileNum"))
        Escudost(i).OffsetX = Val(LeerINI.GetValue("ESC" & i, "OffSetX"))
        Escudost(i).OffsetY = Val(LeerINI.GetValue("ESC" & i, "OffSetX"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Shields.ind" For Binary Access Write As #nF
    
    For i = 0 To 24
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumEscudos
    
    For i = 1 To NumEscudos
        Put #nF, , Escudost(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...Escudos.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Escudos.dat"
End Function

Public Function IndexarFx()

On Error GoTo fallo
    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumFX As Integer
    Dim bTmp As Byte
    
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    Dim LeerINI As New clsIniReader
    
    Call LeerINI.Initialize(ExporDir & "fx.dat")
    
    NumFX = CInt(LeerINI.GetValue("INIT", "NumFXs"))
    
    ReDim Fxst(0 To NumFX) As tFx
    
    For i = 1 To NumFX
        Fxst(i).Animacion = Val(LeerINI.GetValue("FX" & i, "Animacion"))
        Fxst(i).OffsetX = Val(LeerINI.GetValue("FX" & i, "OffSetX"))
        Fxst(i).OffsetY = Val(LeerINI.GetValue("FX" & i, "OffSetX"))
        Fxst(i).Blend = Val(LeerINI.GetValue("FX" & i, "Blend"))
        Fxst(i).color = Val(LeerINI.GetValue("FX" & i, "Color"))
        Fxst(i).angle = Val(LeerINI.GetValue("FX" & i, "Angle"))
    Next i
    
    nF = FreeFile
    Open InitDir & "Fx.ind" For Binary Access Write As #nF
    
    For i = 0 To 34
        Put #nF, , bTmp
    Next i
    
    Put #nF, , NumFX
    
    For i = 1 To NumFX
        Put #nF, , Fxst(i)
    Next
    
    frmMain.lblstatus.Caption = "Guardando...FX.ind"
    DoEvents
    Close #nF
    frmMain.lblstatus.Caption = "Compilado...FX.ind"
    
    Exit Function
fallo:
    MsgBox "Error en Escudos.dat"
End Function
