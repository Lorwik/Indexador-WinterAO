Attribute VB_Name = "modCarga"
Option Explicit

Public grhCount As Long
Public NumCuerpos As Integer
Public NumAtaques As Integer
Public NumWeaponAnims As Integer
Public NumEscudosAnims As Integer
Public NumHeads As Integer
Public NumCascos As Integer
Public NumFxs As Integer

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

    Dim n As Integer
    Dim i As Integer
    Dim LaCabecera As tCabecera
    
    n = FreeFile()
    Open InitDir & "Head.ind" For Binary Access Read As #n
    
        Get #n, , LaCabecera
    
        Get #n, , NumHeads   'cantidad de cabezas

        ReDim heads(0 To NumHeads) As tHead
            
        frmMain.lstGrh(1).Clear
            
        For i = 1 To NumHeads
            Get #n, , heads(i).Std
            Get #n, , heads(i).texture
            Get #n, , heads(i).startX
            Get #n, , heads(i).startY
            
            frmMain.lstGrh(1).AddItem i
        Next i

    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        End If
        
    End If
    
End Sub

Sub CargarHelmets()
On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Integer
    Dim LaCabecera As tCabecera
    
    n = FreeFile()
    Open InitDir & "Helmet.ind" For Binary Access Read As #n
    
        Get #n, , LaCabecera
    
        Get #n, , NumCascos   'cantidad de cascos
             
        ReDim Cascos(0 To NumCascos) As tHead
            
        frmMain.lstGrh(2).Clear
            
        For i = 1 To NumCascos
            Get #n, , Cascos(i).Std
            Get #n, , Cascos(i).texture
            Get #n, , Cascos(i).startX
            Get #n, , Cascos(i).startY
            
            frmMain.lstGrh(2).AddItem i
        Next i
         
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        End If
        
    End If
End Sub

Public Sub CargarBodys()

On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Long
    Dim MisCuerpos() As tIndiceCuerpo
    
    n = FreeFile()
    Open InitDir & "Personajes.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    frmMain.lstGrh(3).Clear
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
            
            frmMain.lstGrh(3).AddItem i
        End If
        
    Next i
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            End
        End If
        
    End If
End Sub

Public Sub CargarAtaques()

On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Long
    Dim MisCuerpos() As tIndiceAtaques
    
    n = FreeFile()
    Open InitDir & "Ataques.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumAtaques
    
    'Resize array
    ReDim AtaqueData(0 To NumAtaques) As AtaqueData
    ReDim MisAtaques(0 To NumAtaques) As tIndiceAtaques
    
    frmMain.lstGrh(7).Clear
    
    If NumAtaques > 0 Then
        For i = 1 To NumCuerpos
            Get #n, , MisAtaques(i)
            
            If MisAtaques(i).Body(1) Then
                Call InitGrh(AtaqueData(i).Walk(1), MisAtaques(i).Body(1), 0)
                Call InitGrh(AtaqueData(i).Walk(2), MisAtaques(i).Body(2), 0)
                Call InitGrh(AtaqueData(i).Walk(3), MisAtaques(i).Body(3), 0)
                Call InitGrh(AtaqueData(i).Walk(4), MisAtaques(i).Body(4), 0)
                
                AtaqueData(i).HeadOffset.x = MisAtaques(i).HeadOffsetX
                AtaqueData(i).HeadOffset.y = MisAtaques(i).HeadOffsetY
                
                frmMain.lstGrh(7).AddItem i
            End If
            
        Next i
    End If
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Ataques.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            End
        End If
        
    End If
End Sub

Public Sub CargarArmas()

On Error GoTo errhandler:

    Dim n As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    n = FreeFile
    Open InitDir & "Armas.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , LaCabecera
    
    'num de cabezas
    Get #n, , NumWeaponAnims
    
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
    
    frmMain.lstGrh(4).Clear
    
    For i = 1 To NumWeaponAnims
        Get #n, , Weapons(i)
        
        If Weapons(i).Weapon(1) Then
        
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).Weapon(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).Weapon(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).Weapon(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).Weapon(4), 0)
            
            frmMain.lstGrh(4).AddItem i
        End If
    Next i
    
    Close #n

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

    Dim n As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    n = FreeFile
    Open InitDir & "Escudos.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , LaCabecera
    
    'num de cabezas
    Get #n, , NumEscudosAnims
    
    'Resize array
    ReDim ShieldAnimData(1 To NumWeaponAnims) As ShieldAnimData
    ReDim Shields(1 To NumWeaponAnims) As tIndiceEscudos
    
    frmMain.lstGrh(5).Clear
    
    For i = 1 To NumEscudosAnims
        Get #n, , Shields(i)
        
        If Shields(i).Shield(1) Then
        
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).Shield(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).Shield(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).Shield(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).Shield(4), 0)
        
            frmMain.lstGrh(5).AddItem i
        End If
    Next i
    
    Close #n

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

    Dim n As Integer
    Dim i As Long
    
    n = FreeFile
    Open InitDir & "FXs.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    frmMain.lstGrh(6).Clear
    
    For i = 1 To NumFxs
        Get #n, , FxData(i)
        
        frmMain.lstGrh(6).AddItem i
    Next i
    
    Close #n
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Fxs.ini no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Sub

Public Function CargarColores() As Boolean

On Error GoTo errhandler:

    If Not FileExist(ExporDir & "colores.dat", vbNormal) Then Exit Function

    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "colores.dat")
    
    Dim i As Long
    
    For i = 0 To MAXCOLORES '48, 49 y 50 reservados para atacables, ciudadano y criminal
        ColoresPJ(i).R = LeerINI.GetValue(CStr(i), "R")
        ColoresPJ(i).G = LeerINI.GetValue(CStr(i), "G")
        ColoresPJ(i).B = LeerINI.GetValue(CStr(i), "B")
    Next i
    
    Set LeerINI = Nothing
    
    CargarColores = True
    
errhandler:

    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo colores.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            End
        End If
        
    End If
End Function

Public Function CargarIndex()
'*************************************
'Autor: Lorwik
'Fecha: 02/05/2020
'Descripción: Carga todos los index
'*************************************
    
    Call LoadGrhData
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Graficos.ind Recargados!"
    
    Call CargarBodys
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Personajes.ind Recargados!"
    
    Call CargarAtaques
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Ataques.ind Recargados!"
    
    Call CargarCabezas
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Cabezas.ind Recargadas!"
    
    Call CargarHelmets
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Cascos.ind Recargados!"
    
    Call CargarArmas
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Armas.ini Recargadas!"
    
    Call CargarEscudos
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Escudos.ind Recargados!"
    
    Call CargarFX
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Fxs.ind Recargados!"
    
    If CargarColores Then _
    If frmMain.Visible Then frmMain.lblstatus.Caption = "Colores.dat Recargados!"
    
    If frmMain.Visible = True Then frmMain.lblstatus.Caption = "Todos los index fueron recargados"
    
End Function
