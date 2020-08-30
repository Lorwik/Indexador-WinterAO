Attribute VB_Name = "modIndexado"
Option Explicit

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
    Call LeerINI.Initialize(ExporDir & "Head.ini")
    
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
    Call LeerINI.Initialize(ExporDir & "Helmet.ini")
    
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
    Call LeerINI.Initialize(ExporDir & "\Armas.ini")
    
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
    Call LeerINI.Initialize(ExporDir & "\Escudos.ini")
    
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

Public Function IndexarParticulas()
'*************************************
'Autor: Lorwik
'Fecha: 26/08/2020
'Descripción: Guarda las particulas en un archivo binario
'*************************************

    Dim N As Integer
    Dim loopc As Long
    Dim i As Long
    Dim ColorSet As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim LaCabecera As tCabecera
    
    Call CargarParticulas
    
    N = FreeFile
    Open InitDir & "\Particulas.ind" For Binary Access Write As #N
    
    Put #N, , LaCabecera
    
    Put #N, , TotalStreams

    For loopc = 1 To TotalStreams
        With StreamData(loopc)
            Put #N, , CLng(.NumOfParticles)
            Put #N, , CLng(.NumGrhs)
            Put #N, , CLng(.id)
            Put #N, , CLng(.X1)
            Put #N, , CLng(.Y1)
            Put #N, , CLng(.X2)
            Put #N, , CLng(.Y2)
            Put #N, , CLng(.angle)
            Put #N, , CLng(.vecx1)
            Put #N, , CLng(.vecx2)
            Put #N, , CLng(.vecy1)
            Put #N, , CLng(.vecy2)
            Put #N, , CLng(.life1)
            Put #N, , CLng(.life2)
            Put #N, , CLng(.friction)
            Put #N, , CByte(.spin)
            Put #N, , CSng(.spin_speedL)
            Put #N, , CSng(.spin_speedH)
            Put #N, , CByte(.alphaBlend)
            Put #N, , CByte(.gravity)
            Put #N, , CLng(.grav_strength)
            Put #N, , CLng(.bounce_strength)
            Put #N, , CByte(.XMove)
            Put #N, , CByte(.YMove)
            Put #N, , CLng(.move_x1)
            Put #N, , CLng(.move_x2)
            Put #N, , CLng(.move_y1)
            Put #N, , CLng(.move_y2)
            Put #N, , CSng(.speed)
            Put #N, , CLng(.life_counter)
                
            For i = 1 To .NumGrhs
                Put #N, , CLng(.grh_list(i))
            Next i
                
            For ColorSet = 1 To 4
                Put #N, , CLng(.colortint(ColorSet - 1).R)
                Put #N, , CLng(.colortint(ColorSet - 1).G)
                Put #N, , CLng(.colortint(ColorSet - 1).B)
            Next ColorSet
    
        End With
        
        frmMain.lblstatus.Caption = "Indexado... Particula: " & loopc & " (" & Format((loopc / TotalStreams * 100), "##") & "%)"
        DoEvents
    Next loopc
            
    Close #N
            
    frmMain.lblstatus.Caption = "Guardando...Particulas.ind"
    DoEvents
    
    frmMain.lblstatus.Caption = "Compilado...Particulas.ind"
End Function

Public Sub IndexarColores()
'*************************************
'Autor: Lorwik
'Fecha: 30/08/2020
'Descripción: Guarda los colores en un archivo binario
'*************************************

    Dim N As Integer
    Dim LaCabecera As tCabecera
    
    If CargarColores Then
    
        N = FreeFile
        Open InitDir & "\Colores.ind" For Binary Access Write As #N
        
        Put #N, , LaCabecera
        
        Put #N, , ColoresPJ
        
        Close #N
        
        frmMain.lblstatus.Caption = "Guardando...Colores.ind"
        DoEvents
        
        frmMain.lblstatus.Caption = "Compilado...Colores.ind"
    
    Else
    
        frmMain.lblstatus.Caption = "Error al indexar Colores.dat. No se ha podido leer el archivo de origen."
    
    End If
    
End Sub

' ====================================================
' ################## DESINDEXADO #####################
' ====================================================

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
    
    If FileExist(ExporDir & "Armas.ini", vbArchive) = True Then Call Kill(ExporDir & "Armas.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumArmas=" & NumWeaponAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumWeaponAnims
        Datos = Datos & "[Arma" & (i) & "]" & vbCrLf
        For N = 1 To 4
            Datos = Datos & "Dir" & (N) & "=" & WeaponAnimData(i).WeaponWalk(N).GrhIndex & vbCrLf & IIf(N = 1, Chr(9) & " ' abajo", "") & IIf(N = 2, Chr(9) & " ' arriba", "") & IIf(N = 3, Chr(9) & " ' izquierda", "") & IIf(N = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Armas.ini"
    DoEvents
    
    Open (ExporDir & "Armas.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Armas.ini"
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
    
    If FileExist(ExporDir & "Escudos.ini", vbArchive) = True Then Call Kill(ExporDir & "Armas.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumEscudos=" & NumEscudosAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumEscudosAnims
        Datos = Datos & "[ESC" & (i) & "]" & vbCrLf
        For N = 1 To 4
            Datos = Datos & "Dir" & (N) & "=" & ShieldAnimData(i).ShieldWalk(N).GrhIndex & vbCrLf & IIf(N = 1, Chr(9) & " ' abajo", "") & IIf(N = 2, Chr(9) & " ' arriba", "") & IIf(N = 3, Chr(9) & " ' izquierda", "") & IIf(N = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.ini"
    DoEvents
    
    Open (ExporDir & "Escudos.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Escudos.ini"
End Function

