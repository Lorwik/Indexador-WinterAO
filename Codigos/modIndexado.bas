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

    Dim i As Integer, j, n, K As Integer
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
    
    n = FreeFile
    Open InitDir & "Personajes.ind" For Binary Access Write As #n
    
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

Public Function IndexarAtaques()
'******************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripcion: Indexa ataques
'********************************

    Dim i As Integer, j, n, K As Integer
    Dim LeerINI As New clsIniReader
    
    'Notificamos que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Call LeerINI.Initialize(ExporDir & "Ataques.ini")
    
    'Total de cuerpos
    NumCuerpos = Val(LeerINI.GetValue("INIT", "NumAtaques"))
    
    ReDim AtackData(0 To NumAtaques + 1) As tIndiceAtaques
    
    For i = 1 To NumAtaques
        AtackData(i).Body(1) = Val(LeerINI.GetValue("Ataque" & (i), "WALK1"))
        AtackData(i).Body(2) = Val(LeerINI.GetValue("Ataque" & (i), "WALK2"))
        AtackData(i).Body(3) = Val(LeerINI.GetValue("Ataque" & (i), "WALK3"))
        AtackData(i).Body(4) = Val(LeerINI.GetValue("Ataque" & (i), "WALK4"))
        AtackData(i).HeadOffsetX = Val(LeerINI.GetValue("Body" & (i), "HeadOffsetX"))
        AtackData(i).HeadOffsetY = Val(LeerINI.GetValue("Body" & (i), "HeadOffsety"))
    Next i
    
    n = FreeFile
    Open InitDir & "Ataques.ind" For Binary Access Write As #n
    
    'Escribimos la cabecera
    Put #n, , MiCabecera
    
    'Guardamos las cabezas
    Put #n, , NumAtaques
    
    For i = 1 To NumAtaques
        Put #n, , AtackData(i)
    Next i
    
    Close #n
    
    frmMain.lblstatus.Caption = "Compilado...Ataques.ind"

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

Public Function IndexarArmas()

    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\Armas.ini")
    
    n = FreeFile
    Open InitDir & "\Armas.ind" For Binary Access Write As #n
    
    Put #n, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumArmas"))
    
    Put #n, , K
    
    ReDim Weapons(1 To K) As tIndiceArmas
    
    For i = 1 To K
        Weapons(i).Weapon(1) = Val(LeerINI.GetValue("Arma" & i, "Dir1"))
        Weapons(i).Weapon(2) = Val(LeerINI.GetValue("Arma" & i, "Dir2"))
        Weapons(i).Weapon(3) = Val(LeerINI.GetValue("Arma" & i, "Dir3"))
        Weapons(i).Weapon(4) = Val(LeerINI.GetValue("Arma" & i, "Dir4"))
    Next
    
    Put #n, , Weapons()
    
    frmMain.lblstatus.Caption = "Guardando...Armas.ind"
    DoEvents
    Close #n
    
    frmMain.lblstatus.Caption = "Compilado...Armas.ind"
End Function

Public Function IndexarEscudos()

    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    'Notificamos de que vamos a indexar
    frmMain.lblstatus.Caption = "Compilando..."
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(ExporDir & "\Escudos.ini")
    
    n = FreeFile
    Open InitDir & "\Escudos.ind" For Binary Access Write As #n
    
    Put #n, , MiCabecera
    
    K = Val(LeerINI.GetValue("INIT", "NumEscudos"))
    
    Put #n, , K
    
    ReDim Shields(1 To K) As tIndiceEscudos
    
    For i = 1 To K
        Shields(i).Shield(1) = Val(LeerINI.GetValue("ESC" & i, "Dir1"))
        Shields(i).Shield(2) = Val(LeerINI.GetValue("ESC" & i, "Dir2"))
        Shields(i).Shield(3) = Val(LeerINI.GetValue("ESC" & i, "Dir3"))
        Shields(i).Shield(4) = Val(LeerINI.GetValue("ESC" & i, "Dir4"))
    Next
    
    Put #n, , Shields()
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.ind"
    DoEvents
    Close #n
    
    frmMain.lblstatus.Caption = "Compilado...Escudos.ind"
End Function

Public Function IndexarParticulas()
'*************************************
'Autor: Lorwik
'Fecha: 26/08/2020
'Descripción: Guarda las particulas en un archivo binario
'*************************************

    Dim n As Integer
    Dim Loopc As Long
    Dim i As Long
    Dim ColorSet As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim LaCabecera As tCabecera
    
    Call CargarParticulas
    
    n = FreeFile
    Open InitDir & "\Particulas.ind" For Binary Access Write As #n
    
    Put #n, , LaCabecera
    
    Put #n, , TotalStreams

    For Loopc = 1 To TotalStreams
        With StreamData(Loopc)
            Put #n, , CLng(.NumOfParticles)
            Put #n, , CLng(.NumGrhs)
            Put #n, , CLng(.id)
            Put #n, , CLng(.X1)
            Put #n, , CLng(.Y1)
            Put #n, , CLng(.X2)
            Put #n, , CLng(.Y2)
            Put #n, , CLng(.Angle)
            Put #n, , CLng(.vecx1)
            Put #n, , CLng(.vecx2)
            Put #n, , CLng(.vecy1)
            Put #n, , CLng(.vecy2)
            Put #n, , CLng(.life1)
            Put #n, , CLng(.life2)
            Put #n, , CLng(.friction)
            Put #n, , CByte(.spin)
            Put #n, , CSng(.spin_speedL)
            Put #n, , CSng(.spin_speedH)
            Put #n, , CByte(.alphaBlend)
            Put #n, , CByte(.gravity)
            Put #n, , CLng(.grav_strength)
            Put #n, , CLng(.bounce_strength)
            Put #n, , CByte(.XMove)
            Put #n, , CByte(.YMove)
            Put #n, , CLng(.move_x1)
            Put #n, , CLng(.move_x2)
            Put #n, , CLng(.move_y1)
            Put #n, , CLng(.move_y2)
            Put #n, , CSng(.speed)
            Put #n, , CLng(.life_counter)
                
            For i = 1 To .NumGrhs
                Put #n, , CLng(.grh_list(i))
            Next i
                
            For ColorSet = 1 To 4
                Put #n, , CLng(.colortint(ColorSet - 1).R)
                Put #n, , CLng(.colortint(ColorSet - 1).G)
                Put #n, , CLng(.colortint(ColorSet - 1).B)
            Next ColorSet
    
        End With
        
        frmMain.lblstatus.Caption = "Indexado... Particula: " & Loopc & " (" & Format((Loopc / TotalStreams * 100), "##") & "%)"
        DoEvents
    Next Loopc
            
    Close #n
            
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

    Dim n As Integer
    Dim LaCabecera As tCabecera
    
    If CargarColores Then
    
        n = FreeFile
        Open InitDir & "\Colores.ind" For Binary Access Write As #n
        
        Put #n, , LaCabecera
        
        Put #n, , ColoresPJ
        
        Close #n
        
        frmMain.lblstatus.Caption = "Guardando...Colores.ind"
        DoEvents
        
        frmMain.lblstatus.Caption = "Compilado...Colores.ind"
    
    Else
    
        frmMain.lblstatus.Caption = "Error al indexar Colores.dat. No se ha podido leer el archivo de origen."
    
    End If
    
End Sub

Public Sub IndexarGUI()
'*************************************
'Autor: Lorwik
'Fecha: 30/08/2020
'Descripción: Guarda la GUI en un archivo binario
'*************************************

    Dim n               As Integer
    Dim LaCabecera      As tCabecera
    Dim Leer            As New clsIniReader
    Dim i               As Integer
    Dim NumButtons      As Integer
    Dim NumConnectMap   As Byte

    If FileExist(ExporDir & "GUI.dat", vbArchive) = True Then
        Call Leer.Initialize(ExporDir & "GUI.dat")
        
        n = FreeFile
        Open InitDir & "\GUI.ind" For Binary Access Write As #n
        
            Put #n, , LaCabecera
            
            NumButtons = Val(Leer.GetValue("INIT", "NumButtons"))
            Put #n, , NumButtons
            
            NumConnectMap = Val(Leer.GetValue("INIT", "NumMaps"))
            Put #n, , NumConnectMap
            
            'Mapas de GUI
            For i = 1 To NumConnectMap
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "Map"))
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "X"))
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "Y"))
            Next i
            
            'Posiciones de los PJ
            For i = 1 To 10
                Put #n, , CInt(Leer.GetValue("PJPos" & i, "X"))
                Put #n, , CInt(Leer.GetValue("PJPos" & i, "Y"))
            Next i
            
            'Posiciones de los botones
            For i = 1 To NumButtons
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "X"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "Y"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "PosX"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "PosY"))
                Put #n, , CLng(Leer.GetValue("BUTTON" & i, "GrhNormal"))
       
            Next i
        
        Close #n
        
        frmMain.lblstatus.Caption = "Guardando...GUI.ind"
        DoEvents
            
        frmMain.lblstatus.Caption = "Compilado...GUI.ind"
        
    Else
    
        frmMain.lblstatus.Caption = "Error al indexar GUID.dat. No se ha encontrado el archivo de origen."
    
    End If
    
    Set Leer = Nothing
End Sub

' ====================================================
' ################## DESINDEXADO #####################
' ====================================================

Public Function DesindexarGraficos()
    On Error Resume Next
    Dim i As Long, j As Integer, K As Integer
    Dim n
    Dim Datos$
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Graficos.ini", vbArchive) = True Then Call Kill(ExporDir & "Graficos.ini")
    
    n = FreeFile
    Open (ExporDir & "Graficos.ini") For Binary Access Write As n
    Put n, , "[INIT]" & vbCrLf & "NumGrh=" & grhCount & vbCrLf & vbCrLf
    K = 0
    
    Put n, , "[Graphics]" & vbCrLf
    
    For i = 1 To grhCount
        K = K + 1
        If K > 100 Then
            frmMain.lblstatus.Caption = "Exportando..." & i & " de MaxGRH"
            DoEvents
            K = 0
        End If
        
        If GrhData(i).NumFrames > 0 Then
            Datos$ = ""
            If GrhData(i).NumFrames = 1 Then
                Datos$ = "1-" & CStr(GrhData(i).FileNum) & "-" & CStr(GrhData(i).sX) & "-" & CStr(GrhData(i).sY) & "-" & CStr(GrhData(i).pixelWidth) & "-" & CStr(GrhData(i).pixelHeight)
                
            Else
                Datos$ = CStr(GrhData(i).NumFrames)
                For j = 1 To GrhData(i).NumFrames
                    Datos$ = Datos$ & "-" & CStr(GrhData(i).Frames(j))
                Next
                Datos$ = Datos$ & "-" & CStr(GrhData(i).speed)
            End If
            
            If Len(Datos$) > 0 Then
                Put n, , "Grh" & CStr(i) & "=" & Datos$ & vbCrLf
            End If
        End If
    Next
    Close #n
    
    frmMain.lblstatus.Caption = "Exportado...Graficos.ini"
End Function

Public Function DesindexarCabezas()
'*************************************
'Autor: Lorwik
'Fecha: 05/04/2021
'Descripción: Desindexa las Cabezas de Winter
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Head.ini", vbArchive) = True Then Call Kill(ExporDir & "Head.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumHeads=" & NumHeads & vbCrLf & vbCrLf
    
    For i = 1 To NumHeads
        Datos = Datos & "[HEAD" & (i) & "]" & vbCrLf

        Datos = Datos & "std=" & heads(i).Std & vbCrLf
        Datos = Datos & "FileNum=" & heads(i).texture & vbCrLf
        Datos = Datos & "OffSetX=" & heads(i).startX & vbCrLf
        Datos = Datos & "OffSetY=" & heads(i).startY & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Head.ini"
    DoEvents
    
    Open (ExporDir & "Head.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Head.ini"
End Function

Public Function DesindexarCascos()
'*************************************
'Autor: Lorwik
'Fecha: 05/04/2021
'Descripción: Desindexa los Cascos de Winter
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Helmet.ini", vbArchive) = True Then Call Kill(ExporDir & "Helmet.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumCascos=" & NumCascos & vbCrLf & vbCrLf
    
    For i = 1 To NumCascos
        Datos = Datos & "[CASCO" & (i) & "]" & vbCrLf

        Datos = Datos & "std=" & Cascos(i).Std & vbCrLf
        Datos = Datos & "FileNum=" & Cascos(i).texture & vbCrLf
        Datos = Datos & "OffSetX=" & Cascos(i).startX & vbCrLf
        Datos = Datos & "OffSetY=" & Cascos(i).startY & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Helmet.ini"
    DoEvents
    
    Open (ExporDir & "Helmet.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Helmet.ini"
End Function

Public Function DesindexarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: 26/05/2020
'Descripción: Desindexa los cuerpos
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Personajes.ini", vbArchive) = True Then Call Kill(ExporDir & "Personajes.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumBodies=" & NumCuerpos & vbCrLf & vbCrLf
    
    For i = 1 To NumCuerpos
        Datos = Datos & "[BODY" & (i) & "]" & vbCrLf
        Debug.Print BodyData(i).Walk(n).GrhIndex
        For n = 1 To 4
            Datos = Datos & "WALK" & (n) & "=" & BodyData(i).Walk(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' abajo", "") & IIf(n = 2, Chr(9) & " ' arriba", "") & IIf(n = 3, Chr(9) & " ' izquierda", "") & IIf(n = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
        
        Datos = Datos & "HeadOffsetX=" & BodyData(i).HeadOffset.x & vbCrLf & "HeadOffsetY=" & BodyData(i).HeadOffset.y & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Personajes.ini"
    DoEvents
    
    Open (ExporDir & "Personajes.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Personajes.ini"
End Function

Public Function DesindexarFxs()
'*************************************
'Autor: Lorwik
'Fecha: 05/04/2021
'Descripción: Desindexa los Fxs
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "FXs.ini", vbArchive) = True Then Call Kill(ExporDir & "FXs.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumFxs=" & NumFxs & vbCrLf & vbCrLf
    
    For i = 1 To NumFxs
        If FxData(i).Animacion > 0 Then
            Datos = Datos & "[FX" & (i) & "]" & vbCrLf
            Datos = Datos & "Animacion=" & FxData(i).Animacion & vbCrLf & "OffsetX=" & FxData(i).OffsetX & vbCrLf & "OffsetY=" & FxData(i).OffsetY & vbCrLf & vbCrLf
        End If
    Next
    
    frmMain.lblstatus.Caption = "Guardando...FXs.ini"
    DoEvents
    
    Open (ExporDir & "FXs.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1
    
    DoEvents
    
    frmMain.lblstatus.Caption = "Exportado...FXs.ini"
End Function

Public Function DesindexarColores()
'*************************************
'Autor: Lorwik
'Fecha: 05/04/2021
'Descripción: Desindexa los Colores
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Colores.dat", vbArchive) = True Then Call Kill(ExporDir & "Colores.dat")
    
    Datos = "'Permite customizar los colores de los PJs"
    Datos = Datos & "'todos los valores deben estar entre 0 y 255"
    Datos = Datos & "'los rangos van de 1 a 48 (inclusive). El 0 y el 49,50 estan reservados. Mas arriba son ignorados." & vbCrLf & vbCrLf
    
    For i = 0 To MAXCOLORES
        Datos = Datos & "[" & (i) & "]" & vbCrLf
        Datos = Datos & "R=" & ColoresPJ(i).R & vbCrLf
        Datos = Datos & "G=" & ColoresPJ(i).R & vbCrLf
        Datos = Datos & "B=" & ColoresPJ(i).R & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Colores.dat"
    DoEvents
    
    Open (ExporDir & "Colores.dat") For Binary Access Write As #1
    Put #1, , Datos
    Close #1
    
    DoEvents
    
    frmMain.lblstatus.Caption = "Exportado...Colores.dat"
End Function

Public Function DesindexarAtaques()
'*************************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripción: Desindexa los ataques
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Ataques.ini", vbArchive) = True Then Call Kill(ExporDir & "Personajes.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumAtaques=" & NumCuerpos & vbCrLf & vbCrLf
    
    For i = 1 To NumAtaques
        Datos = Datos & "[ATAQUE" & (i) & "]" & vbCrLf
        Debug.Print AtaqueData(i).Walk(n).GrhIndex
        For n = 1 To 4
            Datos = Datos & "WALK" & (n) & "=" & AtaqueData(i).Walk(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' abajo", "") & IIf(n = 2, Chr(9) & " ' arriba", "") & IIf(n = 3, Chr(9) & " ' izquierda", "") & IIf(n = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
        
        Datos = Datos & "HeadOffsetX=" & AtaqueData(i).HeadOffset.x & vbCrLf & "HeadOffsetY=" & AtaqueData(i).HeadOffset.y & vbCrLf & vbCrLf
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Ataques.ini"
    DoEvents
    
    Open (ExporDir & "Ataques.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Ataques.ini"
End Function

Public Function DesindexarArmas()
'*************************************
'Autor: Lorwik
'Fecha: 11/06/2020
'Descripción: Desindexa las armas
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Armas.ini", vbArchive) = True Then Call Kill(ExporDir & "Armas.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumArmas=" & NumWeaponAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumWeaponAnims
        Datos = Datos & "[Arma" & (i) & "]" & vbCrLf
        For n = 1 To 4
            Datos = Datos & "Dir" & (n) & "=" & WeaponAnimData(i).WeaponWalk(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' abajo", "") & IIf(n = 2, Chr(9) & " ' arriba", "") & IIf(n = 3, Chr(9) & " ' izquierda", "") & IIf(n = 4, Chr(9) & " ' derecha", "") & vbCrLf
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
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    frmMain.lblstatus.Caption = "Exportando..."
    DoEvents
    
    If FileExist(ExporDir & "Escudos.ini", vbArchive) = True Then Call Kill(ExporDir & "Armas.ini")
    
    Datos = "[INIT]" & vbCrLf & "NumEscudos=" & NumEscudosAnims & vbCrLf & vbCrLf
    
    For i = 1 To NumEscudosAnims
        Datos = Datos & "[ESC" & (i) & "]" & vbCrLf
        For n = 1 To 4
            Datos = Datos & "Dir" & (n) & "=" & ShieldAnimData(i).ShieldWalk(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' abajo", "") & IIf(n = 2, Chr(9) & " ' arriba", "") & IIf(n = 3, Chr(9) & " ' izquierda", "") & IIf(n = 4, Chr(9) & " ' derecha", "") & vbCrLf
        Next
    Next
    
    frmMain.lblstatus.Caption = "Guardando...Escudos.ini"
    DoEvents
    
    Open (ExporDir & "Escudos.ini") For Binary Access Write As #1
        Put #1, , Datos
    Close #1
    
    frmMain.lblstatus.Caption = "Exportado...Escudos.ini"
End Function

