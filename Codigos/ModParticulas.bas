Attribute VB_Name = "ModParticulas"
Option Explicit

'--> Current StreamFile <--
Public CurStreamFile As String

Public Sub NuevaParticula()
#If ModoVisor = 0 Then

    Dim Nombre As String
    Dim NewStreamNumber As Integer
    Dim grhlist(0) As Long
    Dim loopc As Long
    
    'Get name for new stream
    Nombre = InputBox("Por favor inserte un nombre a la particula", "New Stream")
    
    If Nombre = "" Then Exit Sub
    
    'Set new stream #
    NewStreamNumber = frmParticleEditor.List2.ListCount + 1
    
    'Add stream to combo box
    frmParticleEditor.List2.AddItem NewStreamNumber & " - " & Nombre
    
    'Add 1 to TotalStreams
    TotalStreams = TotalStreams + 1
    
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'Add stream data to StreamData array
    StreamData(NewStreamNumber).name = frmParticleEditor.name
    StreamData(NewStreamNumber).NumOfParticles = 126
    StreamData(NewStreamNumber).x1 = 0
    StreamData(NewStreamNumber).y1 = 0
    StreamData(NewStreamNumber).x2 = 0
    StreamData(NewStreamNumber).y2 = 0
    StreamData(NewStreamNumber).angle = 0
    StreamData(NewStreamNumber).vecx1 = -20
    StreamData(NewStreamNumber).vecx2 = 20
    StreamData(NewStreamNumber).vecy1 = -20
    StreamData(NewStreamNumber).vecy2 = 20
    StreamData(NewStreamNumber).life1 = 10
    StreamData(NewStreamNumber).life2 = 50
    StreamData(NewStreamNumber).friction = 8
    StreamData(NewStreamNumber).spin_speedL = 0.1
    StreamData(NewStreamNumber).spin_speedH = 0.1
    StreamData(NewStreamNumber).grav_strength = 2
    StreamData(NewStreamNumber).bounce_strength = -5
    StreamData(NewStreamNumber).alphaBlend = 1
    StreamData(NewStreamNumber).gravity = 0
    
    
    'Select the new stream type in the combo box
    frmParticleEditor.List2.ListIndex = NewStreamNumber - 1
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

Public Sub GuardarParticulas()
#If ModoVisor = 0 Then

    Dim loopc As Long
    Dim StreamFile As String
    Dim Bypass As Boolean
    Dim RetVal
    CurStreamFile = InitDir & "Particulas.dat"
    
    If FileExist(CurStreamFile, vbNormal) = True Then
        RetVal = MsgBox("¡El archivo " & CurStreamFile & " ya existe!" & vbCrLf & "¿Deseas sobreescribirlo?", vbYesNoCancel Or vbQuestion)
        If RetVal = vbNo Then
            Bypass = False
        ElseIf RetVal = vbCancel Then
            Exit Sub
        ElseIf RetVal = vbYes Then
            StreamFile = CurStreamFile
            Bypass = True
        End If
    End If
    
    If Bypass = False Then
    
        StreamFile = CurStreamFile
        
        If FileExist(StreamFile, vbNormal) = True Then
            RetVal = MsgBox("¡El archivo " & StreamFile & " ya existe!" & vbCrLf & "¿Desea sobreescribirlo?", vbYesNo Or vbQuestion)
            If RetVal = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Dim GrhListing As String
    Dim i As Long
    
    'Check for existing data file and kill it
    If FileExist(StreamFile, vbNormal) Then Kill StreamFile
    
    'Write particle data to particle.ini
    General_Var_Write StreamFile, "INIT", "Total", Val(TotalStreams)
    
    For loopc = 1 To TotalStreams
        General_Var_Write StreamFile, Val(loopc), "Name", StreamData(loopc).name
        General_Var_Write StreamFile, Val(loopc), "NumOfParticles", Val(StreamData(loopc).NumOfParticles)
        General_Var_Write StreamFile, Val(loopc), "X1", Val(StreamData(loopc).x1)
        General_Var_Write StreamFile, Val(loopc), "Y1", Val(StreamData(loopc).y1)
        General_Var_Write StreamFile, Val(loopc), "X2", Val(StreamData(loopc).x2)
        General_Var_Write StreamFile, Val(loopc), "Y2", Val(StreamData(loopc).y2)
        General_Var_Write StreamFile, Val(loopc), "Angle", Val(StreamData(loopc).angle)
        General_Var_Write StreamFile, Val(loopc), "VecX1", Val(StreamData(loopc).vecx1)
        General_Var_Write StreamFile, Val(loopc), "VecX2", Val(StreamData(loopc).vecx2)
        General_Var_Write StreamFile, Val(loopc), "VecY1", Val(StreamData(loopc).vecy1)
        General_Var_Write StreamFile, Val(loopc), "VecY2", Val(StreamData(loopc).vecy2)
        General_Var_Write StreamFile, Val(loopc), "Life1", Val(StreamData(loopc).life1)
        General_Var_Write StreamFile, Val(loopc), "Life2", Val(StreamData(loopc).life2)
        General_Var_Write StreamFile, Val(loopc), "Friction", Val(StreamData(loopc).friction)
        General_Var_Write StreamFile, Val(loopc), "Spin", Val(StreamData(loopc).spin)
        General_Var_Write StreamFile, Val(loopc), "Spin_SpeedL", Val(StreamData(loopc).spin_speedL)
        General_Var_Write StreamFile, Val(loopc), "Spin_SpeedH", Val(StreamData(loopc).spin_speedH)
        General_Var_Write StreamFile, Val(loopc), "Grav_Strength", Val(StreamData(loopc).grav_strength)
        General_Var_Write StreamFile, Val(loopc), "Bounce_Strength", Val(StreamData(loopc).bounce_strength)
        
        General_Var_Write StreamFile, Val(loopc), "AlphaBlend", Val(StreamData(loopc).alphaBlend)
        General_Var_Write StreamFile, Val(loopc), "Gravity", Val(StreamData(loopc).gravity)
        
        General_Var_Write StreamFile, Val(loopc), "XMove", Val(StreamData(loopc).XMove)
        General_Var_Write StreamFile, Val(loopc), "YMove", Val(StreamData(loopc).YMove)
        General_Var_Write StreamFile, Val(loopc), "move_x1", Val(StreamData(loopc).move_x1)
        General_Var_Write StreamFile, Val(loopc), "move_x2", Val(StreamData(loopc).move_x2)
        General_Var_Write StreamFile, Val(loopc), "move_y1", Val(StreamData(loopc).move_y1)
        General_Var_Write StreamFile, Val(loopc), "move_y2", Val(StreamData(loopc).move_y2)
        General_Var_Write StreamFile, Val(loopc), "life_counter", Val(StreamData(loopc).life_counter)
        General_Var_Write StreamFile, Val(loopc), "Speed", Str(StreamData(loopc).speed)
        
        General_Var_Write StreamFile, Val(loopc), "NumGrhs", Val(StreamData(loopc).NumGrhs)
        
        GrhListing = vbNullString
        For i = 1 To StreamData(loopc).NumGrhs
            GrhListing = GrhListing & StreamData(loopc).grh_list(i) & ","
        Next i
        
        General_Var_Write StreamFile, Val(loopc), "Grh_List", GrhListing
        
        General_Var_Write StreamFile, Val(loopc), "ColorSet1", StreamData(loopc).colortint(0).r & "," & StreamData(loopc).colortint(0).g & "," & StreamData(loopc).colortint(0).b
        General_Var_Write StreamFile, Val(loopc), "ColorSet2", StreamData(loopc).colortint(1).r & "," & StreamData(loopc).colortint(1).g & "," & StreamData(loopc).colortint(1).b
        General_Var_Write StreamFile, Val(loopc), "ColorSet3", StreamData(loopc).colortint(2).r & "," & StreamData(loopc).colortint(2).g & "," & StreamData(loopc).colortint(2).b
        General_Var_Write StreamFile, Val(loopc), "ColorSet4", StreamData(loopc).colortint(3).r & "," & StreamData(loopc).colortint(3).g & "," & StreamData(loopc).colortint(3).b
        
    Next loopc
    
    'Report the results
    If TotalStreams > 1 Then
        MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
    Else
        MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
    End If
    
    'Set DataChanged variable to false
    DataChanged = False
    CurStreamFile = StreamFile
    
#Else
    MsgBox "Esta opción no esta disponible en el modo visor.", vbCritical
#End If
End Sub

Sub CargarParticulasLista()
    Dim loopc As Long
    Dim DataTemp As Boolean
    DataTemp = DataChanged
    
    'Set the values
    frmParticleEditor.txtPCount.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).NumOfParticles
    frmParticleEditor.txtX1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).x1
    frmParticleEditor.txtY1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).y1
    frmParticleEditor.txtX2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).x2
    frmParticleEditor.txtY2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).y2
    frmParticleEditor.txtAngle.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).angle
    frmParticleEditor.vecx1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).vecx1
    frmParticleEditor.vecx2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).vecx2
    frmParticleEditor.vecy1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).vecy1
    frmParticleEditor.vecy2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).vecy2
    frmParticleEditor.life1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).life1
    frmParticleEditor.life2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).life2
    frmParticleEditor.fric.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).friction
    frmParticleEditor.chkSpin.value = StreamData(frmParticleEditor.List2.ListIndex + 1).spin
    frmParticleEditor.spin_speedL.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).spin_speedL
    frmParticleEditor.spin_speedH.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).spin_speedH
    frmParticleEditor.txtGravStrength.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).grav_strength
    frmParticleEditor.txtBounceStrength.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).bounce_strength
    frmParticleEditor.chkAlphaBlend.value = StreamData(frmParticleEditor.List2.ListIndex + 1).alphaBlend
    frmParticleEditor.chkGravity.value = StreamData(frmParticleEditor.List2.ListIndex + 1).gravity
    frmParticleEditor.chkXMove.value = StreamData(frmParticleEditor.List2.ListIndex + 1).XMove
    frmParticleEditor.chkYMove.value = StreamData(frmParticleEditor.List2.ListIndex + 1).YMove
    frmParticleEditor.move_x1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).move_x1
    frmParticleEditor.move_x2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).move_x2
    frmParticleEditor.move_y1.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).move_y1
    frmParticleEditor.move_y2.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).move_y2
    
    If StreamData(frmParticleEditor.List2.ListIndex + 1).life_counter = -1 Then
        frmParticleEditor.life.Enabled = False
        frmParticleEditor.chkNeverDies.value = vbChecked
    Else
        frmParticleEditor.life.Enabled = True
        frmParticleEditor.life.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).life_counter
        frmParticleEditor.chkNeverDies.value = vbUnchecked
    End If
    
    frmParticleEditor.speed.Text = StreamData(frmParticleEditor.List2.ListIndex + 1).speed
    
    frmParticleEditor.lstSelGrhs.Clear
    
    For loopc = 1 To StreamData(frmParticleEditor.List2.ListIndex + 1).NumGrhs
        frmParticleEditor.lstSelGrhs.AddItem StreamData(frmParticleEditor.List2.ListIndex + 1).grh_list(loopc)
    Next loopc
    
    DataChanged = DataTemp
    
    indexs = frmParticleEditor.List2.ListIndex + 1
    
    General_Particle_Create indexs, 50, 50

End Sub
