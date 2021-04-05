Attribute VB_Name = "Mod_General"
Option Explicit

Public Normal_RGBList(3) As Long

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Sub Main()
On Error Resume Next

    'AgregaGrH (1)
    ChDrive App.Path
    ChDir App.Path
    Windows_Temp_Dir = General_Get_Temp_Dir
    Call GenerateContra
    
    AddtoRichTextBox frmCargando.Status, "Cargando Engine Grafico....", 255, 255, 255
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    frmCargando.Show
    frmCargando.Refresh
    
    AddtoRichTextBox frmCargando.Status, "Cargando Rutas", 255, 255, 255
    Call CargarRutas
    
    AddtoRichTextBox frmCargando.Status, "Cargando Index", 255, 255, 255
    Call CargarIndex
    
    AddtoRichTextBox frmCargando.Status, "Cargando Particulas", 255, 255, 255
    Call CargarParticulas
    
    AddtoRichTextBox frmCargando.Status, "Cargando Colores", 255, 255, 255
    Call CargarColores
    
    AddtoRichTextBox frmCargando.Status, "Inicializando Motor Grafico", 255, 255, 255
    Call engine.Engine_Init
    
    AddtoRichTextBox frmCargando.Status, "¡Bienvenido al Indexador de WinterAO - Desarrollado por Lorwik!", 255, 255, 255
    Unload frmCargando
    
    Call IniciarCabecera
                   
    frmMain.Show

    'Inicialización de variables globales
    prgRun = True
    Dim pausa
    pausa = False
    
    engine.Start
    
Exit Sub

End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).B)

General_Particle_Create = engine.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).alphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).Angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)


End Function

Sub CargarParticulas()
'*************************************
'Autor: ????
'Fecha: ????
'Descripción: Cargar el archivo de particulas en memoria
'*************************************

    Dim StreamFile As String
    Dim Loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    
    StreamFile = InitDir & "Particulas.dat"
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    If TotalStreams < 1 Then Exit Sub
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from particle.ini
    For Loopc = 1 To TotalStreams
        StreamData(Loopc).name = General_Var_Get(StreamFile, Val(Loopc), "Name")
        StreamData(Loopc).NumOfParticles = General_Var_Get(StreamFile, Val(Loopc), "NumOfParticles")
        StreamData(Loopc).X1 = General_Var_Get(StreamFile, Val(Loopc), "X1")
        StreamData(Loopc).Y1 = General_Var_Get(StreamFile, Val(Loopc), "Y1")
        StreamData(Loopc).X2 = General_Var_Get(StreamFile, Val(Loopc), "X2")
        StreamData(Loopc).Y2 = General_Var_Get(StreamFile, Val(Loopc), "Y2")
        StreamData(Loopc).Angle = General_Var_Get(StreamFile, Val(Loopc), "Angle")
        StreamData(Loopc).vecx1 = General_Var_Get(StreamFile, Val(Loopc), "VecX1")
        StreamData(Loopc).vecx2 = General_Var_Get(StreamFile, Val(Loopc), "VecX2")
        StreamData(Loopc).vecy1 = General_Var_Get(StreamFile, Val(Loopc), "VecY1")
        StreamData(Loopc).vecy2 = General_Var_Get(StreamFile, Val(Loopc), "VecY2")
        StreamData(Loopc).life1 = General_Var_Get(StreamFile, Val(Loopc), "Life1")
        StreamData(Loopc).life2 = General_Var_Get(StreamFile, Val(Loopc), "Life2")
        StreamData(Loopc).friction = General_Var_Get(StreamFile, Val(Loopc), "Friction")
        StreamData(Loopc).spin = General_Var_Get(StreamFile, Val(Loopc), "Spin")
        StreamData(Loopc).spin_speedL = General_Var_Get(StreamFile, Val(Loopc), "Spin_SpeedL")
        StreamData(Loopc).spin_speedH = General_Var_Get(StreamFile, Val(Loopc), "Spin_SpeedH")
        StreamData(Loopc).alphaBlend = General_Var_Get(StreamFile, Val(Loopc), "AlphaBlend")
        StreamData(Loopc).gravity = General_Var_Get(StreamFile, Val(Loopc), "Gravity")
        StreamData(Loopc).grav_strength = General_Var_Get(StreamFile, Val(Loopc), "Grav_Strength")
        StreamData(Loopc).bounce_strength = General_Var_Get(StreamFile, Val(Loopc), "Bounce_Strength")
        StreamData(Loopc).XMove = General_Var_Get(StreamFile, Val(Loopc), "XMove")
        StreamData(Loopc).YMove = General_Var_Get(StreamFile, Val(Loopc), "YMove")
        StreamData(Loopc).move_x1 = General_Var_Get(StreamFile, Val(Loopc), "move_x1")
        StreamData(Loopc).move_x2 = General_Var_Get(StreamFile, Val(Loopc), "move_x2")
        StreamData(Loopc).move_y1 = General_Var_Get(StreamFile, Val(Loopc), "move_y1")
        StreamData(Loopc).move_y2 = General_Var_Get(StreamFile, Val(Loopc), "move_y2")
        StreamData(Loopc).life_counter = General_Var_Get(StreamFile, Val(Loopc), "life_counter")
        StreamData(Loopc).speed = Val(General_Var_Get(StreamFile, Val(Loopc), "Speed"))
        StreamData(Loopc).NumGrhs = General_Var_Get(StreamFile, Val(Loopc), "NumGrhs")
        
        ReDim StreamData(Loopc).grh_list(1 To StreamData(Loopc).NumGrhs) As Long
        GrhListing = General_Var_Get(StreamFile, Val(Loopc), "Grh_List")
        
        For i = 1 To StreamData(Loopc).NumGrhs
            StreamData(Loopc).grh_list(i) = CLng(General_Field_Read(Str(i), GrhListing, 44))
        Next i
        
        'StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(Loopc), "ColorSet" & ColorSet)
            StreamData(Loopc).colortint(ColorSet - 1).R = General_Field_Read(1, TempSet, 44)
            StreamData(Loopc).colortint(ColorSet - 1).G = General_Field_Read(2, TempSet, 44)
            StreamData(Loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            frmParticleEditor.List2.AddItem Loopc & " - " & StreamData(Loopc).name
    Next Loopc

End Sub
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function
Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, file
End Sub

Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(Text, LastPos + 1)
    End If
End Function

Public Sub AgregaGrH(ByVal numgrh As Long)
    Dim i As Long
    Dim EsteIndex As Long
    Dim CuentaIndex As Long
    
    GrhData(numgrh).FileNum = 1
    GrhData(numgrh).NumFrames = 1
    GrhData(numgrh).pixelHeight = 32
    GrhData(numgrh).pixelWidth = 32
    GrhData(numgrh).Frames(1) = numgrh
    
    CuentaIndex = -1
    frmParticleEditor.lstGrhs.Clear
    For i = 1 To 17925
        If GrhData(i).NumFrames = 1 Then
            frmParticleEditor.lstGrhs.AddItem i
            CuentaIndex = CuentaIndex + 1
        ElseIf GrhData(i).NumFrames > 1 Then
            frmParticleEditor.lstGrhs.AddItem i & " (animacion)"
            CuentaIndex = CuentaIndex + 1
        End If
        If i = numgrh Then
            EsteIndex = CuentaIndex
        End If
    Next i
    frmParticleEditor.lstGrhs.ListIndex = EsteIndex
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal File_Type As VbFileAttribute) As Boolean
    If Dir(file_path, File_Type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Sub HookSurfaceHwnd(pic As Form)
    Call ReleaseCapture
    Call SendMessage(pic.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
 
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
