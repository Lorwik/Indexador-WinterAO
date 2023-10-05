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

Public Function General_Particle_Create(ByVal ParticulaInd As Long, _
                                        ByVal x As Integer, _
                                        ByVal y As Integer, _
                                        Optional ByVal particle_life As Long = 0) As Long

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

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
                        
'******************************************************************
' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
  
On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
    'Call RegistrarError(Err.Number, Err.Description, "Buscar_Carpeta", Erl)
  
End Function
