Attribute VB_Name = "ModIndexacionAutomatica"
Option Explicit

Public Function BuscarConsecutivo(ByVal Libres As Integer) As String
    Dim i As Long
    Dim Conta As Long
    If IsNumeric(Libres) = False Then Exit Function
    For i = 1 To grhCount
        If GrhData(i).NumFrames = 0 Then
            Conta = Conta + 1
            If Conta = Libres Then
                BuscarConsecutivo = "Desde Grh" & i - (Conta - 1) & " hasta Grh" & i & " se encuentran libres."
                Exit Function
            End If
        ElseIf Conta > 0 Then
            Conta = 0
        End If
    Next
    BuscarConsecutivo = "No se encontraron " & Libres & " GRH Libres Consecutivos"
End Function

Public Function BuscarConsecutivoAutoIndex(ByVal Libres As Integer) As String
    Dim i As Integer
    Dim Conta As Integer
    If IsNumeric(Libres) = False Then Exit Function
    For i = 1 To grhCount
        If GrhData(i).NumFrames = 0 Then
            Conta = Conta + 1
            If Conta = Libres Then
                BuscarConsecutivoAutoIndex = i - (Conta - 1)
                Exit Function
            End If
        ElseIf Conta > 0 Then
            Conta = 0
        End If
    Next
    BuscarConsecutivoAutoIndex = "Nada"
End Function

Public Function AutoIndex_Cuerpos()
    On Error Resume Next
    Dim FileNum As Long
    Dim ImgAlto As Integer
    Dim ImgAncho As Integer
    Dim AnchoAnim As Integer
    Dim AltoAnim As Integer
    Dim resultado As String
    Dim GrhConse As String
    Dim GrhAUsar As Long
    Dim i, j, X, Y As Integer
    
    'Necesitamos información sobre la imagen
    FileNum = Int(InputBox("Indica el numero de la imagen"))
    ImgAlto = Int(InputBox("Indica el Alto de la imagen (181 por defecto)"))
    ImgAncho = Int(InputBox("Indica el Ancho de la imagen (149 por defecto)"))
        
    AnchoAnim = ImgAncho / 6
    AltoAnim = ImgAlto / 4
    
    '¿Hay Grh libres para usar?
    GrhConse = BuscarConsecutivoAutoIndex(26)
    
    If GrhConse = "Nada" Then
        GrhAUsar = grhCount + 1
    Else 'Si no los hay, utilizamos un nuevo Grh
        GrhAUsar = CLng(GrhConse)
    End If
    
    'Inicializamos posiciones
    X = 0
    Y = 0
    
    'Vamos a recorrer las 2 primeras lineas
    For i = 0 To 1
        'Recorremos la animacion
        For j = 0 To 5
            resultado = resultado + "Grh" & GrhAUsar & "=1-" & FileNum & "-" & X & "-" & Y & "-" & AnchoAnim & "-" & AltoAnim & vbCrLf
            GrhAUsar = GrhAUsar + 1
            X = X + AnchoAnim
        Next j
        
        resultado = resultado + "Grh" & GrhAUsar & "=6-" & GrhAUsar - 6 & "-" & GrhAUsar - 5 & "-" & GrhAUsar - 4 & "-" & GrhAUsar - 3 & "-" & GrhAUsar - 2 & "-" & GrhAUsar - 1 & "-555" & vbCrLf
        GrhAUsar = GrhAUsar + 1
        X = 0
        Y = Y + AltoAnim
    Next i
    
    'Reseteamos las posiciones
    X = 0
    Y = Y + AltoAnim 'Pero a este le sumamos
    
    'Vamos a recorrer las 2 ultimas lineas
    For i = 0 To 1
        'Recorremos la animacion
        For j = 0 To 4
            resultado = resultado + "Grh" & GrhAUsar & "=1-" & FileNum & "-" & X & "-" & Y & "-" & AnchoAnim & "-" & AltoAnim & vbCrLf
            GrhAUsar = GrhAUsar + 1
            X = X + AnchoAnim
        Next j
        
        resultado = resultado + "Grh" & GrhAUsar & "=5-" & GrhAUsar - 5 & "-" & GrhAUsar - 4 & "-" & GrhAUsar - 3 & "-" & GrhAUsar - 2 & "-" & GrhAUsar - 1 & "-555" & vbCrLf
        GrhAUsar = GrhAUsar + 1
        X = 0
        Y = Y + AltoAnim
    Next i

    frmresultado.Show
    frmresultado.txtResultado.Text = resultado
End Function
