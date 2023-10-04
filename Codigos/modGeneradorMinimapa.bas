Attribute VB_Name = "modGeneradorMinimapa"
Option Explicit

Public Const COLOR_KEY As Long = &H0

Public Sub GenerarMinimapa()
    Dim file As Long
    If Trabajando Then Exit Sub
    
    Trabajando = True
    
    Dim i As Long
    
    If FileExist(InitDir & "\minimap.dat", vbArchive) Then Kill InitDir & "\minimap.dat"
    
    file = FreeFile
    Open InitDir & "\minimap.dat" For Binary As #file
    
    'Abrimos el generador de Minimap
    frmGeneradorMinimap.Show
    
    For i = 1 To grhCount
        frmGeneradorMinimap.lblstatus.Text = "Grh " & i & "/" & UBound(GrhData())
        If Grh_Check(i) Then
            Put #file, , Grh_get_value(i, frmGeneradorMinimap.ucImage.hdc, 0, 0, False)
        End If

        DoEvents
    Next i
    
    Close #file
    
    frmGeneradorMinimap.lblstatus.Text = "Minimap.dat generado con exito!"
    
    Trabajando = False
End Sub

Function Grh_get_value(ByVal grh_index As Long, ByVal destHDC As Long, ByVal screen_x As Long, ByVal screen_y As Long, Optional transparent As Boolean = False) As Long
'**************************************************************
'Author: David Justus
'Last Modify Date: 10/09/2004
'Modified by Juan Martín Sotuyo Dodero
'*************************************************************
    Dim x As Long
    Dim y As Long
    Dim file_path As String
    Dim src_x As Long
    Dim src_y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long
    Dim OldObj As Long
    Dim Value As Currency
    
    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
    
    file_path = GraphicsDir & "\" & GrhData(grh_index).FileNum & ".png"
    
    If Not FileExist(file_path, vbArchive) Then Exit Function
    
    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
    
    hdcsrc = CreateCompatibleDC(destHDC)
    OldObj = SelectObject(hdcsrc, frmGeneradorMinimap.ucImage.LoadImageFromFile(file_path))
    
    BitBlt destHDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    
    DeleteObject SelectObject(hdcsrc, OldObj)
    DeleteDC hdcsrc
    
    DoEvents
    
    Dim R As Currency
    Dim B As Currency
    Dim G As Currency
    Dim TempR As Integer
    Dim TempG As Integer
    Dim TempB As Integer
    Dim InvalidPixels As Long
    
    For x = 0 To GrhData(grh_index).pixelHeight - 1
        For y = 0 To GrhData(grh_index).pixelWidth - 1
            'Color is not taken into account if the color is transparent
            If GetPixel(destHDC, x, y) = COLOR_KEY Then
                InvalidPixels = InvalidPixels + 1
            Else
                General_Long_Color_to_RGB GetPixel(destHDC, x, y), TempR, TempG, TempB
                R = R + TempR
                G = G + TempG
                B = B + TempB
            End If
            DoEvents
        Next y
    Next x
    
    Dim size As Long
    
    size = src_height * src_width - InvalidPixels
    
    If size = 0 Then size = 1
    Grh_get_value = RGB(B / size, G / size, R / size)
End Function

Public Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grhCount Then
        If GrhData(grh_index).active Then
            Grh_Check = True
        End If
    End If
End Function

Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
'***********************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 2/19/03
'Takes a long value and separates RGB values to the given variables
'***********************************
    Dim temp_color As String
    
    temp_color = Hex(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid(temp_color, 1, 2))
    green = CLng("&H" + mid(temp_color, 3, 2))
    blue = CLng("&H" + mid(temp_color, 5, 2))

End Sub
