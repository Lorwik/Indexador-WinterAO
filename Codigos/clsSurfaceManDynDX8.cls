VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDynDX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D As D3DX8
Private device As Direct3DDevice8

Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
Private mFreeMemoryBytes As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Terminate()
    
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).texture = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
If FileName = 0 Then
Debug.Print "ERROR! GRH = 0"
Exit Function
End If

    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).texture
                Exit Function
            End If
        Next i
    End With

    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    
    Init = True
End Function

Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8
On Error GoTo errhandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    Dim buffer()     As Byte
    Dim InfoHead As INFOHEADER
    
    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            
            #If ModoVisor = 0 Then
        
                Set .texture = mD3D.CreateTextureFromFileEx(device, GraphicsDir & LTrim(Str(Archivo)) & ".png", _
                    D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                    D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
            
            #Else
            
                InfoHead = File_Find(App.Path & "\Recursos\Graficos.WAO", CStr(.FileName) & ".png")
            
                If InfoHead.lngFileSize <> 0 Then
                    Extract_File_Memory Graphics, LCase$(CStr(.FileName) & ".png"), buffer()
                    'Seteamos la textura.
                    Set .texture = mD3D.CreateTextureFromFileInMemoryEx(device, buffer(0), UBound(buffer()) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, _
                                                                             D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                    Erase buffer
                End If
            
            #End If

        
            .texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.width
            .texture_height = texture_info.height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .texture
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.size
        End With
    End With
    
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
errhandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".png"
End Function

Private Function RemoveLRU() As Boolean
    
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
    
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).texture.GetLevelDesc(0, surface_desc)
    
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function



