VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type ChildNode
    Key As String
    value As String
End Type

Private Type MainNode
    name As String
    values() As ChildNode
    numValues As Long
End Type

Private fileData() As MainNode
Private MainNodes As Long

Private Sub Class_Terminate()
    Dim i As Long
    
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        Erase fileData
    End If
End Sub

Public Sub Initialize(ByVal file As String)

    Dim handle As Integer
    Dim Text As String
    Dim Pos As Long
    
    handle = FreeFile
    
    Open file For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, Text
        
        If Len(Text) Then
            If Left$(Text, 1) = "[" Then
                Pos = InStr(2, Text, "]")
                If Pos Then
                    ReDim Preserve fileData(MainNodes) As MainNode
                    fileData(MainNodes).name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    MainNodes = MainNodes + 1
                End If
            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, Text, "=")
                If Pos Then
                    'Is it under any main node??
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As ChildNode
                            
                            .values(.numValues).value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).Key = UCase$(Left$(Text, Pos - 1))
                            
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close handle
    
    Dim i As Long
    
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then _
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
End Sub

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Long, ByVal Last As Long)

    Dim min As Long      'First item in the list
    Dim max As Long      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As ChildNode
    
    min = First
    max = Last
    
    With Node
        comp = .values((min + max) \ 2).Key
    
        Do While min <= max
            Do While .values(min).Key < comp And min < Last
                min = min + 1
            Loop
            Do While .values(max).Key > comp And max > First
                max = max - 1
            Loop
            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    
    If First < max Then SortChildNodes Node, First, max
    If min < Last Then SortChildNodes Node, min, Last
End Sub

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)

    Dim min As Long      'First item in the list
    Dim max As Long      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As MainNode
    
    min = First
    max = Last
    
    comp = fileData((min + max) \ 2).name
    
    Do While min <= max
        Do While fileData(min).name < comp And min < Last
            min = min + 1
        Loop
        Do While fileData(max).name > comp And max > First
            max = max - 1
        Loop
        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If First < max Then SortMainNodes First, max
    If min < Last Then SortMainNodes min, Last
End Sub

Public Function GetValue(ByVal Main As String, ByVal Key As String) As String
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        'If valid, binary search among keys
        j = FindKey(fileData(i), UCase$(Key))
        
        'If we found it we return it
        If j >= 0 Then GetValue = fileData(i).values(j).value
    End If
End Function

Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = Node.numValues - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If Node.values(mid).Key < Key Then
            min = mid + 1
        ElseIf Node.values(mid).Key > Key Then
            max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop

    FindKey = Not mid
End Function

Private Function FindMain(ByVal name As String) As Long
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If fileData(mid).name < name Then
            min = mid + 1
        ElseIf fileData(mid).name > name Then
            max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop

    FindMain = Not mid
End Function




