VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module
Private list() As String
Private listSize As Integer


Public Property Get getList() As Variant
    getList = list
End Property

Public Property Let letList(newList)
    list = newList
    Me.Size
End Property


' Event - triggered when class created
Private Sub Class_Initialize()
    listSize = 0
End Sub


Public Sub Sort()
    Dim indexReorded As Integer
    Dim cpyList() As String
    
    ReDim cpyList(listSize - 1)
    For Position = 0 To listSize - 1
        For i = 0 To listSize - 1
            If Not (list(i) = "") Then
                If (cpyList(Position) = "") Then
                    cpyList(Position) = list(i)
                End If
                If StrComp(cpyList(Position), list(i), 1) >= 0 Then
                    cpyList(Position) = list(i)
                    indexReorded = i
                End If
            End If
        Next
        list(indexReorded) = ""
    Next
    list = cpyList
End Sub

Public Sub SortUpSideDown()
    Dim indexReorded As Integer
    Dim cpyList() As String
    
    ReDim cpyList(listSize - 1)
    For Position = listSize - 1 To 0 Step -1
        For i = listSize - 1 To 0 Step -1
            If Not (list(i) = "") Then
                If (cpyList(Position) = "") Then
                    cpyList(Position) = list(i)
                End If
                If StrComp(cpyList(Position), list(i), 1) >= 0 Then
                    cpyList(Position) = list(i)
                    indexReorded = i
                End If
            End If
        Next
        list(indexReorded) = ""
    Next
    list = cpyList
End Sub

Public Function Size()
    If Not IsNull(list) Then
        On Error GoTo ErrorHandler
        listSize = UBound(list, 1) - LBound(list, 1) + 1
    Else
        listSize = 0
    End If
    Size = listSize
    
    Exit Function
ErrorHandler:
    listSize = 0
    Size = listSize
End Function

Public Sub Add(item As Variant)
    If Not IsNull(item) Then
        ReDim Preserve list(listSize)
        listSize = listSize + 1
        list(listSize - 1) = item
    End If
End Sub
Public Sub AddArray(arrayToAdd As Variant)
    Dim i As Integer
    
    sizeToAdd = SizeOfArray(arrayToAdd)
    If sizeToAdd > 0 Then
        ReDim Preserve list(listSize + sizeToAdd - 1)
        
        i = 0
        For Each item In arrayToAdd
            list(listSize + i) = item
            i = i + 1
        Next
        listSize = listSize + sizeToAdd
    End If
End Sub

Public Sub RemoveStr(item As String)
    Dim i As Integer
    Dim cpyList() As String
    Dim index As Integer
    index = Me.Find(item)
    If index < listSize Then
        If listSize = 1 Then
            Erase list
            listSize = 0
            Exit Sub
        End If
        ReDim cpyList(listSize - 2)
        
        For i = 0 To index - 1
            cpyList(i) = list(i)
        Next
            
        For i = index + 1 To listSize - 1
            cpyList(i - 1) = list(i)
        Next
        listSize = listSize - 1
        ReDim list(listSize - 1)
        list = cpyList
    End If
End Sub

Public Sub Remove(index As Integer)
    Dim i As Integer
    Dim cpyList() As String
    
    If index < listSize Then
        If listSize = 1 Then
            Erase list
            listSize = 0
            Exit Sub
        End If
        ReDim cpyList(listSize - 2)
        
        For i = 0 To index - 1
            cpyList(i) = list(i)
        Next
            
        For i = index + 1 To listSize - 1
            cpyList(i - 1) = list(i)
        Next
        listSize = listSize - 1
        ReDim list(listSize - 1)
        list = cpyList
    End If
End Sub

Public Function retrieve(index As Integer) As String
    Dim i As Integer
    
    If index < listSize Then
        retrieve = list(index)
    End If
End Function

Public Sub RemoveDuplicates()
    Dim i As Integer
    Dim duplListCounter As Integer
    
    duplListCounter = 0
    If listSize > 1 Then
        For Each cv In list
            If IsEmpty(cv) Then
                Exit For
            ElseIf cv <> "" Then
                firstMatch = True
                For i = 0 To listSize - 1
                    If list(i) = cv Then
                        If Not firstMatch Then
                            duplListCounter = duplListCounter + 1
                            list(i) = ""
                        End If
                        firstMatch = False
                    End If
                Next
            End If
        Next
        Me.Sort
        listSize = listSize - duplListCounter
        ReDim Preserve list(listSize - 1)
    End If
End Sub

Public Function Contains(stringToBeFound) As Boolean
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = 0 To Me.Size - 1
            If list(i) = stringToBeFound Then
                Contains = True
                Exit Function
            End If
        Next i
    End If
    Contains = False
End Function

Public Function Find(stringToBeFound)
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = LBound(list) To UBound(list)
            If list(i) = stringToBeFound Then
                Find = i
                Exit Function
            End If
        Next i
    End If
    Find = -1
End Function
