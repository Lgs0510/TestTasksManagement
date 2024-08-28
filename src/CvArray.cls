VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CvArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module

Dim arrayVar() As New ScriptNameObj
Private ArraySize As Integer, trueArraySize As Integer
Private arrayGrowingStep As Integer

Public Property Get getArray() As Variant
    getArray = arrayVar
End Property

Public Property Let letArray(newArray)
    arrayVar = newArray
    Me.Size
End Property


' Event - triggered when class created
Private Sub Class_Initialize()
    ArraySize = 0
    arrayGrowingStep = 1000
    ReDim arrayVar(arrayGrowingStep)
    Me.Size
End Sub

Public Sub Sort()
    Dim indexReorded As Integer
    Dim cpyList() As New ScriptNameObj
    
    ReDim cpyList(ArraySize - 1)
    For Position = 0 To ArraySize - 1
        For i = 0 To ArraySize - 1
            If Not (arrayVar(i).cvNumber = "") Then
                If (cpyList(Position).cvNumber = "") Then
                    cpyList(Position).cvNumber = arrayVar(i).cvNumber
                    cpyList(Position).ScriptName = arrayVar(i).ScriptName
                End If
                If StrComp(cpyList(Position).cvNumber, arrayVar(i).cvNumber, 1) >= 0 Then
                    cpyList(Position).cvNumber = arrayVar(i).cvNumber
                    cpyList(Position).ScriptName = arrayVar(i).ScriptName
                    indexReorded = i
                End If
            End If
        Next
        arrayVar(indexReorded).cvNumber = ""
        arrayVar(indexReorded).ScriptName = ""
    Next
    arrayVar = cpyList
End Sub

Public Function Size()
    If Not IsNull(arrayVar) Then
        On Error GoTo ErrorHandler
        trueArraySize = UBound(arrayVar, 1) - LBound(arrayVar, 1) + 1
        If ArraySize > trueArraySize Then
            ArraySize = trueArraySize
        End If
    Else
        trueArraySize = 0
        ArraySize = 0
    End If
    Size = ArraySize
    
    Exit Function
ErrorHandler:
    ArraySize = 0
    Size = ArraySize
End Function

Public Sub Add(item)
    If Not IsNull(item) Then
        If ArraySize >= trueArraySize Then
            ReDim Preserve arrayVar(trueArraySize + arrayGrowingStep)
            Me.Size
        End If
        arrayVar(ArraySize).Add item
        ArraySize = ArraySize + 1
    End If
End Sub


Public Sub Remove(index As Integer)
    Dim i As Integer
    Dim cpyList() As String
    
    If index < ArraySize Then
        If ArraySize > 0 Then
            ReDim cpyList(ArraySize - 1)
            
            For i = 0 To index - 1
                cpyList(i) = arrayVar(i)
            Next
                
            For i = index + 1 To ArraySize - 1
                cpyList(i - 1) = arrayVar(i)
            Next
            ArraySize = ArraySize - 1
            ScriptNameObj = cpyList
        End If
    End If
End Sub
Public Sub RemoveDuplicates()
    Dim i As Integer
    Dim duplListCounter As Integer
    
    duplListCounter = 0
    If ArraySize > 1 Then
        For Each cv In arrayVar
            firstMatch = True
            For i = 0 To ArraySize - 2
                If cv.cvNumber <> "" Then
                    If arrayVar(i).cvNumber = cv.cvNumber Then
                        If Not firstMatch Then
                            duplListCounter = duplListCounter + 1
                            arrayVar(i).cvNumber = ""
                            arrayVar(i).ScriptName = ""
                        End If
                        firstMatch = False
                    End If
                End If
            Next
        Next
        Me.Sort
        ArraySize = ArraySize - duplListCounter
        ReDim Preserve arrayVar(ArraySize - 1)
    End If
End Sub

Public Function Contains(stringToBeFound)
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = 0 To Me.Size - 1
            If arrayVar(i) = stringToBeFound Then
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
        For i = LBound(arrayVar) To UBound(arrayVar)
            If arrayVar(i).cvNumber = stringToBeFound Then
                Find = i
                Exit Function
            End If
        Next i
    End If
    Find = -1
End Function

Public Function GetCV(index As Integer) As String
    GetCV = arrayVar(index).cvNumber
End Function
Public Function GetScriptName(index As Integer) As String
    GetScriptName = arrayVar(index).ScriptName
End Function

