VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCasesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module

Dim testLists() As New TestCaseObj
Private listSize As Integer, trueArraySize As Integer
Private arrayGrowingStep As Integer

Public Property Get getArray() As Variant
    getArray = testLists
End Property

Public Property Let letArray(newArray)
    testLists = newArray
    listSize = SizeOfArray(newArray) - 1
    Me.Size
End Property


' Event - triggered when class created
Private Sub Class_Initialize()
    listSize = 0
    arrayGrowingStep = 1000
    ReDim testLists(arrayGrowingStep)
    Me.Size
End Sub

Public Sub Sort()
    Dim indexReorded As Integer
    Dim cpyList() As New TestCaseObj
    
    ReDim cpyList(listSize - 1)
    For Position = 0 To listSize - 1
        For i = 0 To listSize - 1
            If Not (testLists(i).cvNumber = "") Then
                If (cpyList(Position).cvNumber = "") Then
                    cpyList(Position).cvNumber = testLists(i).cvNumber
                    cpyList(Position).testStatus = testLists(i).testStatus
                    cpyList(Position).cvOld = testLists(i).cvOld
                End If
                If StrComp(cpyList(Position).cvNumber, testLists(i).cvNumber, 1) >= 0 Then
                    cpyList(Position).cvNumber = testLists(i).cvNumber
                    cpyList(Position).testStatus = testLists(i).testStatus
                    cpyList(Position).cvOld = testLists(i).cvOld
                    indexReorded = i
                End If
            End If
        Next
        testLists(indexReorded).cvNumber = ""
        testLists(indexReorded).testStatus = ""
        testLists(indexReorded).cvOld = ""
    Next
    testLists = cpyList
End Sub

Public Function Size()
    If Not IsNull(testLists) Then
        On Error GoTo ErrorHandler
        trueArraySize = UBound(testLists, 1) - LBound(testLists, 1) + 1
        If listSize > trueArraySize Then
            listSize = trueArraySize
        End If
    Else
        trueArraySize = 0
        listSize = 0
    End If
    Size = listSize
    
    Exit Function
ErrorHandler:
    listSize = 0
    Size = listSize
End Function

Public Sub Add(item)
    If Not IsNull(item) Then
        If listSize >= trueArraySize Then
            ReDim Preserve testLists(trueArraySize + arrayGrowingStep)
            Me.Size
        End If
        testLists(listSize).Add item
        listSize = listSize + 1
    End If
End Sub


Public Sub Remove(index As Integer)
    Dim i As Integer
    Dim cpyList() As New TestCaseObj
    
    If index < listSize Then
        If listSize > 0 Then
            ReDim cpyList(listSize - 1)
            
            For i = 0 To index - 1
                cpyList(i).letTestCase = testLists(i)
            Next
                
            For i = index + 1 To listSize - 1
                cpyList(i - 1).letTestCase = testLists(i)
            Next
            listSize = listSize - 1
            testLists = cpyList
        End If
    End If
End Sub
Public Sub RemoveDuplicates()
    Dim i As Integer
    Dim duplListCounter As Integer
    
    duplListCounter = 0
    If listSize > 1 Then
        For Each cv In testLists
            firstMatch = True
            If cv Is Nothing Then
                Exit For
            End If
            For i = 0 To listSize - 2
                If cv.cvNumber <> "" Then
                    If testLists(i).cvNumber = cv.cvNumber Then
                        If Not firstMatch Then
                            duplListCounter = duplListCounter + 1
                            testLists(i).cvNumber = ""
                            testLists(i).testStatus = ""
                            testLists(i).cvOld = ""
                        End If
                        firstMatch = False
                    End If
                End If
            Next
        Next
        Me.Sort
        listSize = listSize - duplListCounter
        ReDim Preserve testLists(listSize - 1)
    End If
End Sub

Public Function Contains(stringToBeFound)
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = 0 To Me.Size - 1
            If testLists(i).cvNumber = stringToBeFound Then
                Contains = True
                Exit Function
            End If
        Next i
    End If
    Contains = False
End Function
Public Function ContainsOldCv(stringToBeFound)
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = 0 To Me.Size - 1
            If testLists(i).cvOld = stringToBeFound Then
                ContainsOldCv = True
                Exit Function
            End If
        Next i
    End If
    ContainsOldCv = False
End Function

Public Function Find(stringToBeFound As String)
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = LBound(testLists) To UBound(testLists)
            If StrComp(testLists(i).cvNumber, stringToBeFound, 1) = 0 Then
                Find = i
                Exit Function
            End If
        Next i
    End If
    Find = -1
End Function
Public Function FindOldCv(stringToBeFound As String) As Integer
    Dim i
    If Not IsNull(stringToBeFound) Then
        For i = LBound(testLists) To UBound(testLists)
            If StrComp(testLists(i).cvOld, stringToBeFound, 1) = 0 Then
                FindOldCv = i
                Exit Function
            End If
        Next i
    End If
    FindOldCv = -1
End Function
Public Function GetCV(index As Integer) As String
    GetCV = testLists(index).cvNumber
End Function

Public Function GetScriptName(index As Integer) As String
    GetScriptName = testLists(index).testStatus
End Function
Public Function CountApprovedTests() As Integer
    Dim count As Integer
    count = 0
    For Each cv In testLists
        If cv.testStatus = "OK" Then
            count = count + 1
        End If
    Next
    CountApprovedTests = count
End Function

Public Function CountReprovedTests() As Integer
    Dim count As Integer
    count = 0
    For Each cv In testLists
        If cv.testStatus = "NOK" Then
            count = count + 1
        End If
    Next
    CountReprovedTests = count
End Function
Public Function CountNotTested() As Integer
    Dim count As Integer
    count = 0
    For Each cv In testLists
        If cv.testStatus = "0" Then
            count = count + 1
        End If
    Next
    CountNotTested = count
End Function
