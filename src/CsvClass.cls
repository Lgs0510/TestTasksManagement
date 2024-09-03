VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CsvClass module
Option Explicit

Private csvContent As New Collection
Private numberOfLines As Integer
Private Type csvDataTypes
    id() As New csvIDClass
    'Title() As New csvTitleClass
    lastUpdated() As New csvLastUpdatedClass
    linkedWorkItens() As New csvLinkedWorkItensClass
End Type

Private csvFile As csvDataTypes

' Event - triggered when class created
Private Sub Class_Initialize()
End Sub

Public Property Let newCsv(newCsv)
    Me = newCsv
End Property

Public Property Get getPropertyNames() As String()
    Dim csvType As Variant
    Dim namesArray() As String
    Dim arrayPos As Integer
    
    arrayPos = 0
    ReDim namesArray(csvContent.count - 1)
    For Each csvType In csvContent
        namesArray(arrayPos) = csvType(1).getName
        arrayPos = arrayPos + 1
    Next
    getPropertyNames = namesArray
End Property
Public Sub initCsv(csvSize, listIDs)
    Dim i, z As Integer
    ReDim csvFile.id(csvSize)
    'ReDim csvFile.Title(csvSize)
    ReDim csvFile.lastUpdated(csvSize)
    ReDim csvFile.linkedWorkItens(csvSize)
    
    For i = 0 To csvSize
        csvFile.id(i).initClass (listIDs.Find(csvFile.id(i).getName))
        'csvFile.Title(i).initClass (listIDs.Find(csvFile.Title(i).getName))
        csvFile.lastUpdated(i).initClass (listIDs.Find(csvFile.lastUpdated(i).getName))
        csvFile.linkedWorkItens(i).initClass (listIDs.Find(csvFile.linkedWorkItens(i).getName))
    Next
    csvContent.Add item:=csvFile.id, Key:="ID"
    'csvContent.Add item:=csvFile.Title, Key:="Title"
    csvContent.Add item:=csvFile.lastUpdated, Key:="Updated"
    csvContent.Add item:=csvFile.linkedWorkItens, Key:="Linked Work Items"
End Sub

Public Sub addLine(strLine As String)
    Dim lineArray() As String
    Dim i As Integer
    Dim anything As Integer
    lineArray = Split(strLine, ";")
    csvContent("ID")(numberOfLines).setValue (lineArray(csvContent("ID")(numberOfLines).getIndex))
    'csvContent("Title")(numberOfLines).setValue (lineArray(csvContent("Title")(numberOfLines).getIndex))
    csvContent("Linked Work Items")(numberOfLines).setValue (lineArray(csvContent("Linked Work Items")(numberOfLines).getIndex))
    If csvContent("Updated")(numberOfLines).getIndex >= 0 Then
        csvContent("Updated")(numberOfLines).setValue (lineArray(csvContent("Updated")(numberOfLines).getIndex))
    End If
    numberOfLines = numberOfLines + 1
    
End Sub

Public Function getReqList() As String()
    Dim reqList As New list
    Dim id As Variant
    
    For Each id In csvContent(1)
        reqList.Add (id.getValue)
    Next
    getReqList = reqList.getList
End Function
Public Function getReqListNO() As String()
    Dim reqList As New list
    Dim id As Variant
    
    For Each id In csvContent(1)
        reqList.Add (Replace(id.getValue, "CV-", ""))
    Next
    getReqListNO = reqList.getList
End Function

Public Function getReqLikedWkItems(ByVal wantedID As String) As String
    Dim id As csvIDClass
    Dim i As Integer
    
    For i = 0 To SizeOfArray(csvContent("ID")) - 1
        Set id = csvContent("ID")(i)
        If StrComp(id.getValue, wantedID) = 0 Then
            getReqLikedWkItems = csvContent("Linked Work Items")(i).getValue
            Exit Function
        End If
    Next
    getReqLikedWkItems = 0
End Function
