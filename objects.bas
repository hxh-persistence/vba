Attribute VB_Name = "objects"
Option Explicit

'create dictionary
Sub CreateDic()
    Dim dic As Object
    Dim arrKey
    Dim arrItem
    Set dic = CreateObject("Scripting.Dictionary")
    
    With dic
        'add key and item
        .Add 1, "a"
        .Add 2, "b"
        'add key and item to array
        arrKey = .keys
        arrItem = .items
    End With
    
    dic.RemoveAll
    Set dic = Nothing
End Sub

Function CreateInput()
    Dim length As Integer
    Dim width As Integer

    length = InputBox("Enter length", "Enter number")
    width = InputBox("Enter length", "Enter number")
    CreateInput = length * width
End Function

