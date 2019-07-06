Option Explicit

' Offers to save the workbook before the macro is executed to allow the macro's actions to be easily undone (Excel does not
' support native undoing of Macro actions).
Sub SaveBeforeExecute()
    Select Case MsgBox("You can't undo this. Save workbook first?", vbYesNoCancel)
    Case Is = vbYes
        ThisWorkbook.Save
    Case Is = vbCancel
        End ' Terminates program execution
    End Select
End Sub

' Extracts the filename given an absolute directory path.
Function getWorksheetFromPath(path As String) As String
    Dim splitString() As String
    splitString = Split(path, Application.PathSeparator)
    
    getWorksheetFromPath = splitString(UBound(splitString))
End Function

' Prompts the user to select an Excel workbook and opens this workbook.
Function openDesiredFile() As String
    Dim fileName As Variant
    Dim stringFilePath As String
    Dim strippedFileName As String
    
    fileName = Application.GetOpenFilename()

    If fileName <> False Then
        Workbooks.Open (fileName)
        stringFilePath = CStr(fileName)
    Else
        While fileName = False
            fileName = Application.GetOpenFilename()
        Wend
    End If

    strippedFileName = getWorksheetFromPath(stringFilePath)

    openDesiredFile = strippedFileName
End Function

' Converts a given array into a collection.
Function arrToCollection(initArr As Variant) As Collection
    Dim desiredCollection As New Collection
    
    For Each value In initArr
        desiredCollection.add value
    Next value
    
    Set arrToCollection = desiredCollection
End Function

' Determines whether a collection (values) contains a given value (desiredVal).
Function collectionContains(desiredVal As String, values As Collection) As Boolean
    For Each value In values
        If value = desiredVal Then
            collectionContains = True
            Exit Function
        End If
    Next value

    collectionContains = False
End Function

' Determines whether a given string represents a valid size (either XS, S, M, L, XL, or XXL).
Function isSize(potentialSize As String) As Boolean
    Dim validSizes As Collection
    Dim sizeArray As Variant
    
    Set validSizes = arrToCollection(Array("XS", "S", "M", "L", "XL", "XXL"))
    
    isSize = collectionContains(potentialSize, validSizes)
End Function

' Determines whether a given string represents a SKU which can be shipped.
Function isSku(potentialSku As String) As Boolean
    Dim splitString() As String
    Const MAX_SKU_TOKENS As Integer = 2
    Dim arrLength As Integer

    splitString = Split(potentialSku, " ")
    arrLength = UBound(splitString) + 1
    
    If arrLength > MAX_SKU_TOKENS Then
        isSku = False
        Exit Function
    ElseIf arrLength = MAX_SKU_TOKENS Then
        isSku = isSize(splitString(UBound(splitString)))
    Else
        isSku = True
    End If
End Function
