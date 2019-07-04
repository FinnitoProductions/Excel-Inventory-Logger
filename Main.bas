Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
Const ORIGIN_WORKSHEET_NAME As String = "Inventory"

Const SKU_COLUMN As Integer = 1
Const LOCATION_LETTER_COLUMN As Integer = 5
Const LOCATION_NUM_COLUMN As Integer = 6
Dim orderFile As String

Public Type Map
    keyset As Collection
    keyValPairs As Collection
End Type

Function validateWorkbook()
    If ActiveWorkbook.Name <> ORIGIN_WORKBOOK_NAME Then
        MsgBox ("This macro must be executed from " & ORIGIN_WORKBOOK_NAME & ". Please re-open.")
        'End
    End If
End Function

Function SaveBeforeExecute()
    Select Case MsgBox("You can't undo this. Save workbook first?", vbYesNoCancel)
    Case Is = vbYes
        ThisWorkbook.Save
    Case Is = vbCancel
        End ' Terminates program execution
    End Select
End Function

Function getWorksheetFromPath(path As String) As String
    Dim splitString() As String
    splitString = Split(path, "/")
    
    getWorksheetFromPath = splitString(UBound(splitString))
End Function

Function arrToCollection(initArr As Variant) As Collection
    Dim desiredCollection As New Collection
    
    For Each Value In initArr
        desiredCollection.Add Value
    Next Value
    
    Set arrToCollection = desiredCollection
End Function

Function collectionContains(desiredVal As String, values As Collection) As Boolean
    For Each Value In values
        If Value = desiredVal Then
            collectionContains = True
            Exit Function
        End If
    Next Value

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
        Debug.Print (TypeName(splitString(UBound(splitString))))
        isSku = isSize(splitString(UBound(splitString)))
    Else
        isSku = True
    End If
End Function

Function generateSkuDictionary() As Map
    Dim desiredMap As Map

    Dim skus As New Collection
    Dim skuKeyset As New Collection
    
    For i = 2 To Rows.Count
        Dim skuVal As String
        skuVal = CStr(Cells(i, SKU_COLUMN).Value)
        If skuVal <> "" Then
            Dim location As String
            location = CStr(Cells(i, LOCATION_LETTER_COLUMN).Value) & CStr(Cells(i, LOCATION_NUM_COLUMN).Value)

            Call skus.Add(location, skuVal)
            Call skuKeyset.Add(skuVal)
        End If
    Next

    Set desiredMap.keyset = skuKeyset
    Set desiredMap.keyValPairs = skus
    generateSkuDictionary = desiredMap
End Function

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
    Workbooks(strippedFileName).Activate

    openDesiredFile = strippedFileName
End Function

Function retrieveOrder() As Map

End Function

Sub FindDesiredValues()
    'Call SaveBeforeExecute
    Call validateWorkbook
    Application.ScreenUpdating = False 'Prevent new window from displaying
    
    Dim desiredMap As Map
    desiredMap = generateSkuDictionary()
    
    For Each Key In desiredMap.keyset
        Debug.Print (desiredMap.keyValPairs(Key))
    Next

    orderFile = openDesiredFile()
    
    Workbooks(ORIGIN_WORKBOOK_NAME).Activate 'Reset after each execution
End Sub

