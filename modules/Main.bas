Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
Const ORIGIN_WORKSHEET_NAME As String = "Inventory"

Const ORIGIN_SKU_COLUMN As Integer = 1
Const ORIGIN_LOCATION_LETTER_COLUMN As Integer = 5
Const ORIGIN_LOCATION_NUM_COLUMN As Integer = 6

Const ORDER_BOX_LABEL_COLUMN As Integer = 1
Const ORDER_SKU_COLUMN As Integer = 2
Const ORDER_COUNT_COLUMN As Integer = 3

Dim orderFile As String
Dim orderWorksheet As String


' If this macro isn't being run from the master inventory workbook, warns the user to change the active workbook
' before reexcuting the code and terminates the program.
Function validateWorkbook()
    If ActiveWorkbook.Name <> ORIGIN_WORKBOOK_NAME Then
        MsgBox ("This macro must be executed from " & ORIGIN_WORKBOOK_NAME & ". Please re-open.")
        'End
    End If
End Function

' Offers to save the workbook before the macro is executed to allow the macro's actions to be easily undone (Excel does not
' support native undoing of Macro actions).
Function SaveBeforeExecute()
    Select Case MsgBox("You can't undo this. Save workbook first?", vbYesNoCancel)
    Case Is = vbYes
        ThisWorkbook.Save
    Case Is = vbCancel
        End ' Terminates program execution
    End Select
End Function

' Extracts the filename given an absolute directory path.
Function getWorksheetFromPath(path As String) As String
    Dim splitString() As String
    splitString = Split(path, "/")
    
    getWorksheetFromPath = splitString(UBound(splitString))
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
' Determines whether a collection (values) contains a given key (desiredVal).
Function collectionContainsKey(desiredKey As String, values As Collection) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = values(desiredKey)
    collectionContainsKey = (Err.Number = 0)
    Err.Clear
End Function
Function updateCollectionKey(desiredKey As String, newValue As Variant, values As Collection) As Collection
    Call values.remove(desiredKey)
    Call values.add(Item:=newValue, key:=desiredKey)
    Set updateCollectionKey = values
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

' Iterates over the origin spreadsheet to produce a dictionary containing all SKUs currently in inventory corresponding to
' its location in storage.
Function generateSkuDictionary() As Map
    Dim skus As New Map
    
    For i = 2 To Rows.Count
        Dim skuVal As String
        skuVal = CStr(Cells(i, ORIGIN_SKU_COLUMN).value)
        If skuVal <> "" Then
            Dim location As String
            With Workbooks(ORIGIN_WORKBOOK_NAME).Worksheets(ORIGIN_WORKSHEET_NAME)
                location = CStr(.Cells(i, ORIGIN_LOCATION_LETTER_COLUMN).value) & CStr(.Cells(i, ORIGIN_LOCATION_NUM_COLUMN).value)
            End With

            Call skus.add(skuVal, location)
        End If
    Next

    Set generateSkuDictionary = skus
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
    Workbooks(strippedFileName).Activate

    openDesiredFile = strippedFileName
End Function

Function retrieveOrder() As Map
    Dim returnVal As New Map

    With Workbooks(orderFile).Worksheets(orderWorksheet)
        Dim prevBoxLabel As String: prevBoxLabel = ""
        For i = 2 To .Rows.Count
            Dim boxLabel As String: boxLabel = CStr(.Cells(i, ORDER_BOX_LABEL_COLUMN).value)
            If boxLabel <> "" Then
                Dim storedMap As Map: Set storedMap = New Map
                Call returnVal.add(boxLabel, storedMap)
                prevBoxLabel = boxLabel
            End If

            Dim correspondingSku As String: correspondingSku = CStr(.Cells(i, ORDER_SKU_COLUMN).value)
            Dim correspondingCount As Integer: correspondingCount = CInt(.Cells(i, ORDER_COUNT_COLUMN).value)
            If correspondingSku <> "" Then
                Debug.Print (returnVal.contains(prevBoxLabel))
                'Dim retrievedMap As Variant: Set retrievedMap = returnVal.retrieve(prevBoxLabel)
                Call returnVal.retrieve(prevBoxLabel)
                'Debug.Print ("type " & TypeName(retrievedMap))
                'Call retrievedMap.add(correspondingSku, correspondingCount)
                Debug.Print (correspondingSku & ", " & correspondingCount)
            End If

        Next
    End With

    Set retrieveOrder = returnVal
End Function

Sub FindDesiredValues()
    'Call SaveBeforeExecute
    Call validateWorkbook
    Application.ScreenUpdating = False 'Prevent new window from displaying

    Dim desiredMap As Map
    Set desiredMap = generateSkuDictionary()

    For Each key In desiredMap.keyset
        Debug.Print (desiredMap.retrieve(key))
    Next

    orderFile = openDesiredFile()
    orderWorksheet = Workbooks(orderFile).Sheets(1).Name

    Call retrieveOrder

    Workbooks(ORIGIN_WORKBOOK_NAME).Activate 'Reset after each execution
End Sub




