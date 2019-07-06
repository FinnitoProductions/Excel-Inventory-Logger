Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
Const ORIGIN_WORKSHEET_NAME As String = "Inventory"

Const ORIGIN_SKU_COLUMN As Integer = 1
Const ORIGIN_LOCATION_LETTER_COLUMN As Integer = 5
Const ORIGIN_LOCATION_NUM_COLUMN As Integer = 6

Const ORDER_BOX_LABEL_COLUMN As Integer = 1
Const ORDER_SKU_COLUMN As Integer = 2
Const ORDER_COUNT_COLUMN As Integer = 4

Dim orderFile As String
Dim orderWorksheet As String


' If this macro isn't being run from the master inventory workbook, warns the user to change the active workbook
' before reexcuting the code and terminates the program.
Sub validateWorkbook()
    With Application.Workbooks
        For i = 1 To .Count
            If .Item(i).Name = ORIGIN_WORKBOOK_NAME Then
                Debug.Print ("validated")
                Exit Sub
            End If
        Next
        MsgBox ("This macro must be executed with " & ORIGIN_WORKBOOK_NAME & " open. Please re-open.")
        'End
    End With
End Sub

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

    openDesiredFile = strippedFileName
End Function

Function retrieveOrder() As Map
    Dim returnVal As New Map

    With Workbooks(orderFile).Worksheets(orderWorksheet)
        Dim prevBoxLabel As String: prevBoxLabel = ""
        For i = 2 To .Rows.Count
            Dim boxLabel As String: boxLabel = CStr(.Cells(i, ORDER_BOX_LABEL_COLUMN).value)
            If boxLabel <> "" Then
                Call returnVal.add(boxLabel, New Map)
                prevBoxLabel = boxLabel
            End If

            Dim correspondingSku As String: correspondingSku = CStr(.Cells(i, ORDER_SKU_COLUMN).value)
            Dim strCorrespondingCount As String: strCorrespondingCount = CStr(.Cells(i, ORDER_COUNT_COLUMN).value)
            If correspondingSku <> "" And strCorrespondingCount <> "" Then
                Dim intCorrespondingCount As Integer: intCorrespondingCount = CInt(strCorrespondingCount)
                Call returnVal.retrieve(prevBoxLabel).add(correspondingSku, intCorrespondingCount)
                Debug.Print (correspondingSku & ", " & intCorrespondingCount)
            End If
            
            If returnVal.contains(prevBoxLabel) Then
                If returnVal.retrieve(prevBoxLabel).size() = 0 Then
                    returnVal.remove (prevBoxLabel)
                End If
            End If
        Next
    End With

    Set retrieveOrder = returnVal
End Function

Sub FindDesiredValues()
    ' 'Call SaveBeforeExecute
    ' Call validateWorkbook
    ' Application.ScreenUpdating = False 'Prevent new window from displaying

    ' Dim baseInventory As Map
    ' Set baseInventory = generateSkuDictionary()

    ' For Each key In baseInventory.keyset
    '     Debug.Print (baseInventory.retrieve(key))
    ' Next

    orderFile = openDesiredFile()
    orderWorksheet = Workbooks(orderFile).Sheets(1).Name

    Dim desiredGoods As Map
    Set desiredGoods = retrieveOrder()
    Debug.Print (desiredGoods.size())
End Sub






