Option Explicit

Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
Const ORIGIN_WORKSHEET_NAME As String = "Inventory"

Const ORIGIN_SKU_COLUMN As Integer = 1
Const ORIGIN_LOCATION_LETTER_COLUMN As Integer = 5
Const ORIGIN_LOCATION_NUM_COLUMN As Integer = 6
Const ORIGIN_BLANK_ROWS_TO_EXIT As Integer = 3

Const ORDER_BOX_LABEL_COLUMN As Integer = 1
Const ORDER_SKU_COLUMN As Integer = 2
Const ORDER_COUNT_COLUMN As Integer = 4
Const ORDER_BLANK_ROWS_TO_EXIT As Integer = 3

Dim orderFile As String
Dim orderWorksheet As String

' If this macro isn't being run from the master inventory workbook, warns the user to change the active workbook
' before reexcuting the code and terminates the program.
Sub validateWorkbook()
    With Application.Workbooks
        Dim i As Long
        For i = 1 To .count
            If .Item(i).Name = ORIGIN_WORKBOOK_NAME Then
                Exit Sub
            End If
        Next
        MsgBox ("This macro must be executed with " & ORIGIN_WORKBOOK_NAME & " open. Please re-open.")
        'End
    End With
End Sub

' Iterates over the origin spreadsheet to produce a dictionary containing all SKUs currently in inventory corresponding to
' its location in storage.
Function generateSkuDictionary() As Map
    Dim skus As New Map
    Dim numBlankRows As Integer: numBlankRows = 0
    With Workbooks(ORIGIN_WORKBOOK_NAME).Worksheets(ORIGIN_WORKSHEET_NAME)
        Dim i As Long
        For i = 2 To Rows.count
            If WorksheetFunction.CountA(.Rows(i)) = 0 Then
                numBlankRows = numBlankRows + 1
            Else
                numBlankRows = 0
            End If

            If numBlankRows > ORIGIN_BLANK_ROWS_TO_EXIT Then
                Exit For
            End If

            Dim skuVal As String
            skuVal = CStr(.Cells(i, ORIGIN_SKU_COLUMN).value)
            If skuVal <> "" Then
                Dim location As String
                location = CStr(.Cells(i, ORIGIN_LOCATION_LETTER_COLUMN).value) _
                & CStr(.Cells(i, ORIGIN_LOCATION_NUM_COLUMN).value)

                Call skus.add(skuVal, location)
            End If
        Next
    End With

    Set generateSkuDictionary = skus
End Function

Function retrieveOrder(masterInventory As Map) As Map
    Dim returnVal As New Map

    With Workbooks(orderFile).Worksheets(orderWorksheet)
        Dim prevBoxLabel As String: prevBoxLabel = ""

        Dim numBlankRows As Integer: numBlankRows = 0
        Dim i As Long
        For i = 2 To .Rows.count
            If WorksheetFunction.CountA(.Rows(i)) = 0 Then
                numBlankRows = numBlankRows + 1
            Else
                numBlankRows = 0
            End If

            If numBlankRows > ORDER_BLANK_ROWS_TO_EXIT Then
                Exit For
            End If

            Dim boxLabel As String: boxLabel = CStr(.Cells(i, ORDER_BOX_LABEL_COLUMN).value)
            If boxLabel <> "" Then
                Call returnVal.add(boxLabel, New Collection)
                prevBoxLabel = boxLabel
            End If

            Dim correspondingSku As String: correspondingSku = CStr(.Cells(i, ORDER_SKU_COLUMN).value)
            Dim strCorrespondingCount As String: strCorrespondingCount = CStr(.Cells(i, ORDER_COUNT_COLUMN).value)
            If correspondingSku <> "" And strCorrespondingCount <> "" Then
                Dim intCorrespondingCount As Integer: intCorrespondingCount = CInt(strCorrespondingCount)
                
                Dim desiredShelfItem As New ShelfItem
                Dim shelfLocation As String: shelfLocation = masterInventory.retrieve(correspondingSku)
                Call desiredShelfItem.InitiateProperties(correspondingSku, _
                                                         desiredCount:=intCorrespondingCount, _ 
                                                         desiredLocation:= shelfLocation)
                
                Call returnVal.retrieve(prevBoxLabel).add(desiredShelfItem)
            End If
            
            If returnVal.contains(prevBoxLabel) Then
                If returnVal.retrieve(prevBoxLabel).count = 0 Then
                    returnVal.remove (prevBoxLabel)
                End If
            End If
        Next
    End With

    Set retrieveOrder = returnVal
End Function

Sub FindDesiredValues()
    'Call SaveBeforeExecute
    Call validateWorkbook
    Application.ScreenUpdating = False 'Prevent new window from displaying

    Dim baseInventory As Map
    Set baseInventory = generateSkuDictionary()
    Debug.Print (baseInventory.size())

    orderFile = openDesiredFile()
    orderWorksheet = Workbooks(orderFile).Sheets(1).Name

    Dim desiredGoods As Map
    Set desiredGoods = retrieveOrder(baseInventory)
    Debug.Print (desiredGoods.size())

    Dim boxLabel As Variant
    For Each boxLabel In desiredGoods.keyset
        Debug.Print(desiredGoods.retrieve(boxLabel).toString())
    Next boxLabel
End Sub
