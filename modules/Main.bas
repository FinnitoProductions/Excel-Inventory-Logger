'
' Contains all the functions and subroutines specific to the execution of the inventory logging project.
' First validates the workbook, then generates a dictioanry using the master ("origin") SKU list, then
' retrieves a given order (as selected by the user) and matches each item to its location. Concludes by writing
' the data to a printable file and deducting the desired amount from the inventory.
'
' Finn Frankis
' July 2, 2019
'
Option Explicit

Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
Const ORIGIN_WORKSHEET_NAME As String = "Inventory"

Const ORIGIN_SKU_COLUMN As Integer = 1
Const ORIGIN_LOCATION_LETTER_COLUMN As Integer = 4
Const ORIGIN_LOCATION_NUM_COLUMN As Integer = 5
Const ORIGIN_BLANK_ROWS_TO_EXIT As Integer = 3
Const ORIGIN_COUNT_COLUMN As Integer = 3

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
                
                Dim count As Long
                count = CInt(.Cells(i, ORIGIN_COUNT_COLUMN).value)
                
                Dim desiredShelfItem As shelfItem: Set desiredShelfItem = New shelfItem
                Call desiredShelfItem.initiateProperties(skuVal, _
                    desiredLocation:=location, _
                    desiredAvailableCount:=count, _
                    desiredExcelRow:=i)
                    
                Call skus.add(skuVal, desiredShelfItem)

            End If
        Next
    End With

    Set generateSkuDictionary = skus
End Function

Function retrieveOrder(orderWorksheet As Worksheet, masterInventory As Map) As Map
    Dim returnVal As New Map

    With orderWorksheet
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
                If returnVal.contains(prevBoxLabel) Then
                    If returnVal.retrieve(prevBoxLabel).count = 0 Then
                        returnVal.remove (prevBoxLabel)
                    End If
                End If
                Call returnVal.add(boxLabel, New Collection)
                prevBoxLabel = boxLabel
            End If

            Dim correspondingSku As String: correspondingSku = CStr(.Cells(i, ORDER_SKU_COLUMN).value)
            Dim strCorrespondingCount As String: strCorrespondingCount = CStr(.Cells(i, ORDER_COUNT_COLUMN).value)
            If correspondingSku <> "" And strCorrespondingCount <> "" Then
                Dim intCorrespondingCount As Integer: intCorrespondingCount = CInt(strCorrespondingCount)
                
                Dim desiredShelfItem As shelfItem: Set desiredShelfItem = New shelfItem
                Dim correspondingItem As shelfItem: Set correspondingItem = masterInventory.retrieve(correspondingSku)

                Call desiredShelfItem.initiateProperties(correspondingSku, _
                                                         desiredCount:=intCorrespondingCount, _
                                                         desiredLocation:=correspondingItem.location, _
                                                         desiredAvailableCount:=correspondingItem.availableCount)
                                                         
                Call returnVal.retrieve(prevBoxLabel).add(desiredShelfItem)
            End If
        Next

        If returnVal.contains(prevBoxLabel) Then
            If returnVal.retrieve(prevBoxLabel).count = 0 Then
                returnVal.remove (prevBoxLabel)
            End If
        End If
    End With

    Set retrieveOrder = returnVal
End Function

Sub deductInventory(masterInventoryDict As Map, orderDict As Map)
    Dim boxLabel As Variant
    For Each boxLabel In orderDict.keyset
        Dim shelfItems As Collection: Set shelfItems = orderDict.retrieve(boxLabel)
        Dim shelfItem As Variant
        For Each shelfItem In shelfItems
            If masterInventoryDict.contains(shelfItem.sku) Then
                Dim correspondingItem As shelfItem: Set correspondingItem = masterInventoryDict.retrieve(shelfItem.sku)
                With Workbooks(ORIGIN_WORKBOOK_NAME).Sheets(ORIGIN_WORKSHEET_NAME)
                    .Cells(correspondingItem.excelRow, ORIGIN_COUNT_COLUMN).value = .Cells(correspondingItem.excelRow, ORIGIN_COUNT_COLUMN).value - shelfItem.count
                End With
            Else
                Debug.Print ("Sku " & shelfItem.sku & " not contained in master inventory.")
            End If
        Next shelfItem
    Next boxLabel
End Sub

Sub writeDataToFile(data As Map)
    Const FILE_NAME As String = "orderData.txt"
    Dim filePath As String: filePath = Application.DefaultFilePath & FILE_NAME
    
    Open filePath For Output As #1

    Dim boxLabel As Variant
    For Each boxLabel In data.keyset
        Print #1, boxLabel
        Dim shelfItems As Collection: Set shelfItems = data.retrieve(boxLabel)
        Dim shelfItem As Variant
        For Each shelfItem In shelfItems
            Print #1, shelfItem.toString()
        Next shelfItem
        Print #1, vbCrLf ' Print new line
    Next boxLabel

    Close #1
End Sub

Sub FindDesiredValues()
    Call SaveBeforeExecute
    Call validateWorkbook
    Application.ScreenUpdating = False 'Prevent new window from displaying

    Dim baseInventory As Map
    Set baseInventory = generateSkuDictionary()

    orderFile = openDesiredFile()
    orderWorksheet = Workbooks(orderFile).Sheets(1).Name

    Dim desiredGoods As Map
    Set desiredGoods = retrieveOrder(Workbooks(orderFile).Worksheets(orderWorksheet), baseInventory)

    Call writeDataToFile(desiredGoods)
    
    'Call deductInventory(baseInventory, desiredGoods)
End Sub

Sub GenerateInventorySpreadsheet()
    Application.ScreenUpdating = False 'Prevent new window from displaying
    ActiveWindow.DisplayGridlines = False
    
    Dim currentDate As String: currentDate = CStr(Date)
    currentDate = Replace(currentDate, "/", "-")
    Dim NEW_FILE_NAME As String: NEW_FILE_NAME = "inventory-" & currentDate & ".xlsx"
    Dim NEW_FILE_LOC As String: NEW_FILE_LOC = Application.DefaultFilePath & NEW_FILE_NAME

    Dim newWorkbook As Workbook: Set newWorkbook = Workbooks.add
    With newWorkbook
        Call .SaveAs(fileName:=NEW_FILE_LOC)
        Call Workbooks(ORIGIN_WORKBOOK_NAME).Sheets(ORIGIN_WORKSHEET_NAME).Copy(Before:=newWorkbook.Sheets(1))
        
        .Sheets(1).Columns(ORIGIN_LOCATION_NUM_COLUMN).EntireColumn.Delete
        .Sheets(1).Columns(ORIGIN_LOCATION_LETTER_COLUMN).EntireColumn.Delete
    End With
    Call newWorkbook.Close(SaveChanges:=True)
End Sub