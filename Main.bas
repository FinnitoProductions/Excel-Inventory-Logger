Attribute VB_Name = "Main"
Function SaveBeforeExecute()
    Select Case MsgBox("You can't undo this. Save workbook first?", vbYesNoCancel)
    Case Is = vbYes
        ThisWorkbook.Save
    Case Is = vbCancel
        End
    End Select
End Function
Function getWorksheetFromPath(path As String) As String
    Dim splitString() As String
    splitString = Split(path, "/")
    
    getWorksheetFromPath = splitString(UBound(splitString))
End Function
Sub FindDesiredValues()
    'Call SaveBeforeExecute
    Const ORIGIN_WORKBOOK_NAME As String = "harker inventory.xlsm"
    If ActiveWorkbook.Name <> ORIGIN_WORKBOOK_NAME Then
        MsgBox ("This macro must be executed from " & ORIGIN_WORKBOOK_NAME & ". Please re-open.")
        'Exit Sub
    End If
    
    Application.ScreenUpdating = False 'Prevent new window from displaying
    
    Dim fileName As Variant
    Dim stringFilePath As String
    Dim strippedFileName As String
    Dim skus As New Collection
    Dim skuKeyset As New Collection
    
    Const SKU_COLUMN As Integer = 1
    Const LOCATION_LETTER_COLUMN As Integer = 5
    Const LOCATION_NUM_COLUMN As Integer = 6
    
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
    
    For Each sku In skuKeyset
        Debug.Print (sku & " " & skus(sku))
    Next sku
    
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
    
    Workbooks(ORIGIN_WORKBOOK_NAME).Activate 'Reset after each execution
End Sub


