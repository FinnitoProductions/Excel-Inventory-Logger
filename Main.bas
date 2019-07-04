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
    Application.ScreenUpdating = False 'Prevent new window from displaying
    
    Dim fileName As Variant
    Dim stringFilePath As String
    Dim strippedFileName As String
    Dim book As Variant
    
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
    
    For Each Value In Range("A1:A3")
        Debug.Print (Value)
    Next Value
End Sub


