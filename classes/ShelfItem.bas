'
' Represents an item stored on a shelf. Can be used to represent an item which was ordered or an item which is currently
' in inventory.
'
' Finn Frankis
' July 6, 2019
'
Public sku As String
Public location As String
Public count As Integer
Public excelRow As Long
Public excelColumn As Long
Public availableCount As Long

Public Sub initiateProperties(desiredSku As String, Optional desiredCount As Integer = 0, Optional desiredLocation As String = "", _
                                            Optional desiredExcelRow As Long = 0, Optional desiredExcelColumn As Long = 0, Optional desiredAvailableCount As Long = -1)
   sku = desiredSku
   count = desiredCount
   location = desiredLocation
   excelRow = desiredExcelRow
   excelColumn = desiredExcelColumn
   availableCount = desiredAvailableCount
End Sub

Public Function toString() As String
   toString = sku & ": " & count & " items stored at " & location
   If availableCount <> -1 Then
        toString = toString & " (" & availableCount & " available)"
   End If
   
End Function