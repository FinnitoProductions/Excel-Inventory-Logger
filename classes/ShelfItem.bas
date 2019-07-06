' Represents an item stored on a shelf. Can be used to represent an item which was ordered or an item which is currently
' in inventory.
Public sku As String
Public location As String
Public count As Integer

Public Sub initiateProperties(desiredSku As String, Optional desiredCount As Integer = 0, Optional desiredLoc As String = "")
   sku = desiredSku
   count = desiredCount
   location = desiredLoc
End Sub