' This Map class simulates a map/dictionary. Although Collections support the addition of key-value pairs,
' the list of keys cannot be iterated over (only the values), so keeps track of a separate keyset for convenient key-value
' iteration.

Dim keyset As New Collection
Dim keyValPairs As New Collection

Public Function retrieve(key As String) As Variant
   retrieve = keyValPairs(key)
End Function

Private Function getKeysetIndex(key As String) As Integer
   For i = 1 To keyset.Count
      If keyset(i) = key Then
         getKeysetIndex = i
         Exit Function
      End If
   Next
    
   getKeysetIndex = -1
End Function

Public Function contains(key As String) As Boolean
   contains = getKeysetIndex(key) <> -1
End Function

Public Sub add(key As String, value As Variant)
   If contains(key) Then
      Call remove(key)
   End If

Call keyValPairs.add(Item:=value, key:=key)
Call keyset.add(key)
End Sub

Public Sub remove(key As String)
  Call keyValPairs.remove(key)
  keyset.remove (getKeysetIndex(key))
End Sub
