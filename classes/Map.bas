' 
' This Map class simulates a map/dictionary. Although Collections support the addition of key-value pairs,
' the list of keys cannot be iterated over (only the values), so this Map class keeps track of a separate keyset for convenient 
' key-value iteration.
'
' Finn Frankis
' July 4, 2019

Public keyset As New Collection
Public keyValPairs As New Collection

Public Function size() As Integer
    size = keyset.Count
End Function

Public Function retrieve(ByVal key As String) As Variant
    If IsObject(keyValPairs(key)) Then
        Set retrieve = keyValPairs(key)
    Else
        retrieve = keyValPairs(key)
    End If
End Function

Private Function getKeysetIndex(ByVal key As String) As Integer
    For i = 1 To keyset.Count
        If keyset(i) = key Then
            getKeysetIndex = i
            Exit Function
        End If
    Next
    
    getKeysetIndex = -1
End Function

Public Function contains(ByVal key As String) As Boolean
    contains = getKeysetIndex(key) <> -1
End Function

Public Sub add(ByVal key As String, ByVal value As Variant)
  If contains(key) Then
    Call remove(key)
  End If

Call keyValPairs.add(Item:=value, key:=key)
Call keyset.add(key)
End Sub

Public Sub remove(ByVal key As String)
  Call keyValPairs.remove(key)
  keyset.remove (getKeysetIndex(key))
End Sub
