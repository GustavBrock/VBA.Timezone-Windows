Attribute VB_Name = "DataUtil"
Option Compare Database
Option Explicit

' DataUtil v1.1.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Common functions for database and tables.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)


' Searches a connection string for a key.
' If found, the value of the key is returned.
' If not found, an empty string is returned.
'
' 2017-11-16. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function GetConnectionKeyValue( _
    ByVal Connection As String, _
    ByVal Key As String) _
    As String
    
    Const ElementSeparator  As String = ";"
    Const KeyValueSeparator As String = "="
    
    Dim KeyValuePairs   As Variant
    Dim KeyValuePair    As Variant
    
    Dim Element         As Integer
    Dim Value           As String
    
    If Connection <> "" Then
        KeyValuePairs = Split(Connection, ElementSeparator)
        For Element = LBound(KeyValuePairs) To UBound(KeyValuePairs)
            If KeyValuePairs(Element) <> "" Then
                KeyValuePair = Split(KeyValuePairs(Element), KeyValueSeparator)
                If StrComp(Key, KeyValuePair(0)) = 0 Then
                    If LBound(KeyValuePair) < UBound(KeyValuePair) Then
                        Value = KeyValuePair(1)
                    End If
                    Exit For
                End If
            End If
        Next
    End If
    
    GetConnectionKeyValue = Value
    
End Function

' Checks if a named table exists.
' Returns True if so, False if not.
'
' 2018-10-18. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function IsTableDefName( _
    ByVal TableName As String) _
    As Boolean

    Dim Database    As DAO.Database
    Dim Table       As DAO.TableDef
    
    Dim Result      As Boolean
    
    Set Database = CurrentDb
    
    For Each Table In Database.TableDefs
        If Table.Name = TableName Then
            ' Table exists. Exit.
            Result = True
            Exit For
        End If
    Next

    IsTableDefName = Result

End Function
    
' Checks if a relation named RelationName exists.
' Returns True if it is found, False if not.
'
' 2017-11-14. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function IsRelationName( _
    ByVal RelationName As String) _
    As Boolean
    
    Dim Relation    As DAO.Relation
    
    Dim Result      As Boolean
    
    For Each Relation In CurrentDb.Relations
        If Relation.Name = RelationName Then
            Exit For
        End If
    Next
    Result = Not (Relation Is Nothing)
    
    IsRelationName = Result
    
End Function

