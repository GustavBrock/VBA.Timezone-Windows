Attribute VB_Name = "ByteUtil"
Option Compare Text
Option Explicit

' ByteUtil v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Functions for converting integers to hex bytes (octets).
'
' License: MIT (http://opensource.org/licenses/mit-license.php)


Private Type Byte1
    Value(0)    As Byte
End Type

Private Type Byte2
    Value(1)    As Byte
End Type

Private Type Byte4
    Value(3)    As Byte
End Type

Private Type Integer1
    Value(0)    As Integer
End Type

Private Type Long1
    Value(0)    As Long
End Type

' 2018-06-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CByteBytes( _
    ByRef OneBytes() As Byte) _
    As Byte

    Dim ByteItems       As Byte1
    Dim ByteItem        As Byte1
    Dim Index           As Integer
    
    If Index >= LBound(OneBytes) And Index <= UBound(OneBytes) Then
        ByteItems.Value(Index) = OneBytes(Index)
    End If
    
    LSet ByteItem = ByteItems
    
    CByteBytes = ByteItem.Value(LBound(ByteItems.Value))

End Function

' 2018-06-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CIntBytes( _
    ByRef TwoBytes() As Byte) _
    As Integer

    Dim IntegerItems    As Byte2
    Dim IntegerItem     As Integer1
    Dim Index           As Integer
    
    For Index = LBound(IntegerItems.Value) To UBound(IntegerItems.Value)
        If Index >= LBound(TwoBytes) And Index <= UBound(TwoBytes) Then
            IntegerItems.Value(Index) = TwoBytes(Index)
        End If
    Next
    
    LSet IntegerItem = IntegerItems
    
    CIntBytes = IntegerItem.Value(LBound(IntegerItems.Value))

End Function

' 2018-06-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CLngBytes( _
    ByRef FourBytes() As Byte) _
    As Long

    Dim LongItems       As Byte4
    Dim LongItem        As Long1
    Dim Index           As Integer
    
    For Index = LBound(LongItems.Value) To UBound(LongItems.Value)
        If Index >= LBound(FourBytes) And Index <= UBound(FourBytes) Then
            LongItems.Value(Index) = FourBytes(Index)
        End If
    Next
    
    LSet LongItem = LongItems
    
    CLngBytes = LongItem.Value(LBound(LongItems.Value))

End Function

