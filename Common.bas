Attribute VB_Name = "Common"
Option Explicit
'
' Version 1.3.1
' 2020-03-01. Gustav Brock, Cactus Data ApS, CPH.

' General constants.

' Named Ranges.
' Separator between the worksheet name and the name of the Named Range.
Public Const NamedRangeSeparator    As String = "!"

' Resizes a Named Range and cleans excessive rows and columns.
' Applies the interior color of the named range to an expanded range.
'
' Optionally, set EraseExcessRows to False to not clean excessive rows and columns.
'
' If both rows and colums are reduced, the color of the bottom-right corner
' is applied to the excess area at bottom-right.
' Optionally, set ColorPreference to select the adjacent color of either
' rows or columns to fill an excess area.
'
' 2017-09-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ResizeNamedRange( _
    ByVal Index As Variant, _
    ByVal RowSize As Long, _
    ByVal ColumnSize As Long, _
    Optional ByVal EraseExcessRows As Boolean = True, _
    Optional ByVal ColorPreference As XlRowCol)

    Dim Name                As Excel.Name
    Dim NamedRange          As Excel.Range
    
    Dim ExcessRowsCount     As Long
    Dim ExcessColumnsCount  As Long
    Dim ClearColorRows      As Long
    Dim ClearColorColumns   As Long
    Dim ClearColorCorner    As Long
    Dim RangeColor          As Long
    
    Set Name = Names(Index)
    Set NamedRange = Name.RefersToRange
        
    If RowSize > 0 And ColumnSize > 0 Then
        ' Pick color from top-left cell of the Named Range.
        RangeColor = NamedRange.Resize(1, 1).Interior.Color
        
        If EraseExcessRows = True Then
            ExcessRowsCount = NamedRange.Rows.Count - RowSize
            ExcessColumnsCount = NamedRange.Columns.Count - ColumnSize
            If ExcessRowsCount > 0 Or ExcessColumnsCount > 0 Then
                ' Pick color from the adjacent row.
                ClearColorRows = NamedRange.Rows(NamedRange.Rows.Count + 1).Interior.Color
                ' Pick color from the adjacent column.
                ClearColorColumns = NamedRange.Columns(NamedRange.Columns.Count + 1).Interior.Color
                ' Determine the color setting of a bottom-right excess area.
                Select Case ColorPreference
                    Case xlRows
                        ' Use the color from the adjecent row.
                        ClearColorCorner = ClearColorRows
                    Case xlColumns
                        ' Use the color from the adjecent column.
                        ClearColorCorner = ClearColorColumns
                    Case Else
                        ' Pick color from the adjacent corner.
                        ClearColorCorner = _
                            NamedRange.Cells(NamedRange.Rows.Count + 1, NamedRange.Columns.Count + 1).Interior.Color
                End Select
                
                ' Clear excess rows.
                If ExcessRowsCount > 0 Then
                    ' Clear color.
                    NamedRange.Offset(RowSize, 0).Resize(ExcessRowsCount).Interior.Color = ClearColorRows
                    ' Clear contents.
                    NamedRange.Offset(RowSize, 0).Resize(ExcessRowsCount).ClearContents
                End If
                ' Clear excess columns.
                If ExcessColumnsCount > 0 Then
                    ' Clear color.
                    NamedRange.Offset(0, ColumnSize).Resize(, ExcessColumnsCount).Interior.Color = ClearColorColumns
                    ' Clear contents.
                    NamedRange.Offset(0, ColumnSize).Resize(, ExcessColumnsCount).ClearContents
                End If
                ' Clear excess corner.
                If ExcessRowsCount > 0 And ExcessColumnsCount > 0 Then
                    ' Clear color.
                    NamedRange.Offset(RowSize, ColumnSize).Resize(ExcessRowsCount, ExcessColumnsCount).Interior.Color = ClearColorCorner
                End If
                
            End If
        End If
        
        ' Resize the Named Range.
        Name.RefersTo = NamedRange.Resize(RowSize, ColumnSize)
        ' Set (reapply) color of the full Named Range.
        Name.RefersToRange.Interior.Color = RangeColor
    End If
    
    Set NamedRange = Nothing
    Set Name = Nothing

End Sub

' Copies in a Named Range the formula(s) of the specified column(s)
' from the top row to the remaining rows.
'
' 2017-02-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CopyColumnFormulas( _
    ByVal Index As Variant, _
    Optional ByVal StartColumnIndex As Long = 1, _
    Optional ByVal ColumnCount As Long = 1)

    Const FirstRowIndex As Long = 1
    
    Dim Name            As Excel.Name
    Dim NamedRange      As Excel.Range
    
    Dim Formula         As String
    Dim ColumnIndex     As Long
    Dim RowIndex        As Long
    
    Set Name = Names(Index)
    Set NamedRange = Name.RefersToRange
        
    For ColumnIndex = StartColumnIndex To StartColumnIndex + ColumnCount - 1
        ' Copy formula from first row of the column.
        Formula = NamedRange.Cells(FirstRowIndex, ColumnIndex).FormulaR1C1
        ' Paste formula to the other rows of the column.
        For RowIndex = FirstRowIndex + 1 To NamedRange.Rows.Count
            NamedRange.Cells(RowIndex, ColumnIndex).Formula = Formula
        Next
    Next
     
    Set NamedRange = Nothing
    Set Name = Nothing
    
End Sub

' Renames the project containing this workbook.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RenameProject( _
    ByVal Name As String)
    
    Const DefaultProjectName    As String = "VBAProject"
    
    CleanModuleCodeName Name
    If Name = "" Then
        Name = DefaultProjectName
    End If
    ThisWorkbook.VBProject.Name = Name

End Sub

' Renames a code module of this workbook.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RenameModule( _
    ByVal Index As Variant, _
    ByVal Name As String)
    
    CleanModuleCodeName Name
    If Name <> "" Then
        ThisWorkbook.VBProject.VBComponents(Index).Name = Name
    End If

End Sub

' Renames the code module of Workbook.
'
' 2017-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RenameWorkbookModule( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Name As String)

    If Not Workbook Is Nothing Then
        CleanModuleCodeName Name
        If Name <> "" Then
            Workbook.[_CodeName] = Name
        End If
    End If
    
End Sub

' Renames the code module of a worksheet of Workbook.
'
' 2017-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RenameWorksheetModule( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Index As Variant, _
    ByVal Name As String)
    
    Dim CodeIndex   As Long
    
    If Not Workbook Is Nothing Then
        CodeIndex = WorksheetModuleIndex(Workbook, Index)
        RenameModule CodeIndex, Name
    End If
    
End Sub

' Renames a worksheet in this workbook.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RenameWorksheet( _
    ByVal Index As Variant, _
    ByVal Name As String)

    CleanWorksheetName Name
    ThisWorkbook.Worksheets(Index).Name = Name
    
End Sub

' Returns the index of the code module of a worksheet.
'
' 2017-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorksheetModuleIndex( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Index As Variant) _
    As Long
    
    Dim CodeIndex   As Long
    
    If Not Workbook Is Nothing Then
        CodeIndex = ModuleIndex(Workbook, WorksheetModuleName(Workbook, Index))
    End If
        
    WorksheetModuleIndex = CodeIndex
    
End Function

' Returns the name of the code module of a worksheet of Workbook.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorksheetModuleName( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Index As Variant) _
    As String
    
    Dim CodeName    As String
    
    If Not Workbook Is Nothing Then
        CodeName = ThisWorkbook.Worksheets(Index).CodeName
    End If
    
    WorksheetModuleName = CodeName
    
End Function

' Replaces characters in Name that are not allowed in a worksheet name.
' Truncates length of Name to MaxWorksheetNameLength.
' Returns the cleaned name by reference.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CleanWorksheetName(ByRef Name As String)

    ' No special error handling.
    On Error Resume Next
    
    ' Maximum length of a worksheet name in Excel.
    Const MaxWorksheetNameLength    As Long = 31
    ' String containing all not allowed characters.
    Const InvalidCharacters         As String = "\/:*?<>[]"
    ' Character to replace not allowed characters.
    Const ReplaceCharacter          As String * 1 = "-"
    ' Character ellipsis: ….
    Const Ellipsis                  As String * 1 = "…"
    
    Dim Length        As Integer
    Dim Position      As Integer
    Dim Character     As String
    Dim TrimmedName   As String
    
    ' Strip doubled spaces.
    While InStr(Name, Space(2)) > 0
        Name = Replace(Name, Space(2), Space(1))
    Wend
    ' Strip leading and trailing spaces.
    TrimmedName = Trim(Name)
    ' Limit length of name.
    If Len(TrimmedName) > MaxWorksheetNameLength Then
        TrimmedName = Left(TrimmedName, MaxWorksheetNameLength - 1) & Ellipsis
    End If
    Length = Len(TrimmedName)
    For Position = 1 To Length Step 1
        Character = Mid(TrimmedName, Position, 1)
        If InStr(InvalidCharacters, Character) > 0 Then
            Mid(TrimmedName, Position) = ReplaceCharacter
        End If
    Next
    
    ' Return cleaned name.
    Name = TrimmedName

End Sub

' Replaces characters in CodeName that are not allowed in a module codename.
' Truncates length of CodeName to MaxModuleCodeNameLength.
' Returns the cleaned name by reference.
'
' 2017-03-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CleanModuleCodeName(ByRef CodeName As String)
    
    ' Maximum length of a module codename in Excel.
    Const MaxModuleCodeNameLength   As Long = 31
    ' String containing all not allowed characters.
    Const InvalidCharacters         As String = "\/,;.:*?'`""<>|()[]{} @#$%&=+-~^"
    ' String containing all not allowed leading characters.
    Const InvalidLeadCharacters     As String = "_0123456789"
    
    ' Character to replace not allowed characters.
    Const ReplaceCharacter          As String * 1 = "_"
    ' Character to replace not allowed characters.
    Const ReplaceLeadCharacter      As String * 1 = "M"
    
    Dim Length          As Integer
    Dim Position        As Integer
    Dim Character       As String
    Dim TrimmedCodeName As String
    
    ' Strip doubled spaces.
    While InStr(CodeName, Space(2)) > 0
        CodeName = Replace(CodeName, Space(2), Space(1))
    Wend
    ' Strip leading and trailing spaces and limit length of codename.
    TrimmedCodeName = Left(Trim(CodeName), MaxModuleCodeNameLength)
    Length = Len(TrimmedCodeName)
    ' Replace invalid characters.
    For Position = 1 To Length Step 1
        Character = Mid(TrimmedCodeName, Position, 1)
        If InStr(InvalidCharacters, Character) > 0 Then
            Mid(TrimmedCodeName, Position) = ReplaceCharacter
        End If
    Next
    ' Replace a leading invalid character:
    Character = Left(TrimmedCodeName, 1)
    If InStr(InvalidLeadCharacters, Character) > 0 Then
        Mid(TrimmedCodeName, 1) = ReplaceLeadCharacter
    End If
    
    ' Return cleaned code name.
    CodeName = TrimmedCodeName
    
End Sub

' Replaces characters in Name that are not allowed in a name for a Named Range.
' Truncates length of Name to MaxNamedRangeNameLength.
' Returns the cleaned name by reference.
'
' 2017-08-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CleanNamedRangeName(ByRef Name As String)
    
    ' Maximum length of a name of a Named Range.
    Const MaxNamedRangeNameLength   As Long = 255
    ' String containing all not allowed characters.
    Const InvalidCharacters         As String = "/,;:*'`""<>|()[]{} @#$£&=+-~^"
    ' String containing all not allowed leading characters.
    Const InvalidLeadCharacters     As String = "!0123456789"
    
    ' Character to replace not allowed characters.
    Const ReplaceCharacter          As String * 1 = "_"
    ' Character to replace not allowed lead characters.
    Const ReplaceLeadCharacter      As String * 1 = "N"
    
    Dim Length          As Integer
    Dim Position        As Integer
    Dim Character       As String
    Dim TrimmedName     As String
    
    ' Strip doubled spaces.
    While InStr(Name, Space(2)) > 0
        Name = Replace(Name, Space(2), Space(1))
    Wend
    ' Strip leading and trailing spaces and limit length of Name.
    TrimmedName = Left(Trim(Name), MaxNamedRangeNameLength)
    Length = Len(TrimmedName)
    ' Replace invalid characters.
    For Position = 1 To Length Step 1
        Character = Mid(TrimmedName, Position, 1)
        If InStr(InvalidCharacters, Character) > 0 Then
            Mid(TrimmedName, Position) = ReplaceCharacter
        End If
    Next
    ' Replace a leading invalid character:
    Character = Left(TrimmedName, 1)
    If InStr(InvalidLeadCharacters, Character) > 0 Then
        Mid(TrimmedName, 1) = ReplaceLeadCharacter
    End If
    
    ' Return cleaned name.
    Name = TrimmedName
    
End Sub

' Returns a cleaned and truncated string suitable as a worksheet name.
'
' 2017-02-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TrimWorksheetName( _
    ByVal Name As String) _
    As String

    CleanWorksheetName Name

    TrimWorksheetName = Name

End Function

' Returns a cleaned and truncated string suitable as a code module name.
'
' 2017-03-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TrimModuleCodeName( _
    ByVal Name As String) _
    As String
    
    CleanModuleCodeName Name
    
    TrimModuleCodeName = Name
    
End Function

' Returns a cleaned and truncated string suitable as a name for a Named Range.
'
' 2017-03-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function TrimNamedRangeName( _
    ByVal Name As String) _
    As String

    CleanNamedRangeName Name

    TrimNamedRangeName = Name

End Function

' Searches in the collection of worksheets of a workbook for
' a worksheet named Name or with a name starting with Name.
' Returns the index of the worksheet if found.
' Returns zero (0) if the worksheet name is empty or not found.
'
' To lookup the index of a worksheet with an exact name, use:
'
'   Index = ThisWorkbook.Worksheets("Exact Name").Index
'
' 2017-09-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorksheetIndex( _
    ByRef Workbook As Workbook, _
    ByVal Name As String) _
    As Integer

    Dim Worksheet   As Excel.Worksheet
    Dim Index       As Integer
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    ElseIf Name <> "" Then
        ' Loop worksheets.
        For Each Worksheet In Workbook.Worksheets
            If InStr(1, Worksheet.Name, Name, vbTextCompare) = 1 Then
                Index = Worksheet.Index
                Exit For
            End If
        Next
    End If
    
    Set Worksheet = Nothing
    
    WorksheetIndex = Index

End Function

' Searches the collection of worksheets of a workbook for
' a worksheet named Name or with a name starting with Name.
' Returns the full name of the worksheet if found.
' Returns an empty string ("") if the worksheet name is not found.
'
' 2017-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WorksheetName( _
    ByRef Workbook As Workbook, _
    ByVal Name As String) _
    As String

    Dim Worksheet   As Excel.Worksheet
    Dim FullName    As String
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    ElseIf Name <> "" Then
        For Each Worksheet In Workbook.Worksheets
            If InStr(1, Worksheet.Name, Name, vbTextCompare) = 1 Then
                FullName = Worksheet.Name
                Exit For
            End If
        Next
    End If
    
    Set Worksheet = Nothing
    
    WorksheetName = FullName

End Function

' Freezes a worksheet pane top-left down to the top-left
' corner of the cell of RowIndex and ColumIndex.
'
' If RowIndex or ColumnIndex is less than 2 or omitted,
' only columns or rows respectively are frozen.
' If RowIndex and ColumnIndex are less than 2 or omitted,
' freezing of the worksheet is terminated.
'
' 2017-09-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub FreezeWorksheet( _
    ByVal Index As Variant, _
    Optional ByVal RowIndex As Long = 1, _
    Optional ByVal ColumnIndex As Long = 1)
    
    Const NoSplitIndex  As Long = 0
    
    Dim Freeze          As Boolean
    Dim CallIndex       As Long
    Dim ScreenUpdating  As Boolean
    
    ' Switching of the active window may happen, so
    ' disable screen updating while freezing is set.
    ScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    DoEvents
    
    ' Record the index of the currently active worksheet.
    CallIndex = ThisWorkbook.ActiveSheet.Index
    
    ' Activate the worksheet to freeze.
    ThisWorkbook.Worksheets(Index).Activate
    
    ' Hide row and column numbers.
    ActiveWindow.DisplayHeadings = False
    ' Hide formulabar.
    Application.DisplayFormulaBar = False
    
    ' Determine wether to freeze or to terminate freezing.
    Freeze = (RowIndex > 1 Or ColumnIndex > 1)
    If Freeze Then
        ' Remove an already set split.
        If ActiveWindow.Split = True Then
            ActiveWindow.Split = False
        End If
        ' Avoid errors.
        If RowIndex < 1 Then
            RowIndex = 1
        End If
        If ColumnIndex < 1 Then
            ColumnIndex = 1
        End If
        ' Set coordinates and apply freezing.
        ActiveWindow.SplitRow = RowIndex - 1
        ActiveWindow.SplitColumn = ColumnIndex - 1
        ActiveWindow.FreezePanes = True
    Else
        ' Terminate split and freeze.
        ActiveWindow.SplitRow = NoSplitIndex
        ActiveWindow.SplitColumn = NoSplitIndex
        ActiveWindow.Split = False
    End If
    
    ' Return to the previously active worksheet.
    DoEvents
    ThisWorkbook.Worksheets(CallIndex).Activate
    ' Restore status of screen updating.
    Application.ScreenUpdating = ScreenUpdating

End Sub

' Creates or adjusts a Named Range belonging to a workbook or a worksheet
' and located on any worksheet in the workbook.
' Parameter Name can either be the full name ("WorksheetName!NamedRangeName")
' or just the base name of the Named Range.
'
' 2017-09-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SetWorkbookNamedRange( _
    ByVal Index As Variant, _
    ByVal Name As String, _
    Optional ByVal RangeIndex As Variant = Null, _
    Optional ByVal RowIndex As Long, _
    Optional ByVal ColumnIndex As Long, _
    Optional ByVal RowSize As Long, _
    Optional ByVal ColumnSize As Long, _
    Optional ByVal Comment As Variant = Null) _
    As Long

    ' Reference to lost parent worksheet.
    Const Missing   As String = "=#REF"
    
    Dim Workbook    As Excel.Workbook
    Dim Worksheet   As Excel.Worksheet
    Dim Range       As Excel.Range
    Dim NameItem    As Excel.Name
    
    Dim NameParts   As Variant
    Dim SheetName   As String
    Dim NameIndex   As Long
    
    ' Don't waste time when nothing can be done.
    If Name = "" Then Exit Function
    
    If IsNull(Index) Then
        Set Workbook = ThisWorkbook
    ElseIf Index = 0 Then
        Set Workbook = ThisWorkbook
    Else
        ' Raise error if Index is invalid.
        Set Workbook = Application.Workbooks(Index)
    End If
    
    NameParts = Split(Name, NamedRangeSeparator)
    ' Retrieve/build names of Named Ranges to search.
    If UBound(NameParts) = 1 Then
        ' Name is on the form "WorksheetName!NamedRangeName".
        ' The Named Range is owned by the worksheet.
        SheetName = NameParts(0)
    Else
        ' Name is on the form "NamedRangeName".
        ' The Named Range is owned by the workbook.
    End If
    
    ' Look up the index of the Named Range if it already exists.
    For Each NameItem In Workbook.Names
        If StrComp(NameItem.Name, Name, vbTextCompare) = 0 Then
            If Split(NameItem.RefersTo, "!")(0) = Missing Then
                ' The parent worksheet has been deleted, thus the reference to it has been lost.
                ' Delete the Named Range.
                Workbook.Names(Name).Delete
                Exit For
            Else
                NameIndex = NameItem.Index
            End If
            Exit For
        End If
    Next

    If RowIndex = 0 And ColumnIndex = 0 And RowSize = 0 And ColumnSize = 0 Then
        ' Delete the Named Range.
        If NameIndex > 0 Then
            Workbook.Names(Name).Delete
        End If
    Else
        ' Find existing coordinates and replace missing parameters
        ' for an existing Named Range.
        If NameIndex > 0 Then
            Set Range = NameItem.RefersToRange
            If RowIndex = 0 Then
                RowIndex = Range.Rows(1).Row
            End If
            If ColumnIndex = 0 Then
                ColumnIndex = Range.Columns(1).Column
            End If
            If RowSize = 0 Then
                RowSize = Range.Rows.Count
            End If
            If ColumnSize = 0 Then
                ColumnSize = Range.Columns.Count
            End If
        End If
        
        ' Create or adjust the Named Range.
        If Not IsNull(RangeIndex) Then
            ' Override a the worksheet info of a passed Name like "WorksheetName!NamedRangeName".
        Else
            ' No index is specified for the worksheet containing the Named Range.
            ' If the Named Range exists, retrieve the containing worksheet from this.
            ' Else a worksheet name from the passed Name is used as container worksheet.
            If NameIndex > 0 Then
                ' Named Range exists and refers to "=WorksheetName!NamedRangeName".
                ' Extract "=WorksheetName".
                ' Strip leading "=" and find index of the worksheet.
                RangeIndex = Workbook.Worksheets(Mid(Split(NameItem.RefersTo, NamedRangeSeparator)(0), 2)).Index
            Else
                ' Use the worksheet name from the passed Name.
                RangeIndex = SheetName
            End If
        End If
        ' Raise error if RangeIndex is not found or invalid.
        Set Worksheet = Workbook.Worksheets(RangeIndex)
        
        ' Set coordinates of the Named Range.
        Set Range = Worksheet.Range( _
            Worksheet.Cells(RowIndex, ColumnIndex), _
            Worksheet.Cells(RowIndex + RowSize - 1, ColumnIndex + ColumnSize - 1))
        ' Create or adjust the Named Range.
        Workbook.Names.Add Name, Range
        If Not IsNull(Comment) Then
            ' Adjust Comment. An empty string will clear the comment.
            Workbook.Names(Name).Comment = Comment
        End If
        ' Retrieve workbook-level index of the touched or created Named Range.
        NameIndex = Workbook.Names(Name).Index
    End If
    
    Set NameItem = Nothing
    Set Range = Nothing
    Set Worksheet = Nothing
    Set Workbook = Nothing
    
    ' Return workbook-level index of the Named Range.
    SetWorkbookNamedRange = NameIndex

End Function

' Creates or adjusts a Named Range belonging to a worksheet
' located on the owner worksheet or any other worksheet.
' Parameter Name can either be the full name ("WorksheetName!NamedRangeName")
' or just the base name of the Named Range.
'
' 2017-03-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SetWorksheetNamedRange( _
    ByVal OwnerIndex As Variant, _
    ByVal Name As String, _
    Optional ByVal RangeIndex As Variant = Null, _
    Optional ByVal RowIndex As Long, _
    Optional ByVal ColumnIndex As Long, _
    Optional ByVal RowSize As Long, _
    Optional ByVal ColumnSize As Long, _
    Optional ByVal Comment As Variant = Null) _
    As Long

    Dim OwnerSheet  As Excel.Worksheet
    Dim RangeSheet  As Excel.Worksheet
    Dim Range       As Excel.Range
    Dim NameItem    As Excel.Name
    
    Dim NameParts   As Variant
    Dim FullName    As String
    Dim OwnerName   As String
    Dim BaseName    As String
    Dim NameIndex   As Long
    
    ' Don't waste time when nothing can be done.
    If Name = "" Then Exit Function
    
    NameParts = Split(Name, NamedRangeSeparator)
    ' Raise error if OwnerIndex is invalid.
    Set OwnerSheet = ThisWorkbook.Worksheets(OwnerIndex)
    ' Retrieve/build names of Named Ranges to search.
    OwnerName = OwnerSheet.Name
    BaseName = NameParts(UBound(NameParts))
    FullName = OwnerName & NamedRangeSeparator & BaseName
''
FullName = BaseName
''
    ' Look up the index of the Named Range if it already exists.
    For Each NameItem In OwnerSheet.Names
        If NameItem.Name = FullName Then
            NameIndex = ThisWorkbook.Names(FullName).Index
            Exit For
        End If
    Next
    
    If RowIndex = 0 And ColumnIndex = 0 And RowSize = 0 And ColumnSize = 0 Then
        ' Delete the Named Range.
        If NameIndex > 0 Then
            OwnerSheet.Names(FullName).Delete
        End If
    Else
        ' Find existing coordinates and replace missing parameters
        ' for an existing Named Range.
        If NameIndex > 0 Then
            Set Range = NameItem.RefersToRange
            If RowIndex = 0 Then
                RowIndex = Range.Rows(1).Row
            End If
            If ColumnIndex = 0 Then
                ColumnIndex = Range.Columns(1).Column
            End If
            If RowSize = 0 Then
                RowSize = Range.Rows.Count
            End If
            If ColumnSize = 0 Then
                ColumnSize = Range.Columns.Count
            End If
        End If
        
        ' Create or adjust the Named Range.
        If Not IsNull(RangeIndex) Then
            ' Override the worksheet info of a passed Name like "WorksheetName!NamedRangeName".
        Else
            ' No name or index is specified for the worksheet containing the Named Range.
            ' If the Named Range exists, retrieve the containing worksheet from this.
            ' Else the owner worksheet is also the container worksheet.
            If NameIndex > 0 Then
                ' Named Range exists and refers to "=WorksheetName!NamedRangeName".
                ' Extract "=WorksheetName".
                ' Strip leading "=" and find index of the worksheet.
                RangeIndex = ThisWorkbook.Worksheets(Mid(Split(NameItem.RefersTo, NamedRangeSeparator)(0), 2)).Index
            Else
                ' Use the owner worksheet also as container worksheet.
                RangeIndex = OwnerIndex
            End If
        End If
        ' Raise error if RangeIndex is not found or invalid.
        Set RangeSheet = ThisWorkbook.Worksheets(RangeIndex)
    
        ' Set coordinates of the Named Range.
        Set Range = RangeSheet.Range( _
            RangeSheet.Cells(RowIndex, ColumnIndex), _
            RangeSheet.Cells(RowIndex + RowSize - 1, ColumnIndex + ColumnSize - 1))
        ' Create or adjust the Named Range.
        OwnerSheet.Names.Add FullName, Range
        If Not IsNull(Comment) Then
            ' Adjust Comment. An empty string will clear the comment.
            OwnerSheet.Names(FullName).Comment = Comment
        End If
        ' Retrieve workbook-level index of the touched or created Named Range.
        NameIndex = OwnerSheet.Names(FullName).Index
    End If
    
    Set NameItem = Nothing
    Set Range = Nothing
    Set RangeSheet = Nothing
    Set OwnerSheet = Nothing
    
    ' Return workbook-level index of the Named Range.
    SetWorksheetNamedRange = NameIndex
    
End Function

' Turns columns and rows headings on or off for all worksheets
' in the current workbook.
'
' 2017-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub DisplayHeadings(ByVal Show As Boolean)
  
    Dim Worksheet   As Excel.Worksheet
    
    Dim Index       As Variant
    
    ' Avoid flickering.
    Application.ScreenUpdating = False
    ' Record current window.
    Index = Application.ActiveWindow.ActiveSheet.Index
    
    For Each Worksheet In ThisWorkbook.Worksheets
        Worksheet.Activate
        ActiveWindow.DisplayHeadings = Show
    Next Worksheet
    
    ' Reactivate current window.
    ThisWorkbook.Worksheets(Index).Activate
    ' Restore screen updating.
    Application.ScreenUpdating = True
    
    Set Worksheet = Nothing

End Sub

' Turns the Named Range selector and formula bar on or off.
'
' 2017-04-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub DisplayFormulaBar(ByVal Show As Boolean)

    Application.DisplayFormulaBar = Show

End Sub

' Create localized string expression for True or False.
'
' 2017-06-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LocalizedBoolean( _
    ByVal Expression As Boolean) _
    As String
    
    Dim Result  As String
    
    Result = UCase(Format(Expression, "True/False"))
    
    LocalizedBoolean = Result

End Function

' Returns a formula that will return the full path of this workbook:
'
'   SUBSTITUTE(LEFT(CELL("filename",A1),FIND("]",CELL("filename",A1))-1),"[","")
'
' 2017-06-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormulaWorkbookPath() As String

    Const Formula   As String = _
        "SUBSTITUTE(LEFT(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))-1),""["","""")"
    
    FormulaWorkbookPath = Formula

End Function

' Returns a formula that will return the current file name of this workbook:
'
'   REPLACE(LEFT(CELL("filename",A1),FIND("]",CELL("filename",A1))-1),1,FIND("[",CELL("filename",A1)),"")
'
' 2017-06-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormulaWorkbookFilename() As String

    Const Formula   As String = _
        "REPLACE(LEFT(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))-1),1,FIND(""["",CELL(""filename"",A1)),"""")"

    FormulaWorkbookFilename = Formula
    
End Function

'
'   SubAddress must be on the form:
'       "'Name of some worksheet'!AddressOfRangeOrCell"
'   like:
'       "'My Sheet Name'!A1"
'
Public Sub AddLocalHyperlink( _
    ByVal Anchor As Excel.Range, _
    ByVal Target As Excel.Range, _
    Optional ByVal ScreenTip As Variant, _
    Optional ByVal TextToDisplay As Variant, _
    Optional ByVal Parent As Excel.Worksheet)

    ' Address is not used for links to local addresses.
    Const Address   As String = ""
    
    Dim Worksheet   As Excel.Worksheet
    Dim Hyperlink   As Excel.Hyperlink
    
    Dim SubAddress  As String
    
    ' Exit if there is nothing to do.
    If Anchor Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    If Not Parent Is Nothing Then
        Set Worksheet = Parent
    Else
        Set Worksheet = Anchor.Parent
    End If
    SubAddress = HyperlinkSubaddress(Target)
    
    ' Save current range/cell formatting.
    CopyFont Anchor.Font, Nothing
    ' Create hyperlink.
    Set Hyperlink = Worksheet.Hyperlinks.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
    ' Restore range/cell formatting.
    CopyFont Nothing, Anchor.Font
    
    Set Hyperlink = Nothing
    Set Worksheet = Nothing

End Sub

' Copy Font settings from or to a range.
'
' Usage:
'
'   Copy and paste and keep in store:
'       CopyFont RangeToCopyFrom.Font, RangeToPasteTo.Font, False
'
'   Copy and paste and erase from store:
'       CopyFont RangeToCopyFrom.Font, RangeToPasteTo.Font [, True]
'
'   Copy and store:
'       CopyFont RangeToCopyFrom.Font, Nothing
'
'   Paste and keep in store:
'       CopyFont Nothing, RangeToPasteTo.Font, False
'
'   Paste and erase from store:
'       CopyFont Nothing, RangeToPasteTo [, True]
'
' 2017-07-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CopyFont( _
    ByRef SourceFont As Excel.Font, _
    ByRef TargetFont As Excel.Font, _
    Optional ByVal Reset As Boolean = True)
    
    Static SavedFont As VBA.Collection
    
    If Not SourceFont Is Nothing Then
        ' Store Font properties.
        Set SavedFont = New VBA.Collection
        
        ' Skip .Application
        SavedFont.Add SourceFont.Background, "Background"
        SavedFont.Add SourceFont.Bold, "Bold"
        SavedFont.Add SourceFont.Color, "Color"
        SavedFont.Add SourceFont.ColorIndex, "ColorIndex"
        ' Skip .Creator
        SavedFont.Add SourceFont.FontStyle, "FontStyle"
        SavedFont.Add SourceFont.Italic, "Italic"
        SavedFont.Add SourceFont.Name, "Name"
        ' Skip .Parent
        SavedFont.Add SourceFont.Size, "Size"
        SavedFont.Add SourceFont.Strikethrough, "Strikethrough"
        SavedFont.Add SourceFont.Subscript, "Subscript"
        SavedFont.Add SourceFont.Superscript, "Superscript"
        ' Skip .ThemeColor
        SavedFont.Add SourceFont.ThemeFont, "ThemeFont"
        SavedFont.Add SourceFont.TintAndShade, "TintAndShade"
        SavedFont.Add SourceFont.Underline, "Underline"
    End If
    
    If Not TargetFont Is Nothing Then
        If Not SavedFont Is Nothing Then
            If SavedFont.Count > 0 Then
                ' Retrieve Font properties.
                
                ' Skip .Application
                TargetFont.Background = SavedFont("Background")
                TargetFont.Bold = SavedFont("Bold")
                TargetFont.Color = SavedFont("Color")
                TargetFont.ColorIndex = SavedFont("ColorIndex")
                ' Skip .Creator
                TargetFont.FontStyle = SavedFont("FontStyle")
                TargetFont.Italic = SavedFont("Italic")
                TargetFont.Name = SavedFont("Name")
                ' Skip .Parent
                TargetFont.Size = SavedFont("Size")
                TargetFont.Strikethrough = SavedFont("Strikethrough")
                TargetFont.Subscript = SavedFont("Subscript")
                TargetFont.Superscript = SavedFont("Superscript")
                ' Skip .ThemeColor
                TargetFont.ThemeFont = SavedFont("ThemeFont")
                TargetFont.TintAndShade = SavedFont("TintAndShade")
                TargetFont.Underline = SavedFont("Underline")
            End If
        End If
        
        If Reset = True Then
            Set SavedFont = Nothing
        End If
    End If

End Sub

' Copy Interior settings from or to a range.
'
' Usage:
'
'   Copy and paste and keep in store:
'       CopyInterior RangeToCopyFrom.Interior, RangeToPasteTo.Interior, False
'
'   Copy and paste and erase from store:
'       CopyInterior RangeToCopyFrom.Interior, RangeToPasteTo.Interior [, True]
'
'   Copy and store:
'       CopyInterior RangeToCopyFrom.Interior, Nothing
'
'   Paste and keep in store:
'       CopyInterior Nothing, RangeToPasteTo.Interior, False
'
'   Paste and erase from store:
'       CopyInterior Nothing, RangeToPasteTo [, True]
'
' 2017-07-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CopyInterior( _
    ByRef SourceInterior As Excel.Interior, _
    ByRef TargetInterior As Excel.Interior, _
    Optional ByVal Reset As Boolean = True)
    
    Static SavedInterior As VBA.Collection
    
    If Not SourceInterior Is Nothing Then
        ' Store Interior properties.
        Set SavedInterior = New VBA.Collection
        
        ' Skip .Application
        SavedInterior.Add SourceInterior.Color, "Color"
        SavedInterior.Add SourceInterior.ColorIndex, "ColorIndex"
        ' Skip .Creator
        If SourceInterior.Gradient Is Nothing Then
            SavedInterior.Add Null, "Gradient"
        Else
            SavedInterior.Add SourceInterior.Gradient, "Gradient"
        End If
        ' Skip .InvertIfNegative
        ' Skip .Parent
        SavedInterior.Add SourceInterior.Pattern, "Pattern"
        SavedInterior.Add SourceInterior.PatternColor, "PatternColor"
        SavedInterior.Add SourceInterior.PatternColorIndex, "PatternColorIndex"
        ' Skip .PatternThemeColor
        SavedInterior.Add SourceInterior.PatternTintAndShade, "PatternTintAndShade"
        ' Skip .ThemeColor
        SavedInterior.Add SourceInterior.TintAndShade, "TintAndShade"
    End If
    
    If Not TargetInterior Is Nothing Then
        If Not SavedInterior Is Nothing Then
            If SavedInterior.Count > 0 Then
                ' Retrieve Interior properties.
                
                ' Skip .Application
                TargetInterior.Color = SavedInterior("Color")
                TargetInterior.ColorIndex = SavedInterior("ColorIndex")
                ' Skip .Creator
                If Not IsNull(SavedInterior("Gradient")) Then
                    TargetInterior.Gradient = SavedInterior("Gradient")
                End If
                ' Skip .InvertIfNegative
                ' Skip .Parent
                TargetInterior.Pattern = SavedInterior("Pattern")
                TargetInterior.PatternColor = SavedInterior("PatternColor")
                TargetInterior.PatternColorIndex = SavedInterior("PatternColorIndex")
                ' Skip .PatternThemeColor
                TargetInterior.PatternTintAndShade = SavedInterior("PatternTintAndShade")
                ' Skip .ThemeColor
                If Not IsNull(SavedInterior("TintAndShade")) Then
                    TargetInterior.TintAndShade = SavedInterior("TintAndShade")
                End If
            End If
        End If
        
        If Reset = True Then
            Set SavedInterior = Nothing
        End If
    End If

End Sub

'   SubAddress must be on the form:
'       "'Name of some worksheet'!AddressOfRangeOrCell"
'   like:
'       "'My Sheet Name'!A1"
'
Public Function HyperlinkSubaddress( _
    ByVal Range As Excel.Range) _
    As String

    Const Format    As String = "'{0}'" & NamedRangeSeparator & "{1}"
    
    Dim Name        As String
    Dim Address     As String
    Dim SubAddress  As String
    
    If Not Range Is Nothing Then
        Name = Range.Parent.Name
        Address = Range.Address
        SubAddress = Replace(Replace(Format, "{0}", Name), "{1}", Address)
    End If
    
    HyperlinkSubaddress = SubAddress
    
End Function

Public Sub ListHyperlinks(Optional ByVal Delete As Boolean)

    Dim Worksheet   As Excel.Worksheet
    Dim Hyperlink   As Excel.Hyperlink
    Dim Range       As Excel.Range
    
    For Each Worksheet In ThisWorkbook.Worksheets
        For Each Hyperlink In Worksheet.Hyperlinks
            Debug.Print Worksheet.Name, Hyperlink.Range.Address
            If Delete = True Then
                Set Range = Hyperlink.Range
                CopyFont Range.Font, Nothing
                CopyInterior Range.Interior, Nothing
                Hyperlink.Delete
                CopyFont Nothing, Range.Font
                CopyInterior Nothing, Range.Interior
            End If
        Next
    Next
    
    Set Range = Nothing
    Set Hyperlink = Nothing
    Set Worksheet = Nothing
    
End Sub

' Read a Unicode character value from a range (cell).
'
' 2017-07-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function GetRangeUnicode( _
    ByVal Range As Excel.Range) _
    As Long

    Const SignedIntegerMax  As Long = 2 ^ 16
    
    Dim Value   As Long

    If Not Range Is Nothing Then
        If Not IsEmpty(Range.Value) Then
            ' Alternative:
            ' Value = Excel.WorksheetFunction.Unicode(Range.Value)
            Value = (SignedIntegerMax + AscW(Range.Value)) Mod SignedIntegerMax
        End If
    End If
    
    GetRangeUnicode = Value
    
End Function

' Write a Unicode character value to a range (cell).
'
' 2017-07-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SetRangeUnicode( _
    ByVal Range As Excel.Range, _
    ByVal Value As Long) _
    As Boolean
    
    Const SignedIntegerMax      As Long = 2 ^ 16
    Const UnsignedIntegerMin    As Long = -2 ^ 15
    
    Dim Result  As Boolean
    
    If Not Range Is Nothing Then
        If Value > UnsignedIntegerMin And Value < SignedIntegerMax - 1 Then
            Range.Value = ChrW((SignedIntegerMax + Value) Mod SignedIntegerMax)
            Result = True
        End If
    End If
    
    SetRangeUnicode = Result

End Function

' Retrieves the index of the WorkbookConnection of a ListObject
' having a QueryTable.
'
Public Function ListObjectConnection( _
    ByVal ListObject As Excel.ListObject) _
    As Long
    
    Dim ConnectionIndex As Long
    
    If Not ListObject Is Nothing Then
        If Not ListObject.QueryTable Is Nothing Then
            ConnectionIndex = WorkbookConnectionIndex( _
                ListObject.Parent.Parent, _
                ListObject.QueryTable.WorkbookConnection.Name)
        End If
    End If
    
    ListObjectConnection = ConnectionIndex
    
End Function

' Retrieves the current index of a code module of Workbook
' from itself, its name, or its index.
' Returns 0 (zero) if not found.
'
' 2017-09-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ModuleIndex( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Name As Variant) _
    As Long
    
    Dim Index   As Long
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    End If
    
    Index = Workbook.VBProject.VBComponents.Count
    Select Case VarType(Name)
        Case vbByte, vbInteger, vbLong
            If Name > 0 And Name <= Index Then
                Index = Name
            Else
                Index = 0
            End If
        Case vbString
            ' Name could be the name, or the index of a module,
            ' or a module object.
            If Trim(Name) = CStr(Int(Val(Name))) Then
                ' Name could be the index of a module.
                If Val(Name) > 0 And Val(Name) <= Index Then
                    ' Name is supposed to represent the index.
                    Index = Val(Name)
                    ' Stop further processing.
                    Name = ""
                End If
            End If
            
            If Trim(Name) = "" Then
                ' No module for an empty name.
                Index = 0
            Else
                ' Name could be the name of a module.
                ' Loop the names of the module collection.
                Do
                    If StrComp(Workbook.VBProject.VBComponents(Index).Name, Name, vbTextCompare) = 0 Then
                        ' Name exists.
                        Exit Do
                    End If
                    Index = Index - 1
                Loop Until Index = 0
            End If
        Case vbObject
            ' Would only be Nothing.
            Index = 0
        Case Else
            ' Skip other invalid cases.
            Index = 0
    End Select
    
    ModuleIndex = Index
    
End Function

' Retrieves the current index of a Workbook from
' itself, its name, or its index.
' Returns 0 (zero) if not found.
'
' Note:
'   After adding or deleting a workbook, the collection
'   will automatically be reordered by Name asc, thus
'   workbooks after the inserted or deleted workbook
'   will have their index changed by 1.
'
' 2017-09-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function WorkbookIndex( _
    ByVal Name As Variant) _
    As Long
    
    Dim Index   As Long
    
    Index = Workbooks.Count
    Select Case VarType(Name)
        Case vbByte, vbInteger, vbLong
            If Name > 0 And Name <= Index Then
                Index = Name
            Else
                Index = 0
            End If
        Case vbString
            ' Name could be the name, or the index of a workbook,
            ' or a Workbook object.
            If Trim(Name) = CStr(Int(Val(Name))) Then
                ' Name could be the index of a workbook.
                If Val(Name) > 0 And Val(Name) <= Index Then
                    ' Name is supposed to represent the index.
                    Index = Val(Name)
                    ' Stop further processing.
                    Name = ""
                End If
            End If
            
            If Trim(Name) = "" Then
                ' No workbook for an empty name.
                Index = 0
            Else
                ' Name could be the name of a workbook.
                ' Loop the names of the workbook collection.
                Do
                    If StrComp(Workbooks.Item(Index).Name, Name, vbTextCompare) = 0 Then
                        ' Name exists.
                        Exit Do
                    End If
                    Index = Index - 1
                Loop Until Index = 0
            End If
        Case vbObject
            If TypeName(Name) = TypeName(ThisWorkbook) Then
                ' Name is a workbook object.
                ' Loop the workbook collection.
                Do
                    If Name Is Workbooks.Item(Index) Then
                        Exit Do
                    End If
                    Index = Index - 1
                Loop Until Index = 0
            Else
                ' Would most likely be Nothing.
                Index = 0
            End If
        Case Else
            ' Skip other invalid cases.
            Index = 0
    End Select
    
    WorkbookIndex = Index
    
End Function

' Retrieves the current index of a WorkbookConnection of a Workbook
' from itself, its name, or its index.
' Returns 0 (zero) if not found.
'
' Note:
'   After adding or deleting a connection, the collection
'   will automatically be reordered by Name asc, thus
'   connections after the inserted or deleted connection
'   will have their index changed by 1.
'
' 2017-09-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function WorkbookConnectionIndex( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Name As Variant) _
    As Long
    
    Dim Index   As Long
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    End If
    
    Index = Workbook.Connections.Count
    Select Case VarType(Name)
        Case vbByte, vbInteger, vbLong
            If Name > 0 And Name <= Index Then
                Index = Name
            Else
                Index = 0
            End If
        Case vbString
            ' Name could be the name, or the index of a workbook connection,
            ' or a WorkbookConnection object.
            If Trim(Name) = CStr(Int(Val(Name))) Then
                ' Name could be the index of a workbook connection.
                If Val(Name) > 0 And Val(Name) <= Index Then
                    ' Name is supposed to represent the index.
                    Index = Val(Name)
                    ' Stop further processing.
                    Name = ""
                End If
            End If
            
            If Trim(Name) = "" Then
                ' No workbook connection for an empty name.
                Index = 0
            Else
                ' Name could be the name of a workbook connection.
                ' Loop the names of the workbook connection collection.
                Do
                    If StrComp(Workbook.Connections.Item(Index).Name, Name, vbTextCompare) = 0 Then
                        ' Name exists.
                        Exit Do
                    End If
                    Index = Index - 1
                Loop Until Index = 0
            End If
        Case vbObject
            ' Would only be Nothing.
            Index = 0
        Case Else
            ' Skip other invalid cases.
            Index = 0
    End Select
    
    WorkbookConnectionIndex = Index
    
End Function

' Retrieves the current index of a TableStyle of a Workbook
' from itself, its name, its localised name, or its index.
' Returns 0 (zero) if not found.
'
' Note:
'   After adding or deleting a table style, the collection
'   will automatically be reordered by BuiltIn and Name asc,
'   thus table styles after the inserted or deleted table
'   style will have their index changed by 1.
'
' 2017-09-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TableStyleIndex( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Name As Variant) _
    As Long
    
    Dim Index   As Long
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    End If
    
    Index = Workbook.TableStyles.Count
    Select Case VarType(Name)
        Case vbByte, vbInteger, vbLong
            If Name > 0 And Name <= Index Then
                Index = Name
            Else
                Index = 0
            End If
        Case vbString
            ' Name could be the name, the localised name, or the index of a table style,
            ' or a TableStyle object.
            If Trim(Name) = CStr(Int(Val(Name))) Then
                ' Name could be the index of a table style.
                If Val(Name) > 0 And Val(Name) <= Index Then
                    ' Name is supposed to represent the index.
                    Index = Val(Name)
                    ' Stop further processing.
                    Name = ""
                End If
            End If
            
            If Trim(Name) = "" Then
                ' No table style for an empty name.
                Index = 0
            Else
                ' Name could be the (localised) name of a table style.
                ' Loop the names of the table style collection.
                Do
                    If _
                        StrComp(Workbook.TableStyles.Item(Index).Name, Name, vbTextCompare) = 0 Or _
                        StrComp(Workbook.TableStyles.Item(Index).NameLocal, Name, vbTextCompare) = 0 Then
                        ' Name exists.
                        Exit Do
                    End If
                    Index = Index - 1
                Loop Until Index = 0
            End If
        Case vbObject
            ' Would only be Nothing.
            Index = 0
        Case Else
            ' Skip other invalid cases.
            Index = 0
    End Select
        
    TableStyleIndex = Index
    
End Function

' Checks if a table style by itself or its name, its localized name,
' or its index belongs to the TableStyles collection of Workbook.
'
' 2017-08-22. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function IsTableStyle( _
    ByVal Workbook As Excel.Workbook, _
    ByVal Index As Variant) _
    As Boolean
    
    Dim Result      As Boolean
    
    Result = CBool(TableStyleIndex(Workbook, Index))
    
    IsTableStyle = Result
    
End Function

' Export a table style from ThisWorkbook to another workbook,
' optionally with another name.
'
' The scenario will be:
'
'   Source          Target Name         Result
'   ------------------------------------------------------------------------
'   Any             Name of a built-in  False
'   Any             Other name          True. Overwriting an existing target
'   Built-in        None                True. Target named as source with digit suffix
'   Not built-in    None                True. Target named as source, overwriting an existing
'   Not found                           False
'
' The source table style can be specified as a TableStyle object or
' by its index, name, or localised name.
'
' 2017-08-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ExportTableStyle( _
    ByVal TargetFilename As String, _
    ByVal SourceTableStyleIndex As Variant, _
    Optional TargetTableStyleName As String) _
    As Boolean
    
    Dim Source  As Excel.Workbook
    Dim Target  As Excel.Workbook
    
    Dim Success As Boolean
    
    If Dir(TargetFilename, vbNormal) = "" Then
        ' No target file to open. Exit.
        Exit Function
    End If
    
    ' Open the target workbook silently for read/write.
    Application.ScreenUpdating = False
    Set Target = Workbooks.Open(TargetFilename, , False)
    Windows(Target.Name).Visible = False
    Application.ScreenUpdating = True
    
    Set Source = ThisWorkbook
    
    Success = Not TableStyleCopy(Source, Target, SourceTableStyleIndex, TargetTableStyleName) Is Nothing
    
    Target.Close True
    
    Set Source = Nothing
    Set Target = Nothing
    
    ExportTableStyle = Success
    
End Function

' Import a table style from another workbook to ThisWorkbook,
' optionally with another name.
'
' The scenario will be:
'
'   Source          Target Name         Result
'   ------------------------------------------------------------------------
'   Any             Name of a built-in  False
'   Any             Other name          True. Overwriting an existing target
'   Built-in        None                True. Target named as source with digit suffix
'   Not built-in    None                True. Target named as source, overwriting an existing
'   Not found                           False
'
' The source table style can be specified as a TableStyle object or
' by its index, name, or localised name.
'
' 2017-08-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ImportTableStyle( _
    ByVal SourceFilename As String, _
    ByVal SourceTableStyleIndex As Variant, _
    Optional TargetTableStyleName As String) _
    As Boolean
    
    Dim Source  As Excel.Workbook
    Dim Target  As Excel.Workbook
    
    Dim Success As Boolean
    
    If Dir(SourceFilename, vbNormal) = "" Then
        ' No source file to open. Exit.
        Exit Function
    End If
    
    ' Open the source workbook silently as read-only.
    Application.ScreenUpdating = False
    Set Source = Workbooks.Open(SourceFilename, , True)
    Windows(Source.Name).Visible = False
    Application.ScreenUpdating = True
    
    Set Target = ThisWorkbook
    
    Success = Not TableStyleCopy(Source, Target, SourceTableStyleIndex, TargetTableStyleName) Is Nothing
    
    Source.Close False
    
    Set Target = Nothing
    Set Source = Nothing
    
    ImportTableStyle = Success
    
End Function

' Returns a table style copied from one workbook to another,
' optionally with another name.
'
' The returned result will be:
'
'   Source          Target Name         Result
'   ------------------------------------------------------------------------
'   Any             Name of a built-in  Nothing
'   Any             Other name          With target name, overwriting an existing
'   Built-in        None                With source name with digit suffix
'   Not built-in    None                With source name, overwriting an existing
'   Not found                           Nothing
'
' The source table style can be specified as a TableStyle object by
' itself or its index, name, or localised name.
'
' 2017-09-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TableStyleCopy( _
    ByRef Source As Excel.Workbook, _
    ByRef Target As Excel.Workbook, _
    ByVal SourceTableStyleIndex As Variant, _
    Optional TargetTableStyleName As String) _
    As Excel.TableStyle
    
    Const TempTableStylePrefix  As String = "~tmp"
    
    Dim TableStyle              As Excel.TableStyle
    Dim TempTableStyle          As Excel.TableStyle
    
    Dim ActiveWorkbookIndex     As Long
    Dim TargetTableStyleIndex   As Long
    Dim SourceTableStyleName    As String
    Dim TempTableStyleName      As String
    
    Set TableStyle = Nothing
    
    If Source Is Nothing Or Target Is Nothing Then
        ' Nothing to do. Exit.
        Exit Function
    ElseIf Target.ReadOnly = True Then
        ' Cannot copy to a read-only workbook. Exit.
        Exit Function
    Else
        ' Obtain the numeric index of the source table style from its name or index.
        SourceTableStyleIndex = TableStyleIndex(Source, SourceTableStyleIndex)
        If SourceTableStyleIndex = 0 Then
            ' Source table style not found.
            ' Nothing to do. Exit.
            Exit Function
        End If
    End If
    
    TargetTableStyleName = Trim(TargetTableStyleName)
    TargetTableStyleIndex = TableStyleIndex(Target, TargetTableStyleName)
    If TargetTableStyleIndex > 0 Then
        If Target.TableStyles(TargetTableStyleIndex).BuiltIn = True Then
            ' Cannot copy to the name of a built-in table style.
            Exit Function
        End If
    End If
    SourceTableStyleName = Source.TableStyles(SourceTableStyleIndex).Name
    
    If TargetTableStyleName = "" Then
        If Source.TableStyles(SourceTableStyleName).BuiltIn = True Then
            ' Cannot copy a built-in table style to itself.
            ' Exit.
            Exit Function
        Else
            ' Preserve the table style name.
            TargetTableStyleName = SourceTableStyleName
        End If
    End If
    If TargetTableStyleName = SourceTableStyleName Then
        If Target Is Source Then
            ' Cannot copy a table style to itself.
            ' Return the table style as is.
            Set TableStyle = ThisWorkbook.TableStyles(SourceTableStyleName)
        End If
    End If
    
    ' Store name of active workbook.
    ActiveWorkbookIndex = WorkbookIndex(ActiveWorkbook)
    Target.Activate
    
    If TargetTableStyleIndex > 0 Then
        ' Table style exists.
        ' Reset table style of ListObjects having TargetTableStyleName applied.
        ReplaceListObjectsAllTableStyle Target, Target.TableStyles(TargetTableStyleName), Nothing
        ' Overwrite table style by a delete and, later, a duplicate.
        Target.TableStyles(TargetTableStyleName).Delete
    End If
    
    If Source Is Target Then
        ' Create copy.
        Set TableStyle = Target.TableStyles(SourceTableStyleName).Duplicate(TargetTableStyleName)
    Else
        If IsTableStyle(Target, SourceTableStyleName) Then
            ' Table style source name exists in target.
            ' Duplicate it with a temporary name.
            Randomize
            TempTableStyleName = TempTableStylePrefix & Mid(Str(Rnd), 3)
            Set TempTableStyle = Target.TableStyles(SourceTableStyleName).Duplicate(TempTableStyleName)
            ' Replace the table style of ListObjects having the source table style name applied.
            ReplaceListObjectsAllTableStyle Target, Target.TableStyles(SourceTableStyleName), TempTableStyle
            ' Delete table style name to allow the name to be added from the source.
            Target.TableStyles(SourceTableStyleName).Delete
        End If
        
        ' Import the table style using its name in the source.
        Set TableStyle = Target.TableStyles.Add(Source.TableStyles(SourceTableStyleIndex))
        With Source.TableStyles(SourceTableStyleIndex)
            ' Transfer ShowAsAvailable properties.
            TableStyle.ShowAsAvailablePivotTableStyle = .ShowAsAvailablePivotTableStyle
            TableStyle.ShowAsAvailableSlicerStyle = .ShowAsAvailableSlicerStyle
            TableStyle.ShowAsAvailableTableStyle = .ShowAsAvailableTableStyle
            TableStyle.ShowAsAvailableTimelineStyle = .ShowAsAvailableTimelineStyle
        End With
                    
        If TargetTableStyleName <> SourceTableStyleName Then
            ' Create renamed copy of the imported table style.
            Set TableStyle = TableStyle.Duplicate(TargetTableStyleName)
            ' Delete the imported table style.
            Target.TableStyles(SourceTableStyleName).Delete
            
            If Not TempTableStyle Is Nothing Then
                ' Restore the table style from its temporary copy.
                TempTableStyle.Duplicate SourceTableStyleName
                ' Restore the table style of ListObjects which had the source table style name applied.
                ReplaceListObjectsAllTableStyle Target, TempTableStyle, Target.TableStyles(SourceTableStyleName)
                ' Delete the temporary table style.
                TempTableStyle.Delete
            End If
        End If
    End If
    
    If TargetTableStyleIndex > 0 Then
        ' Restore table style of ListObjects previously having TargetTableStyleName applied.
        ReplaceListObjectsAllTableStyle Target, Nothing, Target.TableStyles(TargetTableStyleName)
    End If
    
    ' Reactivate the initially active workbook.
    Workbooks(ActiveWorkbookIndex).Activate
    
    Set TempTableStyle = Nothing
        
    Set TableStyleCopy = TableStyle
    
End Function

' Checks if a workbook object is ThisWorkbook.
'
' 2017-08-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsThisWorkbook( _
    ByRef Workbook As Excel.Workbook) _
    As Boolean
    
    Dim Result  As Boolean
    
    Result = Workbook Is ThisWorkbook
    
    IsThisWorkbook = Result
    
End Function

' Lists the table styles in use in Workbook and
' returns the count of tables (ListObjects).
'
' Output is like:
'
'   Worksheet index, table name, table has style, name of table style
'
' 2017-08-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ListActiveTableStyles( _
    Optional ByRef Workbook As Excel.Workbook) _
    As Integer

    Dim Worksheet       As Excel.Worksheet
    Dim ListObject      As Excel.ListObject
    
    Dim TableNameLength As Integer
    Dim ListCount       As Integer
    
    If Workbook Is Nothing Then
        Set Workbook = ThisWorkbook
    End If
    
    ' Find maximum length of the table names.
    For Each Worksheet In Workbook.Worksheets
        For Each ListObject In Worksheet.ListObjects
            If Len(ListObject.Name) > TableNameLength Then
                TableNameLength = Len(ListObject.Name)
            End If
        Next
        ListCount = ListCount + Worksheet.ListObjects.Count
    Next
    
    ' List the table styles currently assigned.
    For Each Worksheet In Workbook.Worksheets
        For Each ListObject In Worksheet.ListObjects
            ' If no table style, the ListObject.TableStyle object isn't Nothing, it is invalid.
            Debug.Print Worksheet.Index, Left(ListObject.Name & Space(TableNameLength), TableNameLength), IsObject(ListObject.TableStyle),
            If IsObject(ListObject.TableStyle) Then
                Debug.Print ListObject.TableStyle.Name
            Else
                Debug.Print "N/A"
            End If
        Next
    Next
    
    Set ListObject = Nothing
    Set Worksheet = Nothing
    
    ListActiveTableStyles = ListCount
    
End Function

' Lists the custom table styles of Workbook and
' returns the count of these.
'
' Output is like:
'
'   index of table style, ShowAsAvailableTableStyle, name of table style
'
' 2017-09-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ListCustomTableStyles( _
    Optional ByRef Workbook As Excel.Workbook) _
    As Integer
    
    Dim TableStyle  As Excel.TableStyle
    
    Dim ListCount   As Integer
    
    If Workbook Is Nothing Then
        Set Workbook = ThisWorkbook
    End If
    
    For Each TableStyle In Workbook.TableStyles
        If TableStyle.BuiltIn = False Then
            Debug.Print TableStyleIndex(Workbook, TableStyle), TableStyle.ShowAsAvailableTableStyle, TableStyle.Name
            ListCount = ListCount + 1
        End If
    Next
    
    Set TableStyle = Nothing
    
    ListCustomTableStyles = ListCount

End Function

' Returns True if the passed table style currently is assigned a
' list object on Workbook.
'
' 2017-08-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsTableStyleActive( _
    ByRef Workbook As Excel.Workbook, _
    ByVal Name As Variant) _
    As Boolean
    
    Dim Worksheet   As Excel.Worksheet
    Dim ListObject  As Excel.ListObject
    Dim TableStyle  As Excel.TableStyle
    
    Dim Index       As Long
    Dim Result      As Boolean
    
    If Workbook Is Nothing Then
        ' Nothing to do.
        Exit Function
    End If
    
    Index = TableStyleIndex(Workbook, Name)
    If Index > 0 Then
    
        Set TableStyle = Workbook.TableStyles(Index)
        
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                If IsObject(ListObject.TableStyle) Then
                    If ListObject.TableStyle = TableStyle Then
                        Result = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        Set TableStyle = Nothing
        Set ListObject = Nothing
        Set Worksheet = Nothing
        
    End If
    
    IsTableStyleActive = Result
    
End Function

' Replaces a list object's currently assigned table style if
' this matches parameter OldTableStyle.
' OldTableStyle can be Nothing.
'
' If parameter NewTableStyle is Nothing, the list object will be
' left with no assigned table style.
'
' 2017-08-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ReplaceListObjectTableStyle( _
    ByVal ListObject As Excel.ListObject, _
    ByVal OldTableStyle As Excel.TableStyle, _
    ByVal NewTableStyle As Excel.TableStyle)

    If ListObject Is Nothing Then
        ' Nothing to do.
    Else
        If Not OldTableStyle Is Nothing Then
            If Not IsTableStyle(ListObject.Parent.Parent, OldTableStyle) Then
                ' OldTableStyle must belong to the workbook of ListObject.
                Exit Sub
            End If
        End If
        If Not NewTableStyle Is Nothing Then
            If Not IsTableStyle(ListObject.Parent.Parent, NewTableStyle) Then
                ' NewTableStyle must belong to the workbook of ListObject.
                Exit Sub
            End If
        End If
        
        ' Replace table style.
        If OldTableStyle Is Nothing Then
            ' ListObject should have no table style.
            If Not IsObject(ListObject.TableStyle) Then
                ' ListObject does not have a table style.
                ' Apply the table style (can be Nothing).
                ListObject.TableStyle = NewTableStyle
            End If
        Else
            ' ListObject should have a table style.
            If IsObject(ListObject.TableStyle) Then
                ' ListObject does have a table style.
                ' Check that it is the one to replace.
                If ListObject.TableStyle = OldTableStyle Then
                    ' Apply the table style (can be Nothing).
                    ListObject.TableStyle = NewTableStyle
                End If
            End If
        End If
    End If
    
End Sub

' Replaces in Workbook each list object's currently assigned table style
' if this matches parameter OldTableStyle.
' OldTableStyle can be Nothing.
'
' If parameter NewTableStyle is Nothing, each list object will be
' left with no assigned table style.
'
' 2017-08-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ReplaceListObjectsAllTableStyle( _
    ByVal Workbook As Excel.Workbook, _
    ByVal OldTableStyle As Excel.TableStyle, _
    ByVal NewTableStyle As Excel.TableStyle)
    
    Dim Worksheet   As Excel.Worksheet
    Dim ListObject  As Excel.ListObject
    
    If Workbook Is Nothing Then
        ' Nothing to do.
    Else
        If Not OldTableStyle Is Nothing Then
            If Not IsTableStyle(Workbook, OldTableStyle) Then
                ' OldTableStyle must belong to the workbook.
                Exit Sub
            End If
        End If
        If Not NewTableStyle Is Nothing Then
            If Not IsTableStyle(Workbook, NewTableStyle) Then
                ' NewTableStyle must belong to the workbook.
                Exit Sub
            End If
        End If
    
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                ReplaceListObjectTableStyle ListObject, OldTableStyle, NewTableStyle
            Next
        Next
    End If

    Set ListObject = Nothing
    Set Worksheet = Nothing

End Sub

' Returns True if a ListboxObject named Name exists in Workbook.
'
' 2020-03-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsListObject( _
    ByRef Worksheet As Excel.Worksheet, _
    ByVal Name As String) _
    As Boolean
    
    Dim ListObject  As Excel.ListObject
    
    Dim Result      As Boolean
    
    If Not Worksheet Is Nothing Then
        For Each ListObject In Worksheet.ListObjects
            If ListObject.Name = Name Then
                Result = True
                Exit For
            End If
        Next
    End If
    
    IsListObject = Result

End Function


