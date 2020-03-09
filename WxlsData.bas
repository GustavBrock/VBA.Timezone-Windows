Attribute VB_Name = "WxlsData"
Option Explicit

' WxlsData v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Functions to maintain tables holding Windows timezone information
' retrieved from the Registry.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

' Required modules:
'   ByteUtil
'   DateUtil (or project VBA.Date)
'   WtziBase
'   WtziCore

' Required references:
'   Windows Script Host Object Model


' Timezone worksheet name.
Private Const TimezoneWorksheetName As String = "Data"

' Timezone table names.
Private Const TimezoneTableZone     As String = "WindowsTimezone"
Private Const TimezoneTableLocation As String = "WindowsTimezoneLocation"

' Timezone table positions.
Private Const TimezoneRowIndex      As Integer = 1
Private Const TimezoneIndex         As Integer = 1
Private Const TimezoneLocationIndex As Integer = 14

' Timezone field names.
' A field for the Registry key TZI is not included.
Private Const TimezoneMui           As String = "MUI"
Private Const TimezoneMuiDaylight   As String = "MUIDlt"
Private Const TimezoneMuiStandard   As String = "MUIStd"
Private Const TimezoneBias          As String = "Bias"
Private Const TimezoneName          As String = "Name"
Private Const TimezoneUtc           As String = "UTC"
Private Const TimezoneLocations     As String = "Locations"
Private Const TimezoneDlt           As String = "ZoneDlt"
Private Const TimezoneStd           As String = "ZoneStd"
Private Const TimezoneFirstEntry    As String = "FirstEntry"
Private Const TimezoneLastEntry     As String = "LastEntry"
Private Const TimezoneDisplay       As String = "Display"
Private Const TimezoneLocationId    As String = "Id"
Private Const TimezoneLocationMui   As String = "MUI"
Private Const TimezoneLocationName  As String = "Name"

' Create and prepare a worksheet to hold the tables for the timezones.
' Returns True if the worksheet and tables existed or were created.
'
' 2020-03-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CreateTimezoneData() As Boolean

    Dim Worksheet       As Excel.Worksheet
    
    Dim Result          As Boolean

    ' Fetch or create the worksheet holding the tables.
    Set Worksheet = WorksheetData
    
    If Not Worksheet Is Nothing Then
        ' Check that the tables are present. If not, they will be created.
        Result = CreateTimezoneDataTable(Worksheet, TimezoneTableZone)
        If Result = True Then
            Result = CreateTimezoneDataTable(Worksheet, TimezoneTableLocation)
        End If
    End If
    
    CreateTimezoneData = Result

End Function

' Create and prepare a table for the timezones.
' Returns True if the table existed or were created.
'
' 2020-03-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CreateTimezoneDataTable( _
    ByRef Worksheet As Excel.Worksheet, _
    ByVal TableName As String) _
    As Boolean

    Dim ListObject              As Excel.ListObject
    Dim Range                   As Excel.Range
    Dim ListColumn              As Excel.ListColumn
    
    Dim ColumnNames             As Variant
    Dim ColumnIndex             As Integer
    Dim Result                  As Boolean

    On Error GoTo CreateTimezoneDataTable_Error
    
    Select Case TableName
        Case TimezoneTableZone
            Set Range = Worksheet.Cells(TimezoneRowIndex, TimezoneIndex)
        Case TimezoneTableLocation
            Set Range = Worksheet.Cells(TimezoneRowIndex, TimezoneLocationIndex)
    End Select
    
    If Not IsListObject(Worksheet, TableName) Then
        ' Create the table.
        Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Range, , xlYes)
        ListObject.Name = TableName
        
        Select Case TableName
            Case TimezoneTableZone
                ColumnNames = Array( _
                    TimezoneMui, _
                    TimezoneMuiDaylight, _
                    TimezoneMuiStandard, _
                    TimezoneName, _
                    TimezoneBias, _
                    TimezoneUtc, _
                    TimezoneLocations, _
                    TimezoneDlt, _
                    TimezoneStd, _
                    TimezoneFirstEntry, _
                    TimezoneLastEntry, _
                    TimezoneDisplay)
                
            Case TimezoneTableLocation
                ColumnNames = Array( _
                    TimezoneLocationId, _
                    TimezoneLocationMui, _
                    TimezoneLocationName)
        End Select
                
        For ColumnIndex = LBound(ColumnNames) + 1 To UBound(ColumnNames)
            ListObject.ListColumns.Add
        Next
        For ColumnIndex = LBound(ColumnNames) To UBound(ColumnNames)
            Set ListColumn = ListObject.ListColumns(ColumnIndex + 1)
            ListColumn.Name = ColumnNames(ColumnIndex)
            ListColumn.Range.EntireColumn.AutoFit
        Next
        ' The table was created.
        Result = True
    Else
        ' The table is present.
        Result = True
    End If
    
    CreateTimezoneDataTable = Result

CreateTimezoneDataTable_Exit:
    Exit Function

CreateTimezoneDataTable_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WxlsData.CreateTimezoneDataTable."
    Resume CreateTimezoneDataTable_Exit
    
End Function

' Look up and return the worksheet holding the timezone tables.
' If not found, the worksheet will be created.
'
' 2020-03-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function WorksheetData() As Excel.Worksheet

    Dim Worksheet       As Excel.Worksheet
    
    Dim Index           As Integer

    On Error GoTo WorksheetData_Error

    Index = WorksheetIndex(ThisWorkbook, TimezoneWorksheetName)
    If Index = 0 Then
        Set Worksheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Worksheet.Name = TimezoneWorksheetName
        RenameWorksheetModule ThisWorkbook, Worksheet.Index, TimezoneWorksheetName
    Else
        Set Worksheet = ThisWorkbook.Worksheets(Index)
    End If
    
    Set WorksheetData = Worksheet

WorksheetData_Exit:
    Exit Function

WorksheetData_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WxslData.WorksheetData."
    Resume WorksheetData_Exit

End Function

' Fill the timezone tables from the Windows Registry.
' They will be ordered by their bias and list of localised locations.
' Returns True, if the tables were successfully filled.
'
' If the worksheet does not exist, it will be created.
' If the tables don't exist, they will be create of no other tables are present.
'
' 2020-03-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ReloadTimezones() As Boolean
    
    Dim Worksheet       As Excel.Worksheet
    Dim TimezoneList    As Excel.ListObject
    Dim LocationList    As Excel.ListObject
    Dim ListColumn      As Excel.ListColumn
    Dim Range           As Excel.Range
    
    Dim Entries()       As TimezoneEntry
    Dim Items()         As String
    
    Dim Entry           As TimezoneEntry
    Dim Index           As Integer
    Dim SubIndex        As Integer
    Dim Id              As Integer
    Dim Result          As Boolean
    
    On Error GoTo ReloadTimezones_Error

    Entries = RegistryTimezoneItems()
    SortEntriesBiasLocations Entries
    
    Set Worksheet = WorksheetData
    If Not Worksheet Is Nothing Then
        If Worksheet.ListObjects.Count = 0 Then
            ' Worksheet exists but has no tables.
            ' Create the tables.
            CreateTimezoneData
        Else
            ' Do not overwrite existing tables.
        End If
        
        If IsListObject(Worksheet, TimezoneTableZone) And IsListObject(Worksheet, TimezoneTableLocation) Then
            ' Tables exist in the worksheet.
            ' Clear their content if they priviously have been filled.
            Set LocationList = Worksheet.ListObjects(TimezoneTableLocation)
            If Not LocationList.DataBodyRange Is Nothing Then
                LocationList.DataBodyRange.Delete
            End If
            Set TimezoneList = Worksheet.ListObjects(TimezoneTableZone)
            If Not TimezoneList.DataBodyRange Is Nothing Then
                TimezoneList.DataBodyRange.Delete
            End If
            
            ' Fill the tables row by row..
            For Index = LBound(Entries) To UBound(Entries)
                ' Add and fill a row for a timezone.
                Set Range = TimezoneList.ListRows.Add().Range
                Entry = Entries(Index)
                Range(1, 1) = Entry.Mui
                Range(1, 2) = Entry.MuiDaylight
                Range(1, 3) = Entry.MuiStandard
                Range(1, 4) = Entry.Name
                Range(1, 5) = Entry.Bias
                Range(1, 6) = Entry.Utc
                Range(1, 7) = Entry.Locations
                Range(1, 8) = Entry.ZoneDaylight
                Range(1, 9) = Entry.ZoneStandard
                Range(1, 10) = Entry.FirstEntry
                Range(1, 11) = Entry.LastEntry
                ' The formatted column for display and validation.
                Range(1, 12) = FormatBias(Entry.Bias, True, True, Entry.Name) & " " & Entry.Locations
            
                ' Add and fill rows for the locations of the timezone.
                Items = Split(Entry.Locations, ",")
                For SubIndex = LBound(Items) To UBound(Items)
                    If Trim(Items(SubIndex)) <> "" Then
                        Set Range = LocationList.ListRows.Add().Range
                        Id = Id + 1
                        Range(1, 1) = Id
                        Range(1, 2) = Entry.Mui
                        Range(1, 3) = Trim(Items(SubIndex))
                    End If
                Next
            Next
            Result = True
            
            ' Adjust column widths for both tables to fit the content.
            For Index = 1 To TimezoneList.ListColumns.Count
                Set ListColumn = TimezoneList.ListColumns(Index)
                ListColumn.Range.EntireColumn.AutoFit
            Next
            For Index = 1 To LocationList.ListColumns.Count
                Set ListColumn = LocationList.ListColumns(Index)
                ListColumn.Range.EntireColumn.AutoFit
            Next
        Else
            ' No tables (ListObjects) to fill.
        End If
    End If

    ReloadTimezones = Result
    
ReloadTimezones_Exit:
    Exit Function

ReloadTimezones_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WxlsData.ReloadTimezones."
    Resume ReloadTimezones_Exit

End Function
