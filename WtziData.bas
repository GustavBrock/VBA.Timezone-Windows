Attribute VB_Name = "WtziData"
Option Compare Database
Option Explicit

' WtziData v1.0.3
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Functions to maintian tables holding Windows timezone informations
' retrieved from the Registry.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

' Required modules:
'   ByteUtil
'   DataUtil
'   DateUtil (or project VBA.Date)
'   WtziBase
'   WtziCore

' Required references:
'   Windows Script Host Object Model


' Timezone table names.
Private Const TimezoneTableZone     As String = "WindowsTimezone"
Private Const TimezoneTableLocation As String = "WindowsTimezoneLocation"
Private Const TimezoneTableRelation As String = TimezoneTableZone & "_" & TimezoneTableLocation

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
Private Const TimezoneLocationId    As String = "Id"
Private Const TimezoneLocationMui   As String = "MUI"
Private Const TimezoneLocationName  As String = "Name"

' Creates (if missing) the supporting timezone tables.
' Returns True if success, False if not.
'
' 2018-11-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CreateTimezoneData() As Boolean

    Dim Result  As Boolean
    
    ' Create the timezone tables if missing.
    Result = CreateTimezoneDataTable(TimezoneTableZone)
    Result = Result And CreateTimezoneDataTable(TimezoneTableLocation)
    
    If Result = True Then
        ' Enforce referential integrity on the timezone tables.
        Result = CreateTimezoneDataTableRelations()
    End If
    
    CreateTimezoneData = Result

End Function

' Creates a timezone table and its indexes from scratch if missing.
' Returns True if success, False if not.
'
' 2018-11-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CreateTimezoneDataTable( _
    ByVal TableName As String) _
    As Boolean
    
    Const PrimaryKeyName    As String = "PrimaryKey"
    
    Dim Database            As DAO.Database
    Dim Table               As DAO.TableDef
    Dim Field               As DAO.Field
    Dim Index               As DAO.Index
    
    Dim Result              As Boolean
    
    Set Database = CurrentDb
    
    If IsTableDefName(TableName) Then
        Result = True
    Else
        ' Create table.
        Select Case TableName
            Case TimezoneTableZone
                Set Table = Database.CreateTableDef(TableName)
                    Set Field = Table.CreateField(TimezoneMui, dbInteger)
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneMuiDaylight, dbInteger)
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneMuiStandard, dbInteger)
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneName, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneBias, dbInteger)
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneUtc, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneLocations, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneDlt, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneStd, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneFirstEntry, dbInteger)
                    Field.Required = False
                    Field.DefaultValue = "Null"
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneLastEntry, dbInteger)
                    Field.Required = False
                    Field.DefaultValue = "Null"
                Table.Fields.Append Field
                
                    Set Index = Table.CreateIndex(PrimaryKeyName)
                        Set Field = Index.CreateField(TimezoneMui)
                        Index.Fields.Append Field
                        Index.Primary = True
                Table.Indexes.Append Index
                    Set Index = Table.CreateIndex(TimezoneMuiDaylight)
                        Set Field = Index.CreateField(TimezoneMuiDaylight)
                        Index.Fields.Append Field
                        Index.Unique = True
                        Index.Primary = False
                Table.Indexes.Append Index
                    Set Index = Table.CreateIndex(TimezoneMuiStandard)
                        Set Field = Index.CreateField(TimezoneMuiStandard)
                        Index.Fields.Append Field
                        Index.Unique = True
                        Index.Primary = False
                Table.Indexes.Append Index
                    Set Index = Table.CreateIndex(TimezoneName)
                        Set Field = Index.CreateField(TimezoneName)
                        Index.Fields.Append Field
                        Index.Unique = True
                        Index.Primary = False
                Table.Indexes.Append Index
                
            Case TimezoneTableLocation
                Set Table = Database.CreateTableDef(TableName)
                    Set Field = Table.CreateField(TimezoneLocationId, dbLong)
                    Field.Required = True
                    Field.Attributes = Field.Attributes Or dbAutoIncrField
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneLocationMui, dbInteger)
                    Field.Required = True
                Table.Fields.Append Field
                    Set Field = Table.CreateField(TimezoneLocationName, dbText, 50)
                    Field.AllowZeroLength = False
                    Field.Required = True
                Table.Fields.Append Field
                
                ' Don't create an index on MUI as this will
                ' be created when creating referential integrity.
                    Set Index = Table.CreateIndex(PrimaryKeyName)
                        Set Field = Index.CreateField(TimezoneLocationId)
                        Index.Fields.Append Field
                        Index.Primary = True
                Table.Indexes.Append Index
                    Set Index = Table.CreateIndex(TimezoneLocationName)
                        Set Field = Index.CreateField(TimezoneLocationName)
                        Index.Fields.Append Field
                Table.Indexes.Append Index
                
        End Select
        
        If Not Table Is Nothing Then
            ' Append table.
            Database.TableDefs.Append Table
            Result = True
        End If
    End If
    
    CreateTimezoneDataTable = Result
    
End Function

' Creates and appends missing relations between the timezone tables.
' Note, that this will create a hidden index on the foreign table field.
' Returns True if success, False if not, typically because the tables are missing.
'
' 2018-11-01. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CreateTimezoneDataTableRelations() As Boolean

    Dim Database        As DAO.Database
    Dim Field           As DAO.Field
    Dim Relation        As DAO.Relation
    Dim Table           As DAO.TableDef
    Dim ForeignTable    As DAO.TableDef
    
    Dim Name            As String
    Dim ForeignName     As String
    Dim Result          As Boolean
    
    Set Database = CurrentDb
        
    If IsRelationName(TimezoneTableRelation) Then
        Result = True
    ElseIf IsTableDefName(TimezoneTableZone) And IsTableDefName(TimezoneTableLocation) Then
        Set Table = Database.TableDefs(TimezoneTableZone)
        Set ForeignTable = Database.TableDefs(TimezoneTableLocation)
        
        ' Create and append relation RelationName using these fields:
        Name = TimezoneMui
        ForeignName = TimezoneLocationMui
            
            Set Relation = Database.CreateRelation(TimezoneTableRelation)
            Relation.Table = Table.Name
            Relation.ForeignTable = ForeignTable.Name
            Relation.Attributes = dbRelationUpdateCascade
            
                Set Field = Relation.CreateField(Name)
                Field.ForeignName = ForeignName
            Relation.Fields.Append Field
        Database.Relations.Append Relation
    
        Set ForeignTable = Nothing
        Set Table = Nothing
        Result = True
    End If
    
    CreateTimezoneDataTableRelations = Result
    
End Function

' Updates the local timezone tables with the current timezones of Windows.
' If Force is True, the tables will be created if they don't exist.
' Returns True if the tables were created or updated successfully.
'
' 2018-11-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ReloadWindowsTimezoneTables( _
    Optional ByVal Force As Boolean) _
    As Boolean

    Dim Timezones   As DAO.Recordset
    Dim Locations   As DAO.Recordset
    
    Dim Entries()   As TimezoneEntry
    Dim Items()     As String
    
    Dim Index       As Integer
    Dim SubIndex    As Integer
    Dim Sql         As String
    Dim Success     As Boolean
    
    If Force = True Then
        ' Create the timezone tables if they don't exist.
        Success = CreateTimezoneData()
    Else
        ' Check if the timezone tables exist.
        Success = IsRelationName(TimezoneTableZone)
    End If
    
    If Success Then
        Entries = RegistryTimezoneItems()
        
        Sql = "Delete * From " & TimezoneTableLocation & ""
        CurrentDb.Execute Sql
        Sql = "Delete * From " & TimezoneTableZone & ""
        CurrentDb.Execute Sql
        
        Sql = "Select * From " & TimezoneTableZone & ""
        Set Timezones = CurrentDb.OpenRecordset(Sql)
        Sql = "Select * From " & TimezoneTableLocation & ""
        Set Locations = CurrentDb.OpenRecordset(Sql)
        
        For Index = LBound(Entries) To UBound(Entries)
            Timezones.AddNew
                Timezones.Fields(TimezoneMui).Value = Entries(Index).Mui
                Timezones.Fields(TimezoneMuiDaylight).Value = Entries(Index).MuiDaylight
                Timezones.Fields(TimezoneMuiStandard).Value = Entries(Index).MuiStandard
                Timezones.Fields(TimezoneName).Value = Entries(Index).Name
                Timezones.Fields(TimezoneBias).Value = Entries(Index).Bias
                Timezones.Fields(TimezoneUtc).Value = Entries(Index).Utc
                Timezones.Fields(TimezoneLocations).Value = Entries(Index).Locations
                Timezones.Fields(TimezoneDlt).Value = Entries(Index).ZoneDaylight
                Timezones.Fields(TimezoneStd).Value = Entries(Index).ZoneStandard
                Timezones.Fields(TimezoneFirstEntry).Value = Entries(Index).FirstEntry
                Timezones.Fields(TimezoneLastEntry).Value = Entries(Index).LastEntry
            Timezones.Update
            
            Items = Split(Entries(Index).Locations, ",")
            For SubIndex = LBound(Items) To UBound(Items)
                If Trim(Items(SubIndex)) <> "" Then
                    Locations.AddNew
                        Locations.Fields(TimezoneLocationMui).Value = Entries(Index).Mui
                        Locations.Fields(TimezoneLocationName).Value = Trim(Items(SubIndex))
                    Locations.Update
                End If
            Next
        Next
        Locations.Close
        Timezones.Close
        Success = True
    End If
    
    ReloadWindowsTimezoneTables = Success

End Function


