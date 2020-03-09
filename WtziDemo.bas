Attribute VB_Name = "WtziDemo"
Option Compare Text
Option Explicit

' WtziDemo v1.0.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Functions for test or demonstration of the function for Windows timezones.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

' Required modules:
'   ByteUtil
'   DateUtil (or project VBA.Date)
'   WtziBase
'   WtziCore


' List one timezone from the Windows Registry by its name.
' If no name is passed, information for timezone UTC is read.
' If a non-existing name is passed, no information is listed.
'
' Example:
'   ? ListOneTimezone("central europe standard time")
'   ' or:
'   ? ListOneTimezone("central europe")
'   Name:         Central Europe Standard Time
'   UTC Zone:     UTC+01:00
'   Bias:         -60
'   MUI:          -280
'   MUI Std:      -282
'   MUI Dlt:      -281
'   Std:          Centraleuropa, normaltid
'   Dlt:          Centraleuropa, sommertid
'   Bias Std:      0
'   Bias Dlt:     -60
'   Date Std:     2019-10-27 03:00:00
'   Date Dlt:     2019-03-31 02:00:00
'   Locations:    Beograd, Bratislava, Budapest, Ljubljana, Prag
'   True
'
' 2018-11-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ListOneTimezone( _
    ByVal Name As String) _
    As Boolean

    Const UtcName   As String = "UTC"
    
    Dim Entries()   As TimezoneEntry
    
    Dim Entry       As TimezoneEntry
    Dim Result      As Boolean
    
    If Name = "" Then
        ' Use default name.
        Name = UtcName
    End If
    
    Entries = RegistryTimezoneItems(Name)
    ' Only one entry is expected.
    Entry = Entries(LBound(Entries))
    
    ' If the name is not found, no name is returned.
    Result = CBool(Len(Entry.Name))
    
    If Result = True Then
        ' List information for the found timezone.
        Debug.Print "Name:", Entry.Name
        Debug.Print "UTC Zone:", Entry.Utc
        Debug.Print "Bias:", Entry.Bias
        Debug.Print "MUI:", Entry.Mui
        Debug.Print "MUI Std:", Entry.MuiStandard
        Debug.Print "MUI Dlt:", Entry.MuiDaylight
        Debug.Print "Std:", Entry.ZoneStandard
        Debug.Print "Dlt:", Entry.ZoneDaylight
        Debug.Print "Bias Std:", Entry.Tzi.StandardBias
        Debug.Print "Bias Dlt:", Entry.Tzi.DaylightBias
        Debug.Print "Date Std:", Format(DateSystemTime(Entry.Tzi.StandardDate), "yyyy-mm-dd hh:nn:ss")
        Debug.Print "Date Dlt:", Format(DateSystemTime(Entry.Tzi.DaylightDate), "yyyy-mm-dd hh:nn:ss")
        Debug.Print "Locations:", Entry.Locations
    End If
    
    ListOneTimezone = Result

End Function

' List all timezones from the Windows Registry.
' They will be ordered by their names (keys).
'
' 2018-11-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub ListAllTimezones()
    
    Dim Entries()   As TimezoneEntry
    
    Dim Entry       As TimezoneEntry
    Dim Index       As Integer
    
    Entries = RegistryTimezoneItems()
    
    For Index = LBound(Entries) To UBound(Entries)
        Entry = Entries(Index)
        Debug.Print "Mui: " & Entry.Mui, "Bias: " & Str(Entry.Bias), "Name: " & Entry.Name
    Next
    
End Sub

' List all timezones from the Windows Registry.
' They will be ordered by their bias and list of localised locations.
'
' 2018-11-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub ListAllTimezonesSorted()
    
    Dim Entries()   As TimezoneEntry
    
    Dim Entry       As TimezoneEntry
    Dim Index       As Integer
    
    Entries = RegistryTimezoneItems()
    SortEntriesBiasLocations Entries
    
    For Index = LBound(Entries) To UBound(Entries)
        Entry = Entries(Index)
        Debug.Print "Mui: " & Entry.Mui, "Bias: " & Str(Entry.Bias), "Locations: " & Entry.Locations '.Name
    Next
    
End Sub

' List the detail of a Windows timezone.
'
' Example:
'   ListTimezone TimezoneLocation("Osaka")
'   Name:         Tokyo Standard Time
'   UTC Zone:     UTC+09:00
'   Bias:         -540
'   MUI:          -630
'   MUI Std:      -632
'   MUI Dlt:      -631
'   Std:          Tokyo, normaltid
'   Dlt:          Tokyo, sommertid
'   Bias Std:      0
'   Bias Dlt:     -60
'   Date Std:     2018-12-02 00:00:00
'   Date Dlt:     2018-12-02 00:00:00
'   Locations:    Osaka, Sapporo, Tokyo
'
' 2018-11-11. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub ListTimezone(ByRef Entry As TimezoneEntry)

    ' List information for the timezone.
    Debug.Print "Name:", Entry.Name
    Debug.Print "UTC Zone:", Entry.Utc
    Debug.Print "Bias:", Entry.Bias
    Debug.Print "MUI:", Entry.Mui
    Debug.Print "MUI Std:", Entry.MuiStandard
    Debug.Print "MUI Dlt:", Entry.MuiDaylight
    Debug.Print "Std:", Entry.ZoneStandard
    Debug.Print "Dlt:", Entry.ZoneDaylight
    Debug.Print "Bias Std:", Entry.Tzi.StandardBias
    Debug.Print "Bias Dlt:", Entry.Tzi.DaylightBias
    Debug.Print "Date Std:", Format(DateSystemTime(Entry.Tzi.StandardDate), "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Date Dlt:", Format(DateSystemTime(Entry.Tzi.DaylightDate), "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Locations:", Entry.Locations

End Sub

' Retrieve the bias of the current timezone.
'
' 2018-11-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CurrentTimezoneBias() As Integer

    Dim Bias        As Integer
    
    Bias = TimezoneCurrent.Bias
    
    CurrentTimezoneBias = Bias
    
End Function

' Retrieve the name of the current timezone.
'
' 2018-11-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CurrentTimezoneName() As String

    Dim Name        As String
    
    Name = TimezoneCurrent.Name
    
    CurrentTimezoneName = Name
    
End Function

' Retrieve the locations of the current timezone.
'
' 2018-11-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function CurrentTimezoneLocations() As String

    Dim Locations   As String
    
    Locations = TimezoneCurrent.Locations
    
    CurrentTimezoneLocations = Locations
    
End Function

