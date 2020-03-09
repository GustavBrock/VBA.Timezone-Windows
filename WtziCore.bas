Attribute VB_Name = "WtziCore"
Option Compare Text
Option Explicit

' WtziCore v1.0.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Functions for converting integers to hex bytes (octets).
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

' Required modules:
'   ByteUtil
'   DateUtil (or modules from project VBA.Date)
'   WtziBase

' Required references:
'   Windows Script Host Object Model


' Bias.
Private Const UtcBias           As Long = 0

Private Const HKeyLocalMachine  As Long = &H80000002

' Common extension of timezone name. Note the leading space.
Private Const StandardTimeLabel As String = " Standard Time"

' Calculates the date/time of LocalDate in a remote timezone.
' Adds the difference in minutes between the local timezone bias and
' the remote timezone bias, where both bias values are relative to UTC.
'
' Examples:
'
'   RemoteDate = DateRemoteBias(Now(), 60, -600)
'   will return RemoteDate as eleven hours ahead of local time.
'
'   RemoteDate = DateRemoteBias(Now(), -600, 60)
'   will return RemoteDate as eleven hours behind local time.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateRemoteBias( _
    ByVal LocalDate As Date, _
    ByVal LocalBias As Long, _
    ByVal RemoteBias As Long) _
    As Date
  
    Dim RemoteDate  As Date
    Dim Bias        As Long
    
    ' Find difference (in minutes) between timezone bias.
    Bias = BiasDiff(LocalBias, RemoteBias)
    ' Calculate remote date/time.
    RemoteDate = DateAdd("n", Bias, LocalDate)
    
    DateRemoteBias = RemoteDate
  
End Function

' Calculates the date/time of RemoteDate in a local timezone.
' Adds the difference in minutes between the remote timezone bias and
' the local timezone bias, where both bias values are relative to UTC.
'
' Examples:
'
'   LocalDate = DateLocalBias(Now(), 60, -600)
'   will return LocalDate as eleven hours ahead of remote time.
'
'   LocalDate = DateLocalBias(Now(), -600, 60)
'   will return LocalDate as eleven hours behind remote time.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateLocalBias( _
    ByVal RemoteDate As Date, _
    ByVal RemoteBias As Long, _
    ByVal LocalBias As Long) _
    As Date
  
    Dim LocalDate   As Date
    Dim Bias        As Long
    
    ' Find difference (in minutes) between timezone bias.
    Bias = BiasDiff(RemoteBias, LocalBias)
    ' Calculate local date/time.
    LocalDate = DateAdd("n", Bias, RemoteDate)
    
    DateLocalBias = LocalDate
  
End Function

' Calculates the difference in bias (minutes) between two timezones,
' typically from the local timezone to the remote timezone.
' Both timezones must be expressed by their bias relative to
' UTC (Coordinated Universal Time) which is measured in minutes.
'
' 2019-11-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BiasDiff( _
    ByVal LocalBias As Long, _
    ByVal RemoteBias As Long) _
    As Long
  
    Dim Bias    As Long
    
    ' Calculate difference in timezone bias.
    Bias = LocalBias - RemoteBias
    
    BiasDiff = Bias

End Function

' Returns the timezone bias as specified in Windows from
' the name (key) of a timezone entry in the Registry.
' Accepts values without the common trailing "Standard Time".
'
' If Dst is true, and the current date is within daylight saving time,
' bias for daylight saving time is returned.
' If Date1 is specified, the bias of that date is returned.
'
' Returns a bias of zero if a timezone is not found.
'
' Examples:
'   Bias = BiasWindowsTimezone("Argentina")
'   Bias -> 180     ' Found
'
'   Bias = BiasWindowsTimezone("Argentina Standard Time")
'   Bias -> 180     ' Found.
'
'   Bias = BiasWindowsTimezone("Germany")
'   Bias -> 0       ' Not found.
'
'   Bias = BiasWindowsTimezone("Western Europe")
'   Bias -> 0       ' Not found.
'
'   Bias = BiasWindowsTimezone("W. Europe")
'   Bias -> -60     ' Found.
'
'   Bias = BiasWindowsTimezone("Paraguay", True, #2018-07-07#)
'   Bias -> 240     ' Found.
'
'   Bias = BiasWindowsTimezone("Paraguay", True, #2018-02-11#)
'   Bias -> 180     ' Found. DST.
'
'   Bias = BiasWindowsTimezone("Paraguay", False, #2018-02-11#)
'   Bias -> 240     ' Found.
'
' 2018-11-16. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function BiasWindowsTimezone( _
    ByVal TimezoneName As String, _
    Optional Dst As Boolean, _
    Optional Date1 As Date) _
    As Long
    
    Static Entries()    As TimezoneEntry
    Static LastName     As String
    Static LastYear     As Integer
    
    Static Entry        As TimezoneEntry
    Dim ThisName        As String
    Dim ThisYear        As Integer
    Dim StandardDate    As Date
    Dim DaylightDate    As Date
    Dim DeltaBias       As Long
    Dim Bias            As Long
    
    If Trim(TimezoneName) = "" Then
        ' Nothing to look up.
        Exit Function
    Else
        ThisName = Trim(TimezoneName)
        ThisYear = Year(Date1)
        If LastName = ThisName And LastYear = ThisYear Then
            ' Use cached data.
        Else
            ' Retrieve the single entry or - if not found - an empty entry.
            Entries = RegistryTimezoneItems(ThisName, (ThisYear))
            Entry = Entries(LBound(Entries))
            LastName = ThisName
            LastYear = ThisYear
        End If
        If _
            StrComp(Entry.Name, TimezoneName, vbTextCompare) = 0 Or _
            StrComp(Replace(Entry.Name, StandardTimeLabel, ""), TimezoneName, vbTextCompare) = 0 Then
            ' Windows timezone found.
            
            ' Default is standard bias.
            DeltaBias = Entry.Tzi.StandardBias
            If Dst = True Then
                ' Return daylight bias if Date1 is of daylight saving time.
                StandardDate = DateSystemTime(Entry.Tzi.StandardDate)
                DaylightDate = DateSystemTime(Entry.Tzi.DaylightDate)
                
                If DaylightDate < StandardDate Then
                    ' Northern Hemisphere.
                    If DateDiff("s", DaylightDate, Date1) >= 0 And DateDiff("s", Date1, StandardDate) > 0 Then
                        ' Daylight time.
                        DeltaBias = Entry.Tzi.DaylightBias
                    Else
                        ' Standard time.
                    End If
                Else
                    ' Southern Hemisphere.
                    If DateDiff("s", DaylightDate, Date1) >= 0 Or DateDiff("s", Date1, StandardDate) > 0 Then
                        ' Daylight time.
                        DeltaBias = Entry.Tzi.DaylightBias
                    Else
                        ' Standard time.
                    End If
                End If
                
            End If
            ' Calculate total bias.
            Bias = Entry.Bias + DeltaBias
        End If
    End If

    BiasWindowsTimezone = Bias

End Function

' Low-level function to retrieve all timezone data from Windows
' for the current year as an array of timezone entries.
' Optionally, retrieves the (dynamic) data for another year.
'
' To be called from RegistryTimezoneItems.
'
' Required references:
'   Windows Script Host Object Model
'
' 2019-12-14. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function GetRegistryTimezoneItems( _
    Optional ByRef DynamicDstYear As Integer) _
    As TimezoneEntry()

    Const Component     As String = "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv"
    Const DefKey        As Long = HKeyLocalMachine
    Const HKey          As String = "HKLM"
    Const SubKeyPath    As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
    Const DstPath       As String = "Dynamic DST"

    Const DisplayKey    As String = "Display"
    Const DaylightKey   As String = "Dlt"
    Const StandardKey   As String = "Std"
    Const MuiDisplayKey As String = "MUI_Display"
    Const MuiDltKey     As String = "MUI_Dlt"
    Const MuiStdKey     As String = "MUI_Std"
    Const TziKey        As String = "TZI"
    Const FirstEntryKey As String = "FirstEntry"
    Const LastEntryKey  As String = "LastEntry"
    
    Dim SWbemServices   As Object
    Dim WshShell        As WshShell
   
    Dim SubKey          As Variant
    Dim Names           As Variant
    Dim NameKeys        As Variant
    
    Dim Display         As String
    Dim DisplayUtc      As String
    Dim Name            As Variant
    Dim DstEntry        As Variant
    Dim Mui             As Integer
    Dim BiasLabel       As String
    Dim Bias            As Long
    Dim Locations       As String
    Dim TziDetails      As Variant
    Dim TzItems()       As TimezoneEntry
    Dim TzItem          As TimezoneEntry
    Dim Index           As Long
    Dim SubIndex        As Long
    Dim Value           As String
    Dim LBoundItems     As Long
    Dim UBoundItems     As Long
    
    Dim TziInformation  As RegTziFormat

    ' The call is either for another year, or
    ' more than one day has passed since the last call.
    Set SWbemServices = GetObject(Component)
    Set WshShell = New WshShell

    SWbemServices.EnumKey DefKey, SubKeyPath, Names
    ' Retrieve all timezones' base data.
    LBoundItems = LBound(Names)
    UBoundItems = UBound(Names)
    ReDim TzItems(LBoundItems To UBoundItems)
    
    For Index = LBound(Names) To UBound(Names)
        ' Assemble paths and look up key values.
        SubKey = Names(Index)
        
        ' Invariant name of timezone.
        TzItem.Name = SubKey
        
        ' MUI of the timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, MuiDisplayKey), "\")
        Value = WshShell.RegRead(Name)
        Mui = Val(Split(Value, ",")(1))
        TzItem.Mui = Mui
        ' MUI of the standard timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, MuiStdKey), "\")
        Value = WshShell.RegRead(Name)
        Mui = Val(Split(Value, ",")(1))
        TzItem.MuiStandard = Mui
        ' MUI of the DST timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, MuiDltKey), "\")
        Value = WshShell.RegRead(Name)
        Mui = Val(Split(Value, ",")(1))
        TzItem.MuiDaylight = Mui
        
        ' Localised description of the timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, DisplayKey), "\")
        Display = WshShell.RegRead(Name)
        ' Extract the first part, cleaned like "UTC+08:30".
        DisplayUtc = Mid(Split(Display, ")", 2)(0) & "+00:00", 2, 9)
        ' Extract the offset part of first part, like "+08:30".
        BiasLabel = Mid(Split(Display, ")", 2)(0) & "+00:00", 5, 6)
        ' Convert the offset part of the first part to a bias value (signed integer minutes).
        Bias = -Val(Left(BiasLabel, 1) & Str(CDbl(CDate(Mid(BiasLabel, 2))) * 24 * 60))
        ' Extract the last part, holding the location(s).
        Locations = Split(Display, " ", 2)(1)
        TzItem.Bias = Bias
        TzItem.Utc = DisplayUtc
        TzItem.Locations = Locations
        
        ' Localised name of the standard timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, StandardKey), "\")
        TzItem.ZoneStandard = WshShell.RegRead(Name)
        ' Localised name of the DST timezone.
        Name = Join(Array(HKey, SubKeyPath, SubKey, DaylightKey), "\")
        TzItem.ZoneDaylight = WshShell.RegRead(Name)
        
        ' TZI details.
        SWbemServices.GetBinaryValue DefKey, Join(Array(SubKeyPath, SubKey), "\"), TziKey, TziDetails
        FillRegTziFormat TziDetails, TziInformation
        TzItem.Tzi = TziInformation
        ' Default Dynamic DST range.
        TzItem.FirstEntry = Null
        TzItem.LastEntry = Null
        
        ' Check for Dynamic DST info.
        SWbemServices.EnumKey DefKey, Join(Array(SubKeyPath, SubKey), "\"), NameKeys
        If IsArray(NameKeys) Then
            ' This timezone has subkeys. Check if Dynamic DST is present.
            For SubIndex = LBound(NameKeys) To UBound(NameKeys)
                If NameKeys(SubIndex) = DstPath Then
                    ' Dynamic DST details found.
                    ' Record first and last entry.
                    DstEntry = Join(Array(HKey, SubKeyPath, SubKey, DstPath, FirstEntryKey), "\")
                    Value = WshShell.RegRead(DstEntry)
                    TzItem.FirstEntry = Value
                    DstEntry = Join(Array(HKey, SubKeyPath, SubKey, DstPath, LastEntryKey), "\")
                    Value = WshShell.RegRead(DstEntry)
                    TzItem.LastEntry = Value
                    
                    If DynamicDstYear >= TzItems(Index).FirstEntry And _
                        DynamicDstYear <= TzItems(Index).LastEntry Then
                        ' Replace default TZI details with those from the dynamic DST.
                        DstEntry = Join(Array(SubKeyPath, SubKey, DstPath), "\")
                        SWbemServices.GetBinaryValue DefKey, Join(Array(SubKeyPath, SubKey), "\"), , CStr(DynamicDstYear), TziDetails
                        FillRegTziFormat TziDetails, TziInformation
                        TzItem.Tzi = TziInformation
                    Else
                        ' Dynamic DST year was not found.
                        ' Return current year.
                        DynamicDstYear = Year(Date)
                    End If
                    Exit For
                End If
            Next
        End If
        TzItems(Index) = TzItem
    Next
    
    GetRegistryTimezoneItems = TzItems
    
End Function

' Retrieves all timezone data from Windows for the current year as
' an array of timezone entries.
' Optionally, retrieves the data for one timezone only or the
' (dynamic) data for another year.
'
' Will cache the retrieved data to speed up repeated calls.
'
' 2019-12-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RegistryTimezoneItems( _
    Optional ByVal TimezoneName As String, _
    Optional ByRef DynamicDstYear As Integer) _
    As TimezoneEntry()

    Static TzAllItems() As TimezoneEntry
    Static LastCall     As Date
    Static LastYear     As Integer
       
    Dim TzItems()       As TimezoneEntry
    Dim TzItem          As TimezoneEntry
    Dim Index           As Long
    Dim Continue        As Boolean
    
    If LastYear <> DynamicDstYear Or DateDiff("d", LastCall, Date) <> 0 Then
        ' Save 0.5 second for each call by caching the retrieved timezones.
        TzAllItems() = GetRegistryTimezoneItems(DynamicDstYear)
        LastYear = DynamicDstYear
        LastCall = Date
    End If

    If TimezoneName = "" Then
        ' Retrieve all timezones' base data.
        TzItems() = TzAllItems()
    Else
        ' Retrieve one timezone's base data.
        ReDim TzItems(0)
    
        For Index = LBound(TzAllItems) To UBound(TzAllItems)
            TzItem = TzAllItems(Index)
            Continue = Not CBool(StrComp(TzItem.Name, TimezoneName, vbTextCompare))
            If Continue = False Then
                ' Check, if stripping the trailing "Standard Time" form Name will result in a match.
                Continue = Not CBool(StrComp(Replace(TzItem.Name, StandardTimeLabel, ""), TimezoneName, vbTextCompare))
            End If
            
            If Continue = True Then
                TzItems(0) = TzItem
                Exit For
            End If
        Next
    End If
    
    RegistryTimezoneItems = TzItems
    
End Function

' Converts a SystemTime structure to its date/time value.
'
' If SysTime.wYear is zero, SysTime is expected to hold the special set of data
' used for calculation of the beginning or the end of the Daylight Saving Time
' period of the current year.
'
' A value for SystemTime.wMilliseconds will be ignored.
'
' 2018-06-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateSystemTime( _
    ByRef SysTime As SystemTime) _
    As Date
    
    Const FirstDay  As Integer = 1

    Dim DateInMonth As Date
    Dim Occurrence  As Integer
    Dim Weekday     As VBA.VbDayOfWeek
    Dim DateValue   As Date
    Dim TimeValue   As Date
    Dim Value       As Date

    With SysTime
        If .wYear <> 0 Then
            ' Calculate actual date from structure data.
            DateValue = DateSerial(.wYear, .wMonth, .wDay)
        Else
            ' Calculate from a set of data for Daylight Saving Time.
            DateInMonth = DateSerial(Year(Date), .wMonth, FirstDay)
            Occurrence = .wDay
            Weekday = WeekdayFromDayOfWeek(.wDayOfWeek)
            DateValue = DateWeekdayInMonth(DateInMonth, Occurrence, Weekday)
        End If
        TimeValue = TimeSerial(.wHour, .wMinute, .wSecond)
        
        If CDbl(DateValue) >= 0 Then
            Value = DateValue + TimeValue
        Else
            Value = DateValue - TimeValue
        End If
    End With
  
    DateSystemTime = Value

End Function

' Converts a TziDetails variable to a RegTziFormat variable.
' Returns this RegTziFormat value by reference.
'
' TziDetails must be the Tzi element from a TimezoneEntry structure.
' The Tzi element is a Byte array.
'
' 2018-11-11. Gustav Brock. Cactus Data ApS, CPH.
'
Private Sub FillRegTziFormat( _
    ByVal TziDetails As Variant, _
    ByRef RegTziInformation As RegTziFormat)
    
    Dim TimeInfo    As SystemTime
    Dim Item        As Integer
    Dim Bytes(3)    As Byte
    Dim Offset      As Integer
    
    For Item = LBound(TziDetails) To UBound(TziDetails)
        Select Case Item
            ' Bias.
            Case 0 To 3
                Offset = 0
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 3 Then
                    RegTziInformation.Bias = CLngBytes(Bytes)
                End If
                
            ' Standard bias.
            Case 4 To 7
                Offset = 4
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 3 + Offset Then
                    RegTziInformation.StandardBias = CLngBytes(Bytes)
                End If

            ' Daylight bias.
            Case 8 To 11
                Offset = 8
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 3 + Offset Then
                    RegTziInformation.DaylightBias = CLngBytes(Bytes)
                End If
            
            ' Standard time info.
            Case 12, 13
                Offset = 12
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wYear = CIntBytes(Bytes)
                End If
            Case 14, 15
                Offset = 14
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMonth = CIntBytes(Bytes)
                End If
            Case 16, 17
                Offset = 16
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wDayOfWeek = CIntBytes(Bytes)
                End If
            Case 18, 19
                Offset = 18
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wDay = CIntBytes(Bytes)
                End If
            Case 20, 21
                Offset = 20
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wHour = CIntBytes(Bytes)
                End If
            Case 22, 23
                Offset = 22
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMinute = CIntBytes(Bytes)
                End If
            Case 24, 25
                Offset = 24
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wSecond = CIntBytes(Bytes)
                End If
            Case 26, 27
                Offset = 26
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMilliseconds = CIntBytes(Bytes)
                End If
                RegTziInformation.StandardDate = TimeInfo
            
            ' Daylight time info.
            Case 28, 29
                Offset = 28
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wYear = CIntBytes(Bytes)
                End If
            Case 30, 31
                Offset = 30
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMonth = CIntBytes(Bytes)
                End If
            Case 32, 33
                Offset = 32
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wDayOfWeek = CIntBytes(Bytes)
                End If
            Case 34, 35
                Offset = 34
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wDay = CIntBytes(Bytes)
                End If
            Case 36, 37
                Offset = 36
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wHour = CIntBytes(Bytes)
                End If
            Case 38, 39
                Offset = 38
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMinute = CIntBytes(Bytes)
                End If
            Case 40, 41
                Offset = 40
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wSecond = CIntBytes(Bytes)
                End If
            Case 42, 43
                Offset = 42
                Bytes(Item - Offset) = TziDetails(Item)
                If Item = 1 + Offset Then
                    TimeInfo.wMilliseconds = CIntBytes(Bytes)
                End If
                RegTziInformation.DaylightDate = TimeInfo
        End Select
    Next
    
End Sub

' Converts a value of wDayOfWeek from a SystemTime structure to a VBA weekday value.
' An error is raised for invalid input.
'
' Example where Wednesday is the first day of the week:
'   DayOfWeek     Weekday
'   0 Sunday      5
'   1             6
'   2             7
'   3             1 Wednesday
'   4             2
'   5             3
'   6             4
'
' 2019-12-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayFromDayOfWeek( _
    ByVal DayOfWeek As Integer, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As VbDayOfWeek

    Dim Weekday         As VbDayOfWeek
    
    If DayOfWeek < 0 Or DayOfWeek > 6 Then
        ' Raise error.
        Err.Raise DtError.dtInvalidProcedureCallOrArgument, "WeekdayFromDayOfWeek"
        Exit Function
    End If

    Weekday = (DayOfWeek + 1 - FirstDayOfWeek + DaysPerWeek) Mod DaysPerWeek + 1
    
    WeekdayFromDayOfWeek = Weekday

End Function

' Converts a VBA weekday value to the value of wDayOfWeek for a SystemTime structure.
' An error is raised for invalid input.
'
' Example where Wednesday is the first day of the week:
'   Weekday       DayOfWeek
'   1 Wednesday   3
'   2             4
'   3             5
'   4             6
'   5             0 Sunday
'   6             1
'   7             2
'
' 2019-12-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function WeekdayToDayOfWeek( _
    ByVal Weekday As VbDayOfWeek, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Integer
    
    Dim DayOfWeek       As Integer

    ' Validate FirstDayOfWeek.
    Select Case FirstDayOfWeek
        Case _
            vbMonday, _
            vbTuesday, _
            vbWednesday, _
            vbThursday, _
            vbFriday, _
            vbSaturday, _
            vbSunday
        Case Else
            FirstDayOfWeek = vbSunday
    End Select
    
    ' Validate Weekday and calculate DayOfWeek.
    Select Case Weekday
        Case _
            vbMonday, _
            vbTuesday, _
            vbWednesday, _
            vbThursday, _
            vbFriday, _
            vbSaturday, _
            vbSunday
                DayOfWeek = (Weekday + FirstDayOfWeek - 2) Mod DaysPerWeek
        Case Else
            ' Raise error.
            Err.Raise DtError.dtInvalidProcedureCallOrArgument, "WeekdayToDayOfWeek"
            Exit Function
    End Select

    WeekdayToDayOfWeek = DayOfWeek

End Function

' Converts a date/time value to a SystemTime structure.
' Optionally, a value for milliseconds can be passed.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CSystemTime( _
    ByVal Date1 As Date, _
    Optional ByVal Milliseconds As Integer) _
    As SystemTime

    Dim SysTime As SystemTime

    ' Limit milliseconds.
    If Milliseconds < 0 Or Milliseconds > 999 Then
        Milliseconds = 0
    End If
    
    ' Split the date/time value into its components.
    With SysTime
        .wYear = DatePart("yyyy", Date1)
        .wMonth = DatePart("m", Date1)
        .wDay = DatePart("d", Date1)
        .wHour = DatePart("h", Date1)
        .wMinute = DatePart("n", Date1)
        .wSecond = DatePart("s", Date1)
        .wMilliseconds = Milliseconds
        .wDayOfWeek = WeekdayToDayOfWeek(Weekday(Date1))
    End With
  
    CSystemTime = SysTime

End Function

' Converts a present UtcDate from Coordinated Universal Time (UTC) to local time.
' If IgnoreDaylightSetting is True, returned time will always be standard time.
'
' For dates not within the current year, it is preferable to use the function
' ZoneCore.DateFromDistantUtc.
'
' 2017-11-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateFromUtc( _
    ByVal UtcDate As Date, _
    Optional ByVal IgnoreDaylightSetting As Boolean) _
    As Date

    Dim LocalBias   As Long
    Dim LocalDate   As Date
  
    ' Find the local time using standard bias.
    LocalBias = LocalBiasTimezonePresent(UtcDate, True)
    LocalDate = DateRemoteBias(UtcDate, UtcBias, LocalBias)
    If IgnoreDaylightSetting = False Then
        ' The local time should be returned as daylight time.
        If IsCurrentDaylightSavingTime(LocalDate) Then
            ' The local time is daylight time.
            ' Find bias for daylight time.
            LocalBias = LocalBiasTimezonePresent(LocalDate, IgnoreDaylightSetting)
            ' Find the local time using daylight bias.
            LocalDate = DateRemoteBias(UtcDate, UtcBias, LocalBias)
        End If
    End If
    
    DateFromUtc = LocalDate

End Function

' Converts a present LocalDate from local time to Coordinated Universal Time (UTC).
' If IgnoreDaylightSetting is True, LocalDate is considered standard time.
'
' For dates not within the current year, it is preferable to use the function
' ZoneCore.DateToDistantUtc.
'
' Note:
' For a value of the transition interval from daylight time back to standard time,
' where the value could belong to either of these, daylight time is assumed.
' This means that a standard time of the transition interval will return an UTC
' time off by the daylight saving bias.
' Thus, for obtaining the current time in UTC, always use the function UtcNow.
'
' 2017-11-27. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateToUtc( _
    ByVal LocalDate As Date, _
    Optional ByVal IgnoreDaylightSetting As Boolean) _
    As Date

    Dim LocalBias   As Long
    Dim UtcDate     As Date
  
    LocalBias = LocalBiasTimezonePresent(LocalDate, IgnoreDaylightSetting)
    UtcDate = DateRemoteBias(LocalDate, LocalBias, UtcBias)

    DateToUtc = UtcDate

End Function

' Looks up the MUI of a Windows timezone from its name (key).
'
' Examples:
'
'   Mui = MuiWindowsTimezone("Paraguay")
'   Mui -> -960     ' Found.
'
'   Mui = MuiWindowsTimezone("Argentina Standard Time")
'   Bias -> -2080   ' Found.
'
'   Mui = MuiWindowsTimezone("France")
'   Mui -> 0       ' Not found.
'
' 2018-11-16. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function MuiWindowsTimezone( _
    ByVal TimezoneName As String) _
    As Long
    
    Dim Entries()       As TimezoneEntry
    
    Dim Entry           As TimezoneEntry
    Dim Mui             As Long
    
    If Trim(TimezoneName) = "" Then
        ' Nothing to look up.
        Exit Function
    End If
    
    ' Retrieve the single entry or - if not found - an empty entry.
    Entries = RegistryTimezoneItems(TimezoneName)
    Entry = Entries(LBound(Entries))
    If _
        StrComp(Entry.Name, TimezoneName, vbTextCompare) = 0 Or _
        StrComp(Replace(Entry.Name, StandardTimeLabel, ""), TimezoneName, vbTextCompare) = 0 Then
        ' Windows timezone found.
        Mui = Entry.Mui
    End If

    MuiWindowsTimezone = Mui

End Function

' Looks up the Windows timezone from the localised description of the current timezone.
'
' 2019-12-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TimezoneCurrent() As TimezoneEntry

    Dim Entry               As TimezoneEntry

    Entry = TimezoneDescription(LocalTimezoneDescription())
    
    TimezoneCurrent = Entry

End Function

' Looks up a Windows timezone from one of its (localised) description.
' Returns an empty TimezoneEntry if the description is not found.
'
' Examples:
'
'   Entry = TimezoneDescription("Rom, normaltid")
'   Entry.Mui   -> -300
'   Entry.Name  -> "Romance Standard Time"
'
'   Entry = TimezoneDescription("Alaska, Standard Time")
'   Entry.Mui   -> -220
'   Entry.Name  -> "Alaskan Standard Time"
'
'   Entry = TimezoneDescription("Rome, Standard Time")  ' English, not localised.
'   Entry.Mui   -> 0
'   Entry.Name  -> ""
'
' 2019-12-12. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TimezoneDescription( _
    ByVal Description As String) _
    As TimezoneEntry

    Dim Entries()           As TimezoneEntry
    
    Dim Entry               As TimezoneEntry
    Dim EmptyEntry          As TimezoneEntry
    Dim Index               As Integer
    Dim Found               As Boolean
    
    Entries = RegistryTimezoneItems()
    
    For Index = LBound(Entries) To UBound(Entries)
        Entry = Entries(Index)
        If StrComp(Entry.ZoneStandard, Description, vbTextCompare) = 0 Or _
            StrComp(Entry.ZoneDaylight, Description, vbTextCompare) = 0 Then
            Found = True
            Exit For
        End If
    Next
    
    If Not Found Then
        Entry = EmptyEntry
    End If
    
    TimezoneDescription = Entry

End Function

' Looks up a Windows timezone from one of its (localised) locations.
' Returns an empty TimezoneEntry if the location is not found.
'
' Examples:
'
'   Entry = TimezoneLocation("Oslo")
'   Entry.Mui   -> -300
'   Entry.Name  -> "Romance Standard Time"
'
'   Entry = TimezoneLocation("Alaska")
'   Entry.Mui   -> -220
'   Entry.Name  -> "Alaskan Standard Time"
'
'   Entry = TimezoneLocation("Copenhagen")  ' English, not localised.
'   Entry.Mui   -> 0
'   Entry.Name  -> ""
'
' 2019-12-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TimezoneLocation( _
    ByVal Location As String) _
    As TimezoneEntry

    Const LocationDelimiter As String = ", "
    Dim Entries()           As TimezoneEntry
    
    Dim Entry               As TimezoneEntry
    Dim EmptyEntry          As TimezoneEntry
    Dim Index               As Integer
    Dim Locations()         As String
    Dim Item                As Variant
    Dim Found               As Boolean
    
    Entries = RegistryTimezoneItems()
    
    For Index = LBound(Entries) To UBound(Entries)
        Entry = Entries(Index)
        Locations = Split(Entry.Locations, LocationDelimiter)
        For Each Item In Locations
            If StrComp(Item, Location, vbTextCompare) = 0 Then
                Found = True
                Exit For
            End If
        Next
        If Found Then
            Exit For
        End If
    Next
    
    If Not Found Then
        Entry = EmptyEntry
    End If
    
    TimezoneLocation = Entry

End Function

' Looks up a Windows timezone from one of its three Mui keys:
'
'   Mui
'   MuiDaylight
'   MuiStandard
'
' Returns an empty TimezoneEntry if the Mui is not found.
'
' Examples:
'
'   Entry = TimezoneMui(-300")
'   Entry.Name      -> "Romance Standard Time"
'   Entry.Locations -> "København, Stockholm, Oslo, Madrid, Paris"
'
'   Entry = TimezoneMui(-200")
'   Entry.Name      -> "US Mountain Standard Time"
'   Entry.Locations -> "Arizona"
'
'   Entry = TimezoneMui(0)  ' Non-existing Mui.
'   Entry.Name      -> ""
'   Entry.Locations -> ""
'
' 2019-12-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TimezoneMui( _
    ByVal Mui As Integer) _
    As TimezoneEntry

    Dim Entries()           As TimezoneEntry
    
    Dim Entry               As TimezoneEntry
    Dim EmptyEntry          As TimezoneEntry
    Dim Index               As Integer
    Dim Found               As Boolean
    
    Entries = RegistryTimezoneItems()
    
    For Index = LBound(Entries) To UBound(Entries)
        Entry = Entries(Index)
        If Entry.Mui = Mui Or Entry.MuiDaylight = Mui Or Entry.MuiStandard = Mui Then
            Found = True
            Exit For
        End If
    Next
    
    If Not Found Then
        Entry = EmptyEntry
    End If
    
    TimezoneMui = Entry

End Function

' Converts a bias value (minutes) to an offset value of DateTime.
' Note, that a positive bias will return a negative time value
' and vice versa.
'
' The bias must be within 24 hours. If not, an error is raised.
'
' Examples:
'   TimeBias(-60)       ->  #01:00:00#  ' Will display by default as 01:00:00.
'   CDbl(TimeBias(-60)) ->  4.16666666666667E-02
'   TimeBias(360)       -> -#06:00:00#  ' Will display by default as 06:00:00.
'   CDbl(TimeBias(360)) -> -0.25
'   TimeBias(1440)      ->  Error
'
' 2019-12-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function TimeBias( _
    ByVal Bias As Integer) _
    As Date
    
    Dim Result      As Date
    
    If Abs(Bias) >= MinutesPerHour * HoursPerDay Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument, "TimeBias"
    End If
    
    Result = TimeSerial(0, -Bias, 0)
    
    TimeBias = Result
    
End Function

' Formats for display a bias value as a UTC offset.
'
' If argument Name is passed as "UTC", no offet is displayed.
'
' Examples:
'   FormatBias(300)                         -> -05:00
'   FormatBias(300, True)                   -> UTC-05:00
'   FormatBias(300, True, True)             -> (UTC-05:00)
'   FormatBias(0, True)                     -> UTC+00:00
'   FormatBias(0, True, True)               -> (UTC+00:00)
'   FormatBias(0, True, True, "Sao Tome")   -> (UTC+00:00)
'   FormatBias(0, True, True, "UTC")        -> (UTC)
'   FormatBias(0, False, True, "UTC")       -> ()
'   FormatBias(0, False, False, "UTC")      ->
'   FormatBias(-60)                         -> +10:00
'   FormatBias(-60, True)                   -> UTC+10:00
'   FormatBias(-60, True, True)             -> (UTC+10:00)
'
' 2019-12-08. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function FormatBias( _
    ByVal Bias As Integer, _
    Optional PrefixUtc As Boolean, _
    Optional OuterParenthesis As Boolean, _
    Optional Name As String) _
    As String
    
    Const UtcPrefix     As String = "UTC"
    Const OffsetMask    As String = "({0})"
    
    Dim Offset          As Date
    Dim Prefix          As String
    Dim Result          As String
    
    Offset = TimeBias(Bias)
    
    If PrefixUtc = True Then
        Prefix = UtcPrefix
    End If
    If Name = UtcPrefix Then
        Result = Prefix
    Else
        Result = Prefix & FormatSign(Offset, True) & Format(Offset, "hh\:nn")
    End If
    
    If OuterParenthesis = True Then
        Result = Replace(OffsetMask, "{0}", Result)
    End If
    
    FormatBias = Result
    
End Function

' Sorts (by reference) an array holding timezone entries by their bias.
'
' 2019-12-07. Gustav Brock. Cactus Data ApS, CPH.
'
Public Sub SortEntriesBiasLocations(ByRef Entries() As TimezoneEntry)

    Dim Entry   As TimezoneEntry
    Dim Index1  As Integer
    Dim Index2  As Integer
    
    For Index1 = LBound(Entries, 1) To UBound(Entries, 1)
         Entry = Entries(Index1)
         For Index2 = LBound(Entries, 1) To UBound(Entries, 1)
             If Entries(Index2).Bias < Entry.Bias Then
                 Entries(Index1) = Entries(Index2)
                 Entries(Index2) = Entry
                 Entry = Entries(Index1)
             ElseIf Entries(Index2).Bias = Entry.Bias Then
                If Entries(Index2).Locations > Entry.Locations Then
                    Entries(Index1) = Entries(Index2)
                    Entries(Index2) = Entry
                    Entry = Entries(Index1)
                End If
             End If
         Next Index2
     Next Index1
        
End Sub

