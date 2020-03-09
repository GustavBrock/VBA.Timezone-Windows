Attribute VB_Name = "WtziBase"
Option Compare Text
Option Explicit

' WtziBase v1.0.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Timezone
'
' Base functions for handling timezones of Windows.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

' Required modules:
'   DateUtil (or project VBA.Date)


' User defined types for timezone information.

Public Type RegTziFormat
    Bias                            As Long
    StandardBias                    As Long
    DaylightBias                    As Long
    StandardDate                    As SystemTime
    DaylightDate                    As SystemTime
End Type

Public Type TimezoneEntry
    Mui                             As Integer
    MuiDaylight                     As Integer
    MuiStandard                     As Integer
    Bias                            As Integer
    Name                            As String
    Utc                             As String
    Locations                       As String
    ZoneDaylight                    As String
    ZoneStandard                    As String
    FirstEntry                      As Variant
    LastEntry                       As Variant
    Tzi                             As RegTziFormat
End Type

' TimezoneInformation holds information about a timezone.
' The two arrays are null-terminated strings, where each element
' holds the byte code for a character, and the last element is a
' null value, ASCII code 0.
Public Type TimezoneInformation
    Bias                            As Long
    StandardName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
    StandardDate                    As SystemTime
    StandardBias                    As Long
    DaylightName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
    DaylightDate                    As SystemTime
    DaylightBias                    As Long
End Type

' Reference:
'   https://msdn.microsoft.com/en-us/library/windows/desktop/ms724253(v=vs.85).aspx
'
' Not used, for reference only.
' Complete dynamic timezone entry.
' Names must be Unicode arrays.
Public Type DynamicTimezoneInformation
    Bias                            As Long
    StandardName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
    StandardDate                    As SystemTime
    StandardBias                    As Long
    DaylightName(0 To (32 * 2 - 1)) As Byte     ' Unicode.
    DaylightDate                    As SystemTime
    DaylightBias                    As Long
    TimezoneKeyName(0 To 255)       As Byte     ' Unicode.
End Type


' Declarations.

' Returns the current UTC time.
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
    ByRef lpSystemTime As SystemTime)
    
' Retrieves the current timezone settings from Windows.
Private Declare PtrSafe Function GetTimezoneInformation Lib "Kernel32.dll" Alias "GetTimeZoneInformation" ( _
    ByRef lpTimezoneInformation As TimezoneInformation) _
    As Long

' Returns the timezone bias of the current date or, using the current rules,
' for another present date.
' If IgnoreDaylightSetting is True, the returned bias will not include a local
' daylight saving time bias.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LocalBiasTimezonePresent( _
    ByVal Date1 As Date, _
    Optional ByVal IgnoreDaylightSetting As Boolean) _
    As Long

    Dim TzInfo  As TimezoneInformation
    Dim TzId    As TimezoneId
    Dim Bias    As Long
    
    TzId = GetTimezoneInformation(TzInfo)
    
    Select Case TzId
        Case TimezoneId.Standard, TimezoneId.Daylight
            Bias = TzInfo.Bias
            If IgnoreDaylightSetting = False Then
                If IsCurrentDaylightSavingTime(Date1) = True Then
                    Bias = Bias + TzInfo.DaylightBias
                End If
            End If
    End Select
    
    LocalBiasTimezonePresent = Bias
     
End Function

' Returns the localised friendly description of the current timezone.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LocalTimezoneDescription() As String

    Dim TzInfo      As TimezoneInformation
    Dim Description As String
    
    Select Case GetTimezoneInformation(TzInfo)
        Case TimezoneId.Unknown
            Description = "Cannot determine current timezone"
        Case TimezoneId.Standard
            Description = TzInfo.StandardName
        Case TimezoneId.Daylight
            Description = TzInfo.DaylightName
        Case TimezoneId.Invalid
            Description = "Invalid current timezone"
    End Select
    
    LocalTimezoneDescription = Split(Description, vbNullChar)(0)
   
End Function

' Returns True if the passed date/time value is within the local Daylight Saving Time
' period as defined by the rules for the current year.
'
' Note:
' For a value of the transition interval from daylight time back to standard time,
' where the value could belong to either of these, daylight time is assumed.
' This means that a standard time of the transition interval will return True.
'
' Limitation:
' For dates outside the current year, the same rules as for the current year are used,
' thus the returned result will follow a rule that could be true only for recent or
' near future years relative to the current year.
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsCurrentDaylightSavingTime( _
    Optional ByVal Date1 As Date) _
    As Boolean
    
    Dim TzInfo          As TimezoneInformation
    Dim Dst             As Boolean
    Dim TzId            As TimezoneId
    Dim DaylightDate    As Date
    Dim StandardDate    As Date
    Dim Year            As Integer
    
    TzId = GetTimezoneInformation(TzInfo)
    If Date1 = ZeroDateValue Or Date1 = Date Then
        ' GetTimezoneInformation returns the timezone ID for the current date.
        Dst = (TzId = TimezoneId.Daylight)
    Else
        ' Calculate DaylightDate starting date and standard starting date for Year.
        ' wDay is the occurrence of the weekday in the month. 5 is the last occurrence.
        Year = VBA.Year(Date1)
        DaylightDate = _
            DateWeekdayInMonth(DateSerial(Year, TzInfo.DaylightDate.wMonth, 1), TzInfo.DaylightDate.wDay, vbSunday) + _
            TimeSerial(TzInfo.DaylightDate.wHour, TzInfo.DaylightDate.wMinute, TzInfo.DaylightDate.wSecond)
        StandardDate = _
            DateWeekdayInMonth(DateSerial(Year, TzInfo.StandardDate.wMonth, 1), TzInfo.StandardDate.wDay, vbSunday) + _
            TimeSerial(TzInfo.StandardDate.wHour, TzInfo.StandardDate.wMinute, TzInfo.StandardDate.wSecond)
            
        ' Check if Date1 falls within the period of Daylight Saving Time for Year.
        If DaylightDate < StandardDate Then
            ' Northern hemisphere.
            If DateDiff("s", DaylightDate, Date1) >= 0 And DateDiff("s", Date1, StandardDate) > 0 Then
                Dst = True
            End If
        Else
            ' Southern hemisphere.
            If DateDiff("s", StandardDate, Date1) >= 0 And DateDiff("s", Date1, DaylightDate) > 0 Then
                Dst = True
            End If
        End If
    End If
    
    IsCurrentDaylightSavingTime = Dst
  
End Function

' Retrieves the current date from the local computer as UTC.
' Resolution is one day to mimic Date().
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UtcDate() As Date

    Dim SysTime     As SystemTime
    Dim Datetime    As Date
    
    ' Retrieve current UTC date.
    GetSystemTime SysTime
    
    Datetime = DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay)
    
    UtcDate = Datetime
    
End Function

' Retrieves the current time from the local computer as UTC.
' By cutting off the milliseconds, the resolution is one second to mimic Time().
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UtcTime() As Date

    Dim SysTime     As SystemTime
    Dim Datetime    As Date
    
    ' Retrieve current UTC time.
    GetSystemTime SysTime
    
    Datetime = TimeSerial(SysTime.wHour, SysTime.wMinute, SysTime.wSecond)
    
    UtcTime = Datetime
    
End Function

' Retrieves the current date and time from the local computer as UTC.
' By cutting off the milliseconds, the resolution is one second to mimic Now().
'
' 2016-06-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UtcNow() As Date

    Dim SysTime     As SystemTime
    Dim Datetime    As Date
    
    ' Retrieve current UTC date/time.
    GetSystemTime SysTime
    
    Datetime = _
        DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay) + _
        TimeSerial(SysTime.wHour, SysTime.wMinute, SysTime.wSecond)
    
    UtcNow = Datetime
    
End Function


