VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timezone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Worksheet to demonstrate selection of a timezone and
' display of some of the information from this timezone.
'
' 2020-03-05. Gustav Brock, Cactus Data ApS, CPH.


' Have a cell select to simulate a button click.
'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ' Named Range used as "button".
    Const ButtonName            As String = "CurrentTimezone"
    ' Named Range holding the timezone name (the key).
    Const TimezoneName          As String = "TimezoneName"
    
    Dim ButtonRange             As Excel.Range
    Dim NameRange               As Excel.Range
    
    Set ButtonRange = ThisWorkbook.Names(ButtonName).RefersToRange
    Set NameRange = ThisWorkbook.Names(TimezoneName).RefersToRange
    
    If Target.Address = ButtonRange.Address Then
        ' The user has selected the "button" and only this.
        ' Set the timezone to the current timezone displayed in "Windows style".
        NameRange.Value = FormatBias(TimezoneCurrent.Bias, True, True, TimezoneCurrent.Name) & " " & TimezoneCurrent.Locations
        
        ' Reset formulas if needed.
        ResetReferences
    End If
    
End Sub

' Reset the formulas for the timezone look up cells
' if the references for these have been lost.
'
Private Sub ResetReferences()

    Const MuiName               As String = "TimezoneMui"
    Const DescriptionName       As String = "TimezoneDescription"
    
    Const MuiFormula            As String = "=INDEX(WindowsTimezone[MUI],MATCH(TimezoneName,WindowsTimezone[Display],0))"
    Const DescriptionFormula    As String = "=INDEX(WindowsTimezone[ZoneStd],MATCH(TimezoneName,WindowsTimezone[Display],0))"

    Dim MuiRange                As Excel.Range
    Dim DescriptionRange        As Excel.Range
    
    Set MuiRange = ThisWorkbook.Names(MuiName).RefersToRange
    Set DescriptionRange = ThisWorkbook.Names(DescriptionName).RefersToRange
    
    If MuiRange.Formula <> MuiFormula Then
        MuiRange.Formula = MuiFormula
    End If
    If DescriptionRange.Formula <> DescriptionFormula Then
        DescriptionRange.Formula = DescriptionFormula
    End If

End Sub

