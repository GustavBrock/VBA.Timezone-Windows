VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WindowsTimezones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CurrentTimezone_Click()

    Me!Timezones.Value = TimezoneCurrent.Name
    
End Sub

Private Sub Form_Load()
  
    Call StyleCommandButtons(Me)
  
End Sub

