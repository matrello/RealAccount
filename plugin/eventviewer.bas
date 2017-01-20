Attribute VB_Name = "eventViewer"
'    RealAccount v1.2
'    Code by Matro
'    Rome, Italy, 2002-2004
'    matro@realpopup.it
'
'    designed for MS Outlook 10 and later

' original source code for the LogErrorToEventViewer() function
' by John Conwell, http://www.freevbcode.com/ShowCode.asp?ID=3490
' adapted by Matro for RealAccount v1.1 build 171 and later

Option Explicit

Public LogApplication As String

Public Enum enmLogType
   LogError = 1&
   LogWarning = 2&
   LogInfo = 4&
End Enum

Public Enum enmErrLevel
   lInfo = &H60000000
   lWarning = &HA0000000
   lError = &HE0000000
End Enum

Private Declare Function RegisterEventSource _
   Lib "advapi32" Alias "RegisterEventSourceA" _
   (ByVal lpUNCServerName As String, _
    ByVal lpSourceName As String) As Long

Private Declare Function DeregisterEventSource _
   Lib "advapi32" _
   (ByVal hEventLog As Long) As Long

Private Declare Function ReportEvent _
   Lib "advapi32" Alias "ReportEventA" _
   (ByVal hEventLog As Long, _
    ByVal wType As Long, _
    ByVal wCategory As Long, _
    ByVal dwEventID As Long, _
    ByVal lpUserSid As Long, _
    ByVal wNumStrings As Long, _
    ByVal dwDataSize As Long, _
    lpStrings As Any, _
    lpRawData As Any) As Long


Public Function LogErrorToEventViewer(sErrMsg As String, eEventType As LogEventTypeConstants) As Boolean
    Dim lEventLogHwnd As Long
    Dim LogType As enmLogType
    Dim lEventID As Long
    Dim lCategory As Long
    Dim sServerName As String
    Dim lRet As Long
   
    LogErrorToEventViewer = True
    lCategory = 0
    sServerName = vbNullString
            
    If eEventType = vbLogEventTypeError Then
        LogType = LogError
    ElseIf eEventType = vbLogEventTypeInformation Then
        LogType = LogInfo
    ElseIf eEventType = vbLogEventTypeWarning Then
        LogType = LogWarning
    End If
    
    lEventLogHwnd = RegisterEventSource(lpUNCServerName:=sServerName, lpSourceName:=LogApplication)
    
    If lEventLogHwnd = 0 Then
        LogErrorToEventViewer = False
        Exit Function
    End If
    
    lRet = ReportEvent(hEventLog:=lEventLogHwnd, _
                       wType:=LogType, _
                       wCategory:=lCategory, _
                       dwEventID:=1, _
                       lpUserSid:=0, _
                       wNumStrings:=1, _
                       dwDataSize:=0, _
                       lpStrings:=sErrMsg, _
                       lpRawData:=0)
                       
    If lRet = False Then
        LogErrorToEventViewer = False
    End If
                       
    DeregisterEventSource lEventLogHwnd
End Function


