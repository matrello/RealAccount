VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    RealAccount v1.2
'    Code by Matro
'    Rome, Italy, 2002-2004
'    matro@realpopup.it
'
'    designed for MS Outlook 10 and later

Option Explicit

Dim WithEvents myItems As Outlook.items
Attribute myItems.VB_VarHelpID = -1
Dim oMAPIFolder As MAPIFolder

Property Set MAPIFolder(myMAPIFolder As MAPIFolder)

    Set oMAPIFolder = myMAPIFolder
    Set myItems = oMAPIFolder.items

End Property

Property Get MAPIFolder() As MAPIFolder
    
    Set MAPIFolder = oMAPIFolder

End Property

Private Sub Class_Terminate()

    Set oMAPIFolder = Nothing
    Set myItems = Nothing

End Sub

Private Sub myItems_ItemAdd(ByVal Item As Object)

    Dim ClassType As String, itemEntryID As String, itemStoreID As String
    Dim MAPISession, MAPIMessage
    
    Call Log("OnItemAdd", "called", LOG_DEBUG)
    
    On Error Resume Next
    
    Select Case Item.Class
        Case olMail
            ClassType = "Mail"
        Case olMeetingCancellation
            ClassType = "MeetingCancellation"
        Case olMeetingResponseNegative
            ClassType = "MeetingResponseNegative"
        Case olMeetingResponsePositive
            ClassType = "MeetingResponsePositive"
        Case olMeetingResponseTentative
            ClassType = "MeetingResponseTentative"
        Case olMeetingRequest
            ClassType = "MeetingRequest"
        Case olTaskRequestUpdate
            ClassType = "TaskRequestUpdate"
        Case olTaskRequestDecline
            ClassType = "TaskRequestDecline"
        Case olTaskRequestAccept
            ClassType = "TaskRequestAccept"
        Case olTaskRequest
            ClassType = "TaskRequest"
    End Select
    
    If Len(ClassType) > 0 Then
        
        Call Log("OnItemAdd", ClassType & " to '" & oMAPIFolder.Name & "' marking read", LOG_DEBUG)
        itemEntryID = Item.EntryID
        itemStoreID = Item.Parent.StoreID
        Item.UnRead = False
        Item.Save
        
        Call Log("OnItemAdd", ClassType & " to '" & oMAPIFolder.Name & "' marking MAPIMessage read", LOG_DEBUG)
        Set MAPISession = CreateObject("MAPI.Session")
        MAPISession.Logon "", "", False, False
        Set MAPIMessage = MAPISession.GetMessage(itemEntryID, itemStoreID)
        MAPIMessage.UnRead = False
        MAPIMessage.Update True
        MAPISession.Logoff
        Set MAPISession = Nothing
        
        Call Log("OnItemAdd", ClassType & " to '" & oMAPIFolder.Name & "' marked read", LOG_INFO)
        
    End If
    If Err > 0 Then
        Call Log("OnItemAdd", "cannot read class of '" & oMAPIFolder.Name & "' item", LOG_DEBUG)
    End If
    On Error GoTo 0

End Sub

Private Sub myItems_ItemChange(ByVal Item As Object)

    Dim a
    a = 1
    

End Sub
