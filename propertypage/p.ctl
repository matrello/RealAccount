VERSION 5.00
Begin VB.UserControl p 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdOptions 
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "@13"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox chkReplace 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4020
      Width           =   255
   End
   Begin VB.CommandButton cmdManualApply 
      Caption         =   "manual apply (debug)"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optEnable 
      Enabled         =   0   'False
      Height          =   210
      Index           =   1
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   0
      Tag             =   "@2"
      Top             =   400
      Width           =   3255
   End
   Begin VB.Frame frmSettingz 
      Height          =   3530
      Left            =   120
      TabIndex        =   8
      Top             =   385
      Width           =   4455
      Begin VB.OptionButton optFormat 
         Caption         =   "Text"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Tag             =   "@12"
         Top             =   2390
         Width           =   855
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "HTML"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Tag             =   "@11"
         Top             =   2390
         Width           =   855
      End
      Begin VB.CheckBox chkUseFormat 
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Tag             =   "@10"
         Top             =   2030
         Value           =   2  'Grayed
         Width           =   3495
      End
      Begin VB.CheckBox chkMarkRead 
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Tag             =   "@6"
         Top             =   2640
         Value           =   2  'Grayed
         Width           =   3735
      End
      Begin VB.CheckBox chkSignatureTop 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Tag             =   "@5"
         Top             =   1670
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkUseSignature 
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Tag             =   "@4"
         Top             =   950
         Value           =   2  'Grayed
         Width           =   2775
      End
      Begin VB.ComboBox cbSignatures 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1325
         Width           =   2895
      End
      Begin VB.CheckBox chkUseAccount 
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Tag             =   "@3"
         Top             =   230
         Value           =   2  'Grayed
         Width           =   2775
      End
      Begin VB.ComboBox cbAccounts 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   605
         Width           =   2895
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Tag             =   "@7"
         Top             =   3020
         UseMnemonic     =   0   'False
         Width           =   4215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.OptionButton optEnable 
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Tag             =   "@1"
      Top             =   70
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.Label lbAbout 
      Caption         =   "sys 64738"
      Height          =   255
      Left            =   405
      TabIndex        =   17
      Tag             =   "@9"
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Image imAbout 
      Height          =   240
      Left            =   105
      Picture         =   "p.ctx":0000
      Top             =   4560
      Width           =   240
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   390
      TabIndex        =   12
      Tag             =   "@8"
      Top             =   4020
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "p"
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

Const DEBUG_SHOWAPPLYBUTTON = False
Const DEBUG_FOLDERID = "00000000BF7ECAF668FDEB4F8C7A95926ACB64D382800000" ' bbttmrtr posta in arrivo - email.it
'Const DEBUG_FOLDERID = "0000000006702C79BCCD1545AC3C35E9D4E74A9982800000"   ' matrodesktop
'Const DEBUG_FOLDERID = "00000000142B757B1EFD1845A54E55EABAB60149E2810000" ' posta in arrivo
'Const DEBUG_FOLDERID = "00000000142B757B1EFD1845A54E55EABAB6014902820000" '  posta in arrivo - email.it

Implements outlook.PropertyPage

Private oSite As outlook.PropertyPageSite
Private Initializing As Boolean
Private isDirty As Boolean
Private Accounts As New Collection
Private AccountNotDefined As String
Private signatures As New Collection
Private SignatureNotDefined As String
Private currentfolder As MAPIFolder
Private currentfolderID As String
Private currentProfileName As String
Private Language As Long

Public Event PropertyChange(FolderID As String)
Public Event OptionsOK()

Private Sub SetDirty()
    
    If Not oSite Is Nothing And Not Initializing Then
        isDirty = True
        oSite.OnStatusChange
    End If

End Sub

Private Sub cbAccounts_Change()

    Call SetDirty

End Sub

Private Sub cbAccounts_Click()

    Call SetDirty

End Sub

Private Sub cbSignatures_Change()

    Call SetDirty

End Sub

Private Sub cbSignatures_Click()

    Call SetDirty

End Sub

Private Sub chkMarkRead_Click()

    Static istat%
    Static bdone As Boolean
    
    If bdone Then
      bdone = False
      Exit Sub
    End If
    
    If istat% = 1 Then
      bdone = True
      chkMarkRead.Value = 2
    End If

    istat% = chkMarkRead.Value
        
    Call SetControlsEnabled
    Call SetDirty

End Sub

Private Sub chkReplace_Click()

    Call SetDirty

End Sub

Private Sub chkSignatureTop_Click()

    Call SetDirty

End Sub

Private Sub chkUseAccount_Click()

  Static istat%
  Static bdone As Boolean

  If bdone Then
    bdone = False
    Exit Sub
  End If

  If istat% = 1 Then
    bdone = True
    chkUseAccount.Value = 2
  End If

  istat% = chkUseAccount.Value
  
  Call SetControlsEnabled
  Call SetDirty

End Sub

Private Sub chkUseFormat_Click()

  Static istat%
  Static bdone As Boolean

  If bdone Then
    bdone = False
    Exit Sub
  End If

  If istat% = 1 Then
    bdone = True
    chkUseFormat.Value = 2
  End If

  istat% = chkUseFormat.Value
  
  Call SetControlsEnabled
  Call SetDirty

End Sub

Private Sub chkUseSignature_Click()

    Static istat%
    Static bdone As Boolean
    
    If bdone Then
      bdone = False
      Exit Sub
    End If
    
    If istat% = 1 Then
      bdone = True
      chkUseSignature.Value = 2
    End If

    istat% = chkUseSignature.Value
        
    Call SetControlsEnabled
    Call SetDirty
    
End Sub

Private Sub cmdOptions_Click()

    frmOptions.Show vbModal, UserControl
    
    Call SetLanguage
    RaiseEvent OptionsOK
    
End Sub

Private Sub optEnable_Click(Index As Integer)

    Call SetControlsEnabled
    Call SetDirty

End Sub

Private Sub optFormat_Click(Index As Integer)

    Call SetDirty

End Sub

Private Property Get PropertyPage_Dirty() As Boolean
    
    PropertyPage_Dirty = isDirty

End Property

Private Sub PropertyPage_GetPageInfo(HelpFile As String, HelpContext As Long)
    
'        HelpFile = "C:\Matro\sorgenti\RealPopup\RPHelpENG.htm"
'        HelpContext = 0

End Sub

Private Sub UserControl_EnterFocus()
    
    Dim Item
    
    Call Log("UserControl_EnterFocus", "called", LOG_DEBUG)
    
    If currentfolder Is Nothing Then Exit Sub
    
    cbAccounts.Visible = False
    cbAccounts.Clear
    For Each Item In Accounts
        cbAccounts.AddItem Item
    Next Item
    cbAccounts.ListIndex = 0
    cbAccounts.Visible = True
    
    cbSignatures.Visible = False
    cbSignatures.Clear
    For Each Item In signatures
        cbSignatures.AddItem Item
    Next Item
    cbSignatures.ListIndex = 0
    cbSignatures.Visible = True
    
    If RunningIDE Or DEBUG_SHOWAPPLYBUTTON Then cmdManualApply.Visible = True
    
    Call LoadFolderSettings
    Call SetControlsEnabled
    
    UserControl.Refresh
    Initializing = False

End Sub

Private Sub UserControl_ExitFocus()

    Call Log("UserControl_ExitFocus", "called", LOG_DEBUG)

End Sub

Private Sub UserControl_Hide()

    Call Log("UserControl_Hide", "called", LOG_DEBUG)

End Sub

Private Sub UserControl_Initialize()
    
    LogApplication = "RealAccount page"
    Call Log("RealAccountPPage", "version " & App.Major & "." & App.Minor & " build " & Format(App.Revision, "000") & " session started", LOG_STRONGINFO, "RealAccount page.log")

End Sub

Private Sub UserControl_InitProperties()
    
    Dim dummy&, debugMAPISession, CDOVersion As String
    
    Call Log("UserControl_InitProperties", "called", LOG_DEBUG)
    
    On Error Resume Next
    Set oSite = Parent
    Err.Number = 0
    
    isDirty = False
    Initializing = True

    Call SetLanguage

    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, dummy, "UseEntryID")
    UseEntryID = CBool(dummy)

    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, App.Major & "." & App.Minor & " build " & Format(App.Revision, "000"), "RealAccountPPageVersion")
    AccountNotDefined = "(not defined)"
    Accounts.Add AccountNotDefined
    SignatureNotDefined = "(not defined)"
    signatures.Add SignatureNotDefined
    
    If RunningIDE Then
        currentfolderID = DEBUG_FOLDERID
        Set debugMAPISession = outlook.ActiveExplorer.Session
        debugMAPISession.Logon , , False, False
        Set currentfolder = debugMAPISession.GetFolderFromID(currentfolderID)
    Else
        Set currentfolder = oSite.Session.GetFolderFromID(currentfolderID)
    End If
        
    Set debugMAPISession = CreateObject("mapi.session")
    debugMAPISession.Logon , , False, False
    CDOVersion = debugMAPISession.Version
    currentProfileName = debugMAPISession.Name
    Set debugMAPISession = Nothing
        
    If Len(CDOVersion) > 0 Then
        Call Log("UserControl_InitProperties", "CDO version: " & CDOVersion, LOG_INFO)
    Else
        Call Log("UserControl_InitProperties", "CDO library not detected.", LOG_WARNING)
    End If
    
    If Len(currentProfileName) > 0 Then
        Call Log("UserControl_InitProperties", "Current profile name: '" & currentProfileName & "'", LOG_INFO)
    Else
        currentProfileName = "Outlook"
        Call Log("UserControl_InitProperties", "Cannot get current profile name; looking for 'Outlook' profile", LOG_WARNING)
    End If
    
    Call GetOutlookAccounts
    Call EnumSignatures(signatures)
        
End Sub

Public Property Get Name() As Variant
Attribute Name.VB_UserMemId = -518

    Name = "RealAccount"

End Property

Public Property Let FolderID(myFolderID As String)

    currentfolderID = myFolderID

End Property

Private Sub GetOutlookAccounts()

    Dim dummy1 As New Collection, dummy2 As New Collection
    Dim Item, item1, item2, item3, ok As Boolean
    Dim ProfileKey As String

    Call Log("GetOutlookAccounts", "called", LOG_DEBUG)

    If IsNT Then
        ProfileKey = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
    Else
        ProfileKey = "Software\Microsoft\Windows Messaging Subsystem\Profiles\"
    End If

    If EnumRegKey(HKEY_CURRENT_USER, ProfileKey & currentProfileName, dummy1) Then
        For Each item1 In dummy1
            If EnumRegKey(HKEY_CURRENT_USER, ProfileKey & currentProfileName & "\" & item1, dummy2) Then
                If dummy2.Count > 0 Then
                    For Each item2 In dummy2
                        If InStr(1, item2, "000000") > 0 Then
                            If (GetRegValue(HKEY_CURRENT_USER, ProfileKey & currentProfileName & "\" & item1 & "\" & item2, REG_BINARY, Item, "Account Name")) Then
                                If (GetRegValue(HKEY_CURRENT_USER, ProfileKey & currentProfileName & "\" & item1 & "\" & item2, REG_BINARY, item3, "Email")) Or _
                                    (GetRegValue(HKEY_CURRENT_USER, ProfileKey & currentProfileName & "\" & item1 & "\" & item2, REG_BINARY, item3, "Identity Eid")) Then
                                    Accounts.Add Item
                                End If
                            End If
                        End If
                    Next item2
                Else
                    Set dummy2 = New Collection
                End If
            End If
        Next item1
    End If

    Set dummy2 = New Collection

    If EnumRegKey(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts", dummy2) Then
        If dummy2.Count > 0 Then
            For Each item2 In dummy2
                If InStr(1, item2, "000000") > 0 Then
                    If (GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts\" & item1 & "\" & item2, REG_BINARY, Item, "Account Name")) Then
                        If (GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts\" & item1 & "\" & item2, REG_BINARY, item3, "Email")) Or _
                            (GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts\" & item1 & "\" & item2, REG_BINARY, item3, "Identity Eid")) Then
                            ok = True
                            For Each item1 In Accounts
                                If item1 = Item Then ok = False: Exit For
                            Next item1
                            
                            If ok Then Accounts.Add Item
                        End If
                    End If
                End If
            Next item2
        End If
    End If

End Sub

Private Sub LoadFolderSettings()

    Dim enabled&, accountenabled&, signatureenabled&, account$, signature$, signaturetop&
    Dim formatenabled&, eformat&, markread&

    Call Log("LoadFolderSettings", "called", LOG_DEBUG)

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_SZ, account, "Account") Then
        Call SetCombo(cbAccounts, account)
    Else
        cbAccounts.ListIndex = 0
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_SZ, signature, "Signature") Then
        Call SetCombo(cbSignatures, signature)
    Else
        cbSignatures.ListIndex = 0
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, eformat, "Format") Then
        If eformat = 1 Or eformat = 2 Then optFormat(eformat).Value = True
    Else
        optFormat(1).Value = False
        optFormat(2).Value = False
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, markread, "MarkRead") Then
        chkMarkRead.Value = markread
    Else
        chkMarkRead.Value = vbGrayed
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, signaturetop, "SignatureTop") Then
        chkSignatureTop.Value = signaturetop
    Else
        chkSignatureTop.Value = vbChecked
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, accountenabled, "AccountEnabled") Then
        chkUseAccount.Value = accountenabled
    Else
        chkUseAccount.Value = vbGrayed
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, signatureenabled, "SignatureEnabled") Then
        chkUseSignature.Value = signatureenabled
    Else
        chkUseSignature.Value = vbGrayed
    End If
    
    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, formatenabled, "FormatEnabled") Then
        chkUseFormat.Value = formatenabled
    Else
        chkUseFormat.Value = vbGrayed
    End If

    If GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, enabled, "Enabled") Then
        If enabled Then optEnable(1).Value = True
    End If

End Sub

Private Sub SetControlsEnabled()

    optEnable(0).enabled = True
    optEnable(1).enabled = True
    chkReplace.enabled = True
    chkUseAccount.enabled = optEnable(1).Value
    chkUseSignature.enabled = optEnable(1).Value
    chkUseFormat.enabled = optEnable(1).Value
    chkMarkRead.enabled = optEnable(1).Value
    cbAccounts.enabled = chkUseAccount.enabled And (chkUseAccount.Value = vbChecked)
    cbSignatures.enabled = chkUseSignature.enabled And (chkUseSignature.Value = vbChecked)
    optFormat(1).enabled = chkUseFormat.enabled And (chkUseFormat.Value = vbChecked)
    optFormat(2).enabled = optFormat(1).enabled
    chkSignatureTop.enabled = chkUseSignature.enabled And (chkUseSignature.Value = vbChecked)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Log("UserControl_KeyDown", "called: " & KeyCode, LOG_DEBUG)

    If KeyCode = 27 Then
        KeyCode = 0
        Call Log("UserControl_KeyDown", "ESC key inhibited", LOG_INFO)
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    Call Log("UserControl_KeyPress", "called: " & KeyAscii, LOG_DEBUG)

    If KeyAscii = 27 Then KeyAscii = 0

End Sub

Private Sub UserControl_LostFocus()

    Call Log("UserControl_LostFocus", "called", LOG_DEBUG)

End Sub

Private Sub UserControl_Resize()

    imAbout.Top = UserControl.Height - 390
    lbAbout.Top = imAbout.Top
    cmdOptions.Top = imAbout.Top - 120

End Sub

Private Sub UserControl_Terminate()

    Call Log("RealAccountPPage", "version " & App.Major & "." & App.Minor & " build " & Format(App.Revision, "000") & " session closed", LOG_STRONGINFO)

End Sub

Private Sub PropertyPage_Apply()
    
    Dim pfolder As MAPIFolder
    
    Call Log("PropertyPage_Apply", "called", LOG_DEBUG)
    
    If cbAccounts.text = AccountNotDefined And chkUseAccount.Value = vbChecked Then chkUseAccount.Value = vbGrayed
    If cbSignatures.text = SignatureNotDefined And chkUseAccount.Value = vbChecked Then chkUseSignature.Value = vbGrayed
    
    Call OSRegCreateKey(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), 0)
    
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, IIf(optEnable(1).Value, 1, 0), "Enabled")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_SZ, cbAccounts.text, "Account")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_SZ, cbSignatures.text, "Signature")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, chkUseAccount.Value, "AccountEnabled")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, chkUseSignature.Value, "SignatureEnabled")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, chkUseFormat.Value, "FormatEnabled")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, chkSignatureTop.Value, "SignatureTop")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, chkMarkRead.Value, "MarkRead")
    
    If Not optFormat(1).Value And Not optFormat(2).Value Then
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, 0, "Format")
    Else
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(currentfolder), REG_DWORD, IIf(optFormat(1).Value, 1, 2), "Format")
    End If
    
    If chkReplace.Value = vbChecked Then Call ReplaceChildFoldersSettings(currentfolder)
    
    Call Log("PropertyPage_Apply", "Saved settings for folder '" & GetRealAccountFolder(currentfolder) & "'", LOG_INFO)
    
    RaiseEvent PropertyChange(currentfolderID)
    isDirty = False

End Sub

Private Sub ReplaceChildFoldersSettings(Folder As MAPIFolder)

    Dim childFolder As MAPIFolder
    
    For Each childFolder In Folder.Folders
        Call ReplaceChildFoldersSettings(childFolder)
    Next childFolder
    
    If currentfolder <> Folder Then
        Call OSRegCreateKey(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), 0)
        
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, 1, "Enabled")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_SZ, AccountNotDefined, "Account")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_SZ, SignatureNotDefined, "Signature")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, 0, "Format")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, vbGrayed, "AccountEnabled")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, vbGrayed, "SignatureEnabled")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, vbGrayed, "FormatEnabled")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, vbChecked, "SignatureTop")
        Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(Folder), REG_DWORD, vbGrayed, "MarkRead")
    End If
    
End Sub

Private Sub cmdManualApply_Click()

    Call PropertyPage_Apply
    
End Sub

Private Sub SetLanguage()

    Dim Item, strings As New Collection, k&
    Dim RealAccountVersion$, Language&
    
    Call Log("SetLanguage", "called", LOG_DEBUG)
    
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, RealAccountVersion, "RealAccountPluginVersion")
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, Language, "Language")
    
    If Language = 1 Then
        strings.Add "Disable RealAccount for this folder", "@1"
        strings.Add "Enable RealAccount for this folder", "@2"
        strings.Add "Use specified account", "@3"
        strings.Add "Use specified signature", "@4"
        strings.Add "Insert signature on top", "@5"
        strings.Add "Mark all messages read", "@6"
        strings.Add "Grayed options inherit parent's settings (if any).", "@7"
        strings.Add "Update child folders to inherit above settings", "@8"
        strings.Add "RealAccount v" & RealAccountVersion & ".", "@9"
        strings.Add "Use specified format for new emails", "@10"
        strings.Add "HTML", "@11"
        strings.Add "Text", "@12"
        strings.Add "Options...", "@13"
    Else
        strings.Add "Non utilizzare RealAccount per questa cartella", "@1"
        strings.Add "Abilita RealAccount per questa cartella", "@2"
        strings.Add "Utilizza un account specifico", "@3"
        strings.Add "Utilizza una firma specifica", "@4"
        strings.Add "Inserisci la firma in cima al messaggio", "@5"
        strings.Add "Segna tutti i messaggi come già letti", "@6"
        strings.Add "Le opzioni in grigio ereditano le impostazioni dalla cartella padre (se esistente).", "@7"
        strings.Add "Reimposta le cartelle sottostanti per ereditare queste impostazioni", "@8"
        strings.Add "RealAccount v" & RealAccountVersion & ".", "@9"
        strings.Add "Imposta il formato per le nuove email", "@10"
        strings.Add "HTML", "@11"
        strings.Add "Testo", "@12"
        strings.Add "Opzioni...", "@13"
    End If
    
    On Error Resume Next
    For Each Item In Controls
        Item.Caption = strings(Item.Tag)
        If Err.Number > 0 Then
            Err.Number = 0
            Item.text = strings(Item.Tag)
        End If
        Err.Number = 0
    Next
    Err.Number = 0

    optEnable(0).Width = TextWidth(strings("@1")) + 420
    optEnable(1).Width = TextWidth(strings("@2")) + 420
    lblInfo.Top = frmSettingz.Height - lblInfo.Height - 60
    Refresh

End Sub

Private Sub SetCombo(oCombo As ComboBox, text As String)

    Dim k&
    
    For k = 0 To oCombo.ListCount - 1
        If StrComp(oCombo.List(k), text, vbTextCompare) = 0 Then
            oCombo.ListIndex = k
            Exit For
        End If
    Next k
    
    If k = oCombo.ListCount Then
        oCombo.ListIndex = 0
    End If

End Sub
