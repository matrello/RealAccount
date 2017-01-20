VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9630
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   12300
   _ExtentX        =   21696
   _ExtentY        =   16986
   _Version        =   393216
   Description     =   "RealAccount plugin for MS Outlook"
   DisplayName     =   "RealAccount"
   AppName         =   "Microsoft Outlook"
   AppVer          =   "Microsoft Outlook 10.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook"
End
Attribute VB_Name = "Connect"
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

Implements IDTExtensibility2

Const CdoPR_EMAIL = &H39FE001E

Dim BetaExpired As Boolean

Dim WithEvents oApp As Outlook.Application
Attribute oApp.VB_VarHelpID = -1
Dim WithEvents oInspectors As Outlook.Inspectors
Attribute oInspectors.VB_VarHelpID = -1
Dim WithEvents oMailItem As Outlook.MailItem
Attribute oMailItem.VB_VarHelpID = -1
Dim WithEvents oAppNS As Outlook.NameSpace
Attribute oAppNS.VB_VarHelpID = -1
Dim WithEvents oNewRealAccountPPage As RealAccountPPage.p
Attribute oNewRealAccountPPage.VB_VarHelpID = -1
Dim WithEvents oTimer As XTimer
Attribute oTimer.VB_VarHelpID = -1

Dim oItems As New Collection

Dim idCmdAccount&, SignatureInserted As Boolean
Dim curFolder As MAPIFolder
Dim language As Long, MarkReadRetries As Long

Public WorkInProgress As Boolean

Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)

    Call Log("OnAddInsUpdate", "called", LOG_DEBUG)

End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)

    Dim Item
    
    Call Log("OnBeginShutdown", "called", LOG_DEBUG)
    On Error Resume Next

    Set oApp = Nothing
    Set oAppNS = Nothing
    Set oInspectors = Nothing
    For Each Item In oItems
        Set Item = Nothing
    Next Item
    Set oNewRealAccountPPage = Nothing
    
End Sub

Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    Dim dummy&

    LogApplication = "RealAccount addin"
    Call Log("RealAccount", "version " & GetVersion() & " session started", LOG_STRONGINFO, "RealAccount plugin.log")
    Call Log("OnConnection", "called", LOG_DEBUG)
    
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, GetVersion(), "RealAccountPluginVersion")
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, language, "Language")

    Set oApp = Application
    
    Set oInspectors = oApp.Inspectors
    Set oAppNS = oApp.GetNamespace("MAPI")

    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, idCmdAccount, "cmdAccount") Then
        idCmdAccount = 31224
        If Not SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, idCmdAccount, "cmdAccount") Then
            Call Log("OnConnection", "SetRegValue(cmdAccount) failed", LOG_ERROR)
        End If
    End If
    
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, dummy, "UseEntryID")
    UseEntryID = CBool(dummy)
    
End Sub

Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    Call Log("OnDisconnection", "called", LOG_DEBUG)
    Call Log("RealAccount", "version " & GetVersion() & " session closed", LOG_STRONGINFO)
    
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)

    Call Log("OnStartupComplete", "called", LOG_DEBUG)

    MarkReadRetries = 0
    Set oTimer = New XTimer
    oTimer.Interval = 3000
    oTimer.Enabled = True

End Sub

Private Sub oApp_Startup()

    Dim Welcome As Long, WelcomeBeta As Long, CDOCheck As Long, OutlookVersionCheck As Long
    Dim debugMAPISession, CDOVersion As String, VersionParts
    Dim msgCDO As String, msgVersion As String, tit As String
    
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, WelcomeBeta, "WelcomeBeta")
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, Welcome, "Welcome")
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, CDOCheck, "CDOCheck")
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, OutlookVersionCheck, "OutlookVersionCheck")
    
    BetaExpired = APP_BETA And CBool(DateDiff("d", Format(APP_BETA_DAY & "/" & APP_BETA_MONTH & "/" & APP_BETA_YEAR, "dd/mm/yyyy"), Now) >= 0)

    If Welcome = 0 And oApp.Explorers.Count > 0 Then
        Load frmSplash
        frmSplash.language = language
        frmSplash.Show vbModal, Me
    End If
    
    If APP_BETA And (WelcomeBeta < App.Revision Or BetaExpired) And oApp.Explorers.Count > 0 Then
        Load frmSplash
        frmSplash.language = language
        frmSplash.BetaShowMsg = True
        frmSplash.BetaExpired = BetaExpired
        frmSplash.Show vbModal, Me
    End If

    On Error Resume Next
    Set debugMAPISession = CreateObject("mapi.session")
    debugMAPISession.Logon , , False, False
    CDOVersion = debugMAPISession.Version
    debugMAPISession.Logoff
    Set debugMAPISession = Nothing
    On Error GoTo 0
        
    If language = 1 Then
        msgCDO = "RealAccount detected that the Microsoft Collaboration Data Objects (CDO) " & _
            "library is not installed in your system." & vbCrLf & vbCrLf & _
            "RealAccount needs the CDO library to work properly." & vbCrLf & vbCrLf & _
            "In Microsoft Office XP and later, the CDO library is " & _
            "available as optional installation." & vbCrLf & vbCrLf & _
            "To install CDO, go to Control Panel, Add or Remove Programs, select Microsoft Office " & _
            "then press the change button; choose Add/Remove features, " & _
            "then check Outlook/CDO Object library."
        tit = "RealAccount warning message"
        msgVersion = "RealAccount detected an unsupported Outlook version." & vbCrLf & vbCrLf & _
            "RealAccount is compatible with Outlook XP and later; since it is still active, keep on using " & _
            "or disable it through Tools, Options, Other, Advanced Options, COM add ins."
    Else
        msgCDO = "RealAccount ha rilevato che la libreria Microsoft Collaboration Data Objects (CDO) " & _
            "non è installata su questo computer." & vbCrLf & vbCrLf & _
            "RealAccount necessita della libreria CDO per funzionare correttamente." & vbCrLf & vbCrLf & _
            "In Microsoft Office XP e successivi, la libreria CDO è " & _
            "disponibile come pacchetto di installazione opzionale." & vbCrLf & vbCrLf & _
            "Per installare la libreria CDO, aprire il Pannello di Controllo, Aggiungi o Rimuovi Programmi, selezionare Microsoft Office " & _
            "e premere il pulsante Modifica; selezionare Aggiungi/Rimuovi caratteristiche, " & _
            "quindi selezionare Outlook/Oggetti dati di collaborazione."
        tit = "RealAccount messaggio di avvertimento"
        msgVersion = "RealAccount ha rilevato una versione non supportata di Outlook." & vbCrLf & vbCrLf & _
            "RealAccount è compatibile con Outlook XP e versioni successive. RealAccount è comunque attivo; " & _
            "è possibile continuare ad utilizzarlo oppure disabilitarlo con Strumenti, Opzioni, " & _
            "Altro, Opzioni avanzate, Componenti aggiuntivi COM."
    End If
        
    If Len(CDOVersion) > 0 Then
        Call Log("oApp_Startup", "CDO version: " & CDOVersion, LOG_INFO)
    Else
        Call Log("oApp_Startup", "CDO library not detected.", LOG_WARNING)
        
        If CDOCheck = 1 Then
            MsgBox msgCDO, vbOKOnly + vbExclamation, tit
        End If
    End If
        
    VersionParts = Split(oApp.Version, ".")
        
    If Int(VersionParts(LBound(VersionParts))) < 10 And OutlookVersionCheck = 1 Then
        MsgBox msgVersion, vbOKOnly + vbExclamation, tit
    End If
    
    Call SetClickYes

End Sub

Private Sub oInspectors_NewInspector(ByVal Inspector As Outlook.Inspector)

    Call Log("OnNewInspector", "called", LOG_DEBUG)

    If Inspector.CurrentItem.Class = olMail Then
        Set oMailItem = Inspector.CurrentItem
    
        Set curFolder = oApp.ActiveExplorer.CurrentFolder
        If curFolder Is Nothing Then
            Log "OnNewInspector", "cannot find current folder", LOG_ERROR
            Exit Sub
        End If
        
        SignatureInserted = False
    End If

End Sub

Private Sub oMailItem_BeforeCheckNames(Cancel As Boolean)

    Call Log("OnBeforeCheckNames", "called", LOG_DEBUG)
    Call SetMailItemDefaults

End Sub

Private Sub oMailItem_Open(Cancel As Boolean)

    Call Log("OnOpen", "called", LOG_DEBUG)
    Call SetMailItemDefaults

End Sub

Private Sub oMailItem_PropertyChange(ByVal Name As String)

    If StrComp(Name, "InternetCodepage", vbTextCompare) = 0 Then
        Call Log("OnPropertyChange", "called for " & Name & ": event ignored", LOG_DEBUG)
    Else
        Call Log("OnPropertyChange", "called for " & Name, LOG_DEBUG)
        If Not WorkInProgress Then Call SetMailItemDefaults
    End If

End Sub

Private Sub SetMailItemDefaults(Optional SetAccountOnly)

    Dim Account$, signature$, SignatureTop&, Format&, mySignature$, myBody$, myBodyPos&, myBodyPos2&, HTMLBody$
    Dim accCommandBar As CommandBar, accControl As CommandBarPopup, dummy&, Item
    Dim hwnds() As Long, r As Double, curFocus&
    
    Log "SetMailItemDefaults", "called", LOG_DEBUG
    
    If BetaExpired Or oMailItem.Sent Then Exit Sub
    
    On Error GoTo myError
    
    Log "SetMailItemDefaults", "beta is not expired and this is a new message", LOG_DEBUG
    
    If IsMissing(SetAccountOnly) Then SetAccountOnly = False
    
    If Not GetRealAccount(curFolder, Account, signature, SignatureTop, dummy, Format) Then
        Log "SetMailItemDefaults", "no RealAccount for folder " & curFolder.Name, LOG_INFO
        Exit Sub
    End If
    
    Set accCommandBar = oMailItem.GetInspector.CommandBars("Standard")
    If accCommandBar Is Nothing Then
        Log "SetMailItemDefaults", "cannot find 'Standard' CommandBar", LOG_WARNING
    Else
    
        If Format > 0 And Not SetAccountOnly Then
            WorkInProgress = True
            If Len(oMailItem.Body) > 1 Then
                Log "SetMailItemDefaults", "Message format not set: item is not a new email", LOG_DEBUG
            Else
                oMailItem.BodyFormat = Format
                Log "SetMailItemDefaults", "Message format set to " & IIf(Format = olFormatPlain, "text", "html"), LOG_INFO
            End If
            WorkInProgress = False
        End If
        
        Set accControl = accCommandBar.FindControl(msoControlPopup, idCmdAccount, , , True)
        If accControl Is Nothing Then
            Log "SetMailItemDefaults", "Account control (" & idCmdAccount & ") not found", LOG_INFO
            Exit Sub
        Else
            If Not accControl.Enabled Then
                Log "SetMailItemDefaults", "Account control (" & idCmdAccount & ") is not Enabled", LOG_DEBUG
            Else
                For Each Item In accControl.Controls
                    If InStr(Item.Caption, "&") > 0 And StrComp(Right$(Item.Caption, Len(Account)), Account, vbTextCompare) = 0 Then
                        Log "SetMailItemDefaults", "account '" & Account & "' inserted", LOG_INFO
                        Item.Execute
                        Exit For
                    End If
                Next Item
            End If
        End If
    End If
    
    Set accControl = Nothing
    Set accCommandBar = Nothing
    
    If Len(signature) = 0 Then SignatureInserted = True

    If Not SignatureInserted And Not SetAccountOnly Then
    
        WorkInProgress = True
        mySignature = GetSignature(signature, oMailItem.BodyFormat)
        
        If mySignature = "" Then
            Log "SetMailItemDefaults", "cannot load signature '" & signature & "'", LOG_WARNING
            WorkInProgress = False
        Else
                    
            Select Case oMailItem.BodyFormat
                Case olFormatHTML, olFormatRichText
                
                    If Len(ClickYesVBSPath) > 0 And ClickYes = vbChecked Then
                        r = Shell("wscript " & ClickYesVBSPath, vbHide)
                        Log "SetMailItemDefaults", "ClickYes script called, return is " & r, LOG_DEBUG
                    End If
                    HTMLBody = oMailItem.HTMLBody
                    
                    If SignatureTop Then
                        myBody = HTMLBody
                        myBodyPos = InStr(1, myBody, "<body", vbTextCompare)
                        myBodyPos2 = myBodyPos + 5
                        Do While myBodyPos2 <= Len(myBody)
                            If Mid$(myBody, myBodyPos2, 1) = ">" Then Exit Do
                            myBodyPos2 = myBodyPos2 + 1
                        Loop
                        myBody = Mid$(myBody, myBodyPos, myBodyPos2 - myBodyPos + 1)
                        oMailItem.HTMLBody = Replace$(HTMLBody, myBody, myBody & mySignature & "<P>")
                    Else
                        oMailItem.HTMLBody = Replace$(HTMLBody, "</BODY>", "<P>" & mySignature & "</BODY>")
                    End If
                    SendKeys "^{HOME}"
                Case olFormatPlain
                    If SignatureTop Then
                        oMailItem.Body = mySignature & oMailItem.Body
                    Else
                        oMailItem.Body = oMailItem.Body & mySignature
                    End If
                    SendKeys "^{HOME}"
            End Select
                        
            SignatureInserted = True
            WorkInProgress = False
            Log "SetMailItemDefaults", "signature '" & signature & "' inserted", LOG_INFO
        End If
    
    End If

    Exit Sub
    
myError:
    
    Log "SetMailItemDefaults", "error: (" & Err.Number & ") " & Err.Description, LOG_ERROR
    WorkInProgress = False

End Sub

Private Function GetRealAccount(Folder As MAPIFolder, Account As String, signature As String, SignatureTop As Long, MarkRead As Long, Format As Long) As Boolean

    Dim accountOk As Boolean, signatureOk As Boolean, signaturetopOk As Boolean, markreadOk As Boolean, formatOK As Boolean
    Dim SignatureEnabled&, AccountEnabled&, FormatEnabled&, signaturedenied As Boolean, accountdenied As Boolean, formatdenied As Boolean
    Dim myFolder As MAPIFolder, myGetFolder$, Enabled&
    
    Set myFolder = Folder
    myGetFolder = "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(myFolder)

    Do While Not (accountOk And signatureOk And markreadOk And formatOK)
        
        Call Log("GetRealAccount", "looking for folder '" & myGetFolder & "'", LOG_DEBUG)
        
        Enabled = 0
        Call GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, Enabled, "Enabled")
        If CBool(Enabled) Then
        
            Call GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, AccountEnabled, "AccountEnabled")
            Call GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, SignatureEnabled, "SignatureEnabled")
            Call GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, FormatEnabled, "FormatEnabled")
            
            If Not markreadOk Then
                If GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, MarkRead, "MarkRead") Then
                    If MarkRead <> vbGrayed Then GetRealAccount = True: markreadOk = True
                End If
            End If
            
            If myFolder = Folder Then
                If AccountEnabled = vbUnchecked Then accountdenied = True
                If SignatureEnabled = vbUnchecked Then signaturedenied = True
                If FormatEnabled = vbUnchecked Then formatdenied = True
            End If
        
            If AccountEnabled = vbChecked And Not accountdenied And Not accountOk Then
                If GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_SZ, Account, "Account") Then
                    GetRealAccount = True: accountOk = True
                End If
            End If
            
            If SignatureEnabled = vbChecked And Not signaturedenied And Not signatureOk Then
                If GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_SZ, signature, "Signature") Then
                    GetRealAccount = True: signatureOk = True
                End If
            
                If GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, SignatureTop, "SignatureTop") Then
                    GetRealAccount = True: signaturetopOk = True
                End If
            End If
            
            If FormatEnabled = vbChecked And Not formatdenied And Not formatOK Then
                If GetRegValue(HKEY_CURRENT_USER, myGetFolder, REG_DWORD, Format, "Format") Then
                    GetRealAccount = True: formatOK = True
                End If
            End If
            
        End If
            
        If myFolder.Parent Is Nothing Then Exit Do
        If myFolder.Parent.Class <> olFolder Or myFolder.FolderPath = "\" Then Exit Do
        Set myFolder = myFolder.Parent
        myGetFolder = "Software\Matro\RealAccount\folders\" & GetRealAccountFolder(myFolder)
            
    Loop

End Function

Private Sub oAppNS_OptionsPagesAdd(ByVal Pages As Outlook.PropertyPages, ByVal Folder As Outlook.MAPIFolder)

    Call Log("OnOptionsPagesAdd", "called", LOG_DEBUG)
    
    If BetaExpired Then Exit Sub
    
    If Not oNewRealAccountPPage Is Nothing Then
        Set oNewRealAccountPPage = Nothing
    End If
    
    If Folder.DefaultItemType = olMailItem Then
        Call Log("OnOptionsPagesAdd", "attaching RealAccountPPage", LOG_DEBUG)
        If RunningIDE Then
            Set oNewRealAccountPPage = CreateObject("RealAccountPPage.p")
            oNewRealAccountPPage.FolderID = Folder.EntryID
            Call Log("OnOptionsPagesAdd", "RealAccountPPage.FolderID set", LOG_DEBUG)
            Pages.Add oNewRealAccountPPage
            Call Log("OnOptionsPagesAdd", "RealAccountPPage launched", LOG_DEBUG)
        Else
            Pages.Add "RealAccountPPage.p", "RealAccount"
            Call Log("OnOptionsPagesAdd", "RealAccountPPage launched", LOG_DEBUG)
            Set oNewRealAccountPPage = Pages.Item(Pages.Count)
            oNewRealAccountPPage.FolderID = Folder.EntryID
            Call Log("OnOptionsPagesAdd", "RealAccountPPage.FolderID set", LOG_DEBUG)
        End If
    End If
        
End Sub

Private Sub oMailItem_Send(Cancel As Boolean)

    Call SetMailItemDefaults(SetAccountOnly:=True)

End Sub

Private Sub oNewRealAccountPPage_OptionsOK()

    Call SetClickYes

End Sub

Private Sub oNewRealAccountPPage_PropertyChange(FolderID As String)

    Dim FolderName As String, myFolder As MAPIFolder
    Dim MarkRead&, dummy1$, dummy2$, dummy3&, dummy4&, Item, oItem As cItems
    
    Set myFolder = oAppNS.GetFolderFromID(FolderID)
    
    Call Log("OnNewRealAccountPPage_PropertyChange", "called for folder '" & myFolder.Name & "'", LOG_DEBUG)

    Call GetRealAccount(myFolder, dummy1, dummy2, dummy3, MarkRead, dummy4)
    If MarkRead = vbChecked Then
        Set oItem = New cItems
        Set oItem.MAPIFolder = myFolder
        On Error Resume Next
        oItems.Add oItem, FolderID
        On Error GoTo 0
        Call Log("OnNewRealAccountPPage_PropertyChange", "MarkRead for folder '" & myFolder.Name & "'", LOG_INFO)
    Else
        For Each Item In oItems
            If Item.MAPIFolder.EntryID = FolderID Then
                oItems.Remove FolderID
                Set Item = Nothing
                Call Log("OnNewRealAccountPPage_PropertyChange", "MarkRead removed for folder '" & myFolder.Name & "'", LOG_INFO)
                Exit For
            End If
        Next Item
    End If
    
End Sub

Sub GetAllMarkReadFolders()

    Dim myFolder As MAPIFolder, oItem As cItems
    Dim dummy1 As New Collection, dummy2$, dummy3$, dummy4&, dummy5&
    Dim Item, MarkRead As Long, errGetFolderFromID&, ItemsAdded&

    If BetaExpired Then Exit Sub

    On Error GoTo myError

    Call Log("GetAllMarkReadFolders", "Get MarkRead folders - start", LOG_DEBUG)
    If EnumRegKey(HKEY_CURRENT_USER, "Software\Matro\RealAccount\Folders", dummy1) Then
        For Each Item In dummy1

            On Error Resume Next
            Set myFolder = oAppNS.GetFolderFromID("00000000" & Item)
            If Err.Number > 0 Then errGetFolderFromID = Err.Number: Err.Number = 0
            On Error GoTo myError
            
            If errGetFolderFromID = 0 Then
                If Not myFolder Is Nothing Then
                    If GetRealAccount(myFolder, dummy2, dummy3, dummy4, MarkRead, dummy5) Then
                        If MarkRead = vbChecked Then
                            Set oItem = New cItems
                            Set oItem.MAPIFolder = myFolder
                            
                            On Error Resume Next
                            oItems.Add oItem, myFolder.EntryID
                            If Err > 0 Then
                                Call Log("GetAllMarkReadFolders", "Folder already present '" & myFolder.Name & "'", LOG_INFO)
                                Err = 0
                            Else
                                Call Log("GetAllMarkReadFolders", "MarkRead for folder '" & myFolder.Name & "'", LOG_INFO)
                                ItemsAdded = ItemsAdded + 1
                            End If
                            On Error GoTo myError
                        End If
                    Else
                        Call Log("GetAllMarkReadFolders", "GetRealAccount failed for folder '" & Item & "'", LOG_DEBUG)
                    End If
                Else
                    Call Log("GetAllMarkReadFolders", "Invalid folder '" & Item & "'", LOG_DEBUG)
                End If
            End If
            
            Set myFolder = Nothing: MarkRead = False
        Next Item
    End If
    Call Log("GetAllMarkReadFolders", "Completed: Added " & ItemsAdded & " Total " & oItems.Count, LOG_DEBUG)

    Exit Sub

myError:
    
    Log "GetAllMarkReadFolders", "error: (" & Err.Number & ") " & Err.Description, LOG_ERROR
    WorkInProgress = False

End Sub

Private Sub oTimer_Tick()

    MarkReadRetries = MarkReadRetries + 1
    If MarkReadRetries = 3 Then oTimer.Enabled = False
    
    Call Log("oTimer_Timer", "Calling GetAllMarkReadFolders attempt " & MarkReadRetries, LOG_DEBUG)
    Call GetAllMarkReadFolders

End Sub

Private Sub SetClickYes()

    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYes, "ClickYes") Then
        ClickYes = vbUnchecked
        If Not SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYes, "ClickYes") Then
            Call Log("SetClickYes", "SetRegValue(ClickYes) failed", LOG_ERROR)
        End If
    End If
    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYesMls, "ClickYesMls") Then
        ClickYesMls = 1000
        If Not SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYesMls, "ClickYesMls") Then
            Call Log("SetClickYes", "SetRegValue(ClickYesMls) failed", LOG_ERROR)
        End If
    End If

    Call CreateClickYesScript
    
End Sub

