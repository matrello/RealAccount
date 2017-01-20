Attribute VB_Name = "utility"
'    RealAccount v1.2
'    Code by Matro
'    Rome, Italy, 2002-2004
'    matro@realpopup.it
'
'    designed for MS Outlook 10 and later

Option Explicit

Public Const APP_BETA = False
Public Const APP_BETA_YEAR = 2004
Public Const APP_BETA_MONTH = 6
Public Const APP_BETA_DAY = 30

Public Const LOG_ERROR = 0
Public Const LOG_WARNING = 1
Public Const LOG_INFO = 2
Public Const LOG_STRONGINFO = 3
Public Const LOG_DEBUG = 4

Public Const HKEY_CURRENT_USER = &H80000001

Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_SZ = 1
Public Const ERROR_NO_MORE_ITEMS = 259&

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_HISTORY = &H22

Private Const MAX_PATH = 260

Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbname As Long) As Long
Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, ByVal lpbData As String, ByVal cbData As Long) As Long
Declare Function OSRegSetDWordEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Long, ByVal cbData As Long) As Long

Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWnd As Long, ByVal nFolder As Long, Pidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal Pidl As Long, ByVal FolderPath As String) As Long

Declare Function OSGetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ClickYesVBSPath As String
Public ClickYes As Long, ClickYesMls As Long
Public UseEntryID As Boolean
Public LogVerbose As String

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Public Function IsNT() As Boolean

  Dim OSInfo As OSVERSIONINFO, PId As String, ret&
  
  OSInfo.dwOSVersionInfoSize = Len(OSInfo)
  ret& = OSGetVersionEx(OSInfo)
  If ret& = 0 Then
    Call Log("IsNT", "GetVersionEx() returned error.", LOG_WARNING)
  End If
  
  If OSInfo.dwPlatformId = 2 Then IsNT = True

End Function

Public Sub Log(fn As String, msg As String, Optional tipo, Optional start)

    Dim Out As String, eventType As enmLogType

    If IsMissing(tipo) Then tipo = LOG_INFO
    If Not IsMissing(start) Then App.StartLogging start, vbLogAuto

    If LogVerbose = "" Then
        Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, LogVerbose, "Log")
    End If

    Select Case tipo
        Case LOG_ERROR
            Out = "E": If InStr(LogVerbose, "E") > 0 Then eventType = LogError
        Case LOG_WARNING
            Out = "W": If InStr(LogVerbose, "W") > 0 Then eventType = LogWarning
        Case LOG_INFO
            Out = "I": If InStr(LogVerbose, "I") > 0 Then eventType = LogInfo
        Case LOG_DEBUG
            Out = "D": If InStr(LogVerbose, "D") > 0 Then eventType = LogInfo
        Case LOG_STRONGINFO
            Out = "S": eventType = 4
    End Select

    If eventType > 0 Then
        Out = Out & Format(Now, "hh:nn:ss") & " " & fn & ": " & msg
        Debug.Print Out
        Call LogErrorToEventViewer(Out, eventType)
    End If

End Sub

Function EnumRegKey(hKey As Long, sSubKey As String, items As Collection) As Boolean

    Dim KeyHandle&, maxChar&
    Dim valBuffer As String * 1024, sBuffer$
    Dim ret%, Index%, pos&
    
    ret = OSRegOpenKey(hKey, sSubKey, KeyHandle)
    If ret <> 0 Then Exit Function

    Do
        maxChar = 1024: valBuffer = String$(1024, 0)
        ret = OSRegEnumKey(KeyHandle, Index, valBuffer, maxChar)
        If ret <> 0 Then Exit Do
        pos = InStr(1, valBuffer, Chr$(0))
        If pos = 0 Then items.Add valBuffer Else items.Add Left$(valBuffer, pos - 1)
        Index = Index + 1
    Loop

    If ret <> ERROR_NO_MORE_ITEMS Then Exit Function

    EnumRegKey = True

End Function

Function SetRegValue(hKey As Long, sSubKey As String, sType As Long, sVal, Optional Item) As Boolean

    Dim KeyHandle&, lenBuffer&
    Dim valBuffer As String * 1024
    Dim ret%
    
    ret = OSRegOpenKey(hKey, sSubKey, KeyHandle)
    If ret <> 0 Then Exit Function
    
    Select Case sType
        Case REG_SZ
            valBuffer = Trim(sVal): lenBuffer = Len(Trim(sVal))
            ret = OSRegSetValueEx(KeyHandle, Item, 0, sType, ByVal valBuffer, lenBuffer)
        Case REG_DWORD
            ret = OSRegSetDWordEx(KeyHandle, Item, 0, sType, sVal, 4)
    End Select
    If ret <> 0 Then Exit Function
    OSRegCloseKey (KeyHandle)
    SetRegValue = True

End Function

Function GetRegValue(hKey As Long, sSubKey As String, sType As Long, sVal, Optional Item) As Boolean

    Dim KeyHandle&, lenBuffer&
    Dim valBuffer$, longBuffer&
    Dim ret%, pos&
    
    ret = OSRegOpenKey(hKey, sSubKey, KeyHandle)
    If ret <> 0 Then Exit Function
    lenBuffer = 1025: valBuffer = Space(lenBuffer)
    Select Case sType
        Case REG_BINARY
            ret = OSRegQueryValueEx(KeyHandle, Item, 0, sType, ByVal valBuffer, lenBuffer)
            If ret <> 0 Then Exit Function
            sVal = StrConv(valBuffer, vbFromUnicode)
            pos = InStr(1, sVal, Chr$(0))
            If pos > 0 Then sVal = Left$(sVal, pos - 1)
        Case REG_SZ
            ret = OSRegQueryValueEx(KeyHandle, Item, 0, sType, ByVal valBuffer, lenBuffer)
            If ret <> 0 Then Exit Function
            pos = InStr(1, valBuffer, Chr$(0))
            If pos > 0 Then valBuffer = Left$(valBuffer, pos - 1)
            sVal = Trim(valBuffer)
        Case REG_DWORD
            lenBuffer = 4
            ret = OSRegQueryValueEx(KeyHandle, Item, 0, sType, longBuffer, lenBuffer)
            If ret <> 0 Then Exit Function
            sVal = longBuffer
    End Select
    OSRegCloseKey (KeyHandle)
    GetRegValue = True

End Function

Function EnumSignatures(signatures As Collection)

    Dim sigpath$, sigfolder$, ok As Boolean

    ok = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\11.0\Common\General", REG_SZ, sigfolder, "Signatures")
    If Not ok Or Len(sigfolder) = 0 Then ok = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\10.0\Common\General", REG_SZ, sigfolder, "Signatures")

    Call Log("EnumSignatures", "signatures key is '" & sigfolder & "'", LOG_DEBUG)

    If Len(sigfolder) = 0 Then sigfolder = "Signatures"

    sigpath = SpecialFolderPath(CSIDL_APPDATA)
    sigpath = sigpath & IIf(Right$(sigpath, 1) <> "\", "\", "") & "Microsoft\" & sigfolder
    
    Call Log("EnumSignatures", "signatures folder is '" & sigpath & "'", LOG_DEBUG)
    
    sigfolder = Dir(sigpath & "\*.*", vbNormal)
    Do While sigfolder <> ""
        If (InStr(1, sigfolder, ".htm", vbTextCompare) > 0 Or InStr(1, sigfolder, ".txt", vbTextCompare) > 0) And Len(sigfolder) > 4 Then
            sigfolder = Replace$(sigfolder, ".htm", "")
            sigfolder = Trim$(Replace$(sigfolder, ".txt", ""))
            On Error Resume Next
            sigfolder = signatures(sigfolder)
            If Err > 0 Then
                signatures.Add sigfolder, sigfolder
            Call Log("EnumSignatures", "added signature '" & sigfolder & "'", LOG_DEBUG)
            End If
            On Error GoTo 0
        End If
    
       sigfolder = Dir
    Loop

End Function

Function GetSignature(signature As String, tipo As OlBodyFormat) As String

    Dim sigpath$, sigfolder$, sigbin() As Byte, h%, pos&, ok As Boolean

    If signature = "" Then Exit Function
    
    On Error GoTo myError
    
    ok = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\11.0\Common\General", REG_SZ, sigfolder, "Signatures")
    If Not ok Or Len(sigfolder) = 0 Then ok = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Office\10.0\Common\General", REG_SZ, sigfolder, "Signatures")

    Call Log("GetSignature", "signatures key is '" & sigfolder & "'", LOG_DEBUG)

    If Len(sigfolder) = 0 Then sigfolder = "Signatures"
    
    sigpath = SpecialFolderPath(CSIDL_APPDATA)
    sigpath = sigpath & IIf(Right$(sigpath, 1) <> "\", "\", "") & "Microsoft\" & sigfolder & IIf(Right$(sigpath, 1) <> "\", "\", "")
    sigpath = sigpath & signature
    
    Call Log("GetSignature", "signatures folder is '" & sigpath & "'", LOG_DEBUG)
        
    Select Case tipo
        Case olFormatHTML, olFormatRichText
            sigpath = sigpath & ".htm"
            h = FreeFile
            Open sigpath For Binary Access Read As #h
            sigbin = Space$(LOF(h))
            Get h, , sigbin
            Close #h
            GetSignature = StrConv(sigbin, vbUnicode)
            If Len(GetSignature) < 3 Then Exit Function
            pos = InStr(1, GetSignature, "<body>", vbTextCompare)
            If pos > 0 Then GetSignature = Right$(GetSignature, Len(GetSignature) - pos - 5)
            pos = InStr(1, GetSignature, "</body>", vbTextCompare)
            If pos > 0 Then GetSignature = Left$(GetSignature, pos - 1)
            
        Case olFormatPlain
            sigpath = sigpath & ".txt"
            h = FreeFile
            Open sigpath For Binary Access Read As #h
            sigbin = Space$(LOF(h))
            Get h, , sigbin
            Close #h
            GetSignature = CStr(sigbin):
            If Len(GetSignature) < 3 Then Exit Function
            GetSignature = Right$(GetSignature, Len(GetSignature) - 1)
            h = 4
            Do While h <= Len(GetSignature)
                If Asc(Mid$(GetSignature, h, 1)) = 0 Then
                    GetSignature = Left$(GetSignature, h - 1)
                    Exit Do
                End If
                h = h + 1
            Loop
    End Select

    Exit Function
    
myError:
    
    Log "GetSignature", "error: (" & Err.Number & ") " & Err.Description, LOG_ERROR

End Function

Public Function SpecialFolderPath(CSIDL As Long) As String
    
    Dim Pidl As Long
    Dim sFolderPath As String
    
    If SHGetSpecialFolderLocation(0, CSIDL, Pidl) = 0 Then
        sFolderPath = String(MAX_PATH, 0)
        If SHGetPathFromIDList(Pidl, ByVal sFolderPath) Then
            SpecialFolderPath = Left(sFolderPath, InStr(1, sFolderPath, Chr(0)) - 1)
        End If
    End If

End Function

Public Function GetRealAccountFolder(Folder As MAPIFolder) As String

    If UseEntryID Then
        GetRealAccountFolder = Right$(Folder.EntryID, Len(Folder.EntryID) - 8)
    Else
        GetRealAccountFolder = Folder.Name
    End If

End Function

Public Function RunningIDE() As Boolean
    
    Debug.Assert Not TestIDE(RunningIDE)

End Function

Private Function TestIDE(Test As Boolean) As Boolean
    
    Test = True

End Function

Public Function GetVersion() As String

    GetVersion = App.Major & "." & App.Minor & " build " & Format(App.Revision, "000")
    If APP_BETA Then GetVersion = GetVersion & " BETA"

End Function

Public Function CreateClickYesScript() As Boolean

    Dim fs As New Scripting.FileSystemObject, f As Scripting.TextStream
    Dim RealAccountVersion As String
    
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, RealAccountVersion, "RealAccountPluginVersion")
    ClickYesVBSPath = fs.BuildPath(fs.GetSpecialFolder(TemporaryFolder), "RealAccountClickYes.vbs")
    
    On Error Resume Next
    Set f = fs.CreateTextFile(ClickYesVBSPath, True)
    f.WriteLine ("' this is part of RealAccount v" & RealAccountVersion)
    f.WriteLine ("' RealAccount is a freeware plugin for MS Outlook.")
    f.WriteLine ("'")
    f.WriteLine ("' control this script activation through RealAccount options;")
    f.WriteLine ("' to access the options, right click on any mailitem folder,")
    f.WriteLine ("' activate RealAccount tab and press Options button.")
    f.WriteLine ("'")
    f.WriteLine ("set sh=WScript.CreateObject(""WScript.Shell"")" & vbCrLf & "WScript.Sleep(" & (ClickYesMls / 2) & ")")
    f.WriteLine ("activated = False: dt = 100: tw = " & ClickYesMls)
    f.WriteLine ("Do While (Not activated And tw > 0)" & vbCrLf & "activated = sh.AppActivate(""Microsoft Outlook"")" & vbCrLf & "WScript.Sleep (dt): tw = tw - dt" & vbCrLf & "Loop")
    f.WriteLine ("If tw >= dt Then" & vbCrLf & "WScript.Sleep(dt): sh.SendKeys(""{TAB 3}{ENTER}"")" & vbCrLf & "End If")
    f.Close
    
    If Err = 0 Then
        CreateClickYesScript = True
        Call Log("CreateClickYesScript", "ClickYes script created at " & ClickYesVBSPath, LOG_DEBUG)
    Else
        Call Log("CreateClickYesScript", "could not create ClickYes script, error is (" & Err.Number & ") " & Err.Description, LOG_DEBUG)
        ClickYesVBSPath = ""
        Err = 0
    End If
    
    Set fs = Nothing

End Function
