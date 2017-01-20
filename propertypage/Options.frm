VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RealAccount - Options"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "@13"
   Begin VB.TextBox txAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Tag             =   "@14"
      Text            =   "Options.frx":0000
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Tag             =   "@12"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   2640
      TabIndex        =   15
      Tag             =   "@11"
      Top             =   1920
      Width           =   2295
      Begin VB.Image lbLang 
         Height          =   360
         Index           =   1
         Left            =   120
         MouseIcon       =   "Options.frx":000E
         MousePointer    =   99  'Custom
         Picture         =   "Options.frx":0318
         Stretch         =   -1  'True
         ToolTipText     =   "Click here to set language to english"
         Top             =   660
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   1080
         Picture         =   "Options.frx":075A
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label lblLanguage 
         Caption         =   "English"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblLanguage 
         Caption         =   "Italiano"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Image lbLang 
         Height          =   360
         Index           =   0
         Left            =   120
         MouseIcon       =   "Options.frx":0CD2
         MousePointer    =   99  'Custom
         Picture         =   "Options.frx":0FDC
         Stretch         =   -1  'True
         ToolTipText     =   "Fare clic qui per impostare la lingua italiana"
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Tag             =   "@1"
      Top             =   120
      Width           =   4815
      Begin VB.VScrollBar vsClickYesMls 
         Enabled         =   0   'False
         Height          =   220
         Left            =   2610
         Max             =   9999
         Min             =   100
         SmallChange     =   100
         TabIndex        =   2
         Top             =   345
         Value           =   1000
         Width           =   175
      End
      Begin VB.TextBox txtClickYesMls 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   1
         Top             =   315
         Width           =   660
      End
      Begin VB.CheckBox chkClickYes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Tag             =   "@2"
         Top             =   345
         Width           =   975
      End
      Begin VB.Label lblDelay 
         AutoSize        =   -1  'True
         Caption         =   ":-)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Tag             =   "@3"
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "@5"
         Top             =   1350
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "matro"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Tag             =   "@4"
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblDelayMls 
         AutoSize        =   -1  'True
         Caption         =   "mls"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Tag             =   "@6"
      Top             =   1920
      Width           =   2415
      Begin VB.CheckBox chkLog 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Tag             =   "@10"
         Top             =   1095
         Width           =   1695
      End
      Begin VB.CheckBox chkLog 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "@7"
         Top             =   330
         Width           =   1695
      End
      Begin VB.CheckBox chkLog 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Tag             =   "@8"
         Top             =   585
         Width           =   1695
      End
      Begin VB.CheckBox chkLog 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Tag             =   "@9"
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    RealAccount v1.2
'    Code by Matro
'    Rome, Italy, 2002-2004
'    matro@realpopup.it
'
'    designed for MS Outlook 10 and later

Option Explicit

Dim Language As Long
Dim Log As String
Dim ClickYesMls As Long
Dim ClickYes As Long

Private Sub chkClickYes_Click()

    lblDelay.enabled = (chkClickYes.Value)
    lblDelayMls.enabled = lblDelay.enabled
    txtClickYesMls.enabled = lblDelay.enabled
    vsClickYesMls.enabled = lblDelay.enabled

End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Log = IIf(chkLog(0).Value = vbChecked, "E", "") & IIf(chkLog(1).Value = vbChecked, "W", "") & IIf(chkLog(2).Value = vbChecked, "I", "") & IIf(chkLog(3).Value = vbChecked, "D", "")
    ClickYesMls = txtClickYesMls
    ClickYes = chkClickYes.Value
    
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, Language, "Language")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, Log, "Log")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYes, "ClickYes")
    Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYesMls, "ClickYesMls")
       
    Unload Me

End Sub

Private Sub Form_Activate()

    chkLog(0).Value = IIf(InStr(Log, "E") > 0, vbChecked, vbUnchecked)
    chkLog(1).Value = IIf(InStr(Log, "W") > 0, vbChecked, vbUnchecked)
    chkLog(2).Value = IIf(InStr(Log, "I") > 0, vbChecked, vbUnchecked)
    chkLog(3).Value = IIf(InStr(Log, "D") > 0, vbChecked, vbUnchecked)
   
    txtClickYesMls = ClickYesMls
    chkClickYes.Value = ClickYes
    
    Call SetLanguage

End Sub

Private Sub Form_Load()

    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, Language, "Language") Then
        Language = 0
    End If
    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, Log, "Log") Then
        Log = "EW"
    End If
    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYes, "ClickYes") Then
        ClickYes = vbUnchecked
    End If
    If Not GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, ClickYesMls, "ClickYesMls") Then
        ClickYesMls = 1000
    End If

End Sub

Private Sub lbLang_Click(Index As Integer)

    Language = Index

    Call SetLanguage

End Sub

Private Sub txtClickYesMls_Change()
    
    If Val(txtClickYesMls.text) > vsClickYesMls.Min Then
        vsClickYesMls.Value = Val(txtClickYesMls.text)
    End If

End Sub

Private Sub txtClickYesMls_Click()

    txtClickYesMls.SelStart = 0
    txtClickYesMls.SelLength = Len(txtClickYesMls.text)

End Sub

Private Sub txtClickYesMls_KeyPress(KeyAscii As Integer)

    If InStr("0123456789", Chr$(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub vsClickYesMls_Change()

    txtClickYesMls.text = Format(vsClickYesMls.Value, "###")

End Sub

Private Sub SetLanguage()

    Dim Item, strings As New Collection, k&
    Dim RealAccountVersion$
        
    Call GetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_SZ, RealAccountVersion, "RealAccountPluginVersion")
        
    If Language = 1 Then
        strings.Add "ClickYes feature", "@1"
        strings.Add "Enabled", "@2"
        strings.Add "delay time:", "@3"
        strings.Add "The ""ClickYes"" feature suppresses the Security Guard which appears on Outlook XP SP3 (and possibly later) when a message is created and RealAccount is active.", "@4"
        strings.Add "Activate it at your own risk.", "@5"
        strings.Add "Log type", "@6"
        strings.Add "Errors", "@7"
        strings.Add "Warnings", "@8"
        strings.Add "Informations", "@9"
        strings.Add "Debug", "@10"
        strings.Add "Language", "@11"
        strings.Add "Cancel", "@12"
        strings.Add "RealAccount version " & App.Major & "." & App.Minor & vbCrLf & _
            "Developed by Matro, Rome (Italy)" & vbCrLf & vbCrLf & _
            "This plugin for MS Outlook is freeware:" & vbCrLf & _
            "use and distribute as you want!" & vbCrLf & vbCrLf & _
            "RealAccount homepage:" & vbCrLf & _
            "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
            "Send feedbacks to:" & vbCrLf & "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
            "RealAccount addin version " & RealAccountVersion & vbCrLf & _
            "RealAccount page version " & App.Major & "." & App.Minor & " build " & Format(App.Revision, "000"), "@14"
        txtClickYesMls.Left = 2160
        Me.Caption = "RealAccount - Options"
    Else
        strings.Add "Funzione ClickYes", "@1"
        strings.Add "Attiva", "@2"
        strings.Add "ritardo:", "@3"
        strings.Add "La funzione ""ClickYes"" elimina il Security Guard presente su Outlook XP SP3 (e versioni successive) che compare con RealAccount attivo quando si crea un messaggio.", "@4"
        strings.Add "Attivare la funzione a proprio rischio.", "@5"
        strings.Add "Tipo di Log", "@6"
        strings.Add "Errori", "@7"
        strings.Add "Avvertimenti", "@8"
        strings.Add "Informazioni", "@9"
        strings.Add "Debug", "@10"
        strings.Add "Linguaggio", "@11"
        strings.Add "Annulla", "@12"
        strings.Add "RealAccount versione " & App.Major & "." & App.Minor & vbCrLf & _
            "Sviluppato da Matro, Roma (Italia)" & vbCrLf & vbCrLf & _
            "Questo plugin per MS Outlook è freeware:" & vbCrLf & _
            "usalo e distribuiscilo quanto vuoi!" & vbCrLf & vbCrLf & _
            "RealAccount homepage:" & vbCrLf & _
            "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
            "Invia commenti a:" & vbCrLf & "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
            "RealAccount addin versione " & RealAccountVersion & vbCrLf & _
            "RealAccount page versione " & App.Major & "." & App.Minor & " build " & Format(App.Revision, "000"), "@14"
        txtClickYesMls.Left = 1920
        Me.Caption = "RealAccount - Opzioni"
    End If
    
    vsClickYesMls.Left = txtClickYesMls.Left + 450
    lblDelayMls.Left = txtClickYesMls.Left + 720
    
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

    Refresh

End Sub


