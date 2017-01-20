VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWelcome 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CheckBox chkDoNotShow 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmSplash.frx":4D5A
      Top             =   120
      Width           =   3390
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public language As Long
Public BetaShowMsg As Boolean
Public BetaExpired As Boolean

Private Sub Command1_Click()

    If chkDoNotShow.Value = vbChecked Then
        If BetaShowMsg Then
            Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, CLng(App.Revision), "WelcomeBeta")
        Else
            Call SetRegValue(HKEY_CURRENT_USER, "Software\Matro\RealAccount", REG_DWORD, 1, "Welcome")
        End If
    End If

    Unload Me

End Sub

Private Sub Form_Activate()

    Dim msg As String

    If BetaExpired Then chkDoNotShow.Visible = False

    If language = 1 Then
        If BetaShowMsg Then
            If BetaExpired Then
                msg = "THIS BETA HAS EXPIRED!" & vbCrLf & vbCrLf
            End If
            msg = msg & "This is a beta release. " & vbCrLf & vbCrLf & _
                "RealAccount addin version " & GetVersion() & vbCrLf & vbCrLf & _
                "Purpose of beta builds is to test specific features for " & _
                "bugs and compatibility. Read and join the RealAccount mailing list at " & _
                "http://groups.yahoo.com/group/realaccount for support about betas." & vbCrLf & vbCrLf & _
                "This beta will expire on " & Format(APP_BETA_DAY & "/" & APP_BETA_MONTH & "/" & APP_BETA_YEAR, "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & _
                "RealAccount homepage:" & vbCrLf & _
                "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
                "Send feedbacks to:" & vbCrLf & _
                "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
                "Read on the website for info about support."
            Me.Caption = "RealAccount BETA release"
        Else
            msg = "Welcome to RealAccount!" & vbCrLf & vbCrLf & _
                "RealAccount is active as a plug-in for Microsoft Outlook. " & vbCrLf & vbCrLf & _
                "As default, it doesn't affect your normal operation in any way; " & _
                "to enable its features, right click on any email folder - such as " & _
                "Inbox - and select Properties, then choose the RealAccount tab." & vbCrLf & vbCrLf & _
                "RealAccount homepage:" & vbCrLf & _
                "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
                "Send feedbacks to:" & vbCrLf & _
                "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
                "Read on the website for info and support." & vbCrLf & vbCrLf & _
                "This plugin for MS Outlook is freeware:" & vbCrLf & _
                "use and distribute as you want!" & vbCrLf & vbCrLf & _
                "matro :)"
            Me.Caption = "About RealAccount"
        End If
        chkDoNotShow.Caption = "Do not show this message anymore"
    Else
        If BetaShowMsg Then
            If BetaExpired Then
                msg = "QUESTA VERSIONE BETA E' SCADUTA!" & vbCrLf & vbCrLf
            End If
            msg = msg & "Questa è una versione beta. " & vbCrLf & vbCrLf & _
                "RealAccount addin versione " & GetVersion() & vbCrLf & vbCrLf & _
                "Lo scopo delle build beta è quello di testare specifiche funzionalità " & _
                "alla ricerca di bachi e problemi di compatibilità. Leggi e sottoscrivi la RealAccount mailing list all'indirizzo " & _
                "http://groups.yahoo.com/group/realaccount per il supporto delle beta." & vbCrLf & vbCrLf & _
                "Questa versione beta scade il " & Format(APP_BETA_DAY & "/" & APP_BETA_MONTH & "/" & APP_BETA_YEAR, "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & _
                "RealAccount homepage:" & vbCrLf & _
                "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
                "Invia commenti e suggerimenti a:" & vbCrLf & _
                "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
                "Visita il sito web per informazioni sul supporto di RealAccount."
            Me.Caption = "RealAccount release BETA"
        Else
            msg = "Benvenuto in RealAccount!" & vbCrLf & vbCrLf & _
                "RealAccount è un plug-in attivo in Microsoft Outlook. " & vbCrLf & vbCrLf & _
                "Con le impostazioni predefinite, RealAccount non interferisce sulle normali attività in alcun modo; " & _
                "per abilitare le sue funzionalità, fare clic col tasto destro su una qualunque cartella di messaggi - come " & _
                "Posta in arrivo - e selezionare Proprietà, quindi scegliere il riquadro RealAccount." & vbCrLf & vbCrLf & _
                "RealAccount homepage:" & vbCrLf & _
                "http://www.realpopup.it/realaccount" & vbCrLf & vbCrLf & _
                "Invia commenti e suggerimenti a:" & vbCrLf & _
                "realaccount@realpopup.it" & vbCrLf & vbCrLf & _
                "Visita il sito web per informazioni ed il supporto di RealAccount." & vbCrLf & vbCrLf & _
                "Questo plug-in per MS Outlook è freeware:" & vbCrLf & _
                "usalo e distribuiscilo come vuoi!" & vbCrLf & vbCrLf & _
                "matro :)"
            Me.Caption = "Informazioni su RealAccount"
        End If
        chkDoNotShow.Caption = "Non visualizzare questo messaggio in futuro"
    End If
    
    txtWelcome.Text = msg

    Me.SetFocus
    
End Sub

