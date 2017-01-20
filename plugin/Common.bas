Attribute VB_Name = "modCommon"
' ************************************************************
' Copyright © 1996-2001 Slightly Tilted Software
' All rights reserved
' You're absolutely free to use these resources within your
'     own applications, but you may not redistribute them
'     (as source) in any manner whatsoever, whether for profit
'     or not.
' The only legitimate source for the original source code is
'     at the VBPJ site and my own web site at:
'     http://www.SlightlyTiltedSoftware.com
' ************************************************************

' ---------------------------------------------
' Module    : EVENTLOG.BAS
' By        : L.J. Johnson       Date: 04-28-2001
' Comments  : Contains only ReturnApiErrString()
' ---------------------------------------------
Option Explicit

' ---------------------------------------------
' Used to get error messages directly from the
'    system instead of hard-coding them
' ---------------------------------------------
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200&
Private Const FORMAT_MESSAGE_FROM_HMODULE    As Long = &H800&
Private Const LOAD_LIBRARY_AS_DATAFILE       As Long = 2&

' ---------------------------------------------
' Custom error messages for this app
' ---------------------------------------------
Public Const ERR_REG_EVENT_SOURCE      As Long = 9001&
Public Const ERR_REPORT_EVENT          As Long = 9002&
Public Const ERR_DEREG_EVENT_SOURCE    As Long = 9003&
Public Const ERR_NO_CREATE_KEY         As Long = 9004&
Public Const ERR_NO_OPEN_KEY           As Long = 9005&
Public Const ERR_NO_SET_FIRST_VALUE    As Long = 9006&
Public Const ERR_NO_SET_SECOND_VALUE   As Long = 9007&
Public Const ERR_NO_CLOSE_KEY          As Long = 9008&

' ---------------------------------------------
' Status Codes
' ---------------------------------------------
Private Const INVALID_HANDLE_VALUE           As Long = -1&
Public Const ERROR_SUCCESS                   As Long = 0&

' ---------------------------------------------
' Upper and lower bounds of network errors
' ---------------------------------------------
Private Const NERR_BASE                      As Long = 2100&
Private Const MAX_NERR                       As Long = NERR_BASE + 899&

' ---------------------------------------------
' Upper and lower bounds of Internet errors
' ---------------------------------------------
Private Const INTERNET_ERROR_BASE            As Long = 12000&
Private Const INTERNET_ERROR_LAST            As Long = INTERNET_ERROR_BASE + 171&

Private Declare Function FormatMessage _
   Lib "kernel32" Alias "FormatMessageA" _
   (ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long
Private Declare Function LoadLibraryEx _
   Lib "kernel32" Alias "LoadLibraryExA" _
   (ByVal lpLibFileName As String, _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary _
   Lib "kernel32" _
   (ByVal hLibModule As Long) As Long

' *******************************************************
' Routine Name : (PUBLIC in MODULE) Function ReturnApiErrString
' Written By   : L.J. Johnson
' Programmer   : L.J. Johnson [Slightly Tilted Software]
' Date Writen  : 01/16/1999 -- 12:56:46
' Inputs       : ErrorCode:Long - Number returned from API error
' Outputs      : N/A
' Description  : Function returns the error string
'              : The original code appeared in Keith Pleas
'              :     article in VBPJ, April 1996 (OLE Expert
'              :     column).  Thanks, Keith.
' *******************************************************
Public Function ReturnApiErrString(ErrorCode As Long) As String
On Error Resume Next                   ' Don't accept an error here
   Dim p_strBuffer                     As String
   Dim p_lngHwndModule                 As Long
   Dim p_lngFlags                      As Long
   
   ' ------------------------------------------
   ' Separate handling for network errors
   ' netmsg.dll
   ' ------------------------------------------
   If ErrorCode >= NERR_BASE And _
      ErrorCode <= MAX_NERR Then
      
      p_lngHwndModule = LoadLibraryEx(lpLibFileName:="netmsg.dll", _
                        hFile:=0&, _
                        dwFlags:=LOAD_LIBRARY_AS_DATAFILE)
      
      If p_lngHwndModule <> 0 Then
      
         p_lngFlags = FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS Or _
                      FORMAT_MESSAGE_FROM_HMODULE
                      
         ' ------------------------------------
         ' Allocate the string, then get the
         '     system to tell us the error
         '     message associated with this error number
         ' ------------------------------------
         p_strBuffer = String(256, 0)
         FormatMessage dwFlags:=p_lngFlags, _
                       lpSource:=ByVal p_lngHwndModule, _
                       dwMessageId:=ErrorCode, _
                       dwLanguageId:=0&, _
                       lpBuffer:=p_strBuffer, _
                       nSize:=Len(p_strBuffer), _
                       Arguments:=ByVal 0&
      
         ' ------------------------------------
         ' Strip the last null, then the last
         '     CrLf pair if it exists
         ' ------------------------------------
         p_strBuffer = Left(p_strBuffer, InStr(p_strBuffer, vbNullChar) - 1)
         If Right$(p_strBuffer, 2) = Chr$(13) & Chr$(10) Then
            p_strBuffer = Mid$(p_strBuffer, 1, Len(p_strBuffer) - 2)
         End If
         
         FreeLibrary hLibModule:=p_lngHwndModule
      End If
   
   ' ------------------------------------------
   ' Separate handling for Wininet error
   ' Wininet.dll
   ' ------------------------------------------
   ElseIf ErrorCode >= INTERNET_ERROR_BASE And _
      ErrorCode <= INTERNET_ERROR_LAST Then
      
      ' ---------------------------------------
      ' Load the library
      ' ---------------------------------------
      p_lngHwndModule = LoadLibraryEx(lpLibFileName:="Wininet.dll", _
                        hFile:=0&, _
                        dwFlags:=LOAD_LIBRARY_AS_DATAFILE)
      
      If p_lngHwndModule <> 0 Then
      
         p_lngFlags = FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS Or _
                      FORMAT_MESSAGE_FROM_HMODULE
                      
         ' ------------------------------------
         ' Allocate the string, then get the
         '     system to tell us the error
         '     message associated with this error number
         ' ------------------------------------
         p_strBuffer = String(256, 0)
         FormatMessage dwFlags:=p_lngFlags, _
                       lpSource:=ByVal p_lngHwndModule, _
                       dwMessageId:=ErrorCode, _
                       dwLanguageId:=0&, _
                       lpBuffer:=p_strBuffer, _
                       nSize:=Len(p_strBuffer), _
                       Arguments:=ByVal 0&
      
         ' ------------------------------------
         ' Strip the last null, then the last
         '     CrLf pair if it exists
         ' ------------------------------------
         p_strBuffer = Left(p_strBuffer, InStr(p_strBuffer, vbNullChar) - 1)
         If Right$(p_strBuffer, 2) = Chr$(13) & Chr$(10) Then
            p_strBuffer = Mid$(p_strBuffer, 1, Len(p_strBuffer) - 2)
         End If
         
         FreeLibrary hLibModule:=p_lngHwndModule
      End If
   
   ' ------------------------------------------
   ' Wasn't Wininet or NetMsg, so do the standard
   '     API error look-up
   ' ------------------------------------------
   Else
      ' ---------------------------------------
      ' Allocate the string, then get the system
      '     to tell us the error message associated
      '     with this error number
      ' ---------------------------------------
      p_strBuffer = String(256, 0)
      p_lngFlags = FORMAT_MESSAGE_FROM_SYSTEM Or _
                   FORMAT_MESSAGE_IGNORE_INSERTS
      
      FormatMessage dwFlags:=p_lngFlags, _
                    lpSource:=ByVal 0&, _
                    dwMessageId:=ErrorCode, _
                    dwLanguageId:=0&, _
                    lpBuffer:=p_strBuffer, _
                    nSize:=Len(p_strBuffer), _
                    Arguments:=ByVal 0&
   
      ' ---------------------------------------
      ' Strip the last null, then the last CrLf
      '     pair if it exists
      ' ------------------------------------------
      p_strBuffer = Left(p_strBuffer, InStr(p_strBuffer, vbNullChar) - 1)
      If Right$(p_strBuffer, 2) = Chr$(13) & Chr$(10) Then
         p_strBuffer = Mid$(p_strBuffer, 1, Len(p_strBuffer) - 2)
      End If
   End If
   
   ' ------------------------------------------
   ' Set the return value
   ' ------------------------------------------
   ReturnApiErrString = p_strBuffer

End Function
