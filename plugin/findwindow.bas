Attribute VB_Name = "findwindow"
Option Explicit

Declare Function SetFocusAPI Lib "user32" Alias "SetForegroundWindow" _
    (ByVal hwnd As Long) As Long
   Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal wCmd As Long) As Long
   Declare Function GetDesktopWindow Lib "user32" () As Long
   Declare Function GetWindowLW Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
   Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
     ByVal nMaxCount As Long) As Long
   Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) _
     As Long

   Public Const GWL_ID = (-12)
   Public Const GW_HWNDNEXT = 2
   Public Const GW_CHILD = 5
   'FindWindowLike
   ' - Finds the window handles of the windows matching the specified
   '   parameters
   '
   'hwndArray()
   ' - An integer array used to return the window handles
   '
   'hWndStart
   ' - The handle of the window to search under.
   ' - The routine searches through all of this window's children and their
   '   children recursively.
   ' - If hWndStart = 0 then the routine searches through all windows.
   '
   'WindowText
   ' - The pattern used with the Like operator to compare window's text.
   '
   'ClassName
   ' - The pattern used with the Like operator to compare window's class
   '   name.
   '
   'ID
   ' - A child ID number used to identify a window.
   ' - Can be a decimal number or a hex string.
   ' - Prefix hex strings with "&H" or an error will occur.
   ' - To ignore the ID pass the Visual Basic Null function.
   '
   'Returns
   ' - The number of windows that matched the parameters.
   ' - Also returns the window handles in hWndArray()
   '
   '----------------------------------------------------------------------
   Public Function FindWindowLike(hWndArray() As Long, ByVal hWndStart As Long, _
    WindowText As String, Classname As String, ID) As Long
   Dim hwnd As Long
   Dim r As Long
   ' Hold the level of recursion:
   Static level As Long
   ' Hold the number of matching windows:
   Static iFound As Long
   
   Dim sWindowText As String
   Dim sClassname As String
   Dim sID
   ' Initialize if necessary:
   If level = 0 Then
   iFound = 0
   ReDim hWndArray(0 To 0)
   If hWndStart = 0 Then hWndStart = GetDesktopWindow()
   End If
   ' Increase recursion counter:
   level = level + 1
   ' Get first child window:
   hwnd = GetWindow(hWndStart, GW_CHILD)
   Do Until hwnd = 0
   DoEvents ' Not necessary
   ' Search children by recursion:
   r = FindWindowLike(hWndArray(), hwnd, WindowText, Classname, ID)
   ' Get the window text and class name:
   sWindowText = Space(255)
   r = GetWindowText(hwnd, sWindowText, 255)
   sWindowText = Left(sWindowText, r)
   sClassname = Space(255)
   r = GetClassName(hwnd, sClassname, 255)
   sClassname = Left(sClassname, r)
   ' If window is a child get the ID:
   If GetParent(hwnd) <> 0 Then
   r = GetWindowLW(hwnd, GWL_ID)
   sID = CLng("&H" & Hex(r))
   Else
   sID = Null
   End If
   ' Check that window matches the search parameters:
   If sWindowText Like WindowText And sClassname Like Classname Then
   If IsNull(ID) Then
   ' If find a match, increment counter and
   '  add handle to array:
   iFound = iFound + 1
   ReDim Preserve hWndArray(0 To iFound)
   hWndArray(iFound) = hwnd
   ElseIf Not IsNull(sID) Then
   If CLng(sID) = CLng(ID) Then
   ' If find a match increment counter and
   '  add handle to array:
   iFound = iFound + 1
   ReDim Preserve hWndArray(0 To iFound)
   hWndArray(iFound) = hwnd
   End If
   End If
   End If
   ' Get next child window:
   hwnd = GetWindow(hwnd, GW_HWNDNEXT)
   Loop
   ' Decrement recursion counter:
   level = level - 1
   ' Return the number of windows found:
   FindWindowLike = iFound
   End Function

