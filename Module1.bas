Attribute VB_Name = "Module1"
Option Explicit
DefLng A-Z

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Type SIZE
  cx As Long
  cy As Long
End Type

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long
Declare Function GetTextExtentPoint32 _
Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc _
As Long, ByVal lpsz As String, ByVal _
cbString As Long, lpSize As SIZE) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const WM_SETREDRAW = &HB&
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const LB_GETITEMRECT = &H198
Public Const LB_ERR = (-1)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46

Public Function Max(ByVal param1 As Long, _
ByVal param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else _
     Max = param2
End Function


