Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function CreateCursor Lib "user32.dll" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public m_sngMin As Single
Public m_sngMax As Single
Public m_sngValue As Single
Public m_bEnabled As Boolean
Public bReset As Boolean
Public g_hInstance As Long

Public m_bActive As Boolean

Public lOldCursor As Long
Public bOldCursorSet As Boolean

Sub Main()

End Sub
