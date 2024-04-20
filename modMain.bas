Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Function PCName() As String
    Dim sName As String * 255
    GetComputerName sName, 255
    PCName = Trim(Replace(sName, Chr(0), ""))
End Function

Public Sub Sleep(ByVal dwMilliseconds As Long)
    Dim SaveTime As Long
    Dim NowTime As Long
    Dim IsWait As Long
    IsWait = 0
    SaveTime = GetTickCount
    Do
       DoEvents
       NowTime = GetTickCount
       If NowTime - SaveTime >= dwMilliseconds Then
          IsWait = 1
       End If
    Loop While IsWait = 0
End Sub
