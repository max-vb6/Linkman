VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAniTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private m_TimerID As Long
Private m_Interval As Long
Private m_Enabled As Boolean

Public Tag As Variant
Public Event Timer()

Public Property Get Interval() As Long
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal Value As Long)
    m_Interval = Value
    Enabled = m_Enabled
End Property

Public Property Get Enabled() As Boolean
    Interval = m_Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    If Value Then
        m_Enabled = Value
        Call StartTimer
    Else
        Call StopTimer
    End If
End Property

Private Function StartTimer() As Boolean
    If m_TimerID = 0 Then
        If m_Interval > 0 Then
            m_TimerID = SetTimer(0, 0, m_Interval, AddressOf TimeProc)
            If m_TimerID <> 0 Then
                TimerColl.Add ObjPtr(Me), "ID:" & m_TimerID
                StartTimer = True
            End If
        Else
            m_Enabled = True
        End If
    End If
End Function

Friend Sub PulseTimer()
    RaiseEvent Timer
End Sub

Private Sub StopTimer()
    If m_TimerID <> 0 Then
        KillTimer 0, m_TimerID
        TimerColl.Remove "ID:" & m_TimerID
        m_TimerID = 0
        m_Enabled = False
    End If
End Sub

Private Sub Class_Terminate()
    Call StopTimer
End Sub
