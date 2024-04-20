Attribute VB_Name = "mAniControl"
Option Explicit

Public Type AniParams
    ToValue As Single
    Speed As Single
    K As Single
    Attn As Single        'Attenuation
    Mode As AnimationMode
End Type

Public Enum AnimationMode
    Deceleration = 0
    Uniform = 1
    Elasticity = 2
End Enum

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public TimerColl As New VBA.Collection

Public Sub TimeProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim Timer As cAniTimer, lpTimer As Long
    lpTimer = TimerColl("ID:" & idEvent)
    CopyMemory Timer, lpTimer, 4&
    Timer.PulseTimer
    CopyMemory Timer, 0&, 4&
End Sub
