VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAniControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event AnimationEnded()
Event AnimationProgress()

Dim WithEvents tmrAni As cAniTimer
Attribute tmrAni.VB_VarHelpID = -1
Dim oObj As Object, sAttr As String, apPara As AniParams

Private Sub Class_Initialize()
    Set tmrAni = New cAniTimer
    tmrAni.Interval = 1
End Sub

Private Sub Class_Terminate()
    Set tmrAni = Nothing
End Sub

Friend Sub DoAnimation(Object As Object, AttrName As String, Param As AniParams)
    tmrAni.Enabled = False
    Set oObj = Object
    sAttr = AttrName
    apPara = Param
    DoEvents
    tmrAni.Enabled = True
End Sub

Public Sub StopAnimation()
    tmrAni.Enabled = False
    RaiseEvent AnimationEnded
End Sub

Public Function IsRunning() As Boolean
    IsRunning = tmrAni.Enabled
End Function

Private Function GetOrgValue() As Single
    GetOrgValue = CallByName(oObj, sAttr, VbGet)
End Function

Private Function MoveDeceleration(Value As Single, Params As AniParams) As Single
    Dim sTmp As Single
    With Params
        sTmp = (.ToValue - Value) / .Speed
        If Round(sTmp) = 0 Then sTmp = 0
        MoveDeceleration = sTmp
    End With
End Function

Private Function MoveUniform(Value As Single, Params As AniParams) As Single
    With Params
        If Value < .ToValue Then
            If Value + .Speed < .ToValue Then
                MoveUniform = Value + .Speed
            Else
                MoveUniform = .ToValue
            End If
        Else
            If Value - .Speed > .ToValue Then
                MoveUniform = Value - .Speed
            Else
                MoveUniform = .ToValue
            End If
        End If
    End With
End Function

Private Function MoveElasticity(Value As Single, Params As AniParams) As Single
    With Params
        .Speed = (.Speed + (.ToValue - Value) * .K) * .Attn
        MoveElasticity = .Speed
    End With
End Function

Private Sub tmrAni_Timer()
    With apPara
        Select Case .Mode
            Case 0
                CallByName oObj, sAttr, VbLet, GetOrgValue + MoveDeceleration(GetOrgValue, apPara)
                If MoveDeceleration(GetOrgValue, apPara) = 0 Then StopAnimation
            Case 1
                CallByName oObj, sAttr, VbLet, MoveUniform(GetOrgValue, apPara)
                If GetOrgValue = .ToValue Then StopAnimation
            Case 2
                CallByName oObj, sAttr, VbLet, GetOrgValue + MoveElasticity(GetOrgValue, apPara)
                If GetOrgValue = .ToValue Then StopAnimation
        End Select
    End With
    RaiseEvent AnimationProgress
End Sub
