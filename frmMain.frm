VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Linkman"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":4781A
   MousePointer    =   99  'Custom
   ScaleHeight     =   5325
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSta 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1920
      ScaleHeight     =   735
      ScaleWidth      =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   2280
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2100
      End
   End
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picScr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Ani As cAniControl
Attribute Ani.VB_VarHelpID = -1
Dim sAniTag As String

Sub AutoLinkYqyz()
    If LCase(PCName) <> "hygd-pc" Then
        ChangeStatus "请确保此程序运行于学校电脑"
        EndApp
        Exit Sub
    End If

    Const sIPHead As String = "10.40.83.", sSubnet As String = "255.255.255.128", sGway As String = "10.40.83.254", sDNS As String = "10.42.5.6"
    Const lLRange As Long = 130, lURange As Long = 200, lTimeOut As Long = 5, sSkipIP As String = "10.40.83.192"
    
    On Error GoTo ALYErr
    
    If PingIP("10.40.80.4") Then
        ChangeStatus "已连接至网络"
        EndApp
        Exit Sub
    End If
    
    Dim sIP As String, lCount As Long, lRtn As Long
    lCount = 1
    Randomize
LinkProc:
    ChangeStatus "重选 IP..."
    sIP = sIPHead & Trim(Str(Int((lURange - lLRange + 1) * Rnd + lLRange)))
    If sIP = sSkipIP Then GoTo LinkProc
    ChangeStatus "第 " & CStr(lCount) & " 次尝试连接 " & sIP & " ..."
    lRtn = LinkConfig(sIP, sSubnet, sGway, sDNS, "8.8.8.8")
    If lRtn <> 0 Then GoTo ReLink
    ChangeStatus "等待应用设置..."
    Sleep 10000
    ChangeStatus "正在检查连接..."
    If Not PingIP("10.40.80.4") Then GoTo ReLink
    ChangeStatus "连接成功！"
    EndApp
    
    Exit Sub
ReLink:
    If lCount >= lTimeOut Then GoTo ALYErr
    lCount = lCount + 1
    GoTo LinkProc
ALYErr:
    ChangeStatus "自动设置失败！请尝试重新运行程序"
    EndApp
End Sub

Sub ChangeStatus(sSta As String)
    lblStatus.Caption = sSta
End Sub

Sub EndApp()
    Sleep 2000
    Dim aniPara As AniParams
    With aniPara
        .Mode = Deceleration
        .Speed = 6
        .ToValue = Me.ScaleHeight
    End With
    sAniTag = "end"
    Ani.DoAnimation picScr, "Height", aniPara
End Sub

Private Sub Ani_AnimationEnded()
    If sAniTag = "" Then
        Sleep 1000
        AutoLinkYqyz
    Else
        tmrFade.Enabled = True
    End If
End Sub

Private Sub Ani_AnimationProgress()
    With picScr
        .Width = Me.ScaleWidth * (.Height / Me.ScaleHeight)
        .Move (Me.ScaleWidth - .Width) / 2, (Me.ScaleHeight - .Height) / 2
        Me.Cls
        Me.PaintPicture .Picture, .Left, .Top, .Width, .Height, 0, 0, Me.ScaleWidth, Me.ScaleHeight, vbSrcCopy
        picSta.Top = Me.ScaleHeight - (picSta.Height + 1800) * ((Me.ScaleHeight - .Height) / (Me.ScaleHeight * 0.6))
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then EndApp
End Sub

Private Sub Form_Load()
    ChangeStatus "Linkman ver " & App.Major & "." & App.Minor & "." & App.Revision & " By MaxXSoft && 阳一计算机社"
    
    Dim lRtn As Long
    lRtn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    lRtn = lRtn Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, lRtn
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
    SetWindowPos Me.hWnd, -1, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, 0
    
    Set Ani = New cAniControl
    
    BitBlt Me.hDC, 0, 0, Screen.Width, Screen.Height, GetDC(GetDesktopWindow), 0, 0, vbSrcCopy
    picScr.Picture = Me.Image
    picScr.Move 0, 0, Me.Width, Me.Height
    
    Dim aniPara As AniParams
    With aniPara
        .Mode = Deceleration
        .Speed = 7
        .ToValue = Me.ScaleHeight * 0.6
    End With
    Ani.DoAnimation picScr, "Height", aniPara
End Sub

Private Sub Form_Resize()
    picSta.Move 0, Me.ScaleHeight, Me.ScaleWidth
    lblStatus.Move 0, 0, picSta.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Ani = Nothing
End Sub

Private Sub tmrFade_Timer()
    Static lAlp As Long
    lAlp = lAlp + 10
    If lAlp > 255 Then
        tmrFade.Enabled = False
        Unload Me
        Exit Sub
    End If
    SetLayeredWindowAttributes Me.hWnd, 0, 255 - lAlp, LWA_ALPHA
End Sub
