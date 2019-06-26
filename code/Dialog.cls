VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Public displayChar As String, diaChar As String, lastTime As Long, wav As String, zPressed As Boolean, rate As Long
'======================================
Public Sub NewDialog(Str As String, wav As String, rate As Long)
    lastTime = GetTickCount
    displayChar = "": diaChar = Str: zPressed = False
    Me.wav = wav
    Me.rate = rate
    Do While zPressed = False
        ECore.Display
        Sleep 10: DoEvents
    Loop
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    If diaChar <> "" Then
        If ECore.ActivePage <> "GameOverPage" Then
            Page.Paint 0, 60, GH - 200 - 40, GW - 120, 190, argb(255, 0, 0, 0)
            Page.Paint 0, 2 + 60, GH - 200 - 40, GW - 5 - 120, 190, argb(255, 255, 255, 255), size:=5, style:=1
            Page.DrawImageEx "icebig.png", 100, GH - 200 + 48 - 40, 96, 96
        End If
        
        If wav = "ho.wav" Then
            Page.Writes "*", 220, GH - 200 + 50 - 40, 24, argb(255, 255, 255, 255)
            For i = 1 To Len(displayChar)
                Randomize
                Page.Writes Mid(displayChar, i, 1), 220 + 10 + 26 * i + Int(Rnd * 3) * IIf(Int(Rnd * 1) = 1, 1, -1), GH - 200 + 50 - 40 + Int(Rnd * 3) * IIf(Int(Rnd * 1) = 1, 1, -1), 24, argb(255, 255, 255, 255)
            Next
        Else
            Page.Writes "* " & displayChar, 220, GH - 200 + 50 - 40, 24, argb(255, 255, 255, 255)
        End If
        
        If GetTickCount - lastTime >= rate And Len(displayChar) <> Len(diaChar) Then
            lastTime = GetTickCount
            MusicList.Play wav
            displayChar = displayChar & Mid(diaChar, Len(displayChar) + 1, 1)
        End If
        
        If IsKeyUp(VK_Z) And Len(displayChar) = Len(diaChar) Then zPressed = True: diaChar = ""
        If IsKeyPress(VK_X) And Len(displayChar) <> Len(diaChar) Then displayChar = diaChar
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets"
    Page.TopPage = True
    '创建页面
    ECore.Add Page, "Dialog"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
