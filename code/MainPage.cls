VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim br As Single, bg As Single, bb As Single, bt As Long
    Dim tr As Single, tg As Single, tb As Single
    Dim openTime As Long, BattleTime As Long
    Dim NoLoop As Boolean
'======================================
    
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    If GetTickCount - bt >= 3000 Then
        bt = GetTickCount
        Randomize
        Select Case Int(Rnd * 5)
            Case 0
                tr = 255: tg = 255: tb = 255
            Case 1
                tr = 0: tg = 0: tb = 0
            Case 2
                tr = 255: tg = 255: tb = 255
            Case 3
                tr = 0: tg = 0: tb = 0
            Case 4
                tr = 255: tg = 255: tb = 255
            Case 5
                tr = 0: tg = 0: tb = 0
        End Select
    Else
        br = br + (tr - br) / 100
        bg = bg + (tg - bg) / 100
        bb = bb + (tb - bb) / 100
    End If
    
    Page.Clear argb(255, br, bg, bb) '清空画布

    If BattleTime <> 0 Then
        Page.Clear argb(255, 0, 0, 0)
        Dim pro2 As Single
        If GetTickCount - BattleTime <= 500 Then
            Page.DrawImage "heart.png", GW / 2, GH / 2, alpha:=IIf((GetTickCount Mod 100) < 50, 1, 0), Pos:=posOnCenter
        Else
            pro2 = (GetTickCount - BattleTime - 500) / 1000
            If pro2 > 1 Then pro2 = 1: ECore.NewTransform transFadeIn, 500, "BattlePage"
            pro2 = Cubic(pro2, 0, 0, 0, 1)
            Page.DrawImage "heart.png", GW / 2 - 330 * pro2, GH / 2 + 200 * pro2, alpha:=1, Pos:=posOnCenter
        End If
        Exit Sub
    End If

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    If Mouse.state = 2 Then
        'Page.SwitchChannel "attack", "Default"
    End If
    
    Dim pro As Long
    pro = (GetTickCount Mod 5000) + 5000
    
    If pro >= 8000 And pro <= 9000 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 80 - (pro - 8000) / 1000 * 20 / 4, 128 - (pro - 8000) / 1000 * 20, 128 - (pro - 8000) / 1000 * 20, Pos:=posOnCenter
    ElseIf pro >= 9000 And pro <= 9500 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 80 - 20 / 4, 128 - 20, 128 - 20, Pos:=posOnCenter
    ElseIf pro >= 9500 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 80 - 20 / 4 + (pro - 9500) / 500 * 20 / 4, 128 - 20 + (pro - 9500) / 500 * 20, 128 - 20 + (pro - 9500) / 500 * 20, Pos:=posOnCenter
    Else
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 80, 128, 128, Pos:=posOnCenter
    End If
    
    If GetTickCount - openTime >= 8000 And Dialog.diaChar = "" Then
        Dialog.NewDialog "你好，我是冰棍。", "say.wav", 100
        Dialog.NewDialog "你最好的朋友，冰棍。", "say.wav", 80
        Dialog.NewDialog "...", "say.wav", 200
        Dialog.NewDialog "...", "say.wav", 200
        Dialog.NewDialog "你说你很喜欢我是吗？", "say.wav", 200
        Dialog.NewDialog "那么", "say.wav", 100
        BGM.Stops: NoLoop = True
        Dialog.NewDialog "如果我说现在我要杀了你呢？", "ho.wav", 200
        Dialog.NewDialog "你也会很开心地吃掉我吗？", "ho.wav", 200
        MusicList.Play "battle.wav"
        BattleTime = GetTickCount
    End If
    
    'Page.DrawAnimation "attack", GW / 2, GH / 2 - 80
    
    If (BGM.IsPlaying = False) And (Not NoLoop) Then BGM.Play
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
    Page.Res.ClipCircle "icebig.png"
    
    Set BGM = New GMusic
    BGM.Create App.path & "\music\Barrier.mp3"
    BGM.Volume = 0.3
    BGM.Play

    '创建页面
    ECore.Add Page, "MainPage"
    
    openTime = GetTickCount
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
