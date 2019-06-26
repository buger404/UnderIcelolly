VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BattlePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim focus1 As Long, focus2 As Long, lastFocus As Long
    Dim HP As Long, UHP As Long, Mercy As Long
    Dim displayText As String, TargetText As String, lastTime As Long
    Dim roundTime As Long, RoundText As String
    Dim hx As Long, hy As Long
    Dim atkKind As Integer
'======================================
Public Sub AttackProcess()
    Dim fx As Long, fy As Long, fw As Long, fh As Long
    fx = 46 + 200: fy = GH - 310 - 80
    fw = GW - 46 * 2 - 400: fh = 170 + 80
    Dim x As Long, y As Long
    
    Select Case atkKind
        Case 0
            For i = 0 To 2
                For s = 0 To 3
                    x = fx + i / 2 * (fw - 60) + Sin(GetTickCount / 600 + s) * 40
                    y = 40 + fy + s / 3 * (fh - 20 - 80) + Sin(GetTickCount / 800 + s) * 30
                    Page.Paint 0, x, y, 60, 20, argb(255, 255 - s * 40, s * 40, i * 40)
                    Page.Writes "BUG", x + 12, y, 18, argb(255, 255, 255, 255), style:=FontStyleBold
                    If hx >= x And hy >= y And hx <= x + 60 And hy <= y + 20 Then
                        MusicList.Play "hurt.wav"
                        HP = HP - 10
                    End If
                Next
            Next
    End Select
End Sub
Public Sub ShowText(Str As String)
    lastTime = GetTickCount
    TargetText = Str
    displayText = ""
End Sub
Public Sub ShowRoundText()
    Randomize
    Select Case Int(Rnd * 5)
        Case 0: RoundText = "啊。。。讨厌。": ShowText "冰棍看起来有点讨厌你了。": Mercy = Mercy - 1
        Case 1: RoundText = "别这样。": ShowText "冰棍好像不喜欢那样。": Mercy = Mercy - 1
        Case 2: RoundText = "太恶心了。": ShowText "冰棍开始觉得你很恶心。": Mercy = Mercy - 1
        Case 3: RoundText = "你不会害臊吗？": ShowText "冰棍很生气地站在屋顶上吹风。": Mercy = Mercy - 1
        Case 4: RoundText = "别用你的舌头碰我。": ShowText "冰棍好像不希望你吃掉它。"
        Case 5: RoundText = "别用你的舌头碰我。": ShowText "冰棍好像不希望你吃掉它。"
    End Select
End Sub
Public Sub ShowRoundText2()
    Randomize
    Select Case Int(Rnd * 5)
        Case 0: RoundText = "天啊！你在做什么！": ShowText "冰棍被你的行为吓到了。"
        Case 1: RoundText = "别这样。": ShowText "冰棍觉得有些突然。"
        Case 2: RoundText = "太恶心了。": ShowText "冰棍觉得你可能是一个坏蛋。": Mercy = Mercy - 1
        Case 3: RoundText = "或许我有点喜欢这样。": Mercy = Mercy + 2: ShowText "冰棍好像有点融化了。"
        Case 4: RoundText = "哈...哈....": Mercy = Mercy + 1: ShowText "屋顶上回荡着冰棍的笑声。"
        Case 5: RoundText = "或许我有点喜欢这样。": Mercy = Mercy + 2: ShowText "液态冰棍的味道。"
    End Select
End Sub
Public Sub ShowRoundText3()
    Randomize
    Select Case Int(Rnd * 5)
        Case 0: RoundText = "你怎么可以打可爱的冰棍呢！": ShowText "冰棍生气地在屋顶上吹风。": Mercy = Mercy - 1
        Case 1: RoundText = "噢，好热啊。": ShowText "冰棍觉得自己快要融化了。"
        Case 2: RoundText = "冰棍的报复。": ShowText "冰棍现在很生气。": Mercy = Mercy - 2
        Case 3: RoundText = "看来你不想成为冰棍的朋友。": ShowText "冰棍想要和你成为朋友。"
        Case 4: RoundText = "。。。": ShowText "冰棍已经不想理你了。": Mercy = Mercy - 1
        Case 5: RoundText = "。。。": ShowText "冰棍已经不想理你了。": Mercy = Mercy - 1
    End Select
End Sub
Public Sub ChangeFocus(f As Long)
    If f > focus1 Then MusicList.Play "choice.wav"
    If focus1 = 0 Then lastFocus = focus2
    focus1 = f
    focus2 = 0
    If f = 0 Then ShowText TargetText: focus2 = lastFocus
    If f = 9 Then hx = GW / 2 - 10: hy = GH / 2 - 5
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    If BGM.IsPlaying And Dialog.diaChar <> "" Then
        BGM.Stops
        Page.Clear argb(255, 0, 0, 0) '清空画布
        Exit Sub
    End If
    If Dialog.diaChar <> "" Then
        Page.Clear argb(255, 0, 0, 0) '清空画布
        Exit Sub
    End If
    
    Page.Clear argb(255, 0, 0, 0) '清空画布
    Dim by()
    ReDim by(Int(GW / 50) * 2 + 1)
    
    For i = 0 To (UBound(by) - 6) Step 2
        by(i) = i / (UBound(by) - 6) * GW
        by(i + 1) = GH - Sin(i / (UBound(by) - 6) * 2 + GetTickCount / 500) * 50 - 70
    Next
    by(UBound(by) - 5) = GW: by(UBound(by) - 4) = GH - Sin(2 + GetTickCount / 500) * 50 - 70
    by(UBound(by) - 3) = GW: by(UBound(by) - 2) = GH
    by(UBound(by) - 1) = 0: by(UBound(by)) = GH
    Page.PaintPolygon2 argb(120, 255, 0, 0), 0, by
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim bpro As Single
    If focus1 = 9 Then
        bpro = (GetTickCount - roundTime) / 500
        If bpro > 1 Then bpro = 1
        bpro = Cubic(bpro, 0, 1, 1, 1)
    ElseIf GetTickCount - roundTime <= 500 Then
        bpro = 1 - (GetTickCount - roundTime) / 500
        If bpro < 0 Then bpro = 0
        bpro = Cubic(bpro, 0, 1, 1, 1)
    End If
    
    If focus1 = 9 Then
        Page.DrawImage "dialog.png", GW - 310, 15, alpha:=1
        Dim row As Long, rowl As Long
        
        For i = 1 To Len(RoundText)
            If 20 * (i - rowl) > 180 Then
                row = row + 1: rowl = i - 1
            End If
            Page.Writes Mid(RoundText, i, 1), GW - 310 + 20 + 20 * (i - rowl), 40 + Sin(i + GetTickCount / 200) * 5 + 30 * row, 20, argb(255, 0, 0, 0)
        Next
        
    End If
    
    Page.DrawImage "attack" & IIf(focus1 = 0 And focus2 = 0, 1, 0) & ".png", GW / 2 - 330 + 35, GH / 2 + 220, alpha:=1, Pos:=posOnCenter
    
    Page.DrawImage "act" & IIf(focus1 = 0 And focus2 = 1, 1, 0) & ".png", GW / 2 - 330 + 35 + 180, GH / 2 + 220, alpha:=1, Pos:=posOnCenter
    
    Page.Writes "YOU", 46, GH - 120, 28, argb(255, 255, 255, 255), style:=FontStyleBold
    Page.Writes "LV 99", 186, GH - 120, 28, argb(255, 255, 255, 255), style:=FontStyleBold
    
    Page.Writes "HP", 336, GH - 120 + 8, 18, argb(255, 255, 255, 255)
    Page.Paint 0, 366, GH - 120 + 4, 210, 25, argb(255, 255, 0, 0)
    Page.Paint 0, 366, GH - 120 + 4, 210 * (HP / 900), 25, argb(255, 255, 255, 0)
    Page.Writes HP & "/900", GW - 210, GH - 120 + 3, 28, argb(255, 255, 255, 255), style:=FontStyleBold
    
    Page.Paint 0, 46 + bpro * 200, GH - 310 - bpro * 80, GW - 46 * 2 - bpro * 400, 170 + bpro * 80, argb(255, 0, 0, 0), style:=0
    
    'Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 160 - bpro * 40, 128, 128, Pos:=posOnCenter
    
    Dim pro As Long
    pro = (GetTickCount Mod 2000) + 8000
    If pro >= 8000 And pro <= 9000 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 160 - (pro - 8000) / 1000 * 20 / 4 - bpro * 40, 128 - (pro - 8000) / 1000 * 20, 128 - (pro - 8000) / 1000 * 20, Pos:=posOnCenter
    ElseIf pro >= 9000 And pro <= 9500 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 160 - 20 / 4 - bpro * 40, 128 - 20, 128 - 20, Pos:=posOnCenter
    ElseIf pro >= 9500 Then
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 160 - 20 / 4 + (pro - 9500) / 500 * 20 / 4 - bpro * 40, 128 - 20 + (pro - 9500) / 500 * 20, 128 - 20 + (pro - 9500) / 500 * 20, Pos:=posOnCenter
    Else
        Page.DrawImageEx "icebig.png", GW / 2, GH / 2 - 160 - bpro * 40, 128, 128, Pos:=posOnCenter
    End If
    
    
    If focus1 = 0 Then
        Page.Writes "* " & displayText, 70, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    If focus1 = 1 Then
        Page.Writes "* 冰棍", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
        Page.Paint 0, 400, GH - 310 + 29, 150, 21, argb(255, 128, 128, 128)
        Page.Paint 0, 400, GH - 310 + 29, 150 * (UHP / 9000), 21, argb(255, 0, 255, 0)
    End If
    If focus1 = 2 Then
        Page.Writes "* 冰棍", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    If focus1 = 3 Then
        Page.Writes "* 舔一口", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
        Page.Writes "* 含在嘴里", 510, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    
    If focus1 = 0 Then
        Page.DrawImage "heart.png", GW / 2 - 330 + 180 * focus2, GH / 2 + 220, alpha:=1, Pos:=posOnCenter
        If IsKeyUp(VK_Z) Then
            If focus2 = 0 Then
                ChangeFocus 1
            Else
                ChangeFocus 2
            End If
        End If
    End If
    If focus1 = 1 Then
        Page.DrawImage "heart.png", GW / 2 - 330 + 10, GH / 2 + 5, alpha:=1, Pos:=posOnCenter
        If IsKeyPress(VK_X) Then ChangeFocus 0
    End If
    If focus1 = 2 Then
        Page.DrawImage "heart.png", GW / 2 - 330 + 10, GH / 2 + 5, alpha:=1, Pos:=posOnCenter
        If IsKeyPress(VK_X) Then ChangeFocus 0
        If IsKeyUp(VK_Z) Then ChangeFocus 3
    End If
    If focus1 = 3 Then
        Page.DrawImage "heart.png", GW / 2 - 330 + 10 + 400 * focus2, GH / 2 + 5, alpha:=1, Pos:=posOnCenter
        If IsKeyPress(VK_X) Then ChangeFocus 0
        If IsKeyUp(VK_Z) Then
            If focus2 = 0 Then Call ShowRoundText
            If focus2 = 1 Then Call ShowRoundText2
            ChangeFocus 9
            roundTime = GetTickCount ': ShowRoundText
        End If
    End If
    If focus1 = 9 Then
        Page.DrawImage "heart.png", hx, hy, alpha:=1, Pos:=posNormal
        If IsKeyPress(VK_X) Then
            ChangeFocus 0
            roundTime = GetTickCount
            Select Case Mercy
                Case Is < -20
                    Dialog.NewDialog "我以为你想和我做朋友的。", "say.wav", 60
                    Dialog.NewDialog "但是你真的太令我讨厌了。", "ho.wav", 60
                    End
                Case Is < -10
                    ShowText "冰棍现在唯一的愿望是离开这里。"
                Case Is < -5
                    ShowText "冰棍已经不想理你了。"
                Case Is > 15
                    ShowText "液态冰棍。"
                Case Is > 10
                    ShowText "屋顶洋溢着液态冰棍的味道。"
            End Select
        End If
    End If
    
    If BGM.IsPlaying = False Then BGM.Play
    
    If GetTickCount - lastTime >= 60 And Len(displayText) <> Len(TargetText) And focus1 = 0 Then
        lastTime = GetTickCount
        MusicList.Play "say.wav"
        displayText = displayText & Mid(TargetText, Len(displayText) + 1, 1)
    End If
    
    If focus1 = 9 Then
        Call AttackProcess
        If IsKeyPress(VK_LEFT) Then hx = hx - 2
        If IsKeyPress(VK_RIGHT) Then hx = hx + 2
        If IsKeyPress(VK_UP) Then hy = hy - 2
        If IsKeyPress(VK_DOWN) Then hy = hy + 2
        Dim fx As Long, fy As Long, fw As Long, fh As Long
        fx = 46 + 200 + 1: fy = GH - 310 - 80 + 1
        fw = GW - 46 * 2 - 400 - 2: fh = 170 + 80 - 2
        If hx < fx Then hx = fx
        If hx > fx + fw - 20 Then hx = fx + fw - 20
        If hy < fy Then hy = fy
        If hy > fy + fh - 20 Then hy = fy + fh - 20
    End If
    
    Page.Paint 0, 46 + bpro * 200, GH - 310 - bpro * 80, GW - 46 * 2 - bpro * 400, 170 + bpro * 80, argb(255, 255, 255, 255), style:=1, size:=4

    If HP < 0 Then
        BGM.Stops
        GameWindow.Cls
        GameWindow.Refresh
        DoEvents
        MusicList.Play "break.wav"
        Sleep 500
        MusicList.Play "break2.wav"
        Sleep 2000
        ECore.ActivePage = "GameOverPage"
    End If
End Sub
Public Sub Move(dir As Long, key As Long)
    If focus1 = 1 Then Exit Sub
    focus2 = focus2 + dir: MusicList.Play "focus.wav"
    If focus1 = 0 Or focus1 = 3 Then
        If focus2 < 0 Then focus2 = 1
        If focus2 > 1 Then focus2 = 0
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
    BGM.Create App.path & "\music\SpiderDance.mp3"
    BGM.Play
    HP = 900: UHP = 9000
    ShowText "冰棍想要你吃掉它。"
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
    Page.LoadAnimationsFromDir App.path & "\animation"
    Page.CreatePlayAnimation "attack", "attack", "Default"
    
    '创建页面
    ECore.Add Page, "BattlePage"

    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
