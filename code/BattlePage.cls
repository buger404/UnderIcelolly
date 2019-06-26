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
    Dim oldX As Long, oldY As Long
    Dim atkKind As Integer, atkTime As Long
    Dim hurtTime As Long
    Dim ODC As Long, OG As Long
    Dim ADC As Long, AG As Long
    Dim atkFrame As Long, atkStop As Long
    Dim hurtUTime As Long, hurtU As Long
    Dim atkBuf(), atkBuf2(), atkBuf3()
'======================================
Public Sub AttackProcess()
    Dim fx As Long, fy As Long, fw As Long, fh As Long
    fx = 46 + 200: fy = GH - 310 - 80
    fw = GW - 46 * 2 - 400: fh = 170 + 80
    Dim x As Long, y As Long, i As Integer, s As Integer
    Dim Crash As New GCrashBox
    Crash.Reset fw, fh, 1
    Dim Quiet As Boolean
    Quiet = True
    
    '替换
    Page.CDC = ADC: Page.GG = AG
    Page.Clear 0
    
    If GetTickCount - hurtTime > 3000 Then
        Page.DrawImage "heart.png", hx, hy, alpha:=1, Pos:=posNormal
    ElseIf GetTickCount Mod 200 < 100 Then
        Page.DrawImage "heart.png", hx, hy, alpha:=1, Pos:=posNormal
    ElseIf GetTickCount - roundTime <= 3000 Then
        Page.DrawImage "heart.png", hx, hy, alpha:=1, Pos:=posNormal
    End If
    
    Dim pro As Single
    pro = (GetTickCount - roundTime) / atkTime
    Dim pro2 As Single
    
    Dim waitKind As Long, atkK As Long
    atkK = atkKind
    
reatk:
    Select Case atkK
        Case 0
            For i = 0 To 2
                For s = 0 To 3
                    x = i / 2 * (fw - 60) + Sin(GetTickCount / 600 + s) * 40 / (atkTime / 6000)
                    y = 40 + s / 3 * (fh - 20 - 80) + Sin(GetTickCount / 800 + s) * 30 / (atkTime / 6000)
                    Page.Paint 0, x, y, 60, 20, argb(255, 255 - s * 40, s * 40, i * 40)
                    Page.Writes "BUG", x + 12, y, 18, argb(255, 255, 255, 255), style:=FontStyleBold
                    Crash.AddCrashRect x, y, 60, 20, i, s
                Next
            Next
        Case 1
            Page.Paint 0, 0, GH * pro, 200, 20, argb(255, 255, 255, 255)
            Page.Paint 0, fw - 100, GH * pro - 100, 100, 20, argb(255, 255, 255, 255)
            Page.Paint 0, 0, GH * pro - 300, 270, 20, argb(255, 255, 255, 255)
            Crash.AddCrashRect 0, GH * pro, 200, 20, 0, 0
            Crash.AddCrashRect fw - 100, GH * pro - 100, 100, 20, 0, 0
            Crash.AddCrashRect 0, GH * pro - 300, 270, 20, 0, 0
        Case 2
            For i = 0 To 5
                Page.Paint 0, 0, fh * 3 * pro - 40 * i, 100 + 15 * i, 20, argb(255, 255, 255 - i / 10 * 255, 255 - i / 10 * 255)
                Crash.AddCrashRect 0, fh * 3 * pro - 40 * i, 100 + 15 * i, 20, 0, 0
            Next
            For i = 6 To 10
                Page.Paint 0, fw - 100 - 17 * (i - 6), fh * 3 * pro - 40 * i, 100 + 17 * (i - 6), 20, argb(255, 255, 255 - i / 10 * 255, 255 - i / 10 * 255)
                Crash.AddCrashRect fw - 100 - 17 * (i - 6), fh * 3 * pro - 40 * i, 100 + 17 * (i - 6), 20, 0, 0
            Next
        Case 3
            For i = 0 To 5
                Page.Paint 0, 0, fh * 3 * pro - 40 * i, 100 + 15 * i + pro * 150, 20, argb(255, 255, 255 - i / 10 * 255, 255 - i / 10 * 255)
                Crash.AddCrashRect 0, fh * 3 * pro - 40 * i, 100 + 15 * i + pro * 150, 20, 0, 0
            Next
            For i = 6 To 10
                Page.Paint 0, fw - 500 - 17 * (i - 6) + pro * 700, fh * 3 * pro - 40 * i - 100, 500 + 17 * (i - 6) - pro * 700, 20, argb(255, 255, 255 - i / 10 * 255, 255 - i / 10 * 255)
                Crash.AddCrashRect fw - 500 - 17 * (i - 6) + pro * 700, fh * 3 * pro - 40 * i - 100, 500 + 17 * (i - 6) - pro * 700, 20, 0, 0
            Next
        Case 4
            For i = 0 To UBound(atkBuf)
                pro2 = (GetTickCount - roundTime - 500 * i) / (1000 / (atkTime / 10000))
                If pro2 > 1 Then
                    If pro2 > 2 Then
                        pro2 = 1 - (pro2 - 2)
                    Else
                        pro2 = 1
                    End If
                End If
                If pro2 < 0 Then pro2 = 0
                Page.Paint 0, atkBuf(i) / 12 * fw, 0, fw / 12, fh, argb(255 * pro2, 255, 255, 255)
                If pro2 = 1 Then
                    If Val(atkBuf2(i)) = 0 Then atkBuf2(i) = 1: MusicList.Play "attack2.wav"
                    Crash.AddCrashRect atkBuf(i) / 12 * fw, 0, fw / 12, fh, 0, 0
                End If
            Next
        Case 5
            For i = 0 To UBound(atkBuf)
                pro2 = (GetTickCount - roundTime - 200 * i) / (800 / (atkTime / 10000))
                If pro2 > 1 Then
                    If pro2 > 2 Then
                        pro2 = 1 - (pro2 - 2)
                    Else
                        pro2 = 1
                    End If
                End If
                If pro2 < 0 Then pro2 = 0
                Page.Paint 0, atkBuf(i) / 12 * fw, 0, fw / 12, fh, IIf(atkBuf3(i) = 0, argb(255 * pro2, 255, 128, 0), argb(255 * pro2, 0, 128, 255))
                If pro2 = 1 Then
                    If Val(atkBuf2(i)) = 0 Then atkBuf2(i) = 1: MusicList.Play "attack5.wav"
                    Crash.AddCrashRect atkBuf(i) / 12 * fw, 0, fw / 12, fh, Val(atkBuf3(i)), 0
                End If
            Next
        Case 6
            atkK = 0: waitKind = 1
            GoTo reatk
        Case 7
            atkK = 1: waitKind = 5
            GoTo reatk
        Case 8
            atkK = 0: waitKind = 2
            GoTo reatk
        Case 9
            atkK = 1: waitKind = 4
            GoTo reatk
    End Select
    
    If waitKind <> 0 Then atkK = waitKind: waitKind = 0: GoTo reatk
    
    '还原
    Page.CDC = ODC: Page.GG = OG
    BitBlt ODC, fx, fy, fw, fh, ADC, 0, 0, vbSrcCopy
    
    If IsKeyPress(VK_LEFT) Then hx = hx - 2: Quiet = False
    If IsKeyPress(VK_RIGHT) Then hx = hx + 2: Quiet = False
    If IsKeyPress(VK_UP) Then hy = hy - 2: Quiet = False
    If IsKeyPress(VK_DOWN) Then hy = hy + 2: Quiet = False

    If hx < 0 Then hx = 0
    If hx > fw - 20 Then hx = fw - 20
    If hy < 0 Then hy = 0
    If hy > fh - 20 Then hy = fh - 20
    
    If GetTickCount - hurtTime > 3000 Then
        Dim cr As Integer, hurts As Boolean
        cr = Crash.CheckCrashRect(hx, hy, 20, 20, 0, 0, False)
        If cr <> 0 Then
            hurts = True
            If atkKind = 5 Then
                If (Quiet And Crash.GetObj1(cr) = 1) Then hurts = False
                If ((Not Quiet) And Crash.GetObj1(cr) = 0) Then hurts = False
            End If
            If hurts Then
                hurtTime = GetTickCount
                MusicList.Play "hurt.wav"
                HP = HP - Int(Rnd * 50 + 50)
            End If
        End If
    End If
    Set Crash = Nothing
    
    If GetTickCount - roundTime > atkTime Then
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
            Case Is > 20
                ShowText "你和冰棍成为了朋友。"
                MusicList.Play "die.wav"
                UHP = -1
                BGM.Stops
                ECore.NewTransform
            Case Is > 15
                ShowText "液态冰棍。"
            Case Is > 10
                ShowText "屋顶洋溢着液态冰棍的味道。"
        End Select
    End If
    
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
    Dim pp As Long
    pp = Int(Rnd * 5)
    If Mercy > 5 Then pp = 6
    Select Case pp
        Case 0: RoundText = "天啊！你在做什么！": ShowText "冰棍被你的行为吓到了。"
        Case 1: RoundText = "别这样。": ShowText "冰棍觉得有些突然。"
        Case 2: RoundText = "太恶心了。": ShowText "冰棍觉得你可能是一个坏蛋。": Mercy = Mercy - 1
        Case 3: RoundText = "或许我有点喜欢这样。": Mercy = Mercy + 2: ShowText "冰棍好像有点融化了。"
        Case 4: RoundText = "哈...哈....": Mercy = Mercy + 1: ShowText "屋顶上回荡着冰棍的笑声。"
        Case 5: RoundText = "或许我有点喜欢这样。": Mercy = Mercy + 2: ShowText "液态冰棍的味道。"
        Case 6: RoundText = "我喜欢这样。": Mercy = Mercy + 2: ShowText "液态冰棍的味道。"
    End Select
    If Mercy > 10 Then RoundText = "我们其实可以做个朋友？。。。"
End Sub
Public Sub ShowRoundText3()
    Randomize
    Select Case Int(Rnd * 5)
        Case 0: RoundText = "你怎么可以打可爱的冰棍呢！": ShowText "冰棍生气地在屋顶上吹风。": Mercy = Mercy - IIf(Mercy > -3, 1, 0)
        Case 1: RoundText = "噢，好热啊。": ShowText "冰棍觉得自己快要融化了。"
        Case 2: RoundText = "冰棍的报复。": ShowText "冰棍现在很生气。": Mercy = Mercy - IIf(Mercy > -3, 2, 0)
        Case 3: RoundText = "看来你不想成为冰棍的朋友。": ShowText "冰棍想要和你成为朋友。"
        Case 4: RoundText = "。。。": ShowText "冰棍已经不想理你了。": Mercy = Mercy - IIf(Mercy > -3, 1, 0)
        Case 5: RoundText = "。。。": ShowText "冰棍已经不想理你了。": Mercy = Mercy - IIf(Mercy > -3, 1, 0)
    End Select
End Sub
Public Sub ChangeFocus(f As Long)
    If f > focus1 And focus1 <> 6 Then MusicList.Play "choice.wav"
    If focus1 = 0 Then lastFocus = focus2
    focus1 = f
    focus2 = 0
    If f = 0 Then ShowText TargetText: focus2 = lastFocus
End Sub
Public Sub BeginAttack()
    atkKind = Int(Rnd * 10)
    If atkKind = 10 Then atkKind = 9
    
    Select Case atkKind
        Case 0
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) - 20
            atkTime = 6000
            hurtTime = GetTickCount
        Case 1
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
            atkTime = 5000
        Case 2
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
            atkTime = 4000
        Case 3
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
            atkTime = 4000
        Case 4
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
            atkTime = 10000
            Randomize
            ReDim atkBuf(20): ReDim atkBuf2(20)
            For i = 0 To 15
                atkBuf(i) = Int(Rnd * 15) - 1
            Next
        Case 5
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
            atkTime = 10000
            Randomize
            ReDim atkBuf(35): ReDim atkBuf2(35): ReDim atkBuf3(35)
            For i = 0 To 35
                atkBuf(i) = Int(Rnd * 15) - 1
                atkBuf3(i) = Int(Rnd * 2)
                If atkBuf3(i) = 2 Then atkBuf3(i) = 1
            Next
        Case 6
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) - 20
            atkTime = 6000
            hurtTime = GetTickCount
        Case 7
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) - 20
            atkTime = 10000
            ReDim atkBuf(35): ReDim atkBuf2(35): ReDim atkBuf3(35)
            For i = 0 To 35
                atkBuf(i) = Int(Rnd * 15) - 1
                atkBuf3(i) = Int(Rnd * 2)
                If atkBuf3(i) = 2 Then atkBuf3(i) = 1
            Next
        Case 8
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) - 20
            atkTime = 6000
            hurtTime = GetTickCount
        Case 9
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) - 20
            atkTime = 10000
            ReDim atkBuf(20): ReDim atkBuf2(20)
            For i = 0 To 15
                atkBuf(i) = Int(Rnd * 15) - 1
            Next
        Case Else
            hx = (GW - 46 * 2 - 400) / 2 - 10: hy = (170 + 80) / 2 - 10
    End Select
    
    If UHP < 2000 Then
        atkTime = atkTime * 0.7
    ElseIf UHP < 4000 Then
        atkTime = atkTime * 0.8
    ElseIf UHP < 6000 Then
        atkTime = atkTime * 0.9
    End If
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
    Dim hurtX As Long
    Dim atkPro As Single, atkpro2 As Single
    atkPro = (GetTickCount - hurtUTime) / 1000
    
    If atkPro > 1 Then atkPro = 1
    atkpro2 = (GetTickCount - hurtUTime) / 3000
    If atkpro2 > 1 Then atkPro = 1
    
    If focus1 = 6 Then
        If GetTickCount Mod 250 < 120 Then
            hurtX = 10 - 10 * atkpro2
        Else
            hurtX = -10 + 10 * atkpro2
        End If
    End If
    
    If (Not (focus1 = 0 And UHP < 0)) Then
        If focus1 = 6 Then
            Page.DrawImageEx "icebig.png", GW / 2 + hurtX, GH / 2 - 120 - bpro * 80, 128, 128, Pos:=posOnCenter
        Else
            If pro >= 8000 And pro <= 9000 Then
                Page.DrawImageEx "icebig.png", GW / 2 + hurtX, GH / 2 - 120 - (pro - 8000) / 1000 * 20 / 4 - bpro * 80, 128 - (pro - 8000) / 1000 * 20, 128 - (pro - 8000) / 1000 * 20, Pos:=posOnCenter
            ElseIf pro >= 9000 And pro <= 9500 Then
                Page.DrawImageEx "icebig.png", GW / 2 + hurtX, GH / 2 - 120 - 20 / 4 - bpro * 80, 128 - 20, 128 - 20, Pos:=posOnCenter
            ElseIf pro >= 9500 Then
                Page.DrawImageEx "icebig.png", GW / 2 + hurtX, GH / 2 - 120 - 20 / 4 + (pro - 9500) / 500 * 20 / 4 - bpro * 80, 128 - 20 + (pro - 9500) / 500 * 20, 128 - 20 + (pro - 9500) / 500 * 20, Pos:=posOnCenter
            Else
                Page.DrawImageEx "icebig.png", GW / 2 + hurtX, GH / 2 - 120 - bpro * 80, 128, 128, Pos:=posOnCenter
            End If
        End If
    End If
    
    If atkStop <> 0 Then Page.DrawAnimation "attack", GW / 2, GH / 2 - 160
    
    If focus1 = 0 Then
        Page.Writes "* " & displayText, 70, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    If focus1 = 1 Then
        Page.Writes "* 冰棍", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
        Page.Paint 0, 400, GH - 310 + 29, 150, 21, argb(255, 128, 128, 128)
        Page.Paint 0, 400, GH - 310 + 29, 150 * (UHP / 9000), 21, argb(255, 0, 255, 0)
        If IsKeyUp(VK_Z) Then atkFrame = GetTickCount: atkStop = 0: hurtUTime = 0: ChangeFocus 4
    End If
    If focus1 = 2 Then
        Page.Writes "* 冰棍", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    If focus1 = 3 Then
        Page.Writes "* 舔一口", 110, GH - 310 + 25, 25, argb(255, 255, 255, 255)
        Page.Writes "* 含在嘴里", 510, GH - 310 + 25, 25, argb(255, 255, 255, 255)
    End If
    If focus1 = 4 Then
        Page.DrawImage "attackpad.png", 46 + (GW - 46 * 2) / 2, GH - 310 + 170 / 2, alpha:=1, Pos:=posOnCenter
        Page.Paint 0, (GetTickCount - atkFrame) / 1500 * GW, GH - 310 + 170 / 2, 10, 130, argb(255, 255, 255, 255), Pos:=posOnCenter
        Page.Paint 0, (GetTickCount - atkFrame) / 1500 * GW, GH - 310 + 170 / 2, 10, 130, argb(255, 0, 0, 0), size:=1, style:=1, Pos:=posOnCenter
        
        If IsKeyUp(VK_Z) Or (GetTickCount - atkFrame) / 1500 * GW > GW Then
            atkStop = (GetTickCount - atkFrame) / 1500 * GW
            If GW / 2 > atkStop Then
                hurtU = (Rnd * 0.3 + 1) * (350 - (GW / 2 - atkStop) / (GW / 2) * 350 + IIf(UHP < 9000 * 0.2, 300, 0))
            Else
                hurtU = (Rnd * 0.3 + 1) * (350 - (atkStop - GW / 2) / (GW / 2) * 350 + IIf(UHP < 9000 * 0.2, 300, 0))
            End If
            If atkStop > GW Then hurtU = 0
            UHP = UHP - hurtU
            Page.SwitchChannel "attack", "Default"
            ChangeFocus 5
        End If
    End If
    If focus1 = 5 Or focus1 = 6 Then
        Page.DrawImage "attackpad.png", 46 + (GW - 46 * 2) / 2, GH - 310 + 170 / 2, alpha:=1, Pos:=posOnCenter
        Page.Paint 0, atkStop, GH - 310 + 170 / 2, 10, 130, argb(255, 255, 255, 255), Pos:=posOnCenter
        Page.Paint 0, atkStop, GH - 310 + 170 / 2, 10, 130, IIf(GetTickCount Mod 300 < 150, argb(255, 0, 0, 0), argb(255, 0, 0, 255)), size:=1, style:=1, Pos:=posOnCenter
        
        Page.Paint 0, GW / 2 - 400 / 2, 70, 400, 20, argb(255, 128, 128, 128)
        If hurtUTime = 0 Then atkPro = 0
        If atkPro > 1 Then atkPro = 1
        
        Page.Paint 0, GW / 2 - 400 / 2, 70, (UHP + hurtU - (hurtU * atkPro)) / 9000 * 400, 20, argb(255, 0, 255, 0)
        If hurtU = 0 Then
            Page.DrawImage "miss.png", GW / 2 - 122 / 2, 20, alpha:=1
        Else
            For i = 1 To Len(hurtU)
                Page.DrawImage Mid(hurtU, i, 1) & ".png", GW / 2 - (Len(hurtU) - 1) * 34 / 2 + (i - 1) * 34, 20, alpha:=1
            Next
        End If
        
        If GetTickCount - hurtUTime >= 2000 And hurtUTime <> 0 Then
            If UHP < 0 Then
                ShowText "冰棍被你杀死了。"
                MusicList.Play "die.wav"
                ChangeFocus 0
                ECore.NewTransform
            Else
                Call ShowRoundText3
                Call BeginAttack
                ChangeFocus 9
                roundTime = GetTickCount ': ShowRoundText
            End If
        End If
    End If
    
    If focus1 = 0 Then
        Page.DrawImage "heart.png", GW / 2 - 330 + 180 * focus2, GH / 2 + 220, alpha:=1, Pos:=posOnCenter
        If IsKeyUp(VK_Z) And UHP > 0 Then
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
            Call BeginAttack
            ChangeFocus 9
            roundTime = GetTickCount ': ShowRoundText
        End If
    End If
    
    If BGM.IsPlaying = False And UHP > 0 Then BGM.Play
    
    If GetTickCount - lastTime >= 60 And Len(displayText) <> Len(TargetText) And focus1 = 0 Then
        lastTime = GetTickCount
        MusicList.Play "say.wav"
        displayText = displayText & Mid(TargetText, Len(displayText) + 1, 1)
    End If
    
    If focus1 = 9 Then
        Call AttackProcess
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
Public Sub Move(dir As Long, Key As Long)
    If UHP < 0 Then Exit Sub
    If focus1 = 1 Or focus1 = 9 Or focus1 = 5 Or focus1 = 6 Then Exit Sub
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
    If msg = "attack_ok" Then
        focus1 = 6
        hurtUTime = GetTickCount
        If UHP < 0 Then
            MusicList.Play "kill.wav"
            BGM.Stops
        Else
            MusicList.Play "shout.wav"
        End If
    End If
End Sub
Public Sub Enter()
    '页面进入事件
    BGM.Create App.path & "\music\Hopes and Dreams.mp3"
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

    '卷轴
    Dim fx As Long, fy As Long, fw As Long, fh As Long
    fx = 46 + 200: fy = GH - 310 - 80
    fw = GW - 46 * 2 - 400: fh = 170 + 80
    ADC = CreateCDC(fw, fh)
    GdipCreateFromHDC ADC, AG
    GdipSetSmoothingMode AG, SmoothingModeAntiAlias
    GdipSetTextRenderingHint AG, TextRenderingHintSingleBitPerPixel
    ODC = Page.CDC: OG = Page.GG

    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
