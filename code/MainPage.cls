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
'   ҳ�������
    Dim Page As GPage
    Dim br As Single, bg As Single, bb As Single, bt As Long
    Dim tr As Single, tg As Single, tb As Single
    Dim OpenTime As Long, BattleTime As Long
    Dim NoLoop As Boolean
'======================================
    
Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
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
    
    Page.Clear argb(255, br, bg, bb) '��ջ���

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
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
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
    
    If GetTickCount - OpenTime >= 8000 And Dialog.diaChar = "" Then
        Dialog.NewDialog "��ã����Ǳ�����", "say.wav", 100
        Dialog.NewDialog "����õ����ѣ�������", "say.wav", 80
        Dialog.NewDialog "...", "say.wav", 200
        Dialog.NewDialog "...", "say.wav", 200
        Dialog.NewDialog "��˵���ϲ��������", "say.wav", 200
        Dialog.NewDialog "��ô", "say.wav", 100
        BGM.Stops: NoLoop = True
        Dialog.NewDialog "�����˵������Ҫɱ�����أ�", "ho.wav", 200
        Dialog.NewDialog "��Ҳ��ܿ��ĵسԵ�����", "ho.wav", 200
        MusicList.Play "battle.wav"
        BattleTime = GetTickCount
    End If
    
    'Page.DrawAnimation "attack", GW / 2, GH / 2 - 80
    
    If (BGM.IsPlaying = False) And (Not NoLoop) Then BGM.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(ID As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.NewImages App.path & "\assets"
    Page.Res.ClipCircle "icebig.png"
    
    Set BGM = New GMusic
    BGM.Create App.path & "\music\Barrier.mp3"
    BGM.Play

    '����ҳ��
    ECore.Add Page, "MainPage"
    
    OpenTime = GetTickCount
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub