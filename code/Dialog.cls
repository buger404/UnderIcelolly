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
'   ҳ�������
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
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear 0 '��ջ���

    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
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
    Page.TopPage = True
    '����ҳ��
    ECore.Add Page, "Dialog"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub