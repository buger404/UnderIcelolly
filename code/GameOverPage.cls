VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameOverPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   ҳ�������
    Dim Page As GPage
    Dim openTime As Long
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear argb(255, 0, 0, 0) '��ջ���

    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim pro As Single
    pro = (GetTickCount - openTime) / 1000
    If pro > 1 Then pro = 1
    If pro < 0 Then pro = 0
    
    Page.DrawImage "gameover.png", GW / 2, 150, alpha:=pro, Pos:=posOnCenter
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
    openTime = GetTickCount
    BGM.Create App.path & "\music\GameOver.mp3"
    BGM.Play
    Dialog.NewDialog "���û�µģ�������ľ��ģ�", "say.wav", 60
    BGM.Stops
    Dialog.NewDialog "������ԶҲ����������", "ho.wav", 60
    End
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
    '����ҳ��
    ECore.Add Page, "GameOverPage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub