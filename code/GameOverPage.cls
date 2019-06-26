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
'   页面绘制器
    Dim Page As GPage
    Dim openTime As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim pro As Single
    pro = (GetTickCount - openTime) / 1000
    If pro > 1 Then pro = 1
    If pro < 0 Then pro = 0
    
    Page.DrawImage "gameover.png", GW / 2, 150, alpha:=pro, Pos:=posOnCenter
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
    openTime = GetTickCount
    BGM.Create App.path & "\music\GameOver.mp3"
    BGM.Play
    Dialog.NewDialog "你会没事的，保持你的决心！", "say.wav", 60
    BGM.Stops
    Dialog.NewDialog "但你永远也不会醒来。", "ho.wav", 60
    End
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
    '创建页面
    ECore.Add Page, "GameOverPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
