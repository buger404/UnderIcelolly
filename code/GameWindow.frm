VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IcelollySnake"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面类模块声明
    Dim MainPage As New MainPage
    Dim BattlePage As New BattlePage
    Dim GameOverPage As New GameOverPage
'==================================================

Private Sub DrawTimer_Timer()
    '绘制
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) And ECore.ActivePage = "BattlePage" Then
        BattlePage.Move IIf(KeyCode = vbKeyLeft, -1, 1), 0
    End If
End Sub

Private Sub Form_Load()
    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, 805, 556
    '创建字体
    MakeFont "宋体"
    '创建页面管理器
    Set EC = New GMan
    
    '创建存档（可选）
    Set ESave = New GSaving
    ESave.Create "icelollysnake", "55AA22CCICESNAICESN99"
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\sound"

    '开始显示
    Me.Show
    DrawTimer.Enabled = True
    
    '在此处初始化你的页面
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set MainPage = New MainPage
        Set BattlePage = New BattlePage
        Set GameOverPage = New GameOverPage
        Set Dialog = New Dialog
    '=============================================

    '设置活动页面
    EC.ActivePage = "MainPage"
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
    End
End Sub
