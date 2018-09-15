Attribute VB_Name = "Windowtop"
'本模块用于窗口置顶。执行语句：
'置顶：ontop 窗口名称.hwnd
'取消置顶：behind 窗口名称.hwnd
'具体如何实现的不用看……
'Written by zhj.
'Copyright (c) 2009-2012 Createnhance Programs.All rights served.
'使用CC BY-NC-SA 3.0协议授权。详见http://creativecommons.org/licenses/by-nc-sa/3.0/deed.zh


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'hwnd ----------- Long，欲定位的窗口
'hWndInsertAfter - Long，窗口句柄。在窗口列表中，窗口hwnd会置于这个窗口句柄的后面。也可能选用下述值之一：
Public Const HWND_BOTTOM = 1 '将窗口置于窗口列表底部
Public Const HWND_TOP = 0 '将窗口置于Z序列的顶部；Z序列代表在分级结构中，窗口针对一个给定级别的窗口显示的顺序
Public Const HWND_TOPMOST = -1 '将窗口置于列表顶部，并位于任何最顶部窗口的前面
Public Const HWND_NOTOPMOST = -2 '将窗口置于列表顶部，并位于任何最顶部窗口的后面
'x -------------- Long，窗口新的x坐标。如hwnd是一个子窗口，则x用父窗口的客户区坐标表示
'y -------------- Long，窗口新的y坐标。如hwnd是一个子窗口，则y用父窗口的客户区坐标表示
'cx ------------- Long，指定新的窗口宽度
'cy ------------- Long，指定新的窗口高度
'wFlags --------- Long，包含了旗标的一个整数
Public Const SWP_FRAMECHANGED = &H20 '强迫一条WM_NCCALCSIZE消息进入窗口，即使窗口的大小没有改变
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED '围绕窗口画一个框
Public Const SWP_HIDEWINDOW = &H80 '隐藏窗口
Public Const SWP_NOACTIVATE = &H10 '不激活窗口
Public Const SWP_NOMOVE = &H2 '保持当前位置 (x和y设定将被忽略)
Public Const SWP_NOREDRAW = &H8 '窗口不自动重画
Public Const SWP_NOSIZE = &H1 '保持当前大小 (cx和cy会被忽略)
Public Const SWP_NOZORDER = &H4 '保持窗口在列表的当前位置 (hWndInsertAfter将被忽略)
Public Const SWP_SHOWWINDOW = &H40 '显示窗口

Public Sub ontop(ByVal hwnd As Long)
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW Or SWP_NOMOVE
End Sub

Public Sub behind(ByVal hwnd As Long)
Dim Xbef As Integer
Dim ybef As Integer
Xbef = Form1.Top
ybef = Form1.Left
Form1.Hide
SetWindowPos hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOMOVE
Form1.Top = Xbef
Form1.Left = ybef
Form1.Show
Form1.SetFocus
End Sub
