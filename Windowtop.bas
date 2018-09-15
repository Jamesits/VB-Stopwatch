Attribute VB_Name = "Windowtop"
'��ģ�����ڴ����ö���ִ����䣺
'�ö���ontop ��������.hwnd
'ȡ���ö���behind ��������.hwnd
'�������ʵ�ֵĲ��ÿ�����
'Written by zhj.
'Copyright (c) 2009-2012 Createnhance Programs.All rights served.
'ʹ��CC BY-NC-SA 3.0Э����Ȩ�����http://creativecommons.org/licenses/by-nc-sa/3.0/deed.zh


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'hwnd ----------- Long������λ�Ĵ���
'hWndInsertAfter - Long�����ھ�����ڴ����б��У�����hwnd������������ھ���ĺ��档Ҳ����ѡ������ֵ֮һ��
Public Const HWND_BOTTOM = 1 '���������ڴ����б�ײ�
Public Const HWND_TOP = 0 '����������Z���еĶ�����Z���д����ڷּ��ṹ�У��������һ����������Ĵ�����ʾ��˳��
Public Const HWND_TOPMOST = -1 '�����������б�������λ���κ�������ڵ�ǰ��
Public Const HWND_NOTOPMOST = -2 '�����������б�������λ���κ�������ڵĺ���
'x -------------- Long�������µ�x���ꡣ��hwnd��һ���Ӵ��ڣ���x�ø����ڵĿͻ��������ʾ
'y -------------- Long�������µ�y���ꡣ��hwnd��һ���Ӵ��ڣ���y�ø����ڵĿͻ��������ʾ
'cx ------------- Long��ָ���µĴ��ڿ��
'cy ------------- Long��ָ���µĴ��ڸ߶�
'wFlags --------- Long������������һ������
Public Const SWP_FRAMECHANGED = &H20 'ǿ��һ��WM_NCCALCSIZE��Ϣ���봰�ڣ���ʹ���ڵĴ�Сû�иı�
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED 'Χ�ƴ��ڻ�һ����
Public Const SWP_HIDEWINDOW = &H80 '���ش���
Public Const SWP_NOACTIVATE = &H10 '�������
Public Const SWP_NOMOVE = &H2 '���ֵ�ǰλ�� (x��y�趨��������)
Public Const SWP_NOREDRAW = &H8 '���ڲ��Զ��ػ�
Public Const SWP_NOSIZE = &H1 '���ֵ�ǰ��С (cx��cy�ᱻ����)
Public Const SWP_NOZORDER = &H4 '���ִ������б�ĵ�ǰλ�� (hWndInsertAfter��������)
Public Const SWP_SHOWWINDOW = &H40 '��ʾ����

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
