VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���׼�ʱ��"
   ClientHeight    =   1890
   ClientLeft      =   1755
   ClientTop       =   2505
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6045
   Begin VB.CheckBox Check1 
      Caption         =   "�ö�"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ͣ��ʱ"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��ʱ"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��ʱ�������塣


'Written by zhj.
'Copyright (c) 2009-2012 Createnhance Programs.All rights served.
'ʹ��CC BY-NC-SA 3.0Э����Ȩ�����http://creativecommons.org/licenses/by-nc-sa/3.0/deed.zh


Dim a, b, c As Integer
Dim d, e, f As String

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long



Private Sub Check1_Click()
If Check1.Value = 1 Then ontop Me.hwnd Else behind Me.hwnd
If Command1.Enabled = True Then Command1.SetFocus Else Command2.SetFocus
End Sub

Private Sub Command1_Click()
Command1.Caption = "������ʱ"
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Command2.SetFocus
End Sub
Private Sub Command2_Click()
Command1.Caption = "������ʱ"
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command1.SetFocus
End Sub
Private Sub Command3_Click()
If Command1.Enabled = True Then Command1.SetFocus Else Command2.SetFocus
X = MessageBox(Me.hwnd, "��ʱ�� Version " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) + Chr(13), "���ڼ�ʱ��", 0)
End Sub
Private Sub Command4_Click()
a = 0
b = 0
c = 0
Label1.Caption = "00:00:00"
Timer1.Enabled = False
Command1.Caption = "��ʼ��ʱ"
Command1.Enabled = True
Command2.Enabled = False
Command1.SetFocus
End Sub

Private Sub Form_GotFocus()
If Command1.Enabled = True Then Command1.SetFocus Else Command2.SetFocus
End Sub

Private Sub Form_Initialize()  '��������������ע��ؼ������á�XP��ʽ���Ŀؼ������簴ť����������ܻ��ڸ�����xp��ʽ.res��
InitCommonControls
End Sub

Private Sub Form_Load()
a = 0
b = 0
c = 0
Timer1.Enabled = False
Timer1.Interval = 1000
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If a > 0 Or b > 0 Or c > 0 Then
Command1.Caption = "������ʱ"
Command1.Enabled = True
Command2.Enabled = False

End If
Timer1.Enabled = False
X = MsgBox("��л��ʹ�ü�ʱ����" + Chr(13) + "ȷʵҪ�˳���", vbOKCancel, "�˳�")
Cancel = X - 1
If Cancel = 0 Then End
End Sub

Private Sub Timer1_Timer()
c = c + 1
If c = 60 Then
c = 0
b = b + 1
End If
If b = 60 Then
b = 0
a = a + 1
End If
If a > 99 Then
Label1.Caption = "00:00:00"
X = MsgBox("ʱ���Ѿ�����100Сʱ��" & Chr(13) & "�������ݲ�֧��100Сʱ���ϵļ�ʱ���������ü�ʱ����", vbOKOnly, "����")
a = 0
b = 0
c = 0
Timer1.Enabled = False
Label1.Caption = "00 : 00 : 00"
Command1.Caption = "��ʼ��ʱ"
Command1.Enabled = True
Command2.Enabled = False

End If
d = a
e = b
f = c
If a < 10 Then d = "0" & a
If b < 10 Then e = "0" & b
If c < 10 Then f = "0" & c
Label1.Caption = d & ":" & e & ":" & f
End Sub
