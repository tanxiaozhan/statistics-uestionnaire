VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FDFDFD&
   Caption         =   "��ӭʹ�ñ�ϵͳ"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��������ϵͳ
'���򿪷���lc_mtt
'CSDN���ͣ�http://blog.csdn.net/lc_mtt/
'������ҳ��http://www.3lsoft.com
'���䣺3lsoft@163.com
'ע���˴����ֹ������ҵ��;�����޸��߷���һ�ݣ�лл��
'---------------- ��Դ���磬���Ҹ����� ----------------

Private Sub Form_Load()
    Me.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
    imgBack.Move (Width / 15 - imgBack.Width) / 2, (Height / 15 - imgBack.Height) / 2
End Sub
