VERSION 5.00
Begin VB.Form frmChangePWD 
   Caption         =   "�޸�����"
   ClientHeight    =   2895
   ClientLeft      =   1485
   ClientTop       =   1350
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Icon            =   "frmChangePWD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5250
   StartUpPosition =   1  '����������
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtNewconf 
      Height          =   300
      Left            =   1500
      TabIndex        =   4
      Top             =   1305
      Width           =   2595
      _extentx        =   4577
      _extenty        =   529
      font            =   "frmChangePWD.frx":058A
      fontname        =   "����"
      fontsize        =   9
      passwordchar    =   "*"
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtNew 
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      Top             =   825
      Width           =   2595
      _extentx        =   4577
      _extenty        =   529
      font            =   "frmChangePWD.frx":05AE
      fontname        =   "����"
      fontsize        =   9
      passwordchar    =   "*"
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtOld 
      Height          =   300
      Left            =   1500
      TabIndex        =   0
      Top             =   330
      Width           =   2595
      _extentx        =   4577
      _extenty        =   529
      font            =   "frmChangePWD.frx":05D2
      fontname        =   "����"
      fontsize        =   9
      passwordchar    =   "*"
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton XPButton2 
      Height          =   405
      Left            =   3015
      TabIndex        =   7
      Top             =   2115
      Width           =   1245
      _extentx        =   2196
      _extenty        =   714
      caption         =   " ��  �� "
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmChangePWD.frx":05F6
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton XPButton1 
      Height          =   405
      Left            =   1050
      TabIndex        =   6
      Top             =   2115
      Width           =   1245
      _extentx        =   2196
      _extenty        =   714
      caption         =   " ��  �� "
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmChangePWD.frx":061A
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.Label Label3 
      Caption         =   "ȷ�����룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "�����룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   645
      TabIndex        =   3
      Top             =   855
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "ԭ���룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   645
      TabIndex        =   2
      Top             =   360
      Width           =   930
   End
End
Attribute VB_Name = "frmChangePWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub XPButton1_Click()
    Dim rs As ADODB.Recordset
    Dim sql As String
    If txtOld.Text = "" Then
        MsgBox "����������룡", vbExclamation, "�޸�����"
        txtOld.SetFocus
        Exit Sub
    End If
    If txtNew.Text = "" Then
        MsgBox "�����������룡", vbExclamation, "�޸�����"
        txtNew.SetFocus
        Exit Sub
    End If
    If txtNew.Text <> txtNewconf.Text Then
        MsgBox "������������벻��ͬ��", vbExclamation, "�޸�����"
        txtNewconf.SetFocus
        Exit Sub
    End If
    
    DBConnect
    
    sql = "select * from userInfo where uID='" & curID & "' and uPWD='" & GetMD5(txtOld.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, 1, 1
    recc = rs.RecordCount
    rs.Close
    Set rs = Nothing
    If recc <> 1 Then
        MsgBox "���������", vbCritical, "�޸�����"
        Exit Sub
    End If
    
    Conn.Execute "update userInfo set uPWD='" & GetMD5(txtNew.Text) & "' where uID='" & curID & "'"
    MsgBox "�����޸ĳɹ���", vbInformation, "�޸�����"
    Unload Me
    
    frmMain.cmdLeft_Click 1

End Sub

Private Sub XPButton2_Click()
    Unload Me
    frmMain.cmdLeft_Click 1
    
End Sub
