VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   AutoRedraw      =   -1  'True
   Caption         =   "�û�����"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   ControlBox      =   0   'False
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView List1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�û�ID"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�û���"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "�û�����"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":10BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin �����ʾ���Ϣ����ϵͳ.XPButton cmdDel 
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "ɾ��(&D)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin �����ʾ���Ϣ����ϵͳ.XPButton cmdEdit 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "�޸�(&E)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin �����ʾ���Ϣ����ϵͳ.XPButton cmdAdd 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "���(&A)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   45
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   452
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   45
      Width           =   6780
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   60
         Top             =   -15
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   11
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Frame freItem 
      Height          =   2895
      Left            =   240
      TabIndex        =   12
      Top             =   1140
      Visible         =   0   'False
      Width           =   4380
      Begin �����ʾ���Ϣ����ϵͳ.FCombo cboStyle 
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   1800
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnabledText     =   0   'False
         ListIndex       =   -1
      End
      Begin �����ʾ���Ϣ����ϵͳ.FTextBox txtPW 
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "����"
         FontSize        =   9
         PasswordChar    =   "*"
         AutoSelAll      =   -1  'True
      End
      Begin �����ʾ���Ϣ����ϵͳ.FTextBox txtUser 
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "����"
         FontSize        =   9
         AutoSelAll      =   -1  'True
      End
      Begin �����ʾ���Ϣ����ϵͳ.FTextBox txtPW2 
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "����"
         FontSize        =   9
         PasswordChar    =   "*"
         AutoSelAll      =   -1  'True
      End
      Begin �����ʾ���Ϣ����ϵͳ.XPButton cmdExit 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   2940
         TabIndex        =   9
         Top             =   2310
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "ȡ��"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin �����ʾ���Ϣ����ϵͳ.XPButton cmdOK 
         Default         =   -1  'True
         Height          =   345
         Left            =   1740
         TabIndex        =   8
         Top             =   2310
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "���"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lbPW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���벻������"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  �룺"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   915
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û�����"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ  �ϣ�"
         Height          =   180
         Left            =   360
         TabIndex        =   14
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  �ͣ�"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   1875
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    cmdOK.caption = "���"
    freItem.caption = " ����û� "
    txtUser.Text = ""
    txtPW.Text = ""
    txtPW2.Text = ""
    lbPW.Visible = False
    LoadcboStyle
    cboStyle.ListIndex = 0
    cboStyle.Enabled = True
    ShowItemFrame True
    txtUser.SetFocus
End Sub

Private Sub cmdDel_Click()
On Error GoTo aaaa
    Dim Item As ListItem
    Set Item = List1.SelectedItem
    Dim j As Long
    j = CLng(Left$(Item.SubItems(2), 1))
    If j <= curUserStyle Then
        MsgBox "��û��Ȩ��ɾ�����û���", vbExclamation
        List1.SetFocus
        Exit Sub
    End If
    If StrComp(curUserName, Item.SubItems(1), 1) = 0 Then
        MsgBox "����ɾ���Լ���", vbInformation
        Exit Sub
    End If
    
    If MsgBox("ȷ��ɾ������û��� [" & Mid$(Item.SubItems(2), 3) & "] " & Item.SubItems(1), vbInformation + vbOKCancel) = vbCancel Then Exit Sub
    DBConnect
    Conn.Execute "Delete From Userinfo Where Userid='" & Item.SubItems(1) & "'"
    SetSB 2, "ɾ���û� " & Item.SubItems(1) & " �ɹ�."
    List1.ListItems.Remove Item.Index
    List1.SetFocus
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
On Error GoTo aaaa
    Dim Item As ListItem
    Set Item = List1.SelectedItem
    Dim j As Long
    j = CLng(Left$(Item.SubItems(2), 1))
    If j <= curUserStyle Then
        MsgBox "��û��Ȩ�ޱ༭���û���", vbExclamation
        List1.SetFocus
        Exit Sub
    End If
    If StrComp(curUserName, Item.SubItems(1), 1) = 0 Then cboStyle.Enabled = False
    
    txtUser.Text = Item.SubItems(1)
    txtUser.Tag = Item.SubItems(1)
    txtPW.Text = ""
    txtPW2.Text = ""
    LoadcboStyle
    cboStyle.ListIndex = j - 1
    
    lbPW.Visible = True
    cmdOK.caption = "�޸�"
    freItem.caption = " �޸��û� "
    ShowItemFrame True
    txtUser.SetFocus
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdExit_Click()
    ShowItemFrame False
    List1.SetFocus
End Sub

Private Sub cmdOK_Click()
'On Error GoTo aaaa
    If txtUser.Text = "" Then
        MsgBox "������д�û�����", vbInformation
        txtUser.SetFocus
        Exit Sub
    End If
    If cmdOK.caption = "���" Then
        If txtPW.Text = "" Then
            MsgBox "������д���롣", vbInformation
            txtPW.SetFocus
            Exit Sub
        End If
        If txtPW2.Text = "" Then
            MsgBox "������дȷ�����롣", vbInformation
            txtPW2.SetFocus
            Exit Sub
        End If
    End If
    If txtPW.Text <> txtPW2.Text Then
        MsgBox "����ǰ��һ�¡�", vbInformation
        txtPW2.SetFocus
        Exit Sub
    End If
    
    DBConnect
    
    If cmdOK.caption = "���" Then
        Conn.Execute "insert into Userinfo(userid,pwd,levelN) values('" & txtUser.Text & "','" & GetMD5(txtPW.Text) & "'," & CStr(cboStyle.ListIndex + 1) & ")"
        LoadUserList
        SetSB 2, "����û� " & txtUser.Text & " �ɹ�."
    Else
        If txtPW.Text = "" Then
            Conn.Execute "UPDATE Userinfo SET Userid='" & txtUser.Text & "',levelN=" & CStr(cboStyle.ListIndex + 1) & " Where Userid='" & txtUser.Tag & "'"
        Else
            Conn.Execute "UPDATE Userinfo SET Userid='" & txtUser.Text & "',PWD='" & GetMD5(txtPW.Text) & "',levelN=" & CStr(cboStyle.ListIndex + 1) & " Where Userid='" & txtUser.Tag & "'"
        End If
        List1.SelectedItem.SubItems(1) = txtUser.Text
        List1.SelectedItem.SubItems(2) = cboStyle.Text
        SetSB 2, "�޸��û� " & txtUser.Text & " �ɹ�."
    End If
    
    cmdExit_Click
Exit Sub
aaaa:
    MsgBox "����ʧ�ܣ������Ǹ��û����Ѿ����ڣ�", vbCritical
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    imgIcon.Picture = frmMain.cmdLeft(6).Picture
    '��ȡ�û������б�
    LoadUserList
    
    SetCmdState
End Sub

'����cboStyle
Private Sub LoadcboStyle()
    Dim i As Long
    cboStyle.Clear
    For i = 1 To 4
        If i <= 2 Or curUserStyle = 4 Then cboStyle.AddItem i & "��" & GetUserStyleString(i)
    Next
End Sub

'��ȡ�û������б�
Public Sub LoadUserList()
    Dim Item As ListItem, lngUserStyle As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    DBConnect
    List1.ListItems.Clear
    rs.Open "Select * From Userinfo order by UserID Desc", Conn, 1, 1
    iCount = 0
    Do Until rs.EOF
        iCount = iCount + 1
        lngUserStyle = rs("leveln")
        Set Item = List1.ListItems.Add(, , iCount, , lngUserStyle)
        Item.SubItems(1) = rs("Userid")
        Item.SubItems(2) = lngUserStyle & "��" & GetUserStyleString(lngUserStyle)
        rs.MoveNext
    Loop
    SetSB 2, "�� " & rs.RecordCount & " ���û�Ա��¼."
End Sub

Public Function GetUserStyleString(ByVal lngUserStyle As Long) As String
    Select Case lngUserStyle
    Case 1
        GetUserStyleString = "����Ա"
    Case 2
        GetUserStyleString = "��ͨ�û�"
    Case 3
        GetUserStyleString = "�м�����Ա"
    Case 4
        GetUserStyleString = "�߼�����Ա"
    End Select
End Function

Public Sub ShowItemFrame(ByVal b As Boolean)
    List1.Visible = Not b
    freItem.Visible = b
    cmdDel.Enabled = Not b
    cmdEdit.Enabled = Not b
    cmdAdd.Enabled = Not b
End Sub

Private Sub Form_Resize()
On Error Resume Next
    List1.Width = Width / 15 - 40
    List1.Height = Height / 15 - 116
    PicTop.Width = Width / 15 - 16
    Cls
    Line (2, 2)-(Width / 15 - 14, Height / 15 - 29), 10921638, B
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    With List1
        If (ColumnHeader.Index - 1) = .SortKey Then
            .SortOrder = 1 - .SortOrder
            .Sorted = True
        Else
            .Sorted = False
            .SortOrder = 0
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End If
    End With
End Sub

Private Sub List1_DblClick()
On Error GoTo aaaa
    Dim j As Long
    j = List1.SelectedItem.Index
    cmdEdit_Click
aaaa:
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo aaaa
    If KeyCode = vbKeyDelete Then
        Dim j As Long
        j = List1.SelectedItem.Index
        cmdDel_Click
    End If
aaaa:
End Sub
Sub SetCmdState()
    frmMain.cmdLeft(3).Enabled = False
    frmMain.cmdLeft(4).Enabled = False
    frmMain.cmdLeft(5).Enabled = False
    
    frmMain.tbLeft(3).Enabled = False
    frmMain.tbLeft(4).Enabled = False
    frmMain.tbLeft(5).Enabled = False
    
    frmMain.mnuLeft(3).Enabled = False
    frmMain.mnuLeft(4).Enabled = False
    frmMain.mnuLeft(5).Enabled = False
    
End Sub
