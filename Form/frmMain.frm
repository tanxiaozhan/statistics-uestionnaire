VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����ʾ���Ϣ����ϵͳ"
   ClientHeight    =   10530
   ClientLeft      =   6135
   ClientTop       =   2910
   ClientWidth     =   14745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5055
      Top             =   7485
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
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStudentList 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9840
      Left            =   1560
      ScaleHeight     =   9840
      ScaleWidth      =   2025
      TabIndex        =   17
      Top             =   390
      Width           =   2025
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdDel 
         Height          =   330
         Left            =   1005
         TabIndex        =   22
         Top             =   6510
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   582
         Caption         =   "ɾ��"
         BackColor       =   33023
         MouseDownColor  =   64
         MouseOnColor    =   255
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvStudentList 
         Height          =   6045
         Left            =   60
         TabIndex        =   20
         ToolTipText     =   "˫���޸�"
         Top             =   405
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   10663
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdCloseStudentList 
         Height          =   195
         Left            =   1695
         TabIndex        =   19
         Top             =   75
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   344
         Caption         =   "��"
         ToolTip         =   "�ر�"
         BackColor       =   6956042
         ForeColor       =   16777215
         MouseDownColor  =   6956042
         MouseOnColor    =   6956042
         StyleColor      =   0
         Style3dColor1   =   16577259
         Style3dColor2   =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   45
         Picture         =   "frmMain.frx":1998
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ���б�"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   285
         TabIndex        =   18
         Top             =   90
         Width           =   720
      End
      Begin VB.Shape shStudentList 
         BorderColor     =   &H00A6A6A6&
         Height          =   7230
         Left            =   0
         Top             =   345
         Width           =   1965
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H006A240A&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   15
         Top             =   30
         Width           =   1965
      End
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3840
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   983
      TabIndex        =   4
      Top             =   10230
      Width           =   14745
      Begin VB.Image Image2 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":1F22
         Top             =   45
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3150
         Picture         =   "frmMain.frx":22AC
         Top             =   45
         Width           =   240
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   2
         Left            =   3465
         TabIndex        =   10
         Top             =   75
         Width           =   90
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ӭʹ�ñ�ϵͳ"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   9
         Top             =   75
         Width           =   1260
      End
      Begin VB.Shape Shb2 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   3090
         Top             =   30
         Width           =   6885
      End
      Begin VB.Image imgLB 
         Height          =   180
         Left            =   10080
         MousePointer    =   8  'Size NW SE
         Top             =   120
         Width           =   180
      End
      Begin VB.Shape Shb1 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   30
         Top             =   30
         Width           =   3015
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9840
      Left            =   0
      ScaleHeight     =   656
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   1
      Top             =   390
      Width           =   1560
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton XButton2 
         Height          =   885
         Left            =   210
         TabIndex        =   25
         Top             =   6630
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�鿴�ϴ����"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2636
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdLeft 
         Height          =   885
         Index           =   5
         Left            =   210
         TabIndex        =   23
         Top             =   5220
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�ϴ�����"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2F10
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdLeft 
         Height          =   885
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   480
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�  ��"
         ToolTip         =   "��������ʾ���Ϣ"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":37EA
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdClose 
         Height          =   195
         Left            =   1245
         TabIndex        =   3
         Top             =   60
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   344
         Caption         =   "��"
         ToolTip         =   "�ر�"
         BackColor       =   6956042
         ForeColor       =   16777215
         MouseDownColor  =   6956042
         MouseOnColor    =   6956042
         StyleColor      =   0
         Style3dColor1   =   16577259
         Style3dColor2   =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdLeft 
         Height          =   885
         Index           =   3
         Left            =   195
         TabIndex        =   6
         Top             =   2850
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "��  ѯ"
         ToolTip         =   "��������ѯ"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3EE4
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdLeft 
         Height          =   885
         Index           =   2
         Left            =   210
         TabIndex        =   7
         Top             =   1665
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "����¼��"
         ToolTip         =   "�԰�Ϊ��λ¼������ʾ���Ϣ"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":45DE
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton cmdLeft 
         Height          =   885
         Index           =   4
         Left            =   210
         TabIndex        =   8
         Top             =   4035
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�޸�����"
         ToolTip         =   "�޸��û�����"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":52B8
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape ShLeft 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00A6A6A6&
         Height          =   7575
         Left            =   30
         Top             =   330
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   90
         TabIndex        =   2
         Top             =   75
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H006A240A&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   270
         Left            =   30
         Top             =   30
         Width           =   1485
      End
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      BackColor       =   &H00D1D8DB&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   983
      TabIndex        =   0
      Top             =   0
      Width           =   14745
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton XButton3 
         Height          =   270
         Left            =   6540
         TabIndex        =   26
         Top             =   60
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   476
         Caption         =   "�鿴�ϴ����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":5852
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLeft 
         Height          =   270
         Index           =   5
         Left            =   5025
         TabIndex        =   24
         Top             =   60
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   476
         Caption         =   "�ϴ�����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":5DEC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton XButton1 
         Height          =   270
         Left            =   8370
         TabIndex        =   21
         Top             =   75
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   476
         Caption         =   ""
         ToolTip         =   "����"
         BackColor       =   13752539
         MouseDownColor  =   -2147483644
         MouseOnColor    =   -2147483644
         StyleColor      =   0
         Style3dColor1   =   16577259
         Style3dColor2   =   8421504
         Picture         =   "frmMain.frx":67FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLogin 
         Height          =   270
         Left            =   210
         TabIndex        =   11
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         Caption         =   ""
         ToolTip         =   "���ص�½����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":6D98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLeft 
         Height          =   270
         Index           =   2
         Left            =   1620
         TabIndex        =   12
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Caption         =   "¼��"
         ToolTip         =   "�԰�Ϊ��λ¼������ʾ���Ϣ"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":7332
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLeft 
         Height          =   270
         Index           =   3
         Left            =   2655
         TabIndex        =   13
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Caption         =   "��ѯ"
         ToolTip         =   "��������ѯ"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":78CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLeft 
         Height          =   270
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Caption         =   "���"
         ToolTip         =   "��������ʾ���Ϣ"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":7E66
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbLeft 
         Height          =   270
         Index           =   4
         Left            =   3675
         TabIndex        =   15
         Top             =   60
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   476
         Caption         =   "�޸�����"
         ToolTip         =   "�޸��û�����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":8400
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XButton tbExit 
         Height          =   270
         Left            =   8895
         TabIndex        =   16
         Top             =   75
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   476
         Caption         =   ""
         ToolTip         =   "�˳�����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":899A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00A6A6A6&
         X1              =   549
         X2              =   549
         Y1              =   2
         Y2              =   22
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00A6A6A6&
         X1              =   42
         X2              =   42
         Y1              =   3
         Y2              =   23
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileUpload 
         Caption         =   "�����ϴ�"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExItem 
         Caption         =   "��������(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExIncome 
         Caption         =   "��������(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBBackUp 
         Caption         =   "�������ݿ�(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDBResume 
         Caption         =   "�ָ����ݿ�(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "���ص�½����(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSP4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuLeft 
         Caption         =   "�����ѯ(&Q)"
         Index           =   1
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "����¼��(&D)"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "��ѯ(&S)"
         Index           =   3
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "�޸�����(&C)"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuViewSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStudentList 
         Caption         =   "ѧ���б�(&L)"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "������(&G)"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuTB 
         Caption         =   "������(&T)"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSB 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuContent 
         Caption         =   "����(&C)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSupply 
         Caption         =   "����֧��(&S)"
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "���ڱ����(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�϶������API
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TVFIRST = &H1100
Const TVMSETBKCOLOR = TVFIRST + 29

Dim CanResize As Boolean
Public LastFrm As Long
Public ct, cc As Byte

Private Sub cmdAbout_Click()
    mnuAbout_Click
End Sub

Private Sub cmdClose_Click()
    picLeft.Visible = False
    mnuGuide.Checked = False
    SaveINI "Main", "Guide", "n"
End Sub

Private Sub cmdCloseStudentList_Click()
    picStudentList.Visible = False
    mnuStudentList.Checked = False
    SaveINI "Main", "StudentList", "n"

End Sub

Private Sub cmdDel_Click()
    If IsNumeric(Trim(tvStudentList.SelectedItem)) Then
        If MsgBox("ȷʵɾ�����Ϊ " & tvStudentList.SelectedItem & " ���ʾ���", vbYesNo, "ɾ���ʾ�") = vbNo Then Exit Sub
        DBConnect
        Conn.Execute "delete from main where mClass='" & curID & "' and mNo=" & tvStudentList.SelectedItem
        Conn.Close
        Set Conn = Nothing
        tvStudentList.Nodes.Remove (tvStudentList.SelectedItem.Index)
    End If
End Sub

Public Sub cmdLeft_Click(Index As Integer)
    
    If LastFrm = Index And Index < 5 Then Exit Sub
    If LastFrm > 0 Then
        cmdLeft(LastFrm).IfDraw = False
        tbLeft(LastFrm).IfDraw = False
        If LastFrm < 5 Then mnuLeft(LastFrm).Checked = False
        cmdLeft(LastFrm).BackColor = picLeft.BackColor
        tbLeft(LastFrm).BackColor = picTB.BackColor
    Else
        'Unload frmList
    End If
    
    Select Case LastFrm
        Case 1:
        Case 2:
                Select Case curInputForm
                    Case 1:
                        Unload frmInputMainSCL
                    Case 2:
                        Unload frmInputMainSCH
                    Case 3:
                        Unload frmInputMainMS
                End Select
                    
                picStudentList.Visible = False
                mnuStudentList.Checked = False
                picLeft.Visible = True
                mnuGuide.Checked = True

        
        Case 3:
    End Select
    
    LastFrm = Index
    cmdLeft(Index).IfDraw = True
    tbLeft(Index).IfDraw = True
    If Index < 5 Then mnuLeft(Index).Checked = True
    cmdLeft(Index).BackColor = 14210516
    tbLeft(Index).BackColor = 14210516
    SetSB 1, "����λ�ã�" & cmdLeft(Index).caption
    
    
    
    Select Case Index
        Case 1:
        Case 2: picStudentList.Visible = True
                mnuStudentList.Checked = True
                picLeft.Visible = False
                mnuGuide.Checked = False
                frmInputNo.Show vbModal, frmMain
        Case 3:
        Case 4:
                frmChangePWD.Show vbModal, frmMain
        Case 5:
                mnuFileUpload_Click
                
        
    End Select
    
End Sub

Private Sub imgLB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(hwnd, &HA1, 17, 0)
    End If
End Sub

Private Sub MDIForm_Load()
    Dim rs As ADODB.Recordset
    
    '��ȡ����λ��,��ͼ��Ϣ
    If GetINI("Main", "Left") = "" Then
        Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Else
        Move GetLongINI("Main", "Left"), GetLongINI("Main", "Top"), GetLongINI("Main", "Width"), GetLongINI("Main", "Height")
        Dim j As Long
        j = GetLongINI("Main", "WindowState")
        If j = 2 Then Me.WindowState = 2
    End If
    CanResize = True
    If GetINI("Main", "Guide") = "n" Then
        picLeft.Visible = False
        mnuGuide.Checked = False
    End If
    If GetINI("Main", "ToolBar") = "n" Then
        picTB.Visible = False
        mnuTB.Checked = False
    End If
    If GetINI("Main", "StateBar") = "n" Then
        picSB.Visible = False
        mnuSB.Checked = False
    End If
    If GetINI("Main", "StudentList") = "n" Then
        picStudentList.Visible = False
        mnuStudentList.Checked = False
    End If
    
    LastFrm = 0
    
    Me.tvStudentList.Nodes.Add , , , curClass, 1
    Me.tvStudentList.Nodes.Item(1).Selected = True
    
    
    DBConnect
    Set rs = New ADODB.Recordset
    rs.Open "select mNo from main where mClass='" & curID & "'  order by mNo", Conn, 1, 1
    
    Do While Not rs.EOF
        If rs("mNo") < 10 Then
            insertNo = "0" & Trim(rs("mNo"))
        Else
            insertNo = Trim(rs("mNo"))
        End If
        
        Me.tvStudentList.Nodes.Add Me.tvStudentList.SelectedItem, tvwChild, , insertNo, 2, 3
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
    
    expandeAll
    
    SetSB 2, "ѧУ��" & curSchool & "     �༶��" & curClass
        
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    If CanResize = False Then Exit Sub
    If Me.Width < 9900 Then Me.Width = 9900
    If Me.Height < 8370 Then Me.Height = 8370
    SaveINI "Main", "WindowState", CStr(WindowState)
    If Me.WindowState = 0 Then
        SaveINI "Main", "Width", CStr(Width)
        SaveINI "Main", "Height", CStr(Height)
    End If
    picSB_Resize
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
End Sub

Private Sub mnuAbout_Click()
    MsgBox "�����ʾ���Ϣ����ϵͳ���� V1.0" & Chr(13) & Chr(13) & "    2009.03", vbInformation
End Sub

Private Sub mnuContent_Click()
    MsgBox "���ް���������£�", vbInformation
End Sub

Private Sub mnuDBBackUp_Click()
    On Error GoTo errmsg
    
    If Conn.State <> 0 Then
        Conn.Close
    End If
    
    If DirExists(GetApp & "bak") = 0 Then
        MkDir GetApp & "bak"
    End If
    
    Dlg.Filter = "�����ʾ���Ϣ����ϵͳ�����ļ�(*.htb)|*.htb"
    Dlg.FileName = "DATA" & Format(Now(), "yyyy-mm-dd hh.mm.ss") & ".htb"
    Dlg.DialogTitle = "���ݱ���"
    Dlg.InitDir = GetApp & "bak"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    FileCopy GetApp & "\data\data.qta", Dlg.FileName
    MsgBox "���ݱ��ݳɹ���", vbInformation, "���ݱ���"
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݱ���"
End Sub

Private Sub mnuDBResume_Click()
    On Error GoTo errmsg
    
    Dim rs As ADODB.Recordset
    
    If Conn.State <> 0 Then
        Conn.Close
    End If
    If DirExists(GetApp & "bak") <> 0 Then
        Dlg.InitDir = GetApp & "bak"
    End If
    
    Dlg.Filter = "�����ʾ���Ϣ����ϵͳ�����ļ�(*.htb)|*.htb"
    Dlg.DialogTitle = "���ݻָ�"
    Dlg.CancelError = True
    Dlg.ShowOpen
    
    If MsgBox("���棺���ݻָ�����" & Dlg.FileName & "�����ݸ������������ݡ�", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    If MsgBox("ȷ�Ͻ������ݻָ���?", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    FileCopy Dlg.FileName, GetApp & "\data\data.qta"
    MsgBox "���ݻָ��ɹ���", vbInformation, "���ݻָ�"
    
    
    Me.tvStudentList.Nodes.Clear
    Me.tvStudentList.Nodes.Add , , , curClass, 1
    Me.tvStudentList.Nodes.Item(1).Selected = True

    DBConnect
    Set rs = New ADODB.Recordset
    rs.Open "select mNo from main order by mNo", Conn, 1, 1
    
    Do While Not rs.EOF
        If rs("mNo") < 10 Then
            insertNo = "0" & Trim(rs("mNo"))
        Else
            insertNo = Trim(rs("mNo"))
        End If
        
        Me.tvStudentList.Nodes.Add Me.tvStudentList.SelectedItem, tvwChild, , insertNo, 2, 3
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
    
    expandeAll
    
    
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݻָ�"

End Sub

Private Sub mnuExIncome_Click()
    On Error GoTo errmsg
    
    Exit Sub
    
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "����"

End Sub

Private Sub mnuExit_Click()
    'frmList.SaveListColWidth
    Unload Me
End Sub

Private Sub mnuExItem_Click()
    On Error GoTo errmsg
    
    Exit Sub
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "��Ŀ����(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "������Ŀ����"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "����"



End Sub

Private Sub mnuFileUpload_Click()
    frmUpload.Show vbModal, frmMain
End Sub

Private Sub mnuGuide_Click()
    mnuGuide.Checked = Not mnuGuide.Checked
    picLeft.Visible = mnuGuide.Checked
    SaveINI "Main", "Guide", IIf(mnuGuide.Checked = True, "", "n")
End Sub

Private Sub mnuLogin_Click()
On Error Resume Next
    Unload Me
    frmLogin.Show
End Sub

Private Sub mnuStudentList_Click()
    mnuStudentList.Checked = Not mnuStudentList.Checked
    picStudentList.Visible = mnuStudentList.Checked
    SaveINI "Main", "StudentList", IIf(mnuStudentList.Checked = True, "", "n")

End Sub

Private Sub mnuSupply_Click()
    MsgBox "���µ磺31304837", vbInformation
End Sub

Private Sub picSB_Resize()
On Error Resume Next
    Shb2.Width = Me.Width / 15 - IIf(Me.WindowState = 2, 210, 230)
    imgLB.Visible = (Me.WindowState <> 2)
    imgLB.Left = Me.Width / 15 - 20
End Sub

Private Sub mnuSB_Click()
    mnuSB.Checked = Not mnuSB.Checked
    picSB.Visible = mnuSB.Checked
    SaveINI "Main", "StateBar", IIf(mnuSB.Checked = True, "", "n")
End Sub

Private Sub mnuTB_Click()
    mnuTB.Checked = Not mnuTB.Checked
    picTB.Visible = mnuTB.Checked
    SaveINI "Main", "ToolBar", IIf(mnuTB.Checked = True, "", "n")
End Sub

Private Sub picLeft_Resize()
On Error Resume Next
    ShLeft.Height = picLeft.Height / 15 - 23
    shStudentList.Height = picLeft.Height - 25 * 15
    tvStudentList.Height = shStudentList.Height * 0.96
    cmdDel.Top = tvStudentList.Height + tvStudentList.Top
        
End Sub

Private Sub tbExit_Click()
    mnuExit_Click
End Sub

Private Sub tbLeft_Click(Index As Integer)
    cmdLeft_Click Index
End Sub

Private Sub tbLogin_Click()
    mnuLogin_Click
End Sub


Private Sub XButton2_Click()
    frmUploadResult.Show vbModal, frmMain
    
    
End Sub

Private Sub XButton3_Click()
    
    Me.tvStudentList.Nodes.Add Me.tvStudentList.SelectedItem, tvwChild, , cc
    cc = cc + 1
End Sub

Private Sub XButton4_Click()
    expandeAll

End Sub

Private Sub expandeAll()
          Dim nodex     As Node
          For Each nodex In tvStudentList.Nodes
                  nodex.Expanded = True
          Next
  End Sub
    
  Private Sub closeAll()
          Dim nodex     As Node
          For Each nodex In tvStudentList.Nodes
                  nodex.Expanded = False
          Next
  End Sub

Private Sub XButton5_Click()
    closeAll
End Sub

Private Sub tvStudentList_DblClick()
    Dim rs As ADODB.Recordset
    Dim strAnwser As String     '�ʾ��������
    Dim i As Byte
    Dim selNo As String
    selNo = tvStudentList.SelectedItem
    If Not IsNumeric(selNo) Then Exit Sub
    curNo = selNo
    
    '�����ݿ��ȡ��������
    DBConnect
    Set rs = New ADODB.Recordset
    rs.Open "select * from main where mClass='" & curID & "' and mNo=" & curNo, Conn, 1, 1
        
    If rs.RecordCount <> 1 Then
        MsgBox "��ȡ�ʾ������Ϣ����", vbCritical, "�޸��ʾ�"
        Exit Sub
    End If
        
    strAnwser = rs("mAnwser")
    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
        
        
    Select Case curInputForm
        Case 1:
            frmInputMainSCL.Show     'Сѧ1-3�꼶
            FillAnswer frmInputMainSCL, strAnwser
        Case 2:
            frmInputMainSCH.Show     'Сѧ4-6�꼶
            FillAnswer frmInputMainSCH, strAnwser
        Case 3:
            frmInputMainMS.Show      '���С�����
            FillAnswer frmInputMainMS, strAnwser
    End Select
        
    IsEdit = True     '�޸��ʾ����

End Sub

Private Sub CreateUser()
                
                strsql = "insert into userInfo(uID,uPWD,uSchool,uClass) " & " " & _
                               "values('1101101','" & GetMD5("123456") & "','��������԰Сѧ','һ(1)��')"
                DBConnect
                Conn.Execute strsql
                Conn.Close
                Set Conn = Nothing

End Sub

Private Sub XButton1_Click()
    MsgBox "�����ʾ����ݴ���ϵͳ", vbInformation, "����"
End Sub

Private Sub FillAnswer(frm As Form, strAnswer As String)
    
    For Each optionCtl In frm
        
        If TypeOf optionCtl Is OptionButton Then
            
            ctlname = optionCtl.Name
            ctlname = Right(ctlname, (Len(ctlname) - 6))
            If optionCtl.Index = Asc(Mid(strAnswer, ctlname, 1)) - 65 + 1 Then optionCtl.value = True
        End If
    
    Next

End Sub
