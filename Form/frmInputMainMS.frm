VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputMainMS 
   Caption         =   "�����޸�"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   -1620
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "ѧ�����˼���ͥ�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Index           =   1
      Left            =   12390
      TabIndex        =   1
      Top             =   8385
      Width           =   9270
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   1
         Left            =   6915
         TabIndex        =   62
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   645
         Index           =   7
         Left            =   525
         TabIndex        =   57
         Top             =   6255
         Width           =   7830
         Begin VB.OptionButton Option7 
            Caption         =   "C���������������ֵܽ���"
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
            Index           =   3
            Left            =   4485
            TabIndex        =   60
            Top             =   375
            Width           =   2865
         End
         Begin VB.OptionButton Option7 
            Caption         =   "B����һ���ֵܽ���"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   59
            Top             =   375
            Width           =   2160
         End
         Begin VB.OptionButton Option7 
            Caption         =   "A��������Ů"
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
            Index           =   1
            Left            =   375
            TabIndex        =   58
            Top             =   375
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "7�����ڼ��е����: ��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   45
            TabIndex        =   61
            Top             =   60
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   8
         Left            =   435
         TabIndex        =   51
         Top             =   7065
         Width           =   7620
         Begin VB.OptionButton Option8 
            Caption         =   "A����������ĸһ������"
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
            Index           =   1
            Left            =   465
            TabIndex        =   55
            Top             =   420
            Width           =   2610
         End
         Begin VB.OptionButton Option8 
            Caption         =   "B����ĸ�����ɥ��ɥĸ�������ͥ"
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
            Index           =   2
            Left            =   3150
            TabIndex        =   54
            Top             =   420
            Width           =   3600
         End
         Begin VB.OptionButton Option8 
            Caption         =   "D�����׼�ͥ"
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
            Index           =   4
            Left            =   6150
            TabIndex        =   53
            Top             =   720
            Width           =   1620
         End
         Begin VB.OptionButton Option8 
            Caption         =   "C�����ڸ�ĸ��߸����游ĸ�����游ĸ������һ������"
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
            Index           =   3
            Left            =   450
            TabIndex        =   52
            Top             =   720
            Width           =   5550
         End
         Begin VB.Label Label4 
            Caption         =   "8����ļ�ͥ���:��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   150
            TabIndex        =   56
            Top             =   105
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   6
         Left            =   465
         TabIndex        =   45
         Top             =   5460
         Width           =   8340
         Begin VB.OptionButton Option6 
            Caption         =   "C��3-8��"
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
            Index           =   3
            Left            =   3735
            TabIndex        =   49
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton Option6 
            Caption         =   "D��3������"
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
            Index           =   4
            Left            =   5265
            TabIndex        =   48
            Top             =   330
            Width           =   1350
         End
         Begin VB.OptionButton Option6 
            Caption         =   "B��8-20��"
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
            Index           =   2
            Left            =   2250
            TabIndex        =   47
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton Option6 
            Caption         =   "A��20������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   46
            Top             =   330
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "6�����ͥ���������룺(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   45
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   5
         Left            =   450
         TabIndex        =   38
         Top             =   4545
         Width           =   6195
         Begin VB.OptionButton Option5 
            Caption         =   "A���о���������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   42
            Top             =   360
            Width           =   1965
         End
         Begin VB.OptionButton Option5 
            Caption         =   "B����ѧ"
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
            Index           =   2
            Left            =   2520
            TabIndex        =   41
            Top             =   360
            Width           =   1110
         End
         Begin VB.OptionButton Option5 
            Caption         =   "D��Сѧ"
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
            Index           =   4
            Left            =   5025
            TabIndex        =   40
            Top             =   345
            Width           =   1260
         End
         Begin VB.OptionButton Option5 
            Caption         =   "C����ѧ"
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
            Index           =   3
            Left            =   3765
            TabIndex        =   39
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "5����ĸ�׵�ѧ����(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   45
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   645
         Index           =   4
         Left            =   465
         TabIndex        =   32
         Top             =   3630
         Width           =   6360
         Begin VB.OptionButton Option4 
            Caption         =   "C����ѧ"
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
            Index           =   3
            Left            =   3930
            TabIndex        =   36
            Top             =   345
            Width           =   1155
         End
         Begin VB.OptionButton Option4 
            Caption         =   "D��Сѧ"
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
            Index           =   4
            Left            =   5370
            TabIndex        =   35
            Top             =   345
            Width           =   1260
         End
         Begin VB.OptionButton Option4 
            Caption         =   "B����ѧ"
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
            Index           =   2
            Left            =   2625
            TabIndex        =   34
            Top             =   345
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "A���о���������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   33
            Top             =   330
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "4���㸸�׵�ѧ����(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   105
            TabIndex        =   37
            Top             =   75
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   1
         Left            =   570
         TabIndex        =   7
         Top             =   855
         Width           =   8130
         Begin VB.OptionButton Option1 
            Caption         =   "F���붫��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   330
            TabIndex        =   200
            Top             =   495
            Width           =   1290
         End
         Begin VB.OptionButton Option1 
            Caption         =   $"frmInputMainMS.frx":0000
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
            Index           =   10
            Left            =   5595
            TabIndex        =   16
            Top             =   495
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            Caption         =   "I�����޽�"
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
            Index           =   9
            Left            =   4275
            TabIndex        =   15
            Top             =   495
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "H��������"
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
            Index           =   8
            Left            =   2955
            TabIndex        =   14
            Top             =   495
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "G���ϸڽ�"
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
            Index           =   7
            Left            =   1650
            TabIndex        =   13
            Top             =   480
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "E����ɽ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   5595
            TabIndex        =   12
            Top             =   255
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "A�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   330
            TabIndex        =   11
            Top             =   255
            Width           =   1290
         End
         Begin VB.OptionButton Option1 
            Caption         =   "B�����ҽ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1650
            TabIndex        =   10
            Top             =   255
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "C����ɳ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2955
            TabIndex        =   9
            Top             =   255
            Width           =   1320
         End
         Begin VB.OptionButton Option1 
            Caption         =   "D���ĳ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   4275
            TabIndex        =   8
            Top             =   255
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "1����ļ�ͥ���ڵأ�(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   3000
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   960
         Index           =   2
         Left            =   465
         TabIndex        =   18
         Top             =   1725
         Width           =   8040
         Begin VB.OptionButton Option2 
            Caption         =   "D�������񹤻�ũ��"
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
            Index           =   4
            Left            =   2145
            TabIndex        =   24
            Top             =   705
            Width           =   2100
         End
         Begin VB.OptionButton Option2 
            Caption         =   "B������Ա ���� ҽ�� ��ʦ ��ʦ��"
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
            Index           =   2
            Left            =   4470
            TabIndex        =   22
            Top             =   420
            Width           =   3570
         End
         Begin VB.OptionButton Option2 
            Caption         =   "A�����ء�����ҵ��λ�쵼�����ɲ�"
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
            Index           =   1
            Left            =   420
            TabIndex        =   21
            Top             =   420
            Width           =   3780
         End
         Begin VB.OptionButton Option2 
            Caption         =   "E����ҵ������"
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
            Index           =   5
            Left            =   4485
            TabIndex        =   20
            Top             =   720
            Width           =   1710
         End
         Begin VB.OptionButton Option2 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   420
            TabIndex        =   23
            Top             =   720
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "2���㸸�׵�ְҵ��(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   150
            Width           =   3615
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   930
         Index           =   3
         Left            =   375
         TabIndex        =   25
         Top             =   2715
         Width           =   8280
         Begin VB.OptionButton Option3 
            Caption         =   "E����ҵ������"
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
            Index           =   5
            Left            =   4605
            TabIndex        =   30
            Top             =   675
            Width           =   1740
         End
         Begin VB.OptionButton Option3 
            Caption         =   "A�����ء�����ҵ��λ�쵼�����ɲ�"
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
            Index           =   1
            Left            =   510
            TabIndex        =   29
            Top             =   405
            Width           =   3780
         End
         Begin VB.OptionButton Option3 
            Caption         =   "B������Ա ���� ҽ�� ��ʦ ��ʦ��"
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
            Index           =   2
            Left            =   4590
            TabIndex        =   28
            Top             =   390
            Width           =   3735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "D�������񹤻�ũ��"
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
            Index           =   4
            Left            =   2115
            TabIndex        =   26
            Top             =   675
            Width           =   2100
         End
         Begin VB.OptionButton Option3 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   525
            TabIndex        =   27
            Top             =   675
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "3����ĸ�׵�ְҵ��(   )"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   195
            TabIndex        =   31
            Top             =   150
            Width           =   3615
         End
      End
      Begin VB.Label Label5 
         Caption         =   "��һ���֣�ѧ�����˼���ͥ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2580
         TabIndex        =   44
         Top             =   375
         Width           =   4260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ͥ�������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8985
      Index           =   2
      Left            =   12075
      TabIndex        =   66
      Top             =   7860
      Width           =   10005
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   16
         Left            =   510
         TabIndex        =   110
         Top             =   7155
         Width           =   9465
         Begin VB.OptionButton Option16 
            Caption         =   "C����ʼ�����ܣ����´������Ľ������𽥽���"
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
            Index           =   3
            Left            =   405
            TabIndex        =   114
            Top             =   585
            Width           =   4710
         End
         Begin VB.OptionButton Option16 
            Caption         =   "D����̫����"
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
            Index           =   4
            Left            =   5385
            TabIndex        =   113
            Top             =   570
            Width           =   1560
         End
         Begin VB.OptionButton Option16 
            Caption         =   "B����������"
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
            Index           =   2
            Left            =   2565
            TabIndex        =   112
            Top             =   315
            Width           =   1515
         End
         Begin VB.OptionButton Option16 
            Caption         =   "A�����Ľ���"
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
            Index           =   1
            Left            =   405
            TabIndex        =   111
            Top             =   315
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "8���ڶ�������£���Լҳ�����������������̬��?��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   21
            Left            =   105
            TabIndex        =   115
            Top             =   30
            Width           =   6705
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   15
         Left            =   495
         TabIndex        =   104
         Top             =   5967
         Width           =   9465
         Begin VB.OptionButton Option15 
            Caption         =   "C��̸�Ľ��������̶Բ߲�����"
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
            Index           =   3
            Left            =   5295
            TabIndex        =   108
            Top             =   345
            Width           =   3195
         End
         Begin VB.OptionButton Option15 
            Caption         =   "D������ǿ�����Գɼ�����Ҫ�ԣ����ܽ������´β�һ������ô����"
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
            Index           =   4
            Left            =   420
            TabIndex        =   107
            Top             =   645
            Width           =   6645
         End
         Begin VB.OptionButton Option15 
            Caption         =   "B���������ľ��鼮"
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
            Index           =   2
            Left            =   2955
            TabIndex        =   106
            Top             =   330
            Width           =   2130
         End
         Begin VB.OptionButton Option15 
            Caption         =   "A����������ŵ��ЩǮ"
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
            Index           =   1
            Left            =   420
            TabIndex        =   105
            Top             =   330
            Width           =   2355
         End
         Begin VB.Label Label1 
            Caption         =   "7��Ϊ�˹������ܼ���ȡ������ɼ���ͨ���㸸ĸ���ȡ���ִ�ʩ? ��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   20
            Left            =   105
            TabIndex        =   109
            Top             =   30
            Width           =   7860
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   14
         Left            =   495
         TabIndex        =   98
         Top             =   5100
         Width           =   9465
         Begin VB.OptionButton Option14 
            Caption         =   "C��ż���μ�"
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
            Index           =   3
            Left            =   4170
            TabIndex        =   102
            Top             =   330
            Width           =   1530
         End
         Begin VB.OptionButton Option14 
            Caption         =   " D�����ٲμ�"
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
            Index           =   4
            Left            =   6000
            TabIndex        =   101
            Top             =   345
            Width           =   1560
         End
         Begin VB.OptionButton Option14 
            Caption         =   "B���󲿷ֲμ�"
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
            Index           =   2
            Left            =   2130
            TabIndex        =   100
            Top             =   345
            Width           =   1695
         End
         Begin VB.OptionButton Option14 
            Caption         =   "A��ÿ�βμ�"
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
            Index           =   1
            Left            =   405
            TabIndex        =   99
            Top             =   345
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "6��ѧУ�ٿ��ļҳ��ᣬ��ĸ�ĸ��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   19
            Left            =   120
            TabIndex        =   103
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   13
         Left            =   495
         TabIndex        =   92
         Top             =   4233
         Width           =   9465
         Begin VB.OptionButton Option13 
            Caption         =   "C��̸�Ľ���"
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
            Index           =   3
            Left            =   4020
            TabIndex        =   96
            Top             =   360
            Width           =   1500
         End
         Begin VB.OptionButton Option13 
            Caption         =   "D��������ʦ��ͬ����"
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
            Index           =   4
            Left            =   5700
            TabIndex        =   95
            Top             =   360
            Width           =   2490
         End
         Begin VB.OptionButton Option13 
            Caption         =   "B����֮����"
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
            Index           =   2
            Left            =   2445
            TabIndex        =   94
            Top             =   360
            Width           =   1485
         End
         Begin VB.OptionButton Option13 
            Caption         =   "A�������"
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
            Index           =   1
            Left            =   420
            TabIndex        =   93
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "5������ĳɼ�������򷸴�ʱ����ĸ���ȡʲô�ֶν����㣿��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   18
            Left            =   90
            TabIndex        =   97
            Top             =   45
            Width           =   7680
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   12
         Left            =   495
         TabIndex        =   86
         Top             =   3366
         Width           =   9465
         Begin VB.OptionButton Option12 
            Caption         =   "C�������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   90
            Top             =   345
            Width           =   1740
         End
         Begin VB.OptionButton Option12 
            Caption         =   "D���Ӳ�"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   89
            Top             =   285
            Width           =   1560
         End
         Begin VB.OptionButton Option12 
            Caption         =   "B��ż��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   88
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option12 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   87
            Top             =   300
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "4����ĸ�ڼ����и��㱨�������𣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   90
            TabIndex        =   91
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   11
         Left            =   495
         TabIndex        =   80
         Top             =   2499
         Width           =   9465
         Begin VB.OptionButton Option11 
            Caption         =   "C���˶�����"
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
            Index           =   3
            Left            =   4125
            TabIndex        =   84
            Top             =   360
            Width           =   1485
         End
         Begin VB.OptionButton Option11 
            Caption         =   "D������ѧϰ"
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
            Index           =   4
            Left            =   5925
            TabIndex        =   83
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton Option11 
            Caption         =   "B���ȾƳ�K"
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
            Index           =   2
            Left            =   2265
            TabIndex        =   82
            Top             =   360
            Width           =   1380
         End
         Begin VB.OptionButton Option11 
            Caption         =   "A�����ƶĲ�"
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
            Index           =   1
            Left            =   420
            TabIndex        =   81
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "3���㸸ĸ��ҳ�����μӵ�ҵ����ʲô����   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   16
            Left            =   105
            TabIndex        =   85
            Top             =   45
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   10
         Left            =   495
         TabIndex        =   74
         Top             =   1662
         Width           =   9465
         Begin VB.OptionButton Option10 
            Caption         =   "C��������"
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
            Index           =   3
            Left            =   3660
            TabIndex        =   78
            Top             =   315
            Width           =   1350
         End
         Begin VB.OptionButton Option10 
            Caption         =   "D��������"
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
            Index           =   4
            Left            =   5580
            TabIndex        =   77
            Top             =   315
            Width           =   1560
         End
         Begin VB.OptionButton Option10 
            Caption         =   "B��żȻ��"
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
            Index           =   2
            Left            =   1890
            TabIndex        =   76
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option10 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   420
            TabIndex        =   75
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "2�����Ƿ���������   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   14
            Left            =   90
            TabIndex        =   79
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   9
         Left            =   540
         TabIndex        =   67
         Top             =   645
         Width           =   9465
         Begin VB.OptionButton Option9 
            Caption         =   "A��������̸"
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
            Index           =   1
            Left            =   420
            TabIndex        =   71
            Top             =   585
            Width           =   1500
         End
         Begin VB.OptionButton Option9 
            Caption         =   "B��ż����̸"
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
            Index           =   2
            Left            =   2190
            TabIndex        =   70
            Top             =   585
            Width           =   1515
         End
         Begin VB.OptionButton Option9 
            Caption         =   " D���Ӳ���̸"
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
            Index           =   4
            Left            =   6345
            TabIndex        =   69
            Top             =   585
            Width           =   1665
         End
         Begin VB.OptionButton Option9 
            Caption         =   "C������û�н�̸"
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
            Index           =   3
            Left            =   4095
            TabIndex        =   68
            Top             =   585
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "1����ĸ�ĸ��ҳ��������㽻̸ѧУ�������𣿣�   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   22
            Left            =   90
            TabIndex        =   72
            Top             =   255
            Width           =   6225
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   2
         Left            =   6915
         TabIndex        =   429
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   2
         Left            =   5100
         TabIndex        =   437
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label Label3 
         Caption         =   "�ڶ����֣���ͥ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3105
         TabIndex        =   73
         Top             =   405
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ͥ�������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Index           =   3
      Left            =   11895
      TabIndex        =   3
      Top             =   7290
      Width           =   9915
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   23
         Left            =   510
         TabIndex        =   151
         Top             =   7155
         Width           =   7740
         Begin VB.OptionButton Option23 
            Caption         =   "C���������е��·���װ��"
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
            Index           =   3
            Left            =   435
            TabIndex        =   155
            Top             =   585
            Width           =   2805
         End
         Begin VB.OptionButton Option23 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   3675
            TabIndex        =   154
            Top             =   585
            Width           =   1560
         End
         Begin VB.OptionButton Option23 
            Caption         =   "B����������Ϸ���ưɵ�"
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
            Index           =   2
            Left            =   3675
            TabIndex        =   153
            Top             =   315
            Width           =   2775
         End
         Begin VB.OptionButton Option23 
            Caption         =   "A�����鼮�ľߵ�"
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
            Index           =   1
            Left            =   450
            TabIndex        =   152
            Top             =   315
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "15���������Ǯ��Ҫ�����ķ��棨     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   12
            Left            =   90
            TabIndex        =   156
            Top             =   45
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   20
         Left            =   480
         TabIndex        =   145
         Top             =   4221
         Width           =   9465
         Begin VB.OptionButton Option20 
            Caption         =   "E������"
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
            Index           =   5
            Left            =   6150
            TabIndex        =   202
            Top             =   540
            Width           =   1050
         End
         Begin VB.OptionButton Option20 
            Caption         =   "C��ͬѧ"
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
            Index           =   3
            Left            =   3540
            TabIndex        =   149
            Top             =   540
            Width           =   1065
         End
         Begin VB.OptionButton Option20 
            Caption         =   "D����ʦ"
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
            Index           =   4
            Left            =   4860
            TabIndex        =   148
            Top             =   525
            Width           =   1050
         End
         Begin VB.OptionButton Option20 
            Caption         =   "B����������"
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
            Index           =   2
            Left            =   1890
            TabIndex        =   147
            Top             =   540
            Width           =   1470
         End
         Begin VB.OptionButton Option20 
            Caption         =   "A����ĸ��"
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
            Index           =   1
            Left            =   465
            TabIndex        =   146
            Top             =   540
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "12������������˳�ĵ���ʱ���������뵽�����߶����ǣ�      ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   90
            TabIndex        =   150
            Top             =   255
            Width           =   6735
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   17
         Left            =   480
         TabIndex        =   140
         Top             =   615
         Width           =   9465
         Begin VB.OptionButton Option17 
            Caption         =   "B��������г�ͣ����ɷ��������˭˵���е�����˭��"
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
            Index           =   2
            Left            =   435
            TabIndex        =   201
            Top             =   540
            Width           =   5340
         End
         Begin VB.OptionButton Option17 
            Caption         =   $"frmInputMainMS.frx":000D
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
            Index           =   3
            Left            =   435
            TabIndex        =   143
            Top             =   825
            Width           =   5130
         End
         Begin VB.OptionButton Option17 
            Caption         =   $"frmInputMainMS.frx":003C
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
            Index           =   4
            Left            =   435
            TabIndex        =   142
            Top             =   1095
            Width           =   5475
         End
         Begin VB.OptionButton Option17 
            Caption         =   "A��ר��˳���ͣ���ĸ��˵���㣬��ֻ�ܷ���"
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
            Index           =   1
            Left            =   435
            TabIndex        =   141
            Top             =   285
            Width           =   4545
         End
         Begin VB.Label Label1 
            Caption         =   "9����ѡ���㸸ĸ�ڴ�������Ĺ�ϵʱ������������������?��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   10
            Left            =   105
            TabIndex        =   144
            Top             =   45
            Width           =   6915
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   18
         Left            =   480
         TabIndex        =   134
         Top             =   2222
         Width           =   9465
         Begin VB.OptionButton Option18 
            Caption         =   "C������һ��"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   138
            Top             =   405
            Width           =   1740
         End
         Begin VB.OptionButton Option18 
            Caption         =   "D���Ӳ�һ��"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   137
            Top             =   405
            Width           =   1560
         End
         Begin VB.OptionButton Option18 
            Caption         =   "B������һ��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   136
            Top             =   405
            Width           =   1590
         End
         Begin VB.OptionButton Option18 
            Caption         =   "A��һ��"
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
            Index           =   1
            Left            =   435
            TabIndex        =   135
            Top             =   405
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "10����ĸ�ĸ�ڹ�ͬ�������ʱ���Ƿ�ע�Ᵽ�����һ��? ��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   139
            Top             =   135
            Width           =   7650
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   885
         Index           =   19
         Left            =   480
         TabIndex        =   128
         Top             =   3184
         Width           =   9465
         Begin VB.OptionButton Option19 
            Caption         =   "C����������ʦ��Ϊ�㱧��ƽ"
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
            Index           =   3
            Left            =   420
            TabIndex        =   132
            Top             =   570
            Width           =   2955
         End
         Begin VB.OptionButton Option19 
            Caption         =   "D��������֮"
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
            Index           =   4
            Left            =   3975
            TabIndex        =   131
            Top             =   600
            Width           =   1995
         End
         Begin VB.OptionButton Option19 
            Caption         =   "B�������˽���������������˾�����"
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
            Index           =   2
            Left            =   3960
            TabIndex        =   130
            Top             =   345
            Width           =   4050
         End
         Begin VB.OptionButton Option19 
            Caption         =   "A���������߶�ʮһ������һ��"
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
            Index           =   1
            Left            =   420
            TabIndex        =   129
            Top             =   300
            Width           =   3225
         End
         Begin VB.Label Label1 
            Caption         =   "11������ʦ��������""�����״""���㸸ĸ�ᣨ     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   133
            Top             =   45
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   21
         Left            =   495
         TabIndex        =   122
         Top             =   5183
         Width           =   9465
         Begin VB.OptionButton Option21 
            Caption         =   "C������Ϥ"
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
            Index           =   3
            Left            =   3945
            TabIndex        =   126
            Top             =   420
            Width           =   1350
         End
         Begin VB.OptionButton Option21 
            Caption         =   "D���Ӳ�����"
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
            Index           =   4
            Left            =   5790
            TabIndex        =   125
            Top             =   420
            Width           =   1560
         End
         Begin VB.OptionButton Option21 
            Caption         =   "B���Ƚ���Ϥ"
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
            Index           =   2
            Left            =   1905
            TabIndex        =   124
            Top             =   405
            Width           =   1605
         End
         Begin VB.OptionButton Option21 
            Caption         =   "A����Ϥ"
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
            Index           =   1
            Left            =   465
            TabIndex        =   123
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "13����ĸ���Ƿ���Ϥ�㾭�����������ѣ���      ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   90
            TabIndex        =   127
            Top             =   105
            Width           =   5580
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   22
         Left            =   480
         TabIndex        =   116
         Top             =   6055
         Width           =   9465
         Begin VB.OptionButton Option22 
            Caption         =   "C����ʳ��������������õ���"
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
            Index           =   3
            Left            =   465
            TabIndex        =   120
            Top             =   675
            Width           =   3195
         End
         Begin VB.OptionButton Option22 
            Caption         =   "D��Ʒ�¸��е���"
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
            Index           =   4
            Left            =   4050
            TabIndex        =   119
            Top             =   675
            Width           =   2310
         End
         Begin VB.OptionButton Option22 
            Caption         =   "B������ͷ�ع���ҫ�����"
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
            Index           =   2
            Left            =   4065
            TabIndex        =   118
            Top             =   375
            Width           =   2790
         End
         Begin VB.OptionButton Option22 
            Caption         =   "A����׬Ǯ�����ܵ���"
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
            Index           =   1
            Left            =   465
            TabIndex        =   117
            Top             =   375
            Width           =   2790
         End
         Begin VB.Label Label1 
            Caption         =   "14����ĸ�ĸ��ϣ�����Ϊ�������ˣ���     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   121
            Top             =   90
            Width           =   6225
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   3
         Left            =   6915
         TabIndex        =   430
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   3
         Left            =   5100
         TabIndex        =   438
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧУ�������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9525
      Index           =   4
      Left            =   11490
      TabIndex        =   4
      Top             =   6720
      Width           =   10320
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   30
         Left            =   465
         TabIndex        =   193
         Top             =   7215
         Width           =   9465
         Begin VB.OptionButton Option30 
            Caption         =   "A���������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   197
            Top             =   525
            Width           =   1470
         End
         Begin VB.OptionButton Option30 
            Caption         =   "B���������Ҫ�õľͰ�"
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
            Index           =   2
            Left            =   2235
            TabIndex        =   196
            Top             =   525
            Width           =   2550
         End
         Begin VB.OptionButton Option30 
            Caption         =   "D���ظ���"
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
            Index           =   4
            Left            =   6660
            TabIndex        =   195
            Top             =   525
            Width           =   1260
         End
         Begin VB.OptionButton Option30 
            Caption         =   "C�������޹�"
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
            Index           =   3
            Left            =   4980
            TabIndex        =   194
            Top             =   525
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "7������ͬѧ�ܵ���ʦ����ʱ������ô�룿��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   27
            Left            =   90
            TabIndex        =   198
            Top             =   255
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   25
         Left            =   450
         TabIndex        =   187
         Top             =   1737
         Width           =   9465
         Begin VB.OptionButton Option25 
            Caption         =   "A��1Сʱ����"
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
            Index           =   1
            Left            =   435
            TabIndex        =   191
            Top             =   375
            Width           =   1590
         End
         Begin VB.OptionButton Option25 
            Caption         =   "B����Լ1--2Сʱ"
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
            Index           =   2
            Left            =   2190
            TabIndex        =   190
            Top             =   375
            Width           =   1935
         End
         Begin VB.OptionButton Option25 
            Caption         =   $"frmInputMainMS.frx":006D
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
            Index           =   4
            Left            =   6300
            TabIndex        =   189
            Top             =   375
            Width           =   1575
         End
         Begin VB.OptionButton Option25 
            Caption         =   "C����Լ2--3Сʱ"
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
            Index           =   3
            Left            =   4245
            TabIndex        =   188
            Top             =   375
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "2��Ŀǰ����ѧ����ÿ�������ʦ���õ���ҵƽ��Ҫ����ʱ��Ϊ��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   26
            Left            =   90
            TabIndex        =   192
            Top             =   75
            Width           =   7365
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   26
         Left            =   450
         TabIndex        =   181
         Top             =   2724
         Width           =   9465
         Begin VB.OptionButton Option26 
            Caption         =   "A���ȶ���˼����ȷʵ��������̱���"
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
            Index           =   1
            Left            =   450
            TabIndex        =   185
            Top             =   300
            Width           =   3795
         End
         Begin VB.OptionButton Option26 
            Caption         =   "B��ֱ���ʻ�����ͬѧ����������ô��"
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
            Index           =   2
            Left            =   4590
            TabIndex        =   184
            Top             =   300
            Width           =   3840
         End
         Begin VB.OptionButton Option26 
            Caption         =   "D��ʲô������"
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
            Index           =   4
            Left            =   2655
            TabIndex        =   183
            Top             =   570
            Width           =   1935
         End
         Begin VB.OptionButton Option26 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   450
            TabIndex        =   182
            Top             =   570
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "3��������ҵ�����������ѡ������������������ "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   25
            Left            =   90
            TabIndex        =   186
            Top             =   0
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   27
         Left            =   450
         TabIndex        =   175
         Top             =   3711
         Width           =   9465
         Begin VB.OptionButton Option27 
            Caption         =   "A����Լ0--0.5Сʱ"
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
            Index           =   1
            Left            =   435
            TabIndex        =   179
            Top             =   360
            Width           =   2085
         End
         Begin VB.OptionButton Option27 
            Caption         =   "B����Լ0.5--1Сʱ"
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
            Index           =   2
            Left            =   2685
            TabIndex        =   178
            Top             =   360
            Width           =   2100
         End
         Begin VB.OptionButton Option27 
            Caption         =   $"frmInputMainMS.frx":007F
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
            Index           =   4
            Left            =   7110
            TabIndex        =   177
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton Option27 
            Caption         =   "C����Լ1--2Сʱ"
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
            Index           =   3
            Left            =   4935
            TabIndex        =   176
            Top             =   360
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "4����ÿ����ѧУ���������������������������κ����ٵȣ����ʱ���ԼΪ���٣���   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   24
            Left            =   105
            TabIndex        =   180
            Top             =   60
            Width           =   9315
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   28
         Left            =   450
         TabIndex        =   169
         Top             =   4863
         Width           =   9465
         Begin VB.OptionButton Option28 
            Caption         =   "A������û���ӣ��Ժ�����ʦ���Ÿ�"
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
            Index           =   1
            Left            =   450
            TabIndex        =   173
            Top             =   270
            Width           =   3660
         End
         Begin VB.OptionButton Option28 
            Caption         =   "B�����ܽ���"
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
            Index           =   2
            Left            =   4770
            TabIndex        =   172
            Top             =   270
            Width           =   1590
         End
         Begin VB.OptionButton Option28 
            Caption         =   $"frmInputMainMS.frx":0091
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
            Index           =   4
            Left            =   4755
            TabIndex        =   171
            Top             =   510
            Width           =   1290
         End
         Begin VB.OptionButton Option28 
            Caption         =   "C���������ܵ��˺����Ӵ�һ�겻��"
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
            Index           =   3
            Left            =   450
            TabIndex        =   170
            Top             =   510
            Width           =   6465
         End
         Begin VB.Label Label1 
            Caption         =   "5���㱻��ʦ�����������̬���ǣ�   �� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   23
            Left            =   120
            TabIndex        =   174
            Top             =   15
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   1185
         Index           =   29
         Left            =   450
         TabIndex        =   163
         Top             =   5835
         Width           =   9465
         Begin VB.OptionButton Option29 
            Caption         =   "L������"
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
            Index           =   12
            Left            =   7560
            TabIndex        =   210
            Top             =   840
            Width           =   1080
         End
         Begin VB.OptionButton Option29 
            Caption         =   "K����������ѧ������"
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
            Index           =   11
            Left            =   5115
            TabIndex        =   209
            Top             =   840
            Width           =   2430
         End
         Begin VB.OptionButton Option29 
            Caption         =   "J��֪ʶǷԨ��"
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
            Index           =   10
            Left            =   3195
            TabIndex        =   208
            Top             =   825
            Width           =   1665
         End
         Begin VB.OptionButton Option29 
            Caption         =   "I�������ڳ��ϴ������ʦ"
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
            Index           =   9
            Left            =   420
            TabIndex        =   207
            Top             =   840
            Width           =   2775
         End
         Begin VB.OptionButton Option29 
            Caption         =   "H����ѧˮƽ��"
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
            Index           =   8
            Left            =   7590
            TabIndex        =   206
            Top             =   555
            Width           =   1680
         End
         Begin VB.OptionButton Option29 
            Caption         =   "G����������"
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
            Index           =   7
            Left            =   5940
            TabIndex        =   205
            Top             =   555
            Width           =   1470
         End
         Begin VB.OptionButton Option29 
            Caption         =   "F���������󲻺�"
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
            Index           =   6
            Left            =   3960
            TabIndex        =   204
            Top             =   570
            Width           =   1875
         End
         Begin VB.OptionButton Option29 
            Caption         =   "E��������ͬѧ"
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
            Index           =   5
            Left            =   2235
            TabIndex        =   203
            Top             =   570
            Width           =   1710
         End
         Begin VB.OptionButton Option29 
            Caption         =   "A����ѧ������һ��ͬ��"
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
            Index           =   1
            Left            =   390
            TabIndex        =   167
            Top             =   285
            Width           =   2595
         End
         Begin VB.OptionButton Option29 
            Caption         =   "B�����������Դ���̣���������"
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
            Index           =   2
            Left            =   3045
            TabIndex        =   166
            Top             =   285
            Width           =   3450
         End
         Begin VB.OptionButton Option29 
            Caption         =   "D��û����Ĭ��"
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
            Index           =   4
            Left            =   405
            TabIndex        =   165
            Top             =   570
            Width           =   1725
         End
         Begin VB.OptionButton Option29 
            Caption         =   "C������������"
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
            Index           =   3
            Left            =   6705
            TabIndex        =   164
            Top             =   285
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "6�����ϲ��ʲô���͵���ʦ����    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   15
            Left            =   90
            TabIndex        =   168
            Top             =   15
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   24
         Left            =   450
         TabIndex        =   157
         Top             =   885
         Width           =   9465
         Begin VB.OptionButton Option24 
            Caption         =   "A��ϲ��"
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
            Index           =   1
            Left            =   435
            TabIndex        =   161
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option24 
            Caption         =   "B���Ƚ�ϲ��"
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
            Index           =   2
            Left            =   1710
            TabIndex        =   160
            Top             =   375
            Width           =   1530
         End
         Begin VB.OptionButton Option24 
            Caption         =   "D����ϲ��"
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
            Index           =   4
            Left            =   5340
            TabIndex        =   159
            Top             =   375
            Width           =   1560
         End
         Begin VB.OptionButton Option24 
            Caption         =   "C��˵����"
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
            Index           =   3
            Left            =   3420
            TabIndex        =   158
            Top             =   375
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "1����ϲ�������ڵ�ѧУ�𣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   90
            TabIndex        =   162
            Top             =   90
            Width           =   6225
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   4
         Left            =   6915
         TabIndex        =   431
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   4
         Left            =   5100
         TabIndex        =   439
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label Label2 
         Caption         =   "�������֣�ѧУ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   199
         Top             =   420
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧУ�������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8955
      Index           =   5
      Left            =   11145
      TabIndex        =   5
      Top             =   6015
      Width           =   9975
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   36
         Left            =   495
         TabIndex        =   446
         Top             =   5475
         Width           =   9465
         Begin VB.OptionButton Option36 
            Caption         =   "A������̸�ģ������˼ʹ�ϵָ����"
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
            Index           =   1
            Left            =   525
            TabIndex        =   450
            Top             =   390
            Width           =   3735
         End
         Begin VB.OptionButton Option36 
            Caption         =   "B������ĳɼ���ż������һ��������뷨"
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
            Index           =   2
            Left            =   4395
            TabIndex        =   449
            Top             =   345
            Width           =   4335
         End
         Begin VB.OptionButton Option36 
            Caption         =   "D��������������"
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
            Index           =   4
            Left            =   4305
            TabIndex        =   448
            Top             =   645
            Width           =   1995
         End
         Begin VB.OptionButton Option36 
            Caption         =   "C��ֻ���ĳɼ�����������"
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
            Index           =   3
            Left            =   540
            TabIndex        =   447
            Top             =   675
            Width           =   2745
         End
         Begin VB.Label Label1 
            Caption         =   "13����������ʦ����Ĺ��ĳ̶��ǣ�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   451
            Top             =   90
            Width           =   7590
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   31
         Left            =   555
         TabIndex        =   245
         Top             =   645
         Width           =   9465
         Begin VB.OptionButton Option31 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   249
            Top             =   300
            Width           =   1740
         End
         Begin VB.OptionButton Option31 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   248
            Top             =   300
            Width           =   1560
         End
         Begin VB.OptionButton Option31 
            Caption         =   "B�����ձ�"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   247
            Top             =   300
            Width           =   1380
         End
         Begin VB.OptionButton Option31 
            Caption         =   "A��û��"
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
            Index           =   1
            Left            =   375
            TabIndex        =   246
            Top             =   300
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "8��ѧУ�Ƿ�����ʦ��ͬѧ�巣������巣���緣վ�������顢��д�ֵȣ����󣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   13
            Left            =   30
            TabIndex        =   250
            Top             =   45
            Width           =   9210
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   33
         Left            =   195
         TabIndex        =   240
         Top             =   2550
         Width           =   9465
         Begin VB.OptionButton Option33 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3735
            TabIndex        =   243
            Top             =   390
            Width           =   1080
         End
         Begin VB.OptionButton Option33 
            Caption         =   "B�����ձ�"
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
            Index           =   2
            Left            =   2040
            TabIndex        =   242
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton Option33 
            Caption         =   "A��û��"
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
            Index           =   1
            Left            =   735
            TabIndex        =   241
            Top             =   405
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "10����ʦƽʱ�Ͽγٵ����ڿ�����ʹ���ֻ�������ǣ�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   34
            Left            =   375
            TabIndex        =   244
            Top             =   120
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   34
         Left            =   495
         TabIndex        =   235
         Top             =   3495
         Width           =   9465
         Begin VB.OptionButton Option34 
            Caption         =   "C����ϲ��"
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
            Index           =   3
            Left            =   3405
            TabIndex        =   238
            Top             =   375
            Width           =   1740
         End
         Begin VB.OptionButton Option34 
            Caption         =   "B��������"
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
            Index           =   2
            Left            =   1710
            TabIndex        =   237
            Top             =   375
            Width           =   1380
         End
         Begin VB.OptionButton Option34 
            Caption         =   "A��ϲ��"
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
            Index           =   1
            Left            =   435
            TabIndex        =   236
            Top             =   375
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "11����ϲ���������飨��ֽ����־���𣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   33
            Left            =   75
            TabIndex        =   239
            Top             =   90
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   960
         Index           =   35
         Left            =   495
         TabIndex        =   229
         Top             =   4350
         Width           =   9465
         Begin VB.OptionButton Option35 
            Caption         =   "C��һ��"
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
            Index           =   3
            Left            =   4155
            TabIndex        =   233
            Top             =   360
            Width           =   1110
         End
         Begin VB.OptionButton Option35 
            Caption         =   "D�����Կ��Գɼ��ߵ�����ѧ��"
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
            Index           =   4
            Left            =   465
            TabIndex        =   232
            Top             =   675
            Width           =   3270
         End
         Begin VB.OptionButton Option35 
            Caption         =   "B���Ƚ�����"
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
            Index           =   2
            Left            =   1935
            TabIndex        =   231
            Top             =   405
            Width           =   1605
         End
         Begin VB.OptionButton Option35 
            Caption         =   "A��������"
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
            Index           =   1
            Left            =   465
            TabIndex        =   230
            Top             =   405
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "12����İ�������ʦ�Կ��Գɼ��ĸߵ�������ѧ��������ǣ�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   32
            Left            =   90
            TabIndex        =   234
            Top             =   90
            Width           =   7125
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   37
         Left            =   480
         TabIndex        =   223
         Top             =   6555
         Width           =   9465
         Begin VB.OptionButton Option37 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3555
            TabIndex        =   227
            Top             =   375
            Width           =   1740
         End
         Begin VB.OptionButton Option37 
            Caption         =   "D��������"
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
            Index           =   4
            Left            =   5475
            TabIndex        =   226
            Top             =   390
            Width           =   1350
         End
         Begin VB.OptionButton Option37 
            Caption         =   "B��һ��"
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
            Index           =   2
            Left            =   1860
            TabIndex        =   225
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton Option37 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   510
            TabIndex        =   224
            Top             =   390
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "14����������ʦ�Ƿ����ð��ν���˼�����κ���������Ƚ�����    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   31
            Left            =   90
            TabIndex        =   228
            Top             =   90
            Width           =   7590
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   38
         Left            =   480
         TabIndex        =   217
         Top             =   7440
         Width           =   8805
         Begin VB.OptionButton Option38 
            Caption         =   "C.�е�����"
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
            Index           =   3
            Left            =   3540
            TabIndex        =   221
            Top             =   360
            Width           =   1740
         End
         Begin VB.OptionButton Option38 
            Caption         =   "D������û��"
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
            Index           =   4
            Left            =   5475
            TabIndex        =   220
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton Option38 
            Caption         =   "B��ż��"
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
            Index           =   2
            Left            =   1845
            TabIndex        =   219
            Top             =   360
            Width           =   1380
         End
         Begin VB.OptionButton Option38 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   510
            TabIndex        =   218
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "15������ѧУ���ܹ��Խ����𣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   30
            Left            =   90
            TabIndex        =   222
            Top             =   90
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   32
         Left            =   510
         TabIndex        =   211
         Top             =   1575
         Width           =   9465
         Begin VB.OptionButton Option32 
            Caption         =   "C���Լ�Ҳ������"
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
            Index           =   3
            Left            =   420
            TabIndex        =   215
            Top             =   675
            Width           =   1905
         End
         Begin VB.OptionButton Option32 
            Caption         =   $"frmInputMainMS.frx":00A0
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
            Index           =   4
            Left            =   2775
            TabIndex        =   214
            Top             =   645
            Width           =   3555
         End
         Begin VB.OptionButton Option32 
            Caption         =   "B���Լ����������������ܱ�����ô��"
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
            Index           =   2
            Left            =   2790
            TabIndex        =   213
            Top             =   360
            Width           =   3795
         End
         Begin VB.OptionButton Option32 
            Caption         =   "A����ֹ"
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
            Index           =   1
            Left            =   420
            TabIndex        =   212
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "9���������Υ�����󣬱����ӷ�ֽ�����ϴ�����Ϳ�һ�������ʱ�����ǣ�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   28
            Left            =   90
            TabIndex        =   216
            Top             =   90
            Width           =   8085
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   5
         Left            =   6915
         TabIndex        =   432
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   5
         Left            =   5100
         TabIndex        =   440
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8865
      Index           =   6
      Left            =   10785
      TabIndex        =   6
      Top             =   5385
      Width           =   11310
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   39
         Left            =   495
         TabIndex        =   285
         Top             =   855
         Width           =   9465
         Begin VB.OptionButton Option39 
            Caption         =   "A��û��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   345
            TabIndex        =   289
            Top             =   330
            Width           =   1425
         End
         Begin VB.OptionButton Option39 
            Caption         =   "B��1��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2115
            TabIndex        =   288
            Top             =   330
            Width           =   1395
         End
         Begin VB.OptionButton Option39 
            Caption         =   "D��3�һ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   5865
            TabIndex        =   287
            Top             =   330
            Width           =   1575
         End
         Begin VB.OptionButton Option39 
            Caption         =   "C��2��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3990
            TabIndex        =   286
            Top             =   330
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "1����һ�ѧУ�����ж��ټ���Сѧ��������˵ĺ����ɣ���    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   40
            Left            =   15
            TabIndex        =   290
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   40
         Left            =   495
         TabIndex        =   280
         Top             =   1897
         Width           =   9465
         Begin VB.OptionButton Option40 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   315
            TabIndex        =   283
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton Option40 
            Caption         =   "B��ż��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   282
            Top             =   360
            Width           =   1380
         End
         Begin VB.OptionButton Option40 
            Caption         =   "C����ûȥ��"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   281
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "2�����Ƿ��Ӫҵ�����ɣ���    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   39
            Left            =   30
            TabIndex        =   284
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   41
         Left            =   495
         TabIndex        =   275
         Top             =   2925
         Width           =   10575
         Begin VB.OptionButton Option41 
            Caption         =   "A���ܶ�"
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
            Index           =   1
            Left            =   315
            TabIndex        =   278
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton Option41 
            Caption         =   "B������"
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
            Index           =   2
            Left            =   2130
            TabIndex        =   277
            Top             =   360
            Width           =   1380
         End
         Begin VB.OptionButton Option41 
            Caption         =   "C��û��"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   276
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "3��������ס�����������ѧУ���������̵����ʱ̯�������������Ķ�����̻�Σ�յ���ߣ���    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   38
            Left            =   30
            TabIndex        =   279
            Top             =   45
            Width           =   10500
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   42
         Left            =   495
         TabIndex        =   269
         Top             =   3945
         Width           =   9465
         Begin VB.OptionButton Option42 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   315
            TabIndex        =   273
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option42 
            Caption         =   "B��ż��"
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
            Index           =   2
            Left            =   1740
            TabIndex        =   272
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option42 
            Caption         =   "D������û��"
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
            Index           =   4
            Left            =   5025
            TabIndex        =   271
            Top             =   285
            Width           =   1560
         End
         Begin VB.OptionButton Option42 
            Caption         =   "C�������Թ�"
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
            Index           =   3
            Left            =   3090
            TabIndex        =   270
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "4���������𣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   37
            Left            =   30
            TabIndex        =   274
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   43
         Left            =   495
         TabIndex        =   264
         Top             =   4980
         Width           =   9465
         Begin VB.OptionButton Option43 
            Caption         =   "A����������"
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
            Index           =   1
            Left            =   345
            TabIndex        =   267
            Top             =   375
            Width           =   1545
         End
         Begin VB.OptionButton Option43 
            Caption         =   "B��ż��������"
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
            Index           =   2
            Left            =   2340
            TabIndex        =   266
            Top             =   375
            Width           =   1680
         End
         Begin VB.OptionButton Option43 
            Caption         =   "C��û�п�����"
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
            Index           =   3
            Left            =   4635
            TabIndex        =   265
            Top             =   375
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "5�������Լ���ס�������Ƿ�Ŀ�ù����������ڶĲ�����Ź�Ȳ�����Ϊ����    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   36
            Left            =   30
            TabIndex        =   268
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   44
         Left            =   495
         TabIndex        =   258
         Top             =   5955
         Width           =   9465
         Begin VB.OptionButton Option44 
            Caption         =   "A��Ӱ������Ů�������ǵľٶ�"
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
            Index           =   1
            Left            =   315
            TabIndex        =   262
            Top             =   315
            Width           =   3165
         End
         Begin VB.OptionButton Option44 
            Caption         =   "B����ɫ�鿯¼��ɫ��绰����վ�����ɫ��ĵ�����Ϸ"
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
            Index           =   2
            Left            =   3570
            TabIndex        =   261
            Top             =   315
            Width           =   5775
         End
         Begin VB.OptionButton Option44 
            Caption         =   "D��û�нӴ���"
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
            Index           =   4
            Left            =   3585
            TabIndex        =   260
            Top             =   615
            Width           =   1725
         End
         Begin VB.OptionButton Option44 
            Caption         =   "C���������"
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
            Index           =   3
            Left            =   300
            TabIndex        =   259
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "6����Ӵ���������Щ���󣿣�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   35
            Left            =   30
            TabIndex        =   263
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   45
         Left            =   495
         TabIndex        =   252
         Top             =   7230
         Width           =   9465
         Begin VB.OptionButton Option45 
            Caption         =   "A��ѧУ�͸�ĸ"
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
            Index           =   1
            Left            =   330
            TabIndex        =   256
            Top             =   345
            Width           =   1695
         End
         Begin VB.OptionButton Option45 
            Caption         =   "B�������鼮�����ӵ�Ӱ"
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
            Index           =   2
            Left            =   2295
            TabIndex        =   255
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton Option45 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   6375
            TabIndex        =   254
            Top             =   360
            Width           =   1080
         End
         Begin VB.OptionButton Option45 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   5085
            TabIndex        =   253
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "7����Ľ���֪ʶ��Ҫͨ��ʲô�����˽⣿��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   29
            Left            =   30
            TabIndex        =   257
            Top             =   45
            Width           =   8835
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   6
         Left            =   6915
         TabIndex        =   433
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   6
         Left            =   5100
         TabIndex        =   441
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label Label6 
         Caption         =   "���Ĳ��֣����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3780
         TabIndex        =   251
         Top             =   375
         Width           =   3975
      End
   End
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12000
      Begin VB.Label lblSchool 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   780
         TabIndex        =   444
         Top             =   75
         Width           =   9330
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   60
         Top             =   -15
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Index           =   7
      Left            =   10650
      TabIndex        =   63
      Top             =   4860
      Width           =   11370
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   47
         Left            =   480
         TabIndex        =   333
         Top             =   1155
         Width           =   10680
         Begin VB.OptionButton Option47 
            Caption         =   $"frmInputMainMS.frx":00C5
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
            Index           =   3
            Left            =   345
            TabIndex        =   337
            Top             =   600
            Width           =   2745
         End
         Begin VB.OptionButton Option47 
            Caption         =   $"frmInputMainMS.frx":00E6
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
            Index           =   4
            Left            =   3855
            TabIndex        =   336
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton Option47 
            Caption         =   "B����������г�һ�ձ���"
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
            Index           =   2
            Left            =   3840
            TabIndex        =   335
            Top             =   330
            Width           =   3060
         End
         Begin VB.OptionButton Option47 
            Caption         =   "A��Ͷ�����ǣ�Ѱ����Ԯ��"
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
            Index           =   1
            Left            =   345
            TabIndex        =   334
            Top             =   330
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "9��С�ɵĸ�ĸ����������������⣬�⵽ִ����Ա�Ĳ���ƽ�Դ����Դˣ�����ΪС�ɵĸ�ĸӦ�ã� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   48
            Left            =   60
            TabIndex        =   338
            Top             =   30
            Width           =   10350
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   885
         Index           =   48
         Left            =   480
         TabIndex        =   327
         Top             =   2190
         Width           =   9465
         Begin VB.OptionButton Option48 
            Caption         =   "C��������������������˲���ȫ����"
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
            Index           =   3
            Left            =   390
            TabIndex        =   331
            Top             =   615
            Width           =   3990
         End
         Begin VB.OptionButton Option48 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   4410
            TabIndex        =   330
            Top             =   645
            Width           =   1125
         End
         Begin VB.OptionButton Option48 
            Caption         =   "B�������˳����˿ڵľ�ҵѹ��"
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
            Index           =   2
            Left            =   4395
            TabIndex        =   329
            Top             =   360
            Width           =   3225
         End
         Begin VB.OptionButton Option48 
            Caption         =   "A������Ϊ���ҽ�����������Ҫ����"
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
            Index           =   1
            Left            =   390
            TabIndex        =   328
            Top             =   330
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "10��Ŀǰ��������������ԱԽ��Խ�࣬�Դˣ���Ŀ����ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   47
            Left            =   0
            TabIndex        =   332
            Top             =   30
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   49
         Left            =   480
         TabIndex        =   321
         Top             =   3225
         Width           =   9465
         Begin VB.OptionButton Option49 
            Caption         =   "C����̫��ͬ"
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
            Index           =   3
            Left            =   3945
            TabIndex        =   325
            Top             =   330
            Width           =   1560
         End
         Begin VB.OptionButton Option49 
            Caption         =   "D���ܲ���ͬ"
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
            Index           =   4
            Left            =   5880
            TabIndex        =   324
            Top             =   330
            Width           =   1560
         End
         Begin VB.OptionButton Option49 
            Caption         =   "B���Ƚ���ͬ"
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
            Index           =   2
            Left            =   2130
            TabIndex        =   323
            Top             =   330
            Width           =   1575
         End
         Begin VB.OptionButton Option49 
            Caption         =   "A���ǳ���ͬ"
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
            Index           =   1
            Left            =   360
            TabIndex        =   322
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "11������Ϊ����������Ϊ��Ⱦ���̲��Ƿ�ֵ��ͬ�飿��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   46
            Left            =   30
            TabIndex        =   326
            Top             =   30
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   50
         Left            =   480
         TabIndex        =   315
         Top             =   3945
         Width           =   9465
         Begin VB.OptionButton Option50 
            Caption         =   "C�����Ǵ��˵��£��Լ��ܲ���"
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
            Index           =   3
            Left            =   390
            TabIndex        =   319
            Top             =   600
            Width           =   3150
         End
         Begin VB.OptionButton Option50 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   3945
            TabIndex        =   318
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton Option50 
            Caption         =   "B��������������˰������Ӧ�ò���"
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
            Index           =   2
            Left            =   3915
            TabIndex        =   317
            Top             =   300
            Width           =   3810
         End
         Begin VB.OptionButton Option50 
            Caption         =   "A������˰�����˴�Ǯ"
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
            Index           =   1
            Left            =   390
            TabIndex        =   316
            Top             =   315
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "12�������������У�������˰������Ϊ����    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   45
            Left            =   30
            TabIndex        =   320
            Top             =   45
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   51
         Left            =   510
         TabIndex        =   309
         Top             =   4875
         Width           =   10815
         Begin VB.OptionButton Option51 
            Caption         =   "C����������ͬѧ֮������һ�ַ�ʽ"
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
            Index           =   3
            Left            =   345
            TabIndex        =   313
            Top             =   585
            Width           =   3795
         End
         Begin VB.OptionButton Option51 
            Caption         =   "D��������ͬѧ֮��Ļ����ʱȣ�����ʹ���ǹ��ֿ��ء�׷���Ǯ"
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
            Index           =   4
            Left            =   4275
            TabIndex        =   312
            Top             =   600
            Width           =   6405
         End
         Begin VB.OptionButton Option51 
            Caption         =   "B���Լ�Ҳ�������������ﾭ������������"
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
            Index           =   2
            Left            =   5205
            TabIndex        =   311
            Top             =   300
            Width           =   4185
         End
         Begin VB.OptionButton Option51 
            Caption         =   "A�����Ǹ��˵����飬ֻҪ����������������ν"
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
            Index           =   1
            Left            =   345
            TabIndex        =   310
            Top             =   285
            Width           =   4920
         End
         Begin VB.Label Label1 
            Caption         =   "13��������Щ��Сѧ��������ʱ�������ϯ������ͬѧ��������ף���Դ���Ŀ����ǣ���    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   44
            Left            =   30
            TabIndex        =   314
            Top             =   30
            Width           =   9870
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   1080
         Index           =   52
         Left            =   480
         TabIndex        =   303
         Top             =   5880
         Width           =   9465
         Begin VB.OptionButton Option52 
            Caption         =   "E������ʹ����������;����������ĵ�·"
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
            Index           =   5
            Left            =   435
            TabIndex        =   339
            Top             =   885
            Width           =   4395
         End
         Begin VB.OptionButton Option52 
            Caption         =   "C�������ú��ӶԸ�ĸ����Թ��"
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
            Index           =   3
            Left            =   450
            TabIndex        =   307
            Top             =   585
            Width           =   3465
         End
         Begin VB.OptionButton Option52 
            Caption         =   "D������Ӱ�쵽���ӶԼ�ͥ�ͻ����Ŀ���"
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
            Index           =   4
            Left            =   4335
            TabIndex        =   306
            Top             =   555
            Width           =   4200
         End
         Begin VB.OptionButton Option52 
            Caption         =   "B�������ú�����ͬѧ��̧����ͷ"
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
            Index           =   2
            Left            =   4170
            TabIndex        =   305
            Top             =   240
            Width           =   3660
         End
         Begin VB.OptionButton Option52 
            Caption         =   "A������ʹ���Ӹ�����"
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
            Index           =   1
            Left            =   450
            TabIndex        =   304
            Top             =   285
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "14������Ϊ��ĸ���Ժ��ӻ�����ЩӰ�죿��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   43
            Left            =   45
            TabIndex        =   308
            Top             =   0
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Caption         =   "15��С��Ϊ�˵��ϰ೤����һЩͬѧ������Է����Ա��ø����ѡƱ��������������������Ϊ����      ��"
         Height          =   975
         Index           =   53
         Left            =   495
         TabIndex        =   297
         Top             =   7155
         Width           =   10710
         Begin VB.OptionButton Option53 
            Caption         =   "C�������������ڰ����鶼�����������������ʲô��Ҳ�첻��"
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
            Index           =   3
            Left            =   450
            TabIndex        =   301
            Top             =   630
            Width           =   6090
         End
         Begin VB.OptionButton Option53 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   6675
            TabIndex        =   300
            Top             =   615
            Width           =   1560
         End
         Begin VB.OptionButton Option53 
            Caption         =   "B�����ã����������͵����ϰ೤"
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
            Index           =   2
            Left            =   4350
            TabIndex        =   299
            Top             =   330
            Width           =   3615
         End
         Begin VB.OptionButton Option53 
            Caption         =   "A�������Ȳ���ƽҲ������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   298
            Top             =   330
            Width           =   2745
         End
         Begin VB.Label Label1 
            Caption         =   "15��С��Ϊ�˵��ϰ೤����һЩͬѧ������Է����Ա��ø����ѡƱ��������������������Ϊ���� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   42
            Left            =   30
            TabIndex        =   302
            Top             =   30
            Width           =   10590
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   46
         Left            =   465
         TabIndex        =   291
         Top             =   480
         Width           =   9465
         Begin VB.OptionButton Option46 
            Caption         =   "C������ˮ׼�½�"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   295
            Top             =   300
            Width           =   1935
         End
         Begin VB.OptionButton Option46 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   6420
            TabIndex        =   294
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton Option46 
            Caption         =   "B��������"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   293
            Top             =   300
            Width           =   1515
         End
         Begin VB.OptionButton Option46 
            Caption         =   "A������ΰ�"
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
            Index           =   1
            Left            =   345
            TabIndex        =   292
            Top             =   300
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "8������Ϊ����������ע����������ǣ�    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   41
            Left            =   60
            TabIndex        =   296
            Top             =   15
            Width           =   8835
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   7
         Left            =   6915
         TabIndex        =   434
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   7
         Left            =   5100
         TabIndex        =   442
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "����Ʒ�������������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9045
      Index           =   8
      Left            =   10425
      TabIndex        =   64
      Top             =   4260
      Width           =   10575
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   1005
         Index           =   54
         Left            =   525
         TabIndex        =   373
         Top             =   900
         Width           =   9465
         Begin VB.OptionButton Option54 
            Caption         =   "I������"
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
            Index           =   8
            Left            =   6240
            TabIndex        =   427
            Top             =   690
            Width           =   1125
         End
         Begin VB.OptionButton Option54 
            Caption         =   "H��û������"
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
            Index           =   7
            Left            =   4350
            TabIndex        =   426
            Top             =   690
            Width           =   1590
         End
         Begin VB.OptionButton Option54 
            Caption         =   "F����˹����"
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
            Index           =   6
            Left            =   2400
            TabIndex        =   425
            Top             =   660
            Width           =   1590
         End
         Begin VB.OptionButton Option54 
            Caption         =   "E�����ѧ˵"
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
            Index           =   5
            Left            =   315
            TabIndex        =   424
            Top             =   660
            Width           =   1575
         End
         Begin VB.OptionButton Option54 
            Caption         =   "A����������"
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
            Index           =   1
            Left            =   315
            TabIndex        =   377
            Top             =   345
            Width           =   1635
         End
         Begin VB.OptionButton Option54 
            Caption         =   "B�����"
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
            Index           =   2
            Left            =   2385
            TabIndex        =   376
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option54 
            Caption         =   "D������ѧ˵"
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
            Index           =   4
            Left            =   6225
            TabIndex        =   375
            Top             =   300
            Width           =   1530
         End
         Begin VB.OptionButton Option54 
            Caption         =   "C��������"
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
            Index           =   3
            Left            =   4350
            TabIndex        =   374
            Top             =   330
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "1����������ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   55
            Left            =   0
            TabIndex        =   378
            Top             =   -15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   55
         Left            =   495
         TabIndex        =   367
         Top             =   2237
         Width           =   9465
         Begin VB.OptionButton Option55 
            Caption         =   "E������"
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
            Index           =   5
            Left            =   2850
            TabIndex        =   428
            Top             =   630
            Width           =   1080
         End
         Begin VB.OptionButton Option55 
            Caption         =   "A��������Ҫ"
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
            Index           =   1
            Left            =   345
            TabIndex        =   371
            Top             =   315
            Width           =   1545
         End
         Begin VB.OptionButton Option55 
            Caption         =   "B����ѧ��"
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
            Index           =   2
            Left            =   2850
            TabIndex        =   370
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option55 
            Caption         =   "D��Ӱ�ӻ���������"
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
            Index           =   4
            Left            =   345
            TabIndex        =   369
            Top             =   585
            Width           =   2235
         End
         Begin VB.OptionButton Option55 
            Caption         =   "C����ʷ�ܳ�����"
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
            Index           =   3
            Left            =   4695
            TabIndex        =   368
            Top             =   360
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "2���������ͣ����ݣ������ǣ�      ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   54
            Left            =   15
            TabIndex        =   372
            Top             =   30
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   56
         Left            =   480
         TabIndex        =   362
         Top             =   3420
         Width           =   8610
         Begin VB.OptionButton Option56 
            Caption         =   "A��Ϊ����������Լ�"
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
            Index           =   1
            Left            =   360
            TabIndex        =   365
            Top             =   330
            Width           =   2295
         End
         Begin VB.OptionButton Option56 
            Caption         =   "B���𵽴�ͷ����"
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
            Index           =   2
            Left            =   3165
            TabIndex        =   364
            Top             =   330
            Width           =   1920
         End
         Begin VB.OptionButton Option56 
            Caption         =   "C����ȡ�����ʱ�"
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
            Index           =   3
            Left            =   5355
            TabIndex        =   363
            Top             =   315
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "3�������ţ������Ķ����ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   53
            Left            =   45
            TabIndex        =   366
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   57
         Left            =   480
         TabIndex        =   357
         Top             =   4371
         Width           =   9465
         Begin VB.OptionButton Option57 
            Caption         =   "A����Ǯ�Ƕ�����������׵��˵�Ӧ�лر�"
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
            Index           =   1
            Left            =   390
            TabIndex        =   360
            Top             =   360
            Width           =   4215
         End
         Begin VB.OptionButton Option57 
            Caption         =   "B��û��Ǯ�������ܵ�"
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
            Index           =   2
            Left            =   5025
            TabIndex        =   359
            Top             =   345
            Width           =   2520
         End
         Begin VB.OptionButton Option57 
            Caption         =   "C��ǮԽ�࣬�˵ļ�ֵ��Խ��"
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
            Index           =   3
            Left            =   390
            TabIndex        =   358
            Top             =   645
            Width           =   3120
         End
         Begin VB.Label Label1 
            Caption         =   "4����Խ�Ǯ�Ŀ����ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   52
            Left            =   60
            TabIndex        =   361
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   58
         Left            =   480
         TabIndex        =   352
         Top             =   5588
         Width           =   9465
         Begin VB.OptionButton Option58 
            Caption         =   "A��֧��"
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
            Index           =   1
            Left            =   345
            TabIndex        =   355
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option58 
            Caption         =   "B������"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   354
            Top             =   255
            Width           =   1380
         End
         Begin VB.OptionButton Option58 
            Caption         =   "C������ν"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   353
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "5�����������̬���ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   51
            Left            =   60
            TabIndex        =   356
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   59
         Left            =   480
         TabIndex        =   346
         Top             =   6520
         Width           =   9465
         Begin VB.OptionButton Option59 
            Caption         =   "A��ѧУ"
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
            Index           =   1
            Left            =   375
            TabIndex        =   350
            Top             =   315
            Width           =   1410
         End
         Begin VB.OptionButton Option59 
            Caption         =   "B����ͥ"
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
            Index           =   2
            Left            =   2145
            TabIndex        =   349
            Top             =   300
            Width           =   1380
         End
         Begin VB.OptionButton Option59 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   348
            Top             =   270
            Width           =   1245
         End
         Begin VB.OptionButton Option59 
            Caption         =   "C�����"
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
            Index           =   3
            Left            =   4035
            TabIndex        =   347
            Top             =   300
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "6������ò����Ը����ɵ���Ҫ�����ǣ�     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   50
            Left            =   60
            TabIndex        =   351
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   60
         Left            =   450
         TabIndex        =   340
         Top             =   7455
         Width           =   9465
         Begin VB.OptionButton Option60 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   420
            TabIndex        =   344
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton Option60 
            Caption         =   "B��ƽ��"
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
            Index           =   2
            Left            =   2190
            TabIndex        =   343
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option60 
            Caption         =   "D������"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   342
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option60 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   4050
            TabIndex        =   341
            Top             =   270
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "7������Ϊ�Լ��������������������У���     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   49
            Left            =   75
            TabIndex        =   345
            Top             =   0
            Width           =   8835
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdNext 
         Height          =   405
         Index           =   8
         Left            =   6915
         TabIndex        =   435
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   8
         Left            =   5100
         TabIndex        =   443
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "�� һ ҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label Label7 
         Caption         =   "���岿�֣�����Ʒ�������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2895
         TabIndex        =   423
         Top             =   315
         Width           =   4845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����Ʒ�������������(��)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10680
      Index           =   9
      Left            =   225
      TabIndex        =   65
      Top             =   630
      Width           =   9780
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdExit 
         Height          =   405
         Left            =   3930
         TabIndex        =   452
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "������"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   33023
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdSave 
         Height          =   405
         Left            =   2790
         TabIndex        =   445
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "����"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   33023
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   61
         Left            =   510
         TabIndex        =   421
         Top             =   540
         Width           =   9465
         Begin VB.OptionButton Option61 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   345
            TabIndex        =   380
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option61 
            Caption         =   "B������"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   381
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option61 
            Caption         =   "D���Ӳ�"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   383
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton Option61 
            Caption         =   "C����ʱ"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   382
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "8�����Ƿ���Ϊ�Լ������ͺ���ױȱ����ѿ�����     �� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   63
            Left            =   60
            TabIndex        =   422
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   62
         Left            =   510
         TabIndex        =   419
         Top             =   1540
         Width           =   9465
         Begin VB.OptionButton Option62 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   345
            TabIndex        =   384
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton Option62 
            Caption         =   "B������"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   385
            Top             =   255
            Width           =   1380
         End
         Begin VB.OptionButton Option62 
            Caption         =   "D���Ӳ�"
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
            Index           =   4
            Left            =   5865
            TabIndex        =   387
            Top             =   270
            Width           =   1125
         End
         Begin VB.OptionButton Option62 
            Caption         =   "C����ʱ"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   386
            Top             =   270
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "9������һ���»��������������ʶ���ѣ���     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   62
            Left            =   60
            TabIndex        =   420
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   63
         Left            =   435
         TabIndex        =   417
         Top             =   2540
         Width           =   9465
         Begin VB.OptionButton Option63 
            Caption         =   "A������"
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
            Index           =   1
            Left            =   435
            TabIndex        =   388
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton Option63 
            Caption         =   "B������"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   389
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option63 
            Caption         =   "D���Ӳ�"
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
            Index           =   4
            Left            =   5880
            TabIndex        =   391
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton Option63 
            Caption         =   "C����ʱ"
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
            Index           =   3
            Left            =   4005
            TabIndex        =   390
            Top             =   270
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "10����ʹ����һ��ʱ���Ƿ�Ҳ�����Լ��ǹµ���һ���ˣ���     �� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   61
            Left            =   120
            TabIndex        =   418
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   64
         Left            =   510
         TabIndex        =   415
         Top             =   3540
         Width           =   9465
         Begin VB.OptionButton Option64 
            Caption         =   "A����"
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
            Index           =   1
            Left            =   345
            TabIndex        =   392
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton Option64 
            Caption         =   "B����һ��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   393
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option64 
            Caption         =   "D����֪��"
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
            Index           =   4
            Left            =   5490
            TabIndex        =   395
            Top             =   330
            Width           =   1275
         End
         Begin VB.OptionButton Option64 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   394
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "11���㱻��˵�˻������Ƿ���������ȡ�����ж�����     �� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   60
            Left            =   60
            TabIndex        =   416
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   65
         Left            =   510
         TabIndex        =   413
         Top             =   4540
         Width           =   9465
         Begin VB.OptionButton Option65 
            Caption         =   "A����"
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
            Index           =   1
            Left            =   345
            TabIndex        =   396
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton Option65 
            Caption         =   "B����ʱ��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   397
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option65 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   398
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "12�����Ƿ�һ��˵""Ҫ����""����ͽ��ţ���     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   59
            Left            =   60
            TabIndex        =   414
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   66
         Left            =   510
         TabIndex        =   411
         Top             =   5540
         Width           =   9465
         Begin VB.OptionButton Option66 
            Caption         =   "A����"
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
            Index           =   1
            Left            =   345
            TabIndex        =   399
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option66 
            Caption         =   "B����ʱ��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   400
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option66 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   401
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "13�����Ƿ��ܾ��ú�������ע���㣿��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   58
            Left            =   60
            TabIndex        =   412
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   67
         Left            =   525
         TabIndex        =   409
         Top             =   6540
         Width           =   9465
         Begin VB.OptionButton Option67 
            Caption         =   "A����"
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
            Index           =   1
            Left            =   345
            TabIndex        =   402
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton Option67 
            Caption         =   "B����ʱ��"
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
            Index           =   2
            Left            =   2130
            TabIndex        =   403
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option67 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   404
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "14������������������������Ƿ��ر����У���     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   57
            Left            =   60
            TabIndex        =   410
            Top             =   0
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   68
         Left            =   510
         TabIndex        =   379
         Top             =   7545
         Width           =   9465
         Begin VB.OptionButton Option68 
            Caption         =   "A����"
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
            Index           =   1
            Left            =   345
            TabIndex        =   405
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option68 
            Caption         =   "B����ʱ��"
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
            Index           =   2
            Left            =   2115
            TabIndex        =   406
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option68 
            Caption         =   "C������"
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
            Index           =   3
            Left            =   3990
            TabIndex        =   407
            Top             =   285
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "15�����ﲻ���ģ��Ƿ���Ҷ������Ҷ�������     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   56
            Left            =   60
            TabIndex        =   408
            Top             =   15
            Width           =   8835
         End
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdAdd 
         Height          =   405
         Left            =   525
         TabIndex        =   436
         Top             =   8250
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   714
         Caption         =   "���棬¼����һ��"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   33023
      End
      Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdPre 
         Height          =   405
         Index           =   9
         Left            =   5115
         TabIndex        =   453
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "��һҳ"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   End
   Begin MSComctlLib.TabStrip TabStripQuest 
      Height          =   855
      Left            =   10230
      TabIndex        =   2
      Top             =   465
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   1508
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��һ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ڶ�����(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ڶ�����(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���Ĳ���(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���Ĳ���(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���岿��(��)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���岿��(��)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInputMainMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rowHeight As Integer         '�и�
Public rows As Byte                 '����
Public intCurFrame As Byte
Private anwser(68) As String        '�ʾ��
Private Sub chkFinish_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cobCBFS_Expand()
    cmdSave.Enabled = True
End Sub


Private Sub cobType_Expand()
    cmdSave.Enabled = True

End Sub


Private Sub cmdAdd_Click()
    If checkData = False Then Exit Sub
    saveAnwser    '�����ʾ��
    Unload Me
        
    frmInputNo.Show vbModal, frmMain
    
End Sub

Private Sub cmdNext_Click(Index As Integer)
    TabStripQuest.Tabs(Index + 1).Selected = True
End Sub

Private Sub cmdPre_Click(Index As Integer)
    TabStripQuest.Tabs(Index - 1).Selected = True
End Sub

Private Sub Form_Load()
'On Error GoTo aaaa
    Me.WindowState = vbMaximized    '��󻯴���
    intCurFrame = 1
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    For i = 2 To 9
        Frame1(i).Visible = False
    Next
    
    lblSchool.caption = curSchool & "     " & curClass & "    �ʾ��ţ�" & curNo
    
    '�رյ�����
    frmMain.picLeft.Visible = False
    frmMain.mnuGuide.Checked = False
    
    
    '��ѧ���б�
    frmMain.mnuStudentList.Checked = True
    frmMain.picStudentList.Visible = True
    
    
    
Exit Sub
    
aaaa:
    MsgBox Err.Description, vbCritical

End Sub
Private Sub Form_Resize()
On Error Resume Next
    
    Cls
    PicTop.Width = Width
    TabStripQuest.Width = Width
    TabStripQuest.Top = PicTop.Height
    TabStripQuest.Left = 0
    TabStripQuest.Height = Height - PicTop.Height
    
    For i = 1 To 9
        Frame1(i).Top = TabStripQuest.ClientTop + 30
        Frame1(i).Left = TabStripQuest.Left + 50
        Frame1(i).Height = TabStripQuest.ClientHeight - 500
        Frame1(i).Width = TabStripQuest.ClientWidth - 100
    Next

End Sub
Private Sub cmdSave_Click()
    If checkData = False Then Exit Sub
    saveAnwser  '�����ʾ��
    
    MsgBox "������ϣ�", , "�ʾ�¼��/�༭"
    
    frmMain.cmdLeft_Click 1
    frmMain.picStudentList.Visible = False
    frmMain.mnuStudentList.Checked = False
    If GetINI("Main", "Guide") = "n" Then
        frmMain.picLeft.Visible = False
        frmMain.mnuGuide.Checked = False
    Else
        frmMain.picLeft.Visible = True
        frmMain.mnuGuide.Checked = True
    End If
    
    Unload Me
    
End Sub
Private Sub cmdExit_Click()
    IsEdit = False
    frmMain.cmdLeft_Click 1
    frmMain.picStudentList.Visible = False
    frmMain.mnuStudentList.Checked = False
    If GetINI("Main", "Guide") = "n" Then
        frmMain.picLeft.Visible = False
        frmMain.mnuGuide.Checked = False
    Else
        frmMain.picLeft.Visible = True
        frmMain.mnuGuide.Checked = True
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsEdit = False

End Sub

Private Sub TabStripQuest_Click()
    If TabStripQuest.SelectedItem.Index = intCurFrame Then Exit Sub
    Frame1(TabStripQuest.SelectedItem.Index).Visible = True
    Frame1(intCurFrame).Visible = False
    intCurFrame = TabStripQuest.SelectedItem.Index
    'loadItemData TabStripQuest.SelectedItem.Index
End Sub
Private Function checkData() As Boolean
    
    Dim optionCtl As Control
    Dim n As Byte
    Dim i As Byte
    
    checkData = False
    
    For i = 1 To 68
        anwser(i) = "-"
    Next
    
    For Each optionCtl In Me
        If TypeOf optionCtl Is OptionButton Then
            If optionCtl.value = True Then anwser(optionCtl.Container.Index) = Chr(optionCtl.Index + 64)
        End If
    Next

    For i = 1 To 68
        'tmp = tmp & i & ":" & anwser(i) & "     "
        If anwser(i) = "-" Then MsgBox getNo(i), vbExclamation, "¼����ʾ": Exit Function
    Next
    
    checkData = True
    
End Function

Private Function getNo(iNo As Byte) As String
    Dim model(5) As String
    Dim par, no As Byte
    model(1) = "һ"
    model(2) = "��"
    model(3) = "��"
    model(4) = "��"
    model(5) = "��"
    
    If iNo < 9 Then
        par = 1
        no = iNo
    Else
        par = Int((iNo - 9) / 15) + 2
        no = iNo - (par - 2) * 15 - 8
    End If

    getNo = "��" & model(par) & "����  ��" & no & "δ����"

End Function
Private Sub saveAnwser()
    Dim rs As ADODB.Recordset
    Dim sql, strField, strValue As String
    Dim i As Byte
    
    strField = "mClass,mNo,mAnwser"
    
    strValue = ""
    For i = 1 To 68
        strValue = strValue & Trim(anwser(i))
    Next
    If Len(strValue) <> 68 Then MsgBox "�ʾ����Ŀ����", vbCritical, "¼���ʾ�": Exit Sub
    
    strValue = "'" & curID & "'," & curNo & ",'" & strValue & "'"
    
    DBConnect
    
    Set rs = New ADODB.Recordset
    rs.Open "select mNo from main where mClass='" & curID & "' and mNo=" & curNo, Conn, 1, 1
    recc = rs.RecordCount
    rs.Close
    Set rs = Nothing
    
    If recc > 0 Then Conn.Execute "delete from main where mClass='" & curID & "' and mNo=" & curNo
    
    sql = "insert into main(" & strField & ") values(" & strValue & ")"
    
    Conn.Execute sql
    
    
    If Val(curNo) < 10 Then
        insertNo = "0" & Trim(curNo)
    Else
        insertNo = Trim(curNo)
    End If
    
    If IsEdit Then IsEdit = False: Exit Sub
    
    If IsNumeric(frmMain.tvStudentList.SelectedItem) Then
        frmMain.tvStudentList.SelectedItem.Parent.Selected = True
    End If
    
    frmMain.tvStudentList.Nodes.Add frmMain.tvStudentList.SelectedItem, tvwChild, , insertNo, 2, 3
    frmMain.tvStudentList.Sorted = True

End Sub
