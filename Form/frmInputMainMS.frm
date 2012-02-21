VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputMainMS 
   Caption         =   "数据修改"
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
      Caption         =   "学生个人及家庭基本情况"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   1
         Left            =   6915
         TabIndex        =   62
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
            Caption         =   "C、有两个或以上兄弟姐妹"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、有一个兄弟姐妹"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、独生子女"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "7、你在家中的情况: （   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、和亲生父母一起生活"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、父母离异或丧父丧母后重组家庭"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、单亲家庭"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不在父母身边跟随祖父母或外祖父母等亲属一起生活"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "8、你的家庭情况:（   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、3-8万"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、3万以下"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、8-20万"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、20万以上"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "6、你家庭的年总收入：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、研究生或以上"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、大学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、小学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、中学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "5、你母亲的学历：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、中学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、小学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、大学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、研究生或以上"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "4、你父亲的学历：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "F、穗东街"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "I、长洲街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "H、荔联街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "G、南岗街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、红山街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、鱼珠街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、黄埔街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、大沙街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、文冲街"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "1、你的家庭所在地：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、外来务工或农民"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、公务员 军人 医生 教师 律师等"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、机关、企事业单位领导或管理干部"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、无业或其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、工人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "2、你父亲的职业：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、无业或其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、机关、企事业单位领导或管理干部"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、公务员 军人 医生 教师 律师等"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、外来务工或农民"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、工人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "3、你母亲的职业：(   )"
            BeginProperty Font 
               Name            =   "宋体"
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
         Caption         =   "第一部分：学生个人及家庭基本情况"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "家庭教育情况(上)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "C、开始不接受，但下次有所改进，并逐渐接受"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不太接受"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、基本接受"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、诚心接受"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "8、在多数情况下，你对家长的批评教育持哪种态度?（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、谈心交流，共商对策并分析"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、反复强调考试成绩的重要性，不能骄傲，下次不一定有这么好运"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、给你买文具书籍"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、按事先许诺给些钱"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "7、为了鼓励你能继续取得优异成绩，通常你父母会采取哪种措施? （     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、偶尔参加"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   " D、很少参加"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、大部分参加"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、每次参加"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "6、学校召开的家长会，你的父母（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、谈心交流"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、联合老师共同教育"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、置之不理"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、责骂、打罚"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "5、当你的成绩不理想或犯错时，父母会采取什么手段教育你？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、看情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "4、父母在假期有给你报辅导班吗？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、运动健身"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、读书学习"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、喝酒唱K"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、打牌赌博"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "3、你父母或家长最经常参加的业余活动是什么？（   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、基本不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从来不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶然会"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "2、你是否做家务活？（   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常交谈"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔交谈"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   " D、从不交谈"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、几乎没有交谈"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "1、你的父母或家长经常和你交谈学校的事情吗？（   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   2
         Left            =   6915
         TabIndex        =   429
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   2
         Left            =   5100
         TabIndex        =   437
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Caption         =   "第二部分：家庭教育情况"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "家庭教育情况(下)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "C、买最流行的衣服、装饰"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、其它"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、上网、游戏、酒吧等"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、买书籍文具等"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "15、你的利是钱主要用在哪方面（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、同学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、老师"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、其他亲人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、父母亲"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "12、当你遇到不顺心的事时，你首先想到的倾诉对象是（      ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、民主和谐型：自由发表意见，谁说的有道理听谁的"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "A、专制顺从型：父母亲说了算，我只能服从"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "9、请选出你父母在处理与你的关系时，属于以下哪种类型?（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、很难一致"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不一致"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、基本一致"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、一致"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "10、你的父母在共同教育你的时候，是否注意保持意见一致? （     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不相信老师，为你抱不平"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不了了之"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、向你了解情况，如果是你错了就批评"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、不管三七二十一大骂你一顿"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "11、当老师或其他人""告你的状""，你父母会（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不熟悉"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不过问"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、比较熟悉"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、熟悉"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "13、父母亲是否熟悉你经常交往的朋友？（      ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、自食其力，对社会有用的人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、品德高尚的人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、出人头地光宗耀祖的人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、会赚钱懂享受的人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "14、你的父母亲希望你成为怎样的人？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   3
         Left            =   6915
         TabIndex        =   430
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   3
         Left            =   5100
         TabIndex        =   438
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Caption         =   "学校教育情况(上)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "A、特想帮助"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、如果和我要好的就帮"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、特高兴"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、和我无关"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "7、其他同学受到老师批评时，你怎么想？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、1小时以内"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、大约1--2小时"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "C、大约2--3小时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "2、目前，放学后你每天完成老师布置的作业平均要花的时间为（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、先独立思考，确实不会再请教别人"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、直接问会做的同学或其他人怎么做"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、什么都不做"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、抄答案"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "3、遇到作业不会做，你会选择怎样做？（　　） "
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、大约0--0.5小时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、大约0.5--1小时"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "C、大约1--2小时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "4、你每天在学校文艺体育（包括上体育、艺术课和做操等）活动的时间大约为多少？（   ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、觉得没面子，以后与老师对着干"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、接受教育"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "C、自尊心受到伤害，从此一蹶不振"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "5、你被老师批评教育后的态度是（   ） "
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "L、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "K、不善于与学生交往"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "J、知识欠渊博"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "I、不善于承认错误的老师"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "H、教学水平低"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "G、经常拖堂"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "F、外在形象不好"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、不关心同学"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、对学生不能一视同仁"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、不能做到言传身教，处处垂范"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、没有幽默感"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、工作不认真"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "6、你最不喜欢什么类型的老师？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、喜欢"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、比较喜欢"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不喜欢"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、说不清"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "1、你喜欢你所在的学校吗？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   4
         Left            =   6915
         TabIndex        =   431
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   4
         Left            =   5100
         TabIndex        =   439
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Caption         =   "第三部分：学校教育情况"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "学校教育情况(下)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "A、经常谈心，是你人际关系指导者"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、最关心成绩，偶尔关心一下生活和想法"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、根本不关心你"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、只关心成绩，其他不管"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "13、班主任老师对你的关心程度是（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、少数"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、极少"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、很普遍"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "8、学校是否有老师对同学体罚或变相体罚（如罚站、罚抄书、罚写字等）现象？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、极少"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、很普遍"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "10、教师平时上课迟到或在课堂上使用手机的情况是（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不喜欢"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、还可以"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、喜欢"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "11、你喜欢读课外书（报纸或杂志）吗？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、一般"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不以考试成绩高低评价学生"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、比较严重"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、很严重"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "12、你的班主任老师以考试成绩的高低来评价学生的情况是（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、很少"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不进行"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、一般"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "14、班主任老师是否利用班会课进行思想政治和理想信念等教育（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C.有但很少"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从来没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "15、你在学校接受过性教育吗？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、自己也这样做"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "B、自己不这样做，但不管别人怎么做"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、制止"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "9、当你见到违纪现象，比如扔废纸、塑料袋、乱涂乱划等现象时，你是（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   5
         Left            =   6915
         TabIndex        =   432
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   5
         Left            =   5100
         TabIndex        =   440
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Caption         =   "社会教育情况(上)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "A、没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、1家"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、3家或以上"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、2家"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "1、你家或学校附近有多少家中小学生经常光顾的黑网吧？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、从没去过"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "2、你是否进营业性网吧？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、很多"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、较少"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "3、你所居住的社区或你的学校附近有无商店或临时摊贩摆卖不健康的读物、光盘或危险的玩具？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从来没有"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、曾经试过"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "4、你吸烟吗？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、经常看见"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、偶尔看见过"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、没有看见过"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "5、你在自己居住的社区是否目睹过吸毒、聚众赌博、斗殴等不法行为？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、影视中男女过分亲昵的举动"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、黄色书刊录像、色情电话或网站或带有色情的电子游戏"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、没有接触过"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、卖淫嫖娼"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "6、你接触过下面哪些现象？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、学校和父母"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、报刊书籍、电视电影"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、网络"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、社区"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "7、你的禁毒知识主要通过什么方面了解？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   6
         Left            =   6915
         TabIndex        =   433
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   6
         Left            =   5100
         TabIndex        =   441
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Caption         =   "第四部分：社会教育情况"
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
      Caption         =   "社会教育情况(下)"
      BeginProperty Font 
         Name            =   "宋体"
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Caption         =   "B、记在心里，有朝一日报复"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、投诉他们，寻求法律援助"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "9、小飞的父母从外地来黄埔做生意，遭到执法人员的不公平对待。对此，你认为小飞的父母应该（ ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、扰乱了社会秩序，增加了不安全因素"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、增加了城市人口的就业压力"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、他们为黄埔建设做出了重要贡献"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "10、目前，黄埔外来务工人员越来越多，对此，你的看法是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不太赞同"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、很不赞同"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、比较赞同"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、非常赞同"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "11、你认为不正当性行为感染艾滋病是否值得同情？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、那是大人的事，自己管不了"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、公民有依法纳税的义务，应该补缴"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、不逃税挣不了大钱"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "12、如果你的亲戚中，有人逃税，你认为：（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、这是拉近同学之间感情的一种方式"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、会引起同学之间的互相攀比，容易使他们过分看重、追求金钱"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、自己也想这样，但家里经济条件不允许"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、这是个人的事情，只要经济条件允许，无所谓"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "13、现在有些中小学生过生日时，大摆晏席，邀请同学到饭店庆祝，对此你的看法是：（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、离婚会使孩子走上歧途、甚至犯罪的道路"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、离婚会让孩子对父母产生怨恨"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、离婚会影响到孩子对家庭和婚姻的看法"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、离婚会让孩子在同学中抬不起头"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、离婚会使孩子更懂事"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "14、你认为父母离婚对孩子会有哪些影响？（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
         Caption         =   "15、小华为了当上班长，请一些同学到饭店吃饭，以便获得更多的选票。对于这种做法，你认为：（      ）"
         Height          =   975
         Index           =   53
         Left            =   495
         TabIndex        =   297
         Top             =   7155
         Width           =   10710
         Begin VB.OptionButton Option53 
            Caption         =   "C、很正常，现在办事情都是这样，不请客送礼什么事也办不成"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、不好，但不这样就当不上班长"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、这样既不公平也不道德"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "15、小华为了当上班长，请一些同学到饭店吃饭，以便获得更多的选票。对于这种做法，你认为：（ ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、道德水准下降"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、环境恶化"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、社会治安"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "8、你认为黄埔区最急需关注的社会问题是（    ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   7
         Left            =   6915
         TabIndex        =   434
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   7
         Left            =   5100
         TabIndex        =   442
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Caption         =   "个人品质与心理健康情况(上)"
      BeginProperty Font 
         Name            =   "宋体"
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
            Caption         =   "I、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "H、没有信仰"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "F、伊斯兰教"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、儒家学说"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、共产主义"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、佛教"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、道家学说"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、基督教"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "1、你的信仰是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "E、其他"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、国家政要"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、科学家"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、影视或体育明星"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、历史杰出人物"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "2、你最欣赏（或崇拜）的人是（      ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、为了培养提高自己"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、起到带头作用"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、捞取政治资本"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "3、你入团（党）的动机是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、金钱是对社会做出贡献的人的应有回报"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、没有钱是万万不能的"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、钱越多，人的价值就越大"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "4、你对金钱的看法是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、支持"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、反对"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、无所谓"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "5、你对早恋的态度是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、学校"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、家庭"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、朋友"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、社会"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "6、你觉得不良性格养成的主要因素是（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、快乐"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、平静"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、烦躁"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、郁闷"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "7、你认为自己经常处于哪种情绪当中？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   8
         Left            =   6915
         TabIndex        =   435
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "下 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   8
         Left            =   5100
         TabIndex        =   443
         Top             =   8220
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   714
         Caption         =   "上 一 页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Caption         =   "第五部分：个人品质与心理健康情况"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "个人品质与心理健康情况(下)"
      BeginProperty Font 
         Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdExit 
         Height          =   405
         Left            =   3930
         TabIndex        =   452
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "不保存"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdSave 
         Height          =   405
         Left            =   2790
         TabIndex        =   445
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "保存"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Caption         =   "A、总是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、有时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "8、你是否认为自己的体型和面孔比别人难看？（     ） "
            BeginProperty Font 
               Name            =   "宋体"
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
            Name            =   "宋体"
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
            Caption         =   "A、总是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、有时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "9、到了一个新环境，你会主动结识朋友？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Name            =   "宋体"
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
            Caption         =   "A、总是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、经常"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、从不"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、有时"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "10、你和大家在一起时，是否也觉得自己是孤单的一个人？（     ） "
            BeginProperty Font 
               Name            =   "宋体"
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
            Name            =   "宋体"
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
            Caption         =   "A、是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、不一定"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "D、不知道"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不会"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "11、你被人说了坏话，是否想立即采取报复行动？（     ） "
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、有时是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "12、你是否一听说""要考试""心里就紧张？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、有时是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "13、你是否总觉得好象有人注意你？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、有时是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "14、你对收音机和汽车的声音是否特别敏感？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "A、是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "B、有时是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "C、不是"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "15、心里不开心，是否会乱丢、乱砸东西？（     ）"
            BeginProperty Font 
               Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdAdd 
         Height          =   405
         Left            =   525
         TabIndex        =   436
         Top             =   8250
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   714
         Caption         =   "保存，录入下一份"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      Begin 调查问卷信息查询系统.XPButton cmdPre 
         Height          =   405
         Index           =   9
         Left            =   5115
         TabIndex        =   453
         Top             =   8250
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   714
         Caption         =   "上一页"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
            Caption         =   "第一部分"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第二部分(上)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第二部分(下)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第三部分(上)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第三部分(下)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第四部分(上)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第四部分(下)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第五部分(上)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第五部分(下)"
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
Public rowHeight As Integer         '行高
Public rows As Byte                 '行数
Public intCurFrame As Byte
Private anwser(68) As String        '问卷答案
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
    saveAnwser    '保存问卷答案
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
    Me.WindowState = vbMaximized    '最大化窗口
    intCurFrame = 1
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    For i = 2 To 9
        Frame1(i).Visible = False
    Next
    
    lblSchool.caption = curSchool & "     " & curClass & "    问卷编号：" & curNo
    
    '关闭导航栏
    frmMain.picLeft.Visible = False
    frmMain.mnuGuide.Checked = False
    
    
    '打开学生列表
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
    saveAnwser  '保存问卷答案
    
    MsgBox "保存完毕！", , "问卷录入/编辑"
    
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
        If anwser(i) = "-" Then MsgBox getNo(i), vbExclamation, "录入提示": Exit Function
    Next
    
    checkData = True
    
End Function

Private Function getNo(iNo As Byte) As String
    Dim model(5) As String
    Dim par, no As Byte
    model(1) = "一"
    model(2) = "二"
    model(3) = "三"
    model(4) = "四"
    model(5) = "五"
    
    If iNo < 9 Then
        par = 1
        no = iNo
    Else
        par = Int((iNo - 9) / 15) + 2
        no = iNo - (par - 2) * 15 - 8
    End If

    getNo = "第" & model(par) & "部分  第" & no & "未作答！"

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
    If Len(strValue) <> 68 Then MsgBox "问卷答案数目出错！", vbCritical, "录入问卷": Exit Sub
    
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
