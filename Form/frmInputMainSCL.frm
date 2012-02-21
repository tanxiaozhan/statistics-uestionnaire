VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputMainSCL 
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
      Left            =   12420
      TabIndex        =   1
      Top             =   8355
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
            TabIndex        =   176
            Top             =   495
            Width           =   1290
         End
         Begin VB.OptionButton Option1 
            Caption         =   $"frmInputMainSCL.frx":0000
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
      Left            =   12270
      TabIndex        =   65
      Top             =   8070
      Width           =   10005
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   16
         Left            =   510
         TabIndex        =   109
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   114
            Top             =   30
            Width           =   6705
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   15
         Left            =   495
         TabIndex        =   103
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
            TabIndex        =   107
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
            TabIndex        =   106
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   108
            Top             =   30
            Width           =   7860
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   14
         Left            =   495
         TabIndex        =   97
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
            Left            =   4395
            TabIndex        =   101
            Top             =   345
            Width           =   1920
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
            Left            =   6480
            TabIndex        =   100
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
            TabIndex        =   99
            Top             =   345
            Width           =   1800
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
            TabIndex        =   98
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
            TabIndex        =   102
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   13
         Left            =   495
         TabIndex        =   91
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   96
            Top             =   45
            Width           =   7680
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   12
         Left            =   495
         TabIndex        =   85
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   90
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   11
         Left            =   495
         TabIndex        =   79
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   84
            Top             =   45
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   10
         Left            =   495
         TabIndex        =   73
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   78
            Top             =   30
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   9
         Left            =   540
         TabIndex        =   66
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   71
            Top             =   255
            Width           =   6225
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   2
         Left            =   6915
         TabIndex        =   240
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
         TabIndex        =   247
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
         TabIndex        =   72
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
      Top             =   7950
      Width           =   9915
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   19
         Left            =   480
         TabIndex        =   133
         Top             =   2700
         Width           =   9465
         Begin VB.OptionButton Option19 
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
            TabIndex        =   177
            Top             =   540
            Width           =   1050
         End
         Begin VB.OptionButton Option19 
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
            TabIndex        =   137
            Top             =   540
            Width           =   1065
         End
         Begin VB.OptionButton Option19 
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
            TabIndex        =   136
            Top             =   525
            Width           =   1050
         End
         Begin VB.OptionButton Option19 
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
            TabIndex        =   135
            Top             =   540
            Width           =   1470
         End
         Begin VB.OptionButton Option19 
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
            TabIndex        =   134
            Top             =   540
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "11、当你遇到不顺心的事时，你首先想到的倾诉对象是（      ）"
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
            TabIndex        =   138
            Top             =   255
            Width           =   6735
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   17
         Left            =   480
         TabIndex        =   127
         Top             =   675
         Width           =   9465
         Begin VB.OptionButton Option17 
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
            TabIndex        =   131
            Top             =   405
            Width           =   1740
         End
         Begin VB.OptionButton Option17 
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
            TabIndex        =   130
            Top             =   405
            Width           =   1560
         End
         Begin VB.OptionButton Option17 
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
            TabIndex        =   129
            Top             =   405
            Width           =   1590
         End
         Begin VB.OptionButton Option17 
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
            TabIndex        =   128
            Top             =   405
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "9、你的父母在共同教育你的时候，是否注意保持意见一致? （     ）"
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
            TabIndex        =   132
            Top             =   135
            Width           =   7650
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   885
         Index           =   18
         Left            =   480
         TabIndex        =   121
         Top             =   1665
         Width           =   9465
         Begin VB.OptionButton Option18 
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
            TabIndex        =   125
            Top             =   570
            Width           =   2955
         End
         Begin VB.OptionButton Option18 
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
            TabIndex        =   124
            Top             =   600
            Width           =   1995
         End
         Begin VB.OptionButton Option18 
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
            TabIndex        =   123
            Top             =   345
            Width           =   4050
         End
         Begin VB.OptionButton Option18 
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
            TabIndex        =   122
            Top             =   300
            Width           =   3225
         End
         Begin VB.Label Label1 
            Caption         =   "10、当老师或其他人""告你的状""，你父母会（     ）"
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
            TabIndex        =   126
            Top             =   45
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   20
         Left            =   495
         TabIndex        =   115
         Top             =   3810
         Width           =   9465
         Begin VB.OptionButton Option20 
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
            TabIndex        =   119
            Top             =   420
            Width           =   1350
         End
         Begin VB.OptionButton Option20 
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
            TabIndex        =   118
            Top             =   420
            Width           =   1560
         End
         Begin VB.OptionButton Option20 
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
            TabIndex        =   117
            Top             =   405
            Width           =   1605
         End
         Begin VB.OptionButton Option20 
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
            TabIndex        =   116
            Top             =   420
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "12、父母亲是否熟悉你经常交往的朋友？（      ）"
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
            TabIndex        =   120
            Top             =   90
            Width           =   5580
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   3
         Left            =   6915
         TabIndex        =   241
         Top             =   5175
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
         TabIndex        =   248
         Top             =   5175
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
      Left            =   11760
      TabIndex        =   4
      Top             =   7740
      Width           =   10320
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   27
         Left            =   510
         TabIndex        =   275
         Top             =   6945
         Width           =   9465
         Begin VB.OptionButton Option27 
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
            TabIndex        =   279
            Top             =   300
            Width           =   1410
         End
         Begin VB.OptionButton Option27 
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
            TabIndex        =   278
            Top             =   300
            Width           =   1380
         End
         Begin VB.OptionButton Option27 
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
            TabIndex        =   277
            Top             =   300
            Width           =   1560
         End
         Begin VB.OptionButton Option27 
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
            TabIndex        =   276
            Top             =   300
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "7、学校是否有老师对同学体罚或变相体罚（如罚站、罚抄书、罚写字等）现象？（    ）"
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
            TabIndex        =   280
            Top             =   45
            Width           =   9210
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   26
         Left            =   480
         TabIndex        =   169
         Top             =   5775
         Width           =   9465
         Begin VB.OptionButton Option26 
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
            TabIndex        =   173
            Top             =   525
            Width           =   1470
         End
         Begin VB.OptionButton Option26 
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
            TabIndex        =   172
            Top             =   525
            Width           =   2550
         End
         Begin VB.OptionButton Option26 
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
            TabIndex        =   171
            Top             =   525
            Width           =   1260
         End
         Begin VB.OptionButton Option26 
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
            TabIndex        =   170
            Top             =   525
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "6、其他同学受到老师批评时，你怎么想？（    ）"
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
            TabIndex        =   174
            Top             =   255
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   22
         Left            =   450
         TabIndex        =   163
         Top             =   1737
         Width           =   9465
         Begin VB.OptionButton Option22 
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
            TabIndex        =   167
            Top             =   375
            Width           =   1590
         End
         Begin VB.OptionButton Option22 
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
            TabIndex        =   166
            Top             =   375
            Width           =   1935
         End
         Begin VB.OptionButton Option22 
            Caption         =   $"frmInputMainSCL.frx":000D
            BeginProperty Font 
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
            TabIndex        =   165
            Top             =   375
            Width           =   1575
         End
         Begin VB.OptionButton Option22 
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
            TabIndex        =   164
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
            TabIndex        =   168
            Top             =   75
            Width           =   7365
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   23
         Left            =   450
         TabIndex        =   157
         Top             =   2724
         Width           =   9465
         Begin VB.OptionButton Option23 
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
            TabIndex        =   161
            Top             =   300
            Width           =   3795
         End
         Begin VB.OptionButton Option23 
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
            TabIndex        =   160
            Top             =   300
            Width           =   3840
         End
         Begin VB.OptionButton Option23 
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
            TabIndex        =   159
            Top             =   570
            Width           =   1935
         End
         Begin VB.OptionButton Option23 
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
            TabIndex        =   158
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
            TabIndex        =   162
            Top             =   0
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   24
         Left            =   450
         TabIndex        =   151
         Top             =   3711
         Width           =   9465
         Begin VB.OptionButton Option24 
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
            TabIndex        =   155
            Top             =   360
            Width           =   2085
         End
         Begin VB.OptionButton Option24 
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
            TabIndex        =   154
            Top             =   360
            Width           =   2100
         End
         Begin VB.OptionButton Option24 
            Caption         =   $"frmInputMainSCL.frx":001F
            BeginProperty Font 
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
            TabIndex        =   153
            Top             =   360
            Width           =   1560
         End
         Begin VB.OptionButton Option24 
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
            TabIndex        =   152
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
            TabIndex        =   156
            Top             =   60
            Width           =   9315
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   25
         Left            =   450
         TabIndex        =   145
         Top             =   4863
         Width           =   9465
         Begin VB.OptionButton Option25 
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
            TabIndex        =   149
            Top             =   270
            Width           =   3660
         End
         Begin VB.OptionButton Option25 
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
            TabIndex        =   148
            Top             =   270
            Width           =   1590
         End
         Begin VB.OptionButton Option25 
            Caption         =   $"frmInputMainSCL.frx":0031
            BeginProperty Font 
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
            TabIndex        =   147
            Top             =   510
            Width           =   1290
         End
         Begin VB.OptionButton Option25 
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
            TabIndex        =   146
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
            TabIndex        =   150
            Top             =   15
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   21
         Left            =   450
         TabIndex        =   139
         Top             =   885
         Width           =   9465
         Begin VB.OptionButton Option21 
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
            TabIndex        =   143
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton Option21 
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
            TabIndex        =   142
            Top             =   375
            Width           =   1530
         End
         Begin VB.OptionButton Option21 
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
            TabIndex        =   141
            Top             =   375
            Width           =   1560
         End
         Begin VB.OptionButton Option21 
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
            TabIndex        =   140
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
            TabIndex        =   144
            Top             =   90
            Width           =   6225
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   4
         Left            =   6915
         TabIndex        =   242
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
         TabIndex        =   249
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
         TabIndex        =   175
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
      Left            =   11640
      TabIndex        =   5
      Top             =   7545
      Width           =   9975
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   29
         Left            =   195
         TabIndex        =   189
         Top             =   1665
         Width           =   9465
         Begin VB.OptionButton Option29 
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
            TabIndex        =   192
            Top             =   390
            Width           =   1080
         End
         Begin VB.OptionButton Option29 
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
            TabIndex        =   191
            Top             =   390
            Width           =   1290
         End
         Begin VB.OptionButton Option29 
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
            TabIndex        =   190
            Top             =   405
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "9、教师平时上课迟到或在课堂上使用手机的情况是（    ）"
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
            TabIndex        =   193
            Top             =   120
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   30
         Left            =   495
         TabIndex        =   184
         Top             =   2625
         Width           =   9465
         Begin VB.OptionButton Option30 
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
            TabIndex        =   187
            Top             =   375
            Width           =   1740
         End
         Begin VB.OptionButton Option30 
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
            TabIndex        =   186
            Top             =   375
            Width           =   1380
         End
         Begin VB.OptionButton Option30 
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
            TabIndex        =   185
            Top             =   375
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "10、你喜欢读课外书（报纸或杂志）吗？（    ）"
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
            TabIndex        =   188
            Top             =   90
            Width           =   6225
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   28
         Left            =   510
         TabIndex        =   178
         Top             =   525
         Width           =   9465
         Begin VB.OptionButton Option28 
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
            TabIndex        =   182
            Top             =   675
            Width           =   1905
         End
         Begin VB.OptionButton Option28 
            Caption         =   $"frmInputMainSCL.frx":0040
            BeginProperty Font 
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
            TabIndex        =   181
            Top             =   645
            Width           =   3555
         End
         Begin VB.OptionButton Option28 
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
            TabIndex        =   180
            Top             =   360
            Width           =   3795
         End
         Begin VB.OptionButton Option28 
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
            TabIndex        =   179
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "8、当你见到违纪现象，比如扔废纸、塑料袋、乱涂乱划等现象时，你是（    ）"
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
            TabIndex        =   183
            Top             =   90
            Width           =   8085
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   5
         Left            =   6390
         TabIndex        =   243
         Top             =   3750
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
         Left            =   4575
         TabIndex        =   250
         Top             =   3750
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
      Caption         =   "社会教育情况"
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
      Left            =   11445
      TabIndex        =   6
      Top             =   7215
      Width           =   11310
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Caption         =   "15、小华为了当上班长，请一些同学到饭店吃饭，以便获得更多的选票。对于这种做法，你认为：（      ）"
         Height          =   975
         Index           =   33
         Left            =   435
         TabIndex        =   287
         Top             =   3450
         Width           =   10710
         Begin VB.OptionButton Option33 
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
            TabIndex        =   291
            Top             =   330
            Width           =   2745
         End
         Begin VB.OptionButton Option33 
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
            TabIndex        =   290
            Top             =   330
            Width           =   3615
         End
         Begin VB.OptionButton Option33 
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
            TabIndex        =   289
            Top             =   615
            Width           =   1560
         End
         Begin VB.OptionButton Option33 
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
            TabIndex        =   288
            Top             =   630
            Width           =   6090
         End
         Begin VB.Label Label1 
            Caption         =   "3、小华为了当上班长，请一些同学到饭店吃饭，以便获得更多的选票。对于这种做法，你认为：（ ）"
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
            TabIndex        =   292
            Top             =   30
            Width           =   10590
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   32
         Left            =   450
         TabIndex        =   281
         Top             =   2160
         Width           =   10815
         Begin VB.OptionButton Option32 
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
            TabIndex        =   285
            Top             =   285
            Width           =   4920
         End
         Begin VB.OptionButton Option32 
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
            TabIndex        =   284
            Top             =   300
            Width           =   4185
         End
         Begin VB.OptionButton Option32 
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
            TabIndex        =   283
            Top             =   600
            Width           =   6405
         End
         Begin VB.OptionButton Option32 
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
            TabIndex        =   282
            Top             =   585
            Width           =   3795
         End
         Begin VB.Label Label1 
            Caption         =   "2、现在有些中小学生过生日时，大摆晏席，邀请同学到饭店庆祝，对此你的看法是：（    ）"
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
            TabIndex        =   286
            Top             =   30
            Width           =   9870
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   31
         Left            =   510
         TabIndex        =   195
         Top             =   1095
         Width           =   10575
         Begin VB.OptionButton Option31 
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
            TabIndex        =   198
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton Option31 
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
            TabIndex        =   197
            Top             =   360
            Width           =   1380
         End
         Begin VB.OptionButton Option31 
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
            TabIndex        =   196
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "1、你所居住的社区或你的学校附近有无商店或临时摊贩摆卖不健康的读物、光盘或危险的玩具？（    ）"
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
            TabIndex        =   199
            Top             =   45
            Width           =   10500
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   6
         Left            =   7335
         TabIndex        =   244
         Top             =   4770
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
         Left            =   5520
         TabIndex        =   251
         Top             =   4770
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
         TabIndex        =   194
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
         TabIndex        =   253
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
      Index           =   7
      Left            =   11205
      TabIndex        =   63
      Top             =   6660
      Width           =   10575
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
         Index           =   40
         Left            =   405
         TabIndex        =   293
         Top             =   6690
         Width           =   9465
         Begin VB.OptionButton Option40 
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
            TabIndex        =   297
            Top             =   330
            Width           =   1740
         End
         Begin VB.OptionButton Option40 
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
            Left            =   5700
            TabIndex        =   296
            Top             =   330
            Width           =   1365
         End
         Begin VB.OptionButton Option40 
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
            TabIndex        =   295
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option40 
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
            TabIndex        =   294
            Top             =   330
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "7、你被人说了坏话，是否想立即采取报复行动？（     ） "
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
            TabIndex        =   298
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
         Index           =   39
         Left            =   405
         TabIndex        =   269
         Top             =   5850
         Width           =   9465
         Begin VB.OptionButton Option39 
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
            TabIndex        =   273
            Top             =   270
            Width           =   1740
         End
         Begin VB.OptionButton Option39 
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
            TabIndex        =   272
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton Option39 
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
            TabIndex        =   271
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option39 
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
            TabIndex        =   270
            Top             =   270
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "6、你和大家在一起时，是否也觉得自己是孤单的一个人？（     ） "
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
            TabIndex        =   274
            Top             =   0
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
         Index           =   38
         Left            =   495
         TabIndex        =   263
         Top             =   4890
         Width           =   9465
         Begin VB.OptionButton Option38 
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
            TabIndex        =   267
            Top             =   270
            Width           =   1740
         End
         Begin VB.OptionButton Option38 
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
            TabIndex        =   266
            Top             =   270
            Width           =   1125
         End
         Begin VB.OptionButton Option38 
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
            TabIndex        =   265
            Top             =   255
            Width           =   1380
         End
         Begin VB.OptionButton Option38 
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
            TabIndex        =   264
            Top             =   270
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "5、到了一个新环境，你会主动结识朋友？（     ）"
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
            TabIndex        =   268
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
         Index           =   37
         Left            =   465
         TabIndex        =   257
         Top             =   3990
         Width           =   9465
         Begin VB.OptionButton Option37 
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
            TabIndex        =   261
            Top             =   330
            Width           =   1740
         End
         Begin VB.OptionButton Option37 
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
            TabIndex        =   260
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton Option37 
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
            TabIndex        =   259
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option37 
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
            TabIndex        =   258
            Top             =   285
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "4、你是否认为自己的体型和面孔比别人难看？（     ） "
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
            TabIndex        =   262
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   34
         Left            =   495
         TabIndex        =   212
         Top             =   960
         Width           =   9465
         Begin VB.OptionButton Option34 
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
            TabIndex        =   239
            Top             =   630
            Width           =   1080
         End
         Begin VB.OptionButton Option34 
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
            TabIndex        =   216
            Top             =   315
            Width           =   1545
         End
         Begin VB.OptionButton Option34 
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
            TabIndex        =   215
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option34 
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
            TabIndex        =   214
            Top             =   585
            Width           =   2235
         End
         Begin VB.OptionButton Option34 
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
            TabIndex        =   213
            Top             =   360
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "1、你最欣赏（或崇拜）的人是（      ）"
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
            TabIndex        =   217
            Top             =   30
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   35
         Left            =   480
         TabIndex        =   206
         Top             =   2130
         Width           =   9465
         Begin VB.OptionButton Option35 
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
            TabIndex        =   210
            Top             =   315
            Width           =   1410
         End
         Begin VB.OptionButton Option35 
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
            TabIndex        =   209
            Top             =   300
            Width           =   1380
         End
         Begin VB.OptionButton Option35 
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
            TabIndex        =   208
            Top             =   270
            Width           =   1245
         End
         Begin VB.OptionButton Option35 
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
            TabIndex        =   207
            Top             =   300
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "2、你觉得不良性格养成的主要因素是（     ）"
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
            TabIndex        =   211
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   36
         Left            =   450
         TabIndex        =   200
         Top             =   3120
         Width           =   9465
         Begin VB.OptionButton Option36 
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
            TabIndex        =   204
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton Option36 
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
            TabIndex        =   203
            Top             =   270
            Width           =   1380
         End
         Begin VB.OptionButton Option36 
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
            TabIndex        =   202
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option36 
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
            TabIndex        =   201
            Top             =   270
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "3、你认为自己经常处于哪种情绪当中？（     ）"
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
            TabIndex        =   205
            Top             =   0
            Width           =   8835
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdNext 
         Height          =   405
         Index           =   7
         Left            =   6915
         TabIndex        =   245
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
         TabIndex        =   252
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
         TabIndex        =   238
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
      Index           =   8
      Left            =   180
      TabIndex        =   64
      Top             =   525
      Width           =   9780
      Begin 调查问卷信息查询系统.XPButton cmdExit 
         Height          =   405
         Left            =   4425
         TabIndex        =   255
         Top             =   4575
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
         Left            =   3285
         TabIndex        =   254
         Top             =   4575
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
         Height          =   585
         Index           =   41
         Left            =   510
         TabIndex        =   236
         Top             =   660
         Width           =   9465
         Begin VB.OptionButton Option41 
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
            TabIndex        =   219
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton Option41 
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
            TabIndex        =   220
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option41 
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
            Left            =   4005
            TabIndex        =   221
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "8、你是否一听说""要考试""心里就紧张？（     ）"
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
            TabIndex        =   237
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   42
         Left            =   510
         TabIndex        =   234
         Top             =   1575
         Width           =   9465
         Begin VB.OptionButton Option42 
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
            TabIndex        =   222
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option42 
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
            TabIndex        =   223
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option42 
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
            TabIndex        =   224
            Top             =   285
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "9、你是否总觉得好象有人注意你？（     ）"
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
            TabIndex        =   235
            Top             =   15
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   43
         Left            =   525
         TabIndex        =   232
         Top             =   2460
         Width           =   9465
         Begin VB.OptionButton Option43 
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
            TabIndex        =   225
            Top             =   330
            Width           =   1410
         End
         Begin VB.OptionButton Option43 
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
            TabIndex        =   226
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton Option43 
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
            TabIndex        =   227
            Top             =   330
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "10、你对收音机和汽车的声音是否特别敏感？（     ）"
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
            TabIndex        =   233
            Top             =   0
            Width           =   8835
         End
      End
      Begin VB.Frame FrameA 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   44
         Left            =   510
         TabIndex        =   218
         Top             =   3435
         Width           =   9465
         Begin VB.OptionButton Option44 
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
            TabIndex        =   228
            Top             =   285
            Width           =   1410
         End
         Begin VB.OptionButton Option44 
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
            TabIndex        =   229
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton Option44 
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
            TabIndex        =   230
            Top             =   285
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "11、心里不开心，是否会乱丢、乱砸东西？（     ）"
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
            TabIndex        =   231
            Top             =   15
            Width           =   8835
         End
      End
      Begin 调查问卷信息查询系统.XPButton cmdAdd 
         Height          =   405
         Left            =   1020
         TabIndex        =   246
         Top             =   4575
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
         Index           =   8
         Left            =   5610
         TabIndex        =   256
         Top             =   4575
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
         NumTabs         =   8
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
            Caption         =   "第四部分"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第五部分(上)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第五部分(下)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInputMainSCL"
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
    For i = 2 To 8
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
    
    For i = 1 To 44
        anwser(i) = "-"
    Next
    temp = ""
    For Each optionCtl In Me
        If TypeOf optionCtl Is OptionButton Then
            temp = temp & optionCtl.Container.Index & "   "
            If optionCtl.value = True Then anwser(optionCtl.Container.Index) = Chr(optionCtl.Index + 64)
        End If
    Next
    For i = 1 To 44
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
    
        If iNo <= 20 Then
            par = 2: no = iNo - 8
        Else
            If iNo <= 30 Then
                par = 3: no = iNo - 20
            Else
                If iNo <= 33 Then
                    par = 4: no = iNo - 30
                Else
                    If iNo <= 44 Then
                        par = 5: no = iNo - 33
                    End If
                End If
            End If
        End If
    End If

    getNo = "第" & model(par) & "部分  第" & no & "未作答！"

End Function
Private Sub saveAnwser()
    Dim rs As ADODB.Recordset
    Dim sql, strField, strValue As String
    Dim i As Byte
    
    strField = "mClass,mNo,mAnwser"
    
    strValue = ""
    For i = 1 To 44
        strValue = strValue & Trim(anwser(i))
    Next
    If Len(strValue) <> 44 Then MsgBox "问卷答案数目出错！", vbCritical, "录入问卷": Exit Sub
    
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
