VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_uploadShippingList 
   Caption         =   "�ϴ��ͻ�����������Ϣ"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   22040
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "DN�ϴ�"
      TabPicture(0)   =   "Frm_uploadShippingList.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "DN�޸�/ɾ��"
      TabPicture(1)   =   "Frm_uploadShippingList.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "DN����"
      TabPicture(2)   =   "Frm_uploadShippingList.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Toolbar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Fra(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SSTab2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin TabDlg.SSTab SSTab2 
         Height          =   9495
         Left            =   4080
         TabIndex        =   78
         Top             =   2760
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   16748
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "��������"
         TabPicture(0)   =   "Frm_uploadShippingList.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Fra(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "����鿴"
         TabPicture(1)   =   "Frm_uploadShippingList.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fps_Ship_del(1)"
         Tab(1).Control(1)=   "Fps_Ship_del(0)"
         Tab(1).Control(2)=   "Opt1"
         Tab(1).Control(3)=   "Opt2"
         Tab(1).Control(4)=   "Opt3"
         Tab(1).ControlCount=   5
         Begin VB.OptionButton Opt3 
            Caption         =   "����"
            Height          =   375
            Left            =   -72600
            TabIndex        =   95
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Opt2 
            Caption         =   "δ���"
            Height          =   375
            Left            =   -73680
            TabIndex        =   94
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "�����"
            Height          =   375
            Left            =   -74760
            TabIndex        =   93
            Top             =   600
            Width           =   975
         End
         Begin VB.Frame Fra 
            Caption         =   "������"
            ForeColor       =   &H00FF0000&
            Height          =   2415
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   6960
            Width           =   15735
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   2055
               Index           =   2
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   15495
               _Version        =   524288
               _ExtentX        =   27331
               _ExtentY        =   3625
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":008C
               TextTip         =   2
               AppearanceStyle =   0
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "�������������"
            ForeColor       =   &H00FF0000&
            Height          =   6495
            Index           =   2
            Left            =   120
            TabIndex        =   79
            Top             =   480
            Width           =   15700
            Begin VB.CheckBox Chk_All 
               Caption         =   "ȫѡ/ȫ��ѡ"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   240
               Width           =   1695
            End
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   5775
               Index           =   0
               Left            =   120
               TabIndex        =   81
               Top             =   500
               Width           =   8415
               _Version        =   524288
               _ExtentX        =   14843
               _ExtentY        =   10186
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":0525
               TextTip         =   2
               AppearanceStyle =   0
            End
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   5775
               Index           =   1
               Left            =   8640
               TabIndex        =   82
               Top             =   500
               Width           =   6975
               _Version        =   524288
               _ExtentX        =   12303
               _ExtentY        =   10186
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":09BE
               TextTip         =   2
               AppearanceStyle =   0
            End
            Begin VB.TextBox txtID_Text 
               Height          =   495
               Left            =   960
               TabIndex        =   80
               Top             =   5400
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin FPSpreadADO.fpSpread Fps_Ship_del 
            Height          =   7695
            Index           =   0
            Left            =   -74880
            TabIndex        =   91
            Top             =   1320
            Width           =   6615
            _Version        =   524288
            _ExtentX        =   11668
            _ExtentY        =   13573
            _StockProps     =   64
            EditEnterAction =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":0E57
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread Fps_Ship_del 
            Height          =   7695
            Index           =   1
            Left            =   -68160
            TabIndex        =   92
            Top             =   1320
            Width           =   8895
            _Version        =   524288
            _ExtentX        =   15690
            _ExtentY        =   13573
            _StockProps     =   64
            EditEnterAction =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":12F0
            TextTip         =   2
            AppearanceStyle =   0
         End
      End
      Begin VB.Frame Fra 
         Caption         =   "����ά��"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Index           =   1
         Left            =   4200
         TabIndex        =   51
         Top             =   1300
         Width           =   15735
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   9480
            TabIndex        =   64
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   900
            TabIndex        =   63
            Top             =   180
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   3840
            TabIndex        =   62
            Top             =   180
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   6720
            TabIndex        =   61
            Top             =   180
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   6
            ItemData        =   "Frm_uploadShippingList.frx":1789
            Left            =   900
            List            =   "Frm_uploadShippingList.frx":178B
            TabIndex        =   60
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   7
            Left            =   3840
            TabIndex        =   59
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
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
            Left            =   12000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
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
            Left            =   9480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   900
            TabIndex        =   56
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   3840
            TabIndex        =   55
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   6720
            TabIndex        =   54
            Top             =   1020
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   9
            Left            =   6720
            TabIndex        =   53
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "����"
            Height          =   345
            Left            =   11040
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ѡ����/SOLINE����"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   24
            Left            =   11520
            TabIndex        =   102
            Top             =   1080
            Width           =   1710
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0/0"
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   23
            Left            =   13560
            TabIndex        =   101
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ݱ��"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   8760
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   7
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   8
            Left            =   2900
            TabIndex        =   75
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���벿��"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   5800
            TabIndex        =   74
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ջ��ͻ�"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   2900
            TabIndex        =   73
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���䷽ʽ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   72
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ջ���ַ"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   12
            Left            =   5800
            TabIndex        =   71
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��       ע"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   8760
            TabIndex        =   70
            Top             =   720
            Width           =   675
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʒ��"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   120
            TabIndex        =   69
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   2900
            TabIndex        =   68
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ƴ̲���"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   5800
            TabIndex        =   67
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѹ�����/��ѯ������"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   18
            Left            =   8760
            TabIndex        =   66
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0/0"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   19
            Left            =   10560
            TabIndex        =   65
            Top             =   1080
            Width           =   45
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "��ѯ����"
         Height          =   10695
         Left            =   120
         TabIndex        =   33
         Top             =   1300
         Width           =   3855
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   11
            Left            =   1320
            TabIndex        =   99
            Top             =   3840
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   10
            Left            =   1320
            TabIndex        =   98
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Frame Frame9 
            Caption         =   "δOK"
            Height          =   2175
            Left            =   0
            TabIndex        =   86
            Top             =   8160
            Width           =   3735
            Begin VB.ListBox List_dn 
               Columns         =   2
               Height          =   1635
               Index           =   1
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   87
               Top             =   360
               Width           =   3375
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "������"
            Height          =   2895
            Left            =   0
            TabIndex        =   84
            Top             =   5160
            Width           =   3735
            Begin VB.ListBox List_dn 
               Columns         =   2
               Height          =   2310
               Index           =   0
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   85
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.CheckBox chkDZD 
            Caption         =   "����˵���ѡ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   4680
            Width           =   1575
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   8
            Left            =   1320
            TabIndex        =   41
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ѯʱ������ѡ���LOT�Ż��Ϻ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   120
            TabIndex        =   40
            Top             =   4320
            Width           =   3500
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   5
            Left            =   1320
            Style           =   1  'Simple Combo
            TabIndex        =   39
            Top             =   2280
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   4
            Left            =   1320
            Style           =   1  'Simple Combo
            TabIndex        =   38
            Top             =   1920
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   37
            Top             =   1560
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   36
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   35
            Top             =   765
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   0
            ItemData        =   "Frm_uploadShippingList.frx":178D
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1797
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   360
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtShipDate_Ship 
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   2640
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Format          =   114098177
            CurrentDate     =   43879
         End
         Begin VB.TextBox TxtShipDate_Ship 
            Height          =   285
            Left            =   3120
            TabIndex        =   103
            Top             =   2640
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�����"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   100
            Top             =   3600
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "      SO"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   97
            Top             =   3960
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "      DN"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   49
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�PO_NUM"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   48
            Top             =   2355
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��        ��"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   47
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Lot��"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   1635
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ⷿ����"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���߱��"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�����"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   43
            Top             =   435
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "�޸ļ�¼"
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Left            =   -74760
         TabIndex        =   28
         Top             =   8280
         Width           =   20055
         Begin FPSpreadADO.fpSpread fpS_Mod 
            Height          =   3375
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   19575
            _Version        =   524288
            _ExtentX        =   34528
            _ExtentY        =   5953
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":17A6
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "��ѯ���"
         ForeColor       =   &H00FF0000&
         Height          =   5175
         Left            =   -74760
         TabIndex        =   27
         Top             =   2880
         Width           =   20055
         Begin FPSpreadADO.fpSpread fpS_Mod 
            Height          =   4335
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   19575
            _Version        =   524288
            _ExtentX        =   34528
            _ExtentY        =   7646
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":1BC8
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����ѡ��"
         ForeColor       =   &H00FF0000&
         Height          =   2295
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   20055
         Begin VB.ComboBox CbItem_Mod 
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":1FEA
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1FEC
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1920
            Width           =   2055
         End
         Begin VB.TextBox txtShipDate_Mod 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox cbCustomerCode_Mod 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":1FEE
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1FF8
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtDN_Mod 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   19
            Top             =   1575
            Width           =   2055
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   11280
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":2007
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":2C59
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":38AB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":44FD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":514F
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   870
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   17265
            _ExtentX        =   30454
            _ExtentY        =   1535
            ButtonWidth     =   1455
            ButtonHeight    =   1482
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "��ѯ"
                  Key             =   "QUERY_MOD"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ȷ���޸�"
                  Key             =   "SAVE_MOD"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ȷ��ɾ��"
                  Key             =   "DEL_MOD"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "�˳�"
                  Key             =   "HOME_MOD"
                  ImageIndex      =   5
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSComCtl2.DTPicker dtShipDate_mod 
            Height          =   375
            Left            =   8400
            TabIndex        =   23
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   65280
            CalendarTitleForeColor=   65535
            CalendarTrailingForeColor=   16777215
            Format          =   114098177
            CurrentDate     =   43271
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   12000
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
            Flags           =   524800
            MaxFileSize     =   9999
         End
         Begin VB.Label Label6 
            Caption         =   "�޸���Ŀ"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�����:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   26
            Top             =   1230
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   840
            TabIndex        =   25
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ѡ���µĳ�������:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   24
            Top             =   1250
            Width           =   1875
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "������ϸ"
         ForeColor       =   &H00FF0000&
         Height          =   9255
         Left            =   -74760
         TabIndex        =   13
         Top             =   3000
         Width           =   20055
         Begin VB.TextBox txtDNCheck 
            BackColor       =   &H00FFC0FF&
            Height          =   4095
            Left            =   12960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   480
            Width           =   5175
         End
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   8775
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   12495
            _Version        =   524288
            _ExtentX        =   22040
            _ExtentY        =   15478
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   80
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":5DA1
            TextTip         =   2
         End
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   3855
            Index           =   1
            Left            =   12960
            TabIndex        =   16
            Top             =   4920
            Width           =   6855
            _Version        =   524288
            _ExtentX        =   12091
            _ExtentY        =   6800
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":621F
            TextTip         =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DateCode ���"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   12960
            TabIndex        =   17
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����ѡ��"
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   20055
         Begin VB.TextBox txtDN 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4560
            TabIndex        =   5
            Top             =   1215
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cbCustomerCode 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":6691
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":669B
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtFileName 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1680
            Width           =   9015
         End
         Begin VB.TextBox txtShipDate 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1163
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   11280
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":66AA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":72FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":7F4E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":8BA0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":97F2
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   600
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            AllowCustomize  =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "��ѯ"
                  Key             =   "QUERY"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "��"
                  Key             =   "OPEN"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "����"
                  Key             =   "SAVE"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ɾ��"
                  Key             =   "DEL"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "�˳�"
                  Key             =   "HOME"
                  ImageIndex      =   5
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSComCtl2.DTPicker dtShipDate 
            Height          =   375
            Left            =   8880
            TabIndex        =   7
            Top             =   1163
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   65280
            CalendarTitleForeColor=   65535
            CalendarTrailingForeColor=   16777215
            Format          =   114098177
            CurrentDate     =   43271
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   12000
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
            Flags           =   524800
            MaxFileSize     =   9999
         End
         Begin VB.Label lblShipDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7800
            TabIndex        =   12
            Top             =   1230
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4080
            TabIndex        =   11
            Top             =   1230
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblCustomerCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�����:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   10
            Top             =   1230
            Width           =   975
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ļ���(N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label lbl123 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DN�Ѱ�����������,�ϴ�����ѡ���������"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   11880
            TabIndex        =   8
            Top             =   1200
            Width           =   4485
         End
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   870
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   19665
         _ExtentX        =   34687
         _ExtentY        =   1535
         ButtonWidth     =   1455
         ButtonHeight    =   1482
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѯ"
               Key             =   "QUERY_Ship"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȷ��"
               Key             =   "SAVE_Ship"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȷ��ɾ��"
               Key             =   "DEL_Ship"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "HOME_Ship"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
End
Attribute VB_Name = "Frm_uploadShippingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public intFlag           As Integer '�������
Public strRealName As String
Public lsolineqty           As Long 'soline����
Public strsono           As String 'sono
Public strsoline           As String 'soline
Public mailflag As Integer

Private Enum FpsM
    E_CHOOSE = 1        'ѡ��
    e_ID                '����ID
    E_cust              '�ͻ�����
    E_DN                'DN
    e_GDH               '�����Ż�LOT��
    e_BigX              '�����
    e_LH                '�Ϻ�
    e_NUM               '�����
    E_GG                '���
    E_XH                '�ͺ�
    E_UNIT              '��λ
    e_KF                '�����ⷿ

    e_MCol
End Enum
Private Enum FpsD
    e_BigX              '����� �����ֶ�
    E_XH                '���
    e_GDH               '�����Ż�LOT��
    e_LCK               '���̿�
    e_LH                '�Ϻ�
    e_GNum              '�ϸ���
    e_BLNum             '�ͻ�������
    e_ZCNum             '�Ƴ̲�����
    e_ID                '��ϸID
    e_KF                '�����ⷿ
    E_DN                'DN
   ' E_Note               'DN
    E_Note2               'DN
    e_MCol
End Enum

Dim adoPrmReturn        As ADODB.Parameter
Dim adoprm1             As ADODB.Parameter
Dim adoprm2             As ADODB.Parameter
Dim adoPrm3             As ADODB.Parameter
Dim adoPrm4             As ADODB.Parameter
Dim adoPrm5             As ADODB.Parameter
Dim adoPrm6             As ADODB.Parameter
Dim adoPrm7             As ADODB.Parameter
Dim adoPrm8             As ADODB.Parameter
Dim adoPrm9             As ADODB.Parameter
Dim adoprmFG            As ADODB.Parameter
Dim adoprm10            As ADODB.Parameter
Dim adoPrm11            As ADODB.Parameter
Dim adoPrm12            As ADODB.Parameter
Dim adoPrm13            As ADODB.Parameter
Dim adoPrm14            As ADODB.Parameter
Dim adoPrm15            As ADODB.Parameter
Dim adoprm16            As ADODB.Parameter




Private Sub cbCustomerCode_Click()

Select Case cbCustomerCode.text

    Case "37"  '37
        lblDN.Visible = True
        txtDN.Visible = True
        lblShipDate.Visible = True
        dtShipDate.Visible = True
        txtShipDate.Visible = True
        Toolbar1.Buttons(7).Visible = True
        txtDNCheck.Visible = True
        fps(1).Visible = True
        lbl123.Visible = True
        Label1.Visible = True
    Case "SG005"
    
        lblDN.Visible = False
        txtDN.Visible = False
        lblShipDate.Visible = False
        dtShipDate.Visible = False
        txtShipDate.Visible = False
        Toolbar1.Buttons(7).Visible = False
        txtDNCheck.Visible = False
        fps(1).Visible = False
        lbl123.Visible = False
        Label1.Visible = False
    Case Else
        lblDN.Visible = False
        txtDN.Visible = False
        lblShipDate.Visible = False
        dtShipDate.Visible = False
        txtShipDate.Visible = False

End Select

End Sub

Private Function CheckDC() As Boolean
Dim i             As Integer
Dim strDN         As String
Dim strJobID      As String
Dim strDNDC       As String
Dim str37TESTDC   As String
Dim strSql        As String
Dim strWrongJobID As String
Dim bWrongJobID   As Boolean
txtDNCheck.text = ""
txtDNCheck.text = "DN                   JOBID                DNDC          ��ȷDC" & vbCrLf
bWrongJobID = False
strWrongJobID = ""

With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        strDN = Trim$("" & .text)
        .Col = 32
        strJobID = Trim$("" & .text)
        .Col = 52
        strDNDC = Trim$("" & .text)

        If Right$(strJobID, 1) = "M" Then
            str37TESTDC = Get_SqlStr("select DC from erptemp..tbl37testdc_m where JOBID = '" & strJobID & "'")
        Else
            str37TESTDC = Get_OracleStr("select dc from tbl37testdc where jobid = '" & strJobID & "'")

        End If

        If str37TESTDC = "" Then
            MsgBox "JOBID:" & strJobID & "��ѯ�������ض�Ӧ��DC,�޷����ͻ�DN�Ƿ���ȷ,����ϵIT", vbCritical, "����"
            bWrongJobID = True
        Else

            If str37TESTDC <> "" And str37TESTDC <> strDNDC Then
                txtDNCheck.text = txtDNCheck.text & strDN & "       " & strJobID & "           " & strDNDC & "            " & str37TESTDC & vbCrLf
                bWrongJobID = True

            End If

        End If

    Next i

End With

If bWrongJobID = False Then
    CheckDC = True
Else
    CheckDC = False

End If

End Function





Private Sub CancerSelection(ROW_S, ROW_E)
Dim i As Integer
Dim strBigbox As String

     With Fps_Ship(0)
     For i = 1 To .MaxRows
         .Row = i
         .Col = FpsM.E_CHOOSE
         If .text = "1" Then
             Call fps_Ship_Click(0, 1, i)
         End If
     Next
     End With
     
End Sub


Private Sub cbCustomerCode_Mod_Click()

    If cbCustomerCode_Mod.text = "37" Then
        CbItem_Mod.Clear
        CbItem_Mod.AddItem ("�޸ĳ�������")
        CbItem_Mod.AddItem ("�޸ĳ�������")
        CbItem_Mod.AddItem ("ɾ��")
        CbItem_Mod.AddItem ("��ѯ�����޸ļ�¼")
        Toolbar2.Buttons(3).Enabled = True
        Label3.Caption = "D.N:"
    ElseIf cbCustomerCode_Mod.text = "SG005" Then
        CbItem_Mod.Clear
        CbItem_Mod.AddItem ("ɾ��")
        Toolbar2.Buttons(3).Enabled = False
        Label3.Caption = "S.O:"
        Label2.Visible = False
        dtShipDate_mod.Visible = False
        txtShipDate_Mod.Visible = False
    End If
    
End Sub

Private Sub Chk_All_Click()

    Dim i As Integer
    
    If Chk_All.Value = 1 Then

         With Fps_Ship(0)
             For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                If Val(.Value) = 0 Then
                    Call fps_Ship_Click(0, 1, i)
                End If
             Next i
         End With

    
        
    ElseIf Chk_All.Value = 0 Then

         With Fps_Ship(0)
             For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                If Val(.Value) = 1 Then
                    Call fps_Ship_Click(0, 1, i)
                End If
             Next i
         End With
        
    End If
    
    
    
End Sub

Private Sub Cob_Change(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer


 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = True '��������
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '�ͻ�����
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '�ͻ�����
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = ""
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = False '��������
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '�ͻ�����
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '�ͻ�����
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
       
    
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

    

End Sub

Private Sub Cob_Click(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer

 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = True '��������
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '�ͻ�����
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '�ͻ�����
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = 0
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = True '��������
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '�ͻ�����
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '�ͻ�����
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

End Sub


Private Sub Cob_DblClick(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer


 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = True '��������
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '�ͻ�����
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '�ͻ�����
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = 0
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '��������
        dtShipDate_Ship.Visible = True '��������
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '�ͻ�����
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '�ͻ�����
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
       
    
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

End Sub
Private Sub GetOptSelection()
    Dim i As Integer
    Dim j As Integer
    Dim strBigbox As String
    Dim strqty As String
    Dim strBigbox_temp As String
    Dim strBb_temp As String
    Dim lQty_temp As Long
    Dim lQty_select As Long


    strBigbox_temp = ""
    strBb_temp = ""
    strBigbox = ""
    lQty_select = 0
    lQty_temp = 0
        
    With Fps_Ship(0)
    '1.ͳ��ÿ�����������
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.e_BigX
            If InStr(strBigbox, Trim(.text)) = 0 Then
                strBigbox = strBigbox & "," & Trim(.text)
            End If
        Next
        '�ҳ��������
        For i = 0 To UBound(Split(strBigbox, ","))
            strBigbox_temp = Split(strBigbox, ",")(i)
            If strBigbox_temp <> "" Then
                lQty_temp = 0
                For j = 1 To .MaxRows
                    .Row = j
                    .Col = FpsM.e_BigX
                    If Trim(.text) = strBigbox_temp Then
                        .Col = FpsM.e_NUM
                        lQty_temp = lQty_temp + Val(.text)
                    End If
                Next
                If lQty_select + lQty_temp <= lsolineqty Then
                   lQty_select = lQty_select + lQty_temp
                   For j = 1 To .MaxRows
                      .Row = j
                      .Col = FpsM.e_BigX
                      strBb_temp = Trim(.text)
                      .Col = 1
                      If strBb_temp = strBigbox_temp And Val(.Value) = 0 Then
                          Call fps_Ship_Click(0, 1, j)
                      End If
                   Next
                End If
            End If
        Next

     '   For i = 1 To .MaxRows
     '       .Row = i
     '       .Col = FpsM.E_CHOOSE
     '       If Val(.Value) = 0 Then
     '           .Col = FpsM.e_BigX
     '           If Trim(.text) <> strBigbox Then
     '               strBigbox=
     '               For j = i To .MaxCols
      '
                        
     '               Next
                    
                
                
                
            
          '  End If
            
        
        
        'Next
        
        
        
        
    End With
    

End Sub





Private Sub showdata_shiplist(cust As String, ordernolist As String, ShipDate As String)
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim strSql As String
Dim strDN As String
Dim j As Integer
Dim strTemp As String
If Trim(ordernolist) = "" Then
    Exit Sub
End If
If UCase(cust) = "37" Then
    
       strSql = "  SELECT CUST_DEVICE AS �ͻ�����,  Quality as   '����(ea)', HT_DEVICE AS ���ڻ��� ,dn as  'DN(#)',ShipOrder as ���ݺ� ,case BOND when 'A' THEN '��˰' else '�Ǳ�'end as  ��˰�Ǳ� ,SHIP_DATE as �������� ,isnull(remark1,'') AS ���,isnull(remark2,'') AS ������  " & _
                 "  from erpdata..tblShipOrder_Dn where shiporder in ('" & ordernolist & "') order by shiporder "
                
                
                
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(2)
               
                .MaxRows = 0
                 Set .DataSource = rs
                 
                strTemp = "<html><body><table  border=1 cellpadding=10>"
                For i = 0 To .MaxRows
                    
                    .Row = i
                    If i = 0 Then
                        strTemp = strTemp & "<tr bgcolor = #E6E6FA>"
                    Else
                        strTemp = strTemp & "<tr>"
                    End If
                    For j = 1 To .MaxCols
                        .Col = j
                        strTemp = strTemp & add_tr(Trim(.text))
                    Next
                    strTemp = strTemp & "</tr>"
                Next
                strTemp = strTemp & "</table></br></br>" '����
                
               
            End With
            
            strSql = "  SELECT CUST_DEVICE AS �ͻ�����,  Quality as   '����(ea)', HT_DEVICE AS ���ڻ��� ,dn as  'DN(#)',ShipOrder as ���ݺ� ,case BOND when 'A' THEN '��˰' else '�Ǳ�'end as  ��˰�Ǳ� ,SHIP_DATE as �������� ,isnull(remark1,'') AS ��� ,isnull(remark2,'') AS ������ " & _
                     "  from erpdata..tblShipOrder_Dn where isnull(dn,'')<>'' and SHIP_DATE='" & ShipDate & "' order by shiporder "
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If rs.RecordCount > 0 Then
                strTemp = strTemp & "����Ϊ" & ShipDate & "�������� <p></p><Table  border=1 cellpadding=10>"
                rs.MoveFirst
                '����
                strTemp = strTemp & "<tr bgcolor = #E6E6FA >"
                For j = 0 To rs.Fields.count - 1
                    strTemp = strTemp & add_tr(rs.Fields(j).name)
                Next
                strTemp = strTemp & "</tr>"
                For i = 1 To rs.RecordCount
                    strTemp = strTemp & "<tr >"
                    For j = 0 To rs.Fields.count - 1
        
                        strTemp = strTemp & add_tr(rs.Fields(j))
  
                    Next
                    strTemp = strTemp & "</tr>"
                    rs.MoveNext

                Next
                
                strSql = "SELECT sum(Quality) as �ϼ����� from erpdata..tblShipOrder_Dn where isnull(dn,'')<>'' and SHIP_DATE='" & ShipDate & "'"
                strTemp = strTemp & "<tr  border=1><th>�ϼ�</th><th> " & Get_SqlserverNo(strSql) & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
            End If
        
            strTemp = strTemp & "</table></body></html>"
        
        Call SentMesToStock(strTemp)
        
        End If
        
        
        
    ElseIf UCase(cust) = "SG005" Then
        strSql = "  SELECT a.HT_DEVICE AS ������ ,a.ShipOrder as ���ݺ�,a.PCSNUM as ����Ƭ��, a.Quality as  ��������,case a.BOND when 'A' THEN '��˰' else '�Ǳ�'end as  ��˰�Ǳ� , a.SO_NO,a.SO_LINE ,a.remark1 as TERM,a.remark2 as ���� ,case when b.�ⷿ���� like '%����%' then '����' else '��Ʒ' end as '��Ʒ/����'" & _
                 "  from erpdata..tblShipOrder_Dn a , erpdata..tblstock b where a.stockid=b.�ⷿ���� and a.shiporder in ('" & ordernolist & "') order by shiporder"
   
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(2)
               
                .MaxRows = 0
                 Set .DataSource = rs
                strTemp = "<html><body><table  border=1 cellpadding=10>"
                For i = 0 To .MaxRows
                    If i = 0 Then
                        strTemp = strTemp & "<tr bgcolor = #E6E6FA >"
                    Else
                        strTemp = strTemp & "<tr>"
                    End If
                    .Row = i
                    For j = 1 To .MaxCols
                        .Col = j
                        strTemp = strTemp & add_tr(Trim(.text))
                    Next
                    strTemp = strTemp & "</tr>"
                Next
                strTemp = strTemp & "</table></body></html>"
               Call SentMesToStock(strTemp)
            End With
        End If
    End If

End Sub


Private Function add_tr(strTemp As String)
    add_tr = "<th>  " & strTemp & "  </th>"
End Function


Private Sub dtShipDate_Change()
txtShipDate.text = dtShipDate.Value

End Sub

Private Sub dtShipDate_Click()
txtShipDate.text = dtShipDate.Value

End Sub

Private Sub dtShipDate_CloseUp()
txtShipDate.text = dtShipDate.Value

End Sub


Private Sub dtShipDate_mod_CloseUp()
txtShipDate_Mod.text = dtShipDate_mod.Value

End Sub

Private Sub dtShipDate_mod_Change()
txtShipDate_Mod.text = dtShipDate_mod.Value

End Sub

Private Sub dtShipDate_mod_Click()
txtShipDate_Mod.text = dtShipDate_mod.Value
End Sub




Private Sub dtShipDate_Ship_Change()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub



Private Sub dtShipDate_Ship_Click()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub

Private Sub dtShipDate_Ship_CloseUp()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub

Private Sub Form_Load()
InitCtrls
InitCtrl_Ship
End Sub

Private Sub InitCtrls()
Dim i As Integer

dtShipDate.Value = Format(Now(), "yyyy-MM-dd")
dtShipDate_mod.Value = Format(Now(), "yyyy-MM-dd")
dtShipDate_Ship.Value = Format(Now(), "yyyy-MM-dd")

cbCustomerCode.text = "37"
With fps(0)

    .Col = -1
    .Row = -1
    .Lock = True
    .TypeMaxEditLen = 500

End With

With fps(1)
    .TypeMaxEditLen = 500

    .Col = -1
    .Row = -1
    .Lock = True
    
    .SetText 1, 0, "DN"
    .SetText 2, 0, "QTY"
    .SetText 3, 0, "��˰���"
    .SetText 4, 0, "��������"
    
    .ColWidth(1) = 10
    .ColWidth(2) = 10
    .ColWidth(3) = 15

End With

    'Fps��ʼ��
    With fpS_Mod(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '�趨������
        .Col = 1  'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(1) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
    

End Sub





Private Sub fpS_Mod_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String

If Row < 1 Then Exit Sub
If Col <> 1 Then Exit Sub

Select Case CbItem_Mod.text


Case "�޸ĳ�������"


    With fpS_Mod(0)

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)

        If Val(.Value) = 1 Then
            '������һ����DN��ѡ����
            .Col = 2
            .Row = Row
            strTmp = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                If Trim$(.text) = strTmp Then
                    .Col = 1
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
        Else
            '������һ����DN��ѡ����
            .Col = 2
            .Row = Row
            strTmp = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                If Trim$(.text) = strTmp Then
                    .Col = 1
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
    
    



Case "�޸ĳ�������", "ɾ��"

    With fpS_Mod(0)

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)

        If Val(.Value) = 1 Then
            .Col = -1
            .ForeColor = &HFF8080

        Else
            .Col = -1
            .ForeColor = vbBlack

        End If
    End With

End Select




End Sub











Private Sub Fps_Ship_del_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i               As Long
Dim j               As Integer
Dim strorder      As String
Dim rs          As New ADODB.Recordset



    'Fps����¼�
    If Index <> 0 Then Exit Sub
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    With Fps_Ship_del(0)
        .Row = Row
        .Col = 1 '�����ѡ��
        .Value = Abs(Val(.Value) - 1)
        For i = 1 To .MaxRows
             .Row = i
             .Col = 1
             If Val(.Value) = 1 Then
                 .Col = 4
                 If strorder = "" Then
                     strorder = Trim(.text)
                 Else
                     strorder = strorder & "," & Trim(.text)
                 End If
             End If
        Next
    End With
    Set rs = Get_SqlserveRs("select * from erpdata..tblstocksqfhsub where ���ݱ�� in ('" & Replace(strorder, ",", "','") & "') order by ���ݱ�� ,������� ")

    With Fps_Ship_del(1)
        .MaxRows = 0
        Set .DataSource = rs
        
    End With
 

        
        
   
End Sub

Private Sub Opt1_Click()
ListOrderNumData
End Sub

Private Sub Opt2_Click()
ListOrderNumData
End Sub

Private Sub Opt3_Click()
ListOrderNumData
End Sub




Private Sub ListOrderNumData()

Dim strSql As String
Dim rs As ADODB.Recordset

     
  strSql = "select distinct  '0' as ѡ�� ,s.�ͻ�����,s.�ͻ�����,t.���ݱ�� , " & _
        " isnull(��������,'') as �������� from erpdata..tblStockSQfh t inner join erpbase..tblXCustomer s on s.�ͻ�����=t.�ͻ����� where t.���ձ��=0  " & _
        " and t.����Ա='" & txt(1).text & "' and t.��������=1 "
        
       
    If Opt1.Value = True Then '�����
    
     strSql = strSql & " and isnull(t.�������,'')<>''"
            
    End If
           
    If Opt2.Value = True Then 'δ���
    
     strSql = strSql & " and isnull(t.�������,'')=''"
     
    End If
    
       

    strSql = strSql & " order by s.�ͻ�����,t.���ݱ��"
    Set rs = Get_SqlserveRs(strSql)
    
    With Fps_Ship_del(1)
        .MaxRows = 0
        
    End With
    
    With Fps_Ship_del(0)
        .MaxRows = 0
        Set .DataSource = rs
        
    End With

 

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
   control_resize
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)


control_resize
End Sub






Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        QueryData

    Case "OPEN"
        openData

    Case "SAVE"
        If cbCustomerCode.text = "37" Then
            SaveData
        ElseIf cbCustomerCode.text = "SG005" Then
            Call uploadSO(txtFileName.text)
        End If

    Case "DEL"
        delData

    Case "HOME"
        exitFrm

End Select

End Sub

Private Sub QueryData()

If cbCustomerCode.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

Select Case cbCustomerCode.ListIndex

    Case 0
        queryData_37

End Select

End Sub

Private Sub queryData_37()

    Dim strDN  As String
    Dim strSql As String
    Dim rs     As New ADODB.Recordset

    strDN = Trim$(txtDN.text)

    If strDN = "" Then
        If txtShipDate.text = "" Then
            strSql = "select a.lastupdatedate as ��������,a.* from CUSTOMERSHIPPINGUPTBL a order by a.id desc "
        Else
            strSql = "select a.lastupdatedate as ��������,a.* from CUSTOMERSHIPPINGUPTBL a where a.lastupdatedate = '" & Trim(txtShipDate.text) & "'  order by a.id desc "

        End If

    Else
        strSql = "select a.lastupdatedate as ��������,a.* from CUSTOMERSHIPPINGUPTBL a where delivery = '" & strDN & "' order by a.id desc "

    End If

    'Fps
'    Set rs = Get_OracleRs(strSql)
'
'    If Not rs.EOF Then
'
'        With Fps(0)
'            .MaxRows = 0
'            Set .DataSource = rs
'            Toolbar1.Buttons(5).Enabled = False
'
'        End With
'
'    Else
'        MsgBox "��ѯ��������", vbInformation, "��ʾ"
'        Exit Sub
'
'    End If

    'excel
   ExporToExcel (strSql)

End Sub

Private Sub openData()

If openFile Then
    If cbCustomerCode.text = "37" Then
    ShowData
 
    End If
       Toolbar1.Buttons(5).Enabled = True
End If

End Sub

Private Function openFile() As Boolean

On Error GoTo openFile_Err

openFile = False
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then Exit Function
txtFileName.text = Replace(CommonDialog1.filename, Chr(0), ",")
CommonDialog1.filename = ""
openFile = True
Exit Function
openFile_Err:
MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_uploadShippingList.openFile ", vbExclamation + vbOKOnly, "Application Error"

Resume Next

End Function

Private Sub ShowData()
Dim strFileName() As String
Dim i             As Integer

fps(0).MaxRows = 0
fps(1).MaxRows = 0

If InStr(txtFileName.text, ",") > 0 Then
    strFileName = Split(Trim$(txtFileName.text), ",")

    For i = 1 To UBound(strFileName)
        Call ShowFps(strFileName(0) & "\" & strFileName(i))
    Next
Else
    ReDim strFileName(0)
    strFileName(0) = Trim$(txtFileName.text)
    Call ShowFps(strFileName(0))

End If

End Sub

Private Sub ShowFps(strFileName As String)

    On Error GoTo showFps_ErrON

    Dim VBExcel As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i       As Integer
    Dim j       As Integer
    Dim strDN   As String
    Dim strJob  As String
    Dim strShipDate As String
    Dim lQty    As Long
    Dim strChar As String
    
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(strFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 58 Then
        MsgBox "DN��������", vbInformation, "����"
        GoTo showFps_Err
                        
    End If

    With fps(0)

        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count

            If i <> 1 Then .MaxRows = .MaxRows + 1

            For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

                If j > 26 Then
                    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                Else
                    strChar = Chr(96 + j)

                End If

                If i = 1 Then
                    If fps(0).MaxRows = 0 Then
                        .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " "), "'", " "))

                    End If

                Else
                    .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " "), "'", " "))

                End If

                If i > 1 Then
                    If j = 1 Then
                        strDN = Trim$(xlSheet.Range(strChar & i))

                        If strDN = "" Then
                            MsgBox "DN������Ϊ��", vbCritical, "����"
                            GoTo showFps_Err
                        
                        End If
                    
                    ElseIf j = 10 Then
                        strShipDate = Trim$(xlSheet.Range(strChar & i))
                    
                    ElseIf j = 32 Then
                        strJob = Trim$(xlSheet.Range(strChar & i))
                    
                    ElseIf j = 33 Then
                        lQty = CLng(Trim$(xlSheet.Range(strChar & i)))

                    End If

                End If

            Next j

            If i > 1 Then
                If isNewDN(strDN) = True Then
                    If addDNQty(strDN, strJob, lQty, strShipDate) = False Then
                        GoTo showFps_Err

                    End If

                Else

                    If upDateDNQty(strDN, strJob, lQty, strShipDate) = False Then
                        GoTo showFps_Err

                    End If

                End If

            End If

        Next i

    End With

    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
showFps_Err:
    fps(0).MaxRows = 0
    fps(1).MaxRows = 0

    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
showFps_ErrON:
    GoTo showFps_Err
    MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_uploadShippingList.showFps", vbExclamation + vbOKOnly, "Application Error"

End Sub

Private Function isNewDN(strDN As String) As Boolean
Dim i As Integer

isNewDN = True

With fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text = strDN Then
            isNewDN = False

        End If

    Next

End With

End Function

Private Function checkDNHistory(strDN As String) As Boolean
Dim strSql As String

checkDNHistory = False
strSql = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"

If Get_OracleCnt(strSql) > 0 Then
    MsgBox "֮ǰ�Ѿ��ϴ���DN: " & strDN & vbCrLf & "���ν�ֹ�ϴ�, ����ѡ����ȷ��DN�ļ�", vbExclamation, "����"

    Exit Function
End If

checkDNHistory = True

End Function

Private Function upDateDNQty(strDN As String, strJob As String, lQty As Long, strShipDate As String) As Boolean
Dim i       As Integer
Dim strBand As String

upDateDNQty = False
strBand = getBandFlag(strJob)

If strBand = "" Then

    'MsgBox "JOBID: " & strJob & "���ϴ�WOʱû��ά����˰���. ����ϵIT", vbInformation, "��ʾ"
    ' Exit Function
End If

With fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text = strDN Then
            .Col = 2
            .text = CLng(.text) + lQty
            
            .Col = 3
            If .text <> strBand Then
                MsgBox "DN:" & strDN & "���ڱ�˰�Ǳ�˰���JOB, ��ȷ��, ���������ϴ�", vbCritical, "����"
                Exit Function

            End If
            
            .Col = 4
            If .text <> strShipDate Then
                MsgBox "DN:" & strDN & "���ڶ�ʳ�������, ��ȷ��, ���������ϴ�", vbCritical, "����"
                Exit Function

            End If
            
        End If
        
    Next
    
End With

upDateDNQty = True

End Function

Private Function getBandFlag(strJob As String) As String
Dim strBand    As String
Dim strSql     As String
Dim strWaferID As String
Dim strWOID    As String

strJob = Replace$(strJob, "M", "")
'strSql = "select replace(aa.substrateid, '+','') as waferid from mappingdatatest aa  inner join customeroitbl_test bb on to_char(bb.id) = aa.filename and aa.lotid = bb.source_batch_id where bb.test_mtrl_desc = '" & strJob & "'  "
'strWaferID = Get_OracleStr(strSql)
'strSql = "select substratetype from mappingdatatest where substrateid = '" & strWaferID & "' and substratetype is not null "
'strBand = Get_OracleStr(strSql)
'
'If strBand = "" Then
'    strSql = "select distinct jobno from customeroitbl_test where test_mtrl_desc in (select source_mtrl_SLOC from customeroitbl_test where test_mtrl_desc = '" & strJob & "')"
'    If Get_OracleStr(strSql) = "A" Then
'        getBandFlag = "��˰"
'    ElseIf Get_OracleStr(strSql) = "B" Then
'        getBandFlag = "�Ǳ�˰"
'    Else
'        getBandFlag = ""
'    End If
'Else
'
'    If strBand = "0" Then
'        getBandFlag = "�Ǳ�˰"
'    Else
'        getBandFlag = "��˰"
'
'    End If
'
'End If

strSql = "select distinct cc.ordername from customeroitbl_test aa " & _
" inner join mappingdatatest bb on to_char(aa.id) = bb.filename and aa.source_batch_id = bb.lotid " & _
" inner join ib_waferlist cc on bb.substrateid = cc.waferid and bb.lotid = cc.waferlot " & _
" where aa.test_mtrl_desc = '" & strJob & "' "

strWOID = Trim("" & Get_OracleStr(strSql))

If Left$(strWOID, 1) = "A" Then
    getBandFlag = "��˰"

ElseIf Left$(strWOID, 1) = "B" Then
    getBandFlag = "�Ǳ�˰"

Else
    'MsgBox "��ѯ������˰�Ǳ�˰����", vbCritical, "����"
getBandFlag = "��˰"
End If



End Function

Private Function addDNQty(strDN As String, strJob As String, lQty As Long, strShipDate As String) As Boolean
Dim i       As Integer
Dim strBand As String

addDNQty = False
strBand = getBandFlag(strJob)

If strBand = "" Then

    'MsgBox "JOBID: " & strJob & "���ϴ�WOʱû��ά����˰���. ����ϵIT", vbInformation, "��ʾ"
    ' Exit Function
End If

With fps(1)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText 1, i, strDN
    .SetText 2, i, lQty
    .SetText 3, i, strBand
    .SetText 4, i, strShipDate

End With

addDNQty = True

End Function

Private Sub SaveData()
Dim i          As Integer
Dim strDN      As String
Dim strDNQuery As String

If CheckDC = False Then
    MsgBox "DateCode���δͨ��, �����޸�AZ�е�DateCode,�������ϴ�����", vbInformation, "��ʾ"
    Exit Sub
End If

strDNQuery = ""

If cbCustomerCode.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

If fps(0).MaxRows = 0 Then
    MsgBox "�����Ҫ�ϴ����ļ�", vbInformation, "��ʾ"
    Exit Sub

End If

Select Case cbCustomerCode.ListIndex

    Case 0  '37

        With fps(1)

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                strDN = Trim$(.text)

                If checkDNHistory(strDN) = False Then
                    fps(0).MaxRows = 0
                    fps(1).MaxRows = 0
                    Exit Sub
                Else

                    If getData(strDN) = True Then
                        strDNQuery = strDNQuery & strDN & "','"
                        MsgBox "DN:" & strDN & "�ϴ��ɹ�", vbInformation, "��ʾ"

                    End If

                End If

            Next

        End With

        If strDNQuery <> "" Then
            strDNQuery = Left$(strDNQuery, Len(strDNQuery) - 3)
            Call exportSuccessDN(strDNQuery)

        End If

End Select

Toolbar1.Buttons(5).Enabled = False

End Sub

Private Sub exportSuccessDN(strDN As String)
Dim strSql As String

strSql = "select a.lastupdatedate as ��������,a.delivery as DN, sum(a.quantity) as �ϴ����� from CUSTOMERSHIPPINGUPTBL a where a.delivery in ('" & strDN & "') group by a.lastupdatedate,a.delivery order by a.lastupdatedate desc "
ExporToExcel (strSql)

End Sub

Private Function getData(strDN As String) As Boolean

On Error GoTo getData_ErrON

Dim tyDN As DN_DETAILS
Dim i    As Integer
Dim j    As Integer

getData = False
Cnn.BeginTrans
INIadoCon.BeginTrans

With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        
        If .text = "" Then
            MsgBox "���ݳ���", vbCritical, "����"
            GoTo getData_Err
        End If
        
        If .text = strDN Then
            .Col = 1
            tyDN.Delivery = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 2
            tyDN.ItemNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 3
            tyDN.DeliveryCreationDate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 4
            tyDN.Plant = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 5
            tyDN.SalesDocument = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 6
            tyDN.SOItemNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 7
            tyDN.Material = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 8
            tyDN.MarketingPN = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 9
            tyDN.MaterialDescription = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 10
            tyDN.PlannedGIdate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 11
            tyDN.CustomerPartnumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 12
            tyDN.ShiptoName = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 13
            tyDN.ShiptoCustomer = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 14
            tyDN.PurchasingDocNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 15
            tyDN.DateCodeRestrictions = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 16
            tyDN.LabelRequirement = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 17
            tyDN.ReLabelInstructions = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 18
            tyDN.ShipToStreet1 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 19
            tyDN.ShipToStreet2 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 20
            tyDN.ShipToStreet3 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 21
            tyDN.City = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 22
            tyDN.State = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 23
            tyDN.PostalCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 24
            tyDN.CountryKey = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 25
            tyDN.ContactName = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 26
            tyDN.Phone = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 27
            tyDN.Fax = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 28
            tyDN.FreightForwarder = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 29
            tyDN.ShippingInstruction = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 30
            tyDN.AdditionalComments = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 31
            tyDN.StorageLocation = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 32
            tyDN.BatchNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 33
            tyDN.Quantity = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 34
            tyDN.VolumeWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 35
            tyDN.GrossWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 36
            tyDN.netweight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 37
            tyDN.UoMForWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 38
            tyDN.NoOfCartons = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 39
            tyDN.VendorLotNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 40
            tyDN.ShelfLocation = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 41
            tyDN.BOLOrAirwayBillNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 42
            tyDN.ActualShippingDate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 43
            tyDN.PackagingDetails = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 44
            tyDN.PackingStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 45
            tyDN.PickingStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 46
            tyDN.CustomerCalendar = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 47
            tyDN.FatherBatch = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 48
            tyDN.MotherBatch = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 49
            tyDN.FatherBatchDateCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 50
            tyDN.MotherBatchDateCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 51
            tyDN.TransferOrderStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 52
            tyDN.DATECODE = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 53
            tyDN.FatherBatchQty = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 54
            tyDN.ShippingPoint = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 55
            tyDN.ShipmentNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 56
            tyDN.FabSite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 57
            tyDN.AssemblySite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 58
            tyDN.TestSite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            tyDN.id = GetshippingMaxID()

            If checkDNData(tyDN) = False Or saveDataToDB(tyDN) = False Then
                GoTo getData_Err

            End If

        End If

    Next

End With

Cnn.CommitTrans
INIadoCon.CommitTrans
getData = True
Exit Function

getData_Err:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
Exit Function

getData_ErrON:
GoTo getData_Err
MsgBox Err.DESCRIPTION & ":�������ݿ����", vbExclamation, "����"

End Function

Private Function checkDNData(tyDN As DN_DETAILS) As Boolean
checkDNData = False
Dim strSql As String
Dim strDC As String

If tyDN.DATECODE = "" Then
    MsgBox "DATECODE������Ϊ��,����ϵ�ͻ�,ȷ����ֵ", vbCritical, "����"
    Exit Function
Else
    If Right(tyDN.BatchNumber, 1) <> "M" And Right(tyDN.BatchNumber, 1) <> "R" Then
        
        strSql = "select dc from tbl37testdc where jobid = '" & tyDN.BatchNumber & "'   "
        
        strDC = Get_OracleStr(strSql)
        If strDC <> "" And strDC <> tyDN.DATECODE Then
           tyDN.DATECODE = strDC
           
            MsgBox "�ͻ�DC:" & tyDN.DATECODE & " �ͳ���DC" & strDC & "��һ��,����ϵIT,ȷ��һ��", vbCritical, "����"
            Exit Function

        End If

    End If

End If

If tyDN.FabSite <> "" Or tyDN.AssemblySite <> "" Or tyDN.TestSite <> "" Then
    If InStr(UCase(tyDN.LabelRequirement), "SAMSUNG") = 0 Then
        MsgBox "�������SITE��ֵ,������37��SAMSUNG�ı�ǩ", vbCritical, "����"
        Exit Function

    End If

End If

If (InStr(UCase(tyDN.ShiptoName), "PREMIER") > 0) And (tyDN.CustomerPartnumber <> "") Then
    If InStr(UCase(tyDN.LabelRequirement), "SAMSUNG") = 0 Then
        MsgBox "label requirement������SAMSUNG��", vbCritical, "����"
        Exit Function

    End If

End If

If tyDN.LabelRequirement = "" Then
    MsgBox "label requirement��ǩ���Ͳ���Ϊ��", vbCritical, "����"
    Exit Function

End If

checkDNData = True

End Function

Private Function saveDataToDB(tyDN As DN_DETAILS) As Boolean
Dim strSql  As String
Dim strSql2 As String
Dim strSql3 As String

saveDataToDB = False
strSql = "select * from CUSTOMERSHIPPINGUPTBL where Delivery='" & tyDN.Delivery & "'  and batchnumber='" & tyDN.BatchNumber & "' and flag='Y'"

If Get_OracleCnt(strSql) > 0 Then
    strSql = "update CUSTOMERSHIPPINGUPTBL set quantity =(quantity + '" & tyDN.Quantity & "') where delivery = '" & tyDN.Delivery & "' and batchnumber = '" & tyDN.BatchNumber & "'"
'    strSql2 = "update [ERPBASE].[dbo].[tblCustomerShippingUp] set quantity =(quantity + '" & tyDN.Quantity & "')  where delivery = '" & tyDN.Delivery & "' and batchnumber = '" & tyDN.BatchNumber & "'"

    strSql2 = "insert into [ERPBASE].[dbo].[tblCustomerShippingUp](ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',GETDATE(), " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "

    strSql3 = "insert into DNSHIPMENTTBL(ID, Delivery, CUSTOMER_DEVICE,JOB_ID,HT_DEVICE,QUANTITY,SHIP_ORDER,BOND,SHIP_DATE,EXPRESS,REQUEST_FLAG,MAIL_FLAG) " & _
        " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.Material & "','" & tyDN.BatchNumber & "','','" & tyDN.Quantity & "','','" & getBandFlag(tyDN.BatchNumber) & "','" & tyDN.PlannedGIdate & "','',0,0)"
    
    If AddSql(strSql) = 0 Or AddSql2(strSql2) = 0 Or AddSql(strSql3) = 0 Then
        MsgBox "DN: " & tyDN.Delivery & "û���ϴ��ɹ�, ����ϵITȷ�������Ƿ�������", vbCritical, "����"
        Exit Function

    End If
    
    

Else
    strSql = "insert into CUSTOMERSHIPPINGUPTBL(ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',sysdate, " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "
    
    strSql2 = "insert into [ERPBASE].[dbo].[tblCustomerShippingUp](ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',GETDATE(), " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "
    
    strSql3 = "insert into DNSHIPMENTTBL(ID, Delivery, CUSTOMER_DEVICE,JOB_ID,HT_DEVICE,QUANTITY,SHIP_ORDER,BOND,SHIP_DATE,EXPRESS,REQUEST_FLAG,MAIL_FLAG) " & _
        " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.Material & "','" & tyDN.BatchNumber & "','','" & tyDN.Quantity & "','','" & getBandFlag(tyDN.BatchNumber) & "','" & tyDN.PlannedGIdate & "','',0,0)"
    
    If AddSql(strSql) = 0 Or AddSql2(strSql2) = 0 Or AddSql(strSql3) = 0 Then
        MsgBox "DN: " & tyDN.Delivery & "û���ϴ��ɹ�, ����ϵITȷ�������Ƿ�������", vbCritical, "����"
        Exit Function

    End If

End If

saveDataToDB = True

End Function

Private Sub delData()
Dim strDN As String

If txtDN.text = "" Then
    MsgBox "������Ҫɾ����DN", vbInformation, "��ʾ"
    Exit Sub

End If

strDN = Trim$(txtDN.text)

If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'") = 0 Then
    MsgBox "�������DN����ȷ��û���ϴ���¼,����ɾ��", vbInformation, "��ʾ"
    Exit Sub

End If
If Get_OracleCnt("  select * from packing_detailed i WHERE i.DN_NUM= '" & strDN & "'") > 0 Then
    MsgBox "����ѿ�ʼ��ҵ����DN����ɾ��,��֪ͨ���", vbInformation, "��ʾ"
    Exit Sub
End If
    
  
AddSql ("insert into CUSTOMERSHIPPINGUPTBL_BAK select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' ")
MsgBox "DN���ݳɹ�", vbInformation, "��ʾ"
AddSql ("delete from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'")
AddSql ("delete from DNSHIPMENTTBL where delivery = '" & strDN & "'")
AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp] where Delivery = '" & strDN & "'")
MsgBox "�ѳɹ�ɾ��DN:" & strDN, vbInformation, "��ʾ"
txtDN.text = ""

End Sub

Private Sub exitFrm()
Unload Me

End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

    Case "QUERY_MOD"
        QueryData_Mod

    Case "SAVE_MOD"
       'ȷ���޸�
        SAVE_Modify
     
    Case "DEL_MOD"
          DEL_Single

    Case "HOME_MOD"
        exitFrm

End Select

End Sub
Private Sub DEL_Single()

Dim strid  As String
Dim strqty  As String
Dim i As Integer
Dim DelCnt As Integer
DelCnt = 0
With fpS_Mod(0)
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
                DelCnt = DelCnt + 1
            End If

        End If
    Next i
    If DelCnt = 0 Then
        MsgBox "��ѡ��Ҫɾ��������", vbInformation, "��ʾ"
        Exit Sub
    End If
    If MsgBox("��ȷ��Ҫɾ��ѡ�е�" & DelCnt & "��������?", vbOKCancel, "��ʾ") = vbCancel Then
        Exit Sub
    End If
    If cbCustomerCode_Mod.text = "37" Then
         For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            If .text <> "" Then
                If .text = 1 Then
        
                    .Col = 6 'ID
                    strid = Trim(.text) '
                   
                    AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'ɾ��', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                    AddSql ("delete from  CUSTOMERSHIPPINGUPTBL where ID= '" & strid & "' ")
                    AddSql ("delete from  DNSHIPMENTTBL where ID= '" & strid & "' ")
                    AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp] where ID = '" & strid & "' ")
                End If
                
            End If
        Next
    ElseIf cbCustomerCode_Mod.text = "SG005" Then
         For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            If .text <> "" Then
                If .text = 1 Then
        
                    .Col = 2 'ID
                    strid = Trim(.text) '
                   
                    AddSql ("insert into CUSTOMERSHIPPING_SO_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'ɾ��', a.* from CUSTOMERSHIPPINGUPTBL_SO a where a.ID = '" & strid & "' ")
                    AddSql ("delete from  CUSTOMERSHIPPINGUPTBL_SO where ID= '" & strid & "' ")
                  '  AddSql ("delete from  DNSHIPMENTTBL where ID= '" & strid & "' ")
                    AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp_So] where ID = '" & strid & "' ")
                End If
                
            End If
        Next
    End If
    QueryData_Mod
    
End With

    
    

End Sub

Private Sub SAVE_Modify()

If CbItem_Mod.text = "�޸ĳ�������" Then
    Quality_Modify
ElseIf CbItem_Mod.text = "�޸ĳ�������" Then
    ShipDate_Modify
ElseIf CbItem_Mod.text = "ɾ��" Then
    MsgBox "ɾ�����ȷ��ɾ����ť", vbInformation, "��ʾ"
 
Else
    MsgBox "��ѡ���޸���Ŀ", vbInformation, "��ʾ"
End If

End Sub


Private Sub ShipDate_Modify()
'���������޸�
Dim strid  As String
Dim strqty  As String
Dim i As Integer
Dim rs     As New ADODB.Recordset
Dim strmsg As String
Dim ordernumlist As String
Dim strshipdate_old As String



If txtShipDate_Mod.text = "" Then

    MsgBox "��ѡ���µĳ�������", vbInformation, "��ʾ"
    Exit Sub

End If

strshipdate_old = GetSqlServerStr("select distinct ship_date from erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'")

With fpS_Mod(0)

     For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
    
                .Col = 6 'ID
                strid = Trim(.text) '
                
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'�޸�ǰ����', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql ("update  CUSTOMERSHIPPINGUPTBL set PLANNEDGIDATE='" & txtShipDate_Mod.text & "', LASTUPDATEDATE=to_date('" & txtShipDate_Mod.text & "','yyyy/mm/dd')   where ID= '" & strid & "' ")
                AddSql ("update  DNSHIPMENTTBL set SHIP_DATE='" & txtShipDate_Mod.text & "' where ID= '" & strid & "' ")
                AddSql2 ("update   [ERPBASE].[dbo].[tblCustomerShippingUp]  set PLANNEDGIDATE='" & txtShipDate_Mod.text & "', LASTUPDATEDATE='" & txtShipDate_Mod.text & "' where ID= '" & strid & "' ")
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'�޸ĺ�����', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql2 ("update  erpdata..tblShipOrder_Dn set ship_date='" & Format(dtShipDate_mod.Value, "yyyy/mm/dd") & "' where dn='" & txtDN_Mod.text & "'")
                      
            End If
            
        End If
    Next

    
    MsgBox "�޸����"
    QueryData_Mod
End With
 'merry20200422�����dn�Ѿ����ɳ������� , һ���޸ĳ������ݶ�Ӧ�ĳ�������

Set rs = Get_SqlserveRs("select distinct dn,shiporder,ship_date from  erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'")
If rs.RecordCount > 0 Then
    If Get_SqlserverCnt("select distinct ship_date from erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'") > 1 Then
        MsgBox "ͬһ��DN��Ӧ��ͬ����ʱ�䣬�ʼ�����ʧ�ܣ�", vbInformation, "��ʾ"
    Else
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            If ordernumlist = "" Then
                ordernumlist = rs("shiporder")
            Else
                ordernumlist = ordernumlist & "," & rs("shiporder")
            End If
            rs.MoveNext
        Next
        Cob(0).text = "37"
        dtShipDate_Ship.Value = strshipdate_old
        mailflag = 2 '�޸�ʱ��ĸ�ʽ
        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
        dtShipDate_Ship.Value = txtShipDate_Mod.text
        mailflag = 1 '���������ĸ�ʽ
        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
    End If
End If




End Sub



Private Sub Quality_Modify()
'���������޸�
Dim strid  As String
Dim i As Integer
Dim strqty As Long

With fpS_Mod(0)



     For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
    
                .Col = 6 'ID
                strid = Trim(.text) '
                
                .Col = 5 '����
                                        
                If IsNumeric(Trim(.text)) = False Then
                
                    MsgBox "��������" & strqty & "��д���������������ύ", vbInformation, "��ʾ"
                    Exit Sub
                End If
                strqty = Trim(.text)
                If strqty <= 0 Then
                
                    MsgBox "��������" & strqty & "��д���������������ύ", vbInformation, "��ʾ"
                    Exit Sub
                End If
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'�޸�ǰ����', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql ("update  CUSTOMERSHIPPINGUPTBL set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql ("update  DNSHIPMENTTBL set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql2 ("update  [ERPBASE].[dbo].[tblCustomerShippingUp]  set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'�޸ĺ�����', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                
            End If
            
        End If
    Next
    MsgBox "�޸����"
    QueryData_Mod
End With




End Sub


Private Sub QueryData_Mod()
Dim strDN  As String
Dim strSql As String
Dim rs     As New ADODB.Recordset

fpS_Mod(0).MaxRows = 0
fpS_Mod(1).MaxRows = 0

If cbCustomerCode_Mod.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

If txtDN_Mod.text = "" Then
    MsgBox "������DN��", vbInformation, "��ʾ"
    Exit Sub
End If
If CbItem_Mod.text = "" Then
    MsgBox "��ѡ���޸���Ŀ", vbInformation, "��ʾ"
    Exit Sub
End If


If cbCustomerCode_Mod.text = "37" Then

    strDN = Trim$(txtDN_Mod.text)
    strSql = "select 0 as ѡ��,a.delivery as DN, a.PLANNEDGIDATE as ��������,a.BATCHNUMBER AS JobID, a.QUANTITY as �������� ,a.* from CUSTOMERSHIPPINGUPTBL a where delivery = '" & strDN & "' order by a.id desc "
    Set rs = Get_OracleRs(strSql)
    With fpS_Mod(0)
        Set .DataSource = rs
        .Row = -1
        .Col = 1  'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
    End With
   
    
    If CbItem_Mod.text = "�޸ĳ�������" Or CbItem_Mod.text = "ɾ��" Then
         With fpS_Mod(0)
              .Row = -1
              .Col = -1
              .BackColor = &HFFFFFF
              .Col = 5
              .Lock = False
             .BackColor = vbGreen
             
             
             
             
         End With
    ElseIf CbItem_Mod.text = "�޸ĳ�������" Then
         With fpS_Mod(0)
              .Row = -1
              .Col = -1
              .BackColor = &HFFFFFF
              .Col = 3
             .BackColor = vbGreen
  
         End With
    End If
    If Get_OracleCnt("  select * from packing_detailed i WHERE i.DN_NUM= '" & strDN & "'") > 0 Then
        If CbItem_Mod.text = "��ѯ�����޸ļ�¼" Then
            Toolbar2.Buttons(3).Enabled = False
            Toolbar2.Buttons(5).Enabled = False
            
        ElseIf CbItem_Mod.text = "�޸ĳ�������" Then
            Toolbar2.Buttons(3).Enabled = True

        Else
            MsgBox "����ѿ�ʼ��ҵ����DN�����޸�,��֪ͨ���", vbInformation, "��ʾ"
            Toolbar2.Buttons(3).Enabled = False
        End If
        Toolbar2.Buttons(5).Enabled = False
    Else
         If CbItem_Mod.text = "��ѯ�����޸ļ�¼" Then
            Toolbar2.Buttons(3).Enabled = False
            Toolbar2.Buttons(5).Enabled = False
        Else
            Toolbar2.Buttons(3).Enabled = True
            Toolbar2.Buttons(5).Enabled = True
    End If
    End If
    If CbItem_Mod.text = "�޸ĳ�������" Then
        strSql = "select distinct a.delivery as DN, a.MODIFYBY as �޸���Ա,a.MODIFYDATE as �޸�ʱ��,a.MODFIYITEM  as �޸���Ŀ, a.LASTUPDATEDATE as ����ʱ��   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%����%' order by a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "�޸ĳ�������" Then
        strSql = "select distinct  a.delivery as DN, a.MODIFYBY as �޸���Ա,a.MODIFYDATE as �޸�ʱ��,a.MODFIYITEM  as �޸���Ŀ, a.LASTUPDATEDATE as ����ʱ��,a.BatchNumber as JOBID, a.Quantity AS ����   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%����%' order by a.BatchNumber,a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "ɾ��" Then
        strSql = "select distinct a.delivery as DN,  a.MODIFYBY as �޸���Ա,a.MODIFYDATE as �޸�ʱ��,a.MODFIYITEM  as �޸���Ŀ, a.LASTUPDATEDATE as ����ʱ��,a.BatchNumber as JOBID, a.Quantity AS ����   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%ɾ��%' order by a.BatchNumber,a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "��ѯ�����޸ļ�¼" Then
        strSql = "select distinct  a.delivery as DN,  a.MODIFYBY as �޸���Ա,a.MODIFYDATE as �޸�ʱ��,a.MODFIYITEM  as �޸���Ŀ, a.LASTUPDATEDATE as ����ʱ��,a.BatchNumber as JOBID, a.Quantity AS ����   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' order by a.BatchNumber,a.MODIFYDATE desc "
        
     End If
    Set rs = Get_OracleRs(strSql)
    With fpS_Mod(1)
        Set .DataSource = rs
    End With
ElseIf cbCustomerCode_Mod.text = "SG005" Then

    strDN = Trim$(txtDN_Mod.text)
    strSql = "select 0 as ѡ��,a.* from ERPBASE..tblCustomerShippingUp_So a where SO_NO = '" & strDN & "' order by a.SO_LINE "
    Set rs = Get_SqlserveRs(strSql)
    With fpS_Mod(0)
        Set .DataSource = rs
        .Row = -1
        .Col = 1  'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
    End With
   

End If

End Sub



'------------------------------------------------------
'�˺�Ϊ�´���

Private Sub Reconciliation(sDH)
' �����ά��
Dim sFlag As String
Dim strSql As String
Dim rs          As New ADODB.Recordset
sFlag = "N"
If chkDZD.Value = 1 Then
    sFlag = "Y"
Else
    sFlag = "N"
End If

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = INIadoCon

rs.Source = "select * from [erpdata].[dbo].[MDZD_TBL] where SENT_ID = '" & sDH & "'"

rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    strSql = "update [erpdata].[dbo].[MDZD_TBL] set flag = '" & sFlag & "' where SENT_ID = '" & sDH & "'"
    If AddSql2(strSql) = 0 Then
        MsgBox "û��ά��������˵�", vbCritical, "��ʾ"
        Exit Sub
    End If
Else
    strSql = "insert into  [erpdata].[dbo].[MDZD_TBL](SENT_ID, FLAG) values('" & sDH & "', '" & sFlag & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "û��ά��������˵�", vbCritical, "��ʾ"
        Exit Sub
    End If
End If

End Sub


'��������
Private Sub SaveData_Ship()
Dim i           As Long
Dim j           As Long
Dim k           As Long
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strCust     As String   '�ͻ�����
Dim strKF       As String   '�ⷿ
Dim strid       As String   'ID�ַ���
Dim strXH       As String   '����ַ���
Dim strgdh      As String   '�����ַ���
Dim strLCK      As String   '���̿��ַ���
Dim strlps      As String   '�ϸ���
Dim strbls      As String   '������
Dim strzcbls    As String   '�Ƴ̲�����
Dim strDN    As String   'dn
Dim intCount    As Integer  '��¼����
Dim sDH As String
Dim sFlag As String
Dim rs1 As ADODB.Recordset
Dim flag_up As String
Dim qty_total As Long
Dim dn_temp  As String 'dn
Dim bigbox_temp  As String '�����
Dim XH_temp  As String
Dim KF_temp  As String
Dim ordernumlist  As String
Dim strKFList  As String

    mailflag = 1 '�����������ʼ���ʽ
    intFlag = 1
    intCount = 0
    strKF = ""
    strCust = ""
    strKFList = ""
    '���ϼ��
    If txt(1).text = "" Or txt(2).text = "" Or txt(3).text = "" Then
        MsgBox "�������ϵ������ԣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    If intFlag <> 3 Then
        If Cob(6).text = "" Then
            MsgBox "��ѡ�����䷽ʽ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If Cob(7).text = "" Then
            MsgBox "��ѡ���ջ��ͻ���", vbInformation, "��ʾ"
            Exit Sub
        End If
        'If txt(4).Text = "" Then
        If Cob(9).text = "" Then
            MsgBox "��ѡ���ջ���ַ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If Val(txt(6).text) + Val(txt(7).text) + Val(txt(8).text) <= 0 Then
            MsgBox "����ѡ��Ҫ��������ţ�", vbInformation, "��ʾ"
            Exit Sub
        End If
        '��¼�����ⷿ��У������Ŀⷿ�ǲ���һ����
        intCount = 0
        strKF = ""
        With Fps_Ship(0)
            If .MaxRows <= 0 Then
                MsgBox "���Ȳ�ѯ���ϣ�", vbInformation, "��ʾ"
                Exit Sub
            End If
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.E_CHOOSE 'ѡ��
                If .Value = 1 Then
                    intCount = intCount + 1 '������
                    .Col = FpsM.E_cust '�ͻ�
                    If strCust = "" Then
                        strCust = Trim$(.text)
                    End If
                    '.Col = FpsM.e_KF '�ⷿ
                    'If strKF = "" Then
                    '    strKF = Trim$(.Text)
                    'End If
                    'У����һ���Ƿ�һ������һ������ʾ����
                    'If strKF <> Trim(.Text) Then
                    '    MsgBox "����ֻ�ܳ�ͬһ���ⷿ�����ϣ�", vbInformation, "��ʾ"
                    '    Exit Sub
                    'End If
                    .Col = FpsM.e_KF '�ⷿ
                    If InStr(strKFList, Split(Trim$(.text), " ")(0)) = 0 Then
                        If strKFList = "" Then
                            strKFList = Split(Trim$(.text), " ")(0)
                        Else
                            strKFList = strKFList & "," & Split(Trim$(.text), " ")(0)
                        End If
                    End If
                    
                End If
            Next
        End With
        If intCount <= 0 Then
            MsgBox "����ѡ��Ҫ��������ţ�", vbInformation, "��ʾ"
            Exit Sub
        End If
       ' strKF = Left(strKF, InStr(strKF, " ") - 1) '��ȡ�ⷿ���
        '��¼��ϸ����
        With Fps_Ship(1)
            If .MaxRows <= 0 Then
                MsgBox "����ѡ��Ҫ������������ϣ�", vbInformation, "��ʾ"
                Exit Sub
            End If
'����

           If Cob(0).text = "37" Then
                
               dn_temp = ""
               For i = 1 To .MaxRows
                    .Row = i
                    .Col = FpsD.E_DN
                   If .text <> dn_temp Then '�л�dnʱ�л����뵥��
                        dn_temp = Trim(.text)
                        .Col = FpsD.E_XH
                        strXH = Trim(.text)
                        .Col = FpsD.e_BigX
                        bigbox_temp = Trim(.text)
                        .Col = FpsD.e_KF
                        KF_temp = Trim(.text)
                        .SetText FpsD.E_Note2, i, "�л����뵥��"
                   Else
                        .Col = FpsD.e_KF
                        If KF_temp <> Trim(.text) Then
                            MsgBox dn_temp & "���ڲ�ͬ�Ŀⷿ���޷�ͬһ�γ�������ȷ�ϣ�", vbInformation, "��ʾ"
                            Exit Sub
                        End If
                   
                       .Col = FpsD.e_BigX
                       If Trim(.text) <> bigbox_temp Then 'ͬdn��ͬ����ţ��п�����Ҫ������뵥��
                          bigbox_temp = Trim(.text)
                          .Col = FpsD.E_XH    '���
                          XH_temp = Trim(.text)
                          For j = i + 1 To .MaxRows
                             .Row = j
                             .Col = FpsD.e_BigX
                             If Trim(.text) <> bigbox_temp Then Exit For
                             .Col = FpsD.E_XH    '���
                             XH_temp = XH_temp & Trim(.text) & "��"
                          Next
                          
                            If Len(strXH) + Len(XH_temp) < 6000 Then
                                .Col = FpsD.E_XH    '���
                                strXH = strXH & Trim(.text) & "��"
                            Else   '�л����뵥��
                                .Col = FpsD.E_XH    '���
                                strXH = Trim(.text)
                                .SetText FpsD.E_Note2, i, "�л����뵥��"
                            End If
                        Else
                            .Col = FpsD.E_XH    '���
                            strXH = strXH & Trim(.text) & "��"
                        End If
                   End If
               Next
                If CheckQtyByDn <> "" Then '�˶�DN����
                   ' MsgBox CheckQtyByDn & "�������治һ��", vbInformation, "��ʾ"
                    Exit Sub
                End If
            ElseIf UCase(Cob(0).text) = "SG005" Then
                If TxtShipDate_Ship.text = "" Then
                    MsgBox "��ѡ���������,�������Ҳ��Ҫ��ѡһ��", vbInformation, "��ʾ"
                    Exit Sub
                End If
                If CheckQtyBySO <> "" Then '�˶���ѡ�����Ƿ񳬳�SO����
                    Exit Sub
                End If
            End If

           If MsgBox("ȷ��Ҫִ���𣬵�'��'��������'��'ȡ����", vbInformation + vbYesNo, "��ʾ") = vbNo Then
                Exit Sub
           End If
        If UCase(Cob(0).text) = "37" Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsD.E_Note2
                If Trim(.text) <> "" Then
                    If i > 1 Then
                        AutoCode
                        Call Reconciliation(txt(0).text)   '�����ά��
                        If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                            Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                            MsgBox "����ִ��ʧ�ܣ�", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        If ordernumlist = "" Then
                            ordernumlist = txt(0).text
                        Else
                            ordernumlist = ordernumlist & "," & txt(0).text
                        End If
                    End If
                    strid = ""
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                End If
                .Row = i
                .Col = FpsD.e_ID    'ID
                If InStr(strid, Trim$(.text)) <= 0 Then '�����ھ��ۼ�
                    strid = strid & Trim$(.text) & "��"
                End If
                .Col = FpsD.E_XH    '���
                strXH = strXH & Trim(.text) & "��"
                .Col = FpsD.e_GDH   '������
                strgdh = strgdh & Trim(.text) & "��"
                .Col = FpsD.e_LCK   '���̿�
                strLCK = strLCK & Trim(.text) & "��"
                .Col = FpsD.e_GNum  '�ϸ���
                strlps = strlps & Trim(.text) & "��"
                .Col = FpsD.e_BLNum '���ϲ�����
                strbls = strbls & Trim(.text) & "��"
                .Col = FpsD.e_ZCNum '�Ƴ̲�����
                strzcbls = strzcbls & Trim(.text) & "��"
                .Col = FpsD.e_KF '�ⷿ
                strKF = Trim(.text)
                .Col = FpsD.E_DN 'dn
                strDN = Trim(.text)
                If i = .MaxRows Then
                    AutoCode
                    Call Reconciliation(txt(0).text)   '�����ά��
                    If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                        MsgBox "����ִ��ʧ�ܣ�", vbInformation, Me.Caption
                        'Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
                        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                        Exit Sub
                    End If
                    If ordernumlist = "" Then
                        ordernumlist = txt(0).text
                    Else
                        ordernumlist = ordernumlist & "," & txt(0).text
                    End If
                End If
                
            Next
        ElseIf Cob(0).text = "SG005" Then
            For k = 0 To UBound(Split(strKFList, ","))
                strKF = Split(strKFList, ",")(k)
                strid = ""
                strXH = ""
                strgdh = ""
                strLCK = ""
                strlps = ""
                strbls = ""
                strzcbls = ""
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = FpsD.e_KF
                    If Trim(.text) = strKF Then

                        .Row = i
                        .Col = FpsD.e_ID    'ID
                        If InStr(strid, Trim$(.text)) <= 0 Then '�����ھ��ۼ�
                            strid = strid & Trim$(.text) & "��"
                        End If
                        .Col = FpsD.E_XH    '���
                        strXH = strXH & Trim(.text) & "��"
                        .Col = FpsD.e_GDH   '������
                        strgdh = strgdh & Trim(.text) & "��"
                        .Col = FpsD.e_LCK   '���̿�
                        strLCK = strLCK & Trim(.text) & "��"
                        .Col = FpsD.e_GNum  '�ϸ���
                        strlps = strlps & Trim(.text) & "��"
                        .Col = FpsD.e_BLNum '���ϲ�����
                        strbls = strbls & Trim(.text) & "��"
                        .Col = FpsD.e_ZCNum '�Ƴ̲�����
                        strzcbls = strzcbls & Trim(.text) & "��"
                        .Col = FpsD.e_KF '�ⷿ
                        strKF = Trim(.text)
                        .Col = FpsD.E_DN 'dn
                        strDN = Trim(.text)
                    End If
                Next
                AutoCode
                Call Reconciliation(txt(0).text)   '�����ά��
                If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                    MsgBox "����ִ��ʧ�ܣ�", vbInformation, Me.Caption
                   ' Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
                    Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                    Exit Sub
                End If
                If ordernumlist = "" Then
                    ordernumlist = txt(0).text
                Else
                    ordernumlist = ordernumlist & "," & txt(0).text
                End If
            Next
        End If
        End With
    End If
    'Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
    Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
    If Cob(0).text = "37" Then
        LoadDn 'ˢ��DNlist
    ElseIf Cob(0).text = "SG005" Then
        LoadSO 'ˢ��SOlist
    End If
 End Sub
 
 
Private Function ExecProc(intFlag As Integer, strOrderNo As String, strid As String, strXH As String, strgdh As String, strLCK As String, strlps As String, strbls As String, strzcbls As String, strKF As String, strDN As String, strCust As String)
    Dim strSql As String
    Dim rs                  As New ADODB.Recordset
    Dim strShipDate As String
    Dim strTerm As String
    Dim strForwarder As String
    Dim strInstruction As String
    Dim strShipTo As String

On Error GoTo FError
    ExecProc = True
    If Len(strXH) >= 6000 Or Len(strLCK) >= 6000 Then
        MsgBox "ѡȡ�ĳ��������������ޣ������ѡȡ������", vbInformation, "��ʾ"
        ExecProc = False
        Exit Function
    End If

    '��������
    Set adoCmd = New ADODB.Command
     Set adoCmd.ActiveConnection = INIadoCon2
     adoCmd.CommandText = "uspfp_fhsqXH1"
     adoCmd.Parameters.Refresh
     adoCmd.CommandType = adCmdStoredProc
  
        Set adoPrmReturn = New ADODB.Parameter         '����ִ�гɹ����
        adoPrmReturn.type = adInteger
        adoPrmReturn.Direction = adParamReturnValue
        adoCmd.Parameters.Append adoPrmReturn
        
        Set adoprm1 = New ADODB.Parameter               '���ݱ��
        adoprm1.type = adChar
        adoprm1.Size = 50
        adoprm1.Direction = adParamInput
        adoprm1.Value = Trim(strOrderNo)
        adoCmd.Parameters.Append adoprm1
        
        Set adoprm2 = New ADODB.Parameter                 '1ID��
        adoprm2.type = adChar
        adoprm2.Size = 8000
        adoprm2.Direction = adParamInput
        adoprm2.Value = Trim(strid)
        adoCmd.Parameters.Append adoprm2
        
        Set adoPrm3 = New ADODB.Parameter                 '2���
        adoPrm3.type = adChar
        adoPrm3.Size = 80000
        adoPrm3.Direction = adParamInput
        adoPrm3.Value = Trim(strXH)
        adoCmd.Parameters.Append adoPrm3
        
        Set adoPrm4 = New ADODB.Parameter                 '5lck
        adoPrm4.type = adChar
        adoPrm4.Size = 80000
        adoPrm4.Direction = adParamInput
        adoPrm4.Value = Trim(strLCK)
        adoCmd.Parameters.Append adoPrm4
        
        Set adoPrm5 = New ADODB.Parameter                 '6gdh
        adoPrm5.type = adChar
        adoPrm5.Size = 80000
        adoPrm5.Direction = adParamInput
        adoPrm5.Value = Trim(strgdh)
        adoCmd.Parameters.Append adoPrm5
        
        Set adoPrm6 = New ADODB.Parameter                 '7lps
        adoPrm6.type = adChar
        adoPrm6.Size = 80000
        adoPrm6.Direction = adParamInput
        adoPrm6.Value = Trim(strlps)
        adoCmd.Parameters.Append adoPrm6
        
        Set adoPrm7 = New ADODB.Parameter               '8bls
        adoPrm7.type = adChar
        adoPrm7.Size = 80000
        adoPrm7.Direction = adParamInput
        adoPrm7.Value = Trim(strbls)                 '
        adoCmd.Parameters.Append adoPrm7
        
        Set adoPrm8 = New ADODB.Parameter               '9zcbls
        adoPrm8.type = adChar
        adoPrm8.Size = 80000
        adoPrm8.Direction = adParamInput
        adoPrm8.Value = Trim(strzcbls)
        adoCmd.Parameters.Append adoPrm8
        
        
        Set adoPrm9 = New ADODB.Parameter               '������
        adoPrm9.type = adChar
        adoPrm9.Size = 50
        adoPrm9.Direction = adParamInput
        adoPrm9.Value = Trim(txt(1).text)
        adoCmd.Parameters.Append adoPrm9
        
        Set adoprm10 = New ADODB.Parameter               '���벿��
        adoprm10.type = adChar
        adoprm10.Size = 50
        adoprm10.Direction = adParamInput
        adoprm10.Value = Trim(txt(3).text)
        adoCmd.Parameters.Append adoprm10
        
        Set adoPrm11 = New ADODB.Parameter               '�ͻ�����
        adoPrm11.type = adChar
        adoPrm11.Size = 50
        adoPrm11.Direction = adParamInput
        adoPrm11.Value = Trim(strCust)
        adoCmd.Parameters.Append adoPrm11
        
        Set adoPrm12 = New ADODB.Parameter             '������ַ
        adoPrm12.type = adChar
        adoPrm12.Size = 255
        adoPrm12.Direction = adParamInput
        'adoPrm12.Value = Trim(txt(4).Text)
        adoPrm12.Value = Trim(Cob(9).text)
        adoCmd.Parameters.Append adoPrm12
        
        Set adoPrm13 = New ADODB.Parameter               '���䷽ʽ
        adoPrm13.type = adChar
        adoPrm13.Size = 50
        adoPrm13.Direction = adParamInput
        adoPrm13.Value = Trim(Cob(6).text)
        adoCmd.Parameters.Append adoPrm13
        
        Set adoPrm14 = New ADODB.Parameter             '�ⷿ���
        adoPrm14.type = adChar
        adoPrm14.Size = 50
        adoPrm14.Direction = adParamInput
        adoPrm14.Value = strKF
        adoCmd.Parameters.Append adoPrm14
        
        Set adoPrm15 = New ADODB.Parameter             '˵��,��ע
        adoPrm15.type = adChar
        adoPrm15.Size = 50
        adoPrm15.Direction = adParamInput
        adoPrm15.Value = Trim(txt(5).text)
        adoCmd.Parameters.Append adoPrm15
        
        Set adoprmFG = New ADODB.Parameter             '1������;3��ɾ��;2���޸�
        adoprmFG.type = adInteger
        adoprmFG.Direction = adParamInput
        adoprmFG.Value = intFlag
        adoCmd.Parameters.Append adoprmFG
        
        Set adoprm16 = New ADODB.Parameter             '�ջ��ͻ�
        adoprm16.type = adChar
        adoprm16.Size = 50
        adoprm16.Direction = adParamInput
        adoprm16.Value = Trim(Cob(7).text)
        adoCmd.Parameters.Append adoprm16
      
     adoCmd.Execute
     Screen.MousePointer = 0
     
     If adoPrmReturn.Value = 0 Then
        
         strSql = "SELECT * FROM erptemp..SHIP_ORDER_sql a WHERE a.ship_order='" & Trim(strOrderNo) & "'"
         If rs.State = adStateOpen Then rs.Close
         rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
         If rs.RecordCount < 1 Then
          
        '   MsgBox "����ʧ�ܣ�", vbInformation, Me.Caption
           ExecProc = False
           Exit Function
          
         End If
        If Cob(0).text = "37" And strDN <> "" Then
            If DateDiff("d", Format(dtShipDate_Ship.Value, "yyyy/mm/dd"), Format(Now(), "yyyy/mm/dd")) >= 1 Then
            'Ӧ��ǰһ����Ļ�,����û�����ڶ�������������ڴ��ڶ��������
                strShipDate = Format(Now(), "yyyy/mm/dd")
            Else
                strShipDate = Format(dtShipDate_Ship.Value, "yyyy/mm/dd")
            End If
            strInstruction = GetSqlServerStr(" select isnull(ShippingInstruction,'') FROM ERPBASE..tblCustomerShippingUp WHERE rtrim(delivery)='" & Trim(strDN) & "'")
            If InStr(UCase(strInstruction), "FED") > 0 Or InStr(UCase(strInstruction), "DHL") > 0 Then
                strInstruction = "FED"
            Else
                strInstruction = "����"
            End If
            strShipTo = GetSqlServerStr(" select isnull(ShipToName,'') FROM ERPBASE..tblCustomerShippingUp WHERE rtrim(delivery)='" & Trim(strDN) & "'")
            'Shenzhen YH Global Logistics Co. Ltd=HW
            If Trim(UCase(strShipTo)) = "SHENZHEN YH GLOBAL LOGISTICS CO. LTD" Then
                strShipTo = "HW"
            Else
                strShipTo = ""
            End If
            
            strSql = " INSERT INTO erpdata..tblShipOrder_Dn(Dn, ShipOrder, Quality,BOND,CUST_DEVICE,HT_DEVICE,SHIP_DATE,Remark1,Remark2) " & _
            " SELECT '" & Trim(strDN) & "', a.���ݱ��,sum(a.����) AS ���� ,LEFT(a.�󹤵�,1) AS ��˰,c.MPN_DESC ,d.QTECHPTNO ,'" & strShipDate & "','" & strInstruction & "','" & strShipTo & "'" & _
            " FROM erpdata..tblstocksqfhSUB a , erpbase..tblmappingdata   b,erpbase..tblcustomeroi c,erptemp..tbltsvnpiproduct d " & _
            " WHERE a.���̿����=b.SUBSTRATEID AND a.������=b.LOTID AND b.LOTID =c.SOURCE_BATCH_ID AND b.FILENAME =convert(VARCHAR(30),convert(int,c.id)) AND a.�Ϻ�=d.QTECHPTNO2  " & _
            " AND  a.���ݱ�� ='" & Trim(strOrderNo) & "'" & _
            " GROUP BY a.���ݱ��,LEFT(a.�󹤵�,1),c.MPN_DESC ,d.QTECHPTNO "
            AddSql2 (strSql)

        End If
    
        If Cob(0).text = "SG005" And strsono <> "" Then
            strShipDate = Format(dtShipDate_Ship.Value, "yyyy/mm/dd")
            
            strTerm = GetSqlServerStr(" select isnull(TERM,'') FROM ERPBASE..tblCustomerShippingUp_So  WHERE rtrim(SO_NO)='" & Trim(strsono) & "' AND SO_LINE='" & strsoline & "'")
            
            strForwarder = GetSqlServerStr(" select isnull(Forwarder,'') FROM ERPBASE..tblCustomerShippingUp_So  WHERE rtrim(SO_NO)='" & Trim(strsono) & "' AND SO_LINE='" & strsoline & "'")
            
            
            strSql = " INSERT INTO erpdata..tblShipOrder_Dn(SO_NO, SO_LINE, ShipOrder, Quality,BOND,CUST_DEVICE,HT_DEVICE,PCSNUM,SHIPTO,STOCKID,SHIP_DATE,REMARK1,REMARK2) " & _
             " SELECT '" & strsono & "','" & strsoline & "', a.���ݱ��,sum(a.����) AS ���� ,LEFT(a.�󹤵�,1) AS ��˰,c.MPN_DESC ,d.QTECHPTNO,COUNT(*) AS Ƭ��,'" & Trim(Cob(9).text) & "' ,'" & strKF & "','" & strShipDate & "','" & strTerm & "','" & strForwarder & "'" & _
             " FROM erpdata..tblstocksqfhSUB a , erpbase..tblmappingdata   b,erpbase..tblcustomeroi c,erptemp..tbltsvnpiproduct d " & _
             " WHERE a.���̿����=b.SUBSTRATEID AND a.������=b.LOTID AND b.LOTID =c.SOURCE_BATCH_ID AND b.FILENAME =convert(VARCHAR(30),convert(int,c.id)) AND a.�Ϻ�=d.QTECHPTNO2  " & _
             " AND  a.���ݱ�� ='" & Trim(strOrderNo) & "'" & _
             " GROUP BY a.���ݱ��,LEFT(a.�󹤵�,1),c.MPN_DESC ,d.QTECHPTNO "
            
            AddSql2 (strSql)
              
        End If
        
       ' MsgBox "�Ѿ��ɹ�ִ����������", vbInformation, Me.Caption

     Else
        ExecProc = False
       ' MsgBox "����ִ��ʧ�ܣ�", vbInformation, Me.Caption
        Exit Function
        
     End If
     Exit Function
FError:
    MsgBox "ִ��ʧ�ܣ�ԭ��" & Err.DESCRIPTION, vbInformation, "Frm_uploadShippingList.ExecProc"
End Function




'�Զ���õ��ݱ�ţ����ǲ�һ���Ǳ��浽���Ͽ�ĵ��ݱ��
Private Sub AutoCode()
Dim strCompare          As String
Dim strDate             As String
Dim strSql              As String
Dim strSqlin              As String
Dim strSqlin_bak              As String
Dim maxid  As String
Dim rs                  As New ADODB.Recordset
Dim rs1                  As New ADODB.Recordset
    
    txt(0).text = ""
    strCompare = Trim("F" & Format(Now, "yymmdd"))
    strDate = Trim(Format(Now, "yymmdd"))
'    strSql = "SELECT MAX(LEFT(RTRIM(�����嵥���),LEN(RTRIM(�����嵥���))-4))+RIGHT('0000'+CAST(CAST(RIGHT(MAX(RTRIM(�����嵥���)),4) AS INT)+1 AS VARCHAR),4) ���� " & _
'             " FROM tblCodeRule WHERE �����嵥��� LIKE '" & strCompare & "%' "
             
    strSql = " select RIGHT(REPLICATE('0',4)+CAST((ISNULL(MAX(ID),0) + 1) AS varchar(10)),4) AS ����,ISNULL(MAX(ID),0) + 1 as id from erptemp..SHIP_NUM_SEQ a  WHERE  CONVERT(VARCHAR(100),a.create_date ,12)  = '" & strDate & "'  "
    strSqlin_bak = " SELECT MAX(id) FROM erptemp..SHIP_NUM_SEQ_BAK "
     
     
     If rs1.State = adStateOpen Then rs1.Close
     rs1.Open strSqlin_bak, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     maxid = rs1.Fields(0).Value
     rs1.Close
     
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs!���� = "0001" Then
    txt(0).text = strCompare & rs!����
    txtID_Text.text = Val(maxid) + 1
    Else
    txt(0).text = strCompare & Format((Val(rs!����) + 20000 - Val(maxid)), "0000")
    txtID_Text.text = rs.Fields(1).Value
    
    End If
    strSqlin = " INSERT INTO erptemp..SHIP_NUM_SEQ (CREATE_DATE,CREATE_BY,FLAG ) VALUES (GETDATE(),'" & txt(1).text & "','0') "
    
    
    
  If AddSql2(strSqlin) <> 0 Then
     MsgBox "��ȡ���ݺ�" & txt(0).text, vbInformation, "��ʾ"
     End If
    
'    If IsNull(Rs!����) Then
'        txt(0).Text = strCompare & "0001"
'    Else
'        txt(0).Text = Trim$("" & Rs!����)
'    End If
    rs.Close
    
End Sub

Private Sub cmdCreate_Click()
AutoCode
End Sub


'���沼��

Private Sub Form_Resize()
    control_resize

End Sub
Private Sub control_resize()

    On Error Resume Next
    
    SSTab1.Move SSTab1.Left, SSTab1.Top, Me.Width - SSTab1.Left - 350, Me.Height - SSTab1.Top - 400
    Select Case SSTab1.Tab
    
    Case 0
    

    Case 1
    

    
    Case 2
    
    Toolbar3.Move Toolbar3.Left, Toolbar3.Top, SSTab1.Width - Toolbar3.Left - 250, Toolbar3.Height

    SSTab2.Move SSTab2.Left, SSTab2.Top, SSTab1.Width - SSTab2.Left - 350, SSTab1.Height - SSTab2.Top - 400

    
    Select Case SSTab2.Tab
    
    Case 0
    Fra(1).Move Fra(1).Left, Fra(1).Top, SSTab1.Width - Fra(1).Left - 250, Fra(1).Height
    Fra(0).Move Fra(0).Left, SSTab2.Height - 2800, SSTab2.Width - Fra(0).Left - 350, Fra(0).Height
    Fra(2).Move Fra(2).Left, Fra(2).Top, SSTab2.Width - Fra(2).Left - 350, Fra(0).Top - Fra(2).Top - 300
    Fps_Ship(0).Move Fps_Ship(0).Left, Fps_Ship(0).Top, Fra(2).Width - 6800, Fra(2).Height - Fps_Ship(0).Top - 100
    Fps_Ship(1).Move Fra(2).Width - 6700, Fps_Ship(1).Top, 6500, Fra(2).Height - Fps_Ship(1).Top - 100
    Fps_Ship(2).Move Fps_Ship(2).Left, Fps_Ship(2).Top, Fra(0).Width - Fps_Ship(2).Top - 350, Fra(0).Height - Fps_Ship(2).Top - 400
    
    Case 1
    
    Fps_Ship_del(0).Move Fps_Ship_del(0).Left, Fps_Ship_del(0).Top, Fps_Ship_del(0).Width, SSTab2.Height - Fps_Ship_del(0).Top - 100
    Fps_Ship_del(1).Move Fps_Ship_del(1).Left, Fps_Ship_del(1).Top, SSTab2.Width - Fps_Ship_del(1).Left - 350, SSTab2.Height - Fps_Ship_del(1).Top - 100
        
    End Select
    
    End Select
    
    
    
End Sub
'��ʼ���ؼ�

Private Sub InitCtrl_Ship()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
    
    'Fps��ʼ��
    With Fps_Ship(0)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = FpsM.e_MCol - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
        .Col = FpsM.E_CHOOSE
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .SetText FpsM.E_CHOOSE, 0, "ѡ��"
        .SetText FpsM.e_ID, 0, "���"
        .SetText FpsM.E_cust, 0, "�ͻ�"
        .SetText FpsM.e_GDH, 0, "������/LOT��"
        .SetText FpsM.e_BigX, 0, "�����"
        .SetText FpsM.e_LH, 0, "�Ϻ�"
        .SetText FpsM.e_NUM, 0, "�����"
        .SetText FpsM.E_GG, 0, "���"
        .SetText FpsM.E_XH, 0, "�ͺ�"
        .SetText FpsM.E_UNIT, 0, "��λ"
        .SetText FpsM.e_KF, 0, "�����ⷿ"
        .SetText FpsM.E_DN, 0, "DN"
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ColWidth(-1) = 12
        .ColWidth(FpsM.E_CHOOSE) = 4
        .ColWidth(FpsM.e_ID) = 4
        .ColWidth(FpsM.E_cust) = 5
        .ColWidth(FpsM.E_UNIT) = 4
        .ZOrder
        .ReDraw = True
    End With
    With Fps_Ship(1)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = FpsD.e_MCol - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
     '   .Col = FpsD.e_BigX
     '   .Row = -1
        '.ColHidden = True
     '   .Col = FpsD.E_ID
     '   .Row = -1
      '  .ColHidden = True
        .Col = FpsD.E_Note2
        .Row = -1
        .ColHidden = True
        
        .SetText FpsD.e_BigX, 0, "�����"
        .SetText FpsD.E_XH, 0, "���"
        .SetText FpsD.e_GDH, 0, "������/LOT��"
        .SetText FpsD.e_LCK, 0, "���̿�/WaferID"
        .SetText FpsD.e_LH, 0, "�Ϻ�"
        .SetText FpsD.e_GNum, 0, "�ϸ���"
        .SetText FpsD.e_BLNum, 0, "���ϲ�����"
        .SetText FpsD.e_ZCNum, 0, "�Ƴ̲�����"
        .SetText FpsD.e_ID, 0, "���"
        .SetText FpsD.e_KF, 0, "�����ⷿ"
        .SetText FpsD.E_DN, 0, "DN"
        .SetText FpsD.E_Note2, 0, "Note"
'        '�趨�Ƿ�����
'        .UserColAction = UserColActionSort
'        For i = 1 To .MaxCols
'            .Col = i
'            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
'        Next
'        .colwidth(-1) = 12
'        .ZOrder
        .ReDraw = True
    End With
    
     With Fps_Ship_del(0)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = 6
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 1
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .SetText 1, 0, "ѡ��"
        .SetText 2, 0, "�ͻ�����"
        .SetText 3, 0, "�ͻ�����"
        .SetText 4, 0, "���ݱ��"
        .SetText 5, 0, "��������"

        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ColWidth(-1) = 12
        .ColWidth(1) = 4
        .ColWidth(2) = 4
        .ColWidth(3) = 18
        .ColWidth(4) = 10
         
        .ZOrder
        .ReDraw = True
    End With
    
    
    
    strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
   strRealName = Get_SqlStr2(strSql)

    '������
    txt(1).text = Trim(gUserName) & " " & Trim(strRealName)
    '��������
    txt(2).text = Format(Now(), "YYYY-MM-DD")
    '���벿��
   ' txt(3).Text = Trim(strUserDepartNUM) & Space(1) & Trim(strUserDepart)
     txt(3).text = "06 ���۲�"
    '���ؿͻ�������ջ��ͻ�
    strSql = "SELECT �ͻ����� FROM dbo.tblXCustomer"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
  '  Cob(0).Clear
    Cob(7).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
      '      Cob(0).AddItem Trim$("" & rs!�ͻ�����)
            Cob(7).AddItem Trim$("" & rs!�ͻ�����)
            rs.MoveNext
        Next
    Else
        MsgBox "�ͻ��������ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
        Exit Sub
    End If
    rs.Close
    '���ز��߱��
    strSql = "select rtrim(����)+' '+rtrim(˵��) ���߱�� from tblbase where ˵��2='���߱��' order by ˵��"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(1).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(1).AddItem Trim$("" & rs!���߱��)
            rs.MoveNext
        Next
    Else
        MsgBox "���߱�Ǽ���ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
        Exit Sub
    End If
    rs.Close
    '���ؿⷿ����
    strSql = "SELECT  ������� FROM ( SELECT rtrim(�ⷿ����)+' '+rtrim(�ⷿ����) ������� from erpbase..tblstock  where �ֿ�����='��Ʒ��'  AND �ⷿ����  NOT LIKE '%ί��%'  UNION   select rtrim(�ⷿ����)+' '+rtrim(�ⷿ����) ������� from erpbase..tblstock  where  �ⷿ���� = '85'  ) A ORDER BY ������� "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(2).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(2).AddItem Trim$("" & rs!�������)
            rs.MoveNext
        Next
    Else
        MsgBox "�ⷿ���Ƽ���ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
        Exit Sub
    End If
    rs.Close
    '�������䷽ʽ
    strSql = "select RTRIM(���䷽ʽ����)+' '+RTRIM(���䷽ʽ����) ���䷽ʽ from erpdata..tblXTransitMode"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(6).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(6).AddItem Trim$("" & rs!���䷽ʽ)
            rs.MoveNext
        Next
    Else
        MsgBox "���䷽ʽ����ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
        Exit Sub
    End If
    rs.Close
    '����DN
    ' strSql = "SELECT DISTINCT Delivery FROM erpbase..tblCustomerShippingUp WHERE Flag='Y' Order By 1 Desc"
    ' If rs.State = adStateOpen Then rs.Close
    ' rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    ' Cob(8).Clear
    ' If Not rs.EOF Then
        ' For i = 1 To rs.RecordCount
            ' Cob(8).AddItem Trim$("" & rs!Delivery)
            ' rs.MoveNext
        ' Next
    ' Else
        ' MsgBox "DN����ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
        ' Exit Sub
    ' End If
    ' rs.Close
End Sub

Private Sub Cob_DropDown(Index As Integer)
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset

    If Index = 3 Then
        '���ع�����
        strSql = "select distinct rtrim(a.������) ������ from erpdata..tblStockNum a inner join erpdata..tblbase b on a.���߱��=b.���� and b.˵��2='���߱��' " & _
                 " where a.�����>0 and b.����='" & Val(Trim(Cob(1).text)) & "'"
        If Trim$(Cob(2).text) <> "" Then
            If InStr(Cob(2).text, " ") > 0 Then
                strSql = strSql & " and a.�ⷿ���='" & Left(Cob(2).text, InStr(Cob(2).text, " ") - 1) & "'"
            Else
                strSql = strSql & " and a.�ⷿ���='" & Cob(2).text & "'"
            End If
        End If
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(3).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(3).AddItem Trim$("" & rs!������)
                rs.MoveNext
            Next
        End If
        rs.Close
    End If
    If Index = 9 Then
        '�����Ϻ�
        strSql = " SELECT a.SHIP_TO ������ FROM erptemp..customer_information a WHERE a.CUSTOMER = '" & Cob(0).text & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(9).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(9).AddItem Trim$("" & rs!������)
                rs.MoveNext
            Next
        Else
            MsgBox "�����ؼ���ʧ�ܣ�����ϵϵͳ����Ա��", vbInformation, "��ʾ"
            Exit Sub
        End If
        rs.Close
    End If

End Sub



Private Sub LoadDn()
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim date1 As String
Dim date2 As String

Dim strSql As String

   If Cob(0).text = "37" Then
        With Fps_Ship(1)
            .MaxRows = 0
        End With
        With Fps_Ship(0)
            .MaxRows = 0
        End With

        date1 = Format(dtShipDate_Ship.Value, "YYYY/MM/DD")
        date2 = Format(dtShipDate_Ship.Value + 1, "YYYY/MM/DD")

        'merry20200224 ���ؿɳ�����DN
        strSql = "SELECT  DELIVERY ,sum(QUANTITY) AS QTY FROM ERPBASE..tblCustomerShippingUp  WHERE  LASTUPDATEDATE >='" & date1 & "' and LASTUPDATEDATE <'" & date2 & "'  GROUP BY DELIVERY ORDER BY DELIVERY"

        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        List_dn(0).Clear
        List_dn(1).Clear
        If Not rs.EOF Then
             For i = 1 To rs.RecordCount
                If GetDnFlag(Trim(rs("DELIVERY")), rs("QTY")) = "������" Then
                    List_dn(0).AddItem Trim$("" & rs!Delivery) & "#" & rs!QTY
                End If
                If GetDnFlag(Trim(rs("DELIVERY")), rs("QTY")) = "δ���" Then
                    List_dn(1).AddItem Trim$("" & rs!Delivery) & "#" & rs!QTY
                End If
                
                rs.MoveNext
             Next
         End If
         rs.Close
    Else
        List_dn(0).Clear
        List_dn(1).Clear
    
    End If
   
  End Sub
Private Sub LoadSO()
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim date1 As String
Dim date2 As String

Dim strSql As String

   If Cob(0).text <> "SG005" Then
       Exit Sub
   End If
   If Trim(Cob(10).text) = "" Then
       Exit Sub
   End If
   
    With Fps_Ship(1)
        .MaxRows = 0
    End With
    With Fps_Ship(0)
        .MaxRows = 0
    End With
    'merry20200224 ���ؿɳ�����SO
    strSql = "SELECT a.SO_NO,a.SO_LINE,a.SO_QTY  FROM ERPBASE..tblCustomerShippingUp_SO a WHERE NOT EXISTS (SELECT 1 FROM erpdata..tblShipOrder_Dn WHERE SO_NO =a.SO_NO AND SO_LINE=a.SO_LINE)  and  RTRIM(a.DEVICE) ='" & Trim(Cob(10).text) & "' order by a.PSD,a.SO_NO,a.SO_LINE"
  
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(11).Clear
    If Not rs.EOF Then
         For i = 1 To rs.RecordCount
            Cob(11).AddItem Trim$("" & rs!SO_NO) & "#" & rs!SO_LINE & "#" & rs!SO_QTY
            rs.MoveNext
         Next
     End If
     rs.Close
    
  End Sub
Private Function GetDnFlag(dn As String, QTY As Long)
    Dim rs        As New ADODB.Recordset
    Dim i         As Integer
    Dim strSql    As String

    
    On Error Resume Next
    
    GetDnFlag = "������"
     'δ������ϵĲ���ʾ
     ' strSql = "SELECT  a.DELIVERY ,a.QUANTITY AS QTY_DN ,a.BATCHNUMBER ,sum(nvl(b.QTY,0))  as QTY_STOCK ,nvl(b.FLAG,0) as FLAG  FROM CUSTOMERSHIPPINGUPTBL  a  LEFT JOIN packing_detailed  b  ON  a.DELIVERY=b.DN_NUM AND a.BATCHNUMBER=b.JOB_ID " & _
     ' "WHERE a.DELIVERY='" & dn & "' GROUP BY  a.DELIVERY ,a.QUANTITY ,a.BATCHNUMBER ,b.FLAG  ORDER BY a.BATCHNUMBER"
    
     ' rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
     ' rs.MoveFirst
     ' If rs.RecordCount > 0 Then
          ' For i = 1 To rs.RecordCount
              ' If rs("QTY_DN") <> rs("QTY_STOCK") Or rs("FLAG") = 0 Then
                  ' GetDnFlag = "������"
                  ' Exit Function
              ' End If
              ' rs.MoveNext
          ' Next
     ' Else
          ' GetDnFlag = "������"
          ' Exit Function
     ' End If
     ' rs.Close

   strSql = "select isnull(sum(b.����),0) as ������� from erpdata..tblStockNumTree a,erpdata..tblStockNumSub b where  a.���=b.��� and b.�ⷿ��� in ('07','16') and a.DN='" & dn & "'"
   rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
   If rs.RecordCount > 0 Then
      rs.MoveFirst
       If rs("�������") <> QTY Then
           GetDnFlag = "δ���"
           'Exit Function
       End If
   Else
       GetDnFlag = "δ���"
       'Exit Function
   End If
   rs.Close
    
    '����������ģ�������=dn�����Ĳ���ʾ
    strSql = "select isnull(sum(quality),0) as ���������� from [erpdata].[dbo].[tblShipOrder_Dn] where  dn='" & dn & "'"
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
       If rs("����������") >= QTY Then
           GetDnFlag = "�����"
           Exit Function
       End If
    End If
    rs.Close
    
    
    Set rs = Nothing
End Function

Private Sub fps_Ship_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i               As Long
Dim j               As Integer
Dim strBigbox       As String
Dim intID           As Long
Dim strDN       As String
Dim strCustPN       As String

    'Fps����¼�
    If Index <> 0 Then Exit Sub
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    With Fps_Ship(0)
        .Row = Row
        .Col = FpsM.E_CHOOSE '�����ѡ��
        If .Value = 0 Then 'ѡ����
            .Col = FpsM.e_ID        '����ID
            intID = Val(Trim$(.text))
            .Col = FpsM.e_BigX      '�����
            strBigbox = Trim$(.text)
            .Col = FpsM.E_DN      'dn
            strDN = Trim$(.text)
             .Col = FpsM.E_GG      '���
            strCustPN = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.e_BigX
                If strBigbox = Trim$(.text) Then
                    .Col = FpsM.E_CHOOSE
                    .Value = 1
                    .Col = FpsM.E_GG      '���
                    If strCustPN <> Trim$(.text) Then
                        MsgBox "��ͬ���ֲ���װ��ͬһ������", vbInformation, "��ʾ"
                        Exit Sub
                    End If
                    '�޸�������ɫ
                    .Col = -1
                    .ForeColor = &HFF8080
                    
                End If
            Next
            '��ѯ��ϸ��Ϣ����ȡ�������Ϣ��
            Call SearchDetail(strBigbox, strDN, 1)
            '��ѯ�Ƿ����ջ��ͻ�
            Call SerachSHKH(intID, 1)
            '��ѯ������ַ mwl 2017.12.25 add
            Call SearchSHDZ
        Else    'ȡ��
            .Col = FpsM.e_ID        '����ID
            intID = Val(Trim$(.text))
            .Col = FpsM.e_BigX '�����
            strBigbox = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.e_BigX
                If strBigbox = Trim$(.text) Then
                    .Col = FpsM.E_CHOOSE
                    .Value = 0
                    '�޸�������ɫ
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
            'ɾ����ϸ��Ϣ��ɾ��ȡ���Ĵ������Ϣ��
            Call SearchDetail(strBigbox, strDN, 2)
            '��ѯ�Ƿ����ջ��ͻ�
            Call SerachSHKH(intID, 2)
            '��ѯ������ַ mwl 2017.12.25 add
            Call SearchSHDZ
        End If
    End With
    '���㹴ѡ��������������
    Call CalcBoxNum
    
End Sub
'��ѯ�ջ���ַ
Private Sub SearchSHDZ()
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim i                   As Long
Dim strCustTmp          As String
Dim strGDHTmp           As String

    With Fps_Ship(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.E_CHOOSE
            If .Value = 1 Then
                .Col = FpsM.E_cust      '�ͻ�����
                strCustTmp = Trim$(.text)
                .Col = FpsM.e_GDH       '������
                strGDHTmp = strGDHTmp + "," + Trim$(.text)
            End If
        Next
        If InStr(strGDHTmp, ",") > 0 Then
            strGDHTmp = Mid$(strGDHTmp, 2, Len(strGDHTmp) - 1)
        End If
    End With
    '��ѯ���Ͽ�
    strSql = "SELECT ShipTo,COUNT(DISTINCT ShipTo) BS FROM erpdata..tblSale_Shipto WHERE CustCode='" + strCustTmp + "' AND LotID IN('" + Replace$(strGDHTmp, ",", "','") + "')" & _
             " GROUP BY ShipTo"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        If Val(Trim$("" & rs!BS)) > 1 Then
            'txt(4).Text = "��ַ��ͬ"
        Cob(9).text = "��ַ��ͬ"
        Else
          '  txt(4).Text = Trim("" & Rs!ShipTo)
          Cob(9).text = Trim("" & rs!ShipTo)
        End If
    Else
        'txt(4).Text = ""
    End If
    rs.Close
End Sub
'���㹴ѡ��������������
Private Sub CalcBoxNum()
Dim i               As Long
Dim strBox          As String
Dim strBoxDetail    As String
Dim intBoxNum       As Integer
Dim intTotalBox     As Integer
    
    strBox = ""
    strBoxDetail = ""
    intBoxNum = 0
    intTotalBox = 0
    With Fps_Ship(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.e_BigX '�����
            If InStr(strBox, Trim$(.text)) <= 0 Then
                strBox = strBox & Trim$(.text) & ","
                intTotalBox = intTotalBox + 1
            End If
            .Col = FpsM.E_CHOOSE 'ѡ��
            If .Value = 1 Then
                .Col = FpsM.e_BigX '�����
                If InStr(strBoxDetail, Trim$(.text)) <= 0 Then
                    strBoxDetail = strBoxDetail & Trim$(.text) & ","
                    intBoxNum = intBoxNum + 1
                End If
            End If
        Next
    End With
    lbl(19).Caption = Trim$(intBoxNum) & "/" & Trim$(intTotalBox)
    
End Sub

Private Sub SerachSHKH(intID As Long, intBJ As Integer)
'��ѯ�ջ��ͻ�
Dim i               As Long
Dim strSql          As String
Dim rs              As New ADODB.Recordset
    
    '��ѯ����
    If intBJ = 1 And Trim(Cob(7).text) = "" Then
        strSql = "SELECT * FROM erpdata..Vw_AutoSearchSHKH Where ID=" & intID & ""
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            Cob(7).text = Trim$("" & rs!�ջ��ͻ�)
        End If
        rs.Close
    End If
    
    If intBJ = 2 And Trim(Cob(7).text) = "" Then 'ɾ����ϸʱ����ѯѡ�����Щ��Ŀ�ջ��ͻ�
        With Fps_Ship(0)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.E_CHOOSE    'ѡ��
                If Val(.text) = 1 Then
                    .Col = FpsM.e_ID    '����ID
                    strSql = "SELECT * FROM erpdata..Vw_AutoSearchSHKH Where ID=" & Val(Trim$(.text)) & ""
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount > 0 Then
                        Cob(7).text = Trim$("" & rs!�ջ��ͻ�)
                        Exit For
                    End If
                    rs.Close
                End If
            Next
        End With
    End If
    
End Sub
Private Sub SearchDetail(strBigbox As String, strDN As String, intBJ As Integer)
'��ѯ�����Ϣ
Dim i               As Long
Dim strSql          As String
Dim rs              As New ADODB.Recordset
    
    If intBJ = 1 Then '��ѯ��ϸ
        With Fps_Ship(1)
            '���ж��������Ĵ�����������û�м��ع���
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsD.e_BigX '�����
                If strBigbox = Trim$(.text) Then
                    Exit Sub
                End If
            Next
        End With
        '��ѯ����
        strSql = "SELECT '" & strBigbox & "' �����,a.���,b.������,b.���̿����,b.�Ϻ�" & _
                 ",CASE WHEN b.�ϸ���=0 THEN b.���� ELSE 0 END �ϸ��� " & _
                 ",CASE WHEN b.�ϸ���=2 THEN b.���� ELSE 0 END ���ϲ����� " & _
                 ",CASE WHEN b.�ϸ���=1 THEN b.���� ELSE 0 END �Ƴ̲����� " & _
                 ",b.id,b.�ⷿ��� " & _
                 " FROM dbo.f_getChild_1('" & strBigbox & "') a " & _
                 " INNER JOIN erpdata..tblStockNumSub b ON a.��� = b.���"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(1)
                '��Fps�м�������
                For i = 1 To rs.RecordCount
                    .MaxRows = .MaxRows + 1
                    .SetText FpsD.e_BigX, .MaxRows, Trim$("" & rs!�����)
                    .SetText FpsD.E_XH, .MaxRows, Trim$("" & rs!���)
                    .SetText FpsD.e_GDH, .MaxRows, Trim$("" & rs!������)
                    .SetText FpsD.e_LCK, .MaxRows, Trim$("" & rs!���̿����)
                    .SetText FpsD.e_LH, .MaxRows, Trim$("" & rs!�Ϻ�)
                    .SetText FpsD.e_GNum, .MaxRows, Trim$("" & rs!�ϸ���)
                    .SetText FpsD.e_BLNum, .MaxRows, Trim$("" & rs!���ϲ�����)
                    .SetText FpsD.e_ZCNum, .MaxRows, Trim$("" & rs!�Ƴ̲�����)
                    .SetText FpsD.e_ID, .MaxRows, Trim$("" & rs!id)
                    .SetText FpsD.e_KF, .MaxRows, Trim$("" & rs!�ⷿ���)
                    .SetText FpsD.E_DN, .MaxRows, strDN
    
                    rs.MoveNext
                Next
            End With
        End If
        rs.Close
    End If
    
    If intBJ = 2 Then 'ɾ����ϸ
        With Fps_Ship(1)
            Set .DataSource = Nothing
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = FpsD.e_BigX '�����
                If strBigbox = Trim$(.text) Then    '������ɾ��
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
    End If
    '��������
    Call CalcNum
End Sub
'��������
Private Sub CalcNum()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long
    
    lngGNum = 0
    lngBLNum = 0
    lngZCNum = 0
    With Fps_Ship(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsD.e_GNum '�ϸ���
            lngGNum = lngGNum + Val(Trim$(.text))
            .Col = FpsD.e_BLNum '���ϲ�����
            lngBLNum = lngBLNum + Val(Trim$(.text))
            .Col = FpsD.e_ZCNum '�Ƴ̲�����
            lngZCNum = lngZCNum + Val(Trim$(.text))
        Next
    End With
    '��ֵ��������
    txt(6).text = lngGNum
    txt(7).text = lngBLNum
    txt(8).text = lngZCNum
    
    If Cob(11).text <> "" Then '��ʾSO��������ѡ����������ʣ������
        lbl(23).Caption = lngGNum + lngBLNum + lngZCNum & "/" & lsolineqty
        If lngGNum + lngBLNum + lngZCNum > lsolineqty Then
            lbl(23).ForeColor = &HFF&
        Else
            lbl(23).ForeColor = &H8000&
            lbl(23).Caption = lbl(23).Caption & ",����:" & lsolineqty - (lngGNum + lngBLNum + lngZCNum)
     
        End If
    End If
End Sub

'��������
Private Function CheckQtyByDn()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long
Dim lngqty_dn       As Long
Dim strDN           As String
    
    CheckQtyByDn = ""
    
    
    With List_dn(0)
        For i = 0 To .ListCount - 1
            strDN = ""
            lngqty_dn = 0
            If .Selected(i) = True Then
                strDN = Split(Trim$("" & .List(i)), "#")(0)
                lngqty_dn = Val(Split(Trim$("" & .List(i)), "#")(1))
                lngGNum = 0
                lngBLNum = 0
                lngZCNum = 0
                With Fps_Ship(1)
                    For j = 1 To .MaxRows
                        .Row = j
                        .Col = FpsD.E_DN '�ϸ���
                        If Trim(.text) = strDN Then
                            .Col = FpsD.e_GNum '�ϸ���
                            lngGNum = lngGNum + Val(Trim$(.text))
                            .Col = FpsD.e_BLNum '���ϲ�����
                            lngBLNum = lngBLNum + Val(Trim$(.text))
                            .Col = FpsD.e_ZCNum '�Ƴ̲�����
                            lngZCNum = lngZCNum + Val(Trim$(.text))
                         End If
                    Next
                    
                End With
                
                If lngGNum + lngBLNum + lngZCNum <> lngqty_dn Then
                    CheckQtyByDn = strDN
                    MsgBox strDN & "  ��������,��˶� ", vbInformation, "��ʾ"
                    Exit Function
                End If
            End If
        Next
    End With

    


    
End Function
Private Function CheckQtyBySO()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long

    CheckQtyBySO = ""
    
    lngGNum = 0
    lngBLNum = 0
    lngZCNum = 0
    With Fps_Ship(1)
        For j = 1 To .MaxRows
            .Row = j
            .Col = FpsD.e_GNum '�ϸ���
            lngGNum = lngGNum + Val(Trim$(.text))
            .Col = FpsD.e_BLNum '���ϲ�����
            lngBLNum = lngBLNum + Val(Trim$(.text))
            .Col = FpsD.e_ZCNum '�Ƴ̲�����
            lngZCNum = lngZCNum + Val(Trim$(.text))

        Next
        
    End With
    
    If lngGNum + lngBLNum + lngZCNum > lsolineqty Then
        CheckQtyBySO = "ѡ����������SOline����"
        MsgBox " ѡ����������SOline����,��˶� ", vbInformation, "��ʾ"
        Exit Function
    End If


    


    
End Function


'ɾ�����޸ĸ�ֵ��Fps
Public Sub GiveFps(strdjbh As String, intBJ As Integer)
Dim i           As Long
Dim strSql      As String
Dim rs          As New ADODB.Recordset
    
    strSql = "select d.������,b.�ͻ�����,dbo.f_getparent(f.���) �����,A.������ַ,a.��ע,a.���䷽ʽ,a.���ϲ���,b.�ͻ�����,a.���ݱ��,a.���,a.���ϱ��,c.�Ϻ�,isnull(c.����ͺ�,'') as ���,c.�ͺ�," & _
            " c.������λ���� ��λ,a.����Ա,a.��������,isnull(a.���,'') as ���, isnull(a.��˲���,'') as ��˲���,SUM(f.����) ����,c.��������,dbo.usp_date(isnull(�������,'')) as ������� ,e.�ⷿ����+' '+e.�ⷿ���� �ⷿ,a.id,a.�ջ��ͻ� " & _
            " FROM erpdata..tblStockSQfh AS a " & _
            " inner join erpdata..tblStockSQfhsub f on a.���ݱ��=f.���ݱ�� and a.���=f.�������" & _
            " INNER JOIN dbo.tblXCustomer AS b ON a.�ͻ����� = b.�ͻ����� " & _
            " INNER JOIN dbo.tblSmainM2 AS c ON a.���ϱ�� = c.���ϱ��  " & _
            " inner join erpdata..tblstock e on a.�ֿ���=e.�ⷿ����" & _
            " INNER JOIN erpdata..tblstocknum d on a.id=d.id where a.���ݱ��='" & strdjbh & "' and a.���ձ��=0 and a.��������=1" & _
            " group by d.������,b.�ͻ�����,dbo.f_getparent(f.���),A.������ַ,a.��ע,a.���䷽ʽ,a.���ϲ���,b.�ͻ�����,a.���ݱ��,a.���,a.���ϱ��,c.�Ϻ�,isnull(c.����ͺ�,''),c.�ͺ�," & _
            " c.������λ����,a.����Ա,a.��������,isnull(a.���,''), isnull(a.��˲���,''),c.��������,dbo.usp_date(isnull(�������,'')) ,e.�ⷿ����,e.�ⷿ����,a.id,a.�ջ��ͻ�" & _
            " order by 3"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount = 0 Then
      Screen.MousePointer = 0
      MsgBox "��ѡ�����ⵥ�ѱ��ⷿ���գ���ˢ�µ�ǰ�������²�����", vbInformation, Me.Caption
      Exit Sub
   End If
   If rs.RecordCount > 0 Then
        Me.Tag = intBJ
        Me.Caption = "�����嵥" & IIf(intBJ = 1, "(����)", IIf(intBJ = 3, "(ɾ��)", "(�޸�)"))
        txt(0).text = strdjbh
        txt(1).text = Trim("" & rs("����Ա"))
        txt(2).text = Format(Trim("" & rs("��������")), "YYYY-MM-DD")
        txt(3).text = Trim("" & rs("���ϲ���"))
        Cob(0).text = Trim("" & rs("�ͻ�����"))
        Cob(6).text = Trim("" & rs("���䷽ʽ"))
        'txt(4).Text = Trim("" & Rs("������ַ"))
        Cob(9).text = Trim("" & rs("������ַ"))
        txt(5).text = Trim("" & rs("��ע"))
        Cob(2).text = Trim("" & rs("�ⷿ"))
        Cob(7).text = Trim$("" & rs("�ջ��ͻ�"))
        rs.MoveFirst
        With Fps_Ship(0)
            For i = 1 To rs.RecordCount
                .MaxRows = .MaxRows + 1
                .SetText FpsM.e_ID, .MaxRows, Trim$("" & rs!id)
                .SetText FpsM.E_cust, .MaxRows, Trim$("" & rs!�ͻ�����)
                .SetText FpsM.e_GDH, .MaxRows, Trim$("" & rs!������)
                .SetText FpsM.e_BigX, .MaxRows, Trim$("" & rs!�����)
                .SetText FpsM.e_LH, .MaxRows, Trim$("" & rs!�Ϻ�)
                .SetText FpsM.e_NUM, .MaxRows, Trim$("" & rs!����)
                .SetText FpsM.E_GG, .MaxRows, Trim$("" & rs!���)
                .SetText FpsM.E_XH, .MaxRows, Trim$("" & rs!�ͺ�)
                .SetText FpsM.E_UNIT, .MaxRows, Trim$("" & rs!��λ)
                .SetText FpsM.e_KF, .MaxRows, Trim$("" & rs!�ⷿ)
                rs.MoveNext
            Next
        End With
        '��ʾ����
        Screen.MousePointer = 0
        Me.Show vbModal
    End If
    rs.Close
End Sub





Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

    Case "QUERY_Ship"
        QueryData_Ship

    Case "SAVE_Ship"
        SaveData_Ship
        
    Case "DEL_Ship"
        DelData_Ship
      
    Case "HOME_Ship"
        exitFrm

End Select
End Sub


Private Sub DelData_Ship()
On Error GoTo ErrHandle
Dim strdjbh              As String
Dim i                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim strsql1              As String
Dim strSql2              As String
Dim strSql3              As String
Dim strSql4              As String
Dim strSql5              As String
Dim DelCnt               As Integer



   ' MsgBox "������", vbInformation, "��ʾ"
   ' Exit Sub
   DelCnt = 0
    With Fps_Ship_del(0)
    'У��
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If .text = "1" Then
            DelCnt = DelCnt + 1
            .Col = 4
            strdjbh = Trim(.text)
            If Get_SqlserverCnt(" select 1 from erpdata..tblStockSQfh where ���ݱ��='" & strdjbh & "' and  isnull(�������,'')<>''") > 0 Then
                MsgBox strdjbh & "�����Ѿ���������ˣ����ܽ���ɾ����", vbInformation, "��ʾ"
                Exit Sub
            End If
            If Get_SqlserverCnt(" select 1 from erpdata..tblStockSQfhsub where ���ݱ��='" & strdjbh & "' and ��Ʊ���=1 ") > 0 Then
                MsgBox strdjbh & "�����Ѿ���Ʊ�����ܽ���ɾ�����޸ģ�", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    Next
    If DelCnt = 0 Then
        MsgBox "��ѡ����Ҫɾ���ĵ��ݣ�", vbInformation, "��ʾ"
        Exit Sub
    
    End If
    'ɾ��
     For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If .text = "1" Then
            .Col = 4
            strdjbh = Trim(.text)
            If MsgBox("��ȷ��Ҫɾ��" & strdjbh & "��?", vbOKCancel, "��ʾ") = vbCancel Then
            
            Else
                AddSql2 ("insert into erpdata..tblStockSQfh_bak select GETDATE(),'ɾ��','" & gUserName & "', a.* from erpdata..tblStockSQfh a where a.���ݱ��='" & strdjbh & "'")
                AddSql2 ("insert into erpdata..tblStockSQfhsub_bak select GETDATE(),'ɾ��','" & gUserName & "', a.* from erpdata..tblStockSQfhsub a where a.���ݱ��='" & strdjbh & "'")

                AddSql2 ("DELETE FROM erpdata..tblStockSQfh where ���ݱ��='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblStockSQfhsub where ���ݱ��='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblStockMoveRec WHERE ���ݱ��='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erptemp..tblshipreport_new WHERE ship_order='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblshiporder_dn WHERE shiporder='" & strdjbh & "'")
                
                MsgBox "�ѳɹ�ɾ�����ݱ��:" & strdjbh, vbInformation, "��ʾ"
            End If
            
        End If
    Next
    End With
    ListOrderNumData

Exit Sub
ErrHandle:
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.DESCRIPTION, vbExclamation, Me.Caption
End Sub
Private Sub QueryData_Ship()
On Error GoTo ErrHandle
Dim i                   As Long
Dim j                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim strTmpID            As String
Dim strTmpDXH           As String
Dim IsHave              As Boolean
Dim strDNList           As String
Dim strSOList           As String

    If Cob(0).text = "" Then
        MsgBox "��ѡ��ͻ����룡", vbInformation, "��ʾ"
        Exit Sub
    End If
    strDNList = ""
    strSOList = ""
    With List_dn(0)
    If Cob(0).text = "37" Then
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If strDNList = "" Then
                    strDNList = Split(Trim$("" & .List(i)), "#")(0)
                Else
                    strDNList = strDNList & "','" & Split(Trim$("" & .List(i)), "#")(0)
                End If
            End If
        Next
    ElseIf Cob(0).text = "SG005" Then
    End If
    End With

        

    
    
    If Cob(0).text = "37" And (strDNList <> "" Or Cob(8).text <> "") Then
           Cob(6).text = "01 ��������"
           Cob(7).text = "37"
           Cob(9).text = "37"
           
    ElseIf Cob(0).text = "SG005" And strsono <> "" Then
           Cob(6).text = "10 ����"
           Cob(7).text = "SG005"
           Cob(9).text = GetSqlServerStr("Select End_Customer_Label_cust_name from ERPBASE..tblCustomerShippingUp_SO where so_no='" & strsono & "' and so_line='" & strsoline & "'")
    Else
       If Cob(2).text = "" Then
           MsgBox "��ѡ��ⷿ���ƣ�", vbInformation, "��ʾ"
           Exit Sub
       End If
     End If
    strSql = "SELECT 1 AS ѡ��,a.ID,a.�ͻ�����,a.������,dbo.f_getparent(b.���) �����,a.�Ϻ�,d.���,d.�ͺ�,d.������λ���� ��λ" & _
             ",SUM(b.����) ����,c.�ⷿ����+' '+c.�ⷿ���� �ⷿ "
    If strDNList <> "" Or Cob(8).text <> "" Then

        strSql = strSql + ", f.dn as DN "
    Else
        strSql = strSql + ", '' as DN "
    End If
         
    strSql = strSql + "  FROM erpdata..tblStockNum a INNER JOIN erpdata..tblStockNumSub b ON a.id=b.ID " & _
     " INNER JOIN tblstock c ON a.�ⷿ���=c.�ⷿ���� " & _
     " INNER JOIN tblSmainM2 d ON a.���ϱ��=d.���ϱ�� " & _
     " LEFT JOIN erpdata..tblWithWork e ON a.�������=e.������� AND a.�������=e.�������  " & _
     " INNER JOIN erpbase..tblmappingData g ON b.���̿����=g.SUBSTRATEID AND b.������=g.lotid  " & _
     " INNER JOIN ERPBASE..TBLCUSTOMEROI h ON g.lotid=h.SOURCE_BATCH_ID AND g.filename=convert(varchar(20),h.id) "
     
     
    '�����г�����������37�ͻ�DN��ѯ���� 2017.2.14 add mwl------
    
    If strDNList <> "" Then
        strSql = strSql + " INNER JOIN erpdata..tblStockNumTree f ON b.���=f.��� AND f.DN in('" & strDNList & "')"
    Else
        If Cob(8).text <> "" Then
           strSql = strSql + " INNER JOIN erpdata..tblStockNumTree f ON b.���=f.��� AND f.DN in('" & Cob(8).text & "')"
        End If
    End If

    '----------------------------------------------------------
    strSql = strSql + " Where a.����� > 0 "
    If Trim$(Cob(0).text) <> "" Then
        strSql = strSql & " And a.�ͻ�����='" & Trim(Cob(0).text) & "'"
    End If
    If Val(Trim$(Cob(1).text)) > 0 Then
        strSql = strSql & " And a.���߱��='" & Val(Trim(Cob(1).text)) & "'"
    End If
    If Trim$(Cob(2).text) <> "" Then
        strSql = strSql & " And a.�ⷿ���='" & Left(Trim(Cob(2).text), InStr(Trim(Cob(2).text), " ") - 1) & "'"
    Else
        If Cob(0).text = "37" And (strDNList <> "" Or Cob(8).text <> "") Then
           strSql = strSql & " And a.�ⷿ��� in ('07','16')"
        End If
        If Cob(0).text = "SG005" And strsono <> "" Then
           If UCase(Right(Trim$(Cob(10).text), 2)) = "-D" Then 'merry20200408��-D��β������Ʒ����30�ֲ��ҿ��
               strSql = strSql & " And a.�ⷿ��� in ('30')"
           Else
               strSql = strSql & " And a.�ⷿ��� in ('07','20')"
           End If
        End If
    End If
    If Trim$(Cob(3).text) <> "" Then
        strSql = strSql & " And a.������ Like '" & Trim(Cob(3).text) & "%'"
    End If
    If Trim$(Cob(4).text) <> "" Then
        strSql = strSql & " And a.�Ϻ� Like '" & Trim(Cob(4).text) & "%'"
    End If
    If Trim$(Cob(5).text) <> "" Then
        strSql = strSql & " And e.po_num Like '" & Trim(Cob(5).text) & "%'"
    End If
    If Trim$(Cob(10).text) <> "" Then
        If UCase(Right(Trim$(Cob(10).text), 2)) = "-D" Then 'merry20200408��-D��β������Ʒ����-Dȥ��ƥ��ͻ�����
            strSql = strSql & " And rtrim(h.MPN_DESC) = '" & Left(Trim(Cob(10).text), Len(Trim(Cob(10).text)) - 2) & "'"
        Else
            strSql = strSql & " And rtrim(h.MPN_DESC) = '" & Trim(Cob(10).text) & "'"
        End If
    End If
    
    strSql = strSql & " GROUP BY a.ID,a.�ͻ�����,a.������,dbo.f_getparent(b.���),a.�Ϻ�,d.���,d.�ͺ�,d.������λ����,c.�ⷿ����,c.�ⷿ����"
    
    If strDNList <> "" Or Cob(8).text <> "" Then

        strSql = strSql + " , f.dn order by 12, 5"
    Else
        strSql = strSql + " order by  5"
    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If chk.Value = 0 Then 'û�й�ѡ��ѡ�����Ϣ
        Fps_Ship(0).MaxRows = 0
        Fps_Ship(1).MaxRows = 0
        txt(6).text = ""
        txt(7).text = ""
        txt(8).text = ""
    Else
        With Fps_Ship(0)
            Set .DataSource = Nothing
            '��ɾ��û�й�ѡ����Ϣ
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = FpsM.E_CHOOSE
                If .Value = 0 Then
                    .DeleteRows i, 1 'ɾ����
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
    End If
    '��ֵ���ݵ�Fps�ؼ�
    
    If rs.RecordCount > 0 Then
        With Fps_Ship(0)
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                IsHave = False
                '���ж��ѹ�ѡ�������Ƿ��Ѿ����� ��ID�ʹ���ţ�
                If chk.Value = 1 Then
                    For j = 1 To .MaxRows
                        .Row = j
                        .Col = FpsM.e_ID 'ID
                        strTmpID = Trim$(.text)
                        .Col = FpsM.e_BigX  '�����
                        strTmpDXH = Trim$(.text)
                        If Trim$("" & rs!id) = strTmpID And Trim$("" & rs!�����) = strTmpDXH Then
                            IsHave = True
                            Exit For
                        End If
                    Next
                End If
                If IsHave = False Then  '�������Ѿ���ѡ����Ϣ�������Ժ��ȥ
                    .MaxRows = .MaxRows + 1
                    .SetText FpsM.e_ID, .MaxRows, Trim$("" & rs!id)
                    .SetText FpsM.E_cust, .MaxRows, Trim$("" & rs!�ͻ�����)
                    .SetText FpsM.e_GDH, .MaxRows, Trim$("" & rs!������)
                    .SetText FpsM.e_BigX, .MaxRows, Trim$("" & rs!�����)
                    .SetText FpsM.e_LH, .MaxRows, Trim$("" & rs!�Ϻ�)
                    .SetText FpsM.e_NUM, .MaxRows, Trim$("" & rs!����)
                    .SetText FpsM.E_GG, .MaxRows, Trim$("" & rs!���)
                    .SetText FpsM.E_XH, .MaxRows, Trim$("" & rs!�ͺ�)
                    .SetText FpsM.E_UNIT, .MaxRows, Trim$("" & rs!��λ)
                    .SetText FpsM.e_KF, .MaxRows, Trim$("" & rs!�ⷿ)
                    .SetText FpsM.E_DN, .MaxRows, Trim$("" & rs!dn)
                End If
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    '���㹴ѡ��������������
    Call CalcBoxNum
    If Cob(0).text = "SG005" Then
        GetOptSelection '��line�����Զ���ѡ
    End If
Exit Sub
ErrHandle:
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.DESCRIPTION, vbExclamation, Me.Caption
End Sub





Private Sub SentMesToStock(strText As String)
'�����ʼ����ֿ�
Dim strSentTo(100) As String
Dim strSentCC(20)  As String
Dim strSentTitle   As String
Dim strSentText    As String
Dim dirtemp        As String
Dim strTemp        As String
Dim i              As Integer
Dim strBand        As String
Dim CUSTOMER_CODE  As String
Dim REMARK As String
Dim strRealName As String
Dim strSql As String
Dim strFileName As String





If gUserName = "07885" Then
  '   Exit Sub
End If
strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
strRealName = Get_SqlStr2(strSql)

i = 0


If Cob(0).text = "37" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_ShipByDn.cfg"
    strSentTitle = Cob(0).text & "�ͻ�����(����" & Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & "������---�����"
    If mailflag = 1 Then '�����Զ��ʼ�
        strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & "," & Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & "��������" & strText
    ElseIf mailflag = 2 Then '�޸ĳ���ʱ���Զ��ʼ�
        strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & ", ����DN�޸ĳ���ʱ��Ϊ" & txtShipDate_Mod.text & strText
    End If
ElseIf Cob(0).text = "SG005" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_ShipBySO.cfg"
    strSentTitle = Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & " " & Cob(0).text & "����" & Cob(9).text & "   ������Ʊ�ű�ǩ "
    strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & "," & strText
Else
    dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_ShipByDn.cfg"
    strSentTitle = Format(Now(), "YYYY/MM/DD") & " " & Cob(0).text & "����" & Cob(9).text
    strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & "," & strText
End If
If gUserName = "07885" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_ShipByDn_test.cfg"
End If

Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

strSentCC(0) = ""
strSentCC(1) = ""
If SentMess(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "�ʼ��ѷ���", vbInformation, Me.Caption
Else
    MsgBox "�ʼ�����ʧ��", vbCritical, Me.Caption

End If

End Sub


Public Function SentMess(Subject As String, SentText As String, Recipient() As String, Attachment As String, RecipientCC() As String) As Boolean

If gUserName = "07885" Then
    SentMess = True
    'Exit Function
End If

    Dim JM As Object

    Set JM = CreateObject("JMAIL.Message")
    
    Dim Recipients()   As String

    Dim RecipientCCs() As String

    Dim strBodyinfo    As String

    Dim i              As Integer

    Dim strSql         As String

    Dim j              As Integer

    Dim rs             As New ADODB.Recordset

    Dim RsD            As New ADODB.Recordset

    On Error GoTo ErrHandler

    SentMess = False

    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "sqladmin" '�ʺ�
    JM.MailServerPassWord = "ksitadmin" '����
    JM.From = "sqladmin@ht-tech.com"    '����
    JM.FromName = "sqladmin"  '����������
    
    '�ռ���
    For i = 0 To UBound(Recipient) - 1
        If Recipient(i) <> "" Then
            JM.AddRecipient Recipient(i)
        End If
        
    Next
 
    '������
    For i = 0 To UBound(RecipientCC) - 1
        If RecipientCC(i) <> "" Then
            JM.AddRecipientCC RecipientCC(i)
        End If
        
    Next
    
    '����
    If Attachment <> "" Then
        If Dir(Attachment, vbNormal Or vbArchive) = "" Then
            Exit Function
        Else
            JM.AddAttachment Attachment

        End If

    End If
    
    JM.Subject = Subject
    'JM.AppendText SentText
    JM.HTMLBody = SentText
    SentMess = JM.Send("mail.ht-tech.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function

End Function

'-----------------------------------------------------------


Private Sub uploadSO(strFileName As String)

    On Error GoTo uploadSO_ErrON

    Dim VBExcel As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i       As Integer
    Dim j       As Integer
    Dim strso   As String
    Dim fSoLine  As Double
    Dim strShipDate As String
    Dim strCustDevice As String
    Dim lQty    As Long
    Dim strChar As String
    Dim strSOid    As Long
    Dim struploadid    As Long
    Dim insertsqlstr As String
    Dim rs    As New ADODB.Recordset
    Dim strTemp As String
    Dim strSql As String
    Dim strSql2 As String
    
    
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(strFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 43 Then
        MsgBox "SO��������", vbInformation, "����"
        GoTo uploadSO_Err
                        
    End If


    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
       
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If

            ' If i = 1 Then
                ' If Fps(0).MaxRows = 0 Then
                    ' .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " "), "'", " "))

                ' End If

            ' Else
                ' .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " "), "'", " "))

            ' End If

            If i > 1 Then
                If j = 4 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO#������Ϊ��", vbCritical, "����"
                        GoTo uploadSO_Err
                    
                    End If
                
                    strso = Trim$(xlSheet.Range(strChar & i))

                ElseIf j = 5 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO_LINE������Ϊ��", vbCritical, "����"
                        GoTo uploadSO_Err
                    
                    End If
                    fSoLine = Trim$(xlSheet.Range(strChar & i))

                ElseIf j = 3 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "PSD������Ϊ��", vbCritical, "����"
                        GoTo uploadSO_Err
                    
                    End If
                
                    strShipDate = Trim$(xlSheet.Range(strChar & i))
                
                ElseIf j = 7 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "DEVICE������Ϊ��", vbCritical, "����"
                        GoTo uploadSO_Err
                    
                    End If
                    strCustDevice = Trim$(xlSheet.Range(strChar & i))
                
                ElseIf j = 6 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO_QTY������Ϊ��", vbCritical, "����"
                        GoTo uploadSO_Err
                    
                    End If
                    lQty = CLng(Trim$(xlSheet.Range(strChar & i)))

                End If

            End If

        Next j
        strSql = "select * from erpbase..tblCustomerShippingUp_So where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
        If Get_SqlserverCnt(strSql) > 0 Then
        'ֱ�����µĸ��Ǿɵ�
            strSql = "delete  from erpbase..tblCustomerShippingUp_So where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
            AddSql2 (strSql)
            strSql = "delete  from CUSTOMERSHIPPINGUPTBL_SO where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
            AddSql (strSql)
           ' MsgBox strso & " " & fSoLine & "���ϴ������벻Ҫ�ظ��ϴ�", vbInformation, "��ʾ"
          '  GoTo uploadSO_Err

        End If
    Next i
    
    '��ʼ�ϴ�
  '  strSOid = Get_SqlserverNo("select SHIPPING_SO.nextval ID from dual")
  
     struploadid = Get_SqlserverNo("select max(UPLOADID) as maxid from erpbase..tblCustomerShippingUp_So") + 1
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strSOid = Get_SqlserverNo("select max(ID) as maxid from erpbase..tblCustomerShippingUp_So") + 1
        insertsqlstr = "'" & struploadid & "','" & strSOid & "'"
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)
            End If
            strTemp = "'" & Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " "), "'", " ")) & "'"
            insertsqlstr = insertsqlstr & " ," & strTemp
            
        Next j
        strSql = " insert into erpbase..tblCustomerShippingUp_So(UPLOADID,ID, ISSUE_DATE, SUBCON, PSD, SO_NO, SO_LINE, SO_QTY, DEVICE, Sales_Part_ID, Customer_PN, " & _
        " Customer_PO, packing, Description, Forwarder, Ship_to_Address, Country, ZIP, End_Customer_Label_cust_name, End_Customer_SHIP_Name,  " & _
        " End_Customer_attention, End_Customer_TEL, End_Customer_FAX, REMARK_1, REMARK_2, Express_Account, TERM, Ship_To_Location, Label_Version," & _
        " Revision_Index, Additional_Part_Information, Supplier_ID, Customer_Label_Code, Manufacturer_Number, Ship_To_City, Ship_To_Postal_Code, " & _
        " Customer_Part_Description, Customer_Dock_Code, Material_Handling_Code, Ship_to_Address1, Ship_to_Address2, Ship_to_Address3, OVT_Place_of_Origin, Sales_Region, Ship_To_Customer) values("
        strSql = strSql & insertsqlstr & ")"
       
         strSql2 = " insert into CUSTOMERSHIPPINGUPTBL_SO(UPLOADID,ID, ISSUE_DATE, SUBCON, PSD, SO_NO, SO_LINE, SO_QTY, DEVICE, Sales_Part_ID, Customer_PN, " & _
        " Customer_PO, packing, Description, Forwarder, Ship_to_Address, Country, ZIP, End_Customer_Label_cust_name, End_Customer_SHIP_Name,  " & _
        " End_Customer_attention, End_Customer_TEL, End_Customer_FAX, REMARK_1, REMARK_2, Express_Account, TERM, Ship_To_Location, Label_Version," & _
        " Revision_Index, Additional_Part_Information, Supplier_ID, Customer_Label_Code, Manufacturer_Number, Ship_To_City, Ship_To_Postal_Code, " & _
        " Customer_Part_Description, Customer_Dock_Code, Material_Handling_Code, Ship_to_Address1, Ship_to_Address2, Ship_to_Address3, OVT_Place_of_Origin, Sales_Region, Ship_To_Customer) values("
        strSql2 = strSql2 & insertsqlstr & ")"
       
             
       
       If AddSql2(strSql) = 0 Then
           GoTo uploadSO_Err
       End If
       If AddSql(strSql2) = 0 Then
           GoTo uploadSO_Err
        End If
    Next i
'��ʾ�ڽ�����
    Set rs = Get_SqlserveRs("select * from erpbase..tblCustomerShippingUp_So where UPLOADid='" & struploadid & "'")
  fps(0).MaxRows = 0
    If rs.RecordCount > 0 Then
        With fps(0)
            Set .DataSource = rs
        End With
    End If
    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing
        Set xlBook = Nothing

   End If
   MsgBox "�ϴ��ɹ�", vbInformation, "��ʾ"

    Exit Sub
uploadSO_Err:
    MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_uploadShippingList.showFps_SG005", vbExclamation + vbOKOnly, "Application Error"
    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
uploadSO_ErrON:
    GoTo uploadSO_Err
   ' MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_uploadShippingList.showFps_SG005", vbExclamation + vbOKOnly, "Application Error"

End Sub






















