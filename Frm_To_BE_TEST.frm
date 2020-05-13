VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_To_BE_TEST 
   Caption         =   "晶圆待验信息"
   ClientHeight    =   12345
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   18615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Frm_TO_BE_TEST"
   ScaleHeight     =   12345
   ScaleWidth      =   18615
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SST 
      Height          =   12375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   21828
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "信息上传"
      TabPicture(0)   =   "Frm_To_BE_TEST.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtPath"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fpS(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CommonDialog1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdUP(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdQuery"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCreate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ComCustomer"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtText1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkCheck1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkMsgAppend"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMsg"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtLotID"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "待验库存查询"
      TabPicture(1)   =   "Frm_To_BE_TEST.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCommand2"
      Tab(1).Control(1)=   "cmdCommand1"
      Tab(1).Control(2)=   "txtText4"
      Tab(1).Control(3)=   "txtText3"
      Tab(1).Control(4)=   "txtText2"
      Tab(1).Control(5)=   "fpS(1)"
      Tab(1).Control(6)=   "lblLabel4"
      Tab(1).Control(7)=   "lblLabel3"
      Tab(1).Control(8)=   "lblLabel2"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "待验仓晶圆退运"
      TabPicture(2)   =   "Frm_To_BE_TEST.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCommand4"
      Tab(2).Control(1)=   "cmdreject"
      Tab(2).Control(2)=   "chkall"
      Tab(2).Control(3)=   "txtLOT"
      Tab(2).Control(4)=   "cmdCommand3"
      Tab(2).Control(5)=   "fpS(2)"
      Tab(2).Control(6)=   "lblLabel7"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "仓库录入"
      TabPicture(3)   =   "Frm_To_BE_TEST.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fpS_WaferReceivedByStock"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Command5"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame1"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdExport"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69240
         TabIndex        =   81
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "信息填写"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -75000
         TabIndex        =   40
         Top             =   1200
         Width           =   18015
         Begin VB.TextBox txtExpressNumber 
            Height          =   360
            Left            =   9720
            TabIndex        =   77
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtSales 
            Height          =   360
            Left            =   9720
            TabIndex        =   76
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtRemark 
            Height          =   360
            Left            =   9720
            TabIndex        =   75
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox chk_All 
            Caption         =   "全选/全不选"
            Height          =   375
            Left            =   16320
            TabIndex        =   74
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtSupplierno 
            Height          =   375
            Left            =   1680
            TabIndex        =   73
            Top             =   2880
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtStockPos 
            Height          =   360
            Left            =   6000
            TabIndex        =   65
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox txtFABDevice 
            Height          =   360
            Left            =   6000
            TabIndex        =   60
            Top             =   1920
            Width           =   2295
         End
         Begin VB.ComboBox cbPONo 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6000
            TabIndex        =   59
            Top             =   1440
            Width           =   2295
         End
         Begin VB.ComboBox cbPn 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6000
            TabIndex        =   58
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtLot2 
            BackColor       =   &H00FF80FF&
            Height          =   360
            Left            =   1680
            TabIndex        =   46
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtCustomerDevice 
            BackColor       =   &H00FF80FF&
            Height          =   360
            Left            =   1680
            TabIndex        =   45
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00FF80FF&
            Height          =   360
            Left            =   1680
            TabIndex        =   44
            Top             =   2400
            Width           =   2295
         End
         Begin VB.ListBox lsWaferID 
            Columns         =   5
            Height          =   2535
            ItemData        =   "Frm_To_BE_TEST.frx":0070
            Left            =   12360
            List            =   "Frm_To_BE_TEST.frx":00BF
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   840
            Width           =   5415
         End
         Begin VB.ComboBox cbCustomerID 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   42
            Top             =   960
            Width           =   2295
         End
         Begin VB.ComboBox cbBonded 
            BackColor       =   &H00FF80FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_To_BE_TEST.frx":0127
            Left            =   6000
            List            =   "Frm_To_BE_TEST.frx":0131
            TabIndex        =   41
            Top             =   960
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dpArrivingDate 
            Height          =   375
            Left            =   1680
            TabIndex        =   63
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Format          =   106758145
            CurrentDate     =   43947
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "快递单号"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   80
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "销售人员"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   79
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备注"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9000
            TabIndex        =   78
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   8
            Left            =   5400
            TabIndex        =   72
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "双击-->"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4200
            TabIndex        =   71
            Top             =   520
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   7
            Left            =   13320
            TabIndex        =   70
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   69
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   5
            Left            =   5400
            TabIndex        =   68
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "晶圆片号"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12360
            TabIndex        =   67
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "库位"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            TabIndex        =   66
            Top             =   2400
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   64
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到货日期"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "晶圆厂机种"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            TabIndex        =   61
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "料号"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            TabIndex        =   57
            Top             =   480
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "采购单号"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            TabIndex        =   56
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   55
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   54
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   53
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "（*）"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   52
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "批次"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   51
            Top             =   1920
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户机种"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户代码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数量"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   48
            Top             =   2400
            Width           =   510
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "保/非保"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   47
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63960
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "修改"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66600
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71760
         TabIndex        =   37
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增/入待验库"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
      Begin FPSpreadADO.fpSpread fpS_WaferReceivedByStock 
         Height          =   6975
         Left            =   -74880
         TabIndex        =   35
         Top             =   4920
         Width           =   18015
         _Version        =   524288
         _ExtentX        =   31776
         _ExtentY        =   12303
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
         SpreadDesigner  =   "Frm_To_BE_TEST.frx":0143
      End
      Begin VB.CommandButton cmdCommand4 
         Caption         =   "退运记录查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68880
         TabIndex        =   34
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdreject 
         BackColor       =   &H000000FF&
         Caption         =   "退运"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70320
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   720
         Width           =   990
      End
      Begin VB.CheckBox chkall 
         Caption         =   "全选/全不选"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtLOT 
         Height          =   375
         Left            =   -74040
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71880
         TabIndex        =   28
         Top             =   720
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "删除"
         Height          =   360
         Left            =   15000
         TabIndex        =   27
         Top             =   1920
         Width           =   990
      End
      Begin VB.TextBox txtLotID 
         Height          =   375
         Left            =   12360
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtMsg 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   1170
         Left            =   8040
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   24
         Top             =   480
         Width           =   4935
      End
      Begin VB.CheckBox chkMsgAppend 
         Caption         =   "是否需要"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCheck1 
         Caption         =   "来料记录"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70080
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70080
         TabIndex        =   18
         Top             =   720
         Width           =   990
      End
      Begin VB.TextBox txtText4 
         Height          =   495
         Left            =   -73560
         TabIndex        =   17
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtText3 
         Height          =   495
         Left            =   -73560
         TabIndex        =   16
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtText2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73560
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox ComCustomer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "入库单生成"
         Height          =   480
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   480
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdUP 
         BackColor       =   &H00808080&
         Caption         =   "数据上传"
         Height          =   480
         Index           =   0
         Left            =   360
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H80000015&
         Caption         =   "退出"
         Height          =   480
         Index           =   2
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4560
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   9015
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   15901
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_To_BE_TEST.frx":0565
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   9015
         Index           =   1
         Left            =   -74880
         TabIndex        =   11
         Top             =   3120
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   15901
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_To_BE_TEST.frx":0A47
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   6615
         Index           =   2
         Left            =   -74880
         TabIndex        =   31
         Top             =   2040
         Width           =   15855
         _Version        =   524288
         _ExtentX        =   27966
         _ExtentY        =   11668
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_To_BE_TEST.frx":0F29
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblLabel7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "批号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotID"
         Height          =   195
         Left            =   11880
         TabIndex        =   26
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮件正文补充:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6720
         TabIndex        =   22
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "来料记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblLabel4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "批号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   15
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblLabel3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   14
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件路径"
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
         Index           =   1
         Left            =   4320
         TabIndex        =   4
         Top             =   1800
         Width           =   900
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   1800
         Width           =   5655
         VariousPropertyBits=   746604563
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "9975;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "Frm_To_BE_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemp() As String   '记录明细数据
Dim intCount  As Integer  '循环次数

Private Sub InitCustomerCode()
Dim i  As Integer
Dim rs As ADODB.Recordset

Set rs = Get_SqlserveRs("SELECT 客户代码 as PID,客户代码 as productname FROM erpdata.dbo.tblXCustomer order by 客户代码 ")
ComCustomer.Clear

If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        ComCustomer.AddItem Trim(rs("productname"))
      
        rs.MoveNext
    Next i

End If

rs.Close
Set rs = Nothing

End Sub




Private Sub cbBonded_Click()
    Dim rs_new As New ADODB.Recordset
    Dim j As Integer
    
        
    cbPONo.Clear
    If cbBonded.text = "保税" Then
        cbPONo.Enabled = True
        cbPONo.BackColor = &HFF80FF
        If Trim(cbPn.text) <> "" Then
            Set rs_new = Get_SqlserveRs("select distinct 采购单编号 from erpbase..OPENPO_WAFER where isnull(清关日期,'')<>'' and 未到货数量>0 and  料号='" & Trim(cbPn.text) & "'")
            If rs_new.RecordCount > 0 Then
                rs_new.MoveFirst
                For j = 1 To rs_new.RecordCount
                   cbPONo.AddItem Trim(rs_new("采购单编号"))
                   rs_new.MoveNext
                   
                Next
            End If
            
         End If
            
    ElseIf cbBonded.text = "非保税" Then
        cbPONo.text = ""
        cbPONo.Enabled = False
        cbPONo.BackColor = &HC0C0C0
    
    End If
    
End Sub


Private Sub cbPn_Change()
    cbCustomerID.Clear
    cbPONo.Clear
    cbBonded.text = ""
End Sub

Private Sub cbPn_Click()
Dim i  As Integer
Dim rs As ADODB.Recordset

cbCustomerID.Clear
cbPONo.Clear
cbBonded.text = ""
Set rs = Get_SqlserveRs("select DISTINCT CUSTOMERSHORTNAME from erptemp..tbltsvnpiproduct WHERE MARKETLASTUPDATE_BY='" & Trim(cbPn.text) & "'")

If rs.RecordCount > 0 Then
    rs.MoveFirst
    If rs.RecordCount = 1 Then
        cbCustomerID.text = Trim(rs("CUSTOMERSHORTNAME"))
    End If
    For i = 1 To rs.RecordCount
        cbCustomerID.AddItem Trim(rs("CUSTOMERSHORTNAME"))
        rs.MoveNext
    Next i

End If

rs.Close
Set rs = Nothing

End Sub

Private Sub chk_All_Click()
    Dim i As Integer
    
    If chk_All.Value = 1 Then
   
             With lsWaferID
                For i = 0 To .ListCount - 1
                   .Selected(i) = True
                Next
        
            End With
 
        
    ElseIf chk_All.Value = 0 Then

             With lsWaferID
                For i = 0 To .ListCount - 1
                   .Selected(i) = False
                Next
        
            End With
 
        
    End If
End Sub

Private Sub ChkAll_Click()

    Dim i As Integer
    
    If chkall.Value = 1 Then

        For i = 1 To fpS(2).MaxRows

            With fpS(2)
                .Row = i
                .Col = 1
                .text = 1

            End With

        Next i
        
    ElseIf chkall.Value = 0 Then

        For i = 1 To fpS(2).MaxRows

            With fpS(2)
                .Row = i
                .Col = 1
                .text = 0

            End With

        Next i
        
    End If
End Sub


Private Sub cmdCommand1_Click()
Update_OPENPOWAFER
QUERY_2

End Sub

Private Sub cmdCommand2_Click()
' SqlServerExporToExcel ("  SELECT c.客户代码,b.供应商名称,b.供应商编号, ISNULL(d.DEVICE,'') AS DEVICE, RTRIM(a.物料编号) as 物料编号,rtrim(a.批号) LOT,a.库位,a.当前存量,a.建立日期 ,CASE WHEN d.BONDED = 1 THEN '保税' ELSE '非保税'  end  FROM ERPBASE..tblstocknum a  INNER JOIN tblSupplierData  b  ON b.供应商编号 = a.供应商编号" & _
          ' " INNER JOIN erpdata..tblXCustomer  c  ON c.客户名称 = b.供应商名称  LEFT JOIN erptemp..TO_BE_TESTED d ON d.CUSTOMER = c.客户代码 AND d.LOT = a.批号 " & _
         ' " WHERE a.仓库编号 = '54' AND a.当前存量 > 0 ")
         
SqlServerExporToExcel ("  SELECT c.客户代码,isnull(e.料号,'') AS 料号,b.供应商名称,b.供应商编号, ISNULL(d.DEVICE,'') AS DEVICE, RTRIM(a.物料编号) as 物料编号,rtrim(a.批号) LOT,a.库位,a.当前存量,a.建立日期 ,CASE WHEN d.BONDED = 1 THEN '保税' ELSE '非保税'  end " & _
           ",isnull(d.remark3,'') AS 采购单编号  FROM ERPBASE..tblstocknum a  INNER JOIN tblSupplierData  b  ON b.供应商编号 = a.供应商编号" & _
           " INNER JOIN erpdata..tblXCustomer  c  ON c.客户名称 = b.供应商名称  LEFT JOIN erptemp..TO_BE_TESTED d ON d.CUSTOMER = c.客户代码 AND d.LOT = a.批号 " & _
           " LEFT JOIN  erpdata..tblSmainM2 e ON e.物料编号=a.物料编号 " & _
           " WHERE a.仓库编号 = '54' AND a.当前存量 > 0 ")
           
   

End Sub

Private Sub cmdCommand4_Click()
chkall.Visible = False
QUERY_5
End Sub

Private Sub cmdExport_Click()
Call ExportExcel(fpS_WaferReceivedByStock)
End Sub




Private Sub ExportExcel(fpS As fpSpread)
    Dim xlsApp      As Excel.Application
    Dim xlsBook     As Excel.Workbook
    Dim xlsSheet    As Excel.Worksheet
    Dim i           As Long
    Dim j           As Long
    
    On Error GoTo Ert
    
    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True

    End With
   
    With fpS

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
             '   If j <= 14 Then
            '       xlsSheet.Cells(i + 1, j) = .text
            '    Else
                
                    xlsSheet.Cells(i + 1, j) = Trim$(("'" & .text))
            '    End If

            Next j
       
        Next i

    End With
    xlsApp.Visible = True
    
    With xlsSheet.Range("2:" & i + 1)
        .horizontalAlignment = xlLeft
    End With
    xlsSheet.Range("A1").Select
    xlsApp.Columns.AutoFit

    
    
    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing
    Exit Sub
    
Ert:
    MsgBox Err.DESCRIPTION
    
    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    

End Sub



Private Sub cmdquery_Click()

    If chkCheck1.Value = False Then
        QUERY_1
        cmdCreate.Enabled = True
    Else
        QUERY_3

    End If

End Sub



Private Sub cmdreject_Click()
   Dim i As Integer
   Dim sqlStr As String
   Dim Cnt_Reject As Integer
   Dim WaferId_Temp As String

   Cnt_Reject = 0

    With fpS(2)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .text <> "" Then
                If .text = 1 Then
                    Cnt_Reject = Cnt_Reject + 1
                    
                End If
            End If
        Next
        If Cnt_Reject = 0 Then
            MsgBox "请选择要退运的WaferId", vbInformation, "提示"
            Exit Sub
        End If
        
        If MsgBox("共选择了" & Cnt_Reject & "片晶圆,请确认要退运 ?", vbYesNo, "退运提示") = vbNo Then
            MsgBox "退运取消", vbInformation, "提示"
            Exit Sub
    
        End If
       
        
        AddSql2 ("UPDATE erpbase..tblStockNum SET 当前存量 =  当前存量  -  " & Cnt_Reject & "  WHERE 仓库编号 = '54' and  批号 = '" & Trim(txtLOT.text) & "'")
        AddSql ("UPDATE  MAPPINGDATA37 SET WF =  WF  -  " & Cnt_Reject & "  WHERE   BATCH = '" & UCase(Trim(txtLOT.text)) & "'")
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .text <> "" Then
                If .text = 1 Then
                    .Col = 5
                    WaferId_Temp = Trim(.text) 'wafer_id

                    AddSql2 ("UPDATE erptemp..TO_BE_TESTED_sub SET flag = '-1' , LAST_UPDATE_DATE=getdate(),LAST_UPDATE_BY='" & gUserName & "' WHERE   LOT = '" & Trim(txtLOT.text) & "' and wafer_id='" & WaferId_Temp & "'")
                    AddSql ("INSERT INTO mappingdatatest_bak SELECT * FROM mappingdatatest WHERE SUBSTRATEID ='" & UCase(Trim(txtLOT.text)) & WaferId_Temp & "'")
                    AddSql ("DELETE FROM mappingdatatest WHERE SUBSTRATEID ='" & UCase(Trim(txtLOT.text)) & WaferId_Temp & "'")
                    AddSql2 ("INSERT  INTO ERPBASE..tblmappingData_bak SELECT * FROM ERPBASE..tblmappingData WHERE SUBSTRATEID ='" & UCase(Trim(txtLOT.text)) & WaferId_Temp & "'")
                    AddSql2 ("DELETE FROM ERPBASE..tblmappingData WHERE SUBSTRATEID ='" & Trim(txtLOT.text) & WaferId_Temp & "'")
                    AddSql2 ("DELETE FROM tblToInRec_Wafer_Wait WHERE   批号 = '" & Trim(txtLOT.text) & "' and 晶圆ID ='" & Trim(txtLOT.text) & WaferId_Temp & "'")
                    
                    '将mappingdatatest中要退运的晶圆移到将mappingdatatest_bak
                    'ERPBASE..tblmappingData 中要退运的晶圆移到将ERPBASE..tblmappingData_bak
                    
                End If
            End If
        Next
        
        
    End With
    MsgBox "退运已完成", vbInformation, "提示"
End Sub




Private Sub Command1_Click()
Dim strLotID As String
Dim strsql As String
Dim lQty As Long

If txtLotID.text = "" Then
    MsgBox "请输入要删除的LotID", vbInformation, "提示"
    Exit Sub
    
End If

strLotID = Trim$(txtLotID.text)
'strSql = "select WF from MAPPINGDATA37 where batch = '" & strLotID & "'"
'lQty = Get_OracleNo(strSql)
strsql = "select * from  erptemp..to_be_tested_sub  where LOT = '" & strLotID & "'"
lQty = Get_SqlserverCnt(strsql)
If lQty = 0 Then
    MsgBox "查询不到该lot", vbInformation, "提示"
    Exit Sub
End If

If MsgBox("查询到该LOT有" & lQty & "片,请确认是否删除?", vbYesNo, "删除提示") = vbNo Then
    MsgBox "删除取消", vbInformation, "提示"
    Exit Sub
    
End If

AddSql ("delete from MAPPINGDATA37 where batch = '" & strLotID & "' ")
AddSql ("delete from mappingdatatest where lotid = '" & strLotID & "'")
AddSql2 ("delete from [ERPBASE].[dbo].[tblmappingData] where lotid = '" & strLotID & "'")
AddSql2 ("delete from  erptemp..to_be_tested  where LOT = '" & strLotID & "'")
AddSql2 ("delete from  erptemp..to_be_tested_sub  where LOT = '" & strLotID & "'")

MsgBox "删除已完成", vbInformation, "提示"

End Sub


Private Sub Command2_Click()
Dim strsql As String
Dim strWaferID As String
Dim i As Integer

Dim strArrivingDate As String
Dim strCustomerID As String
Dim strCustomerDevice As String
Dim strLot2 As String
Dim intqty As Integer
Dim strPN As String
Dim strPONO As String
Dim intBonded As Integer
Dim strFabDevice As String
Dim strStockPos As String
Dim strExpressNumber As String
Dim strSales As String
Dim strRemark As String


Dim intFlag As String
Dim intID As Integer
Dim strno As String
Dim strmatno As String


    If checkdata_stockdata = False Then
        Exit Sub
    End If
    
  strWaferID = ""
    With lsWaferID
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If strWaferID = "" Then
                    strWaferID = Trim$("" & .List(i))
                Else
                    strWaferID = strWaferID & "." & Trim$("" & .List(i))
                End If
            End If
        Next

    End With
    strArrivingDate = Format(dpArrivingDate.Value, "yyyy/mm/dd")
    strCustomerID = Trim(cbCustomerID.text)
    strCustomerDevice = Trim(txtCustomerDevice.text)
    strLot2 = Trim(txtLot2.text)
    intqty = Trim(txtQty.text)
    strPN = Trim(cbPn.text)
    strmatno = GetSqlServerStr("SELECT 物料编号 FROM erpbase..tblSmainM2 WHERE 料号='" & strPN & "'")
    strPONO = Trim(cbPONo.text)
    If cbBonded.text = "保税" Then
        intBonded = 1
    ElseIf cbBonded.text = "非保税" Then
        intBonded = 0
    End If
    intFlag = 0
    
    strFabDevice = Trim(txtFABDevice.text)
    strStockPos = Trim(txtStockPos.text)
    strExpressNumber = Trim(txtExpressNumber.text)
    strSales = Trim(txtSales.text)
    strRemark = Trim(txtRemark.text)
  
     intID = Get_SqlserverNo("select max(id) from erpbase..WaferReceivedByStock ") + 1
     strsql = "insert into erpbase..WaferReceivedByStock(id,CUSTOMER,DEVICE,LOT,QTY,WAFER_ID,BONDED,FAB_DEVICE,PO_NO,ExpressNumber,Sales,Remark1,料号,库位,ARRIVING_DATE,FLAG,CREATE_DATE,CREATE_BY ) values( " & _
     intID & ", '" & strCustomerID & "','" & strCustomerDevice & "','" & strLot2 & "','" & intqty & "','" & strWaferID & "','" & intBonded & "','" & strFabDevice & "','" & strPONO & "','" & strExpressNumber & "','" & strSales & "','" & strRemark & "','" & strPN & "','" & strStockPos & "','" & strArrivingDate & "','" & intFlag & "', GETDATE() ,'" & gUserName & "')"
     AddSql2 (strsql)

'入库
    ReDim strTemp(5)
      strTemp(1) = intqty & "◆"
      strTemp(2) = Trim$(Now() + 100) + "◆"
      strTemp(3) = strLot2 + "◆"
      strTemp(4) = strLot2 + "◆"
    '  strlot = strLot2
    '  strdevice = Trim$(.Text)
  
      strTemp(0) = strmatno + "◆"
      strTemp(5) = strmatno + "◆"
      intCount = 1
      If SaveDataByStock Then
          '获取入库单号,写明细表
          strno = GetSqlServerStr("select top 1 rtrim(入库单编号) from erpbase..tbltoinrec where  仓库编号='54' and 制单='" & gUserName & "' order by 入库单编号 desc  ")
          With lsWaferID
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    AddSql2 ("insert into erpbase..WaferReceivedByStock_sub(id,LOT,WAFER_ID,FLAG) values(" & intID & ",'" & strLot2 & "','" & Trim$("" & .List(i)) & "'," & intFlag & ")")
                    AddSql2 ("insert into Erpbase..tblToInRec_Wafer_Wait (入库单编号,批号,晶圆ID,FLAG) VALUES('" & strno & "','" & strLot2 & "','" & strLot2 & Trim$("" & .List(i)) & "',0)")
                End If
            Next
    
         End With
     End If
     MsgBox "新增完成", vbInformation, "提示"

End Sub

Private Function checkdata_stockdata()
    Dim intwafercnt As Integer
    Dim i As Integer
    Dim strWaferID As String
    Dim strLot2 As String
    Dim strsql As String
    Dim strSql2 As String
    Dim strCustomerID As String
    Dim flag1 As Boolean
    Dim flag2 As Boolean
    Dim flag3 As Boolean
    
    checkdata_stockdata = False
    txtSupplierno.text = GetSqlServerStr(" SELECT b.供应商编号 FROM ERPBASE..tblXCustomer a," & "  ERPBASE..tblSupplierData  b  WHERE a.客户代码 = '" & Trim(cbCustomerID.text) & "' AND b.供应商名称 = a.客户名称 ")
    If txtSupplierno.text = "" Then
        MsgBox "供应商代码有误，请重新填写", vbInformation, "提示"
        Exit Function
    End If

    '必填项
    If Trim(cbCustomerID.text) = "" Or Trim(txtCustomerDevice.text) = "" Or Trim(txtLot2.text) = "" Or Trim(txtQty.text) = "" Or Trim(cbBonded.text) = "" Then
        MsgBox "带*号的为必填项目，请填写完整", vbInformation, "提示"
        Exit Function
    End If
    If Trim(cbBonded.text) = "保税" And Trim(cbPONo.text) = "" Then
        MsgBox "保税晶圆必须填写采购单号，请填写完整", vbInformation, "提示"
        Exit Function
    End If
    
    strCustomerID = Trim(cbCustomerID.text)
    strLot2 = Trim(txtLot2.text)
    '先判断lot,再判断wafer
    flag1 = False '大仓
    flag2 = False '待验仓
    flag3 = False 'bank
    If Get_SqlserverCnt("SELECT  1 from erpbase..tblToInRec_Wafer where 批号='" & strLot2 & "'") > 0 Or Get_SqlserverCnt("SELECT 1  FROM  erpbase..tblCustomerOI  a INNER JOIN ERPBASE ..tblToInRec_Wafer  c ON c.批号=a.SOURCE_BATCH_ID WHERE a.FAB_CONV_ID='" & strLot2 & "'") > 0 Then
        flag1 = True
    End If
    If Get_SqlserverCnt("SELECT  1 from erpbase..tblToInRec_Wafer_Wait where  FLAG=0 AND  批号='" & strLot2 & "'") > 0 Then
        flag2 = True
    End If
    If Get_SqlserverCnt("SELECT  1 from erptemp..to_be_tested_sub where LOT='" & strLot2 & "'") > 0 Then
        flag3 = True
    End If

    intwafercnt = 0
    With lsWaferID
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If flag1 = True Then
                    strsql = "SELECT  1 from erpbase..tblToInRec_Wafer where  晶圆ID='" & strLot2 & Trim$("" & .List(i)) & "' "
                    strSql2 = "SELECT 1  FROM  erpbase..tblCustomerOI  a INNER JOIN ERPBASE ..tblToInRec_Wafer  c ON c.批号=a.SOURCE_BATCH_ID WHERE a.FAB_CONV_ID='" & strLot2 & "' and c.晶圆ID='" & strLot2 & Trim$("" & .List(i)) & "' "
                    If Get_SqlserverCnt(strsql) > 0 Or Get_SqlserverCnt(strSql2) > 0 Then
                        MsgBox Trim$("" & .List(i)) & "号片已存在大仓，不可入待验仓！", vbInformation, "提示"
                        Exit Function
                    End If
                End If
                If flag2 = True Then
                    strsql = "SELECT  1 from erpbase..tblToInRec_Wafer_Wait where  FLAG=0 AND 晶圆ID='" & strLot2 & Trim$("" & .List(i)) & "' "
                    If Get_SqlserverCnt(strsql) > 0 Then
                        MsgBox Trim$("" & .List(i)) & "号片已存在待验仓，不可重复入！", vbInformation, "提示"
                        Exit Function
                    End If
                End If
                If flag3 = True Then
                    strsql = "SELECT  1 from erptemp..to_be_tested_sub  where LOT='" & strLot2 & "' AND   wafer_id='" & Trim$("" & .List(i)) & "' "
                    If Get_SqlserverCnt(strsql) > 0 Then
                        MsgBox Trim$("" & .List(i)) & "号片已上传Bank，不需重复上传！", vbInformation, "提示"
                        Exit Function
                    End If
                End If
                
                If strWaferID = "" Then
                    strWaferID = Trim$("" & .List(i))
                Else
                    strWaferID = strWaferID & "','" & Trim$("" & .List(i))
                End If
                intwafercnt = intwafercnt + 1

            End If
        Next

    End With
    '数量一致
    If IsNumeric(Trim(txtQty.text)) = False Then
        MsgBox "数量栏位需为数字，请确认", vbInformation, "提示"
        Exit Function
    End If
    If intwafercnt <> Val(Trim(txtQty.text)) Then
        MsgBox "所选片数与数量不一致，请确认", vbInformation, "提示"
        Exit Function
    End If
    checkdata_stockdata = True

End Function






Private Sub Command3_Click()
Dim strsql As String
Dim rs       As New ADODB.Recordset

strsql = "select CUSTOMER as 客户代码,DEVICE as 客户机种,LOT,QTY  as 数量,WAFER_ID,case BONDED when 1 then '保税' else '非保税' end as '保税/否',po_no as 采购单号,ARRIVING_DATE as 到货日期,Fab_Device,料号,ExpressNumber as 快递单号,Sales as 销售人员,库位,Remark1 as 备注 from erpbase..WaferReceivedByStock where 1=1 "

If Trim(cbCustomerID.text) <> "" Then
    strsql = strsql & " and CUSTOMER ='" & Trim(cbCustomerID.text) & "'"
End If

If Trim(txtCustomerDevice.text) <> "" Then
    strsql = strsql & " and DEVICE ='" & Trim(txtCustomerDevice.text) & "'"
End If

If Trim(txtLot2.text) <> "" Then
    strsql = strsql & " and LOT ='" & Trim(txtLot2.text) & "'"
End If
Set rs = Get_SqlserveRs(strsql)
With fpS_WaferReceivedByStock
    .MaxRows = 0
    Set .DataSource = rs
End With




End Sub

Private Sub Form_Load()
InitCustomerCode
cmdCreate.Enabled = False

With fpS(0)
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 1, 0, "客户代码"
    .SetText 2, 0, "机种"
    .SetText 3, 0, "批号"
    .SetText 4, 0, "数量"
    .SetText 5, 0, "WAFER_ID"
    .SetText 6, 0, "保税/非保税"

End With

With fpS(1)
    .Col = -1
    .Row = -1
    .Lock = True

End With

With fpS(2)
    .Col = -1
    .Row = -1
    .Lock = True

End With

End Sub

Private Sub QUERY_1()
Dim querysql As String
Dim rs       As New ADODB.Recordset

If ComCustomer.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

querysql = "  SELECT '' AS 选择,  a.customer AS 客户代码,a.Device AS 机种 ,a.Lot AS 批号,a.qty AS 数量,a.wafer_id AS WAFER_ID  ,'' as  库位 " & _
          "   ,CASE WHEN  a.bonded = 1 THEN '保税' ELSE '非保税' END 保税 ,ISNULL( in_qty,0) AS 已入   " & _
         "   ,a.create_date AS 上传时间 ,a.create_by AS 上传人  ,ISNULL(MAX(c.FNumber),a.DEVICE ) AS 晶圆料号,isnull(a.remark3,'') AS 采购单编号  FROM  erptemp..to_be_tested a  LEFT JOIN erptemp..tbltsvnpiproduct b   ON  b.CUSTOMERPTNO1 = a.DEVICE  " & _
         "    LEFT JOIN AIS20141114094336..t_ICItem c ON c.F_101 = b.MARKETLASTUPDATE_BY " & _
         "    WHERE a.flag = '0'  and a.customer = '" & ComCustomer.text & "' GROUP BY  a.customer ,a.Device,a.Lot ,a.qty ,a.wafer_id ,a.BONDED,a.IN_QTY,a.create_date ,a.create_by  ,isnull(a.remark3,'') "
fpS(0).MaxRows = 0
If rs.State = adStateOpen Then rs.Close
rs.Open querysql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    Call ListDataType(rs)
Else
    MsgBox "无数据", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub QUERY_2()
Dim querysql As String
Dim rs       As New ADODB.Recordset

If Trim(txtText2.text) = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

' querysql = "  SELECT c.客户代码,b.供应商名称,b.供应商编号, ISNULL(d.DEVICE,'') AS DEVICE, RTRIM(a.物料编号) as 物料编号,rtrim(a.批号) LOT,a.库位,a.当前存量,a.建立日期 ,CASE WHEN d.BONDED = 1 THEN '保税' ELSE '非保税'  end  FROM ERPBASE..tblstocknum a  INNER JOIN tblSupplierData  b  ON b.供应商编号 = a.供应商编号" & _
          ' " INNER JOIN erpdata..tblXCustomer  c  ON c.客户名称 = b.供应商名称  LEFT JOIN erptemp..TO_BE_TESTED d ON d.CUSTOMER = c.客户代码 AND d.LOT = a.批号 " & _
         ' " WHERE a.仓库编号 = '54' AND a.当前存量 > 0  AND c.客户代码 =   '" & Trim(txtText2.Text) & "'  "
         
querysql = "  SELECT c.客户代码,isnull(e.料号,'') AS 料号,b.供应商名称,b.供应商编号, ISNULL(d.DEVICE,'') AS DEVICE, RTRIM(a.物料编号) as 物料编号,rtrim(a.批号) LOT,a.库位,a.当前存量,a.建立日期 ,CASE WHEN d.BONDED = 1 THEN '保税' ELSE '非保税'  end as 保税非保" & _
           ",isnull(d.remark3,'') AS 采购单编号  FROM ERPBASE..tblstocknum a  INNER JOIN tblSupplierData  b  ON b.供应商编号 = a.供应商编号" & _
           " INNER JOIN erpdata..tblXCustomer  c  ON c.客户名称 = b.供应商名称  LEFT JOIN erptemp..TO_BE_TESTED d ON d.CUSTOMER = c.客户代码 AND d.LOT = a.批号 " & _
           " LEFT JOIN  erpdata..tblSmainM2 e ON e.物料编号=a.物料编号 " & _
           " WHERE a.仓库编号 = '54' AND a.当前存量 > 0  AND c.客户代码 =   '" & Trim(txtText2.text) & "'  "
           
   
           
           
           

If Trim(txtText3.text) <> "" Then
    querysql = querysql + " AND a.物料编号 = '" & Trim(txtText3.text) & "'  "

End If

If Trim(txtText4.text) <> "" Then
    querysql = querysql + " AND a.批号 = '" & Trim(txtText4.text) & "'  "

End If

fpS(1).MaxRows = 0
If rs.State = adStateOpen Then rs.Close
rs.Open querysql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    Call ListDataType1(rs)
Else
    MsgBox "无数据", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub QUERY_3()
Dim querysql As String
Dim rs       As New ADODB.Recordset

If Trim(ComCustomer.text) = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

querysql = "     SELECT a.CUSTOMER AS 客户代码,a.DEVICE AS 机种 ,a.LOT AS 批号,a.QTY AS 片数,a.WAFER_ID ,CASE WHEN  a.BONDED = 1 THEN '保税' ELSE '非保税' END 保税 ,a.IN_QTY AS 入库数 " & "     ,a.CREATE_DATE AS 来料信息上传时间 ,a.CREATE_BY AS 来料信息上传人 ,a.LAST_UPDATE_DATE AS 来料入库人 ,a.REMARK1 AS 入库单号 " & "      FROM erptemp..to_be_tested a WHERE a.FLAG <> 0  and a.CUSTOMER =   '" & Trim(ComCustomer.text) & "' AND a.CREATE_DATE  > '2019-02-01' "
If Trim(txtText3.text) <> "" Then
    querysql = querysql + " AND a.DEVICE  = '" & Trim(txtText3.text) & "'  "

End If

If Trim(txtText4.text) <> "" Then
    querysql = querysql + " AND a.LOT = '" & Trim(txtText4.text) & "'  "

End If

fpS(1).MaxRows = 0
If rs.State = adStateOpen Then rs.Close
rs.Open querysql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    Call ListDataType(rs)
Else
    MsgBox "无数据", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub ListDataType(rs As ADODB.Recordset)
Dim i As Long

With fpS(0)
    .MaxRows = 0
    Set .DataSource = rs

End With

With fpS(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        .ColWidth(1) = 2
        .CellType = CellTypeCheckBox
        .text = 1
        .Lock = False
        .Col = 7
        .Lock = False
        .BackColor = &HFFFF&
    Next

End With

End Sub

Private Sub ListDataType1(rs As ADODB.Recordset)
Dim i As Long
Dim j As Long
Dim strmatno As String
Dim strbond As String
Dim rs_new       As New ADODB.Recordset
With fpS(1)
    .MaxRows = 0
    Set .DataSource = rs
    
    .Col = 12
    .Row = -1
    .Lock = False
  '  .CellType = CellTypeComboBox
    

    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 6
        strmatno = Trim(.text)
        .Col = 11
        strbond = Trim(.text)
        .Col = 12
        If strbond = "保税" Then
            Set rs_new = Get_SqlserveRs("select distinct 采购单编号 from erpbase..OPENPO_WAFER where isnull(清关日期,'')<>'' and 未到货数量>0 and  物料编号='" & strmatno & "'")
            
            If rs_new.RecordCount > 0 Then
                .CellType = CellTypeComboBox
                rs_new.MoveFirst
                For j = 1 To rs_new.RecordCount
                   .TypeComboBoxList = .TypeComboBoxList & rs_new("采购单编号")
                   rs_new.MoveNext
                   
                Next
            End If
        End If
            
    Next
    
    
    

End With

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)
Dim i As Long

With fpS(2)
    .MaxRows = 0
    Set .DataSource = rs
    
    '.MaxCols = .MaxCols + 1
     For i = 1 To .MaxRows
         .Row = i
         .Col = 1  '选择
      
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .Lock = False
     Next
    
End With

End Sub

Private Sub ListDataType3(rs As ADODB.Recordset)
Dim i As Long

With fpS(2)
    .MaxRows = 0
    Set .DataSource = rs
    
    
End With

End Sub


Private Sub cmdup_Click(Index As Integer)

If ComCustomer.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
End If

If chkMsgAppend.Value = 1 And Trim(txtMsg.text) = "" Then
    MsgBox "您已勾选了邮件补充, 请填写正文 补充内容" & vbCrLf & "否则请取消勾选再上传", vbInformation, "提示"
    Exit Sub
End If

CommonDialog1.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
CommonDialog1.ShowOpen
If CommonDialog1.filename = "" Then
    Exit Sub

End If

txtPath.text = CommonDialog1.filename
CommonDialog1.filename = ""
If txtPath.text = "" Then
    MsgBox "请选择要上传的文件", vbInformation, "提示"
    Exit Sub

End If

Call Upload_0

End Sub

Private Sub Upload_0()

    On Error GoTo ErrHandle

    Dim VBExcel        As Excel.Application
    Dim xlBook         As Excel.Workbook
    Dim xlSheet        As Excel.Worksheet
    Dim strCust        As String
    Dim strdevice      As String
    Dim strKey         As String
    Dim strqty         As Long
    Dim strWafer       As String
    Dim strBonded      As String
    Dim StrCompmpmpmp  As String
    Dim User           As String
    Dim i              As Integer
    Dim rs             As New ADODB.Recordset
    Dim strsql         As String
    Dim waferidNoTemp  As String
    Dim bidWaferID()   As String
    Dim waferStrTemp   As String
    Dim k              As Integer
    Dim N              As Integer
    Dim kk             As Integer
    Dim comparewaferid As String
    Dim mpsid          As Long
    Dim strpo          As String
    Dim strFabDevice   As String

    Dim Remark4_str          As String

    Cnn.BeginTrans
    INIadoCon.BeginTrans
    User = gUserName
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.text)
    Set xlSheet = xlBook.Worksheets(1)

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 8 Then
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Cnn.RollbackTrans
        INIadoCon.RollbackTrans
        Exit Sub

    End If
 
    fpS(0).MaxRows = 0
    Remark4_str = GetID()
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strCust = Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "")
        strdevice = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
        strKey = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "")
        strqty = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        strWafer = Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "")
        strBonded = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "")
        strpo = Replace(Trim(xlSheet.Range("G" & i)), Chr(13) + Chr(10), "")
        strFabDevice = Replace(Trim(xlSheet.Range("H" & i)), Chr(13) + Chr(10), "")
        

        If (Not JudgeMPSBankPT(strdevice)) Then
            MsgBox "这机种：" & strdevice & " 在系统设定表中不存在，请联系市场部与研发部!", vbInformation, "友情提示"
            Close #1
            Cnn.RollbackTrans
            INIadoCon.RollbackTrans
            Exit Sub

        End If

        If InStr(strWafer, ".") > 0 Then
            bidWaferID = Split(strWafer, ".")

            If UBound(bidWaferID) + 1 <> Val(strqty) Then
                MsgBox "Batch号为" & (strKey) & "的条目, 片数与WaferID个数不一致, 请重新调整本次未上传任何batch", vbInformation, "友情提示"
                Close #1
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub

            End If

            For k = 0 To UBound(bidWaferID)

                '根据po号，batch号，判断是否已经上传过
                If (Judge37FabData(strKey, bidWaferID(k))) Then
                    MsgBox "LotID:" & strKey & " WaferID：" & bidWaferID(k) & " 已存在，无需上传!", vbInformation, "友情提示"
                    Close #1
                    Cnn.RollbackTrans
                    INIadoCon.RollbackTrans
                    Exit Sub

                End If

            Next

            For N = 0 To UBound(bidWaferID)
                waferidNoTemp = bidWaferID(N)

                If waferidNoTemp = "" Then
                    MsgBox "WaferId存在空值"
                    Close #1
                    Cnn.RollbackTrans
                    INIadoCon.RollbackTrans
                    Exit Sub
                ElseIf Val(waferidNoTemp) > 25 Or Val(waferidNoTemp) < 1 Then
                    MsgBox "WaferId超出1-25范围"
                    Close #1
                    Cnn.RollbackTrans
                    INIadoCon.RollbackTrans
                    Exit Sub

                End If

                For kk = N + 1 To UBound(bidWaferID)
                    comparewaferid = bidWaferID(kk)

                    If comparewaferid = waferidNoTemp Then
                        MsgBox "WaferId有重复"
                        Close #1
                        Cnn.RollbackTrans
                        INIadoCon.RollbackTrans
                        Exit Sub

                    End If

                Next
            Next
        Else

            If (Judge37FabData(strKey, strWafer)) Then
                MsgBox "这笔：" & strWafer & " 已存在，无需上传!", vbInformation, "友情提示"
                Close #1
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub

            End If

            If Val(strqty) <> 1 Then
                MsgBox "Batch号为" & (strKey) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                Close #1
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub

            End If

            If strWafer = "" Then
                MsgBox "Batch号为" & (strKey) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                Close #1
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub
            ElseIf Val(strWafer) > 25 Or Val(strWafer) < 1 Then
                MsgBox "Batch号为" & (strKey) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                Close #1
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub

            End If

        End If
     
        If strCust = "68" Or strCust = "HK006" Or strCust = "BJ128" Or strCust = "SC081" Then
            mpsid = Get37FabMaxID()
            AddSql (" insert into MAPPINGDATA37 (devicename,batch,wf,PRICE,currency,shippeddt,purchaseno,Purchaseorderlineitem,wafer_id,flag,qtech_created_by,qtech_created_date ,id,Status,Customershortname ) " & " values ('" & strdevice & "','" & strKey & "','" & strqty & "','95','USD',sysdate,'NA','NA','" & strWafer & "' ,'Y', '" & User & "',sysdate , '" & mpsid & "'  ,'0','" & strCust & "' ) ")

            If InStr(strWafer, ".") > 0 Then
                bidWaferID = Split(strWafer, ".")

                For N = 0 To UBound(bidWaferID)
                    waferidNoTemp = bidWaferID(N)

                    If Len(waferidNoTemp) <> 2 Then
                        waferStrTemp = strKey & "0" & waferidNoTemp
                    Else
                        waferStrTemp = strKey & waferidNoTemp

                    End If

                    AddSql ("insert into mappingDataTest(substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,id,filename )" & " values( '" & waferStrTemp & "','" & strKey & "'," & strqty & ",0,'Y','" & User & "',sysdate,'" & waferidNoTemp & "','" & strCust & "',mappingData_SEQ.Nextval,'')")
                    AddSql2 ("insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,filename )" & " values( '" & waferStrTemp & "','" & strKey & "'," & strqty & ",0,'Y','" & User & "',getdate(),'" & waferidNoTemp & "','" & strCust & "','')")
                Next
            Else
                waferidNoTemp = strWafer

                If Len(waferidNoTemp) <> 2 Then
                    waferStrTemp = strKey & "0" & waferidNoTemp

                End If

                waferStrTemp = strKey & waferidNoTemp
                AddSql ("insert into mappingDataTest(substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,id,filename )" & " values( '" & waferStrTemp & "','" & strKey & "'," & strqty & ",0,'Y','" & User & "',sysdate,'" & waferidNoTemp & "','" & strCust & "',mappingData_SEQ.Nextval,'')")
                AddSql2 ("insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,filename )" & " values( '" & waferStrTemp & "','" & strKey & "'," & strqty & ",0,'Y','" & User & "',getdate(),'" & waferidNoTemp & "','" & strCust & "','')")

            End If

        End If


        'REMARK5新增定义FABDEVICE
        'REMARK3新增定义PO
        '20191104 Merry REMARK4新增定义上传编号
        Dim NN As Integer
        Dim Flag_temp As Integer
                                        
        If Get_SqlserverCnt("SELECT c.*  FROM  erpbase..tblCustomerOI  a INNER JOIN ERPBASE ..tblstocknum  c ON c.批号=a.SOURCE_BATCH_ID WHERE a.FAB_CONV_ID='" & strKey & "'  AND c.仓库编号<>54 ") Then
            If Get_SqlserverCnt("SELECT c.*  FROM  erpbase..tblCustomerOI  a INNER JOIN ERPBASE ..tblstocknum  c ON c.批号=a.SOURCE_BATCH_ID WHERE a.FAB_CONV_ID='" & strKey & "'  AND c.仓库编号<>54 AND c.当前存量>0") Then
            '存在于大仓,且大仓有库存，状态传2
                Flag_temp = 2
            Else
                '存在于大仓,且大仓无库存，3个表都不传
                MsgBox "第" & i & "数据" & strWafer & "已存在大仓，不用再上传!", vbInformation, "提示"
                Cnn.RollbackTrans
                INIadoCon.RollbackTrans
                Exit Sub
            End If
        Else
            Flag_temp = 0
        End If
                
        If Get_SqlserverCnt(" SELECT * FROM erptemp..to_be_tested WHERE customer = '" & strCust & " ' AND device = '" & strdevice & " ' AND lot = '" & strKey & " 'AND wafer_id = '" & strWafer & " ' ") = 0 Then
            AddSql2 (" INSERT INTO erptemp..to_be_tested (CUSTOMER,device,LOT,QTY,WAFER_ID,BONDED,FLAG,CREATE_DATE,CREATE_BY,REMARK2,remark5,remark4 ) VALUES " & "  ('" & strCust & "','" & strdevice & "','" & strKey & "','" & strqty & "','" & strWafer & "','" & strBonded & "'," & Flag_temp & ",GETDATE(),'" & User & "','" & strpo & "','" & strFabDevice & "','" & Remark4_str & "') ")
           
          
            If InStr(strWafer, ".") > 0 Then
                bidWaferID = Split(strWafer, ".")

                For NN = 0 To UBound(bidWaferID)
                    waferidNoTemp = bidWaferID(NN)

                    If Len(waferidNoTemp) <> 2 Then
                        waferStrTemp = "0" & waferidNoTemp
                    Else
                        waferStrTemp = waferidNoTemp
                    End If
                    If Get_SqlserverCnt(" SELECT * FROM erptemp..to_be_tested_SUB WHERE customer = '" & strCust & " ' AND device = '" & strdevice & " ' AND lot = '" & strKey & " 'AND wafer_id = '" & waferStrTemp & " ' ") = 0 Then
                        AddSql2 (" INSERT INTO erptemp..to_be_tested_SUB (CUSTOMER,device,LOT,WAFER_ID, flag,REMARK4 ) VALUES " & "  ('" & strCust & "','" & strdevice & "','" & strKey & "','" & waferStrTemp & "'," & Flag_temp & ",'" & Remark4_str & "') ")
                    End If
                    
                Next
            Else
                waferidNoTemp = strWafer

                If Len(waferidNoTemp) <> 2 Then
                    waferStrTemp = "0" & waferidNoTemp
                Else
                    waferStrTemp = waferidNoTemp
                End If
                
                If Get_SqlserverCnt(" SELECT * FROM erptemp..to_be_tested_SUB WHERE customer = '" & strCust & " ' AND device = '" & strdevice & " ' AND lot = '" & strKey & " 'AND wafer_id = '" & waferStrTemp & " ' ") = 0 Then
                    AddSql2 (" INSERT INTO erptemp..to_be_tested_SUB (CUSTOMER,device,LOT,WAFER_ID,flag,REMARK4 ) VALUES " & "  ('" & strCust & "','" & strdevice & "','" & strKey & "','" & waferStrTemp & "'," & Flag_temp & ",'" & Remark4_str & "') ")
                End If
            End If
        Else
    
            MsgBox "第" & i & "数据" & strWafer & "已上传过，请确认!", vbInformation, "提示"
            Cnn.RollbackTrans
            INIadoCon.RollbackTrans

            Exit Sub

        End If

    Next
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    MsgBox "上传完成,根据客户代码查询后可选择生成入库单", vbInformation, "提示"

EXITUPLOAD:
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    ' Add:发送邮件通知
    Call SentMsgToCC(Trim(txtPath.text))
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0

    If Not VBExcel Is Nothing Then
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub

Private Sub SentMsgToCC(strFilePath As String)

Dim strSentTo(100) As String
Dim strSentCC(20)  As String
Dim strSentTitle   As String
Dim strSentText    As String
Dim dirtemp        As String
Dim strTemp        As String
Dim i              As Integer

If strFilePath = "" Then
    MsgBox "没有找到附件,本次邮件发送失败", vbExclamation, "提示"
    Exit Sub
End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_Bank.cfg"
strSentTitle = ComCustomer.text & "Bank wo 上传"
strSentText = "明细见附件:" & vbCrLf
strSentText = strSentText & txtMsg.text
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1
'strSentCC(0) = "xiaobing.yang_ks@ht-tech.com"
'strSentCC(1) = "angel.wu_ks@ht-tech.com"
If SentMes(strSentTitle, strSentText, strSentTo, strFilePath, strSentCC) = True Then
    MsgBox "邮件已发送", vbInformation, Me.Caption
Else
    MsgBox "邮件发送失败", vbCritical, Me.Caption

End If

End Sub

Private Sub cmd_Click(Index As Integer)
Unload Me

End Sub

Private Sub cmdCreate_Click()
If CheckData = True Then
    SaveData
End If

End Sub

Public Function CheckData() As Boolean
Dim i         As Integer
Dim rs        As New ADODB.Recordset
Dim strsql    As String
Dim strdevice As String
Dim strlot    As String

strdevice = ""
strlot = ""
strsql = " SELECT a.客户代码,b.供应商编号 FROM ERPBASE..tblXCustomer a," & "  ERPBASE..tblSupplierData  b  WHERE a.客户代码 = '" & ComCustomer.text & "' AND b.供应商名称 = a.客户名称 "
If rs.State = adStateOpen Then rs.Close
rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    txtText1.text = rs.Fields(1).Value
Else
    MsgBox "客户代码无关联的供应商编码", vbInformation, "提示"
    Exit Function

End If

CheckData = False
intCount = 0
ReDim strTemp(5)

With fpS(0)

    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If Trim$(.text) = "1" Then
        
            .Col = 3  '料号
            If Trim$(.text) <> "" Then
                
                .Col = 5  '数量
                If Val(Trim$(.text)) = 0 Then
                    MsgBox "第" & i & "行没有入库数量！", vbInformation, "提示"
                    Exit Function
    
                End If
    
                strTemp(1) = strTemp(1) + Trim$(.text) + "◆"
                strTemp(2) = strTemp(2) + Trim$(Now() + 100) + "◆"
                .Col = 4 '生产批号
                strTemp(3) = strTemp(3) + Trim$(.text) + "◆"
                '  .Col = 4 '到货批号
                strTemp(4) = strTemp(4) + Trim$(.text) + "◆"
                strlot = Trim$(.text)
                .Col = 7
                AddSql2 ("  UPDATE erptemp..TO_BE_TESTED  SET REMARK2 =  '" & Trim$(.text) & "'  WHERE CUSTOMER = '" & ComCustomer.text & "'  AND LOT = '" & strlot & "'  ")
                .Col = 12
                strdevice = Trim$(.text)
                strTemp(0) = strTemp(0) + Trim$(.text) + "◆"
                strTemp(5) = strTemp(5) + Trim$(.text) + "◆"
                intCount = intCount + 1
    
            End If
        End If

    Next

End With

If intCount = 0 Then
    MsgBox "没有选择任何物料信息！", vbInformation, "提示"
    Exit Function

End If
If intCount > 100 Then
    MsgBox "每次生成入库单不可超出100笔", vbInformation, "提示"
    Exit Function
End If
CheckData = True

End Function

Public Sub SaveData()

On Error GoTo ErrHandle

Dim adoprm1      As ADODB.Parameter
Dim adoprm2      As ADODB.Parameter
Dim adoPrm3      As ADODB.Parameter
Dim adoPrm4      As ADODB.Parameter
Dim adoPrm5      As ADODB.Parameter
Dim adoPrm6      As ADODB.Parameter
Dim adoPrm7      As ADODB.Parameter
Dim adoPrm8      As ADODB.Parameter
Dim adoPrm9      As ADODB.Parameter
Dim adoprm10     As ADODB.Parameter
Dim adoPrm11     As ADODB.Parameter
Dim adoPrm12     As ADODB.Parameter
Dim adoPrm13     As ADODB.Parameter
Dim adoprmFG     As ADODB.Parameter    '添加、修改,删除标记
Dim adoPrmReturn As ADODB.Parameter
Dim User         As String



User = gUserName
Dim CUSTOMER As String
Dim lot      As String
Dim device   As String
Dim WAFER    As String
Dim i        As Integer
Dim j        As Integer
Dim strWafer As String
Dim strno As String

Screen.MousePointer = 11
Set adoCmd = New ADODB.Command
Set adoCmd.ActiveConnection = INIadoCon
adoCmd.CommandText = "uspcg_toinrec_nocgd"
adoCmd.Parameters.Refresh
adoCmd.CommandType = adCmdStoredProc
Set adoPrmReturn = New ADODB.Parameter         '返回执行成功标记
adoPrmReturn.type = adInteger
adoPrmReturn.Direction = adParamReturnValue
adoCmd.Parameters.Append adoPrmReturn
Set adoprmFG = New ADODB.Parameter             '新增，修改，删除
adoprmFG.type = adInteger
adoprmFG.Direction = adParamInput
adoprmFG.Value = 1
adoCmd.Parameters.Append adoprmFG
Set adoprm1 = New ADODB.Parameter              '入库单编号
adoprm1.type = adVarChar
adoprm1.Size = 20
adoprm1.Direction = adParamInput
adoprm1.Value = ""
adoCmd.Parameters.Append adoprm1
Set adoprm2 = New ADODB.Parameter             '1：蓝字入库单
adoprm2.type = adInteger
adoprm2.Direction = adParamInput
adoprm2.Value = 1
adoCmd.Parameters.Append adoprm2
Set adoPrm3 = New ADODB.Parameter             '供应商编号
adoPrm3.type = adVarChar
adoPrm3.Size = 50
adoPrm3.Direction = adParamInput
adoPrm3.Value = Trim(txtText1.text)
adoCmd.Parameters.Append adoPrm3
Set adoPrm4 = New ADODB.Parameter             '仓库编号
adoPrm4.type = adVarChar
adoPrm4.Size = 20
adoPrm4.Direction = adParamInput
adoPrm4.Value = "54"
adoCmd.Parameters.Append adoPrm4
Set adoPrm5 = New ADODB.Parameter             '制单人
adoPrm5.type = adVarChar
adoPrm5.Size = 20
adoPrm5.Direction = adParamInput
adoPrm5.Value = Trim(User)
adoCmd.Parameters.Append adoPrm5
Set adoPrm6 = New ADODB.Parameter             '备注
adoPrm6.type = adVarChar
adoPrm6.Size = 200
adoPrm6.Direction = adParamInput
adoPrm6.Value = ""
adoCmd.Parameters.Append adoPrm6
Set adoPrm7 = New ADODB.Parameter             '保税标记
adoPrm7.type = adInteger
adoPrm7.Direction = adParamInput
adoPrm7.Value = "1"
adoCmd.Parameters.Append adoPrm7
Set adoPrm8 = New ADODB.Parameter             '数量
adoPrm8.type = adVarChar
adoPrm8.Size = 2000
adoPrm8.Direction = adParamInput
adoPrm8.Value = Trim(strTemp(1))
adoCmd.Parameters.Append adoPrm8
Set adoPrm9 = New ADODB.Parameter             '有效期
adoPrm9.type = adVarChar
adoPrm9.Size = 2000
adoPrm9.Direction = adParamInput
adoPrm9.Value = Trim(strTemp(2))
adoCmd.Parameters.Append adoPrm9
Set adoprm10 = New ADODB.Parameter             '生产批号
adoprm10.type = adVarChar
adoprm10.Size = 2000
adoprm10.Direction = adParamInput
adoprm10.Value = Trim(strTemp(3))
adoCmd.Parameters.Append adoprm10
Set adoPrm11 = New ADODB.Parameter             '到货批号
adoPrm11.type = adVarChar
adoPrm11.Size = 2000
adoPrm11.Direction = adParamInput
adoPrm11.Value = Trim(strTemp(4))
adoCmd.Parameters.Append adoPrm11
Set adoPrm12 = New ADODB.Parameter             '物料编号
adoPrm12.type = adVarChar
adoPrm12.Size = 2000
adoPrm12.Direction = adParamInput
adoPrm12.Value = Trim(strTemp(5))
adoCmd.Parameters.Append adoPrm12
Set adoPrm13 = New ADODB.Parameter             '循环次数
adoPrm13.type = adInteger
adoPrm13.Direction = adParamInput
adoPrm13.Value = intCount
adoCmd.Parameters.Append adoPrm13
'adoCmd.Execute
Screen.MousePointer = 0
If adoPrmReturn.Value = 0 Then
    Screen.MousePointer = 0
    '获取入库单号,写明细表
      strno = GetSqlServerStr("select top 1 rtrim(入库单编号) from erpbase..tbltoinrec where  仓库编号='54' and 制单='" & gUserName & "' order by 入库单编号 desc  ")

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If Trim(.text) = "1" Then
                
                .Col = 2
                CUSTOMER = .text
                .Col = 3
                device = .text
                .Col = 4
                lot = .text
                .Col = 6
                WAFER = .text
                AddSql2 ("  UPDATE erptemp..to_be_tested SET flag = 2 WHERE customer = '" & CUSTOMER & "' AND device = '" & device & "' AND lot = '" & lot & "'  AND wafer_id = '" & WAFER & "' and flag = 0 ")
                AddSql2 ("  UPDATE erptemp..to_be_tested_sub SET flag = 2 WHERE customer = '" & CUSTOMER & "' AND device = '" & device & "' AND lot = '" & lot & "' and flag = 0 ")
                For j = 0 To UBound(Split(WAFER, "."))
                    strWafer = Trim(Split(WAFER, ".")(j))
                    If Len(strWafer) = 1 Then
                        strWafer = "0" & strWafer
                    End If
                    AddSql2 ("insert into Erpbase..tblToInRec_Wafer_Wait (入库单编号,批号,晶圆ID,flag) VALUES('" & strno & "','" & lot & "','" & lot & strWafer & "',0)")
                Next
            End If
        Next

    End With

    MsgBox "已经成功执行您的任务！", vbInformation, Me.Caption
    QUERY_1
Else
    GoTo ErrHandle

End If

Exit Sub
ErrHandle:
Screen.MousePointer = 0
MsgBox "执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption

End Sub

Function GetID()
'remark4栏位，默认为ID
'生成方式：YYMMDD +3位流水码
Dim CODE       As String
Dim strsql     As String
Dim YearStr    As String
Dim MonthStr   As String
Dim DayStr     As String
Dim SMR        As New ADODB.Recordset


YearStr = Right(Year(Now()), 2)
If Len(Month(Now())) = 1 Then
    MonthStr = "0" & Month(Now())
Else
    MonthStr = Month(Now())
End If
If Len(Day(Now())) = 1 Then
    DayStr = "0" & Day(Now())
Else
    DayStr = Day(Now())
End If
CODE = YearStr & MonthStr & DayStr

strsql = "Select Isnull(max(remark4),'') as remark4 from erptemp..TO_BE_TESTED where left(remark4,6)='" & CODE & "'"

If SMR.State = adStateOpen Then SMR.Close
SMR.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If SMR("remark4") = "" Then
    GetID = CODE & "001"
Else
    GetID = Val(SMR("remark4")) + 1
End If
SMR.Close
Set SMR = Nothing

End Function


Private Sub cmdCommand3_Click()
    chkall.Visible = True
    QUERY_4

End Sub

Private Sub QUERY_4()
    Dim querysql As String
    Dim rs       As New ADODB.Recordset

    If Get_SqlserverCnt("select * from ERPBASE..tblstocknum where  批号 ='" & Trim(txtLOT.text) & "' and 当前存量 >0 ") > 1 Then
        MsgBox "晶圆分多次入待验仓，无法退运，请联系IT", vbInformation, "提示"
        cmdreject.Enabled = False
        Exit Sub
    End If
    querysql = "select 0 as '选择',b.CUSTOMER ,b.DEVICE ,a.批号,b.wafer_id,a.物料编号,a.供应商编号,a.当前存量,a.仓库编号 from   ERPBASE..tblstocknum a inner join erptemp..TO_BE_TESTED_sub  b on a.批号=b.LOT " & "where a.仓库编号=54 and a.当前存量>0 and a.批号='" & Trim(txtLOT.text) & "' AND isnull(b.FLAG,0)<>-1 "

    querysql = querysql & " order by b.wafer_id"
    
    fpS(2).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open querysql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType2(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Exit Sub

    End If
    
End Sub

Private Sub QUERY_5()
    Dim querysql As String
    Dim rs       As New ADODB.Recordset
      
    querysql = "select b.customer,b.DEVICE,b.LOT,a.WAFER_ID,b.WAFER_ID AS 退运waferid,a.create_date AS 上传BankWO时间,b.LAST_UPDATE_DATE AS 退运时间 ,b.LAST_UPDATE_BY AS 退运人员 from erptemp..TO_BE_TESTED a,erptemp..TO_BE_TESTED_sub b  where b.lot='" & Trim(txtLOT.text) & "' AND b.flag=-1 AND a.lot=b.lot"
    querysql = querysql & " order by b.wafer_id"
    
    fpS(2).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open querysql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType3(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Exit Sub

    End If
    
End Sub





Private Sub Update_OPENPOWAFER()


Dim strsql As String
Dim rs     As New ADODB.Recordset

On Error GoTo Err_Update
AddSql2 ("delete From erpbase..OPENPO_WAFER")
'同步采购表
strsql = "INSERT INTO erpbase..OPENPO_WAFER(采购单编号,采购单项次,物料编号,PO数量,到货数量,未到货数量,料号) SELECT a.采购单编号, a.采购单项次, a.物料编号,a.批准采购数量,0,0,c.F_101 FROM erpbase..tblCPurDataSub a, erpbase..tblCPurData b  ,AIS20141114094336..t_ICItem c WHERE a.采购单编号=b.采购单编号 AND a.物料编号=c.FNumber  and  a.采购单编号 like 'c%' and a.物料编号 LIKE '01.01%' and  b.保税标记=1 AND a.是否禁用=0"
AddSql2 (strsql)

'同步到货数量
strsql = "UPDATE a SET a.到货数量=isnull(t1.到货数量,0),a.未到货数量=a.PO数量-isnull(t1.到货数量,0) from erpbase..OPENPO_WAFER a  left JOIN ( SELECT b.采购单编号 ,b.采购单项次 ,sum(b.到货数量) AS 到货数量 FROM erpbase..tblToRecEntry b  GROUP BY b.采购单编号,b.采购单项次  ) AS t1 ON  a.采购单编号 =t1.采购单编号 AND a.采购单项次=t1.采购单项次"
AddSql2 (strsql)

'同步清关日期
strsql = "UPDATE a SET a.清关日期=replace(replace(isnull(b.入场日期,''),'月','/' ),'日','') from erpbase..OPENPO_WAFER a  left JOIN erptemp..ksimport b on a.采购单编号=b.采购单号 and a.料号=b.料号 where b.flag =0 "
AddSql2 (strsql)

Exit Sub
Err_Update:
MsgBox "Update_OPENPOWAFER 发生错误,原因：" & Err.DESCRIPTION, vbInformation, Me.Caption

End Sub




Private Sub Label11_Click()
Dim i  As Integer
Dim rs As ADODB.Recordset

Set rs = Get_SqlserveRs("select DISTINCT MARKETLASTUPDATE_BY from erptemp..tbltsvnpiproduct WHERE CUSTOMERPTNO1='" & Trim(txtCustomerDevice.text) & "'")
cbPn.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbPn.AddItem Trim(rs("MARKETLASTUPDATE_BY"))
        rs.MoveNext
    Next i

End If

rs.Close
Set rs = Nothing
End Sub

Private Sub txtCustomerDevice_Change()
cbCustomerID.Clear
cbCustomerID.text = ""
cbPn.Clear
cbPn.text = ""
cbBonded.text = ""
cbPONo.Clear


End Sub

Private Sub txtQty_Change()
Dim i As Integer
     With lsWaferID
        For i = 0 To .ListCount - 1
           .Selected(i) = False
        Next

    End With
            
    '数量25时全部勾选
    If IsNumeric(txtQty.text) = False Then
        txtQty.text = ""
    Else
        If Val(txtQty.text) > 25 Then
            MsgBox "数量不能大于25", vbInformation, "提示"
            txtQty.text = ""
        End If
        If Val(txtQty.text) = 25 Then
             With lsWaferID
                For i = 0 To .ListCount - 1
                   .Selected(i) = True
                Next
        
            End With
        End If
    End If
End Sub

Public Function SaveDataByStock()

On Error GoTo ErrHandle

Dim adoprm1      As ADODB.Parameter
Dim adoprm2      As ADODB.Parameter
Dim adoPrm3      As ADODB.Parameter
Dim adoPrm4      As ADODB.Parameter
Dim adoPrm5      As ADODB.Parameter
Dim adoPrm6      As ADODB.Parameter
Dim adoPrm7      As ADODB.Parameter
Dim adoPrm8      As ADODB.Parameter
Dim adoPrm9      As ADODB.Parameter
Dim adoprm10     As ADODB.Parameter
Dim adoPrm11     As ADODB.Parameter
Dim adoPrm12     As ADODB.Parameter
Dim adoPrm13     As ADODB.Parameter
Dim adoprmFG     As ADODB.Parameter    '添加、修改,删除标记
Dim adoPrmReturn As ADODB.Parameter
Dim User         As String

User = gUserName
Dim CUSTOMER As String
Dim lot      As String
Dim device   As String
Dim WAFER    As String
Dim i        As Integer
SaveDataByStock = False
Screen.MousePointer = 11
Set adoCmd = New ADODB.Command
Set adoCmd.ActiveConnection = INIadoCon
adoCmd.CommandText = "uspcg_toinrec_nocgd"
adoCmd.Parameters.Refresh
adoCmd.CommandType = adCmdStoredProc
Set adoPrmReturn = New ADODB.Parameter         '返回执行成功标记
adoPrmReturn.type = adInteger
adoPrmReturn.Direction = adParamReturnValue
adoCmd.Parameters.Append adoPrmReturn
Set adoprmFG = New ADODB.Parameter             '新增，修改，删除
adoprmFG.type = adInteger
adoprmFG.Direction = adParamInput
adoprmFG.Value = 1
adoCmd.Parameters.Append adoprmFG
Set adoprm1 = New ADODB.Parameter              '入库单编号
adoprm1.type = adVarChar
adoprm1.Size = 20
adoprm1.Direction = adParamInput
adoprm1.Value = ""
adoCmd.Parameters.Append adoprm1
Set adoprm2 = New ADODB.Parameter             '1：蓝字入库单
adoprm2.type = adInteger
adoprm2.Direction = adParamInput
adoprm2.Value = 1
adoCmd.Parameters.Append adoprm2
Set adoPrm3 = New ADODB.Parameter             '供应商编号
adoPrm3.type = adVarChar
adoPrm3.Size = 50
adoPrm3.Direction = adParamInput
adoPrm3.Value = Trim(txtSupplierno.text)
adoCmd.Parameters.Append adoPrm3
Set adoPrm4 = New ADODB.Parameter             '仓库编号
adoPrm4.type = adVarChar
adoPrm4.Size = 20
adoPrm4.Direction = adParamInput
adoPrm4.Value = "54"
adoCmd.Parameters.Append adoPrm4
Set adoPrm5 = New ADODB.Parameter             '制单人
adoPrm5.type = adVarChar
adoPrm5.Size = 20
adoPrm5.Direction = adParamInput
adoPrm5.Value = Trim(User)
adoCmd.Parameters.Append adoPrm5
Set adoPrm6 = New ADODB.Parameter             '备注
adoPrm6.type = adVarChar
adoPrm6.Size = 200
adoPrm6.Direction = adParamInput
adoPrm6.Value = ""
adoCmd.Parameters.Append adoPrm6
Set adoPrm7 = New ADODB.Parameter             '保税标记
adoPrm7.type = adInteger
adoPrm7.Direction = adParamInput
adoPrm7.Value = "1"
adoCmd.Parameters.Append adoPrm7
Set adoPrm8 = New ADODB.Parameter             '数量
adoPrm8.type = adVarChar
adoPrm8.Size = 2000
adoPrm8.Direction = adParamInput
adoPrm8.Value = Trim(strTemp(1))
adoCmd.Parameters.Append adoPrm8
Set adoPrm9 = New ADODB.Parameter             '有效期
adoPrm9.type = adVarChar
adoPrm9.Size = 2000
adoPrm9.Direction = adParamInput
adoPrm9.Value = Trim(strTemp(2))
adoCmd.Parameters.Append adoPrm9
Set adoprm10 = New ADODB.Parameter             '生产批号
adoprm10.type = adVarChar
adoprm10.Size = 2000
adoprm10.Direction = adParamInput
adoprm10.Value = Trim(strTemp(3))
adoCmd.Parameters.Append adoprm10
Set adoPrm11 = New ADODB.Parameter             '到货批号
adoPrm11.type = adVarChar
adoPrm11.Size = 2000
adoPrm11.Direction = adParamInput
adoPrm11.Value = Trim(strTemp(4))
adoCmd.Parameters.Append adoPrm11
Set adoPrm12 = New ADODB.Parameter             '物料编号
adoPrm12.type = adVarChar
adoPrm12.Size = 2000
adoPrm12.Direction = adParamInput
adoPrm12.Value = Trim(strTemp(5))
adoCmd.Parameters.Append adoPrm12
Set adoPrm13 = New ADODB.Parameter             '循环次数
adoPrm13.type = adInteger
adoPrm13.Direction = adParamInput
adoPrm13.Value = intCount
adoCmd.Parameters.Append adoPrm13
adoCmd.Execute
Screen.MousePointer = 0
If adoPrmReturn.Value = 0 Then
    Screen.MousePointer = 0
    SaveDataByStock = True
Else
    GoTo ErrHandle

End If

Exit Function
ErrHandle:
Screen.MousePointer = 0
MsgBox "SaveDataByStock执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption

End Function








