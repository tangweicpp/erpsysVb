VERSION 5.00
Begin VB.Form FrmSTQty 
   BackColor       =   &H00C0C0C0&
   Caption         =   "标签核对系统 (LVS)"
   ClientHeight    =   11730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14625
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11730
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frm_TrayP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "卷盘标签"
      Height          =   3375
      Index           =   1
      Left            =   240
      TabIndex        =   37
      Top             =   6960
      Width           =   14055
      Begin VB.TextBox txt_TrST_Qty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2122
         Width           =   1215
      End
      Begin VB.TextBox txt_TrSS_Qty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2122
         Width           =   1215
      End
      Begin VB.TextBox txt_TR_Total 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtTrPkg_SS 
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtTrStatus 
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   9360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton optTrPkg_ST 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3600
         TabIndex        =   41
         Top             =   315
         Width           =   495
      End
      Begin VB.OptionButton optTrPkg_SS 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7680
         TabIndex        =   40
         Top             =   405
         Width           =   255
      End
      Begin VB.TextBox txtTrPkg_ST 
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtScan3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   38
         Top             =   2722
         Width           =   3735
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   12960
         Top             =   1560
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10582
         TabIndex        =   54
         Top             =   360
         Width           =   570
      End
      Begin VB.Label LblOutStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   53
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "明细:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   52
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   8640
         TabIndex        =   51
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   50
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEMTECH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2377
         TabIndex        =   49
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAMSUNG:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6307
         TabIndex        =   48
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Top             =   2760
         Width           =   1050
      End
   End
   Begin VB.Frame Frm_InnerP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "内盒标签"
      Height          =   3375
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   14055
      Begin VB.TextBox txt_IPST_Qty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2122
         Width           =   1215
      End
      Begin VB.TextBox txt_IPSS_Qty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2122
         Width           =   1215
      End
      Begin VB.TextBox txt_IP_Total 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtInPkg_SS 
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtInStatus 
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   9360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton optInPkg_ST 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   315
         Width           =   495
      End
      Begin VB.OptionButton optInPkg_SS 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7680
         TabIndex        =   22
         Top             =   405
         Width           =   255
      End
      Begin VB.TextBox txtInPkg_ST 
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtScan2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   20
         Top             =   2722
         Width           =   3735
      End
      Begin VB.Timer Timer2 
         Interval        =   1200
         Left            =   12960
         Top             =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10582
         TabIndex        =   36
         Top             =   360
         Width           =   570
      End
      Begin VB.Label LblOutStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   35
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "明细:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   8640
         TabIndex        =   33
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEMTECH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2377
         TabIndex        =   31
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAMSUNG:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6307
         TabIndex        =   30
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF8080&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10560
      Width           =   3375
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0FF&
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10560
      Width           =   3255
   End
   Begin VB.Frame Frm_OutP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "外箱标签"
      Height          =   3375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.TextBox txtOP_SS_JOBNO 
         Height          =   975
         Left            =   7320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   2153
         Width           =   1215
      End
      Begin VB.TextBox txtOP_SS_Qty 
         Height          =   285
         Left            =   8160
         TabIndex        =   56
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtOutPkg_SS 
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   720
         Width           =   3495
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   12960
         Top             =   1560
      End
      Begin VB.TextBox txtScan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   2722
         Width           =   3735
      End
      Begin VB.TextBox txtOutPkg_ST 
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton optOutPkg_SS 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7680
         TabIndex        =   12
         Top             =   405
         Width           =   255
      End
      Begin VB.OptionButton optOutPkg_ST 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   315
         Width           =   495
      End
      Begin VB.TextBox txtOutStatus 
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   9360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txt_OP_Total 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt_OPSS_Qty 
         Height          =   960
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_OPST_Qty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2208
         Width           =   1215
      End
      Begin VB.Label lblJOBNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOBNO:"
         Height          =   195
         Left            =   6600
         TabIndex        =   58
         Top             =   2243
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAMSUNG(分):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   16
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAMSUNG(总):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Top             =   2246
         Width           =   570
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   8640
         TabIndex        =   6
         Top             =   2198
         Width           =   570
      End
      Begin VB.Label lblOPDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "明细:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label LblOutStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   840
         Width           =   45
      End
      Begin VB.Label LblOPStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10582
         TabIndex        =   3
         Top             =   360
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmSTQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim tOPLS As tOutPkgLblStatus   ' 外箱核对状态
Dim tIPLS As tOutPkgLblStatus   ' 内箱核对状态
Dim tTRLS As tOutPkgLblStatus

Dim tOPLD_ST As tOutPkgLblData  ' Semtech外箱标签数据
Dim tOPLD_SS(10) As tOutPkgLblData  ' Samsung外箱标签数据

Dim tIPLD_ST(10) As tInPkgLblData   ' Semtech内箱标签
Dim tIPLD_SS(10) As tInPkgLblData   ' Samsung内箱标签

Dim tTRLD_ST(10) As tInPkgLblData
Dim tTRLD_SS(10) As tInPkgLblData

Dim iSS As Integer      ' 外箱
Dim iSS_2 As Integer    ' 内箱
Dim iSS_3 As Integer    ' 卷盘

Private Sub ResetData()

tOPLS.bSS = False
tOPLS.bST = False
    
tOPLD_ST.sInvoice = ""
tOPLD_ST.sCustomerPartNo = ""
tOPLD_ST.sMfgPartNo = ""
tOPLD_ST.sPurchaseOrder = ""
tOPLD_ST.sQty = ""

txt_OP_Total.Text = "0"

For i = 0 To 9
    tOPLD_SS(i).sInvoice = ""
    tOPLD_SS(i).sCustomerPartNo = ""
    tOPLD_SS(i).sMfgPartNo = ""
    tOPLD_SS(i).sPurchaseOrder = ""
    tOPLD_SS(i).sQty = ""
    tOPLD_SS(i).sLotNo = ""
Next

End Sub
Private Sub ResetData2()

tIPLS.bSS = False
tIPLS.bST = False

txt_IP_Total.Text = "0"

For i = 0 To 9
    tIPLD_ST(i).sJobNo = ""
    tIPLD_ST(i).sLotNumber = ""
    tIPLD_ST(i).sMFG = ""
    tIPLD_ST(i).sPartNo = ""
    tIPLD_ST(i).sQty = ""
    tIPLD_ST(i).sStrInfo = ""

    tIPLD_SS(i).sJobNo = ""
    tIPLD_SS(i).sLotNumber = ""
    tIPLD_SS(i).sMFG = ""
    tIPLD_SS(i).sPartNo = ""
    tIPLD_SS(i).sQty = ""
    tIPLD_SS(i).sStrInfo = ""
Next

End Sub

Private Sub ResetData3()

tTRLS.bSS = False
tTRLS.bST = False

txt_TR_Total.Text = "0"

For i = 0 To 9
    tTRLD_ST(i).sJobNo = ""
    tTRLD_ST(i).sLotNumber = ""
    tTRLD_ST(i).sMFG = ""
    tTRLD_ST(i).sPartNo = ""
    tTRLD_ST(i).sQty = ""
    tTRLD_ST(i).sStrInfo = ""

    tTRLD_SS(i).sJobNo = ""
    tTRLD_SS(i).sLotNumber = ""
    tTRLD_SS(i).sMFG = ""
    tTRLD_SS(i).sPartNo = ""
    tTRLD_SS(i).sQty = ""
    tTRLD_SS(i).sStrInfo = ""
Next

End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub cmdReset_Click()

Unload Me
ResetData
FrmSTQty.Show

End Sub

Private Sub Form_Activate()

' 0.定位:txtScan
txtScan.SetFocus

' 1.状态:未核对
tOPLS.bSS = False
tOPLS.bST = False

' 2.扫描对象初始化-> SemTech
optOutPkg_ST.Value = True
optOutPkg_SS.Value = False
iSS = 0
iSS_2 = 0
iSS_3 = 0
txt_OP_Total.Text = "0"
txt_IP_Total.Text = "0"
End Sub

Private Sub Form_Load()

' 1.状态初始化
Call InitStatus

End Sub

Private Sub InitStatus()

' 1.待验
If tOPLS.bST = False Or tOPLS.bST = False Then
    txtOutStatus.Text = "INVOICE+:待验" & vbCrLf & "PURCHASE:待验" & vbCrLf & "CUSTOMER_PART_NO:待验" & vbCrLf & "MFG:待验" & vbCrLf
Else
    
End If

' 2.核对完成

End Sub

Private Sub InitData()

If optOutPkg_ST.Value Then
    txtOutPkg_ST.Text = "(I)INVOICE+:  " & tOPLD_ST.sInvoice & vbCrLf & "(K)PURCHASE ORDER:  " & tOPLD_ST.sPurchaseOrder & vbCrLf & "(P)CUSTOMER PART NO:  " & tOPLD_ST.sCustomerPartNo & vbCrLf & "(Z)MFG PART NO:  " & tOPLD_ST.sMfgPartNo & vbCrLf
    txt_OPST_Qty.Text = Replace(tOPLD_ST.sQty, "Q", "")
Else
    txtOutPkg_SS.Text = "(I)INVOICE+:  " & tOPLD_SS(iSS).sInvoice & vbCrLf & "(K)PURCHASE ORDER:  " & tOPLD_SS(iSS).sPurchaseOrder & vbCrLf & "(P)CUSTOMER PART NO:  " & tOPLD_SS(iSS).sCustomerPartNo & vbCrLf & "(Z)MFG PART NO:  " & tOPLD_SS(iSS).sMfgPartNo & vbCrLf
    txt_OPSS_Qty.Text = Replace(tOPLD_SS(0).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(1).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(2).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(3).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(4).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(5).sQty, "Q", "") & vbCrLf & Replace(tOPLD_SS(6).sQty, "Q", "") & vbCrLf
    txtOP_SS_JOBNO.Text = Replace(tOPLD_SS(0).sLotNo, "P", "") & vbCrLf & Replace(tOPLD_SS(1).sLotNo, "Q", "") & vbCrLf & Replace(tOPLD_SS(2).sLotNo, "Q", "") & vbCrLf & Replace(tOPLD_SS(3).sLotNo, "Q", "") & vbCrLf & Replace(tOPLD_SS(4).sLotNo, "Q", "") & vbCrLf & Replace(tOPLD_SS(5).sLotNo, "Q", "") & vbCrLf & Replace(tOPLD_SS(6).sLotNo, "Q", "") & vbCrLf
End If

End Sub

Private Sub InitData2()

If optInPkg_ST.Value Then
    txtInPkg_ST.Text = "(1T)JOB NUMBER: " & tIPLD_ST(iSS_2).sJobNo & vbCrLf & "(1P)MFG P/N:  " & tIPLD_ST(iSS_2).sMFG & vbCrLf & "(S)LOT NUMBER:  " & tIPLD_ST(iSS_2).sLotNumber & vbCrLf & "(Q)QTY:  " & tIPLD_ST(iSS_2).sQty & vbCrLf
    txt_IPST_Qty.Text = Replace(tIPLD_ST(iSS_2).sQty, "Q", "")
Else
'    txtInPkg_SS.Text = tIPLD_SS(iSS_2).sStrInfo & vbCrLf & "PART NO: " & Left(tIPLD_SS(iSS_2).sStrInfo, InStr(tIPLD_SS(iSS_2).sStrInfo, "DPKT")) & vbCrLf
    txtInPkg_SS.Text = "PART NO: " & tIPLD_SS(iSS_2).sPartNo & vbCrLf
    txt_IPSS_Qty.Text = tIPLD_SS(iSS_2).sQty
    
End If

End Sub

Private Sub InitData3()

If optTrPkg_ST.Value Then
    txtTrPkg_ST.Text = "(1T)JOB NUMBER: " & tTRLD_ST(iSS_3).sJobNo & vbCrLf & "(1P)MFG P/N:  " & tTRLD_ST(iSS_3).sMFG & vbCrLf & "(S)LOT NUMBER:  " & tTRLD_ST(iSS_3).sLotNumber & vbCrLf & "(Q)QTY:  " & tTRLD_ST(iSS_3).sQty & vbCrLf
    txt_TrST_Qty.Text = Replace(tTRLD_ST(iSS_3).sQty, "Q", "")
Else
'    txtInPkg_SS.Text = tIPLD_SS(iSS_2).sStrInfo & vbCrLf & "PART NO: " & Left(tIPLD_SS(iSS_2).sStrInfo, InStr(tIPLD_SS(iSS_2).sStrInfo, "DPKT")) & vbCrLf
    txtTrPkg_SS.Text = "PART NO: " & tTRLD_SS(iSS_3).sPartNo & vbCrLf
    txt_TrSS_Qty.Text = tTRLD_SS(iSS_3).sQty
    
End If

End Sub

Private Sub optOutPkg_SS_Click()
txtScan.SetFocus
End Sub

Private Sub optOutPkg_ST_Click()
txtScan.SetFocus
End Sub

Private Sub Timer1_Timer()

Dim sScanData As String
Dim cScanDataHeader As String

sScanData = Trim(txtScan.Text)
cScanDataHeader = Left$(sScanData, 1)

' 0.数据生成
If optOutPkg_ST.Value Then   'SemTech
    ' 根据首字母判断
    Select Case cScanDataHeader
    Case "I"
        tOPLD_ST.sInvoice = Replace(sScanData, "I", "")
    Case "K"
        tOPLD_ST.sPurchaseOrder = Replace(sScanData, "K", "")
    Case "Z"
        tOPLD_ST.sMfgPartNo = Replace(sScanData, "Z", "")
    Case "P"
        If InStr(sScanData, "-") Then
           tOPLD_ST.sCustomerPartNo = Replace(sScanData, "P", "")
        End If
    Case "Q"
        tOPLD_ST.sQty = Replace(sScanData, "Q", "")
    Case Else
'        If InStr(sScanData, "-") Then
'            tOPLD_ST.sCustomerPartNo = Replace(sScanData, "P", "")
'        End If
    End Select
    
ElseIf optOutPkg_SS.Value Then                       'SamSung
        ' 根据首字母判断
    Select Case cScanDataHeader
    Case "I"
        tOPLD_SS(iSS).sInvoice = Replace(sScanData, "I", "")
    Case "K"
        tOPLD_SS(iSS).sPurchaseOrder = Replace(sScanData, "K", "")
    Case "Z"
        tOPLD_SS(iSS).sMfgPartNo = Replace(sScanData, "Z", "")
    Case "P"
        If InStr(sScanData, "-") Then
           tOPLD_SS(iSS).sCustomerPartNo = Replace(sScanData, "P", "")
        Else
            tOPLD_SS(iSS).sLotNo = Replace(sScanData, "P", "")
        End If
    Case "Q"
        tOPLD_SS(iSS).sQty = Replace(sScanData, "Q", "")
    Case Else

    End Select
    
Else
   
End If

InitData
CheckStatus

txtScan.Text = ""

End Sub

Private Sub Timer2_Timer()

Dim sScanData As String
Dim cScanDataHeader1 As String
Dim cScanDataHeader2 As String

sScanData = Trim(txtScan2.Text)
cScanDataHeader1 = Left$(sScanData, 1)
cScanDataHeader2 = Left$(sScanData, 2)

' 0.数据生成
If optInPkg_ST.Value Then   'SemTech
    ' 根据首字母判断
    Select Case cScanDataHeader1
    Case "S"
        tIPLD_ST(iSS_2).sLotNumber = Replace(sScanData, "S", "")
    Case "Q"
        tIPLD_ST(iSS_2).sQty = Replace(sScanData, "Q", "")
    
    Case Else
        
    End Select
    
    Select Case cScanDataHeader2
    Case "1T"
        ' jobNO
        tIPLD_ST(iSS_2).sJobNo = Replace(sScanData, "1T", "")
        
    Case "T"
         ' jobNO
        tIPLD_ST(iSS_2).sJobNo = Replace(sScanData, "T", "")
    Case "1P"
        ' 机种
        tIPLD_ST(iSS_2).sMFG = Replace(sScanData, "1P", "")
    Case "P"
         ' 机种
        tIPLD_ST(iSS_2).sMFG = Replace(sScanData, "P", "")
    
    Case Else
        
    End Select
    
ElseIf optInPkg_SS.Value Then                       'SamSung
    Timer2.Interval = 2000
    ' 根据首字母判断
    If InStr(sScanData, "0406") Then
        tIPLD_SS(iSS_2).sStrInfo = sScanData
        tIPLD_SS(iSS_2).sPartNo = Left(sScanData, InStr(sScanData, "DP") - 1)
        tIPLD_SS(iSS_2).sQty = Right$(sScanData, Len(sScanData) - InStr(sScanData, "E2") - 1)

    End If
Else
   

End If

InitData2
CheckStatus2

txtScan2.Text = ""

End Sub

Private Sub Timer3_Timer()
Dim sScanData As String
Dim cScanDataHeader1 As String
Dim cScanDataHeader2 As String

sScanData = Trim(txtScan3.Text)
cScanDataHeader1 = Left$(sScanData, 1)
cScanDataHeader2 = Left$(sScanData, 2)

' 0.数据生成
If optTrPkg_ST.Value Then   'SemTech
    ' 根据首字母判断
    Select Case cScanDataHeader1
    Case "S"
        tTRLD_ST(iSS_3).sLotNumber = Replace(sScanData, "S", "")
    Case "Q"
        tTRLD_ST(iSS_3).sQty = Replace(sScanData, "Q", "")
    
    Case Else
        
    End Select
    
    Select Case cScanDataHeader2
    Case "1T"
        ' jobNO
        tTRLD_ST(iSS_3).sJobNo = Replace(sScanData, "1T", "")
        
    Case "T"
         ' jobNO
        tTRLD_ST(iSS_3).sJobNo = Replace(sScanData, "T", "")
    Case "1P"
        ' 机种
        tTRLD_ST(iSS_3).sMFG = Replace(sScanData, "1P", "")
    Case "P"
         ' 机种
        tTRLD_ST(iSS_3).sMFG = Replace(sScanData, "P", "")
    
    Case Else
        
    End Select
    
ElseIf optTrPkg_SS.Value Then  'SamSung
    Timer3.Interval = 2000
    ' 根据首字母判断
    If InStr(sScanData, "0406") Then
        tTRLD_SS(iSS_3).sStrInfo = sScanData
        tTRLD_SS(iSS_3).sPartNo = Left(sScanData, InStr(sScanData, "DP") - 1)
        tTRLD_SS(iSS_3).sQty = Right$(sScanData, Len(sScanData) - InStr(sScanData, "E2") - 1)

    End If
Else
   
End If

InitData3
CheckStatus3

txtScan3.Text = ""

End Sub

Private Sub CheckStatus()

If tOPLS.bST = False And optOutPkg_SS.Value = False And tOPLD_ST.sInvoice <> "" And tOPLD_ST.sPurchaseOrder <> "" And tOPLD_ST.sCustomerPartNo <> "" And tOPLD_ST.sMfgPartNo <> "" And tOPLD_ST.sQty <> "" Then
    txtOutStatus.Text = "三星外箱总标签数据扫描完成" & vbCrLf & "---------------------------" & vbCrLf
    tOPLS.bST = True
    
End If

If tOPLS.bSS = False And tOPLD_SS(iSS).sInvoice <> "" And tOPLD_SS(iSS).sPurchaseOrder <> "" And tOPLD_SS(iSS).sCustomerPartNo <> "" And tOPLD_SS(iSS).sMfgPartNo <> "" And tOPLD_SS(iSS).sQty <> "" And tOPLD_SS(iSS).sLotNo <> "" Then
    txtOutStatus.Text = txtOutStatus.Text & "三星外箱分标签" & tOPLD_SS(iSS).sLotNo & "数据扫描完成" & vbCrLf & "---------------------------" & vbCrLf
    tOPLS.bSS = True
End If

If tOPLS.bST = True And tOPLS.bSS = False Then
    optOutPkg_SS.Value = True
    txtScan.SetFocus

ElseIf tOPLS.bSS = True And tOPLS.bST = False Then
    optOutPkg_ST.Value = True
    txtScan.SetFocus
    
ElseIf tOPLS.bST = True And tOPLS.bSS = True Then
    Timer1.Enabled = False

    Call CheckData
End If

End Sub

Private Sub CheckStatus2()

If tIPLS.bST = False And optInPkg_SS.Value = False And tIPLD_ST(iSS_2).sJobNo <> "" And tIPLD_ST(iSS_2).sMFG <> "" And tIPLD_ST(iSS_2).sLotNumber <> "" And tIPLD_ST(iSS_2).sQty <> "" Then
    txtInStatus.Text = "Semtech内箱JOB: " & tIPLD_ST(iSS_2).sJobNo & "数据扫描完成" & vbCrLf & "---------------------------" & vbCrLf
    tIPLS.bST = True
End If

If tIPLS.bSS = False And tIPLD_SS(iSS_2).sQty <> "" Then
    txtInStatus.Text = txtInStatus.Text & "SamSung内箱标签扫描完成 "
    tIPLS.bSS = True
End If

If tIPLS.bST = True And tIPLS.bSS = False Then
    optInPkg_SS.Value = True
    txtScan2.SetFocus

ElseIf tIPLS.bSS = True And tIPLS.bST = False Then
    optInPkg_ST.Value = True
    txtScan2.SetFocus
    
ElseIf tIPLS.bST = True And tIPLS.bSS = True Then
    Timer2.Enabled = False

    Call CheckData2
End If



End Sub

Private Sub CheckStatus3()

If tTRLS.bST = False And optTrPkg_SS.Value = False And tTRLD_ST(iSS_3).sJobNo <> "" And tTRLD_ST(iSS_3).sMFG <> "" And tTRLD_ST(iSS_3).sLotNumber <> "" And tTRLD_ST(iSS_3).sQty <> "" Then
    txtTrStatus.Text = "Semtech内箱JOB: " & tTRLD_ST(iSS_3).sJobNo & "数据扫描完成" & vbCrLf & "---------------------------" & vbCrLf
    tTRLS.bST = True
End If

If tTRLS.bSS = False And tTRLD_SS(iSS_3).sQty <> "" Then
    txtTrStatus.Text = txtTrStatus.Text & "SamSung卷盘标签扫描完成 "
    tTRLS.bSS = True
End If

If tTRLS.bST = True And tTRLS.bSS = False Then
    optTrPkg_SS.Value = True
    txtScan3.SetFocus

ElseIf tTRLS.bSS = True And tTRLS.bST = False Then
    optTrPkg_ST.Value = True
    txtScan3.SetFocus
    
ElseIf tTRLS.bST = True And tTRLS.bSS = True Then
    Timer3.Enabled = False

    Call CheckData3
End If



End Sub



Private Sub CheckData()

If tOPLD_ST.sInvoice <> tOPLD_SS(iSS).sInvoice Or tOPLD_ST.sPurchaseOrder <> tOPLD_SS(iSS).sPurchaseOrder Or tOPLD_ST.sCustomerPartNo <> tOPLD_SS(iSS).sCustomerPartNo Or tOPLD_ST.sMfgPartNo <> tOPLD_SS(iSS).sMfgPartNo Then
    txtOutStatus.ForeColor = vbRed
    txtOutStatus.Text = txtOutStatus.Text & "SAMSUNG外箱主标签DN,机种等信息与JOBNO:" & tOPLD_SS(iSS).sLotNo & "分标签不一致" & vbCrLf & " 请重新确认" & vbCrLf & "---------------------------" & vbCrLf
    
    ' 清空外箱数据
    ResetData
    txtScan.SetFocus

Else
    txtOutStatus.Text = txtOutStatus.Text & "SAMSUNG外箱主标签DN,机种等信息与JOBNO:" & tOPLD_SS(iSS).sLotNo & "分标签一致" & vbCrLf & "---------------------------" & vbCrLf
    txt_OP_Total.Text = CLng(txt_OP_Total.Text) + CLng(Replace(tOPLD_SS(iSS).sQty, "Q", ""))
    
    If CLng(txt_OPST_Qty) = CLng(txt_OP_Total.Text) Then
        txtOutStatus.Text = txtOutStatus.Text & "SAMSUNG外箱总数核对一致, 准备核对内箱" & vbCrLf & "---------------------------" & vbCrLf
        txtOutStatus.SelStart = Len(txtOutStatus)
        Timer1.Enabled = False
        txtScan2.SetFocus
        optInPkg_ST.Value = True
        
    Else
        txtOutStatus.Text = txtOutStatus.Text & "Next Job" & vbCrLf & "---------------------------" & vbCrLf
        txtOutStatus.SelStart = Len(txtOutStatus)
        
        tOPLS.bSS = False
        
        iSS = iSS + 1
        txtOP_SS_Qty.Text = iSS + 1
        InitData
        
        txtScan.SetFocus
        Timer1.Enabled = True
    End If
End If

End Sub

Private Sub CheckData2()

If tIPLD_ST(iSS_2).sQty <> tIPLD_SS(iSS_2).sQty Then
    txtInStatus.ForeColor = vbRed
    txtInStatus.Text = txtInStatus.Text & "Semtech,Samsung内箱标签数量不一致" & vbCrLf & " 请重新确认" & vbCrLf & "---------------------------" & vbCrLf
    
    ' 清空数据
    ResetData2
    txtScan2.SetFocus
Else
    txtInStatus.Text = txtInStatus.Text & "Semtech,Samsung内箱标签数量,机种一致, 准备核对卷盘" & vbCrLf & "---------------------------" & vbCrLf
    
    txt_IP_Total.Text = tIPLD_ST(iSS_2).sQty
    
    Timer2.Enabled = False
    txtScan3.SetFocus
    optTrPkg_ST.Value = True
End If

End Sub

Private Sub CheckData3()

If tTRLD_ST(iSS_3).sQty <> tTRLD_SS(iSS_3).sQty Then
    txtTrStatus.ForeColor = vbRed
    txtTrStatus.Text = txtTrStatus.Text & "Semtech,Samsung卷盘标签数量不一致" & vbCrLf & " 请重新确认" & vbCrLf & "---------------------------" & vbCrLf
    
    ' 清空数据
    ResetData3
    txtScan3.SetFocus
Else
    txtTrStatus.Text = txtTrStatus.Text & "Semtech,Samsung内箱标签数量,机种一致, 完成核对" & vbCrLf & "---------------------------" & vbCrLf
    
    txt_TR_Total.Text = tTRLD_ST(iSS_3).sQty
    
    Timer3.Enabled = False
'    txtScan3.SetFocus
'    optTrPkg_ST.Value = True
End If

End Sub



