VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmHDLVS 
   Caption         =   "HD标签核对系统"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12870
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
   ScaleHeight     =   8535
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   630
         Width           =   6015
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5895
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   10398
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
         MaxCols         =   6
         MaxRows         =   0
         SpreadDesigner  =   "FrmHDLVS.frx":0000
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描二维码"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmHDLVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

