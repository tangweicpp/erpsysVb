VERSION 5.00
Begin VB.Form FrmLblCheck37 
   Caption         =   "标签核对系统_37"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15405
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
   ScaleHeight     =   11190
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.TextBox txtTrayValue 
         BackColor       =   &H00FFC0FF&
         Height          =   2310
         Left            =   7200
         TabIndex        =   26
         Top             =   5640
         Width           =   5295
      End
      Begin VB.TextBox txtInnerBoxValue 
         BackColor       =   &H00FFC0FF&
         Height          =   2310
         Left            =   7200
         TabIndex        =   24
         Top             =   3120
         Width           =   5295
      End
      Begin VB.TextBox txtOuterBoxValue 
         BackColor       =   &H00FFC0FF&
         Height          =   2310
         Left            =   7200
         TabIndex        =   22
         Top             =   698
         Width           =   5295
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出核对记录"
         Height          =   360
         Left            =   1080
         TabIndex        =   20
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出核对"
         Height          =   360
         Left            =   2640
         TabIndex        =   19
         Top             =   7800
         Width           =   990
      End
      Begin VB.TextBox txtCurTrayID 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   6098
         Width           =   3855
      End
      Begin VB.TextBox txtCurInnerBoxNum 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   5738
         Width           =   3855
      End
      Begin VB.TextBox txtCurOuterBoxNum 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   5378
         Width           =   3855
      End
      Begin VB.TextBox txtTrayID 
         BackColor       =   &H00FFC0FF&
         Height          =   2205
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2880
         Width           =   3855
      End
      Begin VB.TextBox txtInnerBoxNum 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txtOuterBoxNum 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtShipTo 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   997
         Width           =   3855
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   675
         Width           =   3855
      End
      Begin VB.Label lblTrayValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘数据"
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
         Left            =   6240
         TabIndex        =   25
         Top             =   5640
         Width           =   915
      End
      Begin VB.Label lblInnerBoxValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱数据"
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
         Left            =   6240
         TabIndex        =   23
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label lblOuterBoxValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱数据"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblCurTrayID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前卷盘"
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
         TabIndex        =   17
         Top             =   6120
         Width           =   930
      End
      Begin VB.Label lblCurInnerBoxNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前内箱"
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
         TabIndex        =   15
         Top             =   5760
         Width           =   930
      End
      Begin VB.Label lblCurOuterBoxNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱"
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
         TabIndex        =   13
         Top             =   5400
         Width           =   930
      End
      Begin VB.Label lblTrayID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘"
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
         Left            =   600
         TabIndex        =   11
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label lblInnerBoxNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱"
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
         Left            =   600
         TabIndex        =   8
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label lblOuterBoxNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱"
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
         Left            =   600
         TabIndex        =   7
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label lblShipTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHIP TO"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN"
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
         TabIndex        =   3
         Top             =   982
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框"
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
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmLblCheck37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub
