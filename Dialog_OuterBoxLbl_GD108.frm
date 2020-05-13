VERSION 5.00
Begin VB.Form Dialog_OuterBoxLbl_GD108 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GD108外箱标签确认"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtOuterBoxCLblPrinter 
      Height          =   285
      Left            =   4800
      TabIndex        =   14
      Text            =   "ALL_OUT_2B1F_2"
      Top             =   2685
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFFF&
      Caption         =   "确认打印"
      Height          =   360
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "外箱标签预览"
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtSealDate 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   6000
         TabIndex        =   12
         Top             =   1830
         Width           =   2775
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   6000
         TabIndex        =   10
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtDateCode 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   1680
         TabIndex        =   8
         Top             =   1830
         Width           =   2775
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtLotNo 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFC0FF&
         Height          =   390
         Left            =   1680
         TabIndex        =   2
         Top             =   375
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seal Date:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4800
         TabIndex        =   11
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SN:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   5640
         TabIndex        =   9
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Code:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Number:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.Label lblPrinterCLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "外箱C标签打印机"
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   2730
      Width           =   1440
   End
End
Attribute VB_Name = "Dialog_OuterBoxLbl_GD108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Load()
txtPartNo.Text = FrmOuterPkgLblSys.txtPartNo.Text
txtLotNo.Text = FrmOuterPkgLblSys.txtLotNo.Text
txtDateCode.Text = FrmOuterPkgLblSys.txtDateCode.Text
txtQty.Text = FrmOuterPkgLblSys.txtQty.Text
txtSN.Text = FrmOuterPkgLblSys.txtSN.Text
txtSealDate.Text = FrmOuterPkgLblSys.txtSealDate.Text

End Sub

Private Sub cmdPrint_Click()

If txtPartNo.Text = "" Or txtLotNo.Text = "" Or txtDateCode.Text = "" Or txtQty.Text = "" Or txtSN.Text = "" Or txtSealDate.Text = "" Then
    MsgBox "标签信息不完整,无法打印,请确认好再打印", vbCritical, "警告"
    Exit Sub
End If

Call PrintOuterBoxLbl_GD108
Unload Me

End Sub
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintOuterBoxLbl_GD108
' Description:       打印外箱-C标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/14-9:55:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub PrintOuterBoxLbl_GD108()
Dim strSql     As String
Dim strEventID As String

strEventID = Trim$(txtSN.Text)

strSql = "INSERT INTO erpdata.dbo.tblME_PrintInfo(Createdate, Flag, EVENT_ID, PrinterNameID, BartenderName, PRINT_QTY, Content) values(GetDate(),'0','" & strEventID & "','" & Trim(txtOuterBoxCLblPrinter.Text) & "','GD108OUT2.btw','1',  " & " '""CUSTOMER_LOT""' + ',' + '""" & Trim$(txtLotNo.Text) & """' + ';' + '""GD108_DEVICE""' + ',' + '""" & Trim$(txtPartNo.Text) & """'+ ';'+ '""GD108_OUT_QTY""' + ',' + '""" & Trim$(txtQty.Text) & """'+ ';'+ '""GD108_YYWW_MON""' + ',' + '""" & Trim$(txtDateCode.Text) & """' + ';'+ '""LOT_WAFER_CODE_OUT""' + ',' + '""" & Trim$(txtSN.Text) & """'+ ';'+ '""PACKING_DATE1""' + ',' + '""" & Trim$(txtSealDate.Text) & """')  "
AddSql2 (strSql)

End Sub

