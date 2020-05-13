VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_Price_Control 
   Caption         =   "产品价格管理平台"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15645
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
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleMode       =   0  'User
   ScaleWidth      =   15645
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SST01 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   20135
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "产品价格维护"
      TabPicture(0)   =   "Frm_Price_Control.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl01"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl02"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl03"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl04"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDIE"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl06"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl07"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblNER"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Fps(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCust"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDevice"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtWafer"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDIE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbCombo1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbCombo2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdADD"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdmod"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdlos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdexport"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdquery"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtID1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txttestprice"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtNERQTY"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "OPENPO维护"
      TabPicture(1)   =   "Frm_Price_Control.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl001"
      Tab(1).Control(1)=   "lbl002"
      Tab(1).Control(2)=   "lblOPENPO"
      Tab(1).Control(3)=   "Fps(1)"
      Tab(1).Control(4)=   "txtOPcust"
      Tab(1).Control(5)=   "txtOPdevice"
      Tab(1).Control(6)=   "txtOPPO"
      Tab(1).Control(7)=   "cmd01"
      Tab(1).Control(8)=   "cmd02"
      Tab(1).Control(9)=   "txtID2"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtNERQTY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   33
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txttestprice 
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
         Left            =   2400
         TabIndex        =   32
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtID2 
         Height          =   375
         Left            =   -60480
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtID1 
         Height          =   375
         Left            =   14400
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd02 
         Caption         =   "维护"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -68760
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmd01 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -68760
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtOPPO 
         Height          =   375
         Left            =   -72240
         TabIndex        =   24
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtOPdevice 
         Height          =   375
         Left            =   -72240
         TabIndex        =   23
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtOPcust 
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
         Left            =   -72240
         TabIndex        =   22
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdquery 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   12480
         TabIndex        =   18
         Top             =   1320
         Width           =   1110
      End
      Begin VB.CommandButton cmdexport 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12480
         TabIndex        =   17
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton cmdlos 
         Caption         =   "失效"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   14280
         TabIndex        =   16
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdmod 
         Caption         =   "修改"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   14280
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "新增"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   12480
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox cmbCombo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   13
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbCombo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtDIE 
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
         Left            =   2400
         TabIndex        =   11
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtWafer 
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
         Left            =   2400
         TabIndex        =   10
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtDevice 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtCust 
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
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5295
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   15375
         _Version        =   524288
         _ExtentX        =   27120
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm_Price_Control.frx":0038
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7815
         Index           =   1
         Left            =   -75000
         TabIndex        =   27
         Top             =   3480
         Width           =   15615
         _Version        =   524288
         _ExtentX        =   27543
         _ExtentY        =   13785
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
         SpreadDesigner  =   "Frm_Price_Control.frx":0544
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblNER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NER数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   31
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label lbl07 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "测试费单价:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   3360
         Width           =   1770
      End
      Begin VB.Label lblOPENPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPENPO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl002 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO_PRCIE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   20
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lbl001 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         TabIndex        =   19
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label lbl06 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "币别:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblDIE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " DIE单价:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lbl04 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "片单价:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label lbl03 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "事业部:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lbl02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label lbl01 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   840
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Frm_Price_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum fpSprice
    e_ID
    e_choice
    E_cust
    e_device
    E_dept
    E_Wafer
    E_die
    E_test
    E_ner
    E_cur
    e_date
    e_by
    e_seq
    e_MCol
End Enum


Private Enum FpsOP
    e_ID
    E_OPcust
    E_OPprice
    e_OPPO
    e_OPREMARK
    e_OPBY
    e_OPTIME
    e_OPID
    e_MCol
End Enum




Private Sub cmd01_Click()

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 txtID2 = ""

        strSql = " SELECT a.cust_id AS  客户代码,a.po_price ,a.openpo,a.remark, ISNULL(a.last_update_by ,a.create_by) AS  创建人 " & _
                 " ,ISNULL(a.last_update_time,a.create_time )  AS 创建时间,A.ID FROM erptemp..ht_price_config a where a.flag = 0 "

      If Trim(txtOPcust.text) <> "" Then
        
        strSql = strSql + " AND a.cust_id = '" & Trim(txtOPcust.text) & "'"
        
      End If
      
      

   Fps(0).MaxRows = 0
   Fps(1).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType1(rs)
    Else
        MsgBox "没有相关产品价格信息", vbInformation, "提示"
        Exit Sub

    End If





End Sub

Private Sub cmd02_Click()

Dim strup As String
Dim stradd As String

Dim Userid As String
Dim rs         As New ADODB.Recordset
Dim strSql     As String
Dim poid As String


Userid = UCase(gUserName)

poid = GetPOPriceID()

If Trim(txtOPcust.text) = "" Or Trim(txtOPdevice.text) = "" Then
    
  MsgBox "  请输入完整信息 ", vbCritical, "警告"
  Exit Sub
End If
 
 
  strup = " UPDATE erptemp..ht_price_config SET FLAG = ID,last_update_by = '" & Userid & "' ,last_update_time = GETDATE()  WHERE cust_id = '" & Trim(txtOPcust.text) & "' "
  stradd = " INSERT INTO erptemp..ht_price_config (cust_id ,PO_PRICE ,OPENPO,fLAg,create_by,create_time ) VALUES ('" & Trim(txtOPcust.text) & "','" & Trim(txtOPdevice.text) & "' , " & _
          " '" & Trim(txtOPPO.text) & "' ,0,'" & Userid & "',  GETDATE() )  "

    If AddSql2(strup) = 0 Or AddSql2(stradd) = 0 Then
        
       MsgBox "  新增失败,请确认是否存在特殊字符", vbCritical, "警告"
       Exit Sub
       
    End If


     MsgBox "  维护完成", vbInformation, "提示"
                

End Sub

Private Sub cmdADD_Click()
Dim stradd As String
Dim stradd_new As String
Dim Userid          As String

Userid = UCase(gUserName)

txtID1 = ""

'strsql = "select EmpName from XTW..employee where empno = '" & strNPIOwnerNo & "'"
'strNPIOwnerName = Get_SqlStr2(strsql)

stradd = " SELECT * FROM erptemp..ht_price_control A WHERE a.cust_id = '" & Trim(txtCust.text) & "' AND a.cust_device = '" & Trim(txtDevice.text) & "' AND a.flag = 0 "

   If Get_SqlserverCnt(stradd) > 0 Then
        MsgBox "客户机种已存在价格,请确认已维护数据", vbCritical, "警告"
        
        Call price_data(Trim(txtCust.text), Trim(txtDevice.text))
        
        Exit Sub

    End If


If Trim(txtCust.text) = "" Or Trim(txtDevice.text) = "" Or Trim(txtWafer.text) = "" Or Trim(txtDIE.text) = "" Or Trim(cmbCombo1.text) = "" Or Trim(cmbCombo2.text) = "" Then
    
  MsgBox "  所有信息都为必输 , 请输入完整信息 ", vbCritical, "警告"
  Exit Sub
End If

  
  
  stradd_new = " INSERT INTO erptemp..ht_price_control (cust_id,cust_device,dept,wafer_price,die_price,currency,flag, create_by,create_time,test_price,product )" & _
               " VALUES ('" & Trim(txtCust.text) & "','" & Trim(txtDevice.text) & "','" & Trim(cmbCombo2.text) & "','" & Trim(txtWafer.text) & "','" & Trim(txtDIE.text) & "' " & _
               "  ,'" & Trim(cmbCombo1.text) & "',0, '" & Userid & "',GETDATE(),'" & txttestprice.text & "','" & txtNERQTY.text & "' ) "


     If AddSql2(stradd_new) = 0 Then
        
       MsgBox "  新增失败,请确认是否存在特殊字符", vbCritical, "警告"
       
        Exit Sub
    End If
    
   MsgBox "  维护完成", vbInformation, "提示"

End Sub


Private Sub price_data(cust As String, device As String)

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 txtID1 = ""

        strSql = " SELECT '' AS 选择,a.cust_id AS 客户代码,a.cust_device AS 客户机种,a.dept AS 部门,CONVERT(VARCHAR(100),ISNULL(a.wafer_price,0)) AS 片单价,CONVERT(VARCHAR(100),ISNULL(a.die_price,0)) AS DIE单价 " & _
                 " ,a.currency 币别,ISNULL(a.last_update_time,create_time ) AS  建立时间,ISNULL(a.last_update_by,create_by ) AS 创建人 FROM erptemp..ht_price_control A  WHERE a.flag = 0    AND a.cust_id = '" & cust & "'  AND a.cust_device = '" & device & "' "


    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)

    End If


End Sub




Private Sub Form_Load()


 With Fps(0)
 
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 1, 0, "选择"
    .ColWidth(e_choice) = 3
    .SetText 2, 0, "客户代码"
      .ColWidth(E_cust) = 10
    .SetText 3, 0, "客户机种"
      .ColWidth(e_device) = 15
    .SetText 4, 0, "事业部"
      .ColWidth(E_dept) = 8
    .SetText 5, 0, "片单价"
      .ColWidth(E_Wafer) = 8
    .SetText 6, 0, "DIE单价"
      .ColWidth(E_die) = 8
       .SetText 7, 0, "测试费单价"
      .ColWidth(E_test) = 8
         .SetText 8, 0, "NER数量"
      .ColWidth(E_ner) = 8
     .SetText 9, 0, "币别"
      .ColWidth(E_cur) = 8
    .SetText 10, 0, "维护日期"
      .ColWidth(e_date) = 10
    .SetText 11, 0, "维护人"
    .ColWidth(e_by) = 8
    .SetText 12, 0, "ID"
    .ColWidth(e_by) = 5
    
    
 End With
 
 
 
  With Fps(1)
 
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 1, 0, "选择"
    .ColWidth(E_OPcust) = 8
    .SetText 2, 0, "客户代码"
      .ColWidth(E_OPprice) = 5
    .SetText 3, 0, "po_price"
     .ColWidth(E_OPprice) = 5
     .SetText 4, 0, "ID"
     .ColWidth(e_OPPO) = 5
     .SetText 5, 0, "ID"
     .ColWidth(e_OPREMARK) = 18
     .SetText 6, 0, "ID"
     .ColWidth(e_OPBY) = 10
      .SetText 7, 0, "ID"
     .ColWidth(e_OPTIME) = 18
    
 End With
 
 

 Initcurrency
 Initdept
 
 If UCase(gUserName) <> "15236" And UCase(gUserName) <> "07885" Then
 
 cmdquery.Enabled = False
 cmdADD.Enabled = False
 cmdmod.Enabled = False
 cmdlos.Enabled = False
 cmdexport.Enabled = False
 cmd01.Enabled = False
 cmd02.Enabled = False
 
 End If
 

End Sub


Private Sub Initcurrency()

cmbCombo1.AddItem ("人民币")
cmbCombo1.AddItem ("美元")

End Sub


Private Sub Initdept()

cmbCombo2.AddItem ("TSV")
cmbCombo2.AddItem ("WLP")
cmbCombo2.AddItem ("SSP")
cmbCombo2.AddItem ("BUMPING")

End Sub


Private Sub cmdquery_Click()

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 
 txtID1 = ""

        strSql = " SELECT '' AS 选择,a.cust_id AS 客户代码,a.cust_device AS 客户机种,a.dept AS 部门,CONVERT(VARCHAR(100),ISNULL(a.wafer_price,0)) AS 片单价,CONVERT(VARCHAR(100),ISNULL(a.die_price,0)) AS DIE单价 " & _
                 " ,ISNULL(a.test_price,0 ) AS 测试费单价,a.product AS NER数量 ,a.currency 币别,ISNULL(a.last_update_time,create_time ) AS  建立时间,ISNULL(a.last_update_by,create_by ) AS 创建人,ID FROM erptemp..ht_price_control A  WHERE a.flag = 0 "

      If Trim(txtCust.text) <> "" Then
        
        strSql = strSql + " AND a.cust_id = '" & Trim(txtCust.text) & "'"
        
      End If
      
      If Trim(txtDevice.text) <> "" Then
      
       strSql = strSql + " AND a.cust_device = '" & Trim(txtDevice.text) & "'"
        
      End If
      
     If Trim(cmbCombo1.text) <> "" Then
      
       strSql = strSql + " AND  a.currency = '" & Trim(cmbCombo1.text) & "'"
        
      End If
      
      If Trim(cmbCombo2.text) <> "" Then
      
       strSql = strSql + " AND  a.dept = '" & Trim(cmbCombo2.text) & "'"
        
      End If
      
   
    
    Fps(0).MaxRows = 0
    Fps(1).MaxRows = 0
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "没有相关产品价格信息", vbInformation, "提示"
        Exit Sub

    End If


End Sub




Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long

    With Fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
            .text = 1
        Next
        
    End With

End Sub



Private Sub ListDataType1(rs As ADODB.Recordset)

    Dim i As Long

    With Fps(1)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With Fps(1)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
        Next
        
    End With

End Sub





'Private Sub FpsOP_DblClick(Index As Integer, ByVal AdvanceNext As Boolean)
'
'Dim opcust As String
'Dim opdevice As String
'Dim oppo As String
'Dim opid As String
'
'txtID2 = ""
'
'With FpsOP(1)
'    .Row = Row
'
'    If .Row <> 0 Then
'    .Col = 1
'     opcust = Trim(.text)
'     .Col = 2
'     opdevice = Trim(.text)
'     .Col = 3
'     oppo = Trim(.text)
'     .Col = 4
'     opid = Trim(.text)
'
'    End If
'
'txtOPcust.text = opcust
'txtOPdevice.text = opdevice
'txtOPPO.text = oppo
'txtID2.text = opid
'
'cmd02.Enabled = True
'
'End With



'End Sub

Private Sub Fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim cust_id As String
Dim cust_device As String
Dim dept As String
Dim WAFER As String
Dim die As String
Dim Curr As String
Dim id As String
Dim opcust As String
Dim opdevice As String
Dim oppo As String
Dim opid As String

Dim test_price As String
Dim ner As String


txtID1 = ""
txtID2 = ""

If Index = 0 Then

With Fps(0)
    .Row = Row
    
    If .Row <> 0 Then
    .Col = 2
     cust_id = Trim(.text)
     .Col = 3
     cust_device = Trim(.text)
     .Col = 4
     dept = Trim(.text)
    .Col = 5
     WAFER = Trim(.text)
    .Col = 6
     die = Trim(.text)
      .Col = 7
     test_price = Trim(.text)
      .Col = 8
     ner = Trim(.text)
     
    .Col = 9
     Curr = Trim(.text)
     .Col = 10
     id = Trim(.text)
      
    End If

End With

txtCust.text = cust_id
txtDevice.text = cust_device
txtWafer.text = WAFER
txtDIE.text = die
cmbCombo1.text = Curr
cmbCombo2.text = dept
txttestprice.text = test_price
txtNERQTY.text = ner

txtID1 = id

Else

With Fps(1)
    .Row = Row

    If .Row <> 0 Then
    .Col = 1
     opcust = Trim(.text)
     .Col = 2
     opdevice = Trim(.text)
     .Col = 3
     oppo = Trim(.text)
     .Col = 7
     opid = Trim(.text)

    End If

txtOPcust.text = opcust
txtOPdevice.text = opdevice
txtOPPO.text = oppo
txtID2.text = opid

cmd02.Enabled = True

End With



    
End If



End Sub






