VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmNPIProduct 
   Caption         =   "NPI产品名称对照表设定"
   ClientHeight    =   11205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   ScaleHeight     =   11205
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtText3 
      Height          =   375
      Left            =   5400
      TabIndex        =   67
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtcust 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   65
      Top             =   150
      Width           =   2175
   End
   Begin VB.ComboBox DTPicker3 
      Height          =   315
      ItemData        =   "FrmNPIProduct.frx":0000
      Left            =   9360
      List            =   "FrmNPIProduct.frx":000A
      TabIndex        =   63
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox Text1 
      BackColor       =   &H00FFC0FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   62
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox Text2 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "FrmNPIProduct.frx":0016
      Left            =   5400
      List            =   "FrmNPIProduct.frx":0023
      TabIndex        =   61
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtWaferPN 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   13320
      TabIndex        =   56
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查找"
      Height          =   360
      Left            =   1320
      TabIndex        =   54
      Top             =   3600
      Width           =   990
   End
   Begin VB.ComboBox cbMapping 
      BackColor       =   &H00FFC0FF&
      Height          =   315
      ItemData        =   "FrmNPIProduct.frx":0031
      Left            =   17160
      List            =   "FrmNPIProduct.frx":0033
      TabIndex        =   53
      Top             =   1980
      Width           =   1215
   End
   Begin VB.ComboBox txtProEng 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "FrmNPIProduct.frx":0035
      Left            =   17160
      List            =   "FrmNPIProduct.frx":0037
      TabIndex        =   51
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtText2 
      BackColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   13320
      TabIndex        =   48
      Top             =   1335
      Width           =   2415
   End
   Begin VB.TextBox txtNPIOwnerNO 
      BackColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   17160
      TabIndex        =   47
      Top             =   135
      Width           =   975
   End
   Begin VB.TextBox txtPKG 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   44
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TxtCustPT6 
      Height          =   405
      Left            =   17160
      TabIndex        =   43
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox TxtCustPT5 
      Height          =   405
      Left            =   13320
      TabIndex        =   42
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton CmdDelData 
      BackColor       =   &H000000FF&
      Caption         =   "删除"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3600
      Width           =   990
   End
   Begin VB.TextBox TxtCustPT4 
      Height          =   375
      Left            =   9360
      TabIndex        =   37
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtCustPT3 
      Height          =   375
      Left            =   9360
      TabIndex        =   35
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox TxtQtechPT2 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   33
      Top             =   150
      Width           =   2295
   End
   Begin VB.TextBox TxtIDTemp 
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   108658689
      CurrentDate     =   41649
   End
   Begin VB.CommandButton CmdOutReport 
      Caption         =   "导出报表"
      Height          =   360
      Left            =   12840
      TabIndex        =   25
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   360
      Left            =   11160
      TabIndex        =   24
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "清空"
      Height          =   360
      Left            =   9360
      TabIndex        =   23
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "修改"
      Height          =   360
      Left            =   5400
      TabIndex        =   22
      Top             =   3600
      Width           =   990
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   3360
      TabIndex        =   21
      Top             =   3600
      Width           =   990
   End
   Begin VB.TextBox TxtStr3 
      Height          =   375
      Left            =   13320
      TabIndex        =   20
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox TxtStr2 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   18
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox TxtStr1 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox TxtArea 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox TxtXS 
      Height          =   375
      Left            =   17280
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox TxtQtechDie 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   1350
      Width           =   2295
   End
   Begin VB.TextBox TxtCustDie 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1350
      Width           =   2175
   End
   Begin VB.TextBox TxtCustPT2 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox TxtCustPT1 
      BackColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   13320
      TabIndex        =   4
      Top             =   135
      Width           =   2415
   End
   Begin VB.TextBox TxtQtechPT 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   150
      Width           =   2295
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   19815
      _Version        =   524288
      _ExtentX        =   34951
      _ExtentY        =   10821
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
      SpreadDesigner  =   "FrmNPIProduct.frx":0039
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   17640
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761087
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   108658689
      CurrentDate     =   41649
   End
   Begin VB.PictureBox Image21 
      Height          =   1455
      Left            =   7320
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   59
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblPART 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PART："
      Height          =   195
      Left            =   4440
      TabIndex        =   66
      Top             =   840
      Width           =   570
   End
   Begin VB.Label txtNPIOwnerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   18360
      TabIndex        =   64
      Top             =   240
      Width           =   60
   End
   Begin VB.Label lblboned 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保税非保税:"
      Height          =   195
      Left            =   8160
      TabIndex        =   60
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "寸别："
      Height          =   195
      Left            =   4440
      TabIndex        =   58
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "生产部："
      Height          =   195
      Left            =   600
      TabIndex        =   57
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblWaferPN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "晶圆料号："
      Height          =   195
      Left            =   12000
      TabIndex        =   55
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapping是否有："
      Height          =   195
      Left            =   15840
      TabIndex        =   52
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label lblPE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工程量产："
      Height          =   195
      Left            =   16200
      TabIndex        =   50
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打标码："
      Height          =   195
      Left            =   12000
      TabIndex        =   49
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lblLabel18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "负责人工号："
      Height          =   195
      Left            =   16080
      TabIndex        =   46
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabel19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PKG-TYPE"
      Height          =   195
      Left            =   8280
      TabIndex        =   45
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label lbl20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "玻璃规格："
      Height          =   195
      Left            =   16200
      TabIndex        =   41
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lbl19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "清洗程序："
      Height          =   195
      Left            =   12000
      TabIndex        =   40
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CV高度："
      Height          =   195
      Left            =   8160
      TabIndex        =   38
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "清洗步骤："
      Height          =   195
      Left            =   8040
      TabIndex        =   36
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   8160
      TabIndex        =   34
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转小批量日期："
      Height          =   195
      Left            =   4080
      TabIndex        =   30
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第一次打样日期:"
      Height          =   195
      Left            =   0
      TabIndex        =   28
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本3："
      Height          =   195
      Left            =   11880
      TabIndex        =   19
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本2:"
      Height          =   195
      Left            =   8040
      TabIndex        =   17
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本1："
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用领域："
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "像素："
      Height          =   195
      Left            =   16440
      TabIndex        =   11
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "厂内die数："
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户设计die数："
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAB_DEVICE："
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种名1："
      Height          =   195
      Left            =   12000
      TabIndex        =   3
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "厂内项目名称："
      Height          =   195
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码："
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "FrmNPIProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail

    E_SeqId = 1                '序号
    E_CUSTNAME               '客户代码
    E_QtechPT                '厂内项目名称
    E_QtechPT2                '成品料号
    E_CustPT1                '客户机种名1
    E_CustPT2                '客户机种名2
    E_CustPT3                '客户机种名3
    E_CustPT4                '客户机种名4
    E_CustPT5                '客户机种名5
    E_CustPT6                '客户机种名6
    E_CustDie                '客户设计die数
    E_QtechDie                '厂内die数
    E_XS                    '像素
    E_Area                  '应用领域
    E_Stu1                  '封装结构版本1
    E_Stu2                  '封装结构版本2
    E_Stu3                 '封装结构版本3
    E_Time1                '第一次打样日期
    E_Time2                '转小批量日期
    E_Time3                '转MP日期
    E_SecondCode                '二级代码
    E_MARKINGCODE          '打标码
    E_ProduEng             '工程量产
    E_Mapping
    E_Owner
    E_WaferPN
    E_CustPT7
    E_CustPT8
    E_END

End Enum

Dim reportRS   As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2     As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsfab As New ADODB.Recordset

Private Sub CmbCustomer_Change()
TxtQtechPT.SetFocus
End Sub




Private Sub cmdADD_Click()
Dim nPIProductTemp  As NpiProduct
Dim Userid          As String
Dim strNPIOwnerName As String
Dim strNPIOwnerNo   As String
Dim strSql          As String
Dim strSqlfab          As String
Dim strfab  As String
Dim fab_wafer As String

Dim strsqldev          As String

strNPIOwnerNo = Trim$(txtNPIOwnerNO.text)

If strNPIOwnerNo = "" Then
    MsgBox "请填入机种对应NPI负责人的工号", vbInformation, "提示"
    Exit Sub

End If

strSql = "select EmpName from XTW..employee where empno = '" & strNPIOwnerNo & "'"
strNPIOwnerName = Get_SqlStr2(strSql)

If strNPIOwnerName = "" Then
    MsgBox "填入的NPI负责人工号不正确,请确认", vbInformation, "提示"
    Exit Sub

End If

txtNPIOwnerName.Caption = strNPIOwnerName

nPIProductTemp.residual = strNPIOwnerNo

If UCase(Trim(txtcust.text)) = "" Or UCase(Trim(TxtQtechPT.text)) = "" Then
    MsgBox "客户代码或厂内项目名称不可以为空！"
    Exit Sub

End If

If UCase(Trim(TxtCustPT1.text)) = "" And UCase(Trim(TxtCustPT2.text)) = "" Then
    MsgBox "客户机种不可以为空！"
    Exit Sub

End If

If TxtQtechPT2.text = "" Then
    MsgBox "成品料号不可以为空!"
    Exit Sub
Else

    If Get_SqlserverCnt("select * from AIS20141114094336.dbo.t_ICItem where F_101 = '" & UCase(Trim$(TxtQtechPT2.text)) & "' ") = 0 Then
        MsgBox "金蝶未维护该成品料号, 请确认是否输入错误", vbCritical, "警告"
        Exit Sub

    End If

    If Left(Right(UCase(Trim$(TxtQtechPT2.text)), 3), 1) <> "W" And (txtWaferPN.text = "") Then
        MsgBox "晶圆料号不可以为空!"
        Exit Sub
    Else

        If Get_SqlserverCnt("select * from AIS20141114094336.dbo.t_ICItem where F_101 = '" & UCase(Trim$(txtWaferPN.text)) & "' ") = 0 Then
            MsgBox "金蝶未维护该晶圆料号, 请确认是否输入错误", vbCritical, "警告"
            Exit Sub

        End If

    End If

End If

If TxtCustDie.text = "" Then
    MsgBox "客户设计DIE不可以为空"
    Exit Sub

End If

If TxtQtechDie.text = "" Then
    MsgBox "厂内DIE不可以为空"
    Exit Sub

End If

If cbMapping.text = "" Then
    MsgBox "该客户必须填写是否有MAPPING", vbCritical, "警告"
    Exit Sub

End If

If txtText2.text = "" Then
    MsgBox "该客户必须填写打标码长度", vbCritical, "警告"
    Exit Sub

End If

If txtProEng.text = "" Then
    MsgBox "工程量产不可为空"
    Exit Sub

End If

'
If Text1.text = "" Then
    MsgBox "生产部不可为空"
    Exit Sub

End If


If Text2.text = "" Then
    MsgBox "寸别不可为空"
    Exit Sub

End If

If txtPKG.text = "" Then
    MsgBox "PKG-TYPE 不可为空"
    Exit Sub

End If

If TxtStr1.text = "" Then
    MsgBox " 封装结构版本1不可为空"
    Exit Sub

End If

If TxtStr2.text = "" Then
    MsgBox " 封装结构版本2 不可为空"
    Exit Sub

End If

If txtText2.text = "" Then
    MsgBox " 打标码 不可为空"
    Exit Sub

End If

  strSql = "SELECT a.CUSTOMER,a.REMARK1, a.REMARK2  FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(txtcust.text)) & "'  AND a.REMARK1 = 'Y'"
  
   If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
    
    If UCase(Trim(TxtCustPT2.text)) = "" Then
     
     MsgBox "请输入FAB_DEVICE"
      Exit Sub
     End If
     
     
    If rs.Fields(2) = "3" And UCase(Trim(txtText3.text)) = "" Then
     
     MsgBox "请输入PART"
      Exit Sub
     End If
        
    
    strfab = "  select sum(qty) from ( select p.customerptno2,count(distinct p.marketlastupdate_by ) as qty from tbltsvnpiproduct p  where p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' " & _
             "  and p.marketlastupdate_by <>  '" & UCase(Trim(txtWaferPN.text)) & "' group by p.customerptno2 Union select  p.marketlastupdate_by,count(distinct p.customerptno2 ) as qty from tbltsvnpiproduct p   " & _
             "   where p.marketlastupdate_by = '" & UCase(Trim(txtWaferPN.text)) & "'  and p.customerptno2 <> '" & UCase(Trim(TxtCustPT2.text)) & "' group by p.marketlastupdate_by ) X "

    
    fab_wafer = Get_OracleStr(strfab)
    If Val(fab_wafer) <> 0 Then
    
     MsgBox "FAB_DEVICE已存在唯一晶圆料号"
      Exit Sub
        
    End If
    
    
    
    
    If rs.Fields(2).Value = "2" Then
    
    strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "'   group by p.customershortname,p.customerptno1,p.customerptno2 "
     ElseIf rs.Fields(2).Value = "3" Then
     
      strSqlfab = "  select p.customershortname,p.customerptno1,p.customerptno2,p.customerptno3,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'   " & _
                  " and p.customerptno1 =  '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' " & _
                  " and p.customerptno3 = '" & UCase(Trim(txtText3.text)) & "'    group by p.customershortname,p.customerptno1,p.customerptno2 ,p.customerptno3  "
     ElseIf rs.Fields(2).Value = "9" Then
      strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "'   group by p.customershortname,p.customerptno1,p.customerptno2 "
     
     Else
         
       strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "'   group by p.customershortname,p.customerptno1,p.customerptno2 "
     End If
     
   
             
     
      If rsfab.State = adStateOpen Then rsfab.Close
      rsfab.Open strSqlfab, Cnn, adOpenStatic, adLockReadOnly, adCmdText
      
      If Not rsfab.EOF Then
        If rsfab.Fields(3).Value > 0 Then
          
           MsgBox "客户机种+FAB_DEVICE 已经存在唯一成品料号"
           Exit Sub
         
        End If
    End If
 
End If

'Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))
     
   Set bomRS2 = GetNpiProductCheck1(UCase(Trim(txtcust.text)), UCase(Trim(TxtQtechPT2.text)))
     
If bomRS2.RecordCount > 0 Then
    MsgBox "系统中已存在这笔数据，请重新确认输入是否正确 ！"
    Exit Sub

End If

Userid = UCase(gUserName)
nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CUSTOMERSHORTNAME = Replace(UCase(Trim(txtcust.text)), Chr(13) + Chr(10), "")
nPIProductTemp.qtechPTNo = Replace(UCase(Trim(TxtQtechPT.text)), Chr(13) + Chr(10), "")
nPIProductTemp.QtechPTNo2 = Replace(UCase(Trim(TxtQtechPT2.text)), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo1 = Replace(Trim(TxtCustPT1.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo2 = Replace(Trim(TxtCustPT2.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo3 = Replace(Trim(txtText3.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo4 = Replace(Trim(TxtCustPT4.text), Chr(13) + Chr(10), "")
'
nPIProductTemp.CustomerPTNo5 = Replace(Trim(TxtCustPT5.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo6 = Replace(Trim(TxtCustPT6.text), Chr(13) + Chr(10), "")
''''''
nPIProductTemp.CustomerPTNo7 = Replace(Trim(Text1.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo8 = Replace(Trim(Text2.text), Chr(13) + Chr(10), "")
''''''
nPIProductTemp.CustomerDieQty = Replace(UCase(Trim(TxtCustDie.text)), Chr(13) + Chr(10), "")
nPIProductTemp.QtechDieQty = Replace(UCase(Trim(TxtQtechDie.text)), Chr(13) + Chr(10), "")
nPIProductTemp.XiangSu = Replace(UCase(Trim(TxtXS.text)), Chr(13) + Chr(10), "")
nPIProductTemp.UsedArea = Replace(UCase(Trim(TxtArea.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr1 = Replace(UCase(Trim(TxtStr1.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr2 = Replace(UCase(Trim(TxtStr2.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr3 = Replace(UCase(Trim(TxtStr3.text)), Chr(13) + Chr(10), "")
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.text), "", DTPicker3.text)
nPIProductTemp.PKG = Replace(UCase(Trim(txtPKG.text)), Chr(13) + Chr(10), "")
nPIProductTemp.MARKINGCODE = Replace(UCase(Trim(txtText2.text)), Chr(13) + Chr(10), "")
nPIProductTemp.ProducEng = Replace(UCase(Trim(txtProEng.text)), Chr(13) + Chr(10), "")
nPIProductTemp.MAPPING = cbMapping.text
nPIProductTemp.WaferPN = Replace(UCase(Trim(txtWaferPN.text)), Chr(13) + Chr(10), "")
nPIProductTemp.UpdatePrice2 = Replace(UCase(Trim(Text1.text)), Chr(13) + Chr(10), "")
nPIProductTemp.UpdatePrice2 = Replace(UCase(Trim(Text1.text)), Chr(13) + Chr(10), "")

If nPIProductTemp.CUSTOMERSHORTNAME = "37" And Len(nPIProductTemp.PKG) < 1 Then
    MsgBox "请填写PKG"
    Exit Sub

End If

'判断是否有重复的
Dim sOra As String
Dim sId  As String
sOra = "select id from tbltsvnpiproduct where customershortname = '" & nPIProductTemp.CUSTOMERSHORTNAME & "' and  customerptno1 = '" & nPIProductTemp.CustomerPTNo1 & "'  and qtechptno2='" & nPIProductTemp.QtechPTNo2 & "' and qtechptno = '" & nPIProductTemp.qtechPTNo & "'  "

If Get_OracleCnt(sOra) > 0 Then
    ListData

    If MsgBox("已经存在同样的一笔维护信息(客户代码,客户机种,厂内机种,成品料号都一样), 请确认是否要更新? ", vbOKCancel, "提示") = vbOK Then
        sId = Get_OracleStr(sOra)
        Call ModifyNpiProduct(nPIProductTemp, CLng(TxtIDTemp.text))
        MsgBox "修改成功!", vbInformation, "友情提示"
        Call ListData
        Exit Sub

    End If

End If

'Call AddNpiProduct1(nPIProductTemp)
Call AddNpiProduct1(nPIProductTemp)
MsgBox "新增成功!", vbInformation, "友情提示"
ShowData_Where
cmdDel_Click

End Sub

' add
Private Sub AddNpiProduct1(nPIProductTemp As NpiProduct)
Dim sqlTemp  As String
Dim sqlTemp1 As String
Dim sqlid    As String
Dim id       As Long
Dim Rs3      As New ADODB.Recordset
sqlid = "  SELECT  NpiProduct_SEQ.Nextval FROM DUAL "

If Rs3.State = adStateOpen Then Rs3.Close
Rs3.Open sqlid, Cnn, adOpenStatic, adLockReadOnly, adCmdText
id = Val(Rs3.Fields(0).Value)
'sqlTemp = " insert into TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
'          " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
'          " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
'          " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,PKG_TYPE,Residual,MARKING_CODE, P_E ,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
'          " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
'          "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
'          "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',sysdate, " & _
'          "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,   '" & nPIProductTemp.CustomerPTNo3 & "', " & _
'          " '" & nPIProductTemp.CustomerPTNo4 & "',  '" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MarkingCode & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
sqlTemp = " insert into TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
   " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
   " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
   " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,UPDATEPRICE2,UPDATEPRICE1,PKG_TYPE,Residual,MARKING_CODE, P_E ,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
   " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
   "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
   "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',sysdate, " & _
   "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,   '" & nPIProductTemp.CustomerPTNo3 & "', " & _
   " '" & nPIProductTemp.CustomerPTNo4 & "',  '" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.CustomerPTNo7 & "', '" & nPIProductTemp.CustomerPTNo8 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MARKINGCODE & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
'sqlTemp1 = " insert into erptemp..TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
'          " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
'          " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
'          " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,PKG_TYPE,Residual,MARKING_CODE, P_E,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
'          " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
'          "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
'          "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',CONVERT(varchar(100),GETDATE(),20), " & _
'          "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,'" & nPIProductTemp.CustomerPTNo3 & "', " & _
'          " '" & nPIProductTemp.CustomerPTNo4 & "','" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MarkingCode & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
sqlTemp1 = " insert into erptemp..TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
   " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
   " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
   " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,UPDATEPRICE2,UPDATEPRICE1,PKG_TYPE,Residual,MARKING_CODE, P_E ,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
   " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
   "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
   "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',CONVERT(varchar(100),GETDATE(),20), " & _
   "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,   '" & nPIProductTemp.CustomerPTNo3 & "', " & _
   " '" & nPIProductTemp.CustomerPTNo4 & "',  '" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.CustomerPTNo7 & "', '" & nPIProductTemp.CustomerPTNo8 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MARKINGCODE & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
AddSql (sqlTemp)
AddSql2 (sqlTemp1)

End Sub

Private Sub cmdDel_Click()
CmbCustomer.text = ""
cbMapping.text = ""
TxtQtechPT.text = ""
TxtQtechPT2.text = ""
TxtCustPT1.text = ""
TxtCustPT2.text = ""
TxtCustDie.text = ""
TxtQtechDie.text = ""
TxtXS.text = ""
TxtArea.text = ""
TxtArea.text = ""
TxtStr1.text = ""
TxtStr2.text = ""
TxtStr3.text = ""
txtNPIOwnerNO.text = ""
TxtQtechDie.text = ""
txtPKG.text = ""
TxtStr1.text = ""
TxtStr2.text = ""
txtWaferPN.text = ""
txtText2.text = ""
Text1.text = ""
Text2.text = ""
DTPicker3.text = ""
txtProEng.text = ""
TxtCustPT3.text = ""
TxtCustPT4.text = ""
TxtCustPT5.text = ""
TxtCustPT6.text = ""
txtText3.text = ""
txtcust.text = ""

End Sub

Private Sub CmdDelData_Click()
'删除
Dim Userid As String

If TxtIDTemp.text = "" Then
    Exit Sub

End If

If CLng(TxtIDTemp.text) >= 1 Then
    If MsgBox("是否确认要删除", vbOKCancel) = vbOK Then
        Call DelDataNpiProduct(CLng(TxtIDTemp.text))
        MsgBox "删除成功!", vbInformation, "友情提示"
        ShowData_Where
        cmdDel_Click
    Else
        Exit Sub

    End If

Else
    MsgBox "请先双击要删除的行!", vbInformation, "友情提示"

End If

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdModify_Click()
'修改
Dim nPIProductTemp As NpiProduct
Dim Userid         As String

Dim strNPIOwnerName As String
Dim strNPIOwnerNo   As String
Dim strSql          As String

Dim strSqlfab          As String
Dim strfab  As String
Dim fab_wafer As String

strNPIOwnerNo = Trim$(txtNPIOwnerNO.text)

If strNPIOwnerNo = "" Then
    MsgBox "请填入机种对应NPI负责人的工号", vbInformation, "提示"
    Exit Sub

End If

strSql = "select EmpName from XTW..employee where empno = '" & strNPIOwnerNo & "'"
strNPIOwnerName = Get_SqlStr2(strSql)

If strNPIOwnerName = "" Then
    MsgBox "填入的NPI负责人工号不正确,请确认", vbInformation, "提示"
    Exit Sub

End If

txtNPIOwnerName.Caption = strNPIOwnerName

nPIProductTemp.residual = strNPIOwnerNo

If UCase(Trim(CmbCustomer.text)) = "" Or UCase(Trim(TxtQtechPT.text)) = "" Then
    MsgBox "客户代码或厂内项目名称不可以为空！"
    Exit Sub

End If

If UCase(Trim(TxtCustPT1.text)) = "" And UCase(Trim(TxtCustPT2.text)) = "" Then
    MsgBox "客户机种不可以为空！"
    Exit Sub

End If

If UCase(Trim(txtNPIOwnerNO.text)) = "" And UCase(Trim(TxtCustPT2.text)) = "" Then
    MsgBox "OWNER不可以为空！"
    Exit Sub

End If

If TxtQtechPT2.text = "" Then
    MsgBox "成品料号不可以为空!"
    Exit Sub
Else

    If Get_SqlserverCnt("select * from AIS20141114094336.dbo.t_ICItem where F_101 = '" & UCase(Trim$(TxtQtechPT2.text)) & "' ") = 0 Then
        MsgBox "金蝶未维护该成品料号, 请确认是否输入错误", vbCritical, "警告"
        Exit Sub

    End If

    If Left(Right(UCase(Trim$(TxtQtechPT2.text)), 3), 1) <> "W" And (txtWaferPN.text = "") Then
        MsgBox "晶圆料号不可以为空!"
        Exit Sub
    Else

        If Get_SqlserverCnt("select * from AIS20141114094336.dbo.t_ICItem where F_101 = '" & UCase(Trim$(txtWaferPN.text)) & "' ") = 0 Then
            MsgBox "金蝶未维护该晶圆料号, 请确认是否输入错误", vbCritical, "警告"
            Exit Sub

        End If

    End If

End If

If TxtCustDie.text = "" Then
    MsgBox "客户设计DIE不可以为空"
    Exit Sub

End If

If TxtQtechDie.text = "" Then
    MsgBox "厂内DIE不可以为空"
    Exit Sub

End If

If cbMapping.text = "" Then
    MsgBox "该客户必须填写是否有MAPPING", vbCritical, "警告"
    Exit Sub

End If

If txtText2.text = "" Then
    MsgBox "该客户必须填写打标码长度", vbCritical, "警告"
    Exit Sub

End If

If txtProEng.text = "" Then
    MsgBox "工程量产不可为空"
    Exit Sub

End If

'
If Text1.text = "" Then
    MsgBox "生产部不可为空"
    Exit Sub

End If

If Text2.text = "" Then
    MsgBox "寸别不可为空"
    Exit Sub

End If

If txtPKG.text = "" Then
    MsgBox "PKG-TYPE 不可为空"
    Exit Sub

End If

If TxtStr1.text = "" Then
    MsgBox " 封装结构版本1不可为空"
    Exit Sub

End If

If TxtStr2.text = "" Then
    MsgBox " 封装结构版本2 不可为空"
    Exit Sub

End If



   strSql = "SELECT a.CUSTOMER,a.REMARK1, a.REMARK2  FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(txtcust.text)) & "'  AND a.REMARK1 = 'Y'"
  
   If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
    
     If UCase(Trim(TxtCustPT2.text)) = "" Then
     MsgBox "请输入FAB_DEVICE"
      Exit Sub
        
    End If
    
      If rs.Fields(2) = "3" And UCase(Trim(txtText3.text)) = "" Then
     MsgBox "请输入PART"
      Exit Sub
     End If
        
        
    strfab = "  select sum(qty) from ( select p.customerptno2,count(distinct p.marketlastupdate_by ) as qty from tbltsvnpiproduct p  where p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' " & _
             "   and p.marketlastupdate_by <>  '" & UCase(Trim(txtWaferPN.text)) & "' group by p.customerptno2 Union select  p.marketlastupdate_by,count(distinct p.customerptno2 ) as qty from tbltsvnpiproduct p   " & _
             "   where p.marketlastupdate_by = '" & UCase(Trim(txtWaferPN.text)) & "'  and p.customerptno2 <> '" & UCase(Trim(TxtCustPT2.text)) & "' group by p.marketlastupdate_by ) X "

    
    fab_wafer = Get_OracleStr(strfab)
    If Val(fab_wafer) <> 0 Then
    
     MsgBox "FAB_DEVICE已存在唯一晶圆料号"
      Exit Sub
        
    End If
    
    
    If rs.Fields(2).Value = "2" Then
    
    strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' and id <> '" & CLng(TxtIDTemp.text) & "'  group by p.customershortname,p.customerptno1,p.customerptno2 "
     
     
     ElseIf rs.Fields(2).Value = "3" Then
     
      strSqlfab = "   select p.customershortname,p.customerptno1,p.customerptno2, p.customerptno3,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname =  '" & UCase(Trim(txtcust.text)) & "'   " & _
                  " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 =  '" & UCase(Trim(TxtCustPT2.text)) & "'   and  p.customerptno3 =  '" & UCase(Trim(txtText3.text)) & "'  and id <> '" & CLng(TxtIDTemp.text) & "'  " & _
                 "  group by p.customershortname,p.customerptno1,p.customerptno2 , p.customerptno3 "
     
     
     ElseIf rs.Fields(2).Value = "9" Then
     
      strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' and id <> '" & CLng(TxtIDTemp.text) & "'  group by p.customershortname,p.customerptno1,p.customerptno2 "
     
     Else
     
     
     strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname = '" & UCase(Trim(txtcust.text)) & "'     " & _
             " and p.customerptno1 = '" & UCase(Trim(TxtCustPT1.text)) & "'   and  p.customerptno2 = '" & UCase(Trim(TxtCustPT2.text)) & "' and id <> '" & CLng(TxtIDTemp.text) & "'  group by p.customershortname,p.customerptno1,p.customerptno2 "
     
         
     End If
     
     
   
     
     
      If rsfab.State = adStateOpen Then rsfab.Close
      rsfab.Open strSqlfab, Cnn, adOpenStatic, adLockReadOnly, adCmdText
      
      If Not rsfab.EOF Then
        If rsfab.Fields(3).Value > 0 Then
          
           MsgBox "客户机种+FAB_DEVICE 已经存在唯一成品料号"
           Exit Sub
         
        End If
    End If

End If



Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.text)), UCase(Trim(TxtQtechPT.text)), UCase(Trim(TxtCustPT1.text)), UCase(Trim(TxtCustPT2.text)), UCase(Trim(txtText3.text)), CLng(TxtIDTemp.text), UCase(Trim(TxtQtechPT2.text)))

If bomRS2.RecordCount > 0 Then
    MsgBox "系统中已存在这笔数据，请重新确认输入是否正确 ！"
    Exit Sub

End If

Userid = UCase(gUserName)
nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CUSTOMERSHORTNAME = Replace(UCase(Trim(txtcust.text)), Chr(13) + Chr(10), "")
nPIProductTemp.qtechPTNo = Replace(UCase(Trim(TxtQtechPT.text)), Chr(13) + Chr(10), "")
nPIProductTemp.QtechPTNo2 = Replace(UCase(Trim(TxtQtechPT2.text)), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo1 = Replace(Trim(TxtCustPT1.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo2 = Replace(Trim(TxtCustPT2.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo3 = Replace(Trim(txtText3.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo4 = Replace(Trim(TxtCustPT4.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo5 = Replace(Trim(TxtCustPT5.text), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo6 = Replace(Trim(TxtCustPT6.text), Chr(13) + Chr(10), "")
'
nPIProductTemp.CustomerPTNo7 = Replace(UCase(Trim(Text1.text)), Chr(13) + Chr(10), "")
nPIProductTemp.CustomerPTNo8 = Replace(Trim(Text2.text), Chr(13) + Chr(10), "")
'
nPIProductTemp.CustomerDieQty = Replace(UCase(Trim(TxtCustDie.text)), Chr(13) + Chr(10), "")
nPIProductTemp.QtechDieQty = Replace(UCase(Trim(TxtQtechDie.text)), Chr(13) + Chr(10), "")
nPIProductTemp.XiangSu = Replace(UCase(Trim(TxtXS.text)), Chr(13) + Chr(10), "")
nPIProductTemp.UsedArea = Replace(UCase(Trim(TxtArea.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr1 = Replace(UCase(Trim(TxtStr1.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr2 = Replace(UCase(Trim(TxtStr2.text)), Chr(13) + Chr(10), "")
nPIProductTemp.StruckStr3 = Replace(UCase(Trim(TxtStr3.text)), Chr(13) + Chr(10), "")
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.text), "", DTPicker3.text)
nPIProductTemp.PKG = Replace(UCase(Trim(txtPKG.text)), Chr(13) + Chr(10), "")
nPIProductTemp.residual = Replace(UCase(Trim(txtNPIOwnerNO.text)), Chr(13) + Chr(10), "")
nPIProductTemp.MARKINGCODE = Replace(UCase(Trim(txtText2.text)), Chr(13) + Chr(10), "")   ' By Tony, 20170814
nPIProductTemp.ProducEng = Replace(UCase(Trim(txtProEng.text)), Chr(13) + Chr(10), "")
nPIProductTemp.MAPPING = Trim(cbMapping.text)
nPIProductTemp.WaferPN = Replace(Trim(txtWaferPN.text), Chr(13) + Chr(10), "")

If nPIProductTemp.CUSTOMERSHORTNAME = "37" And Len(nPIProductTemp.PKG) < 1 Then
    MsgBox "请填写PKG"
    Exit Sub

End If

If TxtIDTemp.text = "" Then
    MsgBox "无效操作", vbCritical, "警告"
    Exit Sub

End If

'Call ModifyNpiProduct(nPIProductTemp, CLng(TxtIDTemp.Text))
Call ModifyNpiProduct1(nPIProductTemp, CLng(TxtIDTemp.text))
MsgBox "修改成功!", vbInformation, "友情提示"
ListData
cmdDel_Click

End Sub

Private Sub ModifyNpiProduct1(nPIProductTemp As NpiProduct, idTemp As Long)
Dim sqlTemp  As String
Dim sqlTemp1 As String

If gUserName = "16642" Or gUserName = "15236" Or gUserName = "12725" Or gUserName = "16452" Or gUserName = "14117" Or gUserName = "12089" Or gUserName = "15507" Or gUserName = "16368" Or gUserName = "19400" Or gUserName = "08240" Then
    MsgBox "当前账号只有修改客户机种的权限", vbInformation, "提示"
    sqlTemp = " Update TBLTsvNpiProduct set CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "' Where id = " & idTemp & ""
    sqlTemp1 = " Update erptemp..TBLTsvNpiProduct set CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "' Where id = " & idTemp & ""
Else
    '         sqlTemp = " Update TBLTsvNpiProduct " & _
    '           " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
    '           " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
    '           " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
    '           " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
    '           " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=sysdate,CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
    '           " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MarkingCode & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "' , MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'  Where id = " & idTemp & ""
    '
    sqlTemp = " Update TBLTsvNpiProduct " & _
       " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
       " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
       " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
       " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
       " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=sysdate,CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
       " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MARKINGCODE & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "' , UPDATEPRICE2 ='" & nPIProductTemp.CustomerPTNo7 & "',UPDATEPRICE1 ='" & nPIProductTemp.CustomerPTNo8 & "',MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'  Where id = " & idTemp & ""
    '          sqlTemp1 = " Update erptemp..TBLTsvNpiProduct " & _
    '           " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
    '           " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
    '           " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
    '           " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
    '           " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=GetDate(),CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
    '           " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MarkingCode & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "', MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'    Where id = " & idTemp & ""
    '
    sqlTemp1 = " Update erptemp..TBLTsvNpiProduct " & _
       " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
       " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
       " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
       " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
       " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=GetDate(),CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
       " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MARKINGCODE & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "' , UPDATEPRICE2 ='" & nPIProductTemp.CustomerPTNo7 & "',UPDATEPRICE1 ='" & nPIProductTemp.CustomerPTNo8 & "',MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'  Where id = " & idTemp & ""

End If

AddSql (sqlTemp)
AddSql2 (sqlTemp1)

End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String
sqlTemp = "select  id, CUSTOMERSHORTNAME as 客户代码 , QtechPTNo as 厂内项目名称 ,QtechPTNo2 as 成品料号," & _
   " CUSTOMERPTNo1  as 客户机种名1, CUSTOMERPTNo2 as 客户机种名2 ,CUSTOMERPTNo3 as 客户机种名3 ,CUSTOMERPTNo4 as " & _
   " 客户机种名4,CUSTOMERPTNo5 as 客户机种名5,CUSTOMERPTNo6 as 客户机种名6,  " & " CUSTOMERDieQty as 客户设计die数, " & _
   "QtechDieQty as 厂内die数, XiangSu  as 像素, UsedArea as 应用领域, StruckStr1 as 封装结构版本1, StruckStr2 as 封装结构版本2," & _
   "StruckStr3 as 封装结构版本3, ST_DATE as 第一次打样日期,TT_DATE  as 转小批量日期,PT_DATE as 转MP日期 , PKG_TYPE , MARKING_CODE " & _
   "as 打标码 ,  P_E as 工程量产,RESIDUAL as OWNER,MARKETLASTUPDATE_BY as 晶圆料号,UpdatePrice2 as 生产部,UpdatePrice1 as 寸别 " & " From TBLTsvNpiProduct  " & _
   "order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
ExporToExcel (sqlTemp)

End Sub

Private Sub cmdquery_Click()
ListData

End Sub

Private Sub Form_Activate()
'CmbCustomer.SetFocus

End Sub

Private Sub Form_Load()
Dim strSql As String
Dim rs     As ADODB.Recordset
Dim j      As Long
IniCustomerName
IniFpsHeader
cbMapping.AddItem ("Y")
cbMapping.AddItem ("N")
DTPicker1.Value = DateTime.DATE
DTPicker2.Value = DateTime.DATE
'DTPicker3.Value = DateTime.DATE
DTPicker1.Value = Null
DTPicker2.Value = Null
'  DTPicker3.Value = Null
Set rs = New ADODB.Recordset
strSql = "select distinct UPDATEPRICE2 from TBLTsvNpiProduct where UPDATEPRICE2  IS NOT NULL  "
'  If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
Text1.Clear

If rs.RecordCount > 0 Then
    rs.MoveFirst

    For j = 1 To rs.RecordCount
        Text1.AddItem Trim(rs("UPDATEPRICE2"))
        rs.MoveNext
    Next
    rs.Clone
    Set rs = Nothing

End If

txtProEng.AddItem "样品"
txtProEng.AddItem "小批量"
txtProEng.AddItem "量产"
ShowData_Where
Call UserType(UCase(gUserName))

End Sub

Private Sub UserType(nametemp As String)

If nametemp = "12447" Or nametemp = "07885" Or nametemp = "15475" Or nametemp = "15580" Or nametemp = "17226" Or nametemp = "13221" Or nametemp = "13396" Then
    CmdAdd.Enabled = True
    CmdModify.Enabled = True
    CmdDelData.Enabled = True
'ElseIf nametemp = "16642" Or nametemp = "18420" Or nametemp = "15236" Or nametemp = "12725" Or nametemp = "16452" Or nametemp = "14117" Or nametemp = "12089" Or nametemp = "15507" Or nametemp = "16368" Or nametemp = "19400" Or nametemp = "08240" Then
'    cmdADD.Enabled = False
'    CmdModify.Enabled = True
'    CmdDelData.Enabled = False
Else
    CmdAdd.Enabled = False
    CmdModify.Enabled = False
    CmdDelData.Enabled = False

End If

End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").name
CmbCustomer.BoundColumn = mainItemRS("PID").name

End Sub

Private Sub ShowData_Where()
'Set reportRS = GetNPIData()
Set reportRS = GetNPIData1()

With fps(0)
    .MaxRows = 0

    If reportRS.RecordCount > 0 Then
        Set .DataSource = reportRS

    End If

End With

End Sub

Private Function GetNPIData1() As ADODB.Recordset
Dim cmdStr   As String
Dim RSResult As New ADODB.Recordset
'cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3   , CUSTOMERPTNo4 , CUSTOMERPTNo5 , CUSTOMERPTNo6  , " & _
'         " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE ,PKG_TYPE,MARKING_CODE , P_E,MAPPING,RESIDUAL,MARKETLASTUPDATE_BY " & _
'         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3   , CUSTOMERPTNo4 , CUSTOMERPTNo5 , CUSTOMERPTNo6  , " & " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE ,PKG_TYPE,MARKING_CODE , P_E,MAPPING,RESIDUAL,MARKETLASTUPDATE_BY,UpdatePrice2,UpdatePrice1 " & " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
Set RSResult = getStr(cmdStr)
Set GetNPIData1 = RSResult

End Function

Private Sub ListData()
Set reportRS = GetNPIData2()

With fps(0)
    .MaxRows = 0

    If reportRS.RecordCount > 0 Then
        Set .DataSource = reportRS

    End If

End With

End Sub

Private Function GetNPIData2() As ADODB.Recordset
Dim cmdStr   As String
Dim RSResult As New ADODB.Recordset
sApp = "order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2"
'cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3   , CUSTOMERPTNo4 , CUSTOMERPTNo5 , CUSTOMERPTNo6  , " & " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE ,PKG_TYPE,MARKING_CODE , P_E,MAPPING,RESIDUAL,MARKETLASTUPDATE_BY " & " From TBLTsvNpiProduct where 1 = 1  "
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3   , CUSTOMERPTNo4 , CUSTOMERPTNo5 , CUSTOMERPTNo6  , " & " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE ,PKG_TYPE,MARKING_CODE , P_E,MAPPING,RESIDUAL,MARKETLASTUPDATE_BY,UpdatePrice2,UpdatePrice1 " & " From TBLTsvNpiProduct where 1 = 1  "

If txtcust.text <> "" Then
    cmdStr = cmdStr & "and customershortname = '" & Trim(txtcust.text) & "'"

End If

If TxtQtechPT.text <> "" Then
    cmdStr = cmdStr & "and qtechptno = '" & Trim$(TxtQtechPT.text) & "'"

End If

If TxtCustPT1.text <> "" Then
    cmdStr = cmdStr & "and customerptno1 = '" & Trim$(TxtCustPT1.text) & "'"

End If

If TxtQtechPT2.text <> "" Then
    cmdStr = cmdStr & "and qtechptno2 = '" & Trim$(TxtQtechPT2.text) & "'"

End If

cmdStr = cmdStr & sApp
Set RSResult = getStr(cmdStr)
Set GetNPIData2 = RSResult

End Function

Private Sub IniFpsHeader()

With fps(0)
    .ReDraw = False
    .MaxCols = E_FPS0.E_END - 1
    .MaxRows = 0
    '砞竚Α
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    '        .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_FPS0.E_SeqId, 0, "记录号"
    .SetText E_FPS0.E_CUSTNAME, 0, "客户代码"
    .SetText E_FPS0.E_QtechPT, 0, "厂内项目名称"
    .SetText E_FPS0.E_QtechPT2, 0, "成品料号"
    .SetText E_FPS0.E_CustPT1, 0, "客户机种名1"
    .SetText E_FPS0.E_CustPT2, 0, "客户机种名2"
    .SetText E_FPS0.E_CustPT3, 0, "清洗步骤"
    .SetText E_FPS0.E_CustPT4, 0, "CV高度"
    .SetText E_FPS0.E_CustPT5, 0, "清洗程序"
    .SetText E_FPS0.E_CustPT6, 0, "玻璃规格"
    .SetText E_FPS0.E_CustDie, 0, "客户设计die数"
    .SetText E_FPS0.E_QtechDie, 0, "厂内die数"
    .SetText E_FPS0.E_XS, 0, "像素"
    .SetText E_FPS0.E_Area, 0, "应用领域"
    .SetText E_FPS0.E_Stu1, 0, "封装结构版本1"
    .SetText E_FPS0.E_Stu2, 0, "封装结构版本2"
    .SetText E_FPS0.E_Stu3, 0, "封装结构版本3"
    .SetText E_FPS0.E_Time1, 0, "第一次打样日期"
    .SetText E_FPS0.E_Time2, 0, "转小批量日期"
    .SetText E_FPS0.E_Time3, 0, "转MP日期"
    .SetText E_FPS0.E_SecondCode, 0, "PKG_TYPE"
    .SetText E_FPS0.E_MARKINGCODE, 0, "打标码"
    .SetText E_FPS0.E_ProduEng, 0, "工程量产"
    .SetText E_FPS0.E_Mapping, 0, "是否有MAPPING"
    .SetText E_FPS0.E_Owner, 0, "OWNER"
    .SetText E_FPS0.E_WaferPN, 0, "晶圆料号"
    .SetText E_FPS0.E_CustPT7, 0, "生产部"
    .SetText E_FPS0.E_CustPT8, 0, "寸别"
    .ColWidth(E_FPS0.E_SeqId) = 5
    .ColWidth(E_FPS0.E_CUSTNAME) = 6
    .ColWidth(E_FPS0.E_QtechPT) = 10
    .ColWidth(E_FPS0.E_QtechPT2) = 12
    .ColWidth(E_FPS0.E_CustPT1) = 20
    .ColWidth(E_FPS0.E_CustPT2) = 10
    .ColWidth(E_FPS0.E_CustPT3) = 10
    .ColWidth(E_FPS0.E_CustPT4) = 10
    .ColWidth(E_FPS0.E_CustPT5) = 10
    .ColWidth(E_FPS0.E_CustPT6) = 10
    .ColWidth(E_FPS0.E_CustDie) = 10
    .ColWidth(E_FPS0.E_QtechDie) = 10
    .ColWidth(E_FPS0.E_XS) = 10
    .ColWidth(E_FPS0.E_Area) = 10
    .ColWidth(E_FPS0.E_Stu1) = 12
    .ColWidth(E_FPS0.E_Stu2) = 12
    .ColWidth(E_FPS0.E_Stu3) = 12
    .ColWidth(E_FPS0.E_Time1) = 10
    .ColWidth(E_FPS0.E_Time2) = 10
    .ColWidth(E_FPS0.E_Time3) = 10
    .ColWidth(E_FPS0.E_SecondCode) = 10
    .ColWidth(E_FPS0.E_MARKINGCODE) = 10
    .ColWidth(E_FPS0.E_ProduEng) = 10
    .ColWidth(E_FPS0.E_Mapping) = 4
    'add
    .ColWidth(E_FPS0.E_WaferPN) = 12
    .ColWidth(E_FPS0.E_CustPT7) = 10
    .ColWidth(E_FPS0.E_CustPT8) = 10
    '
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

End Sub

Private Sub Fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i As Long

With fps(0)
    .Row = Row
    .Col = 1

    If .Row <> 0 Then
        i = .text

    End If

End With

ShowData (i)

End Sub

Public Function GetNPIDataID1(idTemp As Long) As ADODB.Recordset
'查询GC MarkingCode
Dim cmdStr   As String
Dim RSResult As New ADODB.Recordset
'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3,CUSTOMERPTNo4,CUSTOMERPTNo5,CUSTOMERPTNo6, " & " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE,pkg_type,residual, MARKING_CODE, MAPPING,MARKETLASTUPDATE_BY,UpdatePrice2,UpdatePrice1,P_E   " & " From TBLTsvNpiProduct where id=" & idTemp & "  order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
Set RSResult = getStr(cmdStr)
Set GetNPIDataID1 = RSResult

End Function

Private Sub ShowData(i As Long)
'Set reportRS = GetNPIDataID(I)
Set reportRS = GetNPIDataID1(i)

If reportRS.RecordCount > 0 Then
    CmbCustomer.text = reportRS.Fields("CustomershortName").Value & ""
    txtcust.text = reportRS.Fields("CustomershortName").Value & ""
    txtcust.text = reportRS.Fields("CustomershortName").Value & ""
    TxtQtechPT.text = reportRS.Fields("QtechPTNo").Value & ""
    TxtQtechPT2.text = reportRS.Fields("QtechPTNo2").Value & ""
    TxtCustPT1.text = reportRS.Fields("CustomerPTNo1").Value & ""
    TxtCustPT2.text = reportRS.Fields("CustomerPTNo2").Value & ""
    TxtCustDie.text = reportRS.Fields("CustomerDieQty").Value & ""
    TxtQtechDie.text = reportRS.Fields("QtechDieQty").Value & ""
    TxtXS.text = reportRS.Fields("XiangSu").Value & ""
    TxtArea.text = reportRS.Fields("UsedArea").Value & ""
    TxtStr1.text = reportRS.Fields("StruckStr1").Value & ""
    TxtStr2.text = reportRS.Fields("StruckStr2").Value & ""
    TxtStr3.text = reportRS.Fields("StruckStr3").Value & ""
    DTPicker1.Value = reportRS.Fields("ST_DATE").Value & ""
    DTPicker2.Value = reportRS.Fields("TT_DATE").Value & ""
    DTPicker3.text = reportRS.Fields("PT_DATE").Value & ""
    TxtIDTemp.text = reportRS.Fields("ID").Value & ""
    txtPKG.text = reportRS.Fields("PKG_TYPE").Value & ""
    txtProEng.text = reportRS.Fields("P_E").Value & ""
    txtNPIOwnerNO.text = reportRS.Fields("residual").Value & ""
    txtText2.text = reportRS.Fields("MARKING_CODE").Value & ""
    cbMapping.text = reportRS.Fields("MAPPING").Value & ""
    txtText3.text = reportRS.Fields("CustomerPTNo3").Value & ""
    TxtCustPT4.text = reportRS.Fields("CustomerPTNo4").Value & ""
    TxtCustPT5.text = reportRS.Fields("CustomerPTNo5").Value & ""
    TxtCustPT6.text = reportRS.Fields("CustomerPTNo6").Value & ""
    txtWaferPN.text = reportRS.Fields("MARKETLASTUPDATE_BY").Value & ""
    'add
    Text1.text = reportRS.Fields("UpdatePrice2").Value & ""
    Text2.text = reportRS.Fields("UpdatePrice1").Value & ""

    '
End If

End Sub

Private Sub txtNPIOwnerNO_Change()
Dim strSql As String
strSql = "select EmpName from XTW..employee where empno = '" & Trim$(txtNPIOwnerNO.text) & "'"
txtNPIOwnerName.Caption = Get_SqlStr2(strSql)
End Sub

Private Sub TxtQtechPT_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtQtechPT2.SetFocus

End If

End Sub

Private Sub TxtQtechPT2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtCustPT1.SetFocus

End If

End Sub

Private Sub TxtCustPT1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtCustPT2.SetFocus

End If

End Sub

Private Sub TxtCustPT2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtCustDie.SetFocus

End If

End Sub

Private Sub TxtCustDie_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtQtechDie.SetFocus

End If

End Sub

Private Sub TxtQtechDie_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtXS.SetFocus

End If

End Sub

Private Sub TxtXS_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtArea.SetFocus

End If

End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtStr1.SetFocus

End If

End Sub

Private Sub TxtStr1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtStr2.SetFocus

End If

End Sub

Private Sub TxtStr2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtStr3.SetFocus

End If

End Sub

'ADD
Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text1.SetFocus

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text2.SetFocus

End If

End Sub






