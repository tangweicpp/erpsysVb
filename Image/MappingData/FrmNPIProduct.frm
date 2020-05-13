VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmNPIProduct 
   Caption         =   "NPI产品名称对照表设定"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
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
   ScaleHeight     =   10155
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtText1 
      Height          =   405
      Left            =   17280
      TabIndex        =   49
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPKG 
      Height          =   375
      Left            =   13320
      TabIndex        =   46
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox TxtCustPT6 
      Height          =   405
      Left            =   13320
      TabIndex        =   45
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox TxtCustPT5 
      Height          =   405
      Left            =   13320
      TabIndex        =   44
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton CmdDelData 
      BackColor       =   &H000000FF&
      Caption         =   "删除"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox TxtCustPT4 
      Height          =   375
      Left            =   9360
      TabIndex        =   39
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtCustPT3 
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtQtechPT2 
      Height          =   375
      Left            =   9360
      TabIndex        =   35
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtIDTemp 
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   204275713
      CurrentDate     =   41649
   End
   Begin VB.CommandButton CmdOutReport 
      Caption         =   "导出报表"
      Height          =   360
      Left            =   13200
      TabIndex        =   25
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   360
      Left            =   10920
      TabIndex        =   24
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "清空"
      Height          =   360
      Left            =   9360
      TabIndex        =   23
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "修改"
      Height          =   360
      Left            =   5160
      TabIndex        =   22
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   2880
      TabIndex        =   21
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox TxtStr3 
      Height          =   375
      Left            =   13320
      TabIndex        =   20
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox TxtStr2 
      Height          =   375
      Left            =   9480
      TabIndex        =   18
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox TxtStr1 
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TxtArea 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox TxtXS 
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox TxtQtechDie 
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox TxtCustDie 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox TxtCustPT2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox TxtCustPT1 
      Height          =   375
      Left            =   13320
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox TxtQtechPT 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6375
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   19815
      _Version        =   524288
      _ExtentX        =   34951
      _ExtentY        =   11245
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
      SpreadDesigner  =   "FrmNPIProduct.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1440
      TabIndex        =   27
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   220135425
      CurrentDate     =   41649
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   9360
      TabIndex        =   32
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   220135425
      CurrentDate     =   41649
   End
   Begin VB.Label lblLabel18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Residual："
      Height          =   195
      Left            =   16200
      TabIndex        =   48
      Top             =   240
      Width           =   780
   End
   Begin VB.Label lblLabel19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PKG-TYPE"
      Height          =   195
      Left            =   12000
      TabIndex        =   47
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label lbl20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种名6："
      Height          =   195
      Left            =   12000
      TabIndex        =   43
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label lbl19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种名5："
      Height          =   195
      Left            =   12000
      TabIndex        =   42
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种名4："
      Height          =   195
      Left            =   8160
      TabIndex        =   40
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种名3："
      Height          =   195
      Left            =   4080
      TabIndex        =   38
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   8160
      TabIndex        =   36
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转MP日期："
      Height          =   195
      Left            =   8160
      TabIndex        =   33
      Top             =   2640
      Width           =   930
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转小批量日期："
      Height          =   195
      Left            =   4080
      TabIndex        =   30
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第一次打样日期："
      Height          =   195
      Left            =   0
      TabIndex        =   28
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本3："
      Height          =   195
      Left            =   12000
      TabIndex        =   19
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本2："
      Height          =   195
      Left            =   8160
      TabIndex        =   17
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装结构版本1："
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用领域："
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "像素："
      Height          =   195
      Left            =   8160
      TabIndex        =   11
      Top             =   1440
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
      Caption         =   "客户机种名2："
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1170
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
    E_CustName               '客户代码
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
    E_XS                   '像素
    E_Area                '应用领域
    E_Stu1                '封装结构版本1
    E_Stu2                '封装结构版本2
    E_Stu3                '封装结构版本3
    
    E_Time1                '第一次打样日期
    E_Time2                '转小批量日期
    E_Time3                '转MP日期
    E_secondCode                '二级代码
    E_End
    
End Enum

Dim reportRS As New adodb.Recordset
Dim mainItemRS As New adodb.Recordset
Dim bomRS2        As New adodb.Recordset

Private Sub CmbCustomer_Change()
TxtQtechPT.SetFocus
End Sub

Private Sub CmdAdd_Click()
Dim nPIProductTemp As NpiProduct

Dim userid As String

'校验是否重复


If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtQtechPT.Text)) = "" Then
     MsgBox "客户代码或厂内项目名称不可以为空！"
     Exit Sub
End If

If UCase(Trim(TxtCustPT1.Text)) = "" And UCase(Trim(TxtCustPT2.Text)) = "" Then
     MsgBox "客户机种不可以为空！"
     Exit Sub
End If

'ccs add 20160717
'If CmbCustomer.Text = "GC" And Text1.Text = "" Then
 ' MsgBox "GC客户必须录入二级代码！"
 ' Exit Sub
'End If


 Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))
If bomRS2.RecordCount > 0 Then
    MsgBox "系统中已存在这笔数据，请重新确认输入是否正确 ！"
    Exit Sub
End If

userid = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CustomershortName = UCase(Trim(CmbCustomer.Text))
nPIProductTemp.qtechPTNo = UCase(Trim(TxtQtechPT.Text))
nPIProductTemp.QtechPTNo2 = UCase(Trim(TxtQtechPT2.Text))
nPIProductTemp.CustomerPTNo1 = Trim(TxtCustPT1.Text)
nPIProductTemp.CustomerPTNo2 = Trim(TxtCustPT2.Text)
nPIProductTemp.CustomerPTNo3 = Trim(TxtCustPT3.Text)
nPIProductTemp.CustomerPTNo4 = Trim(TxtCustPT4.Text)
nPIProductTemp.CustomerPTNo5 = Trim(TxtCustPT5.Text)
nPIProductTemp.CustomerPTNo6 = Trim(TxtCustPT6.Text)
nPIProductTemp.CustomerDieQty = UCase(Trim(TxtCustDie.Text))
nPIProductTemp.QtechDieQty = UCase(Trim(TxtQtechDie.Text))
nPIProductTemp.XiangSu = UCase(Trim(TxtXS.Text))
nPIProductTemp.UsedArea = UCase(Trim(TxtArea.Text))
nPIProductTemp.StruckStr1 = UCase(Trim(TxtStr1.Text))
nPIProductTemp.StruckStr2 = UCase(Trim(TxtStr2.Text))
nPIProductTemp.StruckStr3 = UCase(Trim(TxtStr3.Text))
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.Value), "", DTPicker3.Value)
nPIProductTemp.PKG = UCase(Trim(txtPKG.Text))
nPIProductTemp.residual = UCase(Trim(txtText1.Text))

If nPIProductTemp.CustomershortName = "37" And Len(nPIProductTemp.PKG) < 1 Then

MsgBox "请填写PKG"
Exit Sub

End If

Call AddNpiProduct(nPIProductTemp)

 MsgBox "新增成功!", vbInformation, "友情提示"
 
 ShowData_Where

End Sub

Private Sub CmdDel_Click()
CmbCustomer.Text = ""
TxtQtechPT.Text = ""
TxtQtechPT2.Text = ""
TxtCustPT1.Text = ""
TxtCustPT2.Text = ""
TxtCustDie.Text = ""
TxtQtechDie.Text = ""
TxtXS.Text = ""
TxtArea.Text = ""
TxtArea.Text = ""
TxtStr1.Text = ""
TxtStr2.Text = ""
TxtStr3.Text = ""



End Sub

Private Sub CmdDelData_Click()

'修改


Dim userid As String


If CLng(TxtIDTemp.Text) >= 1 Then

Call DelDataNpiProduct(CLng(TxtIDTemp.Text))

 MsgBox "修改成功!", vbInformation, "友情提示"

ShowData_Where

Else

 MsgBox "请先双击要删除的行!", vbInformation, "友情提示"

End If



End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()
'修改

Dim nPIProductTemp As NpiProduct
Dim userid As String
userid = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CustomershortName = UCase(Trim(CmbCustomer.Text))
nPIProductTemp.qtechPTNo = UCase(Trim(TxtQtechPT.Text))
nPIProductTemp.QtechPTNo2 = UCase(Trim(TxtQtechPT2.Text))
nPIProductTemp.CustomerPTNo1 = Trim(TxtCustPT1.Text)
nPIProductTemp.CustomerPTNo2 = Trim(TxtCustPT2.Text)
nPIProductTemp.CustomerPTNo3 = Trim(TxtCustPT3.Text)
nPIProductTemp.CustomerPTNo4 = Trim(TxtCustPT4.Text)
nPIProductTemp.CustomerPTNo5 = Trim(TxtCustPT5.Text)
nPIProductTemp.CustomerPTNo6 = Trim(TxtCustPT6.Text)
nPIProductTemp.CustomerDieQty = UCase(Trim(TxtCustDie.Text))
nPIProductTemp.QtechDieQty = UCase(Trim(TxtQtechDie.Text))
nPIProductTemp.XiangSu = UCase(Trim(TxtXS.Text))
nPIProductTemp.UsedArea = UCase(Trim(TxtArea.Text))
nPIProductTemp.StruckStr1 = UCase(Trim(TxtStr1.Text))
nPIProductTemp.StruckStr2 = UCase(Trim(TxtStr2.Text))
nPIProductTemp.StruckStr3 = UCase(Trim(TxtStr3.Text))
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.Value), "", DTPicker3.Value)
nPIProductTemp.PKG = UCase(Trim(txtPKG.Text))
nPIProductTemp.residual = UCase(Trim(txtText1.Text))

If nPIProductTemp.CustomershortName = "37" And Len(nPIProductTemp.PKG) < 1 Then

MsgBox "请填写PKG"
Exit Sub
End If

Call ModifyNpiProduct(nPIProductTemp, CLng(TxtIDTemp.Text))

 MsgBox "修改成功!", vbInformation, "友情提示"

ShowData_Where

End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id  , CUSTOMERSHORTNAME as 客户代码 , QtechPTNo as 厂内项目名称 ,QtechPTNo2 as 成品料号, CUSTOMERPTNo1  as 客户机种名1, CUSTOMERPTNo2 as 客户机种名2 ,CUSTOMERPTNo3 as 客户机种名3 ,CUSTOMERPTNo4 as 客户机种名4,CUSTOMERPTNo5 as 客户机种名5,CUSTOMERPTNo6 as 客户机种名6,  " & _
         " CUSTOMERDieQty as 客户设计die数, QtechDieQty as 厂内die数, XiangSu  as 像素, UsedArea as 应用领域, StruckStr1 as 封装结构版本1, StruckStr2 as 封装结构版本2, StruckStr3 as 封装结构版本3, ST_DATE as 第一次打样日期,TT_DATE  as 转小批量日期,PT_DATE as 转MP日期 , PKG_TYPE " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
         
  ExporToExcel (sqlTemp)

End Sub

Private Sub Form_Activate()
CmbCustomer.SetFocus

End Sub

Private Sub Form_Load()

IniCustomerName
IniFpsHeader

DTPicker1.Value = DateTime.Date
DTPicker2.Value = DateTime.Date
DTPicker3.Value = DateTime.Date

'DTPicker1.MultiSelect = True
'DTPicker2.MultiSelect = True
'DTPicker3.MultiSelect = True


DTPicker1.Value = Null
DTPicker2.Value = Null
DTPicker3.Value = Null

ShowData_Where

'根据用户名,看是否有修改的权限

Call UserType(UCase(gUserName))



End Sub

Private Sub UserType(nametemp As String)

If nametemp = "13221" Or nametemp = "07885" Or nametemp = "11155" Then
CmdAdd.Enabled = True
CmdModify.Enabled = True
CmdDelData.Enabled = True


Else

CmdAdd.Enabled = False

CmdModify.Enabled = False
CmdDelData.Enabled = False


End If



End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub ShowData_Where()
Set reportRS = GetNPIData()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub



Private Sub IniFpsHeader()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
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
        .SetText E_FPS0.E_CustName, 0, "客户代码"
        .SetText E_FPS0.E_QtechPT, 0, "厂内项目名称"
        .SetText E_FPS0.E_QtechPT2, 0, "成品料号"
        .SetText E_FPS0.E_CustPT1, 0, "客户机种名1"
        .SetText E_FPS0.E_CustPT2, 0, "客户机种名2"
        .SetText E_FPS0.E_CustPT3, 0, "客户机种名3"
        .SetText E_FPS0.E_CustPT4, 0, "客户机种名4"
        .SetText E_FPS0.E_CustPT5, 0, "客户机种名5"
        .SetText E_FPS0.E_CustPT6, 0, "客户机种名6"
                        
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
        .SetText E_FPS0.E_secondCode, 0, "PKG_TYPE"
        
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustName) = 6
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
        .ColWidth(E_FPS0.E_secondCode) = 10
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub



Private Sub fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i As Long

With fps(0)
            .Row = Row
            .Col = 1
       i = .Text

End With

showData (i)

End Sub

Private Sub showData(i As Long)

Set reportRS = GetNPIDataID(i)


 If reportRS.RecordCount > 0 Then
 
 
    CmbCustomer.Text = reportRS.fields("CustomershortName").Value & ""
    TxtQtechPT.Text = reportRS.fields("QtechPTNo").Value & ""
    TxtQtechPT2.Text = reportRS.fields("QtechPTNo2").Value & ""
    TxtCustPT1.Text = reportRS.fields("CustomerPTNo1").Value & ""
    TxtCustPT2.Text = reportRS.fields("CustomerPTNo2").Value & ""
    TxtCustDie.Text = reportRS.fields("CustomerDieQty").Value & ""
    TxtQtechDie.Text = reportRS.fields("QtechDieQty").Value & ""
    TxtXS.Text = reportRS.fields("XiangSu").Value & ""
    TxtArea.Text = reportRS.fields("UsedArea").Value & ""
    TxtStr1.Text = reportRS.fields("StruckStr1").Value & ""
    TxtStr2.Text = reportRS.fields("StruckStr2").Value & ""
    TxtStr3.Text = reportRS.fields("StruckStr3").Value & ""
    DTPicker1.Value = reportRS.fields("ST_DATE").Value
    DTPicker2.Value = reportRS.fields("TT_DATE").Value
    DTPicker3.Value = reportRS.fields("PT_DATE").Value
    TxtIDTemp.Text = reportRS.fields("ID").Value
    txtPKG.Text = reportRS.fields("PKG_TYPE").Value & ""
    txtText1.Text = reportRS.fields("residual").Value & ""
     
 End If



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



