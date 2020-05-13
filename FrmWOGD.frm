VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmWOGD 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   13890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   25170
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   13890
   ScaleWidth      =   25170
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCheck2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8520
      TabIndex        =   26
      Top             =   240
      Width           =   255
   End
   Begin VB.ComboBox ComSFLL 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "FrmWOGD.frx":0000
      Left            =   9840
      List            =   "FrmWOGD.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00808000&
      Caption         =   "启用更新"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin FPSpreadADO.fpSpread Fps1 
      Height          =   6975
      Index           =   1
      Left            =   14400
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
      _Version        =   524288
      _ExtentX        =   11668
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
      MaxCols         =   3
      MaxRows         =   0
      SpreadDesigner  =   "FrmWOGD.frx":0016
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.CheckBox chkCheck1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdall 
      BackColor       =   &H0000FF00&
      Caption         =   "导出excel"
      Height          =   600
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19680
      MaskColor       =   &H008080FF&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   990
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "清空控件值"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   18600
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton command4 
      BackColor       =   &H00008080&
      Caption         =   "标记该记录为不可开立"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   16080
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton command3 
      Caption         =   "查询不可开立信息记录"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   18600
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton command2 
      Caption         =   "二级查询晶圆信息"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13560
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton command1 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13560
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTP 
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   166068225
      CurrentDate     =   43738
   End
   Begin MSComCtl2.DTPicker DTP 
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   13
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   166068225
      CurrentDate     =   43738
   End
   Begin MSDataListLib.DataCombo CboSYB 
      Height          =   345
      Left            =   3360
      TabIndex        =   16
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboKHDM 
      Height          =   345
      Left            =   6600
      TabIndex        =   17
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboKHJZ 
      Height          =   345
      Left            =   9840
      TabIndex        =   18
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   11295
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   2160
      Width           =   23535
      _Version        =   524288
      _ExtentX        =   41513
      _ExtentY        =   19923
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
      MaxCols         =   5
      MaxRows         =   5
      SpreadDesigner  =   "FrmWOGD.frx":04F8
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lblSFLL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否来料"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   8880
      TabIndex        =   25
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否勾选作为条件"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   20
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "来料结束时间"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8640
      TabIndex        =   14
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "来料开始时间"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   12
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label laebl_head 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已上传WO未开工单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   -120
      TabIndex        =   10
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   7
      Top             =   600
      Width           =   720
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   960
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "事业部"
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
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   660
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "FrmWOGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mainItemRS As New ADODB.Recordset
Dim mainItemRS1 As New ADODB.Recordset
Dim strSqls As String
Private Enum fpSDetail
    E_CHOOSE = 1
    e_DJBH = 8
End Enum

Private Sub cmdCheckData_Click()
   CheckData
'MsgBox "功能未开放"
End Sub
Private Function Query()
'查询
Command2.Visible = True
Command4.Visible = True
    Dim strSql       As String

    Dim strsql1     As String
    
    Dim strSql2    As String
    
    Dim rs           As New ADODB.Recordset
    
    Dim SYB         As String

    Dim KHJZ      As String

    Dim KHDM         As String
    
    Dim SFLL As String
    
    Dim create_by       As String
  
    Dim Create_time_start     As String
    
    Dim Create_time_end   As String
    
    
    SYB = Trim(CboSYB.Text)
    KHJZ = Trim(CboKHJZ.Text)
    KHDM = Trim(CboKHDM.Text)
    SFLL = Trim(ComSFLL.Text)
    Create_time_start = DTP(0).Value
    Create_time_end = DTP(1).Value

       strsql1 = "select  CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) as '来料时间', d.UPdateprice2 as '事业部', A.CUSTOMERSHORTNAME as '客户代码'," & _
            "B.MPN_DESC as '客户机种名',B.mtrl_num as '厂内机种名',b.po_num as 'PO',B.SOURCE_BATCH_ID as 'lot号',right(A.SUBSTRATEID,2) as 'wafer号',a.passbincount,a.SUBSTRATETYPE,d.STRUCKSTR2 " & _
            "FROM ERPBASE..tblmappingData a ," & _
            "ERPBASE..tblCustomerOI b,erptemp..tbltsvnpiproduct d " & _
            "WHERE a.FLAG = 'Y' " & _
            "AND CONVERT(VARCHAR(100),b.ID) = a.FILENAME " & _
            "and d.customerptno1 = B.MPN_DESC "
      
      strSql2 = "select  CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) as '来料时间', d.UPdateprice2 as '事业部', A.CUSTOMERSHORTNAME as '客户代码'," & _
            "B.MPN_DESC as '客户机种名',B.mtrl_num as '厂内机种名',b.po_num as 'PO',B.SOURCE_BATCH_ID as 'lot号',A.SUBSTRATEID,a.passbincount,a.SUBSTRATETYPE,d.STRUCKSTR2 " & _
            "FROM ERPBASE..tblmappingData a ," & _
            "ERPBASE..tblCustomerOI b,erptemp..tbltsvnpiproduct d " & _
            "WHERE a.FLAG = 'Y' " & _
            "AND CONVERT(VARCHAR(100),b.ID) = a.FILENAME " & _
            "and d.customerptno1 = B.MPN_DESC "
   
      
    If SYB <> "" Then
        strsql1 = strsql1 + " AND  d.UPdateprice2   = '" & SYB & "'  "
        strSql2 = strSql2 + " AND  d.UPdateprice2   = '" & SYB & "'  "
    End If

    If KHJZ <> "" Then
        strsql1 = strsql1 + " AND B.MPN_DESC = '" & KHJZ & "'  "
        strSql2 = strSql2 + " AND B.MPN_DESC = '" & KHJZ & "'  "
    End If

    If KHDM <> "" Then
        strsql1 = strsql1 + " AND A.CUSTOMERSHORTNAME  = '" & KHDM & "'  "
        strSql2 = strSql2 + " AND A.CUSTOMERSHORTNAME  = '" & KHDM & "'  "
    End If
    
    If chkCheck1 = 1 Then
        strsql1 = strsql1 + "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) >='" & Format(Create_time_start, "yyyy-mm-dd") & "' " & _
                    "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) <='" & Format(Create_time_end, "yyyy-mm-dd") & "' "
        strSql2 = strSql2 + "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) >='" & Format(Create_time_start, "yyyy-mm-dd") & "' " & _
                    "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) <='" & Format(Create_time_end, "yyyy-mm-dd") & "' "
    
    End If
    
    strsql1 = strsql1 + "and not EXISTS ( select * from erpdata .. tblTSVwaferlist c where c.waferid = A.SUBSTRATEID ) " & _
                        "group by CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) , d.UPdateprice2 , A.CUSTOMERSHORTNAME ," & _
                        "B.MPN_DESC ,B.mtrl_num,b.po_num ,B.SOURCE_BATCH_ID,A.SUBSTRATEID,a.passbincount,a.SUBSTRATETYPE,d.STRUCKSTR2 "

    strSql2 = strSql2 + "and not EXISTS ( select * from erpdata .. tblTSVwaferlist c where c.waferid = A.SUBSTRATEID ) " & _
                        "group by CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) , d.UPdateprice2 , A.CUSTOMERSHORTNAME ," & _
                        "B.MPN_DESC ,B.mtrl_num,b.po_num ,B.SOURCE_BATCH_ID,A.SUBSTRATEID,a.passbincount,a.SUBSTRATETYPE,d.STRUCKSTR2 "

    If chkCheck2 = 1 Then
        If SFLL = "否" Then
        strSql = "SELECT * from ( SELECT cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,case when e.仓库编号 is  null then " & _
                " '尚未来料'else e.仓库编号 end as '是否来料存放仓库',cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer片数,cc.wafer序号 from ( " & _
                "SELECT '' as '选择',aa.来料时间,aa.事业部,aa.客户代码,aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount as 'WO过站die',aa.SUBSTRATETYPE as '保税类型',aa.STRUCKSTR2 as '产品结构' " & _
                ",count(aa.SUBSTRATEID) as 'wafer片数',wafer序号 = " & _
                "(STUFF((SELECT ',' + test.wafer号 " & _
                "FROM ( " + strsql1 + " ) test WHERE aa.来料时间 = test.来料时间 and aa.事业部 = test.事业部 and " & _
                "  aa.客户代码 = test.客户代码 and aa.客户机种名 = test.客户机种名 and aa.厂内机种名 = test.厂内机种名 and aa.PO = test.PO " & _
                "  and aa.lot号 = test.lot号 FOR XML PATH('')), 1,  1, '')) from ( " + strSql2 + ") aa group by aa.来料时间,aa.事业部,aa.客户代码, aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount,aa.SUBSTRATETYPE,aa.STRUCKSTR2 )cc " & _
                "left join ERPBASE..tblstocknum e on cc.lot号 = e.批号 and e.当前存量 > 0 and e.仓库编号 <> '54' where not exists (select * from erptemp..BZGDKL f where f.lot = cc.lot号) group by cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,cc.来料时间,cc.事业部, " & _
                "cc.客户代码 , cc.客户机种名, cc.厂内机种名, cc.PO, cc.lot号,e.仓库编号, cc.wafer片数,cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer序号 ) kk where kk.是否来料存放仓库 = '尚未来料' order by kk.来料时间 desc,kk.事业部,kk.客户代码,kk.客户机种名,kk.厂内机种名,kk.PO,kk.lot号,kk.wafer片数,kk.是否来料存放仓库,kk.wafer序号 "
        ElseIf SFLL = "是" Then
          strSql = "SELECT * from ( SELECT cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,case when e.仓库编号 is  null then " & _
            " '尚未来料'else e.仓库编号 end as '是否来料存放仓库',cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer片数,cc.wafer序号 from ( " & _
            "SELECT '' as '选择',aa.来料时间,aa.事业部,aa.客户代码,aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount as 'WO过站die',aa.SUBSTRATETYPE as '保税类型',aa.STRUCKSTR2 as '产品结构' " & _
            ",count(aa.SUBSTRATEID) as 'wafer片数',wafer序号 = " & _
            "(STUFF((SELECT ',' + test.wafer号 " & _
            "FROM ( " + strsql1 + " ) test WHERE aa.来料时间 = test.来料时间 and aa.事业部 = test.事业部 and " & _
            "  aa.客户代码 = test.客户代码 and aa.客户机种名 = test.客户机种名 and aa.厂内机种名 = test.厂内机种名 and aa.PO = test.PO " & _
            "  and aa.lot号 = test.lot号 FOR XML PATH('')), 1,  1, '')) from ( " + strSql2 + ") aa group by aa.来料时间,aa.事业部,aa.客户代码, aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount,aa.SUBSTRATETYPE,aa.STRUCKSTR2 )cc " & _
            "left join ERPBASE..tblstocknum e on cc.lot号 = e.批号 and e.当前存量 > 0 and e.仓库编号 <> '54' where not exists (select * from erptemp..BZGDKL f where f.lot = cc.lot号) group by cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,cc.来料时间,cc.事业部, " & _
            "cc.客户代码 , cc.客户机种名, cc.厂内机种名, cc.PO, cc.lot号,e.仓库编号, cc.wafer片数,cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer序号 ) kk where kk.是否来料存放仓库 <> '尚未来料' order by kk.来料时间 desc,kk.事业部,kk.客户代码,kk.客户机种名,kk.厂内机种名,kk.PO,kk.lot号,kk.wafer片数,kk.是否来料存放仓库,kk.wafer序号 "
        End If
    Else
     strSql = "SELECT cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,case when e.仓库编号 is  null then " & _
            " '尚未来料'else e.仓库编号 end as '是否来料?存放仓库',cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer片数,cc.wafer序号 from ( " & _
            "SELECT '' as '选择',aa.来料时间,aa.事业部,aa.客户代码,aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount as 'WO过站die',aa.SUBSTRATETYPE as '保税类型',aa.STRUCKSTR2 as '产品结构' " & _
            ",count(aa.SUBSTRATEID) as 'wafer片数',wafer序号 = " & _
            "(STUFF((SELECT ',' + test.wafer号 " & _
            "FROM ( " + strsql1 + " ) test WHERE aa.来料时间 = test.来料时间 and aa.事业部 = test.事业部 and " & _
            "  aa.客户代码 = test.客户代码 and aa.客户机种名 = test.客户机种名 and aa.厂内机种名 = test.厂内机种名 and aa.PO = test.PO " & _
            "  and aa.lot号 = test.lot号 FOR XML PATH('')), 1,  1, '')) from ( " + strSql2 + ") aa group by aa.来料时间,aa.事业部,aa.客户代码, aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号,aa.passbincount,aa.SUBSTRATETYPE,aa.STRUCKSTR2 )cc " & _
            "left join ERPBASE..tblstocknum e on cc.lot号 = e.批号 and e.当前存量 > 0 and e.仓库编号 <> '54' where not exists (select * from erptemp..BZGDKL f where f.lot = cc.lot号) group by cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,cc.来料时间,cc.事业部, " & _
            "cc.客户代码 , cc.客户机种名, cc.厂内机种名, cc.PO, cc.lot号,e.仓库编号, cc.wafer片数,cc.WO过站die,cc.保税类型,cc.产品结构,cc.wafer序号 order by cc.来料时间 desc,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,e.仓库编号,cc.wafer序号 "
    End If
    strSqls = strSql  '赋给全局变量
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If

    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Function
    End If

End Function

Private Sub cmd_Click()
    Dim strSql2       As String
    Dim strSqlr  As String
    Dim rs1           As New ADODB.Recordset
    
    Dim i         As Integer
    Dim j         As Integer
    Dim waferid As String
    Dim substratetype As String
    Dim lot         As String
    Dim prelot  As String
    Dim count As Integer
    count = 0
    rcount = 0
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 8
                lot = Trim(.Text)
                If prelot <> lot Then
                    prelot = lot
                    strSqlr = "select a.lotid,a.substrateid,a.substratetype from mappingdatatest a where a.lotid = '" & lot & "'"
    
                       If INIadoCon.State <> adStateOpen Then
                            INIConnectSTART2
                        End If
    
                        If Cnn.State = 0 Then
                        ConOracle
                        End If
                    rs1.Open strSqlr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
                    If rs1.RecordCount > 0 Then
                        For j = 1 To rs1.RecordCount
                             If IsNull(rs1.Fields("substratetype")) Then
                                rs1.MoveNext
                             Else
                                 substratetype = rs1.Fields("substratetype")
                                 waferid = rs1.Fields("substrateid")
                                 strSql2 = "update ERPBASE..tblmappingData set substratetype = '" & substratetype & "' where lotid = '" & lot & "' and substrateid = '" & waferid & "'"
                                 AddSql2 (strSql2)
                                 count = count + 1
                                 rcount = rcount + 1
                             End If
                             rs1.MoveNext
                        Next
                        MsgBox "已更新！lot为:" & lot & ",次数为: " & rcount & ";"
                        rcount = 0
                    Else
                            MsgBox "该lot:" & lot & "未更新"
                    End If
                    rs1.Close
                End If
            End If
        Next i

    End With

    If count = 0 Then
        MsgBox "未选择"
    Else
        MsgBox "更新成功" & "orancle更新到SQL的记录数:" & count & "! "
    End If
    Query
End Sub

Private Sub cmdall_Click()
cmd.Visible = False
'    Dim strSql       As String
'
'    Dim strsql1       As String
'
'    Dim strsql2       As String
'
'    Dim rs           As New ADODB.Recordset
'
'    Dim SYB         As String
'
'    Dim KHJZ      As String
'
'    Dim KHDM         As String
'
'    Dim create_by       As String
'
'    Dim Create_time_start     As String
'
'    Dim Create_time_end   As String
'
'
'    SYB = Trim(CboSYB.Text)
'    KHJZ = Trim(CboKHJZ.Text)
'    KHDM = Trim(CboKHDM.Text)
'    Create_time_start = DTP(0).Value
'    Create_time_end = DTP(1).Value
'
'       strsql1 = "select  CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) as '来料时间', d.UPdateprice2 as '事业部', A.CUSTOMERSHORTNAME as '客户代码'," & _
'            "B.MPN_DESC as '客户机种名',B.mtrl_num as '厂内机种名',b.po_num as 'PO',B.SOURCE_BATCH_ID as 'lot号',right(A.SUBSTRATEID,2) as 'wafer号' " & _
'            "FROM ERPBASE..tblmappingData a ," & _
'            "ERPBASE..tblCustomerOI b,erptemp..tbltsvnpiproduct d " & _
'            "WHERE a.FLAG = 'Y' " & _
'            "AND CONVERT(VARCHAR(100),b.ID) = a.FILENAME " & _
'            "and d.customerptno1 = B.MPN_DESC "
'
'      strsql2 = "select  CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) as '来料时间', d.UPdateprice2 as '事业部', A.CUSTOMERSHORTNAME as '客户代码'," & _
'            "B.MPN_DESC as '客户机种名',B.mtrl_num as '厂内机种名',b.po_num as 'PO',B.SOURCE_BATCH_ID as 'lot号',A.SUBSTRATEID " & _
'            "FROM ERPBASE..tblmappingData a ," & _
'            "ERPBASE..tblCustomerOI b,erptemp..tbltsvnpiproduct d " & _
'            "WHERE a.FLAG = 'Y' " & _
'            "AND CONVERT(VARCHAR(100),b.ID) = a.FILENAME " & _
'            "and d.customerptno1 = B.MPN_DESC "
'
'
'    If SYB <> "" Then
'        strsql1 = strsql1 + " AND  d.UPdateprice2   = '" & SYB & "'  "
'        strsql2 = strsql2 + " AND  d.UPdateprice2   = '" & SYB & "'  "
'    End If
'
'    If KHJZ <> "" Then
'        strsql1 = strsql1 + " AND B.MPN_DESC = '" & KHJZ & "'  "
'        strsql2 = strsql2 + " AND B.MPN_DESC = '" & KHJZ & "'  "
'    End If
'
'    If KHDM <> "" Then
'        strsql1 = strsql1 + " AND A.CUSTOMERSHORTNAME  = '" & KHDM & "'  "
'        strsql2 = strsql2 + " AND A.CUSTOMERSHORTNAME  = '" & KHDM & "'  "
'    End If
'
'    If chkCheck1 = 1 Then
'        strsql1 = strsql1 + "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) >='" & Format(Create_time_start, "yyyy-mm-dd") & "' " & _
'                    "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) <='" & Format(Create_time_end, "yyyy-mm-dd") & "' "
'        strsql2 = strsql2 + "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) >='" & Format(Create_time_start, "yyyy-mm-dd") & "' " & _
'                    "AND CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) <='" & Format(Create_time_end, "yyyy-mm-dd") & "' "
'
'    End If
'
'    strsql1 = strsql1 + "and not EXISTS ( select * from erpdata .. tblTSVwaferlist c where c.waferid = A.SUBSTRATEID ) " & _
'                        "group by CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) , d.UPdateprice2 , A.CUSTOMERSHORTNAME ," & _
'                        "B.MPN_DESC ,B.mtrl_num,b.po_num ,B.SOURCE_BATCH_ID,A.SUBSTRATEID "
'
'    strsql2 = strsql2 + "and not EXISTS ( select * from erpdata .. tblTSVwaferlist c where c.waferid = A.SUBSTRATEID ) " & _
'                        "group by CONVERT(VARCHAR(100),a.QTECH_CREATED_DATE,23) , d.UPdateprice2 , A.CUSTOMERSHORTNAME ," & _
'                        "B.MPN_DESC ,B.mtrl_num,b.po_num ,B.SOURCE_BATCH_ID,A.SUBSTRATEID"
'
'    strSql = "SELECT cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,case when e.仓库编号 is  null then " & _
'            " '尚未来料'else e.仓库编号 end as '是否来料?存放仓库',cc.wafer片数,cc.wafer序号 from ( " & _
'            "SELECT '' as '选择',aa.来料时间,aa.事业部,aa.客户代码,aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号" & _
'            ",count(aa.SUBSTRATEID) as 'wafer片数',wafer序号 = " & _
'            "(STUFF((SELECT ',' + test.wafer号 " & _
'            "FROM ( " + strsql1 + " ) test WHERE aa.来料时间 = test.来料时间 and aa.事业部 = test.事业部 and " & _
'            "  aa.客户代码 = test.客户代码 and aa.客户机种名 = test.客户机种名 and aa.厂内机种名 = test.厂内机种名 and aa.PO = test.PO " & _
'            "  and aa.lot号 = test.lot号 FOR XML PATH('')), 1,  1, '')) from ( " + strsql2 + ") aa group by aa.来料时间,aa.事业部,aa.客户代码, aa.客户机种名,aa.厂内机种名,aa.PO,aa.lot号 )cc " & _
'            "left join ERPBASE..tblstocknum e on cc.lot号 = e.批号 where not exists (select * from erptemp..BZGDKL f where f.lot = cc.lot号) group by cc.选择,cc.来料时间,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,cc.来料时间,cc.事业部, " & _
'            "cc.客户代码 , cc.客户机种名, cc.厂内机种名, cc.PO, cc.lot号, cc.wafer片数, e.仓库编号, cc.wafer序号 order by cc.来料时间 desc,cc.事业部,cc.客户代码,cc.客户机种名,cc.厂内机种名,cc.PO,cc.lot号,cc.wafer片数,e.仓库编号,cc.wafer序号 "
'
If strSqls = "" Then
    MsgBox "没有导出条件"
    Exit Sub
End If
    SqlServerExporToExcel (strSqls)
End Sub

Private Sub CmdClear_Click()
Initial1
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
  Query
  Command2.Visible = True
  Command3.Visible = True
  Command4.Visible = True
  cmd.Visible = True
End Sub
Private Sub ListDataType(rs As ADODB.Recordset)
   Dim i As Long

    With fps(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .Text = 0
            .Col = 4
            .Lock = True
            .Col = 5
            .Lock = True
            .Col = 6
            .Lock = True
            .Col = 7
            .Lock = True
        
        Next
        
    End With
    rs.Close
End Sub
Private Sub ListDataType2(rs As ADODB.Recordset)
   Dim i As Long

    With Fps1(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With Fps1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .Text = 0
            .Col = 4
            .Lock = True
            .Col = 5
            .Lock = True
            .Col = 6
            .Lock = True
            .Col = 7
            .Lock = True
        
        Next
        
    End With
    rs.Close
End Sub

'初始化
Private Function Initial()
    InitSYB
    Initcustomershortname
    InitKHJZname
End Function

Private Sub InitSYB()
Set mainItemRS = GetSYBname()
Set CboSYB.RowSource = mainItemRS
CboSYB.ListField = mainItemRS("事业部").Name
CboSYB.BoundColumn = mainItemRS("事业部").Name

End Sub
Public Function GetSYBname() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
cmdStr = "SELECT d.UPdateprice2 as '事业部' from erptemp..tbltsvnpiproduct d group by d.UPdateprice2 order by d.UPdateprice2"
'cmdStr = "SELECT case when d.UPdateprice2 = '' then '未维护' else d.UPdateprice2 end as '事业部' from erptemp..tbltsvnpiproduct d group by d.UPdateprice2 order by d.UPdateprice2"
Set RSResult = getSqlStr(cmdStr)
Set GetSYBname = RSResult
End Function

Private Sub Initcustomershortname()
Set mainItemRS1 = Getcustomershortname()
Set CboKHDM.RowSource = mainItemRS1
CboKHDM.ListField = mainItemRS1("客户代码").Name
CboKHDM.BoundColumn = mainItemRS1("客户代码").Name

End Sub
Public Function Getcustomershortname() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
cmdStr = "select case when A.CUSTOMERSHORTNAME is null then 'NULL'  ELSE  A.CUSTOMERSHORTNAME end as '客户代码' " & _
        "FROM ERPBASE..tblmappingData a " & _
        "group by A.CUSTOMERSHORTNAME " & _
        "order by A.CUSTOMERSHORTNAME "

Set RSResult = getSqlStr(cmdStr)
Set Getcustomershortname = RSResult
End Function

Private Sub InitKHJZname()
Set mainItemRS1 = GetKHJZname()
Set CboKHJZ.RowSource = mainItemRS1
CboKHJZ.ListField = mainItemRS1("客户机种名").Name
CboKHJZ.BoundColumn = mainItemRS1("客户机种名").Name

End Sub
Public Function GetKHJZname() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
cmdStr = "select case when B.MPN_DESC is null then 'NULL' else B.MPN_DESC end as '客户机种名' " & _
        "from ERPBASE..tblCustomerOI b " & _
        "group by B.MPN_DESC " & _
        "order by B.MPN_DESC "

Set RSResult = getSqlStr(cmdStr)
Set GetKHJZname = RSResult
End Function


'Public Function GetJDCustomerName() As ADODB.Recordset
'
'Dim cmdStr As String
'Dim RSResult As New ADODB.Recordset
'cmdStr = "SELECT d.struckstr1 as '事业部' from erptemp..tbltsvnpiproduct d group by d.struckstr1 order by d.struckstr1"
'
'Set RSResult = getSqlStr(cmdStr)
'Set GetJDCustomerName = RSResult
'End Function

Private Function initSYBs()
'Dim cmdStr As String
'Dim rsSYB           As New ADODB.Recordset
'cmdStr = "select d.struckstr1 as '事业部' from erptemp..tbltsvnpiproduct d group by d.struckstr1 order by d.struckstr1"
'rsSYB.Open cmdStr, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
'    If Not rsSYB.EOF Then
'        CboCustomer.ListField = rsSYB.Fields("事业部")
'    Else
'        CboCustomer.ListField = ""
'        Exit Function
'    End If
'rsSYB.Close
End Function
Private Function CheckData()

MsgBox "功能尚未开放"

'    Dim i As Long
'
'    With Fps(0)
'        .MaxRows = 0
'        Set .DataSource = rs
'
'    End With
'
'    With Fps(0)
'
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1
'            .BackColor = &HFFFF&
'            .ColWidth(1) = 10
'            .CellType = CellTypeCheckBox
'            .Text = 0
'            .Col = 10
'            .Lock = False
'
'        Next
'
'    End With


End Function

Private Function Initial1()
    CboSYB.Text = ""
    CboKHJZ.Text = ""
    CboKHDM.Text = ""

End Function

Private Sub Command2_Click()
  Command2.Visible = False
  Command3.Visible = False
  Command4.Visible = False
  cmd.Visible = False
  Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    Dim i         As Integer
    
    Dim lot         As String
    Dim count As Integer
    

    count = 0
     
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 8
                lot = lot & Trim$("" & Trim(.Text)) & "','"
            End If

        Next i

    End With
    
    If lot = "" Then
        MsgBox "未选择lot号"
        Exit Sub
    End If

    lot = Mid(lot, 1, Len(lot) - 3)
      
    strSql = "select '' as '选择',a.入库单编号,a.批号,a.晶圆ID,a.良品数,a.不良数  FROM ERPBASE..tblToInRec_Wafer a WHERE a.批号 in ('" & lot & "')"
    strSqls = strSql  '赋给全局变量
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If

    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Sub
    End If

    MsgBox "查询完毕"
    
End Sub

Private Sub Command3_Click()
Command2.Visible = False
Command4.Visible = False
cmd.Visible = False
Dim strSql       As String
Dim rs           As New ADODB.Recordset

strSql = "select ' 'as 选择,lot,remark,createpeople,createtime from erptemp..BZGDKL"
strSqls = strSql  '赋给全局变量

  If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If

    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

  If Not rs.EOF Then
      Call ListDataType(rs)
  Else
      MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Sub
    End If
End Sub

Private Sub Command4_Click()
cmd.Visible = False
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    Dim i         As Integer
    
    Dim thedata As String
    Dim lot         As String
    Dim SYZ         As String
    Dim BZ        As String
    Dim count As Integer
    
    SYZ = gUserName
    count = 0
    
    BZ = "该lot无法开立工单"
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 8
                lot = Trim(.Text)
                
                strSql = "select * from erptemp..BZGDKL where lot = '" & lot & "'"
                rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                    MsgBox "已标记！lot为:" & lot & ""
                    count = count + 1
                ElseIf rs.RecordCount = 0 Then
                    thedata = InputBox("请输入标记lot为:" & lot & "的备注信息,取消请点击""取消""", "请输入备注:")
                    If MsgBox("提示：该lot:" & lot & "会被标记,是否继续?", vbOKCancel, "提示") = vbOK Then
                        If thedate = "" Then
                            strSql = "INSERT INTO erptemp.dbo.BZGDKL (lot, remark, createpeople, createtime) VALUES('" & lot & "','" & BZ & "','" & SYZ & "',(getdate()))"
                        Else
                            strSql = "INSERT INTO erptemp.dbo.BZGDKL (lot, remark, createpeople, createtime) VALUES('" & lot & "','" & thedata & "','" & SYZ & "',(getdate()))"
                        End If
                        AddSql2 (strSql)
                        count = count + 1
                    Else
                        MsgBox "该lot:" & lot & "已跳过标记！"
                    End If
                Else
                    MsgBox "标记失败！"
                End If
                rs.Close
            End If

        Next i

    End With

    If count = 0 Then
        MsgBox "未选择"
    Else
        MsgBox "标记成功" & "标记记录数" & count & "! "
    
    End If
    Query
End Sub

Private Sub Form_Load()
Command2.Visible = True
Command4.Visible = True
Command3.Visible = True
cmd.Visible = False
DTP(0).Value = DATE
DTP(1).Value = DATE
Initial

End Sub

Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim lot      As String

    '点击把选择的单号都选上
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
 
       With fps(0)

        .Col = fpSDetail.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
        If Val(.Value) = 1 Then
            '将所有一样的单号选择上
            .Col = fpSDetail.e_DJBH
            .Row = Row
            lot = Trim$(.Text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.Text) = lot Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
            
        Else
            '将所有一样的单号选择上
            .Col = fpSDetail.e_DJBH
            .Row = Row
            lot = Trim$(.Text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.Text) = lot Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
            
        End If
        
    End With
    
End Sub

