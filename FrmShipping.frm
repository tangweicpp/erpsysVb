VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmShipping 
   Caption         =   "客户Shipping资料，客户挑料信息上传"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12525
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "EQ"
      TabPicture(0)   =   "FrmShipping.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SemTech"
      TabPicture(1)   =   "FrmShipping.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "RDA挑料信息"
      TabPicture(2)   =   "FrmShipping.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraPickingListDN"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "GD108"
      TabPicture(3)   =   "FrmShipping.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "上传"
         Height          =   5535
         Left            =   -74520
         TabIndex        =   22
         Top             =   600
         Width           =   10575
         Begin VB.CommandButton cmdEXit 
            Caption         =   "退出"
            Height          =   360
            Left            =   5040
            TabIndex        =   28
            Top             =   4080
            Width           =   990
         End
         Begin VB.CommandButton cmdExportGD 
            Caption         =   "导出"
            Height          =   360
            Left            =   2880
            TabIndex        =   27
            Top             =   4080
            Width           =   990
         End
         Begin VB.CommandButton cmdUploadGD 
            Caption         =   "上传"
            Height          =   360
            Left            =   840
            TabIndex        =   26
            Top             =   4080
            Width           =   990
         End
         Begin VB.CommandButton cmdDia 
            Caption         =   ".."
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   8520
            TabIndex        =   25
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtGD108 
            Height          =   3015
            Left            =   600
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   7695
         End
         Begin MSComDlg.CommonDialog CommonDialog4 
            Left            =   8880
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblXls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   600
            TabIndex        =   23
            Top             =   360
            Width           =   1530
         End
      End
      Begin VB.Frame FraPickingListDN 
         Caption         =   "上传"
         Height          =   2535
         Left            =   -73320
         TabIndex        =   13
         Top             =   1740
         Width           =   7095
         Begin VB.CommandButton cmdOutput 
            Caption         =   "导出报表"
            Height          =   480
            Index           =   1
            Left            =   3720
            TabIndex        =   17
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   16
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton cmd 
            Caption         =   ".."
            Height          =   495
            Index           =   0
            Left            =   6120
            TabIndex        =   15
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   4935
         End
         Begin MSComDlg.CommonDialog CommonDialog3 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblXls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   18
            Top             =   480
            Width           =   1530
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "PickingList DN上传"
         Height          =   5295
         Left            =   -74880
         TabIndex        =   7
         Top             =   1380
         Width           =   11775
         Begin VB.TextBox txtDate 
            Height          =   375
            Left            =   7440
            TabIndex        =   31
            Top             =   3480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "删除"
            Height          =   360
            Left            =   3480
            TabIndex        =   21
            Top             =   4440
            Width           =   990
         End
         Begin VB.TextBox txtDN 
            Height          =   375
            Left            =   1200
            TabIndex        =   20
            Top             =   4440
            Width           =   1935
         End
         Begin VB.TextBox Txtsemtech 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   840
            Width           =   8295
         End
         Begin VB.CommandButton Command3 
            Caption         =   ".."
            Height          =   495
            Left            =   9240
            TabIndex        =   10
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdUpload 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   720
            TabIndex        =   9
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   2400
            TabIndex        =   8
            Top             =   3600
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Index           =   0
            Left            =   7440
            TabIndex        =   29
            Top             =   3840
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
            Format          =   166789121
            CurrentDate     =   43271
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出货日期:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6240
            TabIndex        =   30
            Top             =   3900
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DN:"
            Height          =   180
            Left            =   720
            TabIndex        =   19
            Top             =   4560
            Width           =   270
         End
         Begin VB.Label lblXls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Width           =   1530
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "选择待上传的文件"
         Height          =   2535
         Left            =   480
         TabIndex        =   1
         Top             =   1140
         Width           =   7095
         Begin VB.CommandButton Command8 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   3720
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   4
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   3
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   4935
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblXls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xlsx："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   6
            Top             =   480
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "FrmShipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vtDataTemp As ShippingData

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmd_Click(Index As Integer)

    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog3.Filter = "EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    
    CommonDialog3.ShowOpen
    '得到文件名
    FName = CommonDialog3.filename

    If FName <> "" Then
        Text1.text = FName
    End If

End Sub

Private Sub cmdDel_Click()
MsgBox "该处不可删除, 请至上传37DN", vbCritical, "警告"

Exit Sub
    Dim sDN As String

    sDN = Trim$(txtDN.text)

    If sDN = "" Then
        MsgBox "请输入DN号"

        Exit Sub

    End If

    Dim sOra As String

    Dim sSql As String
    
    AddSql ("insert into CUSTOMERSHIPPINGUPTBL_BAK select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "' ")
    MsgBox "DN备份成功", vbInformation, "提示"
    
    sOra = "delete from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
    sSql = "delete from  [ERPBASE].[dbo].[tblCustomerShippingUp] where Delivery = '" & sDN & "'"

    AddSql (sOra)
    AddSql2 (sSql)

    MsgBox "已成功删除DN:" & sDN, vbInformation, "提示"

End Sub

Private Sub cmdDia_Click()

    Dim FName

    CommonDialog4.flags = cdlOFNAllowMultiselect Or cdlOFNExplorer

    CommonDialog4.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog4.ShowOpen

    FName = CommonDialog4.filename

    If FName <> "" Then
        txtGD108.text = Replace(FName, Chr(0), ",")

    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOutput_Click(Index As Integer)

    Dim strora As String

    strora = "select * from RDABACKDATA"

    ExporToExcel (strora)
End Sub

Private Sub cmdDB_Click()

    If Text1.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub

    End If

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(Text1.text)    '打开文件
    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '2)判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 6 Then
        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"

        Exit Sub

    End If
        
    ' 定义变量
    Dim number   As String

    Dim customer As String

    Dim device   As String

    Dim waferid  As String

    Dim shipment As String

    Dim Userid   As String

    Dim cmdOra   As String

    Dim company  As String

    Userid = gUserName

    ' 遍历表格
    ' 第2行开始,循环更换行号
    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        ' 查询一行的值
        ' 第1列开始,循环增加列数
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & I).Value   '临时保存值

            If j = 1 Then
                number = Trim(tempVal)
            End If
            
            If j = 2 Then
                customer = Trim(tempVal)
            End If
            
            If j = 3 Then
                device = Trim(tempVal)
            End If
            
            If j = 4 Then
                waferid = Trim(tempVal)
            End If
            
            If j = 5 Then
                shipment = Trim(tempVal)
            End If
        
            If j = 6 Then
                company = Trim$(tempVal)
            End If
            
        Next j
    
        ' 判断是否已经上传
        If JudgeRDA(waferid, customer) Then
            MsgBox "已存在该片数据, 请不要重复上传"

            Exit Sub

        End If
    
        ' 上传
        cmdOra = "insert into RDABACKDATA values('" + customer + "', '" + device + "', '" + waferid + "', '" + shipment + "', sysdate, '" + Userid + "', '" + company + "')"
    
        AddSql (cmdOra)
    
        ' add new
        Dim sOra As String
    
        sOra = "select mes_dn_pkg.MES_DN_RDA（'" & waferid & " '） from dual"
        Get_OracleRs (sOra)
    
    Next I

    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    MsgBox "已成功上传", vbInformation, "友情提示"

End Sub

Private Sub cmdUploadGD_Click()
    SumCount = 0
    ErrorInf = ""

    If txtGD108.text = "" Then
        MsgBox "先选择待上传的文件", vbCritical, "警告"

        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Trim(txtGD108.text)

    Dim dirtemp() As String

    Dim I         As Integer
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For I = 1 To UBound(dirtemp)
            UpGD108 (dirtemp(0) + "\" + dirtemp(I))

            Sleep (1000)
        Next
        
    Else
        
        UpGD108 (filename)

    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "提示"
    Else
        MsgBox "没有成功上传"
    End If
    
End Sub

Private Sub Command1_Click()
    '导出
  
    ExporToExcel (" select lastupdatedate as 出货日期,ID,Delivery ,ItemNo ,DeliveryCreationDate ,Plant    , SalesDocument  , " & "  SOItemNo , Material  ,MarketingPN ,MaterialDescription ,PlannedGIDate ,CustomerPartNumber  ,ShipToName  , " & "   ShipToCustomer ,PurchasingDocNo ,DateCodeRestrictions  ,LabelRequirement ,ReLabelInstructions ,ShipToStreet1 ,ShipToStreet2 ,ShipToStreet3 ," & " City  ,State ,PostalCode ,CountryKey , ContactName ,Phone  ,Fax   , FreightForwarder  , " & " ShippingInstruction ,AdditionalComments ,StorageLocation , BatchNumber ,Quantity ,VolumeWeight ,GrossWeight ,Netweight , " & "  UoMForWeight ,NoOfCartons ,VendorLotNumber ,ShelfLocation ,BOLOrAirwayBillNo ,ActualShippingDate ,PackagingDetails ,PackingStatus  , " & " PickingStatus , CustomerCalendar ,customershortname  ,FLAG   ,CREATEDBY ,CREATEDDATE  from CUSTOMERSHIPPINGUPTBL order by CREATEDDATE desc ")
    
End Sub

Private Function checkDNHistory(strSemtech As String) As Boolean

    Dim strFileName   As String

    Dim dirName       As String

    Dim con           As New ADODB.Connection

    Dim rs            As New ADODB.Recordset

    Dim I             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim temp2         As String

    Dim tempVal       As String

    Dim WV_inspect    As String

    Dim Comp_codeTemp As String

    Dim dn_job        As String

    Dim dn_job1       As String

    Dim dn_job_qty    As Long

    Dim dn_job_qty1   As Long
    
    Dim strWaferID    As String
    
    Dim strsql        As String
    
    Dim strBand       As String, strBandThis As String
    
    Dim strJOBNo      As String

    checkDNHistory = False

    If InStrRev(strSemtech, "\") > 0 Then
        strFileName = Mid(strSemtech, InStrRev(strSemtech, "\") + 1)
        dirName = Mid$(strSemtech, 1, InStrRev(strSemtech, "\"))

    End If

    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(strSemtech)
    Set xlSheet = xlBook.Worksheets(1)

    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        For j = 1 To 46

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If

            tempVal = Replace(Replace(xlSheet.Range(strChar & I).Value, ",", " "), "，", " ")
        
            If j = 1 Then
                dnTemp = Trim(tempVal)
            
                If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & dnTemp & "' ") <> 0 Then
            
                    MsgBox "系统存在相同的DN, 此次不允许上传, 请联系IT", vbInformation, "友情提示"
            
                    Exit Function
            
                End If
            
            End If
            
            If j = 32 Then
                strJOBNo = Trim$(tempVal)
                
                strsql = "select replace(aa.substrateid, '+','') as waferid from mappingdatatest aa  inner join customeroitbl_test bb on to_char(bb.id) = aa.filename and aa.lotid = bb.source_batch_id where bb.test_mtrl_desc = '" & strJOBNo & "'  "
                strWaferID = Get_OracleStr(strsql)
                
                strsql = "select substratetype from mappingdatatest where substrateid = '" & strWaferID & "' and substratetype is not null "
                strBandThis = Get_OracleStr(strsql)
                
                If strBandThis <> "" Then
                    
                    If strBand = "" Then
                        strBand = strBandThis
                    Else

                        If strBandThis <> strBand Then
                            MsgBox "保税和非保税不可以混合进一票DN", vbInformation, "提示"
                            Exit Function

                        End If
                        
                    End If
                
                End If
                
            End If
    
        Next j
    
    Next I

    checkDNHistory = True

End Function

Private Sub Up37DN(strSemtech As String)

If checkDNHistory(strSemtech) = False Then
    Exit Sub

End If

SumCount = 0
Dim source_batch_id_Temp As String
Dim dnTemp               As String
Dim dnitemTemp           As String

If Txtsemtech.text = "" Then
    MsgBox "先选择待上传的文件", vbInformation, "提示"
    Exit Sub

End If

Dim dirName  As String
Dim filename As String

If InStrRev(strSemtech, "\") > 0 Then
    strFileName = Mid(strSemtech, InStrRev(strSemtech, "\") + 1)
    dirName = Mid$(strSemtech, 1, InStrRev(strSemtech, "\"))

End If

Dim con As New ADODB.Connection
Dim rs  As New ADODB.Recordset
Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(strSemtech)
Set xlSheet = xlBook.Worksheets(1)
'    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 52 Then
'
'        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
'
'        Exit Sub
'
'    End If
Dim I             As Integer
Dim j             As Integer
Dim id            As Long
Dim TEMP          As String
Dim temp2         As String
Dim temp3         As String
Dim tempVal       As String
Dim WV_inspect    As String
Dim Comp_codeTemp As String
Dim dn_job        As String
Dim dn_job1       As String
Dim dn_job_qty    As Long
Dim dn_job_qty1   As Long
Dim dn_ship       As String
Dim strChk1       As String, strChk2 As String, strChk3 As String
dn_job = ""
dn_job1 = ""
SumCount = 0

For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
    TEMP = ""
    source_batch_id_Temp = ""
    temp3 = ""
    
    For j = 1 To 46

        If j > 26 Then
            strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
        Else
            strChar = Chr(96 + j)

        End If

        tempVal = Replace(Replace(xlSheet.Range(strChar & I).Value, ",", " "), "，", " ")

        If j = 34 Or j = 35 Or j = 36 Or j = 37 Then
            If tempVal = "" Then
                TEMP = TEMP & "," & "0"
            Else
                TEMP = TEMP & "," & newStr("" & tempVal)

            End If

        ElseIf j = 19 Then
            TEMP = TEMP & "," & Replace(newStr("" & tempVal), ",", " ")
        Else
            TEMP = TEMP & "," & newStr("" & tempVal)

        End If

        If j = 1 Then
            dnTemp = tempVal

        End If

        If j = 2 Then
            dnitemTemp = tempVal

        End If

        If j = 32 Then
            dn_job = tempVal

        End If

        If j = 33 Then
            dn_job_qty = tempVal

        End If

        If j = 16 Then
            dn_ship = tempVal

        End If

    Next j

    For j = 56 To 58

        If j > 26 Then
            strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
        Else
            strChar = Chr(96 + j)

        End If

        tempVal = Replace(Replace(xlSheet.Range(strChar & I).Value, ",", " "), "，", " ")
        temp3 = temp3 & "," & newStr("" & tempVal)

        If j = 56 Then
            strChk1 = tempVal

        End If

        If j = 57 Then
            strChk2 = tempVal

        End If

        If j = 58 Then
            strChk3 = tempVal

        End If

    Next j

    If dn_ship = "ADD SAMSUNG E2 LABEL" And (strChk1 = "" Or strChk2 = "" Or strChk3 = "") And xlSheet.Range("A1").CurrentRegion.Columns.count <> 52 Then
        MsgBox "37出三星的挑料信息最后三列不可为空", vbInformation, "提示"
        Exit Sub

    End If

    id = GetshippingMaxID()
    TEMP = id & TEMP
    temp2 = UCase(TEMP & ",'37' ,'Y','" & gUserName & "',GETDATE(),'','" & txtDate.text & "'" & ",'','','','','','','','',''" & temp3)
    TEMP = UCase$(TEMP & ",'37','Y','" & gUserName & "',sysdate,'','" & txtDate.text & "'" & ",'','','','','','','','',''" & temp3)

    If (JudgeFlagStautsShipingUpjob(dnTemp, dn_job)) Then
        Call AddShippingUPDATE(dnTemp, dn_job, dn_job_qty)
    Else
        Call AddShippingUP(TEMP, temp2)

    End If

    sOra = "select mes_dn_pkg.MES_DN_37('" & dnTemp & "') from dual"
    AddSql (sOra)
NextRecord2:
Next I

End Sub

Private Sub UpGD108(strFileName As String)

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    Set VBExcel = CreateObject("excel.application")

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(strFileName)

    Set xlSheet = xlBook.Worksheets(1)



    Dim I             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim temp2         As String

    Dim tempVal       As String

    Dim WV_inspect    As String

    Dim Comp_codeTemp As String

    Dim dn_job        As String

    Dim dn_job1       As String

    Dim dn_job_qty    As Long

    Dim dn_job_qty1   As Long

    dn_job = ""
    dn_job1 = ""

    SumCount = 0

    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        TEMP = ""
        source_batch_id_Temp = ""
        
        For j = 1 To 14
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If
      
            tempVal = Replace(Replace(xlSheet.Range(strChar & I).Value, ",", " "), "，", " ")
                 
    

            If j = 6 Then
                cgd = Trim(tempVal)
            End If
        
            If j = 13 Then
                dn_job = Trim(tempVal)

            End If
        
        
        
     
          
    
        Next j
        
        If dn_job <> "" And cgd <> "" Then
             AddSql2 ("UPDATE tblcpurdatasub SET 单价 = '" & dn_job & "'  where 采购单编号  = '" & cgd & "' ")

        End If
        
          
NextRecord2:
       
    Next I


End Sub

Private Sub cmdUpload_Click()
MsgBox "该处不可上传, 请至上传37DN", vbCritical, "警告"

Exit Sub
If txtDate.text = "" Then
    MsgBox "请选择出货日期", vbInformation, "提示"
    Exit Sub
End If

    SumCount = 0
    ErrorInf = ""

    If Txtsemtech.text = "" Then
        MsgBox "先选择待上传的文件", vbCritical, "警告"

        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Txtsemtech.text

    Dim dirtemp() As String

    Dim I         As Integer
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For I = 1 To UBound(dirtemp)
            Up37DN (dirtemp(0) + "\" + dirtemp(I))

            Sleep (1000)
        Next
        
    Else
        
        Up37DN (filename)

    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "提示"
    Else
        MsgBox "没有成功上传"
    End If
    
End Sub

Private Sub Command3_Click()

    Dim FName

    CommonDialog1.flags = cdlOFNAllowMultiselect Or cdlOFNExplorer

    CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls"
    
    CommonDialog1.ShowOpen

    FName = CommonDialog1.filename

    If FName <> "" Then
        Txtsemtech.text = Replace(FName, Chr(0), ",")

    End If
    
End Sub

Private Sub Command6_Click()

    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.filename

    If FName <> "" Then
        Text3.text = FName
    End If

End Sub

Private Sub Command7_Click()

    UploadVTData

End Sub

Private Sub UploadVTData()

    '上传资料

    Dim source_batch_id_Temp As String

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 23 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"

        Exit Sub

    End If

    Dim I       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0
    BCResultFlag = False

    vtDataTemp.CreatedByTemp = gUserName

    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
 
        TEMP = ""
        source_batch_id_Temp = ""

        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & I).Value   '临时保存值
        
            If j = 1 Then
        
                vtDataTemp.notemp = Trim(tempVal)
            
            ElseIf j = 2 Then
                vtDataTemp.SubConPOTemp = Trim(tempVal)
            
            ElseIf j = 3 Then
                vtDataTemp.itemTemp = Trim(tempVal)
            
            ElseIf j = 4 Then
                vtDataTemp.QuantityTemp = Trim(tempVal)
            
            ElseIf j = 5 Then
                vtDataTemp.devicetemp = Trim(tempVal)
                '-------------
            
            ElseIf j = 6 Then
                vtDataTemp.SPATemp = Trim(tempVal)
            
            ElseIf j = 7 Then
                vtDataTemp.CSDTemp = Trim(tempVal)
            
            ElseIf j = 8 Then
       
                vtDataTemp.lottemp = Trim(tempVal)
            
            ElseIf j = 9 Then
            
                vtDataTemp.DateCode1Temp = Trim(tempVal)
            
            ElseIf j = 10 Then
                vtDataTemp.DeliveryNameTemp = Trim(tempVal)
            
            ElseIf j = 11 Then
                vtDataTemp.DeliveryAddressTemp = Trim(tempVal)
            
            ElseIf j = 12 Then
                vtDataTemp.WarehouseTemp = Trim(tempVal)
            
            ElseIf j = 13 Then
                vtDataTemp.LocationTemp = Trim(tempVal)
            
            ElseIf j = 14 Then
                vtDataTemp.ModeOfDeliveryTemp = Trim(tempVal)
            
            ElseIf j = 15 Then
                vtDataTemp.dateCodeTemp = Trim(tempVal)
            
                '-------------
            
            ElseIf j = 16 Then
                vtDataTemp.soTemp = Trim(tempVal)
            
            ElseIf j = 17 Then
       
                vtDataTemp.CarrierNotesTemp = Trim(tempVal)
            
            ElseIf j = 18 Then
            
                vtDataTemp.lineTemp = Trim(tempVal)
            
            ElseIf j = 19 Then
                vtDataTemp.ScheduleLineTemp = Trim(tempVal)
            
            ElseIf j = 20 Then
                vtDataTemp.CustPNTemp = Trim(tempVal)
            
            ElseIf j = 21 Then
                vtDataTemp.CountryDistributorTemp = Trim(tempVal)
            
            ElseIf j = 22 Then
                vtDataTemp.customerTemp = Trim(tempVal)
            
            ElseIf j = 23 Then
                vtDataTemp.customerPoTemp = Trim(tempVal)
            
            End If

        Next j

        vtDataTemp.idTemp = GetEQShippingMaxID()
        Call AddEQShipping(vtDataTemp)
    
        ' add new
        Dim sOra As String

        sOra = "select mes_dn_pkg.MES_DN_EQ('" & vtDataTemp.SubConPOTemp & "','" & vtDataTemp.customerPoTemp & "','" & vtDataTemp.lineTemp & "') from dual"
        AddSql (sOra)
    
        SumCount = SumCount + 1

        '上传到DB
NextRecord2:

    Next I
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else

        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"

            Exit Sub

        End If
    
    End If

End Sub

Private Sub Command8_Click()

    ExporToExcel ("  select  ID ,NO ,SUBCONPO ,ITEM ,QUANTITY ,DEVICE ,SPA ,CSD ,LOT  ,DATECODE1 ,DELIVERYNAME ,DELIVERYADDRESS ,WAREHOUSE  ," & "   Location , MODEOFDELIVERY, DateCode, SO, CARRIERNOTES, Line, SCHEDULELINE, CUSTPN, COUNTRYDISTRIBUTOR, Customer, CUSTOMERPO ,FLAG, CREATEDBY, CREATEDDATE " & " From customershippingtbl order by id  ")

End Sub

Private Function newStr(TEMP As String)

    If TEMP <> "" Then
        If InStr(TEMP, "'") > 0 Then
            newStr = "'" & Replace(TEMP, "'", "") & "'"
   
        Else
            newStr = "'" & TEMP & "'"
   
        End If

    Else

        newStr = "''"

    End If

End Function

Private Sub DTPicker1_CHANGE(Index As Integer)
txtDate.text = DTPicker1(0).Value
End Sub

Private Sub DTPicker1_Click(Index As Integer)
txtDate.text = DTPicker1(0).Value

End Sub

Private Sub Form_Load()
DTPicker1(0).Value = Format(Now(), "yyyy-MM-dd")

End Sub
