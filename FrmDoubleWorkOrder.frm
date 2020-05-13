VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form FrmDoubleWorkOrder 
   Caption         =   "重工订单批量导入生成"
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11220
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
   ScaleHeight     =   11145
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11295
      Begin VB.CommandButton btnSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "保存"
         Enabled         =   0   'False
         Height          =   735
         Left            =   960
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmDoubleWorkOrder.frx":0000
         Picture         =   "FrmDoubleWorkOrder.frx":3072
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   5175
      End
      Begin VB.CommandButton btnImport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "导入"
         Height          =   735
         Left            =   0
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmDoubleWorkOrder.frx":4D6C
         Picture         =   "FrmDoubleWorkOrder.frx":7DDE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "退出"
         Height          =   735
         Left            =   1800
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmDoubleWorkOrder.frx":9AD8
         Picture         =   "FrmDoubleWorkOrder.frx":CB4A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   9975
         Left            =   0
         TabIndex        =   1
         Top             =   960
         Width           =   6855
         _Version        =   524288
         _ExtentX        =   12091
         _ExtentY        =   17595
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
         SpreadDesigner  =   "FrmDoubleWorkOrder.frx":F2EC
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件路径:"
         Height          =   195
         Left            =   4680
         TabIndex        =   5
         Top             =   480
         Width           =   780
      End
   End
End
Attribute VB_Name = "FrmDoubleWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jSuccess  As Long

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnImport_Click()
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

Call ImportExl

End Sub

Private Sub ImportExl()

On Error GoTo ErrHandle

Dim VBExcel     As Excel.Application
Dim xlBook      As Excel.Workbook
Dim xlSheet     As Excel.Worksheet
Dim i           As Integer
Dim j           As Integer
Dim strChar     As String
Dim strTmp      As String

MousePointer = 11
Fps.MaxRows = 0
Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(txtPath.text)
Set xlSheet = xlBook.Worksheets(1)
If xlSheet.Range("A1").CurrentRegion.Columns.count <> 3 Then
    MousePointer = 0
    MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
    GoTo EXITPRO
    Exit Sub

End If

With Fps

    For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strTmp = Trim(xlSheet.Range("A" & i).Value)
        If Len(strTmp) > 0 Then
            If i <> 1 Then .MaxRows = .MaxRows + 1

            For j = 1 To 7
                If j > 26 Then
                    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                Else
                    strChar = Chr(96 + j)

                End If

                If i = 1 Then
                    .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))
                Else
                    .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))

                End If

            Next

        End If

    Next

End With

MousePointer = 0
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing

 
MsgBox "总共:" & Fps.MaxRows & "笔数据,请点击保存完成导入或退出", vbInformation, "保存提示"
btnImport.Enabled = False
btnSave.Enabled = True

Exit Sub
EXITPRO:

On Error Resume Next

MousePointer = 0
If Not VBExcel Is Nothing Then
    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    VBExcel.Quit

End If

Exit Sub
ErrHandle:
GoTo EXITPRO

End Sub

Private Sub SaveData()
Dim strLotID    As String
Dim strWaferNo  As String
Dim strWaferIDNew As String
Dim strWaferIDOld As String
Dim strGoodDies As String
Dim lGoodDies As Long
Dim i           As Long

Dim strSql As String

jSuccess = 0
With Fps

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        strLotID = Trim$("" & .text)
        .Col = 2
        strWaferNo = Trim$("" & .text)
        .Col = 3
        strGoodDies = Trim$("" & .text)
        If strLotID = "" Or strWaferNo = "" Or strGoodDies = "" Then
            GoTo Continue

        End If
        
        lGoodDies = CLng(strGoodDies)
        If Left$(strWaferNo, 1) = "0" Then
            strWaferNo = Replace$(strWaferNo, "0", "", 1, 1)

        End If

        strSql = "select distinct b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and to_number(b.wafer_id) = '" & strWaferNo & "' and to_char(a.id) = b.filename and a.source_batch_id = b.lotid " & " and a.invflag = 0 and instr(b.substrateid, '+') > 0 and not exists (select 1 from ib_waferlist c where b.substrateid = c.waferid) "
        If Get_OracleCnt(strSql) > 0 Then
            GoTo Continue
        End If
        
        strSql = "select substrateid || '+' as substrateid from mappingdatatest where to_number(wafer_id) = '" & strWaferNo & "' and lotid = '" & strLotID & "' order by to_number(filename) desc"
        strWaferIDNew = Get_OracleStr(strSql)
        strWaferIDOld = Left(strWaferIDNew, Len(strWaferIDNew) - 1)
        
        Call InsertToDB(strWaferIDOld, strWaferIDNew, strWaferNo, lGoodDies, 0, strLotID)
        
Continue:
    Next i

End With

If jSuccess > 0 Then
    MsgBox "已经生成:" & jSuccess & "笔重工WO, 可以开立重工工单", vbInformation, "提示"
Else
    MsgBox "未生成重工WO,请确认", vbCritical, "警告"

End If

End Sub

Private Sub InsertToDB(strWaferID As String, _
                       strWaferIDNew As String, _
                       strWaferNo As String, _
                       lGoodDIe As Long, _
                       lNGDie As Long, _
                       strLotID As String)
Dim cmdStr     As String
Dim cmdStr2    As String
Dim sSeqID     As Long
Dim strCusPN   As String
Dim strCusCode As String
Dim strMark    As String
Dim rs         As New ADODB.Recordset
Dim lID        As Long

On Error GoTo ERRORON

Cnn.BeginTrans
INIadoCon.BeginTrans
' 获取ID
lID = Get_OracleNo("select filename from mappingdatatest where to_number(wafer_id) = '" & strWaferNo & "' and lotid = '" & strLotID & "' order by to_number(filename) desc ")
' 检查
cmdStr = "select * from mappingDataTest where substrateid = '" & strWaferIDNew & "'"
If rs.State = adStateOpen Then rs.Close
rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    MsgBox strWaferIDNew & "已经存在,请不要添加同样的WAFERID", vbCritical, "警告"
    Exit Sub

End If

' 插入
sSeqID = GetMaxID()
cmdStr = " insert into CustomerOItbl_test(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
   " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
   " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
   " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
   " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
   " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
   " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
   " select  distinct '" & sSeqID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
   " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
   " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
   " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
   " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
   " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
   " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
   " ct.custom_part_no,'T','" & gUserName & "',sysdate,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,ct.wafer_visual_inspect, " & _
       " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from CustomerOItbl_test ct, MAPPINGDATATEST mt  where ct.id = " & lID & " and to_char(ct.id) = mt.filename "
cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI](id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
   " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
   " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
   " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
   " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
   " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
   " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
   " select distinct  '" & sSeqID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
   " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
   " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
   " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
   " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
   " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
   " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
   " ct.custom_part_no,'T','" & gUserName & "',GetDate(),ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.wafer_visual_inspect, " & _
   " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from [ERPBASE].[dbo].[tblCustomerOI] ct, [ERPBASE].[dbo].[tblmappingData] mt  where ct.id = " & lID & " and convert(varchar,ct.id) = mt.filename"
AddSql (cmdStr)
AddSql2 (cmdStr2)
cmdStr = "insert into mappingDataTest (id, substrateid, productid, lotid, Wafer_ID, passbincount, failbincount, CustomerShortName, flag, Qtech_Created_By, Qtech_Created_Date,filename) " & " select  mappingData_SEQ.Nextval, '" & strWaferIDNew & "', productid, lotid,Wafer_ID,  '" & lGoodDIe & "', '" & lNGDie & "', CustomerShortName, 'T',  '" & gUserName & "', sysdate, '" & sSeqID & "' " & " from MAPPINGDATATEST  where filename =  '" & lID & "' and to_number(wafer_id) = '" & strWaferNo & "' "
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & " select  '" & strWaferIDNew & "', productid, lotid,Wafer_ID, '" & lGoodDIe & "', '" & lNGDie & "', CustomerShortName, 'T',  '" & gUserName & "', GETDATE(), '" & sSeqID & "' " & " from [ERPBASE].[dbo].[tblmappingData]  where filename = '" & lID & "' and wafer_id in ('" & strWaferNo & "', '0' + '" & strWaferNo & "') "
AddSql (cmdStr)
AddSql2 (cmdStr2)
Cnn.CommitTrans
INIadoCon.CommitTrans
jSuccess = jSuccess + 1

Exit Sub
ERRORON:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
MsgBox "订单生成失败:" & Err.DESCRIPTION, vbInformation, "提示:"

End Sub

Private Sub btnSave_Click()
btnSave.Enabled = False
Call SaveData

btnImport.Enabled = True

End Sub
