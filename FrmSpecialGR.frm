VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmSpecialGR 
   Caption         =   "0 good die GR 生成"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20265
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
   ScaleHeight     =   9810
   ScaleWidth      =   20265
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "数量修改"
      Height          =   1215
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   11055
      Begin VB.CommandButton CmdDel 
         Caption         =   "删除"
         Height          =   480
         Left            =   9360
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdModify 
         Caption         =   "修改"
         Height          =   480
         Left            =   7680
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtAfterDie 
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtAfterPiece 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtBeforDie 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtBeforPiece 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新RejectQty："
         Height          =   195
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新片数："
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原RejectQty："
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原片数："
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成CSV格式GR"
      Height          =   465
      Left            =   16440
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "添加"
      Height          =   480
      Left            =   14760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtBaoFei 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   8055
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7455
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   19935
      _Version        =   196613
      _ExtentX        =   35163
      _ExtentY        =   13150
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
      SpreadDesigner  =   "FrmSpecialGR.frx":0000
      TextTip         =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "有多个WaferID时，请用"",""分隔开"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5160
      TabIndex        =   18
      Top             =   480
      Width           =   2595
   End
   Begin VB.Label LblBaoFei 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报废WaferID："
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label LblLotID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID："
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "FrmSpecialGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
   ' E_SeqId = 1                '序号
    E_PO_Num = 1                'PO_Num
    E_PO_Item                 'PO_Item
    E_Previous_Batch_ID       'Previous_Batch_ID
    E_Previous_Mtrl_Num       'Previous_Mtrl_Num
    E_Batch_ID                'Batch_ID
    E_Mtrl_num                'Mtrl_num
    E_Mtrl_desc               'Mtrl_desc
    E_Mtrl_num_Mtrlgrp        'Mtrl_num_Mtrlgrp
    E_Output_Qty              'Output_Qty
    E_Consumed_Qty            'Consumed_Qty
    E_Reject_Qty              'Reject_Qty
    E_Current_Wafer_Qty       'Current_Wafer_Qty
    E_Film_Frame_Qty          'Film_Frame_Qty
    E_Optical_Quality         'Optical_Quality
    
    E_Country_of_Assembly              'Country_of_Assembly
    E_Offshore_ASM_Company             'Offshore_ASM_Company
    E_Asm_Containment_type             'Asm_Containment_type
    E_Date_code                        'Date_code
    E_Asm_conv_id                      'Asm_conv_id
    E_Asm_excr_id                       'Asm_excr_id
    
    E_Assembly_facility                'Assembly_facility
    E_Country_of_Test                  'Country_of_Test
    E_Offshore_TEST_Company            'Offshore_TEST_Company
    E_Tst_Containment_type             'Tst_Containment_type
    E_Tst_Program_rev                  'Tst_Program_rev
    E_Created_date                      'Created_date
        
    E_Created_time                      'Created_time
    E_Del_Note                          'Del_Note
    E_AWB                               'AWB
    E_Weight                            'Weight
    E_PackageQty                        'PackageQty

    
    
    
    E_End
    
End Enum

Dim oiRS        As New ADODB.Recordset
Dim grRS        As New ADODB.Recordset
Public g_Date           As String
Dim sourceBatchID  As String
Dim sourceBatchIDDel  As String



Private Sub CmdAdd_Click()
Dim lotidTemp As String
Dim grTemp As SpGR
Dim idtemp As Long
Dim baofeiDetail As String
Dim baofeiPieceQty As Integer
Dim baofeiDieQty As Long
Dim waferidTemp As String

Dim strA() As String
Dim i As Integer


lotidTemp = UCase(Trim(TxtLotID.Text))
grTemp.PoLotID = lotidTemp
baofeiDieQty = 0


'判断这个lotId有没有OI
'判断这个lotId有没有维护过，维护过，则不可以再添加

'2014-09-29 jiayun add ,WaferID中有多少个","号
baofeiDetail = Replace(UCase(Trim(TxtBaoFei.Text)), "，", ",")

If baofeiDetail = "" Then
    baofeiPieceQty = 0
Else


strA = Split(baofeiDetail, ",")
baofeiPieceQty = UBound(strA) - LBound(strA)
baofeiPieceQty = baofeiPieceQty + 1

'查询所有Die数的数量

For i = 0 To baofeiPieceQty - 1
waferidTemp = UCase(Trim(strA(i)))

baofeiDieQty = baofeiDieQty + GetAAMaping_GDieQty(waferidTemp)

Next i


End If





If (Not JudgeSpecialGRBillNoOI(lotidTemp)) Then

    MsgBox "没有查询到客户OI信息,请确认!", vbInformation, "友情提示"
    Exit Sub

Else


        If JudgeSpecialGRBillNo(lotidTemp) Then
        
            MsgBox "以前有维护过这个LotID,请确认!", vbInformation, "友情提示"
            Exit Sub
        
        
        Else
        
        
                    '查询OI中的信息
                    Set oiRS = GetAAGROIData(lotidTemp)
                    
                    If (oiRS.RecordCount > 0) Then
                        grTemp.PoNum = getStr(oiRS.fields("po_num").Value)
                        grTemp.PoItem = getStr(oiRS.fields("po_item").Value)
                        grTemp.PreviousMtrl = getStr(oiRS.fields("source_mtrl_num").Value)
                        grTemp.MtrlNum = getStr(oiRS.fields("mpn").Value)
                        grTemp.MtrlDesc = getStr(oiRS.fields("mpn_desc").Value)
                        grTemp.MtrlNumMtr = getStr(oiRS.fields("mtrl_num_mtrlgrp").Value)
                        grTemp.CreatedDate = getStr(oiRS.fields("created_date").Value)
                        grTemp.CreatedTime = getStr(oiRS.fields("created_time").Value)
                        idtemp = CLng(oiRS.fields("ID").Value)
                    End If
                    
                    grTemp.BatchID = "QF" & Right(("000000" & idtemp), 7)
                    
                    '查询不良品数  OI总Die数- 所有以前GR总Die数之和
                    
                    grTemp.RejectQty = GetSpeGRNGDieQty(lotidTemp) - baofeiDieQty
                    grTemp.ConsumedQty = grTemp.RejectQty
                    
                    '更新Die数
                    
                    
                    
                    grTemp.CurrentWaferQty = Format(GetSpeGRNGPieceQty(lotidTemp) - baofeiPieceQty, "0.00")
                    
              
                    
                    '查询以前GR中的信息
                    Set grRS = GetSpecialGRBefor(lotidTemp)
                    If (grRS.RecordCount > 0) Then
                    
                      grTemp.DateCode = getStr(grRS.fields("Date_code").Value)
                      grTemp.TstProgram = getStr(grRS.fields("Tst_Program_rev").Value)
                    
                    Else
                      '查询不到，则查询开工单的日期
                      grTemp.DateCode = GetSpecilGRDt(lotidTemp)
                      grTemp.TstProgram = GetSpecilGRTestVer(lotidTemp)
                     
                    End If
                    
                    
                    '把数据插入表中
                    Call AddSpecialGR(grTemp)
                    
                    
                    ShowData_Where
        
        
        End If

End If

End Sub

Private Function getStr(strTemp As Variant)
getStr = Trim("" & strTemp)
End Function


Private Sub CmdDel_Click()

With fps(0)
    .Row = .ActiveRow
    .Col = 3
     sourceBatchIDDel = .Text
End With

Call DelSpecialGR(sourceBatchIDDel)
ShowData_Where


End Sub

Private Sub CmdModify_Click()
Dim pieceTemp As Double
Dim dieQtyTemp As Integer

If TxtAfterPiece.Text = "" Or TxtAfterDie.Text = "" Then
    MsgBox "请先输入修改后的片数与Die数!", vbInformation, "友情提示"
    TxtAfterPiece.SetFocus
    Exit Sub
Else


        pieceTemp = Format(CDbl(UCase(Trim(TxtAfterPiece.Text))), "0.00")
        dieQtyTemp = CInt(UCase(Trim(TxtAfterDie.Text)))
        
        
        Call ModifySpecialGR(sourceBatchID, pieceTemp, dieQtyTemp)
        
        ShowData_Where
End If

End Sub

Private Sub Command1_Click()
'生成GR文件
'判断是否有资料

Dim i As Integer
Dim j As Integer
Dim flagTemp As Boolean
flagTemp = False

  With fps(0)
        For i = 1 To .MaxRows
               .Row = i
               .Col = 1
               
               If .Text <> "" Then
                 flagTemp = True
               End If

         Next
    End With
    

If flagTemp Then
 
SaveGRTxt

Else
    MsgBox "没有数据信息,请确认!", vbInformation, "友情提示"
    Exit Sub

End If

End Sub

Private Sub Form_Activate()
 g_Date = Format(Now, "YYYY-MM-DD hh:mm:ss")
'TxtBaoFei.Text = 0

End Sub

Private Sub SaveGRTxt()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim Rs          As New ADODB.Recordset

On Error GoTo ErrHandler
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & "_S" & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    ",Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Offshore_ASM_Company,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    ",Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
   
   With fps(0)

        For i = 1 To .MaxRows
             strColData = ""
             For j = 1 To .MaxCols
             
                If j = 27 Then
                     strColData = strColData + Format(g_Date, "hh:mm:ss") + ","
                ElseIf j = 26 Then
                     strColData = strColData + Replace(Format(g_Date, "MM/DD/YYYY"), "-", "/") + ","
                    
                Else
               .Row = i
               .Col = j
               strColData = strColData + Trim(.Text) + ","
               
               End If
               
              
               
               
             Next
             strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
             
         Next
         
    End With
    

    
    
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    MsgBox ("生成成功，请到目录中查看！")
    
    LogFile.Close
    Set LogFile = Nothing
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub ShowData_Where()
Set reportRS = GetSpecialGRDataList()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub


Private Sub Form_Load()
IniFpsHeader

ShowData_Where


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
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
       ' .SetText E_FPS0.E_SeqId, 0, "记录号"
        .SetText E_FPS0.E_PO_Num, 0, "PO_Num"
        .SetText E_FPS0.E_PO_Item, 0, "PO_Item"
        .SetText E_FPS0.E_Previous_Batch_ID, 0, "Previous_Batch_ID"
        .SetText E_FPS0.E_Previous_Mtrl_Num, 0, "Previous_Mtrl_Num"
        .SetText E_FPS0.E_Batch_ID, 0, "Batch_ID"
        .SetText E_FPS0.E_Mtrl_num, 0, "Mtrl_num"
        .SetText E_FPS0.E_Mtrl_desc, 0, "Mtrl_desc"
        
        .SetText E_FPS0.E_Mtrl_num_Mtrlgrp, 0, "Mtrl_num_Mtrlgrp"
        .SetText E_FPS0.E_Output_Qty, 0, "Output_Qty"
        .SetText E_FPS0.E_Consumed_Qty, 0, "Consumed_Qty"
        .SetText E_FPS0.E_Reject_Qty, 0, "Reject_Qty"
        .SetText E_FPS0.E_Current_Wafer_Qty, 0, "Current_Wafer_Qty"
        .SetText E_FPS0.E_Film_Frame_Qty, 0, "Film_Frame_Qty"
        .SetText E_FPS0.E_Optical_Quality, 0, "Optical_Quality"
            
        .SetText E_FPS0.E_Country_of_Assembly, 0, "Country_of_Assembly"
        .SetText E_FPS0.E_Offshore_ASM_Company, 0, "Offshore_ASM_Company"
        .SetText E_FPS0.E_Asm_Containment_type, 0, "Asm_Containment_type"
        .SetText E_FPS0.E_Date_code, 0, "Date_code"
        .SetText E_FPS0.E_Asm_conv_id, 0, "Asm_conv_id"
        .SetText E_FPS0.E_Asm_excr_id, 0, "Asm_excr_id"
      
        .SetText E_FPS0.E_Assembly_facility, 0, "Assembly_facility"
        .SetText E_FPS0.E_Country_of_Test, 0, "Country_of_Test"
        .SetText E_FPS0.E_Offshore_TEST_Company, 0, "Offshore_TEST_Company"
        
        .SetText E_FPS0.E_Tst_Containment_type, 0, "Tst_Containment_type"
        .SetText E_FPS0.E_Tst_Program_rev, 0, "Tst_Program_rev"
        .SetText E_FPS0.E_Created_date, 0, "Created_date"
        
        .SetText E_FPS0.E_Created_time, 0, "Created_time"
        .SetText E_FPS0.E_Del_Note, 0, "Del_Note"
        .SetText E_FPS0.E_AWB, 0, "AWB"
        
        .SetText E_FPS0.E_Weight, 0, "weight(kgs)"
        .SetText E_FPS0.E_PackageQty, 0, "package"
     
  
       ' .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_PO_Num) = 10
        .ColWidth(E_FPS0.E_PO_Item) = 10
        .ColWidth(E_FPS0.E_Previous_Batch_ID) = 10
        .ColWidth(E_FPS0.E_Previous_Mtrl_Num) = 10
        .ColWidth(E_FPS0.E_Batch_ID) = 10
        .ColWidth(E_FPS0.E_Mtrl_num) = 10
        .ColWidth(E_FPS0.E_Mtrl_desc) = 10
        
        .ColWidth(E_FPS0.E_Mtrl_num_Mtrlgrp) = 10
        .ColWidth(E_FPS0.E_Output_Qty) = 10
        .ColWidth(E_FPS0.E_Consumed_Qty) = 10
        .ColWidth(E_FPS0.E_Reject_Qty) = 10
        .ColWidth(E_FPS0.E_Current_Wafer_Qty) = 10
        .ColWidth(E_FPS0.E_Film_Frame_Qty) = 10
        .ColWidth(E_FPS0.E_Optical_Quality) = 10
        
        
        .ColWidth(E_FPS0.E_Country_of_Assembly) = 10
        .ColWidth(E_FPS0.E_Offshore_ASM_Company) = 10
        .ColWidth(E_FPS0.E_Asm_Containment_type) = 10
        .ColWidth(E_FPS0.E_Date_code) = 10
        .ColWidth(E_FPS0.E_Asm_conv_id) = 10
        .ColWidth(E_FPS0.E_Asm_excr_id) = 10
        
        
        .ColWidth(E_FPS0.E_Assembly_facility) = 10
        .ColWidth(E_FPS0.E_Country_of_Test) = 10
        .ColWidth(E_FPS0.E_Offshore_TEST_Company) = 10
        .ColWidth(E_FPS0.E_Tst_Containment_type) = 10
        .ColWidth(E_FPS0.E_Tst_Program_rev) = 10
        .ColWidth(E_FPS0.E_Created_date) = 10
        
        .ColWidth(E_FPS0.E_Created_time) = 10
        .ColWidth(E_FPS0.E_Del_Note) = 10
        .ColWidth(E_FPS0.E_AWB) = 10
        .ColWidth(E_FPS0.E_Weight) = 10
        .ColWidth(E_FPS0.E_PackageQty) = 10
        

        .RowHeight(0) = 30
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i As Long

With fps(0)

            .Row = Row
            .Col = 3
            sourceBatchID = .Text

            .Row = Row
            .Col = 11
            TxtBeforDie.Text = .Text
            
             .Row = Row
            .Col = 12
            TxtBeforPiece.Text = .Text
            
                 

End With

TxtAfterPiece.SetFocus



End Sub
