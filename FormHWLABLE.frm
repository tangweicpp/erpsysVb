VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FormHWLABLE 
   Caption         =   "��Ϊ����ǩ"
   ClientHeight    =   11745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   23085
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
   ScaleHeight     =   11745
   ScaleWidth      =   23085
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   22815
      Begin VB.Frame Frame2 
         Caption         =   "��ǩѡ��"
         Height          =   1815
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   3975
         Begin VB.CheckBox chkMakeUsed 
            Caption         =   "����"
            Height          =   315
            Left            =   2040
            TabIndex        =   10
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton OptCarton 
            Caption         =   "�ɱ�ǩ"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptLablePrint 
            Caption         =   "�±�ǩ"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   480
         Width           =   4335
      End
      Begin VB.CheckBox ChkAll 
         Height          =   255
         Left            =   20760
         TabIndex        =   3
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "��ӡ"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   8280
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdquery 
         Caption         =   "��ѯ"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   5160
         TabIndex        =   1
         Top             =   1320
         Width           =   2295
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6615
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   20535
         _Version        =   524288
         _ExtentX        =   36221
         _ExtentY        =   11668
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
         SpreadDesigner  =   "FormHWLABLE.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "txt����·���趨:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label lbl123 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȫѡ:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   19800
         TabIndex        =   4
         Top             =   1800
         Width           =   900
      End
   End
End
Attribute VB_Name = "FormHWLABLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 00_�����ͷ����
Dim sel_box As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


    
' 01_��ʼ����ͷ
Public Sub InitFpsHeader()

If OptLablePrint.Value Then
    ' With fps(0)
     ' �Զ����ͷ
     '   .ReDraw = False
     '   .MaxCols = E_FPS0.E_End - 1
     '   .MaxRows = 0

     '   '.DAutoHeadings = False
     '   .DAutoCellTypes = False
     '   .DAutoSizeCols = DAutoSizeColsNone
        
     '   .Col = -1
     '   .Row = -1
     '   .Lock = True
     '   .OperationMode = OperationModeNormal
     '   .TypeVAlign = TypeVAlignCenter
     '   .SelForeColor = &HFF8080
        
        
     '   .Col = 1
     '   .CellType = CellTypeCheckBox
     '   .TypeHAlign = TypeHAlignCenter
     '   .TypeVAlign = TypeVAlignCenter
        
     '    For i = 1 To .MaxCols
     '       .Col = i
     '       .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
     '   Next
        
     '   .ZOrder
     '   .ReDraw = True

     '   .ColWidth(1) = 2
     '   .RowHeight(0) = 20
     '   .RowHeight(-1) = 15

     '   .Col = 1
     '   .Lock = False
    
     'End With
    
    ' ��1
     With fps(0)
        .ReDraw = False
        .MaxCols = sel_box
        .MaxRows = 0
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        ' �����ѡ��
        .Col = sel_box
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        ' �����ͷ��
        .SetText 1, 0, "���ϱ���"
        .SetText 2, 0, "����汾"
        .SetText 3, 0, "��������"
        .SetText 4, 0, "��λ(PCS/SET/M��)"
        .SetText 5, 0, "����׼�����"
        .SetText 6, 0, "��Ϊ�����ͺ�"
        .SetText 7, 0, "��Ϊ������������"
        .SetText 8, 0, "��Ϊ����Ӣ������"
        .SetText 9, 0, "���к�"
        .SetText 10, 0, "[���ǩ]"
        .SetText 11, 0, "���̴���"
        .SetText 12, 0, "��������"
        .SetText 13, 0, "ë��"
        .SetText 14, 0, "���̹��"
        .SetText 15, 0, "��Ӧ����������"
        .SetText 16, 0, "�����ʶ"
        .SetText 17, 0, "����"
        .SetText 18, 0, "PO"
        .SetText 19, 0, "09��"
        .SetText 20, 0, "��ע������Ϣ"
        .SetText 21, 0, "ԭ������Ϣ"
        .SetText 22, 0, "�Ѵ�ӡ"
        .SetText 23, 0, "ѡ��"
          
        ' ������
        .ColWidth(1) = 8
        .ColWidth(2) = 8
        .ColWidth(3) = 8
        .ColWidth(4) = 8
        .ColWidth(5) = 8
        .ColWidth(6) = 8
        .ColWidth(7) = 8
        .ColWidth(8) = 8
        .ColWidth(9) = 8
        .ColWidth(10) = 30
        .ColWidth(11) = 5
        .ColWidth(12) = 8
        .ColWidth(13) = 8
        .ColWidth(14) = 8
        .ColWidth(15) = 10
        .ColWidth(16) = 8
        .ColWidth(17) = 10
        .ColWidth(18) = 10
        .ColWidth(19) = 10
        .ColWidth(20) = 5
        .ColWidth(21) = 5
        .ColWidth(22) = 5
        
        ' ����߶�
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        ' �����Ƿ�ɱ༭
        .Col = 23
            .Lock = False
    
        .ReDraw = True
    End With
    
Else
    ' ��2
     With fps(0)
        .ReDraw = False
        .MaxCols = sel_box
        .MaxRows = 0
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        ' �����ѡ��
        .Col = sel_box
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        ' �����ͷ��
        .SetText 1, 0, "bar_code"
        .SetText 2, 0, "po_number"
        .SetText 3, 0, "vendor_code"
        .SetText 4, 0, "item_code"""
        .SetText 5, 0, "item_rev"
        .SetText 6, 0, "rohs"
        .SetText 7, 0, "pcs"
        .SetText 8, 0, "hw_m"
        .SetText 9, 0, "item_desc"
        .SetText 10, 0, "code_09"
        .SetText 11, 0, "mpn"
        .SetText 12, 0, "vendor_lot"
        .SetText 13, 0, "country"
        .SetText 14, 0, "produc_date"
        .SetText 15, 0, "remarks"
        .SetText 16, 0, "�Ѵ�ӡ"
        .SetText 17, 0, "ѡ��"
          
        ' ������
        .ColWidth(1) = 8
        .ColWidth(2) = 8
        .ColWidth(3) = 8
        .ColWidth(4) = 8
        .ColWidth(5) = 8
        .ColWidth(6) = 8
        .ColWidth(7) = 8
        .ColWidth(8) = 8
        .ColWidth(9) = 8
        .ColWidth(10) = 30
        .ColWidth(11) = 5
        .ColWidth(12) = 8
        .ColWidth(13) = 8
        .ColWidth(14) = 8
        .ColWidth(15) = 10
        .ColWidth(16) = 4
        ' ����߶�
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        ' �����Ƿ�ɱ༭
        .Col = 17
            .Lock = False
    
        .ReDraw = True
    End With

End If
    
End Sub


Private Sub ChkAll_Click()
Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = sel_box
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = sel_box
                .Text = 0
            End With
        Next i
        
    End If
End Sub

Private Sub cmdprint_Click()
PrintSelTxt

End Sub

Private Sub Form_Load()

OptLablePrint.Value = True

If OptLablePrint.Value Then
    ' �±�ǩ
    
    txtText1.Text = "\\10.160.1.14\BarCode\37\37HWOUTnew\"
    sel_box = 23
Else
    ' �ɱ�ǩ
    txtText1.Text = "\\10.160.1.14\BarCode\37\37HWOUTold\"
    sel_box = 17
End If

InitFpsHeader

End Sub



Private Sub cmdQuery_Click()

If OptLablePrint.Value Then
    ' ��1
    sel_box = 23
Else
    ' ��2
    sel_box = 17
End If

InitFpsHeader
ShowData

End Sub

Private Sub ShowData()
Dim mainItemRS As New adodb.Recordset
Dim cmd_ora  As String

If OptLablePrint.Value Then
    ' LablePrint��ǩ
    If chkMakeUsed Then
        ' ����
        cmd_ora = "select a.part_no,a.ver,a.qty,a.unit,a.sn_tn,a.hw_m,a.item_desc_cn, replace(a.item_desc_en, ',', '-'),a.sn" & _
                    ",a.pkg_id || ';' || a.ver ||  ';' || a.part_no ||  ';' || a.po ||  ';' || a.mfg_code ||  ';' || a.man_date || ';' ||  a.g_w ||  ';' || a.m_lot || ';' ||  a.qty || ';' || a.sn " & _
                    ",a.mfg_code,a.man_date,a.g_w,a.mpn,a.m_lot,a.law,a.pkg_id,a.po,a.code_09,a.remark,a.country,a.print, ''  from HUAWEI_LABLE a"
    Else
        cmd_ora = "select a.part_no,a.ver,a.qty,a.unit,a.sn_tn,a.hw_m,a.item_desc_cn, replace(a.item_desc_en, ',', '-'),a.sn" & _
                    ",a.pkg_id || ';' || a.ver ||  ';' || a.part_no ||  ';' || a.po ||  ';' || a.mfg_code ||  ';' || a.man_date || ';' ||  a.g_w ||  ';' || a.m_lot || ';' ||  a.qty || ';' || a.sn " & _
                    ",a.mfg_code,a.man_date,a.g_w,a.mpn,a.m_lot,a.law,a.pkg_id,a.po,a.code_09,a.remark,a.country,a.print, ''  from HUAWEI_LABLE a where a.print is null"
    End If
Else
    ' Carton��ǩ
    If chkMakeUsed Then
        ' ����
        cmd_ora = "select b.bar_code,b.po_number,b.vendor_code,b.item_code,b.item_rev,b.rohs,b.pcs,b.hw_m,b.item_desc,b.code_09," & _
                    "b.mpn,b.vendor_lot,b.country,b.produc_date,b.remarks,b.print, ''  from HUAWEI_CARTON  b"
    Else
        cmd_ora = "select b.bar_code,b.po_number,b.vendor_code,b.item_code,b.item_rev,b.rohs,b.pcs,b.hw_m,b.item_desc,b.code_09," & _
                    "b.mpn,b.vendor_lot,b.country,b.produc_date,b.remarks,b.print, ''  from HUAWEI_CARTON  b where b.print is null"
    
    End If

End If

Set mainItemRS = getStr(cmd_ora)

With fps(0)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
        End If
End With

End Sub
' ��ӡ��ǩ�ı�
Private Sub PrintSelTxt()
Dim out_path As String
Dim file_name As String
Dim j As Integer
Dim i As Integer
Dim strTmp As String
Dim strFlag As String
Dim strOra As String
Dim sqlFlag As Boolean

' ����·��
out_path = txtText1.Text
sqlFlag = False
strTmp = ""
strFlag = ""

If OptLablePrint.Value Then

    With fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = 23
            If .Text = 1 Then
                For j = 1 To 22
                    .Row = i
                    .Col = j
                    strTmp = strTmp & .Text & ","
            
                Next j
            
                ' ��ȡһ�����ݴ�ӡһ��txt�ı�
                strTmp = strTmp & Format(DateTime.Now, "yyyy-MM-dd")
                ' �ļ���
                file_name = Format(DateTime.Now, "yyyyMMddHHmmss") & "_HUAWEI_LABLE"
                
                Call addLabelTxt(file_name, strTmp, out_path)
               
                ' ��ӡ��־��ֵ, Remarks
                .Row = i
                .Col = 20
                strFlag = .Text
               
                strOra = "update HUAWEI_LABLE hl set hl.print = 'Y' where hl.remark = '" & strFlag & "'"
               
                Call AddSql(strOra)
                
                Sleep (1000)
            End If

        Next i
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 23
            If .Text = 1 Then
                sqlFlag = True
            End If
        Next
        
        If sqlFlag = False Then
            MsgBox "�빴ѡ���ӡ����Ŀ"
        Else
            MsgBox "����HUAWEI_LABLE��ǩ,�����ı��ɹ�, ������_ " & out_path & ""
        End If
        
    End With
Else
     With fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = 17
            If .Text = 1 Then
                For j = 1 To 16
                    .Row = i
                    .Col = j
                    strTmp = strTmp & .Text & ","
            
                Next j
            
                ' ��ȡһ�����ݴ�ӡһ��txt�ı�
                strTmp = strTmp & Format(DateTime.Now, "yyyy-MM-dd")
                ' �ļ���
                file_name = Format(DateTime.Now, "yyyyMMddHHmmss") & "_HUAWEI_CARTON"
                
                Call addLabelTxt(file_name, strTmp, out_path)
                
                 ' ��ӡ��־��ֵ, Remarks
                .Row = i
                .Col = 15
                strFlag = .Text
               
                strOra = "update HUAWEI_CARTON hc set hc.print = 'Y' where hc.remarks = '" & strFlag & "'"
               
                Call AddSql(strOra)
                
                Sleep (1000)
            End If

        Next i
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 17
            If .Text = 1 Then
                sqlFlag = True
            End If
        Next
        
        If sqlFlag = False Then
            MsgBox "�빴ѡ���ӡ����Ŀ"
        Else
            MsgBox "����HUAWEI_CARTON��ǩ,�����ı��ɹ�, ������_ " & out_path & ""
        End If

    End With
    
End If

End Sub

Private Sub OptCarton_Click()
If OptCarton.Value = True Then
    txtText1.Text = "\\10.160.1.14\BarCode\37\37HWOUTold\"
End If

End Sub

Private Sub OptLablePrint_Click()
If OptLablePrint.Value = True Then
     txtText1.Text = "\\10.160.1.14\BarCode\37\37HWOUTnew\"
End If
End Sub
