VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmDummy 
   Caption         =   "TSV �¹��� (Dummy wafer)"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   18765
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdM 
      BackColor       =   &H0080C0FF&
      Caption         =   "�ֹ���������"
      Height          =   480
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBom 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bom���趨"
      Height          =   480
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "�������"
      Height          =   480
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����Detail"
      Height          =   480
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����Header"
      Height          =   480
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton ComSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "���湤��"
      Height          =   480
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "����Detail"
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   18615
      Begin VB.CheckBox ChkAll 
         Height          =   255
         Left            =   13320
         TabIndex        =   59
         Top             =   120
         Width           =   255
      End
      Begin VB.ListBox Lst 
         Height          =   5010
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   54
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         Height          =   360
         Left            =   2040
         TabIndex        =   53
         Top             =   2040
         Width           =   615
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   0
         Left            =   2760
         TabIndex        =   51
         Top             =   360
         Width           =   15855
         _Version        =   196613
         _ExtentX        =   27966
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
         SpreadDesigner  =   "FrmDummy.frx":0000
         TextTip         =   2
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��LotId"
         Height          =   195
         Left            =   600
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����Header"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18615
      Begin VB.TextBox applyUserTxt 
         Height          =   405
         Left            =   8760
         TabIndex        =   70
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox TxtWoDept 
         Height          =   285
         Left            =   3960
         TabIndex        =   68
         Top             =   3000
         Width           =   3375
      End
      Begin VB.ComboBox CmbCheckCustomer 
         Height          =   315
         Left            =   1080
         TabIndex        =   57
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox TxtShipSite 
         Height          =   285
         Left            =   15960
         TabIndex        =   55
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtRequestDate 
         Height          =   285
         Left            =   12960
         TabIndex        =   49
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtMpn 
         Height          =   285
         Left            =   6840
         TabIndex        =   47
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtLotStatus 
         Height          =   285
         Left            =   3960
         TabIndex        =   45
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtFilmApld 
         Height          =   285
         Left            =   10200
         TabIndex        =   43
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtPoItem 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtMMaterial 
         Height          =   285
         Left            =   15960
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox TxtCounFab 
         Height          =   285
         Left            =   12960
         TabIndex        =   37
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   10200
         TabIndex        =   35
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox TxtMarkingcode 
         Height          =   285
         Left            =   6840
         TabIndex        =   33
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3960
         TabIndex        =   31
         Text            =   "Y"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Txt260 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   15960
         TabIndex        =   27
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtDesignId 
         Height          =   285
         Left            =   12960
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCusRev 
         Height          =   285
         Left            =   10200
         TabIndex        =   23
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtFab 
         Height          =   285
         Left            =   6840
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCustomerPT 
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtPo 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   15960
         TabIndex        =   16
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   178978817
         CurrentDate     =   40882
      End
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   12960
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtNum 
         Height          =   285
         Left            =   10200
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmDummy.frx":4474
         Left            =   3960
         List            =   "FrmDummy.frx":4476
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ѯOI"
         Enabled         =   0   'False
         Height          =   360
         Left            =   6360
         TabIndex        =   4
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox TxtSourceBatchId 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo Text3 
         Height          =   315
         Left            =   6840
         TabIndex        =   61
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   960
         TabIndex        =   67
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������ˣ�"
         Height          =   195
         Left            =   7680
         TabIndex        =   71
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label LblWoDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ţ�"
         Height          =   195
         Left            =   3120
         TabIndex        =   69
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "�ӿ��еĿͻ�"
         Height          =   435
         Left            =   480
         TabIndex        =   58
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ShipSite"
         Height          =   195
         Left            =   15360
         TabIndex        =   56
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�������"
         Height          =   195
         Left            =   12000
         TabIndex        =   50
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn"
         Height          =   195
         Left            =   6360
         TabIndex        =   48
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotStatus"
         Height          =   195
         Left            =   3120
         TabIndex        =   46
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProtectiveFilmApld"
         Height          =   195
         Left            =   8760
         TabIndex        =   44
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PoItem"
         Height          =   195
         Left            =   480
         TabIndex        =   42
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MicronMaterial"
         Height          =   195
         Left            =   14880
         TabIndex        =   40
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CountryFab"
         Height          =   195
         Left            =   12000
         TabIndex        =   38
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(*)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9600
         TabIndex        =   36
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MarkingCode"
         Height          =   195
         Left            =   5880
         TabIndex        =   34
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NG��־"
         Height          =   195
         Left            =   3240
         TabIndex        =   32
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level260"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level235"
         Height          =   195
         Left            =   15240
         TabIndex        =   28
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DesignId"
         Height          =   195
         Left            =   12240
         TabIndex        =   26
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ImagerCustomerRev"
         Height          =   195
         Left            =   8640
         TabIndex        =   24
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB�豸"
         Height          =   195
         Left            =   6120
         TabIndex        =   22
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ��Ϻ�"
         Height          =   195
         Left            =   3120
         TabIndex        =   20
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���깤��"
         Height          =   195
         Left            =   15000
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�ƿ�����"
         Height          =   195
         Left            =   12000
         TabIndex        =   14
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   195
         Left            =   9360
         TabIndex        =   12
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ�Ϻ�"
         Height          =   195
         Left            =   6000
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source_batch_id"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail�֭�
    E_ID = 1                 'id��
    E_WaferID                'Waferid
    E_CompleteFlag           '��ɱ�־�W
    E_TotalDie               '������
    E_GoodDie                'good������
    E_WaferLot               'wafer
    E_MarkingCode            'markingcode
    E_OK                     'ѡ��֭�
    E_End
    
End Enum

Private Enum E_FPS1          'Bom�֭�
    E_ID = 0                 'id��
    E_BomID                  '���Ϲ淶���
    E_PT                     '�Ϻ�
    E_Mt                     '���ϱ��
    E_Name                   '���ƪ�
    E_Qty                    'ÿֻ����
    E_Unit                   '��λ
    
    E_Pt2                     '�Ϻ�2
    E_Mt2                     '���ϱ��2
    E_Name2                   '����2��
    E_Qty2                    'ÿֻ����2
    E_Unit2                   '��λ2
    
    E_End
    
End Enum


Dim oiRS        As New ADODB.Recordset
Dim listRS        As New ADODB.Recordset
Dim bomRS        As New ADODB.Recordset

Dim mainItemRS As New ADODB.Recordset


Private Sub ChkAll_Click()

Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next i
        
    End If


End Sub

Private Sub CmdBom_Click()
FrmTSV_Bom.Show

End Sub

Private Sub CmdM_Click()
FormM.Show

Unload Me

End Sub

Private Sub Command1_Click()
'7468859023
'TxtSourceBatchId.Text = "7468859023"
If CmbCustomer.Text = "" Or TxtSourceBatchId.Text = "" Then
    MsgBox "����ѡ��ͻ���������Lot�š���ȷ��!", vbInformation, "������ʾ"
    Exit Sub
Else

    Set oiRS = GetOIData(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtSourceBatchId.Text)))
    If (oiRS.RecordCount > 0) Then

        TxtPO.Text = getStr(oiRS.fields("po_num").Value)
        TxtCustomerPT.Text = getStr(oiRS.fields("mpn_desc").Value)
        TxtFab.Text = getStr(oiRS.fields("fabrication_facility").Value)
        TxtCusRev.Text = getStr(oiRS.fields("imager_customer_rev").Value)
        TxtDesignId.Text = getStr(oiRS.fields("design_id").Value)
        Txt260.Text = getStr(oiRS.fields("shipping_mst_260").Value)
        Text11.Text = getStr(oiRS.fields("shipping_mst_level").Value)
        TxtMarkingcode.Text = getStr(oiRS.fields("encoded_mark_id").Value)
        TxtCounFab.Text = getStr(oiRS.fields("country_of_fab").Value)
        TxtMMaterial.Text = getStr(oiRS.fields("micron_material").Value)
        TxtPOItem.Text = getStr(oiRS.fields("po_item").Value)
        TxtLotStatus.Text = getStr(oiRS.fields("lot_status").Value)
        TxtMpn.Text = getStr(oiRS.fields("mpn").Value)
        
        If getStr(oiRS.fields("protective_film_apld").Value) = "YES" Then
            TxtFilmApld.Text = "PF"
        Else
            TxtFilmApld.Text = getStr(oiRS.fields("protective_film_apld").Value)
        End If
        
        TxtRequestDate.Text = getStr(oiRS.fields("lot_priority").Value)
        TxtShipSite.Text = getStr(oiRS.fields("ship_site").Value)
        
        If TxtShipSite.Text = "Qtech" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "WLC"
            
        ElseIf TxtShipSite.Text = "SG" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "AA"
            
        ElseIf UCase(Trim(CmbCustomer.Text)) = "GC" Then
             CmbCheckCustomer.Text = "GC"
        End If
        
        Call IniProductTwo(UCase(Trim(CmbCustomer.Text)))
        
        '��ʼ����ߵ�Lot��ϸ��
        
        Call InitListBox(UCase(Trim(CmbCustomer.Text)))
        
        
        
        '2012-11-05 jiayun add  ��ѯ��Ʒ�Ϻ�
        
        
        
        
        
    Else
        MsgBox "��ѯ�������ݣ���ȷ�� SourceBatchId "
        Exit Sub
 
    
    End If
    

    
    
End If
End Sub
Private Function getStr(strTemp As Variant)
getStr = Trim("" & strTemp)
End Function

Private Sub Command2_Click()
Dim strTmp As String
Dim strTemp As String
strTemp = ""
With Lst
        '��ʼ���Ҹ�ֵ
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = .List(i) & "','"
                strTemp = strTemp & strTmp

            End If
        Next
 End With
 
 If strTemp = "" Then
 
 MsgBox "����ѡ��LotId !"
 Exit Sub
 
 Else
 
 strTemp = Mid(strTemp, 1, Len(strTemp) - 3)
 
Call GetFpsData(strTemp, UCase(Trim(CmbCustomer.Text)))

ChkAll.Value = 1
ChkAll_Click

End If


End Sub

Private Sub GetFpsData(strwhereTemp As String, customerTemp As String)
'��ϸ����

Set listRS = GetFps(strwhereTemp, customerTemp)
If listRS.RecordCount <= 0 Then
    MsgBox "��ϸ����û��������ݣ���ȷ��"
    Exit Sub
End If

With fps(0)
        .MaxRows = 0
        If listRS.RecordCount > 0 Then
            Set .DataSource = listRS
        End If
End With

End Sub

Private Sub GetBomData(ptTemp As String)
'��ϸ����

Set bomRS = GetFpsBom(ptTemp)
If bomRS.RecordCount <= 0 Then
    MsgBox "��ϸ����û��������ݣ���ȷ��"
    Exit Sub
End If

With fps(1)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub



Private Sub InitListBox(customerTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailData(customerTemp)
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
      
        

listRS.Close
Set listRS = Nothing

End Sub

Private Sub Command3_Click()
  
 Dim sqlTemp As String
 sqlTemp = "select SEQ_IBWO,ORDERNAME,ORDERTYPE,DESCRIPTION,EVENTTYPE,ERPUSER,PRODUCT,PRODUCTREVISION,QTY,PRODUCTBOM,ERPCREATEDATE,PLANSTARTDATE,PLANENDDATE," & _
         " Customer , SalesOrder, PRODUCTFAMILY, ModifyFlag, CUSTOMERPN, FabFacility, ImagerRev, Designid, MLevel235, Mlevel260, NGFlag, Para1, Para2, Para3, Para4, Para5, Para6, PARA7, PARA8, PARA9, PARA10, Protective_Film_Apld, LOT_STATUS, MPN " & _
         " From IB_WOHISTORY where ORDERNAME='" + Text2.Text + "'order by SEQ_IBWO desc "
  ExporToExcel (sqlTemp)
End Sub

Private Sub Command4_Click()

 Dim sqlTemp As String
 sqlTemp = "select ORDERNAME,WAFERID,COMPLETEFLAG,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE from IB_WAFERLIST where ordername ='" + Text2.Text + "' order by ORDERNAME, WAFERID"
  ExporToExcel (sqlTemp)

End Sub

Private Sub Command5_Click()

ClearData

End Sub

Private Sub ClearData()
'�����һ�ʵ�����
TxtSourceBatchId.Text = ""
Text2.Text = ""
Text3.Text = ""
TxtNum.Text = ""
TxtPO.Text = ""
TxtCustomerPT.Text = ""
TxtFab.Text = ""
TxtCusRev.Text = ""
TxtDesignId.Text = ""
Text11.Text = ""
Txt260.Text = ""
Text13.Text = ""
TxtMarkingcode.Text = ""
Text15.Text = ""
TxtCounFab.Text = ""
TxtMMaterial.Text = ""
TxtPOItem.Text = ""
TxtLotStatus.Text = ""
TxtMpn.Text = ""
TxtFilmApld.Text = ""
TxtRequestDate.Text = ""
TxtShipSite.Text = ""
CmbCheckCustomer.Text = ""
Lst.Clear

fps(0).MaxRows = 0


End Sub

Private Sub ComSave_Click()
'���湤��
Dim headerTemp As BillHeader
Dim detailTemp As BillDetail
Dim typeId As Integer
Dim SumQty As Long
Dim i As Integer
ComSave.Enabled = False

SumQty = 0

'Check���������Ƿ���д
If Trim(Text15.Text) = "" Then
     MsgBox "���ʲ�����Ϊ�գ�"
     ComSave.Enabled = True
     Exit Sub
End If

'��ֵ
 headerTemp.id = GetSeqID()
 headerTemp.OrderName = UCase(Trim(Text2.Text))
 
 If UCase(Trim(TxtNum.Text)) = "" Then
      MsgBox "������������"
      ComSave.Enabled = True
     Exit Sub

 End If
 
 If CInt(TxtNum.Text) < 1 Then
    MsgBox "�������ԣ�"
    ComSave.Enabled = True
     Exit Sub
 End If
 
   If Len(UCase(Trim(Text2.Text))) <> 12 Then
      MsgBox "�����ų��Ȳ��ԣ�"
      ComSave.Enabled = True
     Exit Sub
 
 End If
 
 
 
   
  If UCase(Trim(TxtWoDept.Text)) = "" Then
      MsgBox "�������Ų�����Ϊ�գ�"
      ComSave.Enabled = True
     Exit Sub
 
 End If
  
  
' '2012-11-30 jiayun add �ж��Ϻŵ�bom�Ƿ����
'Set bomRS2 = GetProductBom(Text3.Text)
'If bomRS2.RecordCount <= 0 Then
'    MsgBox "��ϵͳ�����Ϻŵ�Bom�����ڣ�����ϵ��ص��ˣ���ά��Bom ��"
'    ComSave.Enabled = True
'    Exit Sub
'End If


'GetWLOWo


 '2012-12-19 jiayun add У���Ϻ��Ƿ����
Set bomRS2 = GetProduct_Check(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "�ϺŲ����ڣ�����ϵ��ص��ˣ���ά���Ϻ� ��"
    ComSave.Enabled = True
    Exit Sub
End If

 
'  '2014-01-14 jiayun add �ж���ERP bom ��û��ǩ�˹�
'
'Set bomRS2 = GetProductBomERpSign(Text3.Text)
'If bomRS2.RecordCount <= 0 Then
'    MsgBox "��ϵͳ�����Ϻŵ�Bomû�б����ͨ��������ϵ���̲���"
'    ComSave.Enabled = True
'    Exit Sub
'End If
 
 
 
Select Case Combo2.Text
Case "һ�㹤��"
    typeId = 1
Case "�ټӹ�����"
    typeId = 5
Case "ί�⹤��"
    typeId = 7
    
Case "�ع�ί�⹤��"
    typeId = 8
    
Case "���ʽ����"
    typeId = 11
    
Case "Ԥ�⹤��"
    typeId = 13
Case "�Բ�����"
    typeId = 15
    
Case Else
   typeId = 0
End Select

 headerTemp.OrderType = CStr(typeId)
 headerTemp.EventType = "CREATED"
 headerTemp.ERPUser = "Auto"
 headerTemp.product = Text3.Text
                            
 headerTemp.RequestDate = Now
 headerTemp.ERPCreateDate = DateTime.Date
 headerTemp.PlanStartDate = CDate(TxtDate.Text)
 headerTemp.PlanEndDate = DTPicker1.Value
 headerTemp.Customer = CmbCustomer.Text
 headerTemp.SalesOrder = TxtPO.Text
 headerTemp.ModifyFlag = 0
 headerTemp.CustomerERPN = TxtCustomerPT.Text
 headerTemp.FabFacility = TxtFab.Text
headerTemp.ImagerRev = TxtCusRev.Text
headerTemp.DesignId = TxtDesignId.Text
headerTemp.MLevel235 = Text11.Text
headerTemp.Mlevel260 = Txt260.Text

headerTemp.NGFlag = Val(Text13.Text)

headerTemp.Para1 = TxtMarkingcode.Text
headerTemp.Para2 = Text15.Text
'headerTemp.Para3 = TxtCounFab.Text
headerTemp.Para4 = TxtMMaterial.Text
headerTemp.Para5 = TxtPOItem.Text
headerTemp.Para6 = TxtShipSite.Text
headerTemp.Para8 = TxtWoDept.Text

headerTemp.Protective_Film_Apld = TxtFilmApld.Text
headerTemp.Lot_Stauts = TxtLotStatus.Text
headerTemp.MPN = TxtMpn.Text

headerTemp.Para3 = Trim(applyUserTxt.Text)


 
'With fps(0)
'
'For i = 1 To .MaxRows
'    .Row = i
'    .Col = 8
'    If .Text = 1 Then
'        .Row = i
'        .Col = 4
'        SumQty = SumQty + CInt(.Text)
'    End If
'
'Next i
'
'End With

headerTemp.qty = CInt(TxtNum.Text)


'2016-01-12  У����������ţ��Ƿ��Ѵ���

If JudgeDummyWo1Stauts(headerTemp.OrderName) Then
    MsgBox "��Ҫ�ظ�����һ�������� ��"
    ComSave.Enabled = True
    Exit Sub

End If

If JudgeDummyWo2Stauts(headerTemp.OrderName) Then
    MsgBox "��Ҫ�ظ�����һ�������� ��"
    ComSave.Enabled = True
    Exit Sub

End If


 '--��ֵEnd
  Call AddBillHeaderWoDummy(headerTemp)
  
'--����Heand End

'--- Begin Detail

'�ж���ʹ�������Ӧ�ͻ���OI,�Ƿ�������



'MsgBox "������" & Text2.Text & "�����ɹ� !"


ComSave.Enabled = True

End Sub

Private Sub Form_Activate()
Text15.Text = "25"
End Sub

Private Sub Form_Load()

IniCustomerName

CmbCheckCustomer.AddItem ("AA")
CmbCheckCustomer.AddItem ("WLC")
CmbCheckCustomer.AddItem ("GC")
CmbCheckCustomer.AddItem ("SX")
CmbCheckCustomer.AddItem ("SY")
CmbCheckCustomer.AddItem ("ENG")


IniProduct

TxtDate.Text = Format(Now, "yyyy-mm-dd")
DTPicker1.Value = TxtDate.Text

Combo2.AddItem ("һ�㹤��")
Combo2.AddItem ("�ټӹ�����")
Combo2.AddItem ("ί�⹤��")
Combo2.AddItem ("�ع�ί�⹤��")
Combo2.AddItem ("���ʽ����")
Combo2.AddItem ("Ԥ�⹤��")
Combo2.AddItem ("�Բ�����")
Combo2.AddItem ("С�����Բ�����")






IniFpsHeader
'IniFpsBom

Frame2.Visible = False
CmdBom.Visible = False



End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub IniProduct()
Set mainItemRS = GetDummyProduct()
Set Text3.RowSource = mainItemRS
Text3.ListField = mainItemRS("productname").Name
Text3.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub IniProductTwo(customerTemp As String)
If customerTemp = "AA" Then
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductAA()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
 ElseIf customerTemp = "GC" Then
    
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductBB()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
End If

'Set mainItemRS = GetProduct()
'Set Text3.RowSource = mainItemRS
'Text3.ListField = mainItemRS("productname").Name
'Text3.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub IniFpsHeader()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS0.E_ID, 0, "���"
        .SetText E_FPS0.E_WaferID, 0, "WaferId"
        .SetText E_FPS0.E_CompleteFlag, 0, "��ɱ�־"
        .SetText E_FPS0.E_TotalDie, 0, "TotalDie����"
        .SetText E_FPS0.E_GoodDie, 0, "GoodDie����"
        .SetText E_FPS0.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS0.E_MarkingCode, 0, "MarkingCode"
        .SetText E_FPS0.E_OK, 0, "ѡ��"
        
        
        .ColWidth(E_FPS0.E_ID) = 10
        .ColWidth(E_FPS0.E_WaferID) = 15
        .ColWidth(E_FPS0.E_CompleteFlag) = 10
        .ColWidth(E_FPS0.E_TotalDie) = 12
        .ColWidth(E_FPS0.E_GoodDie) = 12
        .ColWidth(E_FPS0.E_WaferLot) = 10
        .ColWidth(E_FPS0.E_MarkingCode) = 10
        .ColWidth(E_FPS0.E_OK) = 10

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS0.E_OK
        .Lock = False
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsBom()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
      
        
        .SetText E_FPS1.E_ID, 0, "���"
        .SetText E_FPS1.E_BomID, 0, "���Ϲ淶���"
        .SetText E_FPS1.E_PT, 0, "�Ϻ�"
        .SetText E_FPS1.E_Mt, 0, "���ϱ��"
        .SetText E_FPS1.E_Name, 0, "����"
        .SetText E_FPS1.E_Qty, 0, "ÿֻ����"
        .SetText E_FPS1.E_Unit, 0, "��λ"
        
        .SetText E_FPS1.E_Pt2, 0, "�����Ϻ�"
        .SetText E_FPS1.E_Mt2, 0, "�������ϱ��"
        .SetText E_FPS1.E_Name2, 0, "��������"
        .SetText E_FPS1.E_Qty2, 0, "����ÿֻ����"
        .SetText E_FPS1.E_Unit2, 0, "���ϵ�λ"
    
        
        
        .ColWidth(E_FPS1.E_ID) = 6
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
        .ColWidth(E_FPS1.E_Qty) = 10
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Pt2) = 14
        .ColWidth(E_FPS1.E_Mt2) = 14
        .ColWidth(E_FPS1.E_Name2) = 14
        .ColWidth(E_FPS1.E_Qty2) = 10
        .ColWidth(E_FPS1.E_Unit2) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
''���ɹ�����
''��������+��λ����
'Dim FirstChar As String
'Dim SeqChar As String
''2012-11-06 ���¾�ϵͳ����ʱȡ���Զ�����
'
''FirstChar = UCase(Trim(Text2.Text))
'' If KeyAscii = 13 Then
''    If FirstChar = "" Then
''        MsgBox "�����빤��ǰ��λ!"
''        Exit Sub
''    End If
''
''    FirstChar = FirstChar & "-" & Right(Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
''
''    SeqChar = Right("000" & CStr(CInt(GetSeqChar()) + 1), 4)
''
''    Text2.Text = FirstChar & SeqChar
''
''    If Mid$(Trim(Text2.Text), 2, 1) = "P" Then
''        Combo2.Text = "һ�㹤��"
''    End If
''
''    If Mid$(Trim(Text2.Text), 2, 1) = "T" Then
''        Combo2.Text = "С�����Բ�����"
''    End If
''
''
'' End If

'���ɹ�����
'��������+��λ����
Dim FirstChar As String
Dim SeqChar As String
Dim typenameTemp As String
Dim yMonthTemp As String
Dim seqTemp As Integer
Dim headChar As String
Dim mdChar As String
Dim id As Long





'2012-11-06 ���¾�ϵͳ����ʱȡ���Զ�����

FirstChar = UCase(Trim(Text2.Text))
 If KeyAscii = 13 Then
    If FirstChar = "" Then
        MsgBox "�����빤��ǰ��λ!"
        Exit Sub
    End If
    
     If Len(FirstChar) <> 3 Then
        MsgBox "�����빤��ǰ��λ!"
        Exit Sub
    End If

headChar = FirstChar

    SeqChar = GetWoIDTemp(FirstChar)
    mdChar = Right(Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
    FirstChar = FirstChar & "-" & mdChar
    
    SeqChar = Right("000" & CStr(CInt(SeqChar)), 4)
    
    id = CLng(SeqChar)
    
    Text2.Text = FirstChar & SeqChar

    If Mid$(Trim(Text2.Text), 2, 1) = "P" Then
        Combo2.Text = "һ�㹤��"
    End If

    If Mid$(Trim(Text2.Text), 2, 1) = "T" Then
        Combo2.Text = "С�����Բ�����"
    End If
    
    '�����к�д������
    
  cmdStr = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceid,flag) values ( '" & headChar & "','" & mdChar & "'," & id & ", 'Y' ) "
  AddSql (cmdStr)
    
 End If


End Sub

Private Sub Text3_Change()
'ѡ���Ʒ�Ϻţ�����ʾBom
'Dim ptTemp As String
''ptTemp = Text3.Text
'
'ptTemp = "18V117FD00CF"
' Call GetBomData(ptTemp)


Dim deptId As String


TxtWoDept.Text = GetWoDept(Text3.Text)

'���ݲ��Ų����

deptId = GetGWoDeptID(TxtWoDept.Text)

TxtWoDept.Text = TxtWoDept.Text & "_" & deptId


End Sub

