VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Frm_ShippingReceipt 
   Caption         =   "��������ǩ"
   ClientHeight    =   13710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13710
   ScaleWidth      =   20535
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "����ѡ��"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   20055
      Begin VB.ComboBox cbItem 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         ItemData        =   "Frm_ShippingReceipt.frx":0000
         Left            =   1320
         List            =   "Frm_ShippingReceipt.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtFileTitle 
         BackColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10560
         TabIndex        =   11
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   9015
      End
      Begin VB.ComboBox cbCustomerCode 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         ItemData        =   "Frm_ShippingReceipt.frx":001E
         Left            =   1320
         List            =   "Frm_ShippingReceipt.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtDJBH 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   4560
         TabIndex        =   2
         Top             =   1215
         Width           =   2055
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11280
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":0037
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":0C89
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":18DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":252D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":317F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ShippingReceipt.frx":3DD1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   1058
         ButtonWidth     =   1984
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѯ"
               Key             =   "QUERY"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "��"
               Key             =   "OPEN"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ϴ�"
               Key             =   "SAVE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "DEL"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "EXPORT"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "HOME"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12000
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "�����ļ�(*.*)|*.*|PDF�ļ�(*.PDF)|*.PDF"
         Flags           =   524800
         MaxFileSize     =   9999
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   107216897
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   107216897
         CurrentDate     =   41387
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Ŀ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ���(N):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label lblCustomerCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ��:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   6
         Top             =   1230
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   15015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   26485
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Frm_ShippingReceipt.frx":4A23
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fpS_ShippingReceipt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "AcroPDF1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   9255
         Left            =   8640
         TabIndex        =   10
         Top             =   3360
         Visible         =   0   'False
         Width           =   11535
         _cx             =   5080
         _cy             =   5080
      End
      Begin FPSpreadADO.fpSpread fpS_ShippingReceipt 
         Height          =   9015
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   8415
         _Version        =   524288
         _ExtentX        =   14843
         _ExtentY        =   15901
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
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "Frm_ShippingReceipt.frx":4A3F
      End
   End
End
Attribute VB_Name = "Frm_ShippingReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum E_ShippingReceipt
  ' E_CHOOSE = 1
    E_cust = 1
    e_DJBH
    E_UploadDate
    E_UploadUser
    E_Path
    E_END
     
End Enum


    

Private Sub cbItem_Click()
    If cbItem.text = "���ϴ�" Then
        Label3.Visible = False
        Label4.Visible = False
        DTP(0).Visible = False
        DTP(1).Visible = False
    ElseIf cbItem.text = "δ�ϴ�" Then
        Label3.Visible = True
        Label4.Visible = True
        DTP(0).Visible = True
        DTP(1).Visible = True
        AcroPDF1.LoadFile ("")
        AcroPDF1.Visible = False
        
    End If
End Sub



Private Sub Form_Load()
InitCtrl
End Sub

'��ʼ���ؼ�
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim strdjbh              As String
Dim rs                  As New ADODB.Recordset
    
    strdjbh = ""

    '���ؿͻ�����
    strSql = "SELECT DISTINCT �ͻ����� FROM dbo.tblXCustomer  "
    If rs.State = 1 Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
   cbCustomerCode.Clear
    If Not rs.EOF Then
        With cbCustomerCode
            Do While Not rs.EOF
                .AddItem Trim$("" & rs!�ͻ�����)
                rs.MoveNext
            Loop

        End With
    End If
    rs.Close

    
    'Fps��ʼ��
    With fpS_ShippingReceipt
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '�趨������
      '  .Col = E_ShippingReceipt.E_CHOOSE   'ѡ��
     '   .CellType = CellTypeCheckBox
     '   .TypeHAlign = TypeVAlignCenter
     '  .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
   '     .ColWidth(E_ShippingReceipt.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
   DTP(0).Value = Format(Now(), "YYYY/MM/DD")
   DTP(1).Value = Format(Now(), "YYYY/MM/DD")
   cbItem.text = "���ϴ�"
End Sub






Private Sub fpS_ShippingReceipt_Click(ByVal Col As Long, ByVal Row As Long)
    If cbItem.text = "δ�ϴ�" Then
        Exit Sub
    End If
    AcroPDF1.LoadFile ("")
    With fpS_ShippingReceipt
        .Row = Row
        .Col = E_ShippingReceipt.E_Path
    
        If Trim(.text) <> "" Then
            AcroPDF1.Visible = True
            AcroPDF1.LoadFile (Trim(.text))
           ' AcroPDF1.setZoom (100)
            AcroPDF1.setShowToolbar (False)
             
        End If

    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        QueryData


    Case "SAVE"
        SaveData

    Case "DEL"
        delData
        
    Case "EXPORT"
        exportData

    Case "HOME"
        exitFrm

End Select

End Sub



Private Sub QueryData()

    Dim strDN  As String
    Dim strSql As String
    Dim rs     As New ADODB.Recordset

    fpS_ShippingReceipt.MaxRows = 0
    Select Case cbItem.text
    Case "���ϴ�"
    
        strSql = "SELECT �ͻ�����, ���ݱ��, convert(varchar(20),�ϴ�ʱ��,111) as �ϴ�ʱ�� , �ϴ���Ա, �ļ�·�� FROM erpdata..tblShippingReceipt where 1=1 "
    
        If Trim$(txtDJBH.text) <> "" Then
            strSql = strSql & " and ���ݱ��='" & Trim$(txtDJBH.text) & "'"
        Else
            If cbItem.text = "δ�ϴ�" Then
            
                strSql = strSql & " and �ϴ�ʱ�� >='" & Format(DTP(0).Value, "yyyy/mm/dd") & "' and �ϴ�ʱ�� <'" & Format(DTP(1).Value + 1, "yyyy/mm/dd") & "'"
            End If
            
        End If
        If Trim$(txtDJBH.text) <> "" Then
            strSql = strSql & " and �ͻ�����='" & Trim$(cbCustomerCode.text) & "'"
        End If
    
    Case "δ�ϴ�"
       strSql = "SELECT distinct  a.���ݱ��,a.�ͻ����� FROM erpdata..tblStockSQfh a WHERE  left(a.���ݱ��,1)='F' AND a.���ձ��=1 and not exists( select 1 from erpdata..tblShippingReceipt where ���ݱ��=a.���ݱ�� )"
    

        If Trim$(txtDJBH.text) <> "" Then
            strSql = strSql & " and ���ݱ��='" & Trim$(txtDJBH.text) & "'"
        Else
            If cbItem.text = "δ�ϴ�" Then
            
                strSql = strSql & " and �������� >='" & Format(DTP(0).Value, "yyyy/mm/dd") & "' and �������� <'" & Format(DTP(1).Value + 1, "yyyy/mm/dd") & "'"
            End If
            
        End If
        If Trim$(txtDJBH.text) <> "" Then
            strSql = strSql & " and �ͻ�����='" & Trim$(cbCustomerCode.text) & "'"
        End If
    
    End Select
    
    

    'fpS_ShippingReceipt
    Set rs = Get_SqlserveRs(strSql)
    
    If Not rs.EOF Then

        With fpS_ShippingReceipt
            .MaxRows = 0
            Set .DataSource = rs


        End With

    Else
        MsgBox "��ѯ��������", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub
Private Sub exitFrm()
Unload Me

End Sub

Private Sub SaveData()
If cbCustomerCode.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub
End If
If txtDJBH.text = "" Then
    MsgBox "��ѡ�񵥾ݱ��", vbInformation, "��ʾ"
    Exit Sub
End If

If Get_SqlserverCnt("select * from erpdata..tblShippingReceipt where ���ݱ�� = '" & Trim(txtDJBH.text) & "'") > 0 Then
    MsgBox "������ĵ��ݱ�������ϴ���¼,��Ҫ�޸�����ɾ��", vbInformation, "��ʾ"
    Exit Sub

End If


If openFile Then
    UploadData
End If
QueryData
End Sub

Private Function openFile() As Boolean
Dim strFileTitle  As String


On Error GoTo openFile_Err

openFile = False
CommonDialog1.Filter = "PDF�ļ�(*.pdf)|*.pdf"
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then Exit Function
strFileTitle = Replace(UCase(Trim(CommonDialog1.FileTitle)), ".PDF", "")
If UCase(Trim(cbCustomerCode.text)) & "-" & UCase(Trim(txtDJBH.text)) <> strFileTitle Then
    MsgBox "�ļ�����ͻ�����-���ݱ�Ų�һ�£���ȷ�ϣ�", vbInformation, "��ʾ"
    Exit Function
End If

txtFileName.text = Replace(CommonDialog1.filename, Chr(0), ",")
txtFileTitle.text = CommonDialog1.FileTitle
CommonDialog1.filename = ""
'�˶��ļ����뵥�ݱ���Ƿ�һ��

openFile = True
Exit Function
openFile_Err:
MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_ShippingReceipt.openFile ", vbExclamation + vbOKOnly, "Application Error"

Resume Next

End Function

Private Sub UploadData()
Dim strFilePath As String
Dim strSql As String
Dim strNewPath As String

On Error GoTo uploadData_Err

If txtFileName.text = "" Then

    Exit Sub
    
End If

'���Ƶ�����
If Dir("\\10.160.1.84\public\FileServer\37.�ֿ����֪ͨ����ǩ", vbDirectory) = "" Then
    MsgBox "\\10.160.1.84\public\FileServer\37.�ֿ����֪ͨ����ǩ ·�������ڣ��뷴��", vbInformation, "��ʾ"
    Exit Sub
    
End If

strFilePath = "\\10.160.1.84\public\FileServer\37.�ֿ����֪ͨ����ǩ\" & cbCustomerCode.text
If Dir(strFilePath, vbDirectory) = "" Then
    MkDir strFilePath

End If

Call CopyFileToFtp(txtFileName.text, strFilePath & "\")
strNewPath = strFilePath & "\" & txtFileTitle.text
'�ϴ����ݿ�
gUserName = ""
strSql = " insert into erpdata..tblShippingReceipt(�ͻ�����,���ݱ��,�ļ�·��,�ϴ�ʱ��,�ϴ���Ա)" & _
       " values('" & cbCustomerCode.text & "','" & txtDJBH.text & "','" & strNewPath & "', getdate() ,'" & gUserName & "')"

AddSql2 (strSql)

Exit Sub
uploadData_Err:
MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1.Frm_ShippingReceipt.uploadData ", vbExclamation + vbOKOnly, "Application Error"

Resume Next

End Sub


Private Sub delData()

Dim strdjbh As String

If txtDJBH.text = "" Then
    MsgBox "������Ҫɾ���ĵ��ݱ��", vbInformation, "��ʾ"
    Exit Sub

End If

strdjbh = Trim$(txtDJBH.text)

If Get_SqlserverCnt("select * from erpdata..tblShippingReceipt where ���ݱ�� = '" & strdjbh & "'") = 0 Then
    MsgBox "������ĵ��ݱ�Ų���ȷ��û���ϴ���¼,����ɾ��", vbInformation, "��ʾ"
    Exit Sub

End If

gUserName = ""
AddSql2 ("insert into erpdata..tblShippingReceipt_bak select getdate(),'ɾ��','" & gUserName & "',* from erpdata..tblShippingReceipt where ���ݱ�� = '" & strdjbh & "' ")
MsgBox "���ݳɹ�", vbInformation, "��ʾ"
AddSql2 ("delete from erpdata..tblShippingReceipt  where ���ݱ�� = '" & strdjbh & "'")
MsgBox "�ѳɹ�ɾ��:" & strdjbh, vbInformation, "��ʾ"
txtDJBH.text = ""
QueryData

End Sub



Private Sub exportData()
    Call FpsToExcel(fpS_ShippingReceipt)
End Sub

Private Sub FpsToExcel(fps As fpSpread)
    If fps.MaxRows = 0 Then
        MsgBox "û�����ݿ��Ե���", vbInformation, "��ʾ"
        Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    
    Dim xlApp      As Excel.Application
    Dim xlBook     As Excel.Workbook
    Dim xlSheet    As Excel.Worksheet
    

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With xlApp
        .Rows(1).Font.Bold = True
    End With
    
 On Error GoTo Ert
    With fps

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
                xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
            Next j
       
        Next i
        
    End With

    '�����и�ʽ����
    'For j = 1 To Fps.MaxCols
    '    If Trim(xlSheet.Cells(1, j)) = "�ͻ����GoodDie" Or Trim(xlSheet.Cells(1, j)) = "��Ƭ����" Or Trim(xlSheet.Cells(1, j)) = "GOODDIE����" Or Trim(xlSheet.Cells(1, j)) = "����Ƭ��" Or 'Trim(xlSheet.Cells(1, j)) = "����NG" Or Trim(xlSheet.Cells(1, j)) = "�������" Then
   '         For i = 2 To Fps.MaxRows + 1
    '            xlSheet.Cells(i, j) = Replace(xlSheet.Cells(i, j), "'", "")
    '        Next
    '    End If
    'Next
    With xlSheet.Range("2:" & fps.MaxRows + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    xlApp.Application.Visible = True
    
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
Ert:
    If Not (xlApp Is Nothing) Then
        
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    End If
    
    
End Sub












