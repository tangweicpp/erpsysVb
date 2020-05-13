VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_UploadCusShipAddress 
   Caption         =   "上传客户发货地址"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
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
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   11295
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   20175
      Begin VB.CommandButton Command5 
         Caption         =   "修改"
         Height          =   600
         Left            =   6240
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9480
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3480
         TabIndex        =   12
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查找"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   840
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cbCusCode 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   2655
         Index           =   1
         Left            =   -240
         TabIndex        =   13
         Top             =   2640
         Width           =   17535
         _Version        =   524288
         _ExtentX        =   30930
         _ExtentY        =   4683
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "Frm_UploadCusShipAddress.frx":0000
         Appearance      =   2
         TextTip         =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHIP_TO:"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   540
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   20175
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8160
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   495
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   315
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   570
      End
   End
End
Attribute VB_Name = "Frm_UploadCusShipAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sCus As String
Dim sShip As String


Private Sub cbCusCode_Click()
    InitShipTo
End Sub



Private Sub cmdDel_Click()

    Dim strsql As String

    If cbCusCode.Text = "" Then
        MsgBox "请输入客户代码", vbCritical, "提示"

        Exit Sub
    Else
        strCus = Trim$(cbCusCode.Text)

    End If

    If comBo2.Text = "" Then
        MsgBox "请输入SHIP_TO", vbCritical, "提示"

        Exit Sub
    Else
        strShipTo = Trim$(comBo2.Text)

    End If

    strsql = "delete from erptemp..customer_information where  customer = '" & strCus & "' and ship_to = '" & strShipTo & "'"

    If MsgBox("确认要删除", vbOKCancel) = vbOK Then
        AddSql2 (strsql)
        MsgBox "删除成功", vbInformation, "提示"
    
        Call cmdQuery_Click
    End If

End Sub

Private Sub cmdQuery_Click()

    Dim strCus    As String

    Dim strShipTo As String

    Dim rs        As ADODB.Recordset

    If cbCusCode.Text = "" Then
        MsgBox "请输入客户代码", vbCritical, "提示"

        Exit Sub

    End If

    If comBo2.Text = "" Then
        MsgBox "请输入SHIP_TO", vbCritical, "提示"

        Exit Sub

    End If

    strCus = Trim(cbCusCode.Text)
    strShipTo = Trim$(comBo2.Text)

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "SELECT customer, SHIP_TO,SHIPPER,SOLD_TO,BILL_TO,SHIP_TO_AD,SOLD_BY, PAYMENT_TERMS, CURRENCY, BANK_INFORMATION,TK,PO,SHIPPER_PACK FROM erptemp..customer_information where customer = '" & strCus & "' and ship_to = '" & strShipTo & "'"

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        With Fps(1)
            .MaxRows = 0
            Set .DataSource = rs
        End With

    Else
        MsgBox "查询不到信息", vbExclamation, "警告"

        Exit Sub

    End If
  
End Sub

Private Sub Command1_Click()

    On Error Resume Next

    Dim FName
 
    CommonDialog1.Filter = "EXCEL文件(*.xlsx)|*.xlsx|CSV文件(*.csv)|*.csv"
    
    CommonDialog1.ShowOpen

    FName = CommonDialog1.filename

    If FName <> "" Then
        txtPath.Text = FName
    End If
    
End Sub

Private Sub Command2_Click()

    Dim ship     As CUSSHIPADDRESS

    Dim dirName  As String

    Dim filename As String

    If txtPath.Text = "" Then
        MsgBox "先选择待上传的文件", vbCritical, "警告"

        Exit Sub

    End If

    Set VBExcel = CreateObject("excel.application")

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)

    Set xlSheet = xlBook.Worksheets(1)

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"

        Exit Sub

    End If

    Dim i       As Integer

    Dim J       As Integer

    Dim ID      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

        TEMP = ""

        For J = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
            strChar = Chr(96 + J)
            tempVal = Replace(xlSheet.Range(strChar & i).Value, "'", "")

            If J = 1 Then

                ship.CUSTOMER = Trim(tempVal)

            ElseIf J = 2 Then
                ship.ship_to = Trim(tempVal)

            ElseIf J = 3 Then
                ship.SHIPPER = Trim(tempVal)

            ElseIf J = 4 Then
                ship.SOLD_TO = Trim(tempVal)

            ElseIf J = 5 Then
                ship.BILL_TO = Trim(tempVal)
                
            ElseIf J = 6 Then
                ship.SHIP_TO_AD = Trim$(tempVal)
                
            ElseIf J = 7 Then
                ship.SOLD_BY = Trim$(tempVal)
                
            ElseIf J = 8 Then
                ship.PAYMENT_TERMS = Trim$(tempVal)
                
            ElseIf J = 9 Then
                ship.CURRENCY = Trim$(tempVal)
                
            ElseIf J = 10 Then
                ship.BANKINFO = Trim$(tempVal)
                
            ElseIf J = 11 Then
                ship.TK = Trim$(tempVal)
                
            ElseIf J = 12 Then
                ship.PO = Trim$(tempVal)
                
            ElseIf J = 13 Then
                ship.SHIPPER_PACK = Trim$(tempVal)
                
            End If

        Next J

        Dim strCustomer As String

        Dim ship_to     As String
       
        strCustomer = ship.CUSTOMER
        ship_to = ship.ship_to
       
        If CheckIsExist(strCustomer, ship_to) = True Then
            Call UpdateCustomerShipAddress2(ship)
        
        Else
            Call AddCustomerShipAddress(ship)
        End If

NextRecord2:

    Next i

    xlBook.Close

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    MsgBox "上传完成, 请导出确认", vbInformation, "友情提醒"

End Sub

Private Function CheckIsExist(strCus As String, strShipTo As String) As Boolean

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select * from erptemp..customer_information where customer = '" & strCus & "' and SHIP_TO = '" & strShipTo & "' "

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckIsExist = True
    Else
        CheckIsExist = False
    End If
  
    rs.Close
    Set rs = Nothing

End Function

Private Sub Command3_Click()
    SqlServer2ExporToExcel ("select * from erptemp..customer_information")

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()

    Dim ship As CUSSHIPADDRESS
    
    If Fps(1).MaxRows = 0 Then
        MsgBox "请先查找对应信息", vbCritical, "提示"
        Exit Sub
    End If
    
    With Fps(1)
    
        .Row = 1
        
        .Col = 1
        ship.CUSTOMER = Trim$(.Text)
        
        .Col = 2
         ship.ship_to = Trim$(.Text)
         
        .Col = 3
        ship.SHIPPER = Trim$(.Text)
        
        .Col = 4
         ship.SOLD_TO = Trim$(.Text)
        .Col = 5
           ship.BILL_TO = Trim$(.Text)
        .Col = 6
        ship.SHIP_TO_AD = Trim$(.Text)
        .Col = 7
        ship.SOLD_BY = Trim$(.Text)
        .Col = 8
            ship.PAYMENT_TERMS = Trim$(.Text)
        .Col = 9
           ship.CURRENCY = Trim$(.Text)
        .Col = 10
          
         ship.BANKINFO = Trim$(.Text)
        .Col = 11
            ship.TK = Trim$(.Text)
        .Col = 12
      ship.PO = Trim$(.Text)
      .Col = 13
      ship.SHIPPER_PACK = Trim$(.Text)
      
      
        Dim strCustomer As String

        Dim ship_to     As String
       
        strCustomer = ship.CUSTOMER
        ship_to = ship.ship_to
       
'        If CheckIsExist(strCustomer, ship_to) = True Then
'            Call UpdateCustomerShipAddress(ship)
'
'        Else
'            Call AddCustomerShipAddress(ship)
'        End If
      
      Call UpdateCustomerShipAddress(ship)

      MsgBox "修改成功", vbInformation, "提示"
      
      cbCusCode.Text = ship.CUSTOMER
      comBo2.Text = ship.ship_to
      
      Call cmdQuery_Click
      
    End With
   
End Sub
Private Sub UpdateCustomerShipAddress2(ship As CUSSHIPADDRESS)
Dim strsql  As String

strsql = "update erptemp..customer_information set CUSTOMER='" & ship.CUSTOMER & "', SHIP_TO='" & ship.ship_to & "', SHIPPER= '" & ship.SHIPPER & "', SOLD_TO='" & ship.SOLD_TO & "', BILL_TO= '" & ship.BILL_TO & "', SHIP_TO_AD='" & ship.SHIP_TO_AD & "', SOLD_BY='" & ship.SOLD_BY & "', PAYMENT_TERMS= '" & ship.PAYMENT_TERMS & "', CURRENCY='" & ship.CURRENCY & "', BANK_INFORMATION= '" & ship.BANKINFO & "', TK='" & ship.TK & "', PO='" & ship.PO & "', LASTUPDATE_BY='" & gUserName & "', LASTUPDATE_DATE= GETDATE(), SHIPPER_PACK='" & ship.SHIPPER_PACK & "'   " & _
"where CUSTOMER = '" & ship.CUSTOMER & "' and ship_to = '" & ship.ship_to & "'"

If AddSql2(strsql) = 0 Then

    MsgBox "更新失败, 请导出确认", vbCritical, "警告"

End If

End Sub

Private Sub UpdateCustomerShipAddress(ship As CUSSHIPADDRESS)
Dim strsql  As String

strsql = "update erptemp..customer_information set CUSTOMER='" & ship.CUSTOMER & "', SHIP_TO='" & ship.ship_to & "', SHIPPER= '" & ship.SHIPPER & "', SOLD_TO='" & ship.SOLD_TO & "', BILL_TO= '" & ship.BILL_TO & "', SHIP_TO_AD='" & ship.SHIP_TO_AD & "', SOLD_BY='" & ship.SOLD_BY & "', PAYMENT_TERMS= '" & ship.PAYMENT_TERMS & "', CURRENCY='" & ship.CURRENCY & "', BANK_INFORMATION= '" & ship.BANKINFO & "', TK='" & ship.TK & "', PO='" & ship.PO & "', LASTUPDATE_BY='" & gUserName & "', LASTUPDATE_DATE= GETDATE(), SHIPPER_PACK='" & ship.SHIPPER_PACK & "'   " & _
"where CUSTOMER = '" & Trim(cbCusCode.Text) & "' and ship_to = '" & Trim$(comBo2.Text) & "'"

If AddSql2(strsql) = 0 Then

    MsgBox "更新失败, 请导出确认", vbCritical, "警告"

End If

End Sub


Private Sub Form_Load()
    InitFps
    InitCuscode
    ' InitShipTo
End Sub

Private Sub InitFps()

    With Fps(1)
        .ReDraw = False
        
        .MaxRows = 0
        
        .DAutoHeadings = True
        .DAutoCellTypes = True
        .DAutoSizeCols = DAutoSizeColsMax
        '
        '        .Col = -1
        '        .Row = -1
        '        .Lock = True
        '        .OperationMode = OperationModeNormal
        '        .TypeVAlign = TypeVAlignCenter
        '        .SelForeColor = &HFF8080
        '
        '        .Col = 1
        '        .CellType = CellTypeCheckBox
        '        .TypeHAlign = TypeHAlignCenter
        '        .TypeVAlign = TypeVAlignCenter
       
    End With

End Sub

Private Sub InitCuscode()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "SELECT distinct Customer FROM erptemp..customer_information "

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    cbCusCode.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbCusCode.AddItem Trim(rs("Customer"))
            rs.MoveNext
        Next i

    End If
  
    rs.Close
    Set rs = Nothing

End Sub

Private Sub InitShipTo()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    
    If cbCusCode.Text = "" Then

        Exit Sub

    Else
        rs.Source = "SELECT distinct SHIP_TO FROM erptemp..customer_information  where Customer = '" & Trim$(cbCusCode.Text) & "' "
    End If
    
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    comBo2.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            comBo2.AddItem Trim(rs("SHIP_TO"))
            rs.MoveNext
        Next i

    End If
  
    rs.Close
    Set rs = Nothing

End Sub
