VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SetData_bak 
   Caption         =   "test"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
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
   ScaleHeight     =   10995
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   13095
      _Version        =   524288
      _ExtentX        =   23098
      _ExtentY        =   13996
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
      SpreadDesigner  =   "Form07.frx":0000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   1640
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      Appearance      =   1
      ImageList       =   "ImageListtest"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询"
            Key             =   "s"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "i"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "u"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "d"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "t"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageListtest 
         Left            =   11640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form07.frx":0414
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form07.frx":1066
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form07.frx":1CB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form07.frx":290A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form07.frx":355C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q_DATE"
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM_COME"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM_NAME"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   870
   End
End
Attribute VB_Name = "Frm_SetData_bak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ForQuery
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key

        Case "s"
            ForQuery
        Case "i"
            Insert
        Case "u"
            Update
         Case "d"
            Delete
        Case "t"
            Unload Me
        
    End Select
End Sub
Private Sub ForQuery()
    Dim rs     As New ADODB.Recordset

    Dim strSql As String
    
    Dim itemName As String
    
    itemName = Text1.Text

    If Text1.Text = "" Then
        strSql = "select ITEM_NAME,ITEM_CODE,Q_DATE,'' as '√' from erptemp..Material_Shelf_Life"
    Else
        strSql = "select ITEM_NAME,ITEM_CODE,Q_DATE,'' as '√' from erptemp..Material_Shelf_Life where ITEM_NAME='" & itemName & "'"
    End If
    
    fpSpread1.MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "暂时没有数据", vbInformation, "提示"
        Exit Sub

    End If

End Sub
Private Sub ListDataType(rs As ADODB.Recordset)
    Dim i As Long

    With fpSpread1
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    With fpSpread1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            .ColWidth(4) = 2
            .CellType = CellTypeCheckBox
        Next
        
    End With
End Sub
Private Sub Insert()
    Dim rs     As New ADODB.Recordset

    Dim strSql As String
    
    Dim itemName As String
    
    Dim itemCome As String
    
    Dim qDate As String
    
    
    If Text1.Text = "" Then
        MsgBox "请输入ITEM_NAME", vbInformation, "提示"
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "请输入ITEM_COME", vbInformation, "提示"
        Exit Sub
    End If
    If Text3.Text = "" Then
        MsgBox "请输入Q_DATE", vbInformation, "提示"
        Exit Sub
         Else
        If Not IsNumeric(Text3.Text) Then
             MsgBox "类型必须为数值", vbInformation, "提示"
        Exit Sub
        End If
    End If
    
    itemName = Trim$(Text1.Text)
    itemCome = Trim$(Text2.Text)
    qDate = Trim$(Text3.Text)
    
    Rem 校验料号是否存在
    strSql = "SELECT * FROM erpdata..tblSmainM2 a WHERE a.料号 IN ('" & itemName & "')"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    '存在则新增
    If Not (rs.EOF And rs.BOF) Then
        AddSql2 ("INSERT INTO [erptemp].[dbo].[Material_Shelf_Life]([ITEM_NAME],[ITEM_CODE],[Q_DATE],CREATE_BY,CREATE_DATE)VALUES('" & itemName & "','" & itemCome & "'," & qDate & ",'" & gUserName & "',CONVERT(varchar(10),getdate(),120))")
    Else
        '否则退出
         MsgBox "没有该料号，请重新输入", vbInformation, "提示"
         Exit Sub
    End If
    
    MsgBox "新增成功", vbInformation, "提示"
    '新增后清空文本框中的值
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    ForQuery
End Sub
Private Sub Update()

    Dim bFlag As Boolean

    bFlag = False

    With fpSpread1

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要修改的行", vbInformation, "提示"
        Exit Sub

    End If
   Dim itemName As String
    
    Dim itemCome As String
    
    Dim qDate As String
    
    With fpSpread1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            
            If .Text = "1" Then
                .Col = 1
                itemName = Trim$(.Text)
                
                .Col = 2
                itemCome = Trim$(.Text)
                
                .Col = 3
                qDate = Trim$(.Text)
                
                AddSql2 ("INSERT INTO [erptemp].[dbo].[Material_Shelf_Life_bak]([ITEM_NAME],[ITEM_CODE],[Q_DATE],CREATE_BY,CREATE_DATE)select [ITEM_NAME],[ITEM_CODE],[Q_DATE],'" & gUserName & "',getdate() from erptemp..Material_Shelf_Life where ITEM_NAME='" & itemName & "'")
                AddSql2 ("update [erptemp].[dbo].[Material_Shelf_Life] set [ITEM_CODE] = '" & itemCome & "', [Q_DATE] = '" & qDate & "',CREATE_BY='" & gUserName & "',CREATE_DATE=Convert(VarChar(10), getdate(), 120) where [ITEM_NAME] = '" & itemName & "'")
            
            End If
            
        Next
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    ForQuery
End Sub
Private Sub Delete()

    Dim i As Integer
    
    Dim bFlag As Boolean
    
    Dim itemName As String

    bFlag = False

    With fpSpread1

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Value = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strno As String
    
    With fpSpread1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                .Col = 1
                itemName = Trim$(.Text)
                AddSql2 ("INSERT INTO [erptemp].[dbo].[Material_Shelf_Life_bak]([ITEM_NAME],[ITEM_CODE],[Q_DATE],CREATE_BY,CREATE_DATE)select [ITEM_NAME],[ITEM_CODE],[Q_DATE],'" & gUserName & "',getdate() from erptemp..Material_Shelf_Life where ITEM_NAME='" & itemName & "'")
                AddSql2 ("delete from [erptemp].[dbo].[Material_Shelf_Life] where ITEM_NAME = '" & itemName & "'  ")
            
            End If
            
        Next
    
    End With

    ForQuery
    
End Sub
