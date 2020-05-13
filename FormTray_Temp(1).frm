VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
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
   ScaleHeight     =   8445
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab SSTTab0 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   14843
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "批号标签打印"
      TabPicture(0)   =   "FormTray_Temp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtText1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPrint"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtText2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FormTray_Temp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FormTray_Temp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtText2 
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000D&
         Caption         =   "打印"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   960
         MaskColor       =   &H0000FFFF&
         TabIndex        =   3
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txtText1 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生产批号："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   1800
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim sqlDB As String
Dim sqlDBRS As New ADODB.Recordset
Dim qty As Long
Dim label As String
Dim j As Integer
Dim i As Integer



sqlDB = "select a.ht_device || ',' || b.propertyvalue || ',' ||e.test_mtrl_desc || ',' ||a.bonded ,sum(c.gross_die_qty) ,a.bonded " & _
" from shop_order a,shop_order_property b,shop_order_detail c,mappingdatatest d,customeroitbl_test e " & _
" where a.shop_order = '" & txtText1.Text & "' and b.shop_order = a.shop_order and b.propertyname = 'CUST_PART_NUM1' " & _
" and c.shop_order = b.shop_order and d.substrateid = c.wafer_id and to_char(e.id) = d.filename " & _
" group by  a.ht_device,b.propertyvalue,e.test_mtrl_desc,a.bonded "

 If sqlDBRS.State = adStateOpen Then sqlDBRS.Close
    sqlDBRS.Open sqlDB, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
  qty = sqlDBRS.Fields(1).Value / 15000
    
  
  For j = 0 To qty
  
  i = j + 1
  
  If j = 0 Then
   
  label = sqlDBRS.Fields(2).Value + "M-R" + Trim(Str(i)) + "," + sqlDBRS.Fields(0).Value
  
  Else
  
  label = sqlDBRS.Fields(2).Value + "-R" + Trim(Str(i)) + "," + sqlDBRS.Fields(0).Value
  
  End If
  
  Call addLabelTxt(sqlDBRS.Fields(2).Value + "-R" + Trim(Str(i)), label, "\\10.160.1.14\BarCode\37\37TRAY_TEMP\")
 
 Next j

End Sub
