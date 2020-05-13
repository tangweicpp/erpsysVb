VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_OrderHistory 
   Caption         =   "开工单记录"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
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
   ScaleHeight     =   9060
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   14055
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   14055
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         Height          =   720
         Left            =   3720
         TabIndex        =   7
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   720
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   548405249
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   548405249
         CurrentDate     =   41424
      End
      Begin VB.Label lblCloseDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间："
         Height          =   195
         Left            =   3480
         TabIndex        =   2
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间："
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   1320
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_OrderHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdQuery_Click()
Dim cus As String
Dim sqlTemp As String
Dim sql1  As String
Dim sql2 As String
Dim sql3 As String


beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")



'sql1 = "select distinct a.ordername as 工单号, a.product as 料号, a.qty as 数量,a.erpcreatedate as 日期, TO_CHAR(wmsys.wm_concat(distinct b.waferlot)) as LOT from ib_wohistory a" & _
' " left join ib_waferlist b on a.ordername = b.ordername where a.customer = '" & cus & "' and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<=to_date('" + endTime + "','YYYY/MM/DD') group by a.ordername, a.product, a.qty, a.erpcreatedate order by a.ordername "
'



sql1 = "select t3.PRODUCT, t2.料号, t1.工单号, t1.物料编号, t1.用量, t1.实领数量 " & _
"  from erpbase..tblllplan t1 " & _
" inner join ERPBASE.dbo.tblSmainM2 t2 " & _
"    on t1.物料编号 = t2.物料编号 " & _
" inner join erpdata..tblTSVworkorder t3 " & _
"    on t3.ORDERNAME = t1.工单号 " & _
" where t3.erpcreatedate >= CONVERT(VARCHAR(24),'" & beginTime & "' , 121) " & _
"   and t3.erpcreatedate <= CONVERT(VARCHAR(24), '" & endTime & "', 121) " & _
"  order by t1.工单号 "


SqlServerExporToExcel (sql1)

End Sub

Private Sub Form_Load()

DTP1.Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")
DTP2.Value = Format(Now(), "yyyy-MM-dd")

End Sub
