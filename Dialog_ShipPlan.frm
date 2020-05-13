VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Dialog_ShipPlan 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "出货日期选择"
   ClientHeight    =   660
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DT_ShipDate 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   245891073
      CurrentDate     =   43594
   End
End
Attribute VB_Name = "Dialog_ShipPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
frmShippingScheduleSystem.txtShipDate = ""
Unload Me
End Sub

Private Sub Form_Load()
frmShippingScheduleSystem.txtShipDate = ""
DT_ShipDate.Value = Format(Now() + 1, "yyyy-MM-dd")

End Sub

Private Sub OKButton_Click()

frmShippingScheduleSystem.txtShipDate = Format(DT_ShipDate.Value, "yyyy-MM-dd")
Unload Me

End Sub
