VERSION 5.00
Begin VB.Form FrmSplitLot 
   Caption         =   "拆批校验"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
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
   ScaleHeight     =   5865
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtMappingQty 
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox TxtAEFQty 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox TxtAQty 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox TxtRejectCodeQty 
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox TxtEQty 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox TxtCustmerFQty 
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtFQty 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   360
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox TxtWaferID 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapping总数："
      Height          =   195
      Left            =   7560
      TabIndex        =   15
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-A数量："
      Height          =   195
      Left            =   960
      TabIndex        =   14
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-A+-E+-F："
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RejectCode量："
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-E数量："
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户来料不良："
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-F数量："
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label LblWaferID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaferID："
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "FrmSplitLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim waferidTemp As String
waferidTemp = Trim(TxtWaferID.Text)
TxtFQty.Text = GetWaferFQty(waferidTemp)
TxtCustmerFQty.Text = GetWaferCustomerFQty(waferidTemp)

TxtEQty.Text = GetWaferEQty(waferidTemp)
TxtRejectCodeQty.Text = GetWafeRejectCodeQty(waferidTemp)

TxtAQty.Text = GetWaferAQty(waferidTemp)

TxtMappingQty.Text = GetWaferCustomerMapQty(waferidTemp)

TxtAEFQty.Text = CInt(TxtFQty.Text) + CInt(TxtEQty.Text) + CInt(TxtAQty.Text)



End Sub
