VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmWLPDelivery 
   Caption         =   "WLP������ǩ"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10425
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
   ScaleHeight     =   8250
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      TabCaption(0)   =   "WLP������ǩ��ӡ"
      TabPicture(0)   =   "FrmWLPDelivery.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSubBoxID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPriBoxID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblReelID"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBindBoxID"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrinter"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSubBoxIDList"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPriBoxIDList"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdScan"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSubBoxID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "WLP������ǩ�˶�"
      TabPicture(1)   =   "FrmWLPDelivery.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.TextBox txtSubBoxID 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   795
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdScan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ɨ��"
         Height          =   285
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   435
         Width           =   990
      End
      Begin VB.TextBox txtPriBoxIDList 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtSubBoxIDList 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�˳�"
         Height          =   285
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   435
         Width           =   990
      End
      Begin VB.CommandButton cmdPrinter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "��ӡ>>>>"
         Height          =   285
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   990
      End
      Begin VB.CommandButton cmdBindBoxID 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����>>>>"
         Height          =   285
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   990
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   435
         Width           =   2535
      End
      Begin VB.Label lblReelID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lblPriBoxID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   195
         Left            =   5160
         TabIndex        =   5
         Top             =   7800
         Width           =   525
      End
      Begin VB.Label lblSubBoxID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ѿ���ID"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   7800
         Width           =   705
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D N"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmWLPDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type T_SUBBOXINFO_GD108

    T_PART_NO As String
    T_LOT_NO As String
    T_QUANTITY As Long
    T_SN As String
    T_DATE_CODE As String
    T_SEAL_DATE As String

End Type

Private Type T_PRIBOXINFO_GD108

    T_PART_NO As String
    T_LOT_NO As String
    T_QUANTITY As Long
    T_SN As String
    T_DATE_CODE As String
    T_SEAL_DATE As String

End Type

Private gStrLotID          As String

Private Sub cmdExit_Click()
Unload Me

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cmdScan_Click
' Description:       DN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/9/11-8:43:28
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdScan_Click()
Dim strSql As String
Dim strDN  As String

If txtDN.Text = "" Then
    MsgBox "DN����Ϊ��", vbInformation, "��ʾ"
    Exit Sub

End If

strDN = Trim$(txtDN.Text)
'���DN��¼���Ƿ��и�DN
strSql = "SELECT * FROM erptemp..ht_dn where DN_NUM = '" & strDN & "'"
If Get_SqlserverCnt(strSql) = 0 Then
    MsgBox "DN������,��ȷ��DN�Ƿ���ȷ", vbInformation, "��ʾ"
    txtDN.Text = ""
    Exit Sub

End If

'����DN�Ƿ��д�ӡ��¼
strSql = "select * from packing_detailed_gd108 where ship_dn = '" & strDN & "'"
If Get_OracleCnt(strSql) > 0 Then
    MsgBox "DN�Ѵ��ڴ�ӡ��¼,�޷��ٴδ�ӡ", vbInformation, "��ʾ"
    txtDN.Text = ""
    Exit Sub

End If

txtSubBoxID.Visible = True
txtSubBoxID.SetFocus

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       txtSubBoxID_KeyPress
' Description:       ɨ��С���ID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/9/11-8:43:07
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtSubBoxID_KeyPress(KeyAscii As Integer)
Dim strSql      As String
Dim strSubBoxID As String
Dim strLotID    As String

If KeyAscii <> vbKeyReturn Then Exit Sub
If Trim(txtSubBoxID.Text) = "" Then Exit Sub
strSubBoxID = UCase(Trim(txtSubBoxID.Text))
'����Ƿ��ظ�ɨ��
If InStr(txtSubBoxIDList.Text, strSubBoxID) > 0 Then
    MsgBox "���: " & strSubBoxID & "��ɨ��,�����ظ�ɨ��", vbInformation, "��ʾ"
    txtSubBoxID.Text = ""
    Exit Sub

End If

'��������Ƿ��и����
strSql = "select ������ from erpdata..tblStockNumSub where ��� = '" & strSubBoxID & "' "
strLotID = Get_SqlStr(strSql)
If strLotID = "" Then
    MsgBox "����ֲ��������: " & strSubBoxID & " ,��ȷ�����ID�Ƿ���ȷ", vbInformation, "��ʾ"
    txtSubBoxID.Text = ""
    Exit Sub
Else
    If gStrLotID = "" Then
        gStrLotID = strLotID
    Else
        If strLotID <> gStrLotID Then
            MsgBox "��ɨ��ͬһ����: " & gStrLotID & "����� ,��ͬ���β��ɻ��", vbInformation, "��ʾ"
            txtSubBoxID.Text = ""
            Exit Sub

        End If

    End If

End If

txtSubBoxIDList.Text = strSubBoxID & "','" & txtSubBoxIDList.Text
txtSubBoxID.Text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cmdBindBoxID_Click
' Description:       ����
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/9/11-8:59:23
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdBindBoxID_Click()
Dim t_PriBoxInfo As T_PRIBOXINFO_GD108
Dim strArr() As String

t_PriBoxInfo.T_LOT_NO = gStrLotID


'GD108
If txtSubBoxIDList.Text = "" Then
    MsgBox "û�о������,�����Ժ���", vbInformation, "��ʾ"
    Exit Sub
End If

strArr = Split(txtSubBoxIDList.Text, vbCrLf)


End Sub
