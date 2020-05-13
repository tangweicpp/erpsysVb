VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UploadShipSide 
   Caption         =   "�ϴ�������ַ"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
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
   ScaleHeight     =   7665
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "�ϴ�"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6360
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H000000FF&
         Caption         =   ".."
         Height          =   405
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FF80FF&
         Caption         =   "��������"
         Height          =   840
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FFFF00&
         Caption         =   "�ϴ�"
         Height          =   840
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·��:"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   420
      End
   End
End
Attribute VB_Name = "Frm_UploadShipSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShipSideTmp As ShipSideData

Private Sub cmd_Click()

On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xlsx)|*.xlsx"
    
    CommonDialog1.ShowOpen
    '�õ��ļ���
    FName = CommonDialog1.filename
    If FName <> "" Then
       txtText1.Text = FName
    End If
    
End Sub

Private Sub cmdOpen_Click()

Dim source_batch_id_Temp As String
Dim dirName As String
Dim filename As String

If txtText1.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(txtText1.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 5 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String



SumCount = 0
BCResultFlag = False

 ShipSideTmp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

    temp = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

        If j = 1 Then

            ShipSideTmp.CustomerCode = Trim(tempVal)

        ElseIf j = 2 Then
            ShipSideTmp.GULFDeviceName = Trim(tempVal)

        ElseIf j = 3 Then
            ShipSideTmp.GULFLotID = Trim(tempVal)

        ElseIf j = 4 Then
            ShipSideTmp.WaferQTY = Trim(tempVal)

        ElseIf j = 5 Then
            ShipSideTmp.ShipTo = Trim(tempVal)

        End If

    Next j

    If (JudgeShipSideData(ShipSideTmp.GULFLotID)) Then
       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
       GoTo NextRecord2

    End If


    Call AddShipSideData(ShipSideTmp)
    SumCount = SumCount + 1

    '�ϴ���DB
NextRecord2:

Next i



     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"

Else
        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
            Exit Sub
        End If

End If

End Sub



Private Sub CmdSave_Click()

SqlServer2ExporToExcel ("SELECT ID, CustCode, DeviceName, LotID, WaferQty, ShipTo, Memo " & _
"FROM tblSale_Shipto ")

End Sub
