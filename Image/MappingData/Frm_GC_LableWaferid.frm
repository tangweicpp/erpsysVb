VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Frm_GC_LableWaferid 
   Caption         =   "GC WLA ��ұ�ǩ Waferid�趨"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
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
   ScaleHeight     =   9300
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "����WLA"
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   5640
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��WLA"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   480
      Left            =   3720
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�޸�"
      Height          =   480
      Left            =   1560
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox TxtWaferID 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5040
      Width           =   3255
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "WO�ϴ� �����趨��ұ�ǩ�ϵ�WaferID"
      Height          =   2535
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   9855
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton Command6 
         Caption         =   ".."
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�ϴ�DB"
         Height          =   480
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "����"
         Height          =   480
         Left            =   4440
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3000
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ����ϴ���CSV��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   1545
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaferID��"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   5160
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID��"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   4440
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(���԰�LotID��WaferID���޸�)"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�޸�WLA��ǣ�"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   3480
      Width           =   1230
   End
End
Attribute VB_Name = "Frm_GC_LableWaferid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim mapTemp As MapRecord
Dim gcHeaderTemp As GCHeader
Dim eqISHeaderTemp As EQISHeader

Dim gcDetailTemp As GCDetail
'Dim SumCount As Integer
Dim ErrorInf As String

Dim updateRS                As New ADODB.Recordset



Private Sub Cmd_Click()
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    Com.Filter = "XML�ļ�(*.xml)|*.xml"
    Com.ShowOpen
    '�õ��ļ���
    FName = Com.FileName
    If FName <> "" Then
       Text1.Text = Replace(FName, Chr(160), ",")
    End If
End Sub

Private Sub CmdClearOI_Click()
ClearData
End Sub

Private Sub ClearData()
TxtCustomer.Text = ""
TxtPO.Text = ""
TxtPOItem.Text = ""
TxtLotID.Text = ""
TxtMpn.Text = ""


TxtMpnDesc.Text = ""
TxtWaferQty.Text = ""
TxtDieQty.Text = ""
TxtDesign.Text = ""
TxtCountryFab.Text = ""

TxtImageRev.Text = ""
TxtFFacility.Text = ""
TxtMarkId.Text = ""
TxtLotPriority.Text = ""
TxtFilmApld.Text = ""

TxtShip260.Text = ""
TxtShipLevel.Text = ""
TxtMicMaterial.Text = ""
TxtShipSite.Text = ""
TxtLotStatus.Text = ""

TxtCustomer.SetFocus



End Sub


Private Sub CmdSaveOI_Click()
Dim oiRecordTemp As OIRecord

If TxtWaferQty.Text = "" Then
MsgBox "Ƭ��������Ϊ�գ�"
Exit Sub
End If

If TxtDieQty.Text = "" Then
MsgBox "Ƭ��������Ϊ�գ�"
Exit Sub
End If

oiRecordTemp.id = GetMaxID()
oiRecordTemp.PoNum = Trim(TxtPO.Text)
oiRecordTemp.PoItem = Trim(TxtPOItem.Text)
oiRecordTemp.lotid = Trim(TxtLotID.Text)
oiRecordTemp.MPN = Trim(TxtMpn.Text)
oiRecordTemp.MPNDec = Trim(TxtMpnDesc.Text)


oiRecordTemp.WaferQty = CInt(Trim(TxtWaferQty.Text))
oiRecordTemp.DieQty = CInt(Trim(TxtDieQty.Text))
oiRecordTemp.DesignId = Trim(TxtDesign.Text)
oiRecordTemp.CountryFab = Trim(TxtCountryFab.Text)
oiRecordTemp.ImageRev = Trim(TxtImageRev.Text)

oiRecordTemp.FFacility = Trim(TxtFFacility.Text)
oiRecordTemp.MarkId = Trim(TxtMarkId.Text)
oiRecordTemp.LotPriority = Trim(TxtLotPriority.Text)
oiRecordTemp.FilmApld = Trim(TxtFilmApld.Text)
oiRecordTemp.Ship260 = Trim(TxtShip260.Text)


oiRecordTemp.ShipLevel = Trim(TxtShipLevel.Text)
oiRecordTemp.MicMaterial = Trim(TxtMicMaterial.Text)
oiRecordTemp.ShipSite = Trim(TxtShipSite.Text)
oiRecordTemp.LotStatus = Trim(TxtLotStatus.Text)
oiRecordTemp.customerName = Trim(TxtCustomer.Text)

oiRecordTemp.Flag = "Y"
oiRecordTemp.CreateBy = "Auto"


Call AddOIRecord(oiRecordTemp)



ClearData

End Sub

Private Sub Qtech_OrderMapping()

   SumCount = 0
    ErrorInf = ""
    If Text1.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
    
    End If
    
    Dim FileName As String
    FileName = Text1.Text
    Dim dirtemp() As String
    Dim i As Integer
    
    If InStr(1, FileName, ",") > 0 Then
        dirtemp = Split(FileName, ",")
        
        For i = 1 To UBound(dirtemp)
            UpMxlForQtech (dirtemp(0) + "\" + dirtemp(i))
        Next
        
    Else
        
        UpMxlForQtech (FileName)
    End If
    
    
    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
    If ErrorInf <> "" Then
           MsgBox "�ϴ�ʧ�ܵ���:" + ErrorInf + "���ݿ����Ѵ��ڣ�"
    End If


End Sub

Private Sub Command1_Click()
'SumCount = 0
''2013-01-21 jiayun add  Qtech �Թ� Mapping
'If Combo1.Text = "�Թ�" Then
'
'Qtech_OrderMapping
'
'Else
'
'    SumCount = 0
'    ErrorInf = ""
'    If Text1.Text = "" Then
'    MsgBox "��ѡ����ϴ����ļ�"
'    Exit Sub
'
'    End If
'
'    Dim FileName As String
'    FileName = Text1.Text
'    Dim dirtemp() As String
'    Dim i As Integer
'
'    If InStr(1, FileName, ",") > 0 Then
'        dirtemp = Split(FileName, ",")
'
'        For i = 1 To UBound(dirtemp)
'            UpMxl (dirtemp(0) + "\" + dirtemp(i))
'        Next
'
'    Else
'
'        UpMxl (FileName)
'    End If
'
'
'    If SumCount > 0 Then
'        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
'    End If
'
'    If ErrorInf <> "" Then
'           MsgBox "�ϴ�ʧ�ܵ���:" + ErrorInf + "���ݿ����Ѵ��ڣ�"
'    End If
'
'End If

Dim lotIDTemp As String
Dim waferIdTemp As String
Dim statusFlag As String
Dim sqltemp As String



If Option2.Value = True Then
    statusFlag = "����WLA"

Else
    statusFlag = "��WLA"

End If


lotIDTemp = UCase(Trim(TxtLotID.Text))
waferIdTemp = UCase(Trim(TxtWaferID.Text))


If UCase(Trim(TxtLotID.Text)) = "" And UCase(Trim(TxtWaferID.Text)) = "" Then
     MsgBox "������LotID��WaferId��"
     Exit Sub

End If


If UCase(Trim(TxtLotID.Text)) <> "" Then
'��lotid����


    If (Not JudgeGCLableLotIDData(lotIDTemp)) Then
       MsgBox "��ʣ�" & lotIDTemp & " �����ڣ������޸ģ�", vbInformation, "������ʾ"
    Exit Sub
    
    End If



sqltemp = " update TSV_GCLable_SETWLA  set WLAFlag='" & statusFlag & "' ,LASTUPDATEDATE =sysdate ,LASTUPDATEBY='" & gUserName & "'   where lotid='" & lotIDTemp & "'"

AddSql (sqltemp)

    MsgBox "�޸ĳɹ���", vbInformation, "������ʾ"


Else

    If (Not JudgeGCLableWaferIDData(waferIdTemp)) Then
       MsgBox "��ʣ�" & lotIDTemp & " �����ڣ������޸ģ�", vbInformation, "������ʾ"
    Exit Sub
    
    End If

sqltemp = " update TSV_GCLable_SETWLA  set WLAFlag='" & statusFlag & "' ,LASTUPDATEDATE =sysdate ,LASTUPDATEBY='" & gUserName & "'   where Waferid='" & waferIdTemp & "'"

AddSql (sqltemp)

    MsgBox "�޸ĳɹ���", vbInformation, "������ʾ"
 
End If




End Sub

Private Sub UpMxl(dirtemp As String)


'--����XML

Dim XMLDoc As DOMDocument
Dim xn As IXMLDOMNode
Dim xn01 As IXMLDOMNode
Dim xn02 As IXMLDOMNode
Dim xn03 As IXMLDOMNode
Dim Flag As Integer
Dim JudgeFlag As Boolean

Dim goodDieQty As Integer
Dim badDieQty As Integer

Dim customerNameTemp As String
customerNameTemp = ""

customerNameTemp = Combo1.Text

If customerNameTemp = "" Then
    customerNameTemp = "AA"
End If

                

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)


Set XMLDoc = New DOMDocument
XMLDoc.Load dirtemp

Set xn = XMLDoc.documentElement
'SumCount = 0

If Not xn Is Nothing Then
'ѭ�� Map

    For Each xn01 In xn.childNodes
        JudgeFlag = False
        goodDieQty = 0
        badDieQty = 0

        mapTemp.SubstrateId = xn01.Attributes(1).nodeValue
        
       ' �ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
'          MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
          ErrorInf = ErrorInf + "," + mapTemp.SubstrateId

          GoTo NextRecord

       End If


        mapTemp.SubstrateType = xn01.Attributes(2).nodeValue

        'ѭ�� Device
        If xn01.nodeName = "Map" Then
            For Each xn02 In xn01.childNodes

                mapTemp.lotid = xn02.Attributes(1).nodeValue
                mapTemp.lotid = Replace$(mapTemp.lotid, ".", "")
                mapTemp.ProductId = xn02.Attributes(6).nodeValue
                mapTemp.CreateDate = xn02.Attributes(8).nodeValue
                mapTemp.MicronLotId = xn02.Attributes(14).nodeValue
                mapTemp.MicronLotId = Replace$(mapTemp.MicronLotId, ".", "")

                'ѭ�� ReferenceDevice
                If xn02.nodeName = "Device" Then
                    Flag = 0
                    For Each xn03 In xn02.childNodes
                        '������һ�еģ�������ʱ����
                        Dim field1 As String
                        Dim field2 As String
                        Dim field3 As String
                        Dim field1Value As String
                        Dim field2Value As String
                        Dim field3Value As String
                        
                        If xn03.nodeName = "Bin" Then
                            '2012-10-25 ����ֻ�������ؼ��� BinCode ,BinCount ,BinQuality
                            field1 = xn03.Attributes(0).nodeName
                            field1Value = xn03.Attributes(0).nodeValue
                            
                            field2 = xn03.Attributes(1).nodeName
                            field2Value = xn03.Attributes(1).nodeValue
                            
                            field3 = xn03.Attributes(2).nodeName
                            field3Value = xn03.Attributes(2).nodeValue
                            
                            If (field1 = "BinCode" And field1Value = "1") Or (field2 = "BinCode" And field2Value = "1") Or (field3 = "BinCode" And field3Value = "1") Then
                            
                            '˵��Ϊ��Ʒ��
                                If field1 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field1Value)
                                
                                ElseIf field2 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field2Value)
                                
                                ElseIf field3 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field3Value)
                                
                                End If
                            End If
                            
                            If (field1 = "BinCode" And (field1Value = "3" Or field1Value = "4" Or field1Value = "5")) Or (field2 = "BinCode" And (field2Value = "3" Or field2Value = "4" Or field2Value = "5")) Or (field3 = "BinCode" And (field3Value = "3" Or field3Value = "4" Or field3Value = "5")) Then
                            '˵��Ϊ����Ʒ��
                               If field1 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field1Value)
                                
                                ElseIf field2 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field2Value)
                                
                                ElseIf field3 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field3Value)
                                
                                End If
                            End If
                            
                        ElseIf xn03.nodeName = "Data" Then

                            Exit For
                              
                        End If
                                  
                    Next   '  xn03 end
                    
              End If   'Device end
                    
             mapTemp.PassBinCount = goodDieQty
             mapTemp.FailBinCount = badDieQty
                            
            Next
            

        '�ϴ���DB
        mapTemp.FileName = fileNameTemp
        
        '2014-04-22 jiayun  ���Y��ͷ�ģ��滻lotid Ϊ�ļ�����
        
'        If UCase(Mid(fileNameTemp, 1, 2)) = "YP" Then
'
'        mapTemp.lotid = Replace(Replace(fileNameTemp, ".xml", ""), ".XML", "")
'
'        End If
        
  
        
        mapTemp.lotid = Replace(Replace(fileNameTemp, ".xml", ""), ".XML", "")
            

        
        
        
        Call AddMap(mapTemp, customerNameTemp)
      
    End If

NextRecord:
Next


End If


End Sub


Private Sub UpMxlForQtech(dirtemp As String)
'Qtech �Թ�Mapping ����

'--����XML

Dim XMLDoc As DOMDocument
Dim xn As IXMLDOMNode
Dim xn01 As IXMLDOMNode
Dim xn02 As IXMLDOMNode
Dim xn03 As IXMLDOMNode
Dim Flag As Integer
Dim JudgeFlag As Boolean

Dim goodDieQty As Integer
Dim badDieQty As Integer
                

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)


Set XMLDoc = New DOMDocument
XMLDoc.Load dirtemp

Set xn = XMLDoc.documentElement
'SumCount = 0

If Not xn Is Nothing Then
'ѭ�� Map

    For Each xn01 In xn.childNodes
        JudgeFlag = False
        goodDieQty = 0
        badDieQty = 0

        mapTemp.SubstrateId = xn01.Attributes(1).nodeValue
        
        '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
'          MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
          ErrorInf = ErrorInf + "," + mapTemp.SubstrateId

          GoTo NextRecord

       End If


        mapTemp.SubstrateType = xn01.Attributes(2).nodeValue

        'ѭ�� Device
        If xn01.nodeName = "Map" Then
            For Each xn02 In xn01.childNodes

                mapTemp.lotid = xn02.Attributes(1).nodeValue
                mapTemp.lotid = Replace$(mapTemp.lotid, ".", "")
                mapTemp.ProductId = xn02.Attributes(6).nodeValue
                mapTemp.CreateDate = xn02.Attributes(8).nodeValue
                mapTemp.MicronLotId = xn02.Attributes(14).nodeValue
                mapTemp.MicronLotId = Replace$(mapTemp.MicronLotId, ".", "")

                'ѭ�� ReferenceDevice
                If xn02.nodeName = "Device" Then
                    Flag = 0
                    For Each xn03 In xn02.childNodes
                        '������һ�еģ�������ʱ����
                        Dim field1 As String
                        Dim field2 As String
                        Dim field3 As String
                        Dim field1Value As String
                        Dim field2Value As String
                        Dim field3Value As String
                        
                        If xn03.nodeName = "Bin" Then
                            '2012-10-25 ����ֻ�������ؼ��� BinCode ,BinCount ,BinQuality
                            field1 = xn03.Attributes(0).nodeName
                            field1Value = xn03.Attributes(0).nodeValue
                            
                            field2 = xn03.Attributes(1).nodeName
                            field2Value = xn03.Attributes(1).nodeValue
                            
                            field3 = xn03.Attributes(2).nodeName
                            field3Value = xn03.Attributes(2).nodeValue
                            
                            If (field1 = "BinCode" And field1Value = "G") Or (field2 = "BinCode" And field2Value = "G") Or (field3 = "BinCode" And field3Value = "G") Then
                            
                            '˵��Ϊ��Ʒ��
                                If field1 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field1Value)
                                
                                ElseIf field2 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field2Value)
                                
                                ElseIf field3 = "BinCount" Then
                                goodDieQty = goodDieQty + CInt(field3Value)
                                
                                End If
                            End If
                            
                            If (field1 = "BinCode" And (field1Value = "X")) Or (field2 = "BinCode" And (field2Value = "X")) Or (field3 = "BinCode" And (field3Value = "X")) Then
                            '˵��Ϊ����Ʒ��
                               If field1 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field1Value)
                                
                                ElseIf field2 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field2Value)
                                
                                ElseIf field3 = "BinCount" Then
                                badDieQty = badDieQty + CInt(field3Value)
                                
                                End If
                            End If
                            
                        ElseIf xn03.nodeName = "Data" Then

                            Exit For
                              
                        End If
                                  
                    Next   '  xn03 end
                    
              End If   'Device end
                    
             mapTemp.PassBinCount = goodDieQty
             mapTemp.FailBinCount = badDieQty
                            
            Next
            

        '�ϴ���DB
        mapTemp.FileName = fileNameTemp
        Call AddMap(mapTemp, "QT")
        SumCount = SumCount + 1
    End If

NextRecord:
Next


End If


End Sub




Private Sub Command10_Click()
If TxtCustomerName.Text = "" Then
    MsgBox "��������ͻ����룡"
    Exit Sub
    
Else
 
 ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname ='" & Trim(TxtCustomerName.Text) & "' and qtech_created_date>sysdate-30  order by qtech_created_date desc , lotid, substrateid")
End If


End Sub

Private Sub Command11_Click()

Dim mapTemp As MapRecord

If TxtCustomerName.Text = "" Then
    MsgBox "��������ͻ����룡"
    Exit Sub
End If

If Text4.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String


    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text4.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets("sheet1")        '��sheet�еı�

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
           
        If j = 1 Then
            mapTemp.SubstrateId = Trim(tempVal) 'WaferId
            
                    '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
           If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
              MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
'              ErrorInf = ErrorInf + "," + mapTemp.SubstrateId
              
              GoTo NextRecord2
    
           End If
           
            
        End If
        
        If j = 2 Then
             mapTemp.lotid = Trim(tempVal) 'LotId
        End If
        
        If j = 3 Then
             mapTemp.ProductId = Trim(tempVal) 'ProductId
        End If
        
        If j = 4 Then
             mapTemp.PassBinCount = Trim(tempVal) 'PassBinCount
        End If
        
        If j = 5 Then
             mapTemp.FailBinCount = Trim(tempVal) 'FailBinCount
        End If
        
        
    Next j
    
    mapTemp.CreateDate = ""
    mapTemp.MicronLotId = ""
    mapTemp.FileName = ""
    
  Call AddMap2(mapTemp, Trim$(TxtCustomerName.Text))
SumCount = SumCount + 1
      

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

Private Sub Command12_Click()
'��ѡ���ļ�
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '�õ��ļ���
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Text4.Text = FName
    End If


End Sub

Private Sub Command13_Click()
On Error Resume Next
'si map

Dim FName
    '˧ѡ�ļ�
    ComSI.Filter = "map�ļ�(*.map)|*.map|txt�ļ�(*.txt)|*.txt"
    

    ComSI.ShowOpen
    '�õ��ļ���
    FName = ComSI.FileName
    If FName <> "" Then
       TxtSI.Text = Replace(FName, Chr(160), ",")
    End If
End Sub

Private Sub Command14_Click()
'si map


If CmbCustomer.Text = "" Then
 MsgBox "����ѡ��ͻ���"
 Exit Sub
End If


SumCount = 0
    ErrorInf = ""
    If TxtSI.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
    
    End If
    
    Dim FileName As String
    FileName = TxtSI.Text
    Dim dirtemp() As String
    Dim i As Integer
    
    If InStr(1, FileName, ",") > 0 Then
        dirtemp = Split(FileName, ",")
        
        For i = 1 To UBound(dirtemp)
             If CmbCustomer.Text = "GT" Or CmbCustomer.Text = "SI" Then
             
                UpMap (dirtemp(0) + "\" + dirtemp(i))
            
             ElseIf CmbCustomer.Text = "HD" Then
                  'HD�ͻ�
                 UpMapHD (dirtemp(0) + "\" + dirtemp(i))
                 
             ElseIf CmbCustomer.Text = "GC" Then
                  'HD�ͻ�
                 UpMapGCWlt (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.Text = "MG" Then
                  'MG�ͻ�
                  UpMapMG (dirtemp(0) + "\" + dirtemp(i))
                  
            ElseIf CmbCustomer.Text = "56" Then
             
                UpMap56 (dirtemp(0) + "\" + dirtemp(i))
            
            End If
            
        Next
        
    Else
       If CmbCustomer.Text = "GT" Or CmbCustomer.Text = "SI" Then
        
        UpMap (FileName)
        
       ElseIf CmbCustomer.Text = "HD" Then
          'HD�ͻ�
         UpMapHD (FileName)
         
       ElseIf CmbCustomer.Text = "GC" Then
          'GC�ͻ�   2015-03-20 jiayun add
         UpMapGCWlt (FileName)
         
         
        ElseIf CmbCustomer.Text = "MG" Then
         UpMapMG (FileName)
         
        ElseIf CmbCustomer.Text = "56" Then
        
        UpMap56 (FileName)
        
       End If
        
    End If
    
    
    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
    If ErrorInf <> "" Then
           MsgBox "�ϴ�ʧ�ܵ���:" + ErrorInf + "���ݿ����Ѵ��ڣ�"
    End If


End Sub

Private Sub UpMap(dirtemp As String)
Dim Flag As Integer
Dim JudgeFlag As Boolean
Dim customerNameTemp As String

Dim waferIDSeq As String
Dim allDieQty As Integer
Dim goodDieQty As Integer
Dim badDieQty As Integer

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.FileName = fileNameTemp
customerNameTemp = "GT"
 
'SI Mapping

Dim TextLine As String
Open dirtemp For Input As #1
' ���ļ���
Do While Not EOF(1)
' ѭ�����ļ�β��
Line Input #1, TextLine

    '�ж����У��Ƿ�Ҫȡ���ϣ�������������һ��
    If InStr(TextLine, "LOT_NO") > 0 Then
    'lotid
     mapTemp.lotid = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
     mapTemp.SubstrateId = mapTemp.lotid & waferIDSeq
     
     
    End If
    
    If InStr(TextLine, "GOOD_DIE") > 0 Then
    'qty
     mapTemp.PassBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
     allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
     
     mapTemp.FailBinCount = allDieQty - mapTemp.PassBinCount
    
    End If


    If InStr(TextLine, "[FLAT") > 0 Then
      GoTo ContinueFlag
    
    End If



Loop


ContinueFlag:


Close #1    ' �ر��ļ���

       ' �ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       
       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
            MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
       
       Else
       
            Call AddMap(mapTemp, customerNameTemp)

       End If

End Sub



Private Sub UpMap56(dirtemp As String)
Dim Flag As Integer
Dim JudgeFlag As Boolean
Dim customerNameTemp As String
Dim productaNameTenp As String

Dim waferIDSeq As String
Dim allDieQty As Long
Dim goodDieQty As Long
Dim badDieQty As Long

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.FileName = fileNameTemp
customerNameTemp = "56"
 
'56 Mapping

Dim TextLine As String
Open dirtemp For Input As #1
' ���ļ���
Do While Not EOF(1)
' ѭ�����ļ�β��
Line Input #1, TextLine

    '�ж����У��Ƿ�Ҫȡ���ϣ�������������һ��
    If InStr(TextLine, "Product Name") > 0 Then
    
    mapTemp.SubstrateType = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
    
    End If
    
    
     If InStr(TextLine, "Lot id") > 0 Then
    mapTemp.lotid = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
    End If
    
    
     If InStr(TextLine, "Wafer ID") > 0 Then
     waferIDSeq = Right("0" & Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20)), 2)
     mapTemp.SubstrateId = mapTemp.lotid & waferIDSeq
    End If
    
     If InStr(TextLine, "Gross Dice") > 0 Then
    'qty
     allDieQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
     End If
    
     If InStr(TextLine, "Good Dice") > 0 Then
    'qty
     mapTemp.PassBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
     mapTemp.FailBinCount = CLng(allDieQty) - mapTemp.PassBinCount
    
    End If
    

    If InStr(TextLine, "Yield") > 0 Then
      GoTo ContinueFlag
    
    End If



Loop


ContinueFlag:


Close #1    ' �ر��ļ���

       ' �ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       
       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
            MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
       
       Else
       
            Call AddMap(mapTemp, customerNameTemp)

       End If

End Sub




'2015-04-20 jiayun add MG

Private Sub UpMapMG(dirtemp As String)
Dim Flag As Integer
Dim JudgeFlag As Boolean
Dim customerNameTemp As String

Dim waferIDSeq As String
Dim allDieQty As Integer
Dim goodDieQty As Integer
Dim badDieQty As Integer

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.FileName = fileNameTemp
customerNameTemp = "MG"
 
'SI Mapping

Dim TextLine As String
Open dirtemp For Input As #1
' ���ļ���
Do While Not EOF(1)
' ѭ�����ļ�β��
Line Input #1, TextLine

    '�ж����У��Ƿ�Ҫȡ���ϣ�������������һ��
    If InStr(TextLine, "LOT_NO") > 0 Then
    'lotid
     mapTemp.lotid = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, 3))
     mapTemp.SubstrateId = mapTemp.lotid & waferIDSeq
     
     
    End If
    
    If InStr(TextLine, "GOOD_DIE") > 0 Then
    'qty
     mapTemp.PassBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
     allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
     
     mapTemp.FailBinCount = allDieQty - mapTemp.PassBinCount
    
    End If


    If InStr(TextLine, "TEST_TIME") > 0 Then
      GoTo ContinueFlag
    
    End If



Loop


ContinueFlag:


Close #1    ' �ر��ļ���

       ' �ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       
'       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
'            MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
'
'       Else
       
            'Call AddMap(mapTemp, customerNameTemp)
            
            Call updateMGMap(mapTemp.SubstrateId, mapTemp.PassBinCount, mapTemp.FailBinCount)
            

'       End If

End Sub



Private Sub UpMapHD(dirtemp As String)
Dim Flag As Integer
Dim JudgeFlag As Boolean
Dim customerNameTemp As String
Dim waferIdTemp As String


Dim waferIDSeq As String
Dim allDieQty As Integer
Dim goodDieQty As Integer
Dim badDieQty As Integer

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.FileName = fileNameTemp
customerNameTemp = "HD"
 
'SI Mapping

Dim TextLine As String
Open dirtemp For Input As #1
' ���ļ���
Do While Not EOF(1)
' ѭ�����ļ�β��
Line Input #1, TextLine

    '�ж����У��Ƿ�Ҫȡ���ϣ�������������һ��
    'LotID
    If InStr(TextLine, "Lot No") > 0 Then
    'lotid
     mapTemp.lotid = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
'     waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
'     mapTemp.SubstrateId = mapTemp.lotID & waferIDSeq
     
     
    End If
    
   'WaferID
  If InStr(TextLine, "Wafer ID") > 0 Then
    'lotid
    ' mapTemp.lotID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
     'D02939-1
     waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     waferIdTemp = Mid(waferIdTemp, InStr(waferIdTemp, "-") + 1, 2)
     
     waferIDSeq = Right("0" & waferIdTemp, 2)
     mapTemp.SubstrateId = mapTemp.lotid & waferIDSeq
     
    End If
    
    
    If InStr(TextLine, "Total Pass") > 0 Then
    'qty
     mapTemp.PassBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
'     allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
'
'     mapTemp.FailBinCount = allDieQty - mapTemp.PassBinCount
    
    End If
    
     If InStr(TextLine, "Total Fail") > 0 Then
    'qty
     mapTemp.FailBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
    'allDieQty = mapTemp.PassBinCount + mapTemp.FailBinCount
'
'     mapTemp.FailBinCount = allDieQty - mapTemp.PassBinCount
    
    End If


    If InStr(TextLine, "Yield") > 0 Then
      GoTo ContinueFlag
    
    End If



Loop


ContinueFlag:


Close #1    ' �ر��ļ���

       ' �ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
       
       If (JudgeFlagStautsMapping2(mapTemp.SubstrateId)) Then
            MsgBox "��ʣ�" & mapTemp.SubstrateId & "�Ѵ��ڣ������ϴ�!"
       
       Else
       
            Call AddTSVMap(mapTemp, customerNameTemp)

       End If

End Sub

Private Sub UpMapGCWlt(dirtemp As String)


Dim customerNameTemp As String


Dim waferidGCTemp As String
Dim gcGoodDieQty As Long

Dim fileNameTemp As String
fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.FileName = fileNameTemp
customerNameTemp = "GC"
 
'SI Mapping

Dim TextLine As String
Open dirtemp For Input As #1
' ���ļ���
Do While Not EOF(1)
' ѭ�����ļ�β��
Line Input #1, TextLine

    '�ж����У��Ƿ�Ҫȡ���ϣ�������������һ��
    'LotID
    If InStr(TextLine, "Lot:") > 0 Then

     waferidGCTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
    End If
    
  
    If InStr(TextLine, "BIN_1") > 0 Then
    
      gcGoodDieQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
      
      GoTo ContinueFlag
    
    End If



Loop


ContinueFlag:


Close #1    ' �ر��ļ���

    
       
Call updateGCWltMap(waferidGCTemp, gcGoodDieQty)

    

End Sub




Private Sub Command15_Click()
   ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname  in ('SI','GT')  and qtech_created_date>sysdate-30  order by qtech_created_date desc , lotid, substrateid")
End Sub

Private Sub Command2_Click()
'
'On Error Resume Next
'Dim FName
'    '˧ѡ�ļ�
'    CommonDialog1.Filter = "CSV�ļ�(*.csv)|*.csv"
'    CommonDialog1.ShowOpen
'    '�õ��ļ���
'    FName = CommonDialog1.FileName
'    If FName <> "" Then
'       Text2.Text = FName
'    End If

Unload Me

End Sub

Private Sub Command3_Click()
Dim source_batch_id_Temp As String
'�ϴ�OI��CSV
'�����ļ���
If Text2.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text2.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
    End If
    

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

'con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText

'2012-07-03 jiayunzhang �޸Ķ�CSV�ķ�ʽ

  '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text2.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�


  '�ж������Excel�еĺ��趨���Ƿ���ͬ
  '2012-10-08 jiayunzhang �г���Ҫ������һ�� comp_code

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 73 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If







Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim WV_inspect As String
Dim Comp_codeTemp As String



Dim SumCount As Integer
SumCount = 0
'Rs.MoveFirst
'For i = 0 To Rs.RecordCount - 1

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count


    temp = ""
    source_batch_id_Temp = ""
'    For j = 0 To Rs.fields.Count - 1

'2012-07-03 ��ͻ�OI����ֶΣ����ݿ����������һ�У����Գ���Ҫ���⴦�� ��������xlSheet.Range("A1").CurrentRegion.Columns.Count ��Ϊ 71
      For j = 1 To 71
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)
            End If

      
'        strChar = Chr(96 + j)
        
        
        
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
    
        If j = 46 Or j = 60 Then
            temp = temp & "," & newStrDate("" & tempVal)
        
        Else
            If j = 61 Then
            tempVal = Format(xlSheet.Range(strChar & i).Value, "HH:mm:SS")
            temp = temp & "," & newStr("" & tempVal)
            Else
            
            temp = temp & "," & newStr("" & tempVal)
            End If
        
        End If
        
        If j = 3 Then
            source_batch_id_Temp = tempVal
        End If
    
    Next j
    
    j = 72
    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
    
    WV_inspect = newStr("" & xlSheet.Range(strChar & i).Value)
    
    j = 73
    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
    
    Comp_codeTemp = newStr("" & xlSheet.Range(strChar & i).Value)
    
    
    
    
    'ȡĿǰDB����ID��
    id = GetMaxID()
    temp = id & temp
    temp2 = temp & ",'Y','" & gUserName & "',GETDATE(),'','','AA',0," & WV_inspect & "," & Comp_codeTemp
    temp = temp & ",'Y','" & gUserName & "',sysdate,'','','AA',0,1," & WV_inspect & "," & Comp_codeTemp
    
'    Debug.Print temp

'             '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagStautsOI(source_batch_id_Temp)) Then
       MsgBox "��ʣ�" & source_batch_id_Temp & "�Ѵ��ڣ������ϴ�!"
       GoTo NextRecord2

    End If

    
    Call AddOI(temp, temp2)
     SumCount = SumCount + 1
    
    '�ϴ���DB
    
NextRecord2:
'    Rs.MoveNext

Next i


If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
End If


End Sub

Private Function newStrDate(temp As String)
Dim mmTemp As String
Dim ddTemp As String
Dim newTemp As String
'2012-09-14 jiayunzhang Modify ʱ���ʽ����ת����
If temp <> "" Then

'    mmTemp = Mid$(temp, 6, InStr(6, temp, "-") - 6)
'    ddTemp = Right$(temp, Len(temp) - InStr(6, temp, "-"))
    
'    If Val(mmTemp) >= 1 And Val(mmTemp) <= 12 And Val(ddTemp) >= 1 And Val(ddTemp) <= 12 Then
'        '��ʱ��Ҫת��
'
'        newTemp = Left$(temp, 4) & "-" & ddTemp & "-" & mmTemp
'        newStrDate = "'" & newTemp & "'"
'
'    Else
        newStrDate = "'" & temp & "'"
'    End If

Else
newStrDate = "''"

End If

End Function

Private Function newStr(temp As String)
If temp <> "" Then
newStr = "'" & temp & "'"
Else
newStr = "''"

End If

End Function


Private Sub Command4_Click()
    ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname ='AA' and qtech_created_date>sysdate-90  order by qtech_created_date desc , lotid, substrateid")
End Sub

Private Sub Command5_Click()

'    ExporToExcel (" select ID,PO_NUM,PO_ITEM,SOURCE_BATCH_ID,SOURCE_MTRL_NUM,MTRL_NUM,MTRL_DESC,TEST_MTRL_NUM,TEST_MTRL_DESC, MPN, " & _
'                 " MPN_DESC, SOURCE_MTRL_SLOC, MTRL_NUM_MTRLGRP,PROBE_SHIP_PART_TYPE, OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY, CURRENT_WAFER_QTY, DIE_QTY, DESIGN_ID,COUNTRY_OF_FAB," & _
'                 " FAB_CONV_ID,FAB_EXCR_ID,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,WAFER_SIZE,IMAGER_CUSTOMER_REV, CHROMATICITY, MICRO_LENS_SHIFT, TEMPERATURE_SPEC," & _
'                 " PRB_CONTAINMENT_TYPE, FABRICATION_FACILITY, PRB_EXCR_ID, BATCH_COMMENT_PROBE, ASSY_PROCESS_ID, DARK_BOND_PAD_ASSY, ASSY_SERIAL_TYPE, STICKY_BACKS_TO_SAVE, OPTICAL_QUALITY, ENCODED_MARK_ID, " & _
'                 " PLANNED_LASER_SCRIBE, PACKAGE_LID_TYPE, PACKAGE_TYPE, PB_FREE_PACKAGE, TARGET_WAF_THICKNESS, RELIABILITY_SAMPLING, LOT_PRIORITY, WAFER_BOX_TYPE, TEST_SITE,ASSEMBLY_FACILITY, " & _
'                 " BATCH_COMMENT_ASSY, TST_PROCESS_ID,ELEC_SPECIAL_TEST, BOX_TYPE, PROTECTIVE_FILM_APLD, SHIPPING_MST_260,SHIPPING_MST_LEVEL, T_PRICE, SHIP_COMMENT, BATCH_COMMENT_TEST, " & _
'                 " CREATED_DATE, CREATED_TIME, UNIT_PRICE,REF_PO, REF_PO_ITEM, COUNTRY_OF_ASSEMBLY, MICRON_MATERIAL,DATE_CODE, SHIP_SITE, SPECIAL_PROCESS_LOT, " & _
'                 " LOT_STATUS, CUSTOM_PART_NO, FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE, QTECH_LASTUPDATE_BY, QTECH_LASTUPDATE_DATE from CustomerOItbl_test  where (customershortname = 'AA' or customershortname is null)  and (source_batch_id like '6%' or source_batch_id like '7%')  order by id ")
'
    
   '2012-05-15 jiayunzhang Modify
    
    ExporToExcel (" select ID,PO_NUM,PO_ITEM,SOURCE_BATCH_ID,SOURCE_MTRL_NUM,MTRL_NUM,MTRL_DESC,TEST_MTRL_NUM,TEST_MTRL_DESC, MPN, " & _
                 " MPN_DESC, SOURCE_MTRL_SLOC, MTRL_NUM_MTRLGRP,PROBE_SHIP_PART_TYPE, OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY, CURRENT_WAFER_QTY, DIE_QTY, DESIGN_ID,COUNTRY_OF_FAB," & _
                 " FAB_CONV_ID,FAB_EXCR_ID,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,WAFER_SIZE,IMAGER_CUSTOMER_REV, CHROMATICITY, MICRO_LENS_SHIFT, TEMPERATURE_SPEC," & _
                 " PRB_CONTAINMENT_TYPE, FABRICATION_FACILITY, PRB_EXCR_ID, BATCH_COMMENT_PROBE, ASSY_PROCESS_ID, DARK_BOND_PAD_ASSY, ASSY_SERIAL_TYPE, STICKY_BACKS_TO_SAVE, OPTICAL_QUALITY, ENCODED_MARK_ID, " & _
                 " PLANNED_LASER_SCRIBE, PACKAGE_LID_TYPE, PACKAGE_TYPE, PB_FREE_PACKAGE, TARGET_WAF_THICKNESS, RELIABILITY_SAMPLING, LOT_PRIORITY, WAFER_BOX_TYPE, TEST_SITE,ASSEMBLY_FACILITY, " & _
                 " BATCH_COMMENT_ASSY, TST_PROCESS_ID,ELEC_SPECIAL_TEST, BOX_TYPE, PROTECTIVE_FILM_APLD, SHIPPING_MST_260,SHIPPING_MST_LEVEL, T_PRICE, SHIP_COMMENT, BATCH_COMMENT_TEST, " & _
                 " CREATED_DATE, CREATED_TIME, UNIT_PRICE,REF_PO, REF_PO_ITEM, COUNTRY_OF_ASSEMBLY, MICRON_MATERIAL,DATE_CODE, SHIP_SITE, SPECIAL_PROCESS_LOT, " & _
                 " LOT_STATUS, CUSTOM_PART_NO, wafer_visual_inspect, comp_code,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE, QTECH_LASTUPDATE_BY, QTECH_LASTUPDATE_DATE from CustomerOItbl_test  where (customershortname = 'AA' or customershortname is null)   order by id desc ")
    
    
    
    
End Sub

Private Sub Command6_Click()
'GC
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog2.Filter = "CSV�ļ�(*.csv)|*.csv|EXCEL�ļ�(*.xlsx)|*.xlsx|EXCEL�ļ�(*.xls)|*.xls"
    
    CommonDialog2.ShowOpen
    '�õ��ļ���
    FName = CommonDialog2.FileName
    If FName <> "" Then
       Text3.Text = FName
    End If

End Sub

Private Sub UploadGC()
'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String

customerTemp = "GC"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text3.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    End If

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
        Rs.open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        Dim str01 As String
        Dim str03 As String
        SumCount = 0
        Rs.MoveFirst
        
        GCHeaderFlag = False
        
        For i = 0 To Rs.RecordCount - 1
            temp = ""
            id = 0
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            gcHeaderTemp.Supplier = Rs.fields(2).Value
            gcHeaderTemp.ShipTo = Rs.fields(3).Value
            gcHeaderTemp.FAB_Device = Rs.fields(4).Value
            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
            gcHeaderTemp.GC_Version = Rs.fields(6).Value
            gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
            
   
            str01 = Rs.fields(8).Value
            
            If InStr(str01, "��") > 0 Then
            
            str03 = Replace(str01, "��", "-")
            str03 = Replace(str03, "��", "")
            str03 = Year(Date) & "-" & str03
            gcHeaderTemp.GC_Date = str03
            
            Else
            
            gcHeaderTemp.GC_Date = Rs.fields(8).Value
            
            End If
            
            gcHeaderTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Wafer_id = Rs.fields(10).Value
            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(11).Value)
            gcHeaderTemp.WO_NO = Rs.fields(12).Value
            gcHeaderTemp.Ship_Out = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value)
            
            '2015-02-03 jiayunadd check shipOut
            '�����COG�ģ��򲻿���Ϊ��
            
            If Left(gcHeaderTemp.Lot_ID, 3) = "GXS" Then
                If gcHeaderTemp.Ship_Out = "" Then
                  MsgBox "GC COG�����һ�з�����ַ�������пգ�"
                  Exit Sub
                
                End If
            
                
            End If
            
            
            
            '2013-12-05 jiayun add
            '�ж�wo�Ƿ�Ϊ��
            
            If Trim(gcHeaderTemp.WO_NO) = "" Then
            
                MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
                Exit Sub

            End If
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
                    Exit Sub
            End If
            
            
            '2012-11-05 jiayun �޸� GC
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
                '��lotid�и���ʱ�����ѯ�ϴε�id
                
                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                
                
                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
                
            End If
           
            
            Rs.MoveNext
        
        Next i
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub

Private Function GetGCWLT(txtTemp As String) As String
        GetGCWLT = "F"
        
        Dim CusDevice As String
        Dim GCVersion As String
        
        Dim FName As String
        Dim Nextline As String
        FName = Trim(Text3.Text)
        Open FName For Input As #2
        
        Do Until EOF(2)
        Line Input #2, Nextline
        
             If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then
             
                Dim bid
                bid = Split(Nextline, ",")
                
                CusDevice = bid(5)
                GCVersion = bid(6)
                
                '�ж��ǲ���WLT
                
                If CusDevice = "GC0312-3" And Right(GCVersion, 1) = "C" Then
                GetGCWLT = "T"
            
                Else
                GetGCWLT = "F"
                End If
                Close #2
              Exit Function
             End If
        
        Loop
        Close #2
        
End Function

Private Sub UploadGCNew()
 Dim SumCount As Integer
 Dim userNameTemp As String
 Dim poidTemp As String
 Dim gcdeviceTemp As String
 Dim lotIDTemp As String
 Dim waferIdTemp As String
 Dim woNoTemp As String
 
 SumCount = 0

If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If

        Dim k As Integer
        
        Dim FName As String
        Dim Nextline As String
        FName = Trim(Text3.Text)
        Open FName For Input As #1
        
        Do Until EOF(1)
        Line Input #1, Nextline
              cusPTTemp = ""
              gcVerTemp = ""
              gcVerLastTemp = ""
              
             If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then
             Dim bid
             bid = Split(Nextline, ",")
             
        
            '��ֵ
            userNameTemp = gUserName
            poidTemp = bid(1)
            gcdeviceTemp = bid(5)
            lotIDTemp = bid(9)
            waferIdTemp = lotIDTemp & Right("0" & bid(10), 2)
            woNoTemp = bid(12)
            
             
            If (JudgeGCLableWlaID(waferIdTemp)) Then
               MsgBox "GC ��ʣ�" & waferIdTemp & " �Ѵ��ڣ��������ϴ�!"
               
            Else
     
                Call AddGCLableWLAWaferid(userNameTemp, poidTemp, gcdeviceTemp, lotIDTemp, waferIdTemp, woNoTemp)
    
                SumCount = SumCount + 1
                
            End If
                    
          
           
        End If
        
        Loop
        Close #1
        
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub

Private Sub UploadGCNewWLDT()
'2015-04-28 jiayun add WLDT

'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim cusPTTemp As String
Dim gcVerTemp As String
Dim gcVerLastTemp As String
Dim waferIdTemp As String

Dim wo_HT_Temp As String


wo_HT_Temp = "WONO_" & Replace(Replace(Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", ""), ":", ""), " ", "")

customerTemp = "GC"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
'Dim dirName As String
'Dim FileName As String

''��ȡ�ļ���
'    If InStrRev(Trim(Text3.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
'        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
'    End If


'�ж� GC���ͣ��ǲ���
'If GetGCWLT(Trim(Text3.Text)) = "T" Then
'UploadGCWLTNew
'
'Exit Sub
'End If


        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        Dim str01 As String
        Dim str03 As String
        SumCount = 0
 
        
        GCHeaderFlag = False
        
        

        Dim k As Integer
        
        Dim FName As String
        Dim Nextline As String
        FName = Trim(Text3.Text)
        Open FName For Input As #1
        
        Do Until EOF(1)
        Line Input #1, Nextline
              cusPTTemp = ""
              gcVerTemp = ""
              gcVerLastTemp = ""
              waferIdTemp = ""
              
             If UCase(Left(Trim(Nextline), 2)) <> "NO" Then
             Dim bid
             bid = Split(Nextline, ",")
             
            id = 0
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = bid(0)
            gcHeaderTemp.PO_NO = bid(6)
            gcHeaderTemp.Supplier = bid(1)
            gcHeaderTemp.ShipTo = bid(2)
            gcHeaderTemp.FAB_Device = bid(3)
            
            gcHeaderTemp.Customer_Device = bid(4) & "-3"
            cusPTTemp = Trim(gcHeaderTemp.Customer_Device)
            gcHeaderTemp.GC_Version = bid(5)
            gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
            
            '2015-04-27 jiayun add ����λϵͳ�Զ���
'            gcVerLastTemp = GetGCVerLastChar(cusPTTemp)
'
'            If gcVerLastTemp <> "" Then
'                 gcHeaderTemp.GC_Version = gcVerTemp & gcVerLastTemp
'
'            Else
'
'                If cusPTTemp = "GC1004-3" Then
'
'                      If Mid(gcVerTemp, 1, 1) = "A" Or Mid(gcVerTemp, 1, 1) = "B" Or Mid(gcVerTemp, 1, 1) = "C" Or Mid(gcVerTemp, 1, 1) = "D" Then
'                       gcHeaderTemp.GC_Version = gcVerTemp & "A"
'                      Else
'                       gcHeaderTemp.GC_Version = gcVerTemp & "B"
'                      End If
'
'
'                ElseIf cusPTTemp = "GC0329-3" Then
'                         If Len(gcVerTemp) = 2 Then
'                            gcHeaderTemp.GC_Version = gcVerTemp & "D"
'
'                         ElseIf Len(gcVerTemp) = 3 Then
'                             gcHeaderTemp.GC_Version = gcVerTemp
'
'                         Else
'                            MsgBox "GC WO�У�GCVersion�����ݲ��ԣ���ȷ��Wo!"
'                            Exit Sub
'
'                         End If
'
'
'
'                Else
'                    '�жϳ����Ƿ�Ϊ3 ������ǣ����г��������ϴ�����������ʾ����
'                    If Len(gcVerTemp) = 3 Then
'                         gcHeaderTemp.GC_Version = gcVerTemp
'
'                    Else
'                            MsgBox "GC WO�У�GCVersion�����ݲ��ԣ���ȷ��Wo!"
'                            Exit Sub
'
'                    End If
'
'
'
'
'                End If
'
'
'
'            End If
            
            
            waferIdTemp = bid(10) & Right("0" & bid(11), 2)
            
            
            gcDetailTemp.Marking_Lot_ID = GetGCWLDMaringCode(waferIdTemp)
            
   
            str01 = bid(9)
            
            If InStr(str01, "��") > 0 Then
            
            str03 = Replace(str01, "��", "-")
            str03 = Replace(str03, "��", "")
            str03 = Year(Date) & "-" & str03
            gcHeaderTemp.GC_Date = str03
            
            Else
            
            gcHeaderTemp.GC_Date = bid(8)
            
            End If
            
            gcHeaderTemp.Lot_ID = bid(10)
            gcDetailTemp.Lot_ID = bid(10)
            gcDetailTemp.Wafer_id = bid(11)
            gcDetailTemp.Good_Die_Qty = CInt(bid(12))
            gcHeaderTemp.WO_NO = wo_HT_Temp
            gcHeaderTemp.Ship_Out = bid(16)
            
            '2015-02-03 jiayunadd check shipOut
            '�����COG�ģ��򲻿���Ϊ��
            
            If Left(gcHeaderTemp.Lot_ID, 3) = "GXS" Then
                If gcHeaderTemp.Ship_Out = "" Then
                  MsgBox "GC COG�����һ�з�����ַ�������пգ�"
                  Exit Sub
                
                End If
            
                
            End If
            
            
            
            '2013-12-05 jiayun add
            '�ж�wo�Ƿ�Ϊ��
            
            If Trim(gcHeaderTemp.WO_NO) = "" Then
            
                MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
                Exit Sub

            End If
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
                    Exit Sub
            End If
            
            
            '2012-11-05 jiayun �޸� GC
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
                '��lotid�и���ʱ�����ѯ�ϴε�id
                
                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                
                
                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_id), 2)
            
            
            If (JudgeGCDetailIdWLD(gcDetailTemp.Lot_ID, gcDetailTemp.item)) Then
               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.item & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   'gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                   
                   'gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                   
                   
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
                
            End If
           
            
 
        End If
        
        Loop
        Close #1
        
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub


Private Sub UploadGCWLTNew()
'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim wo_HT_Temp As String


wo_HT_Temp = "WONO_" & Replace(Replace(Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", ""), ":", ""), " ", "")

customerTemp = "GC"

'�ϴ�OI��CSV
'�����ļ���
'If Text3.Text = "" Then
'    MsgBox "��ѡ����ϴ����ļ�"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String

''��ȡ�ļ���
'    If InStrRev(Trim(Text3.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
'        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
'    End If

        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        Dim str01 As String
        Dim str03 As String
        SumCount = 0
 
        
        GCHeaderFlag = False
        
        

        Dim k As Integer
        
        Dim FName As String
        Dim Nextline As String
        FName = Trim(Text3.Text)
        Open FName For Input As #3
        
        Do Until EOF(3)
        Line Input #3, Nextline
        
             If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then
             Dim bid
             bid = Split(Nextline, ",")
             
            id = 0
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = bid(0)
            gcHeaderTemp.PO_NO = bid(1)
            gcHeaderTemp.Supplier = bid(2)
            gcHeaderTemp.ShipTo = bid(3)
            gcHeaderTemp.FAB_Device = bid(4)
            
            gcHeaderTemp.Customer_Device = bid(5)
            gcHeaderTemp.GC_Version = bid(6)
            gcDetailTemp.Marking_Lot_ID = bid(7)
            
   
            str01 = bid(8)
            
            If InStr(str01, "��") > 0 Then
            
            str03 = Replace(str01, "��", "-")
            str03 = Replace(str03, "��", "")
            str03 = Year(Date) & "-" & str03
            gcHeaderTemp.GC_Date = str03
            
            Else
            
            gcHeaderTemp.GC_Date = bid(8)
            
            End If
            
            gcHeaderTemp.Lot_ID = bid(9)
            gcDetailTemp.Lot_ID = bid(9)
            gcDetailTemp.Wafer_id = bid(10)
            gcDetailTemp.Good_Die_Qty = CInt(bid(11))
            gcDetailTemp.Remark = "WLT"
            gcHeaderTemp.WO_NO = wo_HT_Temp
            gcHeaderTemp.Ship_Out = bid(13)
            
           
        
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
                    Exit Sub
            End If
            
            
            '2012-11-05 jiayun �޸� GC
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
                '��lotid�и���ʱ�����ѯ�ϴε�id
                
                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                
                
                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
'            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
'               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "�Ѵ��ڣ������ϴ�!"
'
'            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                    Call AddGCWLTDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
             
'            End If
           
            
 
        End If
        
        Loop
        Close #3
        
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub




Private Sub UploadEQ()
'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String

customerTemp = "EQ"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text3.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    End If

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
        Rs.open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        Dim str01 As String
        Dim str03 As String
        SumCount = 0
        Rs.MoveFirst
        
        GCHeaderFlag = False
        
        For i = 0 To Rs.RecordCount - 1
            temp = ""
            id = 0
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            gcHeaderTemp.Supplier = Rs.fields(2).Value
            gcHeaderTemp.ShipTo = Rs.fields(3).Value
            gcHeaderTemp.FAB_Device2 = IIf(IsNull(Rs.fields(4).Value), "", Rs.fields(4).Value)
            
            gcHeaderTemp.FAB_Device = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
                   
            
            gcHeaderTemp.Customer_Device = IIf(IsNull(Rs.fields(5).Value), "", Rs.fields(5).Value)
            gcHeaderTemp.GC_Version = IIf(IsNull(Rs.fields(6).Value), "", Rs.fields(6).Value)
            'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
            gcHeaderTemp.GC_Date = Rs.fields(7).Value
            
            
            gcHeaderTemp.Lot_ID = Rs.fields(8).Value
            gcDetailTemp.Lot_ID = Rs.fields(8).Value
            gcDetailTemp.Wafer_id = Rs.fields(9).Value
            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(10).Value)
            gcHeaderTemp.WO_NO = IIf(IsNull(Rs.fields(11).Value), "", Rs.fields(11).Value)
            gcHeaderTemp.remarkTemp = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
            gcHeaderTemp.Date_Code = IIf(IsNull(Rs.fields(13).Value), "", Rs.fields(13).Value)
            gcHeaderTemp.Marking_Lot_ID1 = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value)
            gcHeaderTemp.Marking_Lot_ID2 = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
            gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value) & " " & IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)

            
            
            '2013-12-05 jiayun add
            '�ж�wo�Ƿ�Ϊ��
            
           ' If Trim(gcHeaderTemp.WO_NO) = "" Then
            
               ' MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
               ' Exit Sub

          '  End If
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
'            If gcDetailTemp.Good_Die_Qty <= 0 Then
'                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
'                    Exit Sub
'            End If
            
            
            '2012-11-05 jiayun �޸� GC
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeEQHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO, gcHeaderTemp.PO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
                '��lotid�и���ʱ�����ѯ�ϴε�id
                
               id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                
                
                    Call AddEQHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
                
            End If
           
            
            Rs.MoveNext
        
        Next i
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub



Private Sub UploadEQ_IS()

Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "EQ"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 30 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
       ' strChar = Chr(96 + j)
        
        
        If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
        Else
                strChar = Chr(96 + j)
        End If
             
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            eqISHeaderTemp.Created_By = gUserName
            If j = 1 Then
                eqISHeaderTemp.Created_Datetime = Trim(tempVal)
            End If
            
            If j = 2 Then
                eqISHeaderTemp.Vendor = Trim(tempVal)
            End If
            
            If j = 3 Then
                eqISHeaderTemp.Process = Trim(tempVal)
            End If
            
            If j = 4 Then
                eqISHeaderTemp.OrderType = Trim(tempVal)
            End If
            
            If j = 5 Then
                 eqISHeaderTemp.ESR_No = Trim(tempVal)
            End If
            '------
            If j = 6 Then
                eqISHeaderTemp.AssemblyDateCode = Trim(tempVal)
            End If
            
            If j = 7 Then
                 eqISHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                eqISHeaderTemp.WO_NO = Trim(tempVal)
             
            End If
            
            If j = 9 Then
                eqISHeaderTemp.WorkOrder_PartNo = Trim(tempVal)
            End If
            
             If j = 10 Then
                eqISHeaderTemp.Device = Trim(tempVal)
                
            End If
            '--------
            If j = 11 Then
                eqISHeaderTemp.WaferQty = Trim(tempVal)
            End If
            
            If j = 12 Then
                eqISHeaderTemp.AssyQty = Trim(tempVal)
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
                
            End If
            
            If j = 13 Then
                eqISHeaderTemp.Package = Trim(tempVal)
            End If
            
            If j = 14 Then
                eqISHeaderTemp.FabLotNo = Trim(tempVal)
            End If
            
            If j = 15 Then
                eqISHeaderTemp.TSM_A = Trim(tempVal)
            End If
            '--------------------
              If j = 16 Then
                eqISHeaderTemp.TSM_B = Trim(tempVal)
            End If
            
            If j = 17 Then
                eqISHeaderTemp.TSM_C = Trim(tempVal)
            End If
            
            If j = 18 Then
                eqISHeaderTemp.TSM_D = Trim(tempVal)
            End If
            
            If j = 19 Then
                eqISHeaderTemp.BondingDiagram = Trim(tempVal)
            End If
            
            If j = 20 Then
                eqISHeaderTemp.CompleteLotno = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
                
            End If
            
            
            '----------------------
            
            If j = 21 Then
                eqISHeaderTemp.Remarks = Trim(tempVal)
            End If
            If j = 22 Then
                eqISHeaderTemp.MarketingPartNumber = Trim(tempVal)
            End If
            If j = 23 Then
                eqISHeaderTemp.SPA = Trim(tempVal)
            End If
            If j = 24 Then
                eqISHeaderTemp.DateCode = Trim(tempVal)
            End If
            If j = 25 Then
                eqISHeaderTemp.DieID = Trim(tempVal)
            End If
            
            '---------------------
            
              
            If j = 26 Then
                eqISHeaderTemp.LabelFormat = Trim(tempVal)
            End If
            If j = 27 Then
                eqISHeaderTemp.waferid = Trim(tempVal)
                gcDetailTemp.Wafer_id = Trim(tempVal)
                  
            End If
            If j = 28 Then
                eqISHeaderTemp.SPADESC = Trim(tempVal)
            End If
            If j = 29 Then
                eqISHeaderTemp.Attention = Trim(tempVal)
            End If
            If j = 30 Then
                eqISHeaderTemp.CompanyName = Trim(tempVal)
            End If
               
            
        
    Next j
    
    
    
    
    
     If (JudgeEQISHeaderId(eqISHeaderTemp.PO_NO, eqISHeaderTemp.WO_NO, eqISHeaderTemp.CompleteLotno)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetEQISLotIDPOId(eqISHeaderTemp.CompleteLotno, eqISHeaderTemp.PO_NO)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddEQISHeader(eqISHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
'    '�ж�lotID��Detail�����Ƿ��Ѵ���
'
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "EQ ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"

    Else
'    '�ϴ���Detail����

    gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)

    Call AddEQDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1

    End If
    
    ' ��ϸ���´��ٸ�------------------------

    
    
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        




'------------------
'��ȡCSV
'Dim source_batch_id_Temp As String
'Dim customerTemp As String
'
'customerTemp = "EQ"
'
''�ϴ�OI��CSV
''�����ļ���
'If Text3.Text = "" Then
'    MsgBox "��ѡ����ϴ����ļ�"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''��ȡ�ļ���
'    If InStrRev(Trim(Text3.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
'        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
'    End If
'
'Dim con As New ADODB.Connection
'Dim Rs As New ADODB.Recordset
'
'
'        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'        Rs.open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
'
'        Dim i As Integer
'        Dim j As Integer
'        Dim id As Long
'        Dim temp As String
'        Dim SumCount As Integer
'        Dim GCHeaderFlag As Boolean
'        Dim str01 As String
'        Dim str03 As String
'        SumCount = 0
'        Rs.MoveFirst
'
'        GCHeaderFlag = False
'
'        For i = 0 To Rs.RecordCount - 1
'            temp = ""
'            id = 0
'
'            '��ֵ
'            gcHeaderTemp.Created_By = gUserName
'            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
'            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
'            gcHeaderTemp.Supplier = Rs.fields(2).Value
'            gcHeaderTemp.ShipTo = Rs.fields(3).Value
'            gcHeaderTemp.FAB_Device = Rs.fields(4).Value
'            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
'            gcHeaderTemp.GC_Version = IIf(IsNull(Rs.fields(6).Value), "", Rs.fields(6).Value)
'            'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
'            gcHeaderTemp.GC_Date = Rs.fields(7).Value
'
'
'            gcHeaderTemp.Lot_ID = Rs.fields(8).Value
'            gcDetailTemp.Lot_ID = Rs.fields(8).Value
'            gcDetailTemp.Wafer_ID = Rs.fields(9).Value
'            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(10).Value)
'            gcHeaderTemp.WO_NO = Rs.fields(11).Value
'            gcHeaderTemp.remarkTemp = Rs.fields(12).Value
'            gcHeaderTemp.Date_Code = Rs.fields(13).Value
'            gcHeaderTemp.Marking_Lot_ID1 = Rs.fields(14).Value
'            gcHeaderTemp.Marking_Lot_ID2 = Rs.fields(15).Value
'            gcDetailTemp.Marking_Lot_ID = Rs.fields(14).Value & " " & Rs.fields(15).Value
'
'
'
'            '2013-12-05 jiayun add
'            '�ж�wo�Ƿ�Ϊ��
'
'            If Trim(gcHeaderTemp.WO_NO) = "" Then
'
'                MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
'                Exit Sub
'
'            End If
'
'            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
'
'            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
'
'            '2013-12-27 jiayun add
'
''            If gcDetailTemp.Good_Die_Qty <= 0 Then
''                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
''                    Exit Sub
''            End If
'
'
'            '2012-11-05 jiayun �޸� GC
'
'            '�ж�lotID��Header�����Ƿ��Ѵ���
'
'            If (JudgeEQHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO, gcHeaderTemp.PO_NO)) Then
'
'                If GCHeaderFlag = False Then
'        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
'                End If
'
'                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
'                '��lotid�и���ʱ�����ѯ�ϴε�id
'
''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
'
'            Else
'            '�ϴ���Header����
'                'ȡĿǰDB����ID��
'                id = GetMaxID()
'                '2013-01-11 jiayun add �ͻ����
'
'                If id = 0 Then
'                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
'                    Exit Sub
'
'                Else
'
'
'                    Call AddEQHeader(gcHeaderTemp, id, customerTemp)
'                    GCHeaderFlag = True
'
'                End If
'
'            End If
'
'
'            '�ж�lotID��Detail�����Ƿ��Ѵ���
'
'            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
'               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "�Ѵ��ڣ������ϴ�!"
'
'            Else
'            '�ϴ���Detail����
'
'                   '2012-11-05 jiayun �޸� GCT
'
'
'                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
'
'
'                If id = 0 Then
'                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
'                    Exit Sub
'
'                Else
'                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
'                    SumCount = SumCount + 1
'
'                End If
'
'
'            End If
'
'
'            Rs.MoveNext
'
'        Next i
'
'
'        If SumCount > 0 Then
'            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
'        End If


End Sub



Private Sub UploadMC()
'��ȡCSV
'2013-12-17 jiayun add MC
Dim source_batch_id_Temp As String
Dim customerTemp As String

customerTemp = "MC"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text3.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    End If

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
        Rs.open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        SumCount = 0
        Rs.MoveFirst
        
        GCHeaderFlag = False
        
        For i = 0 To Rs.RecordCount - 1
            temp = ""
            id = 0
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            gcHeaderTemp.Supplier = Rs.fields(2).Value
            gcHeaderTemp.ShipTo = Rs.fields(3).Value
            gcHeaderTemp.FAB_Device = Rs.fields(4).Value
            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
            gcHeaderTemp.GC_Version = IIf(IsNull(Rs.fields(6).Value), "", Rs.fields(6).Value)
            gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
            gcHeaderTemp.GC_Date = Rs.fields(8).Value
            gcHeaderTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Wafer_id = Rs.fields(10).Value
            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(11).Value)
            gcHeaderTemp.WO_NO = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
            
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeMCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
                '��lotid�и���ʱ�����ѯ�ϴε�id
                
                id = GetMCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                
                
                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   
'                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                   
                
                 gcDetailTemp.item = gcDetailTemp.Wafer_id
                 
                 gcDetailTemp.Wafer_id = Right(gcDetailTemp.Wafer_id, 2)
                   
                   
                If id = 0 Then
                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
                    Exit Sub
                
                Else
                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
                
            End If
           
            
            Rs.MoveNext
        
        Next i
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub


'2014-02-10 jiayun add
Private Sub UploadNormalCustomer(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim mCodetemp As String
Dim yTemp As String
Dim mTemp As String
Dim charTemp As Long


   


SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(Date), 1)
                    mTemp = GetMonthChar(Month(Date))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
                
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
            
               If customerNameTemp = "MR" Then
                gcDetailTemp.Wafer_id = Right(Trim(tempVal), 2)
                
               Else
            
                        If IsNumeric(Trim(tempVal)) = False Then
                         MsgBox "WaferId���Ͳ��ԣ���˶�Ҫ�ϴ���Դ�ĵ� !"
                         Exit Sub
                        
                        Else
                         
                         gcDetailTemp.Wafer_id = Trim(tempVal)
                         
                         End If
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
               If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "GT" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
             
                   ElseIf customerNameTemp = "MR" Then
                   
                  gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_id), 2)
                
                  Else
                
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call AddGCDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadNormalCustomer77(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim mCodetemp As String
Dim yTemp As String
Dim mTemp As String
Dim charTemp As Long


   


SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(Date), 1)
                    mTemp = GetMonthChar(Month(Date))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
                
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
            
               If customerNameTemp = "MR" Then
                gcDetailTemp.Wafer_id = Right(Trim(tempVal), 2)
                
               Else
            
'                        If IsNumeric(Trim(tempVal)) = False Then
'                         MsgBox "WaferId���Ͳ��ԣ���˶�Ҫ�ϴ���Դ�ĵ� !"
'                         Exit Sub
'
'                        Else
                         
                         gcDetailTemp.Wafer_id = Trim(tempVal)
                         
                         'End If
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
               If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "GT" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
             
                   ElseIf customerNameTemp = "MR" Then
                   
                  gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_id), 2)
                
                  Else
                
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call AddGCDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub




'2015-09-11 jiayun add 56
Private Sub UploadNormalCustomer56(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim mCodetemp As String
Dim yTemp As String
Dim mTemp As String
Dim charTemp As Long
Dim waferAllDieQty As Long

   


SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    waferAllDieQty = 0
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(Date), 1)
                    mTemp = GetMonthChar(Month(Date))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
                
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
            
               If customerNameTemp = "MR" Then
                gcDetailTemp.Wafer_id = Right(Trim(tempVal), 2)
                
               Else
            
                        If IsNumeric(Trim(tempVal)) = False Then
                         MsgBox "WaferId���Ͳ��ԣ���˶�Ҫ�ϴ���Դ�ĵ� !"
                         Exit Sub
                        
                        Else
                         
                         gcDetailTemp.Wafer_id = Trim(tempVal)
                         
                         End If
                
                End If
                
            End If
            
            If j = 12 Then
                'gcDetailTemp.Good_Die_Qty = Trim(tempVal)
                waferAllDieQty = CLng(Trim(tempVal))
                
            End If
            
            
             If j = 13 Then
                gcDetailTemp.Good_Die_Qty = 0
             
                gcDetailTemp.NG_Die_Qty = 0
                
                 
            End If
            
               If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
               End If
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "GT" Then
          SumCount = SumCount + 1
          
    ElseIf customerNameTemp = "56" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
             
                   ElseIf customerNameTemp = "MR" Then
                   
                  gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_id), 2)
                
                  Else
                
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call Add56Detail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



'2014-02-10 jiayun add
Private Sub UploadNormalCustomerZL(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim mCodetemp As String
Dim yTemp As String
Dim mTemp As String
Dim charTemp As Long


   


SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(Date), 1)
                    mTemp = GetMonthChar(Month(Date))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
                
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
            
               If IsNumeric(Trim(tempVal)) = False Then
                MsgBox "WaferId���Ͳ��ԣ���˶�Ҫ�ϴ���Դ�ĵ� !"
                Exit Sub
               
               Else
               
                gcDetailTemp.Wafer_id = Trim(tempVal)
                
                End If
                
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty
            End If
            
    
            
            
               If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "GT" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
                   Else
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call AddGCDetailZL(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

'2015-04-08 jiayun add
Private Sub UploadNormalCustomerCS(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 15 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim mCodetemp As String
Dim yTemp As String
Dim mTemp As String
Dim charTemp As Long


   


SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(Date), 1)
                    mTemp = GetMonthChar(Month(Date))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
                
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
            
               If IsNumeric(Trim(tempVal)) = False Then
                MsgBox "WaferId���Ͳ��ԣ���˶�Ҫ�ϴ���Դ�ĵ� !"
                Exit Sub
               
               Else
               
                gcDetailTemp.Wafer_id = Trim(tempVal)
                
                End If
                
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty
            End If
            
    
            
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
            If j = 15 Then
                gcHeaderTemp.Date_Code = Trim(tempVal)
            End If
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddCSHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "GT" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
                   Else
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call AddGCDetailZL(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



'2014-09-17 jiayun add
Private Sub UploadQR(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
               If j = 13 Then
                gcDetailTemp.NG_Die_Qty = Trim(tempVal) - gcDetailTemp.Good_Die_Qty
            End If
            
               If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    '2014-03-04 jiayun add  CN Wo  ���������ݵ�Mapping��

      If customerNameTemp = "CN" Then
         SumCount = SumCount + 1
      
      ElseIf customerNameTemp = "SI" Then
          SumCount = SumCount + 1
      
      Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
                   
                   If customerNameTemp = "CN" Then
                   gcDetailTemp.item = gcDetailTemp.Wafer_id
                   
                   Else
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                   End If
                   

                   Call AddQRDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
     End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

'2015-09-07 jiayun add  QR�ڶ��λ���
Private Sub UploadQRV2(customerNameTemp As String)
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = customerNameTemp

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
               
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
               If j = 13 Then
                gcDetailTemp.NG_Die_Qty = Trim(tempVal) - gcDetailTemp.Good_Die_Qty
            End If
            
               If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
        
    Next j
    
    

     If (JudgeQR2HeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetQR2LotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddQR2Header(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    

    If (JudgeQR2DetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����

           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)

           Call AddQR2Detail(gcDetailTemp, customerTemp, id)
           
        SumCount = SumCount + 1
      
    End If
            
  
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub




Private Sub UploadHY()
'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String

customerTemp = "HY"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text3.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    End If

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
        Rs.open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        SumCount = 0
        Rs.MoveFirst
        
        GCHeaderFlag = False
        
        For i = 0 To Rs.RecordCount - 1
            temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            gcHeaderTemp.Supplier = IIf(IsNull(Rs.fields(2).Value), "", Rs.fields(2).Value)
            gcHeaderTemp.ShipTo = IIf(IsNull(Rs.fields(3).Value), "", Rs.fields(3).Value)
            gcHeaderTemp.FAB_Device = IIf(IsNull(Rs.fields(4).Value), "", Rs.fields(4).Value)
            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
            gcHeaderTemp.GC_Version = Rs.fields(6).Value
            gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
            gcHeaderTemp.GC_Date = Rs.fields(8).Value
            gcHeaderTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Wafer_id = Rs.fields(10).Value
            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(11).Value)
            gcHeaderTemp.WO_NO = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(gcHeaderTemp.Customer_Device, gcDetailTemp.Good_Die_Qty)
   
            
            
            '2012-11-05 jiayun �޸� GC
            
            
            
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "HY ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
              
            End If
           
            
            Rs.MoveNext
        
        Next i
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub


Private Sub UploadHT()
'��ȡCSV
Dim source_batch_id_Temp As String
Dim customerTemp As String

customerTemp = "HT"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Text3.Text), "\") > 0 Then
        StrFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    End If

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
        Rs.open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        SumCount = 0
        Rs.MoveFirst
        
        GCHeaderFlag = False
        
        For i = 0 To Rs.RecordCount - 1
            temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.item = Rs.fields(0).Value
            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            gcHeaderTemp.Supplier = Rs.fields(2).Value
            gcHeaderTemp.ShipTo = Rs.fields(3).Value
            gcHeaderTemp.FAB_Device = Rs.fields(4).Value
            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
            gcHeaderTemp.GC_Version = Rs.fields(6).Value
            gcDetailTemp.Marking_Lot_ID = Rs.fields(7).Value
            gcHeaderTemp.GC_Date = Rs.fields(8).Value
            gcHeaderTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Lot_ID = Rs.fields(9).Value
            gcDetailTemp.Wafer_id = Rs.fields(10).Value
            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(11).Value)
            gcHeaderTemp.WO_NO = Rs.fields(12).Value
            
            '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
  
            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(gcHeaderTemp.Customer_Device, gcDetailTemp.Good_Die_Qty)
   
            
            
            '2012-11-05 jiayun �޸� GC
            
            
            
            
            '�ж�lotID��Header�����Ƿ��Ѵ���
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
            Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
                '2013-01-11 jiayun add �ͻ����
                
                
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
            End If
            
            
            '�ж�lotID��Detail�����Ƿ��Ѵ���
            
            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
               MsgBox "HT ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
               
            Else
            '�ϴ���Detail����
            
                   '2012-11-05 jiayun �޸� GCT
                   
                   
                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
                   
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
              
            End If
           
            
            Rs.MoveNext
        
        Next i
        
        
        If SumCount > 0 Then
            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
        End If


End Sub



Private Sub UploadSX36()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "36"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "SX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



Private Sub UploadHJ()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "HJ"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "SX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



Private Sub UploadSX()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "SX"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "SX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub Upload59()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "59"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                 gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                'gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "59 ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub


Private Sub UploadZX()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "ZX"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                 id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "ZX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadOT()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "OT"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
    

                
                
    
   If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                 id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "OT ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub





Private Sub UploadRD()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "RD"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
    If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "RD ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadDN()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer
Dim dnRemark As String

customerTemp = "DN"
dnRemark = ""

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    dnRemark = ""
    
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
            If j = 14 Then
                dnRemark = Trim(tempVal)
            End If
            
        
    Next j
    
    

     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                 
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "RD ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddDNDetail(gcDetailTemp, customerTemp, id, dnRemark)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadPT()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "PT"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgePTHeaderId(gcHeaderTemp.Lot_ID)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "PT ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
'           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
           '2013-03-04 jiayun modify
           gcDetailTemp.item = gcDetailTemp.Wafer_id
           
           gcDetailTemp.Wafer_id = Right$(Trim(gcDetailTemp.Wafer_id), 2)
           
           
           
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadBD()
'2013-06-17 jiayun add BD
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "BD"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    
    
    

Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
   
Dim PShortNameTemp As String



SumCount = 0

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    PShortNameTemp = ""

    
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
            If j = 14 Then
                PShortNameTemp = Trim(tempVal)
            End If
            
            
            
        
    Next j
    
    '2013-12-05 jiayun add У��po���Ƿ�Ϊ��
    
    If Trim(gcHeaderTemp.PO_NO) = "" Then
        MsgBox "PO_NO������Ϊ��ֵ����ȷ�ϣ�", vbInformation, "��ʾ"
        Exit Sub
    
    End If
    
    
     If (JudgePTHeaderId(gcHeaderTemp.Lot_ID)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddBDHeader(gcHeaderTemp, id, customerTemp, PShortNameTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "BD ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
           '2013-03-04 jiayun modify
'           gcDetailTemp.item = gcDetailTemp.Wafer_ID
           
           gcDetailTemp.Wafer_id = Right$(Trim(gcDetailTemp.Wafer_id), 2)
           
           
           
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub


Private Sub UploadSY()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "SY"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgePTHeaderId(gcHeaderTemp.Lot_ID)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "PT ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
'           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
           '2013-03-04 jiayun modify
           gcDetailTemp.item = gcDetailTemp.Wafer_id
           
           gcDetailTemp.Wafer_id = Right$(Trim(gcDetailTemp.Wafer_id), 2)
           
           
           
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



Private Sub UploadSX34()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "34"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 13 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "SX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub

Private Sub UploadSX32()
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim SumCount As Integer

customerTemp = "32"

'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�


    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�
    
    
      '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 14 Then

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

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    source_batch_id_Temp = ""
    
    '��ѯһ�е�ֵ
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ

          temp = ""
        
            '��ֵ
            gcHeaderTemp.Created_By = gUserName
            If j = 1 Then
                gcDetailTemp.item = Trim(tempVal)
            End If
            
            If j = 2 Then
                gcHeaderTemp.PO_NO = Trim(tempVal)
            End If
            
            If j = 3 Then
                gcHeaderTemp.Supplier = Trim(tempVal)
            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)
            End If
            
            If j = 5 Then
                 gcHeaderTemp.FAB_Device = Trim(tempVal)
            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            End If
            
            If j = 7 Then
                 gcHeaderTemp.GC_Version = Trim(tempVal)
            End If
            
            If j = 8 Then
'                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)
            End If
            
             If j = 10 Then
                gcHeaderTemp.Lot_ID = Trim(tempVal)
                gcDetailTemp.Lot_ID = Trim(tempVal)
            End If
            
            If j = 11 Then
                gcDetailTemp.Wafer_id = Trim(tempVal)
            End If
            
'            If j = 12 Then
'                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
'            End If
'
'            If j = 13 Then
'                gcHeaderTemp.WO_NO = Trim(tempVal)
'            End If
'
'
               If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            End If
            
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty
            End If
            
    
            
            
               If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)
            End If
            
        
    Next j
    
    
     If (JudgeSXHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)) Then
            
                If GCHeaderFlag = False Then
        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
                End If
                
                id = GetSXLotIDPOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.PO_NO, gcHeaderTemp.Customer_Device)
                
    Else
            '�ϴ���Header����
                'ȡĿǰDB����ID��
                id = GetMaxID()
       
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
              
     End If
            
            
    '�ж�lotID��Detail�����Ƿ��Ѵ���
    
    If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)) Then
       MsgBox "SX ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_id & "�Ѵ��ڣ������ϴ�!"
       
    Else
    '�ϴ���Detail����
           
           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_id), 2)
           
        Call AddGCDetail(gcDetailTemp, customerTemp, id)
        SumCount = SumCount + 1
      
    End If
    
     
    '�ϴ���DB
NextRecord2:

Next i


     
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
    End If
    
        
End Sub



Private Sub Command7_Click()

UploadGCNew

End Sub


Private Function GetGCGoodDieQty(productNameTemp As String, dieQtyTemp As Long) As Integer
'2013-12-26 jiayun add
'����Gc pt ��ѯ����

GetGCGoodDieQty = 0

Set updateRS = GetWO_GC_Die(productNameTemp)
GetGCGoodDieQty = CInt(updateRS.fields("dieqty").Value)

'Dim productNameTemp2 As String
'
'If productNameTemp <> "" And dieQtyTemp > 0 Then
'    productNameTemp2 = UCase(Left(Trim(productNameTemp), Len(Trim(productNameTemp)) - 2))
'
'    Select Case productNameTemp2
'
'    Case "GC6113"
'        GetGCGoodDieQty = 6975
'
'    Case "GC0311"
'        GetGCGoodDieQty = 5584
'
'    Case "GC0329"
'        GetGCGoodDieQty = 4722
'
'    Case "GC0313"
'        GetGCGoodDieQty = 5364
'
'    Case "GC2035"
'        GetGCGoodDieQty = 1547
'
'    Case "GC6123"
'        'GetGCGoodDieQty = 8688
'        '2013-11-04 jiayun modify
'
'        GetGCGoodDieQty = 8706
'
'    Case "GC0328"
'        GetGCGoodDieQty = 3382
'
'    Case "GC1004"
'        GetGCGoodDieQty = 1302
'
'    Case Else
'        GetGCGoodDieQty = 0
'
'    End Select
'
'Else
'
'    GetGCGoodDieQty = 0
'End If


End Function



Private Function GetGCVerLastChar(ptTemp As String) As String
'2013-12-26 jiayun add
'����Gc pt ��ѯ����

GetGCVerLastChar = ""

Set updateRS = GetWO_GC_Ver(ptTemp)
If updateRS.RecordCount > 0 Then


GetGCVerLastChar = CStr(updateRS.fields("Gcversion").Value)

Else

GetGCVerLastChar = ""
End If

End Function





Private Sub Command8_Click()

If CmbCustomer.Text = "" Then
 MsgBox "����ѡ��ͻ���"
 Exit Sub
End If

 ExporToExcel ("  select po_num as PO_NO, ship_site as Supplier,test_site as Ship_To, fab_conv_id as FAB_Device, mpn_desc as Customer_Device," & _
               " imager_customer_rev as GC_Version,created_date as GC_Date,source_batch_id  as Lot_ID, mtrl_num   As WO_NO " & _
               " From CustomerOItbl_test  where CustomerShortName = '" & CmbCustomer.Text & "'order by id ")
 



End Sub

Private Sub Command9_Click()

  ExporToExcel ("select ID,WO_NO,PO_NO,CustomerDevice,LotID,Waferid,WLAFlag,CREATEDDATE,LASTUPDATEDATE from  TSV_GCLable_SETWLA  order by id desc ")
 
 
End Sub

Private Sub Form_Load()


'Com.flags = &H80200
'
'ComSI.flags = &H80200

'CmbCustomer.AddItem ("GC")
'CmbCustomer.AddItem ("GC_WLD/T")
'CmbCustomer.AddItem ("SX")
'CmbCustomer.AddItem ("HJ")
'
'CmbCustomer.AddItem ("PT")
'CmbCustomer.AddItem ("SY")
'CmbCustomer.AddItem ("RD")
'CmbCustomer.AddItem ("DN")
'CmbCustomer.AddItem ("BD")
'CmbCustomer.AddItem ("ZX")
'CmbCustomer.AddItem ("HY")
'CmbCustomer.AddItem ("HT")
'CmbCustomer.AddItem ("OT")
'CmbCustomer.AddItem ("MC")
''2014-09-17 jiayun modify si ��ΪGT
'CmbCustomer.AddItem ("GT")
'
'CmbCustomer.AddItem ("CN")
'CmbCustomer.AddItem ("KT")
'CmbCustomer.AddItem ("HD")
'
'CmbCustomer.AddItem ("RS")
'CmbCustomer.AddItem ("SD")
'
'CmbCustomer.AddItem ("QR")
'CmbCustomer.AddItem ("QR2")
'
'CmbCustomer.AddItem ("MG")
'CmbCustomer.AddItem ("LX")
'CmbCustomer.AddItem ("GD")
'CmbCustomer.AddItem ("AM")
'CmbCustomer.AddItem ("EQ")
'CmbCustomer.AddItem ("EQ_IS")
'CmbCustomer.AddItem ("ZL")
'CmbCustomer.AddItem ("YW")
'CmbCustomer.AddItem ("RO")
'CmbCustomer.AddItem ("MR")
'CmbCustomer.AddItem ("CS")
'
'CmbCustomer.AddItem ("36")
'CmbCustomer.AddItem ("34")
'CmbCustomer.AddItem ("32")
'CmbCustomer.AddItem ("45")
'CmbCustomer.AddItem ("50")
'CmbCustomer.AddItem ("60")
'
'CmbCustomer.AddItem ("30")
'CmbCustomer.AddItem ("55")
'CmbCustomer.AddItem ("54")
'CmbCustomer.AddItem ("56")
'CmbCustomer.AddItem ("49")
'CmbCustomer.AddItem ("59")
'CmbCustomer.AddItem ("64")
'
'CmbCustomer.AddItem ("68")
'CmbCustomer.AddItem ("70")
'CmbCustomer.AddItem ("69")
'CmbCustomer.AddItem ("80")
'
'
'CmbCustomer.AddItem ("XW")
'
'
'CmbCustomer.AddItem ("YX")
'
'CmbCustomer.AddItem ("37")
'CmbCustomer.AddItem ("77")
'
'
'CmbCustomer.AddItem ("XA")
'CmbCustomer.AddItem ("HH")
'CmbCustomer.AddItem ("SL")
'
'
'Combo1.AddItem ("AA")
'Combo1.AddItem ("�Թ�")
'Combo1.AddItem ("CN")
'

End Sub

Private Sub SSTab1_DblClick()

End Sub
