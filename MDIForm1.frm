VERSION 5.00
Object = "{C30D627D-2D3A-433E-B3B6-6D83CC5D0B98}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "����Ƽ�����ɽ�� TSV ��������ϵͳ"
   ClientHeight    =   10755
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19845
   LinkTopic       =   "MDIForm"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin HookMenu.ctxHookMenu ctxHookMenu2 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   900
      _ExtentY        =   900
      MenuGradientColor=   15003890
      MenuForeColor   =   -2147483640
      MenuBorderColor =   -2147483632
      MenuGradientSelectColor=   12775167
      PopupBorderColor=   -2147483632
      PopupBorderColor=   -2147483640
      PopupGradientSelectColor=   0
      SideBarColor    =   15118989
      SideBarGradientColor=   14743798
      CheckForeColor  =   -2147483641
      ShadowColor     =   0
      OfficeMenuTheme =   0
      BmpCount        =   6
      Bmp:1           =   "MDIForm1.frx":10039E
      Key:1           =   "#BGATUI"
      Bmp:2           =   "MDIForm1.frx":1007C6
      Key:2           =   "#si_h"
      Bmp:3           =   "MDIForm1.frx":100BEE
      Key:3           =   "#lxpian"
      Bmp:4           =   "MDIForm1.frx":101016
      Mask:4          =   16777215
      Key:4           =   "#CusPOInfoSys"
      Bmp:5           =   "MDIForm1.frx":101C68
      Mask:5          =   16777215
      Key:5           =   "#ShippingSchedule"
      Bmp:6           =   "MDIForm1.frx":1028BA
      Mask:6          =   16777215
      Key:6           =   "#POPRICESYS"
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MAINTAIN 
      Caption         =   "[1]&�ͻ���Ϣ�ϴ�"
      Begin VB.Menu uploadncmr 
         Caption         =   "�ϴ�NCMR"
      End
      Begin VB.Menu �ϴ�MNINVOICE 
         Caption         =   "�ϴ�MNINVOICE"
      End
      Begin VB.Menu �ϴ�INVOICE 
         Caption         =   "�ϴ�INVOICE"
      End
      Begin VB.Menu UploadHW 
         Caption         =   "�ϴ�HUAWEI��ǩ����"
      End
      Begin VB.Menu Frm_UploadAWPo 
         Caption         =   "�ϴ���ΪPO"
      End
      Begin VB.Menu Frm_UP_OI 
         Caption         =   "�ϴ��ͻ�OI,WO����"
      End
      Begin VB.Menu CusPOInfoSys 
         Caption         =   "�ͻ�����ά��ϵͳ"
      End
      Begin VB.Menu Frm_UP_OINew 
         Caption         =   "�ϴ��ͻ������붩����Ϣ"
      End
      Begin VB.Menu Frm_BC_Upload 
         Caption         =   "�ϴ��ͻ�BC"
      End
      Begin VB.Menu Frm_PDM_Upload 
         Caption         =   "�ϴ��ͻ�PDM"
      End
      Begin VB.Menu Frm_BI_Upload 
         Caption         =   "�ϴ��ͻ�BI"
      End
      Begin VB.Menu UploadCusShipAddress 
         Caption         =   "�ϴ��ͻ�������ַ"
      End
      Begin VB.Menu Frm_VT_Upload 
         Caption         =   "ί��ػ������ϴ�"
      End
      Begin VB.Menu UpShippingInfo 
         Caption         =   "�ϴ�Shipping��ͻ�������Ϣ"
      End
      Begin VB.Menu Frm_Shipping 
         Caption         =   "DN/SO����-37/SG005"
      End
      Begin VB.Menu UploadShipSite 
         Caption         =   "�ϴ�������ַ"
      End
      Begin VB.Menu UploadKR001WaferSeq 
         Caption         =   "�ϴ�US026Wafer���"
      End
      Begin VB.Menu AH033_SN_PK 
         Caption         =   "AH033�ͻ�SN����_PK����"
      End
      Begin VB.Menu UploadB2B 
         Caption         =   "US026��Ʊ��Ϣά��"
      End
      Begin VB.Menu Frm_AASpecialGR 
         Caption         =   "AA����GR"
      End
      Begin VB.Menu Frm_AAGR 
         Caption         =   "�ͻ�������Ϣ"
      End
      Begin VB.Menu Frm_GCGR_V2 
         Caption         =   "GC�ͻ�������Ϣ(�°��ʽ)"
      End
      Begin VB.Menu Frm_AW_OUTPUT_DATA 
         Caption         =   "��Ϊ��������"
      End
      Begin VB.Menu Frm_Set_Price 
         Caption         =   "�г���NPI��Ʒ�۸�ά��"
      End
      Begin VB.Menu add1 
         Caption         =   "-"
      End
      Begin VB.Menu ShippingSchedule 
         Caption         =   "�����ƻ�"
      End
      Begin VB.Menu aaline 
         Caption         =   "-"
      End
      Begin VB.Menu Frm_ONBC_Upload 
         Caption         =   "ON_BC�ϴ�"
      End
      Begin VB.Menu Frm_ONForecast_Upload 
         Caption         =   "ON_ForeCast�ϴ�"
      End
      Begin VB.Menu Frm_ON_MSLevel 
         Caption         =   "MS-Level��Ϣ"
      End
      Begin VB.Menu Frm_ON_MPN 
         Caption         =   "MPN_Attributes"
      End
      Begin VB.Menu Frm_ON_PTCross 
         Caption         =   "HuatianPartCross"
      End
      Begin VB.Menu Frm_ON_EBR 
         Caption         =   "ON_EBR"
      End
      Begin VB.Menu Frm_ON_MarkingCode 
         Caption         =   "MPS�ͻ���Ʊ��"
      End
      Begin VB.Menu woP 
         Caption         =   "-"
      End
      Begin VB.Menu POPRICESYS 
         Caption         =   "�г��������۸�ά��"
      End
      Begin VB.Menu Frm_WO_PriceSplit 
         Caption         =   "�г��������۸��PO"
      End
      Begin VB.Menu GC_DIE_SECCODE 
         Caption         =   "GC�ϴ�WO��������ά��"
      End
      Begin VB.Menu frm_gc_doublecode 
         Caption         =   "GC������������ά��"
      End
      Begin VB.Menu GC_LABEL_SENDREPORT01 
         Caption         =   "GC��ǩ�������϶�������ά��"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu BaseSetting 
      Caption         =   "[2]TSV�¹���"
      Begin VB.Menu Frm_App_WO 
         Caption         =   "PMC������"
      End
      Begin VB.Menu PMCPROTEST 
         Caption         =   "PMC����������"
      End
      Begin VB.Menu ORDERDEL 
         Caption         =   "PMC����ά��"
      End
      Begin VB.Menu CCCCCC 
         Caption         =   "�ع�WO��������"
      End
      Begin VB.Menu WLCORDER 
         Caption         =   "WLC������"
      End
      Begin VB.Menu OrderHistory 
         Caption         =   "��������¼"
      End
      Begin VB.Menu Frm_App_WO2 
         Caption         =   "��������"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Frm_ToERP_WO2 
         Caption         =   "������ϵͳ(����,��Ʒ,�ع�)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Frm_NotToERP_WO2 
         Caption         =   "��������ϵͳ(����,��Ʒ,�ع�)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Frm_CheckGCWO 
         Caption         =   "GC WO ȷ��"
      End
      Begin VB.Menu Frm_Set_Time 
         Caption         =   "ʱ�� ��ע �趨"
      End
      Begin VB.Menu Frm_RunCard 
         Caption         =   "Ŀǰ���̿����԰汾���趨"
      End
      Begin VB.Menu Frm_Close_WO 
         Caption         =   "�رչ���"
      End
      Begin VB.Menu Frm_DealDoubleData 
         Caption         =   "�׹��������ظ�����"
      End
      Begin VB.Menu FrmTSVBomModify 
         Caption         =   "TSV����Bom�����޸�"
      End
      Begin VB.Menu Frm_GCTraySplitWafer 
         Caption         =   "GC Tray ��WLA��Normal�����趨"
      End
      Begin VB.Menu Frm_GCLableSplitWafer 
         Caption         =   "GC��ұ�ǩWLA��WaferID�趨"
      End
      Begin VB.Menu Frm_GCCT 
         Caption         =   "�ͻ�CT�����ܱ����ѯ"
      End
      Begin VB.Menu gdpc 
         Caption         =   "�����ų�"
      End
   End
   Begin VB.Menu SPCDATE 
      Caption         =   "[3]���Ϻ�����趨"
      Begin VB.Menu Frm_Set_PF 
         Caption         =   "PF�����趨"
      End
      Begin VB.Menu Frm_Set_test01 
         Caption         =   "���԰汾�趨"
      End
      Begin VB.Menu Frm_Set_test02 
         Caption         =   "���԰汾��OI�趨"
      End
      Begin VB.Menu Frm_Set_PT 
         Caption         =   "�ͻ��Ϻ��趨"
      End
      Begin VB.Menu Frm_QueryLot 
         Caption         =   "��ѯLot��"
      End
      Begin VB.Menu Frm_MergQueryLot 
         Caption         =   "��ѯ������Ϣ"
      End
      Begin VB.Menu BOMUP 
         Caption         =   "BOM�ϴ�"
      End
      Begin VB.Menu Frm_bom_setup 
         Caption         =   "TSVBom�趨"
      End
      Begin VB.Menu Frm_bom 
         Caption         =   "TSVBom��ѯ���޸�_ģ��Bom�ϴ�"
      End
      Begin VB.Menu Frm_Set_PT2 
         Caption         =   "�ͻ��Ϻ��볧���ϺŹ�ϵ�趨(��AA)"
      End
      Begin VB.Menu Frm_SetCode 
         Caption         =   "������ά��"
      End
      Begin VB.Menu Frm_Set_PT3 
         Caption         =   "NPI��Ʒ���ƶ��ձ�ά��"
      End
      Begin VB.Menu ubm 
         Caption         =   "UBM ��С"
         Visible         =   0   'False
      End
      Begin VB.Menu Litho_Via 
         Caption         =   "Litho Via"
         Visible         =   0   'False
      End
      Begin VB.Menu LithoTrench 
         Caption         =   "LithoTrench"
         Visible         =   0   'False
      End
      Begin VB.Menu Etch 
         Caption         =   "Thickness Etch1"
         Visible         =   0   'False
      End
      Begin VB.Menu Leadwidth 
         Caption         =   "Lead width ��·���"
         Visible         =   0   'False
      End
      Begin VB.Menu WaferThickNessAfterEtch1 
         Caption         =   "WaferThickNessAfterEtch1"
         Visible         =   0   'False
      End
      Begin VB.Menu SI_TH 
         Caption         =   "Silicon thickness"
         Visible         =   0   'False
      End
      Begin VB.Menu OpenOfSMF 
         Caption         =   "Open Of SMF"
         Visible         =   0   'False
      End
      Begin VB.Menu GC_via_thickness 
         Caption         =   "Via thickness after via etch"
         Visible         =   0   'False
      End
      Begin VB.Menu GC_Wafer_thickness 
         Caption         =   "��Բ��ȼ��"
         Visible         =   0   'False
      End
      Begin VB.Menu GC_LE_PY 
         Caption         =   "LEƫ��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu RPT 
      Caption         =   "[4]�����ѯ"
      Begin VB.Menu Frm_AAWIP 
         Caption         =   "��ѯAA�ͻ�ʵʱWIP"
      End
      Begin VB.Menu Frm_WLObill 
         Caption         =   "WLO��ʷ������ѯ"
      End
      Begin VB.Menu Frm_GCQbox 
         Caption         =   "GC��Ŵ�������ϲ�ѯ"
      End
      Begin VB.Menu Frm_Splitwo 
         Caption         =   "�������ѯ"
      End
      Begin VB.Menu FrmPMCbillQuery 
         Caption         =   "PMC������ѯ"
      End
      Begin VB.Menu Frm_NormalWIP 
         Caption         =   "ͨ��WIP�����ѯ"
      End
      Begin VB.Menu Frm_GC_QBoxQuery 
         Caption         =   "GC����Ʒ�����ѯ"
      End
      Begin VB.Menu Frm_AA_MPNQuery 
         Caption         =   "�����Ϣ��ѯ"
      End
      Begin VB.Menu Frm_MD_OIQuery 
         Caption         =   "���ϻ��ܱ����ѯ"
      End
      Begin VB.Menu Frm_MD_OIDetailQuery 
         Caption         =   "������ϸ�����ѯ"
      End
      Begin VB.Menu Frm_MD_SOQuery 
         Caption         =   "�������ܱ����ѯ"
      End
      Begin VB.Menu Frm_GC_TrayQuery 
         Caption         =   "GCTrayǷ�ϱ����ѯ"
      End
      Begin VB.Menu Frm_GCTheoryTrayQuery 
         Caption         =   "GCTray�������������ѯ"
      End
      Begin VB.Menu Frm_37Portal 
         Caption         =   "SemtechPortalTX��������"
      End
      Begin VB.Menu Frm_37Excel 
         Caption         =   "SemtechExcel������������"
      End
      Begin VB.Menu MN_FPC95 
         Caption         =   "95FPCί����ҵ"
      End
      Begin VB.Menu Frm_ORDER_HISTORY 
         Caption         =   "��������¼"
      End
      Begin VB.Menu GWREPORT 
         Caption         =   "���񱨱�"
      End
      Begin VB.Menu FOJFK 
         Caption         =   "FO�ⷿ������ϸ"
      End
      Begin VB.Menu Shipments 
         Caption         =   "����\�������ͨ�ð�"
      End
      Begin VB.Menu wafer��Ϣ��ѯ 
         Caption         =   "wafer��Ϣ��ѯ"
      End
      Begin VB.Menu KHJZCXWOWKGD 
         Caption         =   "���ϴ�WOδ������"
      End
      Begin VB.Menu Frm_wwreport 
         Caption         =   "ί�ⱨ��"
      End
   End
   Begin VB.Menu WLOSetting 
      Caption         =   "[5]��Ϣά��"
      Begin VB.Menu SJZHWH 
         Caption         =   "�����ۺ�ά��ƽ̨"
      End
      Begin VB.Menu JZDBMWHPT 
         Caption         =   "���ִ����ά��ƽ̨"
      End
      Begin VB.Menu FBSJYRKWH 
         Caption         =   "�Ǳ�˰��Բ���ά��"
      End
      Begin VB.Menu WLYXQ 
         Caption         =   "������Ч���޸�"
      End
      Begin VB.Menu SetData 
         Caption         =   "��Ϣά��"
      End
      Begin VB.Menu UploadData 
         Caption         =   "VT����"
      End
      Begin VB.Menu WAFERTOINFO 
         Caption         =   "��Բ������Ϣ"
      End
      Begin VB.Menu CGVH 
         Caption         =   "�ɹ�ά��"
      End
      Begin VB.Menu GWZL 
         Caption         =   "��������ά��"
      End
      Begin VB.Menu JYCWWH 
         Caption         =   "���-��λ.����.�ߴ�ά��"
      End
      Begin VB.Menu ReelInPos_37 
         Caption         =   "37���̲�λ¼��"
      End
      Begin VB.Menu UPLOAD_MO_US023 
         Caption         =   "�ϴ�MO_US023"
      End
      Begin VB.Menu WaferMark 
         Caption         =   "WaferMark"
      End
      Begin VB.Menu gdrkxx 
         Caption         =   "���������Ϣ"
      End
      Begin VB.Menu frzsgc 
         Caption         =   "NPI�����ض���Ϣά��"
      End
      Begin VB.Menu GZBBWH 
         Caption         =   "���ְ汾ά��"
      End
      Begin VB.Menu Frm_Werwai 
         Caption         =   "ί�����"
      End
      Begin VB.Menu Frm_WAFERSEND 
         Caption         =   "��Բ��������"
      End
      Begin VB.Menu Frm_PriceControl 
         Caption         =   "��Ʒ�۸����"
      End
      Begin VB.Menu XBCCX 
         Caption         =   "�ֲ߲߱�ѯ"
      End
      Begin VB.Menu FrmShippingReceipt 
         Caption         =   "����֪ͨ����ǩ�ϴ�"
      End
      Begin VB.Menu FrmMASend 
         Caption         =   "���������Զ�����"
      End
   End
   Begin VB.Menu Qbox 
      Caption         =   "[6]����ͬ��������ǩ"
      Begin VB.Menu LblPrintSysNew 
         Caption         =   "37��ǩ��ӡϵͳ(��)"
      End
      Begin VB.Menu LblMatchSysNew 
         Caption         =   "37��ǩ�˶�ϵͳ(��)"
      End
      Begin VB.Menu CPVS 
         Caption         =   "37����˶�ϵͳ"
      End
      Begin VB.Menu SH103_CARTON_PRINT 
         Caption         =   "SH103�����ǩ��ӡ"
      End
      Begin VB.Menu DA69WXBQ 
         Caption         =   "DA69�����ǩ"
      End
      Begin VB.Menu Frm_Set_Tray 
         Caption         =   "��ʱ���̱�ǩ"
      End
      Begin VB.Menu MatchLabelSys 
         Caption         =   "��ǩ�ȶ��ۺϰ�"
      End
      Begin VB.Menu Frm_App_ToERPQbox 
         Caption         =   "ɾ��ERPԭ���"
      End
      Begin VB.Menu Frm_App_HYWaferid 
         Caption         =   "PackList(WaferID)SETI"
      End
      Begin VB.Menu Frm_App_GCWaferid 
         Caption         =   "PackList(WaferID)GC"
      End
      Begin VB.Menu Frm_App_WaiBaoWaferid 
         Caption         =   "��Բί��(WaferID)��ǩ"
      End
      Begin VB.Menu MN_37_WG 
         Caption         =   "37��ұ�ǩ"
      End
      Begin VB.Menu Frm_App_36Waferid 
         Caption         =   "ί��(WaferID)36��MG�ͻ�"
      End
      Begin VB.Menu Frm_App_QRWaferid 
         Caption         =   "ί��(WaferID)QR�ͻ�"
      End
      Begin VB.Menu Frm_App_45Waferid 
         Caption         =   "��Ƭ(WaferID)45�ͻ�"
      End
      Begin VB.Menu Frm_App_semtech 
         Caption         =   "SemTech�ڡ������ǩ"
      End
      Begin VB.Menu Frm_QboxSize_semtech 
         Caption         =   "SemTech�����ߴ�ά��"
      End
      Begin VB.Menu MN_SemTechLablePrint 
         Caption         =   "SemTech��ǩ��ӡ"
      End
      Begin VB.Menu Frm_Mes_Edc 
         Caption         =   "����EDC��Mes��ֵ"
      End
      Begin VB.Menu Frm_Qbox 
         Caption         =   "����쳣����"
      End
      Begin VB.Menu FrmSemtechPo 
         Caption         =   "37�ͻ�����PO_NUM"
      End
      Begin VB.Menu MN_WeightFor37 
         Caption         =   "37WaferID����"
      End
      Begin VB.Menu Frm_ShelfNoQuery 
         Caption         =   "���ܺŲ�ѯ"
      End
      Begin VB.Menu Frm_Packing 
         Caption         =   "37���������˶�"
      End
      Begin VB.Menu Frm_HK037 
         Caption         =   "HK037����"
      End
      Begin VB.Menu Frm_ResetBoxNo 
         Caption         =   "�����������"
      End
      Begin VB.Menu Frm_HW_OUT 
         Caption         =   "��Ϊ��ǩ����"
      End
   End
   Begin VB.Menu Label_CHK 
      Caption         =   "[7]��ǩ��ӡ��˶�"
      Begin VB.Menu XB37 
         Caption         =   "�°�37��ӡ"
      End
      Begin VB.Menu XB37HD 
         Caption         =   "�°�37�˶�"
      End
      Begin VB.Menu QR_Chk 
         Caption         =   "��ά��˶�"
      End
      Begin VB.Menu BarCode_Chk 
         Caption         =   "����˶�"
      End
      Begin VB.Menu HuaWeiLabelPrint 
         Caption         =   "37-��Ϊ��ǩ��ӡ"
      End
      Begin VB.Menu HuaWeiLabelVerify 
         Caption         =   "37-��Ϊ��ǩ�˶�"
      End
      Begin VB.Menu HuaweiASNUpload 
         Caption         =   "37-��ΪASN�ϴ�"
      End
      Begin VB.Menu OPkgShipLblPrint 
         Caption         =   "���������ǩ����ͨ�ð�"
      End
      Begin VB.Menu TmpMatch 
         Caption         =   "��ʱ��ǩ����ȶ�(һ��һ)"
      End
      Begin VB.Menu SemTechToHW_ONELOT 
         Caption         =   "37-��Ϊ��ǩ��ӡ��ONELOT��"
      End
      Begin VB.Menu SemTechToHW_ONELOT_MATCH 
         Caption         =   "37-��Ϊ��ǩ�˶ԣ�ONELOT��"
      End
      Begin VB.Menu JPZHGJ 
         Caption         =   "37���̶�ά��ת������"
      End
      Begin VB.Menu PSN_ASN_57 
         Caption         =   "57-REEL_PSN_ASN�󶨺˶�"
      End
      Begin VB.Menu T_57XLJ 
         Caption         =   "57��������"
      End
      Begin VB.Menu WLPTLDN 
         Caption         =   "WLP����DN"
      End
      Begin VB.Menu TYWBDYSYS 
         Caption         =   "���������ǩ��ӡϵͳ(ͨ�ð�1.0)"
      End
      Begin VB.Menu AH017DN 
         Caption         =   "AH017 DNά�����ǩ����"
      End
      Begin VB.Menu HK037CHZLBD 
         Caption         =   "HK037/AC70������ǩ/�������ϱȶ�"
      End
      Begin VB.Menu AT71NWXBD 
         Caption         =   "AT71������ȶ�"
      End
      Begin VB.Menu SH48BQ 
         Caption         =   "SH48��ǩ����"
      End
      Begin VB.Menu WLP_ 
         Caption         =   "-"
      End
      Begin VB.Menu WLPRKSF 
         Caption         =   "WLP��⼴�շ�"
         Begin VB.Menu WLP_GD108 
            Caption         =   "GD108"
         End
         Begin VB.Menu WLP_AT71 
            Caption         =   "AT71"
         End
      End
   End
   Begin VB.Menu Baofei 
      Caption         =   "[8]���ڱ���"
      Begin VB.Menu Frm_Baofei 
         Caption         =   "�ǼǱ�����Ϣ"
      End
      Begin VB.Menu Frm_Baofei_Sigh 
         Caption         =   "��˱�����Ϣ"
      End
      Begin VB.Menu Frm_Baofei_Query 
         Caption         =   "���ϲ�ѯ"
      End
      Begin VB.Menu Frm_RmWo 
         Caption         =   "ɾ����������"
      End
   End
   Begin VB.Menu Frm_develop 
      Caption         =   "[9]����ģʽ"
      Begin VB.Menu SysCustomosReport 
         Caption         =   "���񱨱���ϵͳ"
      End
      Begin VB.Menu Frm_Resize 
         Caption         =   "������������"
      End
      Begin VB.Menu UploadFile 
         Caption         =   "�ϴ��ļ�"
      End
      Begin VB.Menu CreateAccount 
         Caption         =   "��ͨ�˺�"
      End
      Begin VB.Menu MyWorkOrderNew 
         Caption         =   "��������beta"
      End
      Begin VB.Menu TEST1 
         Caption         =   "TEST"
      End
      Begin VB.Menu MRKINGCODERE 
         Caption         =   "����벹��"
      End
      Begin VB.Menu TEST2 
         Caption         =   "37ɨ��ά��"
      End
      Begin VB.Menu SJZHWHPT 
         Caption         =   "�����ۺ�ά��ƽ̨"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BGATUI_Click()
FrmQury.Show
End Sub

Private Sub chipping_Click()
FRM_2_2.Show

End Sub

Private Sub CV_Click()
FRM_CV_HIGHT.Show

End Sub

Private Sub diechg_Click()
Frm_DIECKG.Show
End Sub

Private Sub AH017DN_Click()
FrmDNModule.Show
End Sub

Private Sub AH033_SN_PK_Click()
FrmAH033.Show
End Sub

Private Sub AT71NWXBD_Click()
FrmAT71Match.Show 1
End Sub

Private Sub BarCode_Chk_Click()

Form_MLS.Show
End Sub

Private Sub BOMUP_Click()
Frm_bom_up.Show
End Sub

Private Sub CCCCCC_Click()
FrmDoubleWorkOrder.Show 1
End Sub

Private Sub CGVH_Click()
Frm_CGVH.Show
End Sub

Private Sub CPVS_Click()
    Form_CPVS.Show
End Sub

Private Sub CreateAccount_Click()
If gUserName <> "07885" Then
    MsgBox "���˺��޹���ԱȨ��", vbInformation, "��ʾ"
    Exit Sub
End If

FrmCreateAccount.Show
End Sub

Private Sub CusPOInfoSys_Click()
Form_POSys.Show
End Sub

Private Sub DA69WXBQ_Click()
Frm_DA69_CARTON.Show
End Sub

Private Sub Etch_Click()
Frm_ThicknessEtch1.Show
End Sub

Private Sub Frm_AA_WIP_Click()

End Sub

Private Sub FBSJYRKWH_Click()
If InStr("07952166281282507885", gUserName) = 0 Then
    MsgBox "��û��Ȩ�޽��д������", vbInformation, "����"
    Exit Sub
End If

FrmWaferRK.Show 1
End Sub

Private Sub FOJFK_Click()
frmxskcmxcxfo.Show
End Sub

Private Sub Frm_37Excel_Click()
FrmSemtech_Report.Show

End Sub

Private Sub Frm_37Portal_Click()
Frm37Portal.Show
End Sub

Private Sub Frm_AA_MPNQuery_Click()
Frm_TSV_AAMPNData.Show
End Sub

Private Sub Frm_AAGR_Click()
'2013-11-18 jiayun add aa gr
FrmAAGr.Show
End Sub

Private Sub Frm_AASpecialGR_Click()
FrmSpecialGR.Show
End Sub

Private Sub Frm_AAWIP_Click()
Frm_AA_WIP.Show
End Sub

Private Sub Frm_App_36Waferid_Click()
Frm_36_Label.Show
End Sub

Private Sub Frm_App_45Waferid_Click()
Frm_45_Label.Show
End Sub

Private Sub Frm_App_Dummy_Click()
'FrmDummy.Show
Frm_ProductionPlan.Show
End Sub

Private Sub Frm_App_GCWaferid_Click()
Frm_HY_WaferId_Label.Show
End Sub

Private Sub Frm_App_HYWaferid_Click()
Frm_PackList_Label.Show
End Sub

Private Sub Frm_App_QRWaferid_Click()
Frm_QR_Label.Show
End Sub

Private Sub Frm_App_semtech_Click()
Frm_37_QboxLabel.Show

End Sub

Private Sub Frm_App_ToERPQbox_Click()
FrmDelERPQbox.Show
End Sub

Private Sub Frm_App_WaiBaoWaferid_Click()
Frm_WaiBao_Label.Show
End Sub

Private Sub Frm_App_WO_Click()
'FrmApplyWO.Show
Frm_ProductionPlan.Show
'Frm_ProductionPlanNew.Show
End Sub

Private Sub Frm_App_WO_WLO_Click()
'2013-05-23 add WLO
FrmWLOApplyWO.Show
End Sub

Private Sub Frm_App_WO2_Click()
FrmApplyWO2.Show
End Sub

Private Sub Frm_AW_OUTPUT_DATA_Click()
Frm_AW_Data.Show
End Sub

Private Sub Frm_Baofei_Click()
FrmBaoFeiPutIn.Show
End Sub

Private Sub Frm_Baofei_Query_Click()

FrmBaoFeiQuery.Show

End Sub

Private Sub Frm_Baofei_Sigh_Click()
FrmBaoFeiSign.Show

End Sub

Private Sub Frm_BC_Upload_Click()
'PMC BC�����ϴ�
FrmUpLoadBC.Show
End Sub

Private Sub Frm_BI_Upload_Click()
'2013-03-29 add
FrmUpLoadBI.Show
End Sub

Private Sub Frm_bom_Click()
'2013-03-28 add
FrmTSV_Bom_Query.Show

End Sub

Private Sub Frm_bom_setup_Click()
FrmTSV_Bom_Setup.Show
End Sub

Private Sub Frm_Chang_Wo_Click()
FormchangeWO.Show
End Sub

Private Sub Frm_CheckGCWO_Click()
FrmCheckGCData.Show
End Sub

Private Sub Frm_Close_WO_Click()
FrmCloseWO.Show
End Sub

Private Sub Frm_DealDoubleData_Click()
FrmDoubleData.Show

End Sub

Private Sub Frm_GC_Die_Click()
FrmGCDie.Show

End Sub

Private Sub frm_gc_doublecode_Click()

'If gc_doublecode Is Nothing Then
'        gc_doublecode.Show
'    Else
'        gc_doublecode.ZOrder
'    End If
    
Frm_Gcversion_shipreport.Show
End Sub

Private Sub Frm_GC_QBoxQuery_Click()
Frm_TSV_ERPQbox.Show
End Sub

Private Sub Frm_GC_TrayQuery_Click()
Frm_TSV_GCTray.Show
End Sub

Private Sub Frm_GCCT_Click()
FrmGCCT.Show
End Sub

Private Sub Frm_GCGR_Click()
FrmGCGr.Show
End Sub

Private Sub Frm_GCGR_V2_Click()
FrmGCGrV2.Show
End Sub

Private Sub Frm_GCLableSplitWafer_Click()
Frm_GC_LableWaferid.Show
End Sub

Private Sub Frm_GCQbox_Click()
'GC��������ϲ�ѯ


FrmGCNeedIn.Show



End Sub

Private Sub Frm_GCTheoryTrayQuery_Click()
Frm_TSV_GC_Theory_Tray.Show
End Sub

Private Sub Frm_GCTraySplitWafer_Click()
Frm_GCTray_SplitWafer.Show
End Sub

Private Sub Frm_HK037_Click()
FormHK037.Show
End Sub

Private Sub Frm_HW_OUT_Click()
FormHWLABLE.Show
End Sub

Private Sub Frm_MD_OIDetailQuery_Click()
Frm_TSV_MDOIDetail.Show
End Sub

Private Sub Frm_MD_OIQuery_Click()
Frm_TSV_MDOI.Show
End Sub

Private Sub Frm_MD_SOQuery_Click()
Frm_TSV_MDSO.Show
End Sub

Private Sub Frm_MergQueryLot_Click()
FrmMergQury.Show
End Sub

Private Sub Frm_Mes_Edc_Click()
FrmUpEDC.Show
End Sub

Private Sub Frm_NormalWIP_Click()
Frm_Normal_WIP.Show
End Sub

Private Sub Frm_NotToERP_WO2_Click()
FrmNotToERPApplyWO2.Show
End Sub

Private Sub Frm_ON_EBR_Click()
FormOn_EBR.Show

End Sub

Private Sub Frm_ON_MarkingCode_Click()
FrmOnMarkingCode.Show
End Sub

Private Sub Frm_ON_MPN_Click()
FrmOnMPN.Show
End Sub

Private Sub Frm_ON_MSLevel_Click()
FrmOnMSLevel.Show
End Sub

Private Sub Frm_ON_PTCross_Click()
FrmOnHTPTCross.Show
End Sub

Private Sub Frm_ONBC_Upload_Click()
FrmUpLoadONBC.Show
End Sub

Private Sub Frm_ONForecast_Upload_Click()
FrmUpLoadONForeCast.Show
End Sub

Private Sub Frm_ORDER_HISTORY_Click()
Frm_OrderHistory.Show
End Sub

Private Sub Frm_Packing_Click()
Frm_PackingQty.Show
End Sub

Private Sub Frm_PDM_Upload_Click()
FrmPDM.Show
End Sub

Private Sub Frm_Product_Click()
FrmCheckGCData.Show
End Sub


Private Sub Frm_PWD_Click()
'�ϴ��ͻ���Ϣ
FrmUpLoadOI.Show
End Sub

Private Sub Frm_PriceControl_Click()
Frm_Price_Control.Show
End Sub

Private Sub Frm_Qbox_Click()
If Frm_AlterQbox Is Nothing Then
        Frm_AlterQbox.Show
Else
       Frm_AlterQbox.ZOrder
   End If
End Sub

Private Sub Frm_QboxSize_semtech_Click()
Frm_37_QboxSize.Show
End Sub

Private Sub Frm_QueryLot_Click()
FrmQury.Show
End Sub

Private Sub Frm_ResetBoxNo_Click()
Form_NEWBOX.Show
End Sub

Private Sub Frm_Resize_Click()
Frm_OrderResize.Show
End Sub

Private Sub Frm_RmWo_Click()
Frm_deve.Show
End Sub

Private Sub Frm_RunCard_Click()
FrmTestSetUp.Show
End Sub

Private Sub Frm_SL_Click()
Form4.Show
End Sub

Private Sub Frm_UclLcl_Click()
'PMC ������
FrmApplyWO.Show

End Sub

Private Sub Frm_Set_PF_Click()
FrmSetPF.Show
End Sub

Private Sub Frm_Set_Price_Click()
FrmNPIProductPrice.Show

End Sub

Private Sub Frm_Set_PT_Click()
FrmSetPT.Show
End Sub

Private Sub Frm_Set_PT2_Click()
FrmSetPT2.Show

End Sub

Private Sub Frm_Set_PT3_Click()
FrmNPIProduct.Show
End Sub

Private Sub Frm_Set_test01_Click()
FrmTestNo.Show
End Sub

Private Sub Frm_Set_test02_Click()
FrmTestNo2.Show
End Sub

Private Sub Frm_Set_Time_Click()
FrmSetTime.Show
End Sub

Private Sub Frm_Set_Tray_Click()
Frm_Tray_tmp.Show
End Sub

Private Sub Frm_ShelfNO_Click()

End Sub

Private Sub Frm_SetCode_Click()
From_SetCode.Show
End Sub

Private Sub Frm_SetData_Click()

End Sub

Private Sub Frm_ShelfNoQuery_Click()
Form_ShelfNo.Show
End Sub

Private Sub Frm_Shipping_Click()
If gUserName = "07885" Then
    Frm_uploadShippingList.Show
Else
    'FrmShipping.Show
    Frm_uploadShippingList.Show
End If
'FrmShipping.Show
End Sub

Private Sub Frm_Splitwo_Click()
FrmSplitLot.Show
End Sub

Private Sub Frm_STQTY_Click()

Frm_LVSPLUS.Show
End Sub

Private Sub Frm_ToERP_WO2_Click()
FrmToERPApplyWO2.Show
'FrmApplyWO.Show
End Sub

Private Sub Frm_UP_OI_Click()
FrmUpLoadOI.Show
End Sub

Private Sub Frm_UP_OINew_Click()
FrmOINew.Show

End Sub

Private Sub Frm_UploadAWPo_Click()
Frm_UploadAw.Show
End Sub

Private Sub Frm_VT_Upload_Click()
FrmVT.Show
End Sub

Private Sub Frm_WAFER_SEND_Click()
Frm_WAFER_SEND.Show
End Sub

Private Sub Frm_WAFERSEND_Click()
Frm_WAFER_SEND.Show
End Sub

Private Sub Frm_Werwai_Click()
Frm_ww.Show
End Sub

Private Sub Frm_WLObill_Click()
Frm_WLO_BILL.Show
End Sub

Private Sub Frm_WO_Price_Click()
FrmMDPriceCreate.Show
End Sub

Private Sub Frm_WO_PriceModify_Click()
FrmMDPriceModify.Show
End Sub

Private Sub Frm_WO_PriceQuery_Click()
FrmMDPriceQuery.Show
End Sub

Private Sub Frm_WO_PriceSplit_Click()
'FrmPOPriceSys_NEW.Show 1
FrmMDPriceSplit.Show
End Sub

Private Sub Frm_WORKORDER_PLUS_Click()
Frm_WORK_ORDER.Show
End Sub

Private Sub Frm_wwreport_Click()
Frm_vtreport.Show

End Sub


Private Sub FrmMASend_Click()
Form_MASend.Show
End Sub

Private Sub FrmPMCbillQuery_Click()
Frm_TSV_BILLQuery.Show
End Sub

Private Sub FrmSemtechPo_Click()
If FrmSemtechPoAlter Is Nothing Then
       FrmSemtechPoAlter.Show
Else
        FrmSemtechPoAlter.ZOrder
End If

End Sub

Private Sub FrmShippingReceipt_Click()
Frm_ShippingReceipt.Show
End Sub

Private Sub FrmTSVBomModify_Click()
FrmTSV_WOBo.Show
End Sub

Private Sub frzsgc_Click()
frm_zsgc.Show
End Sub

Private Sub GC_DIE_SECCODE_Click()
Frm_GC_DIE_SEC_CODE.Show 1

End Sub

Private Sub GC_LABEL_SENDREPORT01_Click()
If GC_LABEL_SENDREPORT Is Nothing Then
        GC_LABEL_SENDREPORT.Show
Else
        GC_LABEL_SENDREPORT.ZOrder
    End If

End Sub

Private Sub GC_LE_PY_Click()
'2012-03-02 add Leƫ��  jiayun
FRM_GC_LE_PY.Show
End Sub

Private Sub GC_via_thickness_Click()
'2011-07-11 add jiayunzhang
'GC�ͻ� Via thickness after via etch
FRM_GC_Via_thickness.Show
End Sub

Private Sub GC_Wafer_thickness_Click()
'2012-02-15 jiayun add ��Բ��ȼ��
FRM_GC_WaferThickness.Show

End Sub

Private Sub gexit_Click()
End
End Sub



Private Sub Laser_Click()

FrmTestNo2.Show
End Sub



Private Sub LablePrintSys_Click()

' ����LPS����
LPS.Show

End Sub



Private Sub GCHD_Click()
Frm_Lbl_Verify_GC.Show
End Sub

Private Sub gdpc_Click()
Form_Prod_Control.Show
End Sub

Private Sub gdrkxx_Click()
FrmGet.Show
End Sub

Private Sub GWREPORT_Click()
FrmGW_Report.Show
End Sub

Private Sub GWZL_Click()
Frm_GWZLWH.Show
End Sub

Private Sub GZBBWH_Click()
FrmGZBB.Show
End Sub

Private Sub HDBQHDXT_Click()
FrmHDLVS.Show 1
End Sub

Private Sub HK037CHZLBD_Click()
FrmHK037ShipCheck.Show 1
End Sub

Private Sub HuaweiASNUpload_Click()
Frm_HWAsn.Show
End Sub

Private Sub HuaWeiLabelPrint_Click()
Dim strValidUserIDLst As String

strValidUserIDLst = "07885103541359815034154326219144081676713598150341676706219144081440312439"
If InStr(strValidUserIDLst, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ�޴�", vbInformation, "��ʾ:"
    Exit Sub
End If

'Frm_37LblPrint.Show
FrmLblPrint_37ToHW.Show

End Sub

Private Sub HuaWeiLabelVerify_Click()
Dim sUser As String

Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "078851404912910151291255115398153971263912439"
Else
    
    sUser = "078851404912910151291255115398153971263912439"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

Frm_LblMatchSys_ONELOT.Show
End Sub

Private Sub JPZHGJ_Click()
Frm37QRLblConverter.Show 1
End Sub

Private Sub JYCWWH_Click()
Frm_WaferSide.Show
End Sub

Private Sub LblMatchSys_Click()
Dim sUser As String

Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "0788514049129101512912551153981539712639"
Else
    
    sUser = "0788514049"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

'Frm_LblMatchSys.Show
'Frm_LblMatchSysNew.Show
End Sub

Private Sub JZDBMWHPT_Click()
FrmMarkCodeRep.Show
End Sub

Private Sub KHJZCXWOWKGD_Click()
FrmWOGD.Show
End Sub

Private Sub LblMatchSysNew_Click()
Dim sUser As String

Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "0788514049129101512912551153981539712639"
Else
    
    sUser = "0788514049129101512912551153981539712639"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

'Frm_LblMatchSys.Show
Frm_LblMatchSysNew.Show
End Sub

Private Sub LblPrintSys_Click()
Dim sUser As String
Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "078851035413598150341543262191440816767135981503416767062191440814403"
Else
    
    sUser = "078851035413598150341543262191440816767135981503416767062191440814403"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

'MyLabelSystem.Show
'MyLabelSystemNew.Show
End Sub

Private Sub LblPrintSysNew_Click()
Dim sUser As String
Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "07885103541359815034154326219144081676713598150341676706219144081440312439"
Else
    
    sUser = "07885103541359815034154326219144081676712439"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

'MyLabelSystem.Show
MyLabelSystemNew.Show
End Sub

Private Sub Leadwidth_Click()
FRM_LeadWith1.Show
End Sub

Private Sub Litho_Via_Click()
FRM_LithoVia.Show
End Sub

Private Sub LithoTrench_Click()
FRM_LithoTrench.Show
End Sub

Private Sub lxpian_Click()
FRM_2_5.Show
End Sub



Private Sub MatchLabel_Click()
Frm_MatchLabel.Show
End Sub

Private Sub MatchLabelSys_Click()
    Form_MLS.Show
End Sub

Private Sub MDIForm_Load()
C_SysName = "TSVSys"
 MenuGrant Me        'Ȩ�޿���

'StatusBar1.Panels(4) = "�û�����" & UCase(gUserName)
''�ļ����Ƿ񴴽�
If Dir(g_Path, vbDirectory) = "" Then
    MkDir g_Path '& "\ErpEngFaliaoTempFile"
End If

If Dir(g_PathNewOrder, vbDirectory) = "" Then
    MkDir g_PathNewOrder '& "\ErpEngFaliaoTempFile"
End If

'If Dir(g_Path37, vbDirectory) = "" Then
'    MkDir g_Path37
'End If

If Dir("C:" & "\SemTechReport", vbDirectory) = "" Then MkDir "C:" & "\SemTechReport"


End Sub

Private Sub MN_37_WG_Click()
Frm_37_WG_Label.Show
End Sub

Private Sub MN_FPC95_Click() 'FPCί��������ҵ
    If Frm95FPC Is Nothing Then
        Frm95FPC.Show
    Else
        Frm95FPC.ZOrder
    End If
End Sub

Private Sub MN_SemTechLablePrint_Click()
If FrmSemtech_LablePrint Is Nothing Then
        FrmSemtech_LablePrint.Show
    Else
        FrmSemtech_LablePrint.ZOrder
    End If
End Sub
'37WaferID����
Private Sub MN_WeightFor37_Click()
    If FrmWeightFor37 Is Nothing Then
        FrmWeightFor37.Show
    Else
        FrmWeightFor37.ZOrder
    End If

End Sub

Private Sub MyLMS_Click()
    Frm_MyLMS.Show
End Sub

Private Sub MyLPS_Click()
    MyLabelSystem.Show
End Sub

Private Sub MyWorkOrderNew_Click()
'Frm_ProductionPlan.Show
FrmApplyWO.Show
End Sub

Private Sub OpenOfSMF_Click()
'GC�ͻ� Open of SMF
FRM_GC_SMF.Show
End Sub

Private Sub Plating_Click()

FrmSetPT.Show

End Sub

Private Sub QIUGAO_Click()

FrmTray.Show
End Sub

Private Sub si_h_Click()
FrmSetPF.Show
End Sub

Private Sub OPkgShipLblPrint_Click()
Frm_OPkgPrintGen.Show
End Sub

Private Sub ORDERDEL_Click()
Frm_WOMOD.Show
End Sub

Private Sub OrderHistory_Click()
Form_OrderHistory.Show
End Sub

Private Sub SemTech_LPS_Click()
frmSemtechLPS.Show
End Sub

Private Sub PMCPROTEST_Click()
Frm_ProductionPlanNew.Show
End Sub

Private Sub POPRICESYS_Click()
FrmPOPriceSys_NEW.Show
End Sub

Private Sub PSN_ASN_57_Click()
'FrmCheckLblSys_57.Show
Frm57HW.Show
End Sub

Private Sub QR_Chk_Click()
Frm_Label_Checking_System.Show
End Sub

Private Sub ReelInPos_37_Click()
Frm_ReelInPos_37.Show

End Sub

Private Sub SemTechToHW_ONELOT_Click()
Dim sUser As String
Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "07885103541359815034154326219144081676713598150341676706219144081440312439"
Else
    
    sUser = "07885103541359815034154326219144081676712439"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If
'Frm_37LblPrint_ONELOT.Show
FrmLblPrint_37ToHW.Show
End Sub

Private Sub SemTechToHW_ONELOT_MATCH_Click()
Dim sUser As String

Dim strDate As String

strDate = Now

If strDate < "2019/1/25" Then
    sUser = "078851404912910151291255115398153971263912439"
Else
    
    sUser = "078851404912910151291255115398153971263912439"
End If

If InStr(sUser, gUserName) = 0 Then
    MsgBox "��ǰ�û���Ȩ��", vbInformation, "��ʾ:"
    Exit Sub
End If

Frm_LblMatchSys_ONELOT.Show
End Sub

Private Sub SetData_Click()
Frm_SetData.Show
End Sub

Private Sub SH103_CARTON_PRINT_Click()
Frm_SH103Carton_Print.Show
End Sub

Private Sub SH48BQ_Click()
FrmSH48BD.Show 1

End Sub

Private Sub Shipments_Click()
    Frm_SHIPMENT.Show
End Sub

Private Sub ShippingSchedule_Click()

frmShippingScheduleSystem.Show

End Sub

Private Sub SI_TH_Click()
FRM_Siliconthickness.Show
End Sub

Private Sub SSPC_Click()
FRM_4_1.Show
End Sub

Private Sub SJZHWH_Click()
FrmMaintainSys.Show 1
End Sub

Private Sub SJZHWHPT_Click()
FrmMaintainSys.Show 1
End Sub

Private Sub SysCustomosReport_Click()
FrmCustomosReport.Show
End Sub

Private Sub T_57XLJ_Click()
Frm57Pro.Show
End Sub

Private Sub TEST1_Click()
FrmDNModule.Show
'test.Show
'FrmWaferRK.Show 1
End Sub

Private Sub TraySeting_Click()
FrmTray.Show
End Sub

Private Sub TEST2_Click()
FrmLblPrint37_QrCode.Show
'Frm_WORKORDER.Show
End Sub

Private Sub TmpMatch_Click()
Frm_MatchTmp.Show
End Sub

Private Sub TYWBDYSYS_Click()
FrmOuterPkgLblSys.Show
End Sub

Private Sub ubm_Click()
FRM_UBM.Show
End Sub

Private Sub via12_Click()
FrmTestNo.Show

End Sub

Private Sub UPLOAD_MO_US023_Click()
Frm_Upload_MO_US023.Show
End Sub

Private Sub UploadB2B_Click()
Frm_B2BUpload.Show
End Sub

Private Sub UploadCusShipAddress_Click()
Frm_UploadCusShipAddress.Show
End Sub

Private Sub UploadData_Click()
LPS.Show
End Sub

Private Sub UploadFile_Click()
FrmUploadFile.Show
End Sub

Private Sub UploadHW_Click()
Frm_HuaWei.Show
'FormHWLABLE.Show
End Sub

Private Sub UploadKR001WaferSeq_Click()
Form_UploadKR001WaferSeq.Show
End Sub

Private Sub uploadncmr_Click()
Formupncmr.Show
End Sub

Private Sub UploadShipSite_Click()
Frm_UploadShipSide.Show
End Sub

Private Sub UpShippingInfo_Click()
FrmShipping.Show
End Sub

Private Sub WaferMark_Click()
Frm_WaferMark.Show
End Sub

Private Sub WaferThickNessAfterEtch1_Click()
FRM_WaferThicknessAfterEtch1.Show
End Sub

Private Sub WAFERTOINFO_Click()
Frm_To_BE_TEST.Show
End Sub

Private Sub wafer��Ϣ��ѯ_Click()
If InStr("16819", gUserName) = 0 And InStr("07885", gUserName) = 0 Then
    MsgBox "��û��Ȩ�޽��д������", vbInformation, "����"
    Exit Sub
End If

FrmPJDZ.Show
End Sub

Private Sub WLCORDER_Click()
FrmWLOApplyWO.Show
End Sub

Private Sub WLPCH_Click()
FrmWLPDelivery.Show 1
End Sub

Private Sub WLP_AT71_Click()
FrmWareHousingCharges.Show
FrmWareHousingCharges.Caption = "WLP��⼴�շ�ͨ��ƽ̨_AT71"
FrmWareHousingCharges.Tag = "AT71"
End Sub

Private Sub WLP_GD108_Click()
FrmWareHousingCharges.Show
FrmWareHousingCharges.Caption = "WLP��⼴�շ�ͨ��ƽ̨_GD108"
FrmWareHousingCharges.Tag = "GD108"
End Sub

Private Sub WLPTLDN_Click()
FormWLP_DN.Show
End Sub

Private Sub WLYXQ_Click()
Frm_SetData_bak.Show
End Sub

Private Sub WLYXQXG_Click()
Frm_SetData_bak.Show
End Sub

Private Sub xb37_Click()
FrmLblPrint37.Show

End Sub

Private Sub XB37HD_Click()
'FrmLblCheck37.Show
Frm_LblMatchSys.Show 1
End Sub

Private Sub XBCCX_Click()
Frm_XBCWH.Show
End Sub

Private Sub �ϴ�INVOICE_Click()
Frminup.Show
End Sub

Private Sub �ϴ�MNINVOICE_Click()
Frm_GSJFP_UpLoad.Show
End Sub
