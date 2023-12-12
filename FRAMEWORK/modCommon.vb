Imports System.Data.SqlClient

Module modCommon
    Public retAddUpdate As Boolean = False
    Public retCode As String
    Public FPNL As Byte  '' This variable is used for F-First, P-Previous, N-Next, L-Last
    Public strSql As String
    Public NodeBranchID As String = ""
    Public strSearchReturn1 As String = "" ''CFL RETURN VALUE 1
    Public strSearchReturn2 As String = "" ''CFL RETURN VALUE 2
    Public strSearchReturn3 As String = "" ''CFL RETURN VALUE 3
    Public strSearchReturn4 As String = "" ''CFL RETURN VALUE 4
    Public strSearchReturn5 As String = "" ''CFL RETURN VALUE 5
    Public strSearchReturn6 As String = "" ''CFL RETURN VALUE 6
    Public strSearchReturn7 As String = "" ''CFL RETURN VALUE 7
    Public strSearchReturn8 As String = "" ''CFL RETURN VALUE 8
    Public strSearchReturn9 As String = "" ''CFL RETURN VALUE 9
    Public strSearchReturn10 As String = "" ''CFL RETURN VALUE 10
    Public strSearchReturn11 As String = "" ''CFL RETURN VALUE 11
    Public dt As DataTable
    Public dtempload As DataTable
    Public da As SqlDataAdapter
    Public cmd As SqlCommand
    Public con As New SqlConnection
    Public Maincon As SqlConnection
    Public TCSAMOUNTRANGE As Double = 0
    Public CompanyMainCon As SAPbobsCOM.Company
    Public GoodsReturnTag As String = "N"
    Public MODECON As Boolean = False
    Public MODECOmm As Boolean = False
    Public EventHandlerStatus2 As Integer
    Public Tran As SqlTransaction
    Public MDBName1 As String
    Public MDBMainName As String
    Public MDBServer As String
    Public MDBServerIP As String
    Public TGoldRate As String
    Public PGoldRate As String
    Public CTGoldRate As String
    Public COIGoldRate As String
    Public NodeBranch As String
    Public DBStatused As String
    Public TSilverRate As Double = 0
    Public LoginActive As Boolean
    Public Customervalid As Boolean = False
    'Public comm As InitSerialPort
    Public MachineGrsWt As TextBox
    ''GLOBAL COLORS********************************************
    Public frmBGColorRed As Integer = 234  ' Form Backcolor1
    Public frmBGColorGreen As Integer = 241 ' Form Backcolor2
    Public frmBGColorBlue As Integer = 246 ' Form Backcolor3

    Public frmButtonColorRed As Integer = 255  ' Buttoncolor1
    Public frmButtonColorGreen As Integer = 238 ' Buttoncolor2
    Public frmButtonColorBlue As Integer = 159 ' Buttoncolor3

    Public frmGridColorRed As Integer = 217  ' GridBackcolor1
    Public frmGridColorGreen As Integer = 229 ' GridBackcolor2
    Public frmGridColorBlue As Integer = 242 ' GridBackcolor3

    Public frmTitleColorRed As Integer = 46  ' TitlebarColor1
    Public frmTitleColorGreen As Integer = 79 ' TitlebarColor2
    Public frmTitleColorBlue As Integer = 152 ' TitlebarColor3
    ''*********************************************************

    Public MIPL_Fontsize As Integer = 7
    Public MIPL_Font As String = "Tahoma"
    Public MIPL_FormBackColor As Color = Color.FromArgb(frmBGColorRed, frmBGColorGreen, frmBGColorBlue)
    Public MIPL_ButtonColor As Color = Color.FromArgb(frmButtonColorRed, frmButtonColorGreen, frmButtonColorBlue)
    Public MIPL_Seperator As Char = "/"

    Public ChangingDescText As String = ""
    Public ChangingDescBold As Boolean = False
    Public ChangingDescItalic As Boolean = False

    Public Const MIPL_Msgbox_Title1 As String = "Jewel Addon"
    Public MIPL_StatusBar_Msg As String = "" '' STATUS BAR MESSAGE
    Public MIPL_CtrlKey As Boolean = False '' This is for Changing Description
    Public MIPL_SystemInformation As Boolean = False '' Status Bar Visible in Every Form

    Public MIPL_DefineNewReturn As String = "" '' RETURN VALUE FROM DEFINE NEW TABLE

    Public MIPL_Authorize As Boolean = False '' Authorization for Every Form Open
    Public Remarks As String = ""
    Public strFormToOpen As String = ""
    Public strReasonReturn As String = ""
    Public Fingerassgin As Integer
    Public FingerFILEPATH As String = ""

    ''Customer Sample for order
    Public dtorderbooking As DataTable
    Public dtSalesEstimation As DataTable
    Public dtWorkorder As DataTable
    Public dtmaterialReceipt As DataTable
    Public formname_custsample As String

    Public formname As String
    Public orderno As String
    Public Salesestimationno As String
    Public WeightDecimalPlace As Byte = 3 ''Weight Decimal Place(Grosswt,Netwt,Purewt,Lesswt)
    Public AmountDecimalPlace As Byte = 2 ''Amount Decimal Place(Amount,Price)
    Public PercentageDecimalPlace As Byte = 2 ''Percentage Decimal Place

    Public Username As String ''User Name for Enter user
    Public Userid As String ''User id for Enter user
    Public SAPUserid As String ''User id for Enter user

    Public dtbranchcommon As DataTable ''Select Branch All

    Public dtcommonfreight As DataTable  ''Datatable for Freight Charges

    Public dtcreditcard As DataTable ''Data Table for Credit  Card
    Public dtdescr As DataTable ''Data Table for Gift Description  Card


    Public dtmultiselect As DataTable  ''FOr Multi select Details
    Public dtmultiselectorder As DataTable  ''FOr Multi select Details
    Public dtmultioffer As DataTable

    Public ocr1 As String  'for othercharges reason1
    Public ocr2 As String  'for othercharges reason2
    Public ocr3 As String  'for othercharges reason3
    Public oca1 As String  'for othercharges amount1
    Public oca2 As String  'for othercharges amount2
    Public oca3 As String  'for othercharges amount3



    Public ocrM1 As String  'for othercharges reason1
    Public ocrM2 As String  'for othercharges reason2
    Public ocrM3 As String  'for othercharges reason3
    Public ocrM4 As String  'for othercharges reason1
    Public ocrM5 As String  'for othercharges reason2
    Public ocrM6 As String  'for othercharges reason3
    Public ocaM1 As String  'for othercharges amount1
    Public ocaM2 As String  'for othercharges amount2
    Public ocaM3 As String  'for othercharges amount3
    Public ocaM4 As String  'for othercharges amount1
    Public ocaM5 As String  'for othercharges amount2
    Public ocaM6 As String  'for othercharges amount3

    Public Verfication_Sucess As String  'Double conformation of chit card creation
    Public Customer_Code_From_MIPLCM As String  'for customer details
    Public vendor_Code_From_MIPLCM As String    'for Vendor Details
    Public Update_Customer_Details As String  'for CustomerUPdate Details

    Public MIPL_ENVIRONMENT_NOEDEID As String
    Public MIPL_ADDON_VERSION As String
    Public MIPL_ENVIRONMENT_BRANCHID As String
    Public MDAURUM As String
    Public INSTA_TABLEID As String
    Public dtrateRET As DataTable
    Public dtrateSRET As DataTable

    Public dt_Payment_CC As DataTable ''Payment For CC
    Public Payment_Cash As Decimal  ''Payment Cash
    Public Payment_Modes_limit As String = "SI"  ''Payment Cash
    Public Payment_ES_Chit As Decimal  ''Payment Cash
    Public Payment_order As Boolean = False   ''Payment Cash
    Public Payment_Cash1 As Decimal ''Payment Out Cash
    Public totalsoudhand As Double ''Payment of sodexo and credit ( Handling chr)
    Public OFFERCOUPONAMT As Double ''OFFER COUPON AMOUNT
    Public OFFERCOUPONNO As String ''OFFER COUPON NO
    Public OFFERCOUPONDIS As String ''OFFER COUPON DESCRIPION
    Public dt_Payment_Cheque As DataTable
    Public dt_Payment_Cheque1 As DataTable
    Public dt_Payment_Chit As DataTable
    Public dt_PT_Issue As DataTable
    Public dt_Payment_OG As DataTable
    Public dt_Reason As DataTable ''Payment For CC
    Public dt_NEFT_RECIVED As DataTable
    ''Customer Sample for Repair Module
    Public dtRMReceipt As DataTable
    Public dtGSIssue As DataTable
    Public dtGSReturn As DataTable
    Public dtRMDelivery As DataTable

    Public Customer_Sales_Invoice_No As String
    Public Customer_Repair_No As String
    Public Customer_Maturity_No As String

    Public Customer_CHitNo As String
    Public Customer_IntroducerSlipName As String
    Public Customer_CHit_ClosedCardNo As String

    Public MIPL_VENDOR_CODE As String
    Public Cancel_remaks_Estimation As String

    Public INTRO_SLIP_SCHEMENAME As String
    Public INTRO_SLIP_GROUPAMOUNT As String
    Public INTRO_SLIP_EMPID As String
    Public INTRO_SLIP_BRANCH As String
    Public INTRO_SLIP_SLIPNAME As String
    Public MsgBoxOutPut As String
    Public ApprovalUName As String = ""
    Public ApprovalPswd As String = ""

    Public dt_Payment_Advance, dt_Payment_Privilage As DataTable
    Public pRIVILAGE_CARD_ADJUSTED_NUMBER, pRIVILAGE_CARD_ADJUSTED_AMOUNT, OG_EXCHANGE_MSGBOX As String
    Public xmlstring As String
    Public xmlstringline As String
    Public user As String = "N"
    'Public objMdi_Form As New MDI_JEWELADDON

    Function ConvertText(ByVal str As String, Optional ByVal isLike As Boolean = True) As String
        If isLike Then
            Return "'" + str.Replace("'", "''").Replace("*", "%") + "'"
        Else
            Return "'" + str.Replace("'", "''") + "'"
        End If
    End Function


    Public bool As Boolean




    

End Module
