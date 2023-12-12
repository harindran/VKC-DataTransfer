Imports System.IO
Imports System.Data.SqlClient

Public Class clsAddOn
    'DECLARE SBO OBJECT
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public objCompanyMain As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    'DECLARE LIBRARY OBJECTS
    Public centmasterid As String = ""
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public isItemMaster As Boolean
    Public conns As Boolean
    Private ogeneralmenu
    Private ogeneralmenuchit
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    Dim mcnt As Integer = 0
    ''FORM CLASS OBJECT DECLARATION**********************
    Dim objProMast As frmMProduct ''PRODUCT MASTER
    Dim objSubItemMast As frmMSubProduct  ''SUBITEM MASTER
    Dim objDiscMast As frmMDiscount ''DISCOUNT MASTER
    Dim objPurInv As frmDPurchaseInvoice ''PURCHASE INVOICE
    Dim objGReceipt As frmDMaterialReceipt ''GOODS RECEIPTS
    Dim objLot As frmDLotCreation ''LOT CREATION
    Dim objTag As frmDTagGeneration ''TAG GENERATION
    Dim objcustsmaple As frmdCustomerSample ''Customer Sample
    Dim objsizevb As frmMSize  ''Size Master in vb
    Dim objPurityVB As frmMPurity ''Purity Master in VB
    Dim objHallmarkVB As frmMHallmarkCharges  ''Hallmark Charges in VB
    Dim objSaleCentVB As frmMCentSale ''Sales Cent Master in VB
    Public objpurcent As frmMCentPurchase   ''Purchase cent master in VB
    Public objPwvb As frmMPurchaseWastage  ''Purchase Wastage Master
    Public objIncen As frmMIncentive ''Incentive Master
    Public objRwm As frmMRepairWastage ''Repair Wastage Master
    Public objSwast As frmMSpecialSales ''Special Sales Wastage Master
    Public objswastage As frmMSalesWastage ''Sales Wastage Master
    Public objrate1 As frmMRate 'Rate Master
    Public objratebranchwise As frmMRateMasterBranchWise
    Public objOGRate As frmMOGCashRate ''OG Rate Master
    Public objMetalType As frmMMetalType ''Metal Type Master
    Public objReason As frmMReason ''Reason Master
    Public objOGEst As frmOGEstimation ''OG Estimation
    Public objSalEst As frmDSalesEstimation ''Sales Estimation
    Public objGI As frmDGoodsIssue ''Goods Issue
    Public objGR As frmDGoodsReturn  ''Goods Return
    Public objsalesinvoice As frmDSalesInvoice  ''Sales Invoice
    Public objob As frmDOrderBooking  ''Order Booking
    Public objrmr As frmDRepairMaterialReceipt ''Repair Material Receipt
    Public objrmi As frmDRepairMaterialIssue  ''Repair Material Issue
    Public objwo As frmDWorkOrder  ''Work Order
    Public objOGCashPayment As frmDOGCashPayment
    Public objgsissue As frmDGoldSmithIssue ''GoldSmith Issue
    Public objDGSReturn As frmdGoldSmithReturn


    ''CHIT CREATION
    Public objscheme As frmMSchemeMaster ''Scheme Master
    Public objgrp As frmMGroupMaster     ''Scheme Master

    Public ochit As frmDChitCreation ''chit creation
    Public objreverse As frmDChequeReverse ''Cheque Reverse
    Public objchitmaturity As frmDChitMaturity_Closing

    Dim objGenSet As frmGeneralSettings
    Dim objAuth As frmAuthorization ''Authorization
    Dim objLogin As frmUserNamePassword '' Login Form

    Dim objLct As clsLctDup_Emp '' Location mapping in Branch Definenew(Employee Master)
    Dim objMDI As MDI_JEWELADDON
    ''***************************************************
    Dim Tran As SqlTransaction

    Dim OBJMainCompany As New SAPbobsCOM.Company
    Dim objConnect As New clsGlobalMethods
    Public Sub Intialize(ByVal user As String)
        'Dim objstart As New frmOpeningform(1)
        'objstart.ShowDialog()
        'Dim objstart1 As New frmOpeningform()
        'objstart1.Show()

        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        'objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        ' Dim objConnect As New clsGlobalMethods
        'objApplication = objCompany.Application()
        Try

            ''*********************DATABASE CONNECTION *****************************************
            If IO.File.Exists(Application.StartupPath + "\DBINFO.INI") = False Then
                MsgBox("Check the Database Informations")
                Application.Exit()
                Exit Sub
            End If

            'Dim objConnect As New clsGlobalMethods
            If objConnect.Connection(user) = False Then
                MsgBox("Connecion Cound not be established")
                Exit Sub
            Else
                ' MemoryClr()
            End If

            ''**********************************************************************************
            objCompany = objSBOConnector.GetCompany(user)
            If user.ToString = "N" Then
                objCompanyMain = objCompany
            Else
                objCompanyMain = CompanyMainCon 'objSBOConnector.OBJMainCompanys()
            End If


            If objCompany.CompanyDB.ToString <> "" Then
                ' objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
                conns = True
            Else
                conns = False
                MsgBox("You are not connected to a company")
                Application.Exit()
            End If

            'LoadInitialize()
            'objstart1.Close()

        Catch ex As Exception
            'MsgBox(ex.ToString)
            MsgBox("You are not connected to a company")
            Exit Sub
        End Try
    End Sub
    Private Sub MemoryClr()
        'Dim objConnect1 As clsGlobalMethods
        Dim flagsed As Boolean = False
        strSql = "select (physical_memory_in_use_kb/1024) AS Memory_usedby_Sqlserver_MB From sys.dm_os_process_memory"
        Dim dted As New DataTable
        Dim dats As SqlDataAdapter
        Dim cmdes As SqlCommand

        dats = New SqlDataAdapter(strSql, con)
        dats.Fill(dted)

        If Val(dted.Rows(0).Item("Memory_usedby_Sqlserver_MB")) >= Val(13000) Then
            'Tran = con.BeginTransaction
            'strSql = " EXEC  sp_configure'max server memory (MB)',17000;"
            ''strSql += vbCrLf + " RECONFIGURE; "
            'cmdes = New SqlCommand(strSql, con, Tran)
            'cmdes.ExecuteNonQuery()
            'Tran.Commit()
            'Tran.Dispose()
            objConnect.MemClr(0)


            'strSql = " RECONFIGURE; "
            'cmdes = New SqlCommand(strSql, con)
            'cmdes.ExecuteNonQuery()
            flagsed = True
            System.Threading.Thread.Sleep(1000)
        End If

        If flagsed = True Then
            System.Threading.Thread.Sleep(1000)
            strSql = ""
            'Tran = con.BeginTransaction
            'strSql = " EXEC  sp_configure'max server memory (MB)',29000;"
            'cmd = New SqlCommand(strSql, con, Tran)
            'cmd.ExecuteNonQuery()
            'Tran.Commit()
            'Tran.Dispose()
            'strSql = " RECONFIGURE; "
            'cmd = New SqlCommand(strSql, con)
            'cmd.ExecuteNonQuery()
            objConnect.MemClr(1)
        End If
    End Sub
    Public Function OBJMainCompanys() As SAPbobsCOM.Company

        Dim lretcodes As Long
        Dim OBJMains As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
        'objGVar.MDBName = objMUtil.xDBName.ToString.Trim
        'objGVar.MUID = objMUtil.xUID.ToString.Trim
        'objGVar.MPWD = objMUtil.xPWD.ToString.Trim
        'objGVar.MSAPUID = objMUtil.xSAPUser.ToString.Trim
        'objGVar.MSAPPWd = objMUtil.xSAPPwd.ToString.Trim


        'strSql = "select lsrv ,@@SERVERNAME [servername] from [SBO-COMMON]..SLIC "
        'Dim dttemp As DataTable
        'dttemp = New DataTable
        'da = New SqlDataAdapter(strSql, con)
        'da.Fill(dttemp)
        OBJMainCompany.Server = OBJMains.xServerName.ToString.Trim
        OBJMainCompany.LicenseServer = OBJMains.xServerName.ToString.Trim + ":30000"
        'objCompany.Server = dttemp.Rows(0).Item("servername").ToString
        'objCompany.LicenseServer = dttemp.Rows(0).Item("lsrv").ToString
        objCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        objCompany.UseTrusted = False
        objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
        objCompany.DbUserName = OBJMains.xUID.ToString.Trim
        objCompany.DbPassword = OBJMains.xPWD.ToString.Trim
        objCompany.CompanyDB = MDBMainName.ToString.Trim
        objCompany.UserName = OBJMains.xSAPUser.ToString.Trim
        objCompany.Password = OBJMains.xSAPPwd.ToString.Trim


        lretcodes = objCompany.Connect()


        If lretcodes <> 0 Then
            'objCompany.GetLastError(lretcodes, retStr)
            'MsgBox("Error " & lretcodes & " " & retStr)
            'MsgBox(objCompany.GetLastErrorDescription)
            Exit Function
        Else
            objCompanyMain = OBJMainCompany
            Return OBJMainCompany
        End If

    End Function
    Private Function LoadLoginMenu() As Integer
        mcnt = objApplication.Menus.Item("43520").SubMenus.Item("3328").SubMenus.Count
        ogeneralmenu = CreateMenu(Application.StartupPath & "\VGN.bmp", mcnt, "AVR Login", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLLOGIN", objApplication.Menus.Item("43520").SubMenus.Item("3328"))

        Call CreateMenu("", mcnt, "Login", SAPbouiCOM.BoMenuType.mt_STRING, "Login", ogeneralmenu)
    End Function

    Private Sub CreateUDOs()
        Dim ct1(1) As String
        Dim Ct2(1) As String
        ''PURITY MASTER
        createUDOG("MIPLPM", "MIPURITY", "MIPLPM", SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        ''SIZE MASTER
        createUDOG("MIPLSM", "MIPLSM", "MIPLSM", SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        ''HALLMARK CHARGES MASTER
        createUDOG("MIPLHMC", "MIPLHMC", "MIPLHMC", SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        '' SALES CENT MASTER
        ct1(0) = "MIPLCM1"
        createUDOC("MIPLCM", "MIPLCM", "MIPLCM", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''Ct2(1) = "MIPLSPM1" ''SUBPRODUCT MASTER
        ''createUDOC("MIPLSPM", "MIPLSPM", "MIPLSPM", Ct2, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        ReDim ct1(1)
        ct1(0) = "MIPLPWM1"
        createUDOC("MIPLPWM", "MIPLPWM", "MIPLPWM", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Dim Ct3(1) As String
        Ct3(0) = "MIPLRM1"
        createUDOC("MIPLRM", "MIPLRM", "MIPLRM", Ct3, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Dim Ct4(1) As String
        Ct4(0) = "MIPLDM1"
        createUDOC("MIPLDM", "MIPLDM", "DISCOUNT MASTER", Ct4, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Dim Ct5(1) As String
        Ct5(0) = "MIPLIM1"
        createUDOC("MIPLIM", "MIPLIM", "MIPLIM", Ct5, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Dim Ct6(1) As String
        Ct6(0) = "MIPLSWM1"
        createUDOC("MIPLSWM", "MIPLSWM", "MIPLSWM", Ct6, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        '' PURCHASE CENT MASTER
        Dim ct7(1) As String
        ct7(0) = "MIPLPCM1"
        createUDOC("MIPLPCM", "MIPLPCM", "MIPLPCM", ct7, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''REPAIR WASTAGE MASTER
        Dim ct8(1) As String
        ct8(0) = "MIPLRWM1"
        createUDOC("MIPLRWM", "MIPLRWM", "MIPLRWM", ct8, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''INCENTIVE MASTER
        Dim ct9(1) As String
        createUDOC("MIPLICM", "MIPLICM", "MIPLICM", ct9, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''Special Sales Wastage Master
        Dim Ct10(1) As String
        Ct10(0) = "MIPLSSWM1"
        createUDOC("MIPLSSWM", "MIPLSSWM", "MIPLSSWM", Ct10, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''MetalType Master
        Dim Ct11(1) As String
        Ct11(0) = "MIPLMT"
        createUDOG("MIPLMT", "MIPLMT", "MIPLMT", SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''LOT CREATION (Ponparani)
        Dim Ct12(1) As String
        Ct12(0) = "MIPLLOT1"
        createUDOC("MIPLLOT", "MIPLLOT", "MIPLLOT", Ct12, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)

        ''OG RATE MASTER (Helen)
        Dim Ct13(1) As String
        Ct13(0) = "MIPLOGRM1"
        createUDOC("MIPLOGRM", "MIPLOGRM", "MIPLOGRM", Ct13, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ''REASON MASTER(Helen)
        createUDOG("MIPLREASON", "MIPLREASON", "MIPLREASON", SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        'OG Estimation (Tamizh)
        Dim ct(1) As String
        ct(0) = "MIPLOGL"
        createUDOC("MIPLOGH", "MIPLOGH", "OG Estimation", ct, SAPbobsCOM.BoUDOObjType.boud_Document)

    End Sub
    Function funcInsert() As Integer
        Dim objrs As SAPbobsCOM.Recordset
        Try
            ' Tran = con.BeginTransaction
            str = "  IF OBJECT_ID (N'dbo.Get_OLCTK', N'FN') IS NOT NULL"
            str += vbCrLf + " DROP FUNCTION Get_OLCTK;"
            str += vbCrLf + " GO "
            str += vbCrLf + " Create Function [dbo].[Get_OLCTK] (@Location As varchar(100)) Returns Varchar(100)"
            str += vbCrLf + " begin"
            str += vbCrLf + " if ISNUMERIC(@Location)=1"
            str += vbCrLf + " Return isnull((Select Location [Location] from [OLCT] where Code=@Location),'')"
            str += vbCrLf + " Return isnull((Select Code [Location] from [OLCT] where Location=@Location),'')	"
            str += vbCrLf + " End"
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(str)
            'cmd = New SqlCommand(str, con, Tran)
            'cmd.ExecuteNonQuery()
            'Tran.Commit()
            'Tran.Dispose()
        Catch ex As Exception
            If Not Tran Is Nothing Then
                'objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Tran.Rollback()
                Tran.Dispose()
            End If

        End Try
        Try
            Tran = con.BeginTransaction
            str = "  IF OBJECT_ID (N'dbo.MT_STWASTMAST', N'FN') IS NOT NULL"
            str += vbCrLf + " DROP FUNCTION MT_STWASTMAST;"
            str += vbCrLf + " GO"
            str += vbCrLf + " Create Function [dbo].[MT_STWASTMAST] (@ValidValue As varchar(10)) Returns Varchar(100)"
            str += vbCrLf + " begin"
            str += vbCrLf + " return (Select top 1 U_VALIDDESCR from [@VALIDVALUES] where U_PROCESS='STWASTMAST' and U_TYPE='MT' and U_VALIDVALUE=@ValidValue)"
            str += vbCrLf + " End"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteScalar()
            Tran.Commit()
            Tran.Dispose()
        Catch ex As Exception
            If Not Tran Is Nothing Then
                ' objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Tran.Rollback()
                Tran.Dispose()
            End If
        End Try
        Try
            Tran = con.BeginTransaction
            str = "  IF OBJECT_ID (N'dbo.WT_STWASTMAST', N'FN') IS NOT NULL"
            str += vbCrLf + " DROP FUNCTION WT_STWASTMAST;"
            str += vbCrLf + " GO"
            str += vbCrLf + " Create Function [dbo].[WT_STWASTMAST] (@ValidValue As varchar(10)) Returns Varchar(100)"
            str += vbCrLf + " begin"
            str += vbCrLf + " return (Select top 1 U_VALIDDESCR from [@VALIDVALUES] where U_PROCESS='STWASTMAST' and U_TYPE='WT' and U_VALIDVALUE=@ValidValue)"
            str += vbCrLf + " End"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()
            Tran.Commit()
            Tran.Dispose()
        Catch ex As Exception
            If Not Tran Is Nothing Then
                'objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Tran.Rollback()
                Tran.Dispose()
            End If
        End Try
        Try
            Tran = con.BeginTransaction
            str = ""
            str = "  IF OBJECT_ID (N'dbo.Get_MIPLMTK', N'FN') IS NOT NULL"
            str += vbCrLf + " DROP FUNCTION Get_MIPLMTK;"
            str += vbCrLf + " GO"
            str += vbCrLf + " Create Function [dbo].[Get_MIPLMTK] (@Metal As varchar(100)) Returns Varchar(100)"
            str += vbCrLf + " begin"
            str += vbCrLf + " if ISNUMERIC(@Metal)=1"
            str += vbCrLf + " Return isnull((Select U_METALTYPE [MetalName] from [@MIPLMT] where Code=@Metal),'')"
            str += vbCrLf + " Return isnull((Select Code [MetalCode] from [@MIPLMT] where U_METALTYPE=@Metal),'')	"
            str += vbCrLf + " End"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()
            Tran.Commit()
            Tran.Dispose()
        Catch ex As Exception
            If Not Tran Is Nothing Then
                'objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Tran.Rollback()
                Tran.Dispose()
            End If
        End Try
        Try
            Tran = con.BeginTransaction
            str = ""
            str = "  IF OBJECT_ID (N'dbo.Get_OWGTK', N'FN') IS NOT NULL"
            str += vbCrLf + " DROP FUNCTION Get_OWGTK;"
            str += vbCrLf + " GO"
            str += vbCrLf + " Create Function [dbo].[Get_OWGTK] (@UOM As varchar(100)) Returns Varchar(100)"
            str += vbCrLf + " begin"
            str += vbCrLf + " if ISNUMERIC(@UOM)=1"
            str += vbCrLf + " Return isnull((Select UnitName [UOMName] from [OWGT] where UnitCode=@UOM),'')"
            str += vbCrLf + " Return isnull((Select UnitCode [UOMID] from [OWGT] where UnitName=@UOM),'')	"
            str += vbCrLf + " End"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()
            Tran.Commit()
            Tran.Dispose()
        Catch ex As Exception
            If Not Tran Is Nothing Then
                objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Tran.Rollback()
                Tran.Dispose()
            End If
        End Try
    End Function
    Private Sub createObjects()
        'Library Object Initilisation
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)

        objrate1 = New frmMRate
        objratebranchwise = New frmMRateMasterBranchWise
        objswastage = New frmMSalesWastage
        objProMast = New frmMProduct
        objSubItemMast = New frmMSubProduct
        objDiscMast = New frmMDiscount
        objGenSet = New frmGeneralSettings
        objMDI = New MDI_JEWELADDON

        objPurInv = New frmDPurchaseInvoice
        objGReceipt = New frmDMaterialReceipt
        objLot = New frmDLotCreation
        objTag = New frmDTagGeneration
        objsizevb = New frmMSize
        objPurityVB = New frmMPurity
        objHallmarkVB = New frmMHallmarkCharges
        objSaleCentVB = New frmMCentSale
        objpurcent = New frmMCentPurchase
        objPwvb = New frmMPurchaseWastage
        objIncen = New frmMIncentive
        objRwm = New frmMRepairWastage
        objSwast = New frmMSpecialSales
        objOGRate = New frmMOGCashRate
        objMetalType = New frmMMetalType
        objReason = New frmMReason
        objOGEst = New frmOGEstimation
        objOGCashPayment = New frmDOGCashPayment
        objcustsmaple = New frmdCustomerSample
        objSalEst = New frmDSalesEstimation
        objGI = New frmDGoodsIssue
        objGR = New frmDGoodsReturn

        objsalesinvoice = New frmDSalesInvoice
        objob = New frmDOrderBooking
        objrmr = New frmDRepairMaterialReceipt
        objrmi = New frmDRepairMaterialIssue
        objwo = New frmDWorkOrder
        objgsissue = New frmDGoldSmithIssue
        objDGSReturn = New frmdGoldSmithReturn
        ''CHit
        ochit = New frmDChitCreation
        objscheme = New frmMSchemeMaster
        objgrp = New frmMGroupMaster
        objreverse = New frmDChequeReverse
        objchitmaturity = New frmDChitMaturity_Closing

        objAuth = New frmAuthorization
        objLogin = New frmUserNamePassword
        objLct = New clsLctDup_Emp
    End Sub

    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                'Case clsPurity.formtype
                '    objPurity.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "150"
                    'objItemMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case clssize.Formtype
                    '    objsize.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case Clscent.FormType
                    '    objcent.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case ClsPurchase_Wastage.Formtype
                    '    objPur.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case clsHallmarkCharges.FormType
                    '    objHallmarkcharges.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case clsRate.FormType
                    '    objrate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "943"
                    objLct.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
        End Try
    End Sub

    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent

        'If BusinessObjectInfo.FormTypeEx = "MIPLFree2" Then
        '    objFree2.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        'End If
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function

    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
        Else
            Select Case pVal.MenuUID

                Case frmMProduct.formType
                    objProMast.BrowseFolderDialog()
                Case frmMSubProduct.formType
                    objSubItemMast.BrowseFolderDialog()
                Case frmMDiscount.formType
                    objDiscMast.BrowseFolderDialog()
                Case frmGeneralSettings.formType
                    objGenSet.BrowseFolderDialog()
                Case frmDPurchaseInvoice.formType
                    objPurInv.BrowseFolderDialog()
                Case frmDMaterialReceipt.formType
                    objGReceipt.BrowseFolderDialog()
                Case frmDLotCreation.formType
                    objLot.BrowseFolderDialog()
                Case frmDTagGeneration.formType
                    objTag.BrowseFolderDialog()
                Case MDI_JEWELADDON.formType
                    objMDI.BrowseFolderDialog()
                Case frmMRate.formtype
                    objrate1.BrowseFolderDialog()
                Case frmMRateMasterBranchWise.formType
                    objratebranchwise.BrowseFolderDialog()
                Case frmMSize.formType
                    objsizevb.BrowseFOlderDialog()
                Case frmMPurity.formtype
                    objPurityVB.BrowseFolderDialog()
                Case frmMHallmarkCharges.formType
                    objHallmarkVB.BrowseFolderDialog()
                Case frmMCentSale.formType
                    objSaleCentVB.BrowseFolderDialog()
                Case frmMCentPurchase.formType
                    objpurcent.BrowseFolderDialog()
                Case frmMPurchaseWastage.formType
                    objPwvb.BrowseFOlderDialog()
                Case frmMIncentive.formType
                    objIncen.BrowseFolderDialog()
                Case frmMRepairWastage.formType
                    objRwm.BrowseFOlderDialog()
                Case frmMSpecialSales.formtype
                    ' MsgBox("Special Sales Wastage Master Is Under Construction", MsgBoxStyle.Information, MIPL_Msgbox_Title1)
                    objSwast.BrowseFolderDialog()
                Case frmMSalesWastage.formType
                    objswastage.BrowseFolderDialog()
                Case frmMOGCashRate.formType
                    objOGRate.BrowseFolderDialog()
                Case frmMMetalType.formtype
                    ''Dim objUserPass As New frmUserNamePassword("frmMMetalType")
                    ''objUserPass.BrowseFolderDialog()
                    'If Not MIPL_Authorize Then Exit Select
                    objMetalType.BrowseFolderDialog()
                Case frmAuthorization.formType
                    objAuth.BrowseFolderDialog()
                Case frmMReason.formtype
                    objReason.BrowseFolderDialog()
                Case frmUserNamePassword.formType
                    objLogin.BrowseFolderDialog()
                Case frmdCustomerSample.formType
                    objcustsmaple.BrowseFolderDialog()
                Case frmDWorkOrder.formtype
                    objwo.BrowseFOlderDialog()
                    ''Chit Master
                Case frmMSchemeMaster.formType
                    objscheme.BrowseFolderDialog()
                Case frmMGroupMaster.formType
                    objgrp.BrowseFolderDialog()
                    ''Chit Document
                Case frmDChequeReverse.formtype
                    objreverse.BrowseFolderDialog()
                Case frmDChitCreation.formType
                    ochit.BrowseFolderDialog()
                Case frmDChitMaturity_Closing.formType
                    objchitmaturity.BrowseFolderDialog()

                Case "MIPLOGDoc"
                    Try
                        Dim menus As SAPbouiCOM.Menus
                        menus = objAddOn.objApplication.Menus.Item("51200").SubMenus
                        Dim i As Integer
                        For i = 0 To menus.Count - 1
                            'MsgBox(menus.Item(i).String)
                            If menus.Item(i).String.Contains("MIPLOGDOC") Then
                                objAddOn.objApplication.ActivateMenuItem(menus.Item(i).UID)
                                Exit For
                            End If
                        Next i
                    Catch ex As Exception
                    End Try
                Case "MIPLCAT"
                    Try
                        Dim menus As SAPbouiCOM.Menus
                        menus = objAddOn.objApplication.Menus.Item("51200").SubMenus
                        Dim i As Integer
                        For i = 0 To menus.Count - 1
                            'MsgBox(menus.Item(i).String)
                            If menus.Item(i).String.Contains("MIPLCAT") Then
                                objAddOn.objApplication.ActivateMenuItem(menus.Item(i).UID)
                                Exit For
                            End If
                        Next i
                    Catch ex As Exception
                    End Try
                Case frmDOGCashPayment.formtype
                    objOGCashPayment.BrowseFolderDialog()
                Case frmOGEstimation.formtype
                    objOGEst.BrowseFolderDialog()
                Case frmDSalesEstimation.formtype
                    objSalEst.BrowseFolderDialog()
                Case frmDGoodsIssue.formtype
                    objGI.BrowseFolderDialog()
                Case frmDGoodsReturn.formType
                    objGR.BrowseFolderDialog()

                Case frmDSalesInvoice.formtype
                    'objsalesinvoice.BrowseFolderDialog()
                Case frmDOrderBooking.formtype
                    objob.BrowseFolderDialog()
                Case frmDRepairMaterialReceipt.formtype
                    objrmr.BrowseFolderDialog()
                Case frmDRepairMaterialIssue.formtype
                    objrmi.BrowseFOlderDialog()


            End Select
            If pVal.MenuUID = "1281" Then
                Select Case objAddOn.objApplication.Forms.ActiveForm.TypeEx
                    'Case clssize.Formtype
                    '    objsize.MenuEvent(pVal, BubbleEvent)
                    'Case clsPurity.formtype
                    '    objPurity.MenuEvent(pVal, BubbleEvent)
                    'Case Clscent.FormType
                    '    objcent.MenuEvent(pVal, BubbleEvent)
                    'Case ClsPurchase_Wastage.Formtype
                    '    objPur.MenuEvent(pVal, BubbleEvent)
                    'Case clsHallmarkCharges.FormType
                    '    objHallmarkcharges.MenuEvent(pVal, BubbleEvent)
                    'Case clsRate.FormType
                    '    objrate.MenuEvent(pVal, BubbleEvent)
                End Select
            End If
            If pVal.MenuUID = "1282" Then
                Select Case objAddOn.objApplication.Forms.ActiveForm.TypeEx
                    'Case clssize.Formtype
                    '    objsize.MenuEvent(pVal, BubbleEvent)
                    'Case clsPurity.formtype
                    '    objPurity.MenuEvent(pVal, BubbleEvent)
                    'Case Clscent.FormType
                    '    objcent.MenuEvent(pVal, BubbleEvent)
                    'Case ClsPurchase_Wastage.Formtype
                    '    objPur.MenuEvent(pVal, BubbleEvent)
                    'Case clsHallmarkCharges.FormType
                    '    objHallmarkcharges.MenuEvent(pVal, BubbleEvent)
                    'Case clsRate.FormType
                    '    objrate.MenuEvent(pVal, BubbleEvent)
                End Select
                If pVal.MenuUID = "1293" Then
                    Select Case objAddOn.objApplication.Forms.ActiveForm.TypeEx
                        'Case Clscent.FormType
                        '    objcent.MenuEvent(pVal, BubbleEvent)
                    End Select

                End If
            End If
        End If
    End Sub
    Public Sub loadMenu()
        Dim SubMenu_Master
        Dim SubMenu_Masterchit
        Dim SubMenu_Purchase
        Dim SubMenu_Transaction
        Dim SubMenu_TransactionChit
        Dim SubMenu_LOT
        Dim SubMenu_Sales
        Dim SubMenu_OrdRep
        Dim SubMenu_Utility
        Dim SubMenu_MDI
        Dim SubMenu_OG
        ''If objApplication.Menus.Item("43520").SubMenus.Exists("MIPLAVR") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count
        ogeneralmenu = CreateMenu(Application.StartupPath & "\VGN.bmp", MenuCount, "AVR JEWELLERY", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLAVR", objApplication.Menus.Item("43520"))

        SubMenu_Master = CreateMenu("", 1, "Masters", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLMAS", objApplication.Menus.Item("MIPLAVR"))

        ''MASTERS submenu 
        Call CreateMenu("", MenuCount, "MetalType Master", SAPbouiCOM.BoMenuType.mt_STRING, "MetalType", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Purity Master", SAPbouiCOM.BoMenuType.mt_STRING, "PurityVB", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Discount Master", SAPbouiCOM.BoMenuType.mt_STRING, "Discount", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Hallmark Charges", SAPbouiCOM.BoMenuType.mt_STRING, "HallmarkCharges", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Product Master", SAPbouiCOM.BoMenuType.mt_STRING, "Product", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Sub Product Master", SAPbouiCOM.BoMenuType.mt_STRING, "SubProduct", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Size Master", SAPbouiCOM.BoMenuType.mt_STRING, "Sizevb", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Rate Master", SAPbouiCOM.BoMenuType.mt_STRING, "Rate1", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Rate Master Branch Wise", SAPbouiCOM.BoMenuType.mt_STRING, "RateBranchWise", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Purchase Cent Master", SAPbouiCOM.BoMenuType.mt_STRING, "purchasecentmaster", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Sales Cent Master", SAPbouiCOM.BoMenuType.mt_STRING, "centmastersale", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Purchase Wastage Master", SAPbouiCOM.BoMenuType.mt_STRING, "PurchaseWastagevb", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Repair Wastage Master", SAPbouiCOM.BoMenuType.mt_STRING, "RepairWastage", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Sales Wastage Master", SAPbouiCOM.BoMenuType.mt_STRING, "SalesWastage", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Special Sales Wastage Master", SAPbouiCOM.BoMenuType.mt_STRING, "SpecialSalesWastage", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Incentive Master", SAPbouiCOM.BoMenuType.mt_STRING, "Incentive", SubMenu_Master)
        'Call CreateMenu("", MenuCount, "OG Rate Master", SAPbouiCOM.BoMenuType.mt_STRING, "OGRate", SubMenu_Master)
        Call CreateMenu("", MenuCount, "Reason Master", SAPbouiCOM.BoMenuType.mt_STRING, "Reason", SubMenu_Master)

        ''Transaction Menu
        SubMenu_Transaction = CreateMenu("", 2, "Transaction", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLTRNS", objApplication.Menus.Item("MIPLAVR"))
        ''PURCHASE SUBMENU
        'SubMenu_Purchase = CreateMenu("", 3, "Purchase", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLPUR", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", MenuCount, "Purchase Invoice", SAPbouiCOM.BoMenuType.mt_STRING, "APInvoice", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Material Receipt", SAPbouiCOM.BoMenuType.mt_STRING, "GoodsReceipt", SubMenu_Transaction)
        ''LOT/TAG SUBMENU
        'SubMenu_LOT = CreateMenu("", 4, "Lot/Tag", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLLOT", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", MenuCount, "LOT Creation", SAPbouiCOM.BoMenuType.mt_STRING, "LOT", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "TAG Generation", SAPbouiCOM.BoMenuType.mt_STRING, "TAG", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Work Order", SAPbouiCOM.BoMenuType.mt_STRING, "WorkOrder", SubMenu_Transaction)
        ''SALES SUBMENU
        'SubMenu_Sales = CreateMenu("", 5, "Sales", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLSALES", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", MenuCount, "Sales Estimation", SAPbouiCOM.BoMenuType.mt_STRING, "SalesEst", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Order Booking", SAPbouiCOM.BoMenuType.mt_STRING, "OrderBooking", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Goods Issue", SAPbouiCOM.BoMenuType.mt_STRING, "GoodsIssue", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Goods Return", SAPbouiCOM.BoMenuType.mt_STRING, "GoodsReturn", SubMenu_Transaction)

        Call CreateMenu("", MenuCount, "Sales Invoice", SAPbouiCOM.BoMenuType.mt_STRING, "Sales Invoice", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Repair Material Receipt", SAPbouiCOM.BoMenuType.mt_STRING, "RepairMaterialReceipt", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Repair Material Delivery", SAPbouiCOM.BoMenuType.mt_STRING, "RepairMaterialIssue", SubMenu_Transaction)
        Call CreateMenu("", MenuCount, "Customer Sample", SAPbouiCOM.BoMenuType.mt_STRING, "CustomerSample", SubMenu_Transaction)
        '' ''ORDER/REPAIR SUBMENU
        ''SubMenu_OrdRep = CreateMenu("", 6, "Order/Repair", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLORDREP", objApplication.Menus.Item("MIPLAVR"))
        ''Call CreateMenu("", MenuCount, "Order Booking", SAPbouiCOM.BoMenuType.mt_STRING, "OrderBook", SubMenu_OrdRep)
        ''UTILITIES SUBMENU
        SubMenu_Utility = CreateMenu("", 6, "Utilities", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLOTH", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", MenuCount, "Addon General Setting", SAPbouiCOM.BoMenuType.mt_STRING, "AddonGenSetting", SubMenu_Utility)
        Call CreateMenu("", MenuCount, "Authorization", SAPbouiCOM.BoMenuType.mt_STRING, "Authorize", SubMenu_Utility)

        SubMenu_MDI = CreateMenu("", 7, "MDI Testing", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLMDI", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", MenuCount, "MDI Form", SAPbouiCOM.BoMenuType.mt_STRING, "MDITest", SubMenu_MDI)

        SubMenu_OG = CreateMenu("", 8, "OG", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLOG", objApplication.Menus.Item("MIPLAVR"))
        Call CreateMenu("", 0, "OG Category", SAPbouiCOM.BoMenuType.mt_STRING, "MIPLCAT", SubMenu_OG)
        Call CreateMenu("", 1, "OG Document Setting", SAPbouiCOM.BoMenuType.mt_STRING, "MIPLOGDoc", SubMenu_OG)
        Call CreateMenu("", 2, "OG Estimation", SAPbouiCOM.BoMenuType.mt_STRING, "MIPLOGEst", SubMenu_OG)
        Call CreateMenu("", 3, "OG Cash Payment", SAPbouiCOM.BoMenuType.mt_STRING, "MIPLOGCash", SubMenu_OG)

        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count
        ogeneralmenuchit = CreateMenu(Application.StartupPath & "\VGN.bmp", MenuCount, "AVR CHIT", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLAVRCHIT", objApplication.Menus.Item("43520"))
        SubMenu_Masterchit = CreateMenu("", 1, "Chit Masters", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLCHITMAS", objApplication.Menus.Item("MIPLAVRCHIT"))
        ''Chit Master
        Call CreateMenu("", MenuCount, "Chit Scheme Master", SAPbouiCOM.BoMenuType.mt_STRING, "schememaster", SubMenu_Masterchit)
        Call CreateMenu("", MenuCount, "Chit Group Master", SAPbouiCOM.BoMenuType.mt_STRING, "grpmaster", SubMenu_Masterchit)
        SubMenu_TransactionChit = CreateMenu("", 2, "Chit Transaction", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLCHITTRNS", objApplication.Menus.Item("MIPLAVRCHIT"))
        Call CreateMenu("", MenuCount, "Chit Creation", SAPbouiCOM.BoMenuType.mt_STRING, "CHIT", SubMenu_TransactionChit)
        Call CreateMenu("", MenuCount, "Cheque Reverse", SAPbouiCOM.BoMenuType.mt_STRING, "Cheque_Reverse", SubMenu_TransactionChit)
        Call CreateMenu("", MenuCount, "Chit Maturity", SAPbouiCOM.BoMenuType.mt_STRING, "ChitMaturity", SubMenu_TransactionChit)
        ''objLogin.btnCancel_Click(Me, New EventArgs())
        ''RemoveMenu("", mcnt, "Login", SAPbouiCOM.BoMenuType.mt_STRING, "Login", ogeneralmenu)
        ''RemoveMenu(Application.StartupPath & "\VGN.bmp", mcnt, "AVR Login", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLLOGIN", objApplication.Menus.Item("43520").SubMenus.Item("3328"))
    End Sub

    ' For Menu Creation
    Public Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function

    Public Function RemoveMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.RemoveEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Not Removed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            MsgBox(objAddOn.objCompany.GetLastErrorDescription)
        End Try
    End Function

    Private Sub createTables()
        Try
            Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
            objAddOn.objApplication.SetStatusBarMessage("Checking Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objUDFEngine.CreateTable("MIPLPM", "Purity Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLPM", "MetalType", "MetalType", 30)
            objUDFEngine.AddAlphaField("MIPLPM", "PurityID", "PurityID", 20)
            objUDFEngine.AddAlphaField("MIPLPM", "Description", "Description", 100)
            objUDFEngine.AddFloatField("MIPLPM", "Purity", "Purity", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddAlphaField("MIPLPM", "Basecode", "BaseCode", 20)
            objUDFEngine.AddFloatField("MIPLPM", "Amtinper", "AmountinPer", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLPM", "AmtinRs", "AmountinRupees", SAPbobsCOM.BoFldSubTypes.st_Price)

            ''Script to Create UDF in Item Master
            objUDFEngine.AddAlphaField("OITM", "ProdCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("OITM", "ProdName", "Product Name", 100)
            objUDFEngine.AddAlphaField("OITM", "MetalType", "MetalType", 16)
            objUDFEngine.AddAlphaField("OITM", "Category", "Category", 16)
            objUDFEngine.AddAlphaField("OITM", "StockType", "StockType", 16) '9 [T-Taged, N-NonTaged]
            objUDFEngine.AddAlphaField("OITM", "SalesMode", "SalesMode", 16) ''[G-Gross Weight, N-Net Weight, R-Rate, F-Fixed]
            objUDFEngine.AddAlphaField("OITM", "PurchaseMode", "PurchaseMode", 16) ''[G-Gross Weight, N-Net Weight, R-Rate]
            objUDFEngine.AddAlphaField("OITM", "Brand", "Brand", 30)
            objUDFEngine.AddAlphaField("OITM", "MCType", "MCType", 16)
            objUDFEngine.AddAlphaField("OITM", "OGRate", "OGRate", 16)
            objUDFEngine.AddAlphaField("OITM", "TaxCode", "TaxCode", 30)
            objUDFEngine.AddNumericField("OITM", "DefPcs", "DefPcs", 4)
            objUDFEngine.AddAlphaField("OITM", "PurityID", "PurityID", 20)
            objUDFEngine.AddAlphaField("OITM", "Description", "Description", 100)
            objUDFEngine.AddFloatField("OITM", "Purity", "Purity", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddAlphaField("OITM", "DiscountID", "Discount ID", 20)
            objUDFEngine.AddAlphaField("OITM", "HallmarkID", "Hallmark ID", 20)
            objUDFEngine.AddAlphaField("OITM", "Size", "Size", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "Stones", "Stones", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "Discount", "Discount", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "HallmkChrge", "Hallmark Charges", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "Wastage", "Wastage", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "OtherChrge", "Other Charges", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "BRateDiff", "Billing RateDiff", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "BWastage", "Billing Wastage", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "LessTaxOnRate", "Less Tax On Rate", 1) ' [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("OITM", "MultiMetal", "Multi Metal", 1) ' [Y-Yes, N-No]
            'End of Item Master Script

            '' Employee Master UDF (Chandra)   
            objUDFEngine.AddAlphaField("OSLP", "Section", "Section", 20)

            'Size master
            objUDFEngine.CreateTable("MIPLSM", "Size Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@MIPLSM", "Prodcode", "Product code", "20")
            objUDFEngine.AddAlphaField("@MIPLSM", "ProdName", "Product Name", "100")
            objUDFEngine.AddAlphaField("@MIPLSM", "Sizecode", "Size code", "20")
            objUDFEngine.AddAlphaField("@MIPLSM", "SizeName", "Size Name", "100")

            'Cent Master
            objUDFEngine.CreateTable("MIPLCM", "Cent Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@MIPLCM", "Prodcode", "Product code", "20")
            objUDFEngine.AddAlphaField("@MIPLCM", "SubProdcode", "Sub Product code", 20)
            objUDFEngine.AddAlphaField("@MIPLCM", "ProdName", "Product Name", 100)
            objUDFEngine.AddAlphaField("@MIPLCM", "SubProdName", "Sub Product Name", 100)
            objUDFEngine.AddAlphaField("@MIPLCM", "location", "Branch Location", 200)
            objUDFEngine.AddNumericField("@MIPLCM", "locationcode", "Branch code", 4)

            objUDFEngine.CreateTable("MIPLCM1", "Cent Master Row", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@MIPLCM1", "FromCent", "FromCent", 20)
            objUDFEngine.AddAlphaField("@MIPLCM1", "ToCent", "ToCent", 20)
            objUDFEngine.AddFloatField("@MIPLCM1", "RatePerCent", "Rate Per cent", SAPbobsCOM.BoFldSubTypes.st_Price)

            ''LOT CREATION
            objUDFEngine.CreateTable("MIPLLOT", "LOT", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("MIPLLOT", "series", "series", 30)
            objUDFEngine.AddAlphaField("MIPLLOT", "lotno", "Lot No", 30)
            objUDFEngine.AddDateField("MIPLLOT", "ldate", "Lot Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLLOT", "Vendorcode", "Vendor Code", 30)
            objUDFEngine.AddAlphaField("MIPLLOT", "Vendorname", "Vendor Name", 100)
            objUDFEngine.AddAlphaField("MIPLLOT", "Vendorshortname", "Vendor Short Name", 100)
            objUDFEngine.AddAlphaField("MIPLLOT", "grno", "GrNo", 30)
            objUDFEngine.AddDateField("MIPLLOT", "gdate", "GrDate", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLLOT", "crmemo", "Memo", 30)
            objUDFEngine.AddDateField("MIPLLOT", "memodate", "MemoDate", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLLOT", "trayno", "TrayNo", 30)
            objUDFEngine.AddAlphaField("MIPLLOT", "tagstatus", "TagStatus", 10)
            'objUDFEngine.AddNumericField("MIPLLOT", "orderno", "Order No", 10)

            '' LOT CHILD TABLE CREATION
            objUDFEngine.CreateTable("MIPLLOT1", "LOT Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddNumericField("MIPLLOT1", "sno", "sno", 20)
            objUDFEngine.AddAlphaField("MIPLLOT1", "spcode", "SpCode", 30)
            objUDFEngine.AddAlphaField("MIPLLOT1", "spname", "SpName", 100)
            objUDFEngine.AddAlphaField("MIPLLOT1", "pcode", "PCode", 30)
            objUDFEngine.AddAlphaField("MIPLLOT1", "pname", "PName", 100)
            objUDFEngine.AddAlphaField("MIPLLOT1", "uom", "Uom", 30)
            objUDFEngine.AddFloatField("MIPLLOT1", "pieces", "Pieces", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLLOT1", "grweight", "Grweight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLLOT1", "stweight", "Stweight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLLOT1", "netweight", "Netweight", SAPbobsCOM.BoFldSubTypes.st_Measurement)

            objUDFEngine.AddAlphaField("MIPLLOT1", "councode", "Councode", 20)
            objUDFEngine.AddAlphaField("MIPLLOT1", "counname", "Counname", 100)
            objUDFEngine.AddAlphaField("MIPLLOT1", "branch", "Branch", 100)



            'Purchase Watage Master (Praveen)
            'Header Table
            objUDFEngine.CreateTable("MIPLPWM", "Purchase Wastage Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLPWM", "VCode", "Vendor Code", 15)
            objUDFEngine.AddAlphaField("MIPLPWM", "VName", "Vendor Name", 100)
            objUDFEngine.AddAlphaField("MIPLPWM", "Wastagetype", "Wastage Type", 5)
            objUDFEngine.AddFloatField("MIPLPWM", "Wastagefamt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLPWM", "Wastagefgrams", "Wastage Grams", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLPWM", "Wastageftouch", "Wastage Touch", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLPWM", "WastagefPercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddAlphaField("MIPLPWM", "Makingchrgtype", "Making Charges Type", 5)
            objUDFEngine.AddFloatField("MIPLPWM", "Makingchrg", "Making Charges", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("MIPLPWM", "BranchName", "BranchName", 40)
            objUDFEngine.AddAlphaField("MIPLPWM", "BranchCode", "BranchCode", 30)
            objUDFEngine.AddAlphaField("MIPLPWM", "Active", "Active", 1)
            objUDFEngine.AddDateField("MIPLPWM", "Activefromdate", "Active From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLPWM", "Activetodate", "Active To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLPWM", "Activeremarks", "Active Remarks", 100)
            objUDFEngine.AddAlphaField("MIPLPWM", "Inactive", "InActive", 1)
            objUDFEngine.AddDateField("MIPLPWM", "Inactivefromdate", "InActive From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLPWM", "Inactivetodate", "InActive To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLPWM", "Inactiveremarks", "InActive Remarks", 100)
            'Line table
            objUDFEngine.CreateTable("MIPLPWM1", "Purchase Wastage maste Line", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("MIPLPWM1", "PCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLPWM1", "PName", "Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLPWM1", "SPCode", "Sub Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLPWM1", "SPName", "Sub Product Name", 100)
            objUDFEngine.AddFloatField("MIPLPWM1", "fromweight", "From Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLPWM1", "toweight", "To Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddAlphaField("MIPLPWM1", "size", "Size", 100)
            objUDFEngine.AddFloatField("MIPLPWM1", "Wastageamt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLPWM1", "Wastagegrams", "Wastage Grams", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLPWM1", "Wastagetouch", "Wastage Touch", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLPWM1", "WastagePercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLPWM1", "StoneChargePge", "Stone Charge Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLPWM1", "StoneChargeAmt", "Stone Charge Amount", SAPbobsCOM.BoFldSubTypes.st_Price)


            'Repair Watage Master (Praveen)
            'Header Table
            objUDFEngine.CreateTable("MIPLRWM", "Repair Wastage Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLRWM", "VCode", "Vendor Code", 15)
            objUDFEngine.AddAlphaField("MIPLRWM", "VName", "Vendor Name", 100)
            objUDFEngine.AddAlphaField("MIPLRWM", "Wastagetype", "Wastage Type", 5)
            objUDFEngine.AddFloatField("MIPLRWM", "Wastagefamt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLRWM", "Wastagefgrams", "Wastage Grams", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLRWM", "Wastageftouch", "Wastage Touch", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLRWM", "WastagefPercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddAlphaField("MIPLRWM", "Makingchrgtype", "Making Charges Type", 5)
            objUDFEngine.AddFloatField("MIPLRWM", "Makingchrg", "Making Charges", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("MIPLRWM", "BranchName", "BranchName", 40)
            objUDFEngine.AddAlphaField("MIPLRWM", "BranchCode", "BranchCode", 30)
            objUDFEngine.AddAlphaField("MIPLRWM", "Active", "Active", 1)
            objUDFEngine.AddDateField("MIPLRWM", "Activefromdate", "Active From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLRWM", "Activetodate", "Active To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLRWM", "Activeremarks", "Active Remarks", 100)
            objUDFEngine.AddAlphaField("MIPLRWM", "Inactive", "InActive", 1)
            objUDFEngine.AddDateField("MIPLRWM", "Inactivefromdate", "InActive From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLRWM", "Inactivetodate", "InActive To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLRWM", "Inactiveremarks", "InActive Remarks", 100)
            'Line table
            objUDFEngine.CreateTable("MIPLRWM1", "Repair Wastage maste Line", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("MIPLRWM1", "PCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLRWM1", "PName", "Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLRWM1", "SPCode", "Sub Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLRWM1", "SPName", "Sub Product Name", 100)
            objUDFEngine.AddFloatField("MIPLRWM1", "fromweight", "From Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLRWM1", "toweight", "To Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddAlphaField("MIPLRWM1", "size", "Size", 100)
            objUDFEngine.AddFloatField("MIPLRWM1", "Wastageamt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLRWM1", "Wastagegrams", "Wastage Grams", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLRWM1", "Wastagetouch", "Wastage Touch", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLRWM1", "WastagePercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLRWM1", "StoneChargePge", "Stone Charge Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLRWM1", "StoneChargeAmt", "Stone Charge Amount", SAPbobsCOM.BoFldSubTypes.st_Price)

            'HallMarlCharges (Chandra)
            objUDFEngine.CreateTable("MIPLHMC", "HallmarkCharge", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLHMC", "HallmarkId", "Hallmark ID", 100)
            objUDFEngine.AddAlphaField("MIPLHMC", "HallmarkType", "HallmarkTyp", 20)
            objUDFEngine.AddFloatField("MIPLHMC", "HallmarkPer", "Hallmark Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLHMC", "HallmarkAmt", "Hallmark Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLHMC", "HallmarkRate", "HallmarkRt", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("MIPLHMC", "Remarks", "Remarks", 254)
            objUDFEngine.AddAlphaField("MIPLHMC", "BranchCode", "Branch Code", 30)

            'Rate Master (Praveen)
            objUDFEngine.CreateTable("MIPLRM", "Rate Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddDateField("@MIPLRM", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@MIPLRM", "Time", "Time", SAPbobsCOM.BoFldSubTypes.st_Time)
            objUDFEngine.AddAlphaField("@MIPLRM", "Time1", "Time1", 5)
            objUDFEngine.AddAlphaField("@MIPLRM", "branchname", "Branch Name", 100)
            objUDFEngine.AddAlphaField("@MIPLRM", "branchcode", "Branch Code", 20)
            'Rate Master Line Master
            objUDFEngine.CreateTable("MIPLRM1", "Rate Master Line Table", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@MIPLRM1", "PurityCode", "Purity Code", 20)
            objUDFEngine.AddAlphaField("@MIPLRM1", "Description", "Purity Description", 100)
            objUDFEngine.AddFloatField("@MIPLRM1", "Rate", "Rate", SAPbobsCOM.BoFldSubTypes.st_Rate)


            ''purchase cent master(PARANI)
            objUDFEngine.CreateTable("MIPLPCM", "Purchase cent master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@MIPLPCM", "vencode", "Vendor Code", 15)
            objUDFEngine.AddAlphaField("@MIPLPCM", "venname", "Vendor Name", 100)
            objUDFEngine.AddAlphaField("@MIPLPCM", "Prodcode", "Product code", "20")
            objUDFEngine.AddAlphaField("@MIPLPCM", "SubProdcode", "Sub Product code", 20)
            objUDFEngine.AddAlphaField("@MIPLPCM", "ProdName", "Product Name", 100)
            objUDFEngine.AddAlphaField("@MIPLPCM", "SubProdName", "Sub Product Name", 100)
            objUDFEngine.AddAlphaField("@MIPLPCM", "location", "Branch Location", 200)
            objUDFEngine.AddNumericField("@MIPLPCM", "locationcode", "Branch code", 4)

            ''Purchase Cent Master Line
            objUDFEngine.CreateTable("MIPLPCM1", "Purchase Cent Master row", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@MIPLPCM1", "FromCent", "FromCent", 20)
            objUDFEngine.AddAlphaField("@MIPLPCM1", "ToCent", "ToCent", 20)
            objUDFEngine.AddFloatField("@MIPLPCM1", "RatePerCent", "Rate Per cent", SAPbobsCOM.BoFldSubTypes.st_Price)

            'Incentive Master  (Chandra)
            objUDFEngine.CreateTable("MIPLICM", "Incentive", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@MIPLICM", "IncentiveID", "Incentive ID", 100)
            objUDFEngine.AddAlphaField("@MIPLICM", "SalesEmployee", "SalesEmployee", 20)
            objUDFEngine.AddAlphaField("@MIPLICM", "Section", "Section", 20)
            objUDFEngine.AddAlphaField("@MIPLICM", "SalesEmpID", "SalesEmpID", 20)
            objUDFEngine.AddAlphaField("@MIPLICM", "IncentiveType", "Incentive Type", 20)
            objUDFEngine.AddAlphaField("@MIPLICM", "Quantity", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@MIPLICM", "IncentiveMode", "IncMode", 20)
            objUDFEngine.AddAlphaField("@MIPLICM", "IncentiveAmount", "IncAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("@MIPLICM", "IncentivePer", "IncPer", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddDateField("@MIPLICM", "FromDate", "FromDt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@MIPLICM", "ToDate", "ToDt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@MIPLICM", "ActiveStatus", "ActStatus", 20)
            objUDFEngine.AddDateField("@MIPLICM", "ActiveFrDt", "Actfrdt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@MIPLICM", "ActiveToDt", "Acttodt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@MIPLICM", "ActiveRemarks", "ActRemarks", 254)
            objUDFEngine.AddAlphaField("@MIPLICM", "InActiveStatus", "InActStatus", 20)
            objUDFEngine.AddDateField("@MIPLICM", "InActiveFrDt", "InActfrdt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@MIPLICM", "InActiveToDt", "InActtodt", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@MIPLICM", "InActiveRemarks", "Remarks", 254)

            ''RAJESH**********************************************************************************************************
            ''CFL SETTING NO OBJECT TABLE(Common Table)[UDF]
            objUDFEngine.CreateTable("MICFLSET", "CFL Setting Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUDFEngine.AddNumericField("MICFLSET", "Sno", "Serial No", 10)
            objUDFEngine.AddAlphaField("MICFLSET", "FieldName", "Field Name", 100)
            objUDFEngine.AddAlphaField("MICFLSET", "FieldDesc", "Field Description", 100)
            objUDFEngine.AddAlphaField("MICFLSET", "TableName", "Table Name", 100)
            objUDFEngine.AddAlphaField("MICFLSET", "DefField", "Default Field", 1)
            objUDFEngine.AddAlphaField("MICFLSET", "ActField", "Active Field", 1)
            objUDFEngine.addField("MICFLSET", "AscDesc", "Ascending or Descending", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")


            ''DISCOUNT MASTERDATA TABLE [UDO]
            objUDFEngine.CreateTable("MIPLDM", "Discount Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.addField("MIPLDM", "DiscountGroup", "DiscountGroup", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "A,W,S,B", "AmountBase,WeightBase,STDBase,BoardBase", "")
            objUDFEngine.AddAlphaField("MIPLDM", "DiscID", "Discount ID", 20)
            objUDFEngine.AddFloatField("MIPLDM", "DiscPercent", "Discount %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLDM", "DiscAmount", "Discount Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLDM", "Quantity", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("MIPLDM", "Active", "Active", 1)
            objUDFEngine.AddDateField("MIPLDM", "ActiveFromdate", "Active From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLDM", "ActiveTodate", "Active To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLDM", "ActiveRemarks", "Active Remarks", 100)
            objUDFEngine.AddAlphaField("MIPLDM", "Inactive", "Inactive", 1)
            objUDFEngine.AddDateField("MIPLDM", "InactiveFromdate", "Inactive From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLDM", "InactiveTodate", "Inactive To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLDM", "InactiveRemarks", "Inactive Remarks", 100)
            objUDFEngine.AddAlphaField("MIPLDM", "branchcode", "Branch Code", 4)
            objUDFEngine.AddAlphaField("MIPLDM", "branchname", "Branch Name", 200)

            ''DISCOUNT MASTERDATALINES TABLE [UDO]
            objUDFEngine.CreateTable("MIPLDM1", "Discount Master Row", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddFloatField("MIPLDM1", "FromAmount", "From Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLDM1", "ToAmount", "To Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLDM1", "DiscPercent", "Disc Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLDM1", "DiscAmount", "Disc Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLDM1", "2levelpercent", "2nd Level Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLDM1", "2levelAmount", "2nd level Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLDM1", "finalpercent", "Final Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)


            ''ADDON GENERAL SETTING TABLE
            objUDFEngine.CreateTable("MIPLAGS", "Addon General Settings", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLAGS", "Font", "Font Name", 30)
            objUDFEngine.AddNumericField("MIPLAGS", "Fontsize", "Font Size", 10)
            objUDFEngine.AddAlphaField("MIPLAGS", "FormBColor", "Form Back Color", 30)
            objUDFEngine.AddAlphaField("MIPLAGS", "ButtonBColor", "Button Back Color", 30)
            objUDFEngine.AddAlphaField("MIPLAGS", "SystemInfo", "System Information", 1)
            objUDFEngine.AddNumericField("MIPLAGS", "EntryMinYear", "Entry Minimum Year", 4)

            ''CHANGING DESCRIPTION TABLE
            objUDFEngine.CreateTable("MIPLCD", "Changing Description", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLCD", "FormName", "Form Name", 128)
            objUDFEngine.AddAlphaField("MIPLCD", "ControlName", "Control Name", 128)
            objUDFEngine.AddAlphaField("MIPLCD", "DefaultControlText", "Default Control Text", 128)
            objUDFEngine.AddAlphaField("MIPLCD", "NewControlText", "New Control Text", 128)
            objUDFEngine.AddAlphaField("MIPLCD", "IsBold", "Is Bold", 1)
            objUDFEngine.AddAlphaField("MIPLCD", "IsItalic", "Is Italic", 1)

            ''VALIDVALUES
            objUDFEngine.CreateTable("VALIDVALUES", "Valid Values", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUDFEngine.AddAlphaField("VALIDVALUES", "Type", "Type", 5)
            objUDFEngine.AddAlphaField("VALIDVALUES", "Process", "Process", 10)
            objUDFEngine.AddAlphaField("VALIDVALUES", "ValidValue", "ValidValue", 2)
            objUDFEngine.AddAlphaField("VALIDVALUES", "ValidDescr", "ValidDescription", 30)


            ''PRODUCT MASTER TABLE
            objUDFEngine.CreateTable("MIPLIM", "Product Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLIM", "ProdCode", "Product Code", 20) '1
            objUDFEngine.AddAlphaField("MIPLIM", "ProdName", "Product Name", 100) '2
            objUDFEngine.AddAlphaField("MIPLIM", "ProdGroup", "Product Group", 10) '3
            objUDFEngine.AddAlphaField("MIPLIM", "UOM", "UOM", 10) '4
            objUDFEngine.AddAlphaField("MIPLIM", "MetalType", "Metal Type", 16) '5 ''[M-Metal, O-Ornament]
            objUDFEngine.AddAlphaField("MIPLIM", "Category", "Category", 16)
            objUDFEngine.AddAlphaField("MIPLIM", "PurityID", "PurityID", 20) '6
            objUDFEngine.AddAlphaField("MIPLIM", "PurityDesc", "Purity Description", 100) '7
            objUDFEngine.AddFloatField("MIPLIM", "Purity", "Purity Perentage", SAPbobsCOM.BoFldSubTypes.st_Percentage) '8
            objUDFEngine.AddAlphaField("MIPLIM", "StockType", "Stock Type", 16) '9 [T-Taged, N-NonTaged]
            objUDFEngine.AddAlphaField("MIPLIM", "Size", "Size", 1) '10 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "Stones", "Stones", 1) '11 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "BRateDiff", "Billing RateDiff", 1) '12 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "BWastage", "Billing Wastage", 1) '13 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "Wastage", "Wastage", 1) '14 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "SalesMode", "Sales Mode", 16) '15 ''[G-Gross Weight, N-Net Weight, R-Rate, F-Fixed]
            objUDFEngine.AddAlphaField("MIPLIM", "PurMode", "Purchase Mode", 16) '16 ''[G-Gross Weight, N-Net Weight, R-Rate]
            objUDFEngine.AddAlphaField("MIPLIM", "MultiMetal", "Multi Metal", 1) '17 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "Discount", "Discount", 1) '18 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "DiscountID", "Discount ID", 20) '19
            objUDFEngine.AddAlphaField("MIPLIM", "HallmkChrge", "Hallmark Charges", 1) '20 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "HallmkID", "Hallmark ID", 20) '21
            objUDFEngine.AddAlphaField("MIPLIM", "OtherChrge", "Other Charges", 1) '22 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "LessTaxOnRate", "Less Tax On Rate", 1) '23 [Y-Yes, N-No]
            objUDFEngine.AddAlphaField("MIPLIM", "Brand", "Brand", 10) '26
            objUDFEngine.AddAlphaField("MIPLIM", "TaxCode", "TaxCode", 30) '24
            objUDFEngine.AddAlphaField("MIPLIM", "Active", "Active", 1)
            objUDFEngine.AddDateField("MIPLIM", "ActiveFromdate", "Active From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLIM", "ActiveTodate", "Active To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLIM", "ActiveRemarks", "Active Remarks", 100)
            objUDFEngine.AddAlphaField("MIPLIM", "Inactive", "Inactive", 1)
            objUDFEngine.AddDateField("MIPLIM", "InactiveFromdate", "Inactive From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLIM", "InactiveTodate", "Inactive To Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLIM", "InactiveRemarks", "Inactive Remarks", 100)

            ''PRODUCT MASTER LINE DATA
            objUDFEngine.CreateTable("MIPLIM1", "Product Master Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("MIPLIM1", "WhsCode", "Warehouse Code", 20)
            objUDFEngine.AddAlphaField("MIPLIM1", "WhsName", "Warehouse Name", 100)
            objUDFEngine.addField("MIPLIM1", "Locked", "Locked", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
            objUDFEngine.AddFloatField("MIPLIM1", "InStock", "InStock", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("MIPLIM1", "Branch", "Branch", 50)

            ''SALES WASTAGE MASTER 
            objUDFEngine.CreateTable("MIPLSWM", "Sales Wastage Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLSWM", "ProductCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLSWM", "Productname", "Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLSWM", "Subproductcode", "Sub Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLSWM", "Subproductname", "Sub Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLSWM", "WastageID", "WastageID", 20)
            objUDFEngine.AddAlphaField("MIPLSWM", "WastageType", "WastageType", 2)
            objUDFEngine.AddAlphaField("MIPLSWM", "makingchrg", "Making chrage type", 2)
            objUDFEngine.AddFloatField("MIPLSWM", "Pgepergram", "Percentage Per Gram", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLSWM", "Fixedgram", "Fixed Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLSWM", "FixedAmount", "Fixed Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLSWM", "Amtpergram", "Amount Per Gram", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLSWM", "Wastagepercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddAlphaField("MIPLSWM", "Branch", "Branch", 100)

            ''SALES WASTAGE LINE MASTER
            objUDFEngine.CreateTable("MIPLSWM1", "Sales Wastage Master Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddFloatField("MIPLSWM1", "fromweight", "From Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLSWM1", "toweight", "To Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddAlphaField("MIPLSWM1", "Size", "Size", 100)
            objUDFEngine.AddFloatField("MIPLSWM1", "MaxWastageper", "Maximum Wastage percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLSWM1", "MinWastageper", "Minimum Wastage percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLSWM1", "MaxWastageGm", "Maximum Wastage Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLSWM1", "MinWastageGm", "Minimum Wastage Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLSWM1", "Wastagefixedamt", "Wastage Fixed Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLSWM1", "Wastagefixedgram", "Wastage Fixed Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLSWM1", "stonechrgamt", "Stone Charges Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLSWM1", "Stonechrgper", "Stone Charges Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

            objUDFEngine.CreateTable("MIPLDNEW", "Define New Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUDFEngine.AddAlphaField("MIPLDNEW", "DNewType", "Define New Type", 20)
            objUDFEngine.AddNumericField("MIPLDNEW", "DNewCode", "Define New Code", 8)
            objUDFEngine.AddAlphaField("MIPLDNEW", "DNewName", "Define New Name", 50)
            ''**********************************************************************************************************

            ''Helen

            'Special Sales Wastage Master

            objUDFEngine.CreateTable("MIPLSSWM", "Special Sales Wastage Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLSSWM", "EventName", "Event Name", 50)
            objUDFEngine.AddDateField("MIPLSSWM", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)

            'Special Sales Wastage Master Line Table
            objUDFEngine.CreateTable("MIPLSSWM1", "SpecialSalesWastageLine", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "Branch", "Branch", 30)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "ProductCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "ProductName", "Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "SubProdcode", "Sub Product code", 20)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "SubProdName", "Sub Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLSSWM1", "WastType", "Wastage Type", 16)
            objUDFEngine.AddFloatField("MIPLSSWM1", "WastPer", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLSSWM1", "WastAmt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddDateField("MIPLSSWM1", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLSSWM1", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)

            ''Metal Type Master (Helen)
            objUDFEngine.CreateTable("MIPLMT", "Metal Type Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLMT", "MetalType", "MetalType", 50)

            ''Employee Master(OHEM) UDF (Rajesh)
            objUDFEngine.AddAlphaField("OHEM", "LoginUser", "Login User Name", 30)
            objUDFEngine.AddAlphaField("OHEM", "LoginPassword", "Login Password", 30)

            ''Branch Master(Employee Master Definenew) (Rajesh)
            objUDFEngine.AddNumericField("OUBR", "LctCode", "Location Code", 4)

            'Old Gold Rate Master (Helen)
            objUDFEngine.CreateTable("MIPLOGRM", "Old Gold Rate Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddDateField("MIPLOGRM", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("MIPLOGRM", "Time", "Time", SAPbobsCOM.BoFldSubTypes.st_Time)
            objUDFEngine.AddAlphaField("MIPLOGRM", "Branch", "Branch", 30)

            objUDFEngine.CreateTable("MIPLOGRM1", "Old Gold Rate Master1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("MIPLOGRM1", "MetalType", "MetalType", 30)
            objUDFEngine.AddFloatField("", "CashPer", "CashinPercentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("", "CashAmt", "CashinAmount", SAPbobsCOM.BoFldSubTypes.st_Price)

            'OG Reason master (Helen)
            objUDFEngine.CreateTable("MIPLREASON", "Reason Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("MIPLREASON", "screen", "Screen", 100)
            objUDFEngine.AddAlphaField("MIPLREASON", "reason", "Reason", 200)

            ''Transaction Tables Start***************************************************

            ''Material Receipt (Praveen)
            objUDFEngine.AddAlphaField("OIGN", "Workordertype", "Work Order Type", 2)
            objUDFEngine.AddAlphaField("OIGN", "Workorderno", "Work Order No", 30)
            objUDFEngine.AddAlphaField("OIGN", "invoiceno", "Invoice No", 30)
            objUDFEngine.AddAlphaField("OIGN", "Vendorcode", "Vendor Code", 15)
            objUDFEngine.AddAlphaField("OIGN", "Vendorname", "Vendor Name", 100)
            objUDFEngine.AddAlphaField("OIGN", "Vendorshortname", "Vendor Short Name", 100)
            objUDFEngine.AddAlphaField("OIGN", "Vendorref", "Vendor Ref", 30)
            objUDFEngine.AddAlphaField("OIGN", "lotstatus", "Lot Status", 2)
            objUDFEngine.AddDateField("OIGN", "workorderdate", "Work Order Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("OIGN", "invoicedate", "Invoice Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("OIGN", "transactiontype", "Transaction Type", 2)
            objUDFEngine.AddAlphaField("OIGN", "Employeename", "Employee name", 100)


            objUDFEngine.AddAlphaField("IGN1", "Prodcode", "Product Code", 20)
            objUDFEngine.AddAlphaField("IGN1", "Prodname", "Product Name", 100)
            objUDFEngine.AddFloatField("IGN1", "LessWeight", "Less Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("IGN1", "NetWeight", "Net Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("IGN1", "Wastagepercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("IGN1", "wastagegram", "Wastage Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("IGN1", "Puritypercen", "Purity Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("IGN1", "Purityweight", "Purity Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddNumericField("IGN1", "pieces", "Pieces", 10)
            objUDFEngine.AddFloatField("IGN1", "makingchrgamt", "Making Charge Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("IGN1", "Hallmrkvendor", "Hall Mark Vendor", 30)
            objUDFEngine.AddAlphaField("IGN1", "lotstatus", "Lot Status", 2)
            objUDFEngine.AddAlphaField("IGN1", "countername", "Counter Name", 100)

            ''Purchase Invoice (Praveen)
            objUDFEngine.AddAlphaField("OPCH", "Matrecstatus", "MaterialReceiptstatus", 2)

            ''Purchase Invoice Line Table (Rajesh)
            objUDFEngine.AddAlphaField("PCH1", "ProdCode", "Product Code", 20)
            objUDFEngine.AddAlphaField("PCH1", "ProdName", "Product Name", 100)
            objUDFEngine.AddFloatField("PCH1", "LessWeight", "Less Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("PCH1", "NetWeight", "Net Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("PCH1", "WastagePercen", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("PCH1", "WastageGram", "Wastage Gram", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("PCH1", "WastageAmt", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("PCH1", "PurityPercen", "Purity Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("PCH1", "PureWeight", "Pure Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddNumericField("PCH1", "Pieces", "Pieces", 10)
            objUDFEngine.AddFloatField("PCH1", "MakingChrg", "Making Charges", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("PCH1", "BeforeDisc", "Before Discount", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("PCH1", "CounterName", "Counter Name", 100)
            objUDFEngine.AddAlphaField("PCH1", "Matrecstatus", "MaterialReceiptstatus", 2)


            'Category Master
            objUDFEngine.CreateTable("MIPLCAT", "Category Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUDFEngine.AddAlphaField("MIPLCAT", "category", "Category", 50)
            objUDFEngine.AddAlphaField("MIPLCAT", "process", "Process", 100)

            'OG Estimation Header
            objUDFEngine.CreateTable("MIPLOGH", "AVR-OGEstimation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("MIPLOGH", "docnum", "OG Estimation No", 20)
            objUDFEngine.AddDateField("MIPLOGH", "docdate", "Estimation Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLOGH", "canceled", "Cancelled", 2)
            objUDFEngine.AddAlphaField("MIPLOGH", "remittance", "Remittance", 2)
            objUDFEngine.AddAlphaField("MIPLOGH", "remit", "Remittance", 10)
            objUDFEngine.AddAlphaField("MIPLOGH", "cardcode", "Customer Code", 30)
            objUDFEngine.AddAlphaField("MIPLOGH", "cardname", "Customer Name", 100)
            objUDFEngine.AddAlphaField("MIPLOGH", "address", "Address", 254)
            objUDFEngine.AddAlphaField("MIPLOGH", "contactno", "Contact No", 20)

            objUDFEngine.AddAlphaField("MIPLOGH", "series", "Series", 20)
            objUDFEngine.AddAlphaField("MIPLOGH", "docstatus", "Status  Code", 2)
            objUDFEngine.AddAlphaField("MIPLOGH", "status", "Document Status", 10)
            'objUDFEngine.AddAlphaField("MIPLOGH", "makingtype", "Making Type", 2)
            objUDFEngine.AddAlphaField("MIPLOGH", "making", "Making", 10)
            objUDFEngine.AddAlphaField("MIPLOGH", "saldocnum", "Sales Document Number", 20)
            objUDFEngine.AddDateField("MIPLOGH", "saldocdate", "Sales Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

            objUDFEngine.AddAlphaField("MIPLOGH", "empid", "Employee ID", 20)
            objUDFEngine.AddAlphaField("MIPLOGH", "remarks", "Remarks", 254)
            objUDFEngine.AddFloatField("MIPLOGH", "baseamt", "Base Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGH", "discount", "Discount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGH", "taxamt", "Tax	Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGH", "othercharges", "Other Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGH", "doctotal", "Document Total", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddAlphaField("MIPLOGH", "rackno", "Rack No", 20)
            objUDFEngine.AddAlphaField("MIPLOGH", "approval1", "Approval 1", 40)
            objUDFEngine.AddAlphaField("MIPLOGH", "approval2", "Approval 2", 40)
            objUDFEngine.AddAlphaField("MIPLOGH", "approvedby", "Approved by", 40)

            'OG Estimation Lines
            objUDFEngine.CreateTable("MIPLOGL", "AVR-OGEstimation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            objUDFEngine.AddAlphaField("MIPLOGL", "productcode", "Product Code", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "productname", "Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLOGL", "subproductcode", "Sub Product Code", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "subproductname", "Sub Product Name", 100)
            objUDFEngine.AddAlphaField("MIPLOGL", "uom", "Uom", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "makingtype", "Making Type", 10)
            objUDFEngine.AddAlphaField("MIPLOGL", "metaltype", "Metal Type", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "category", "Category", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "process", "Process", 30)
            objUDFEngine.AddFloatField("MIPLOGL", "grossweight", "Gross Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLOGL", "stoneweight", "Stone Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLOGL", "dustweight", "Dust Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLOGL", "netweight", "Net Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLOGL", "Wastageper", "Wastage Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGL", "wastagegrams", "Wastage Grams", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddFloatField("MIPLOGL", "wastageamount", "Wastage Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGL", "purityper", "Purity Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGL", "pureweight", "Pure Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
            objUDFEngine.AddNumericField("MIPLOGL", "pieces", "Pieces", 10)
            objUDFEngine.AddFloatField("MIPLOGL", "rate", "Rate", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddFloatField("MIPLOGL", "stonecharges", "Stone Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGL", "makingcharges", "Making Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGL", "hallmarkcharges", "Hallmark Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGL", "othercharges", "Other Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGL", "linetotal", "LineTotal", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddAlphaField("MIPLOGL", "salestagno", "SalesTagNo", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "taxcode", "Tax Code", 30)
            objUDFEngine.AddFloatField("MIPLOGL", "taxamount", "Tax Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddAlphaField("MIPLOGL", "countercode", "Counter Code", 30)
            objUDFEngine.AddAlphaField("MIPLOGL", "countername", "Counter Name", 100)
            objUDFEngine.AddAlphaField("MIPLOGL", "branch", "Branch Name", 100)
            'objUDFEngine.AddAlphaField("MIPLOGL", "warehousecode", "Warehouse Code", 30)
            'objUDFEngine.AddAlphaField("MIPLOGL", "warehousename", "Warehouse Name", 100)
            'objUDFEngine.AddAlphaField("MIPLOGL", "locationcode", "Location Code", 30)
            objUDFEngine.AddFloatField("MIPLOGL", "openqty", "OpenQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("MIPLOGL", "linestatus", "LineStatus", 2)
            objUDFEngine.AddAlphaField("MIPLOGL", "lstatus", "LineStatus", 10)

            'OG Document setting
            objUDFEngine.CreateTable("MIPLOGDOC", "OG Document Settings", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUDFEngine.AddDateField("MIPLOGDOC", "effdate", "Effect Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("MIPLOGDOC", "stdocno", "Start Document no", 10)
            objUDFEngine.AddAlphaField("MIPLOGDOC", "nextdocno", "Next", 10)
            objUDFEngine.AddFloatField("MIPLOGDOC", "ornpurper", "ornament purchase Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGDOC", "cashremper", "cash Remittance %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGDOC", "cashremamt", "Cash Remittance Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddFloatField("MIPLOGDOC", "gcpurper", "Gold Coin Purchase %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            objUDFEngine.AddFloatField("MIPLOGDOC", "gcpuramt", "Gold coin Purchase Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
            objUDFEngine.AddNumericField("MIPLOGDOC", "srdays", "SR - Days", 10)
            objUDFEngine.AddFloatField("MIPLOGDOC", "platinumper", "Platinum %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'objUDFEngine.AddFloatField("MIPLOGDOC", "platinumamt", "Platinum Amount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'objUDFEngine.AddFloatField("MIPLOGDOC", "silverper", "Silver %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'objUDFEngine.AddFloatField("MIPLOGDOC", "silveramt", "Silver Amount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'objUDFEngine.AddFloatField("MIPLOGDOC", "diamondper", "Diamond %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'objUDFEngine.AddFloatField("MIPLOGDOC", "diamondamt", "Diamond Amount", SAPbobsCOM.BoFldSubTypes.st_Percentage)


            ''Transaction Tables End*****************************************************
            objAddOn.objApplication.SetStatusBarMessage("Checking Process Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        'If eventInfo.FormUID.Contains(Clscent.FormType) Then
        '    objcent.RightClickEvent(eventInfo, BubbleEvent)
        'End If
        'If eventInfo.FormUID.Contains(ClsPurchase_Wastage.Formtype) Then
        '    objPur.RightClickEvent(eventInfo, BubbleEvent)
        'End If
    End Sub

    Private Sub createUDOG(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False, Optional ByVal LogForm As Boolean = True)
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        'Dim i As Integer
        Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udocode) Then
            oUserObjectMD.Code = udocode
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                oUserObjectMD.FormColumns.Add()
                oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                oUserObjectMD.FormColumns.Add()
            Else
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "MIPLPM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_MetalType"
                            oUserObjectMD.FindColumns.ColumnDescription = "MetalType"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_PurityID"
                            oUserObjectMD.FindColumns.ColumnDescription = "PurityID"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Description"
                            oUserObjectMD.FindColumns.ColumnDescription = "Description"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Purity"
                            oUserObjectMD.FindColumns.ColumnDescription = "Purity"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Basecode"
                            oUserObjectMD.FindColumns.ColumnDescription = " BaseCode"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Amtinper"
                            oUserObjectMD.FindColumns.ColumnDescription = "AmountinPer"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_AmtinRs"
                            oUserObjectMD.FindColumns.ColumnDescription = "AmountinRupees"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLSM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Prodcode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ProdName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Sizecode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Size Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_SizeName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Size Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLPWM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_VENDORCODE"
                            oUserObjectMD.FindColumns.ColumnDescription = "Vendor Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_VENDORNAME"
                            oUserObjectMD.FindColumns.ColumnDescription = "Vendor Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLHMC"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_BranchCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Branch Code"
                            oUserObjectMD.FindColumns.Add()

                            oUserObjectMD.FindColumns.ColumnAlias = "U_HallmarkId"
                            oUserObjectMD.FindColumns.ColumnDescription = "HallmarkID"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_HallmarkPer"
                            oUserObjectMD.FindColumns.ColumnDescription = "HallmarkPer"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_HallmarkAmt"
                            oUserObjectMD.FindColumns.ColumnDescription = "HallmarkAmt"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Remarks"
                            oUserObjectMD.FindColumns.ColumnDescription = "Remarks"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLSSWM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_EventName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Event Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Date"
                            oUserObjectMD.FindColumns.ColumnDescription = "Date"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLMT"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_MetalType"
                            oUserObjectMD.FindColumns.ColumnDescription = "MetalType"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLREASON"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                    End Select
                ElseIf type = SAPbobsCOM.BoUDOObjType.boud_Document Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        'Maintenance
                    End Select
                End If
            End If

            If LogForm Then oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES


            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                MsgBox("error" + CStr(lRetCode))
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
            End If
        End If
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Successfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Sub createUDOC(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udocode) Then
            oUserObjectMD.Code = udocode
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            ' oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_Document Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "MIPLLOT"

                    End Select
                Else
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "MIPLSPM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ItemCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Item Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ItemName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Item Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_SubProdcode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Sub Product Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_SubProdName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Sub Product Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLCM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Prodcode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ProdName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_SubProdcode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Sub Product Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_SubProdName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Sub Product Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLRM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Date"
                            oUserObjectMD.FindColumns.ColumnDescription = "Date"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Time"
                            oUserObjectMD.FindColumns.ColumnDescription = "Time"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLIM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ProdCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ProdName"
                            oUserObjectMD.FindColumns.ColumnDescription = "Product Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLSSWM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLOGRM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Date"
                            oUserObjectMD.FindColumns.ColumnDescription = "Date"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Time"
                            oUserObjectMD.FindColumns.ColumnDescription = "Time"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Branch"
                            oUserObjectMD.FindColumns.ColumnDescription = "Branch"
                            oUserObjectMD.FindColumns.Add()
                    End Select
                End If
            End If
            If childTable.Length > 0 Then
                For i = 0 To childTable.Length - 2
                    If Trim(childTable(i)) <> "" Then
                        oUserObjectMD.ChildTables.TableName = childTable(i)
                        oUserObjectMD.ChildTables.Add()
                    End If
                Next
            End If
            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                MsgBox("Error Code : " + CStr(lRetCode))
                MsgBox("Error Desc : " + objAddOn.objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
            End If
        End If
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Suceessfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()

                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
    End Sub

    Public Sub LoadInitialize()
        Dim FBColor As String = ""
        Dim BBColor As String = ""
        strSql = "SELECT ISNULL(U_FONT,'TAHOMA')FONT"
        strSql += vbCrLf + " , ISNULL(U_FONTSIZE,7)FONTSIZE  "
        strSql += vbCrLf + " ,ISNULL(U_FORMBCOLOR,'CLASSIC')FORMBCOLOR"
        strSql += vbCrLf + " , ISNULL(U_BUTTONBCOLOR,'CLASSIC')BUTTONBCOLOR "
        strSql += vbCrLf + " , ISNULL(U_SYSTEMINFO,'N')SYSTEMINFO  "
        strSql += vbCrLf + " FROM [@MIPLAGS]"
        dt = New DataTable
        da = New SqlDataAdapter(strSql, con)
        da.Fill(dt)
        If dt.Rows.Count > 0 Then
            MIPL_Fontsize = Val(dt.Rows(0)("FONTSIZE").ToString)
            MIPL_Font = dt.Rows(0)("FONT").ToString
            FBColor = dt.Rows(0)("FORMBCOLOR").ToString
            BBColor = dt.Rows(0)("BUTTONBCOLOR").ToString
            MIPL_SystemInformation = IIf(dt.Rows(0)("SYSTEMINFO").ToString.ToUpper = "Y", True, False)

            Select Case FBColor.ToUpper
                Case "CLASSIC"
                    frmBGColorRed = 234
                    frmBGColorGreen = 241
                    frmBGColorBlue = 246
                Case "COMBINED"
                    frmBGColorRed = 234
                    frmBGColorGreen = 241
                    frmBGColorBlue = 246
                Case "COOL GREY"
                    frmBGColorRed = 243
                    frmBGColorGreen = 243
                    frmBGColorBlue = 243
                Case "VIOLET"
                    frmBGColorRed = 241
                    frmBGColorGreen = 238
                    frmBGColorBlue = 250
                Case "TURQUOISE"
                    frmBGColorRed = 240
                    frmBGColorGreen = 255
                    frmBGColorBlue = 249
                Case "MINT"
                    frmBGColorRed = 240
                    frmBGColorGreen = 249
                    frmBGColorBlue = 228
                Case "LEMON"
                    frmBGColorRed = 255
                    frmBGColorGreen = 254
                    frmBGColorBlue = 238
                Case "CREAM"
                    frmBGColorRed = 255
                    frmBGColorGreen = 248
                    frmBGColorBlue = 219
                Case "PINK"
                    frmBGColorRed = 248
                    frmBGColorGreen = 232
                    frmBGColorBlue = 241
                Case "COFFEE"
                    frmBGColorRed = 241
                    frmBGColorGreen = 234
                    frmBGColorBlue = 225
                Case Else
                    frmBGColorRed = 234
                    frmBGColorGreen = 241
                    frmBGColorBlue = 246
            End Select
            MIPL_FormBackColor = Color.FromArgb(frmBGColorRed, frmBGColorGreen, frmBGColorBlue)

            Select Case BBColor.ToUpper
                Case "DEFAULT"
                    frmButtonColorRed = 255
                    frmButtonColorGreen = 238
                    frmButtonColorBlue = 159
                Case Else
                    frmButtonColorRed = 150
                    frmButtonColorGreen = 150
                    frmButtonColorBlue = 250
            End Select
            MIPL_ButtonColor = Color.FromArgb(frmButtonColorRed, frmButtonColorGreen, frmButtonColorBlue)

        End If
    End Sub

    Function funcInsertValidValues() As Integer
        Try
            Dim str As String = ""

            Tran = con.BeginTransaction

            str = "DELETE [@VALIDVALUES]"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()

            ''CATEGORY(CT) Common(Product,Subproduct)
            str = vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('1','1','CT','COMMON','M','Metal')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('2','2','CT','COMMON','O','Ornament')"

            ''SALESMODE(SM) Common(Product,Subproduct)
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('3','3','SM','COMMON','G','Gross Weight')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('4','4','SM','COMMON','N','Net Weight')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('5','5','SM','COMMON','R','Rate')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('6','6','SM','COMMON','F','Fixed')"
            ''PURCHASEMODE(PM) Product Master   
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('7','7','PM','COMMON','G','Gross Weight')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('8','8','PM','COMMON','N','Net Weight')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('9','9','PM','COMMON','R','Rate')"

            ''STOCKTYPE(ST) Common(Product,Subproduct)
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('10','10','ST','COMMON','T','Taged')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('11','11','ST','COMMON','N','NonTaged')"

            ''WASTAGETYPE(WT) Purchase Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('12','12','WT','PWASTMAST','FG','Fixed Grams')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('13','13','WT','PWASTMAST','FA','Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('14','14','WT','PWASTMAST','FP','Fixed Percentage')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('15','15','WT','PWASTMAST','FT','Fixed Touch')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('16','16','WT','PWASTMAST','PR','Percentage in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('17','17','WT','PWASTMAST','AR','Amount in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('18','18','WT','PWASTMAST','GR','Grams in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('19','19','WT','PWASTMAST','TR','Touch in Range')"

            ''MAKINGCHARGESTYPE(MCT) Purchase Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('20','20','MCT','PWASTMAST','PG','% Per Gram')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('21','21','MCT','PWASTMAST','FG','Fixed Grams')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('22','22','MCT','PWASTMAST','FA','Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('23','23','MCT','PWASTMAST','AG','Amount Per Gram')"

            ''HALLMARKTYPE(HT) Hallmark Charges Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('24','24','HT','HALLMAST','P','Percentage')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('25','25','HT','HALLMAST','R','Rate')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('26','26','HT','HALLMAST','A','Amount')"

            ''DISCOUNTGROUP(DG) Discount Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('27','27','DG','DISCMAST','A','Amount Base')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('28','28','DG','DISCMAST','W','Weight Base')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('29','29','DG','DISCMAST','S','Std Base')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('30','30','DG','DISCMAST','B','Board Rate')"

            ''INCENTIVETYPE(IT) Incentive Master
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('31','31','IT','ICMAST','WT','Weight')"
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('32','32','IT','ICMAST','CT','Cent')"
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('33','33','IT','ICMAST','PE','Piece')"

            ''MC TYPE(MCT) Subproduct Master
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('34','34','MCT','SPRODMAST','I','Inclusive')"
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('35','35','MCT','SPRODMAST','E','Exclusive')"

            ''OG Rate(OGR) Subproduct Master
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('36','36','OGR','SPRODMAST','C','Cash Rate')"
            str += vbCrLf + "Insert into [@VALIDVALUES] (CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('37','37','OGR','SPRODMAST','E','Exchange Rate')"

            ''WASTAGETYPE(WT) Repair Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('38','38','WT','RWASTMAST','FG','Fixed Grams')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('39','39','WT','RWASTMAST','FA','Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('40','40','WT','RWASTMAST','FP','Fixed Percentage')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('41','41','WT','RWASTMAST','FT','Fixed Touch')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('42','42','WT','RWASTMAST','PR','Percentage in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('43','43','WT','RWASTMAST','AR','Amount in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('44','44','WT','RWASTMAST','GR','Grams in Range')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('45','45','WT','RWASTMAST','TR','Touch in Range')"

            ''MAKINGCHARGESTYPE(MCT) Repair Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('46','46','MCT','RWASTMAST','PG','% Per Gram')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('47','47','MCT','RWASTMAST','FG','Fixed Grams')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('48','48','MCT','RWASTMAST','FA','Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('49','49','MCT','RWASTMAST','AG','Amount Per Gram')"

            ''Wastage Type(WT) Sales Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('50','50','WT','SWASTMAST','FP','Fixed Percentage')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('51','51','WT','SWASTMAST','SP','Slab Percentage')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('52','52','WT','SWASTMAST','SG','Slab Grams')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('53','53','WT','SWASTMAST','SA','Slab Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('54','54','WT','SWASTMAST','FA','Slab Fixed Gram')"

            ''Making Charges Type(MT) Sales Wastage Master
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('55','55','MT','SWASTMAST','PG','% Per Gram')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('56','56','MT','SWASTMAST','FG','Fixed Gram')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('57','57','MT','SWASTMAST','FA','Fixed Amount')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('58','58','MT','SWASTMAST','AG','Amount Per Gram')"

            ''DocStatus(DS) Common[Purchase Invoice]
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('59','59','DS','COMMON','O','Open')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('60','60','DS','COMMON','C','Closed')"

            ''Work Order Type(WOT)   MeterialReceipt
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('61','61','WOT','MRECEIPT','I','Invoice')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('62','62','WOT','MRECEIPT','W','Work Order')"

            ''Transaction Type(TT) Meterial Receipt
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('63','63','TT','MRECEIPT','P','Pure Weight')"
            str += vbCrLf + " INSERT INTO [@VALIDVALUES](CODE,NAME,U_TYPE,U_PROCESS,U_VALIDVALUE,U_VALIDDESCR) VALUES('64','64','TT','MRECEIPT','D','Direct Payment')"

            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()

            Tran.Commit()
            Tran.Dispose()

        Catch ex As Exception
            If Not Tran Is Nothing Then
                Tran.Rollback()
                Tran.Dispose()
            End If
        End Try
    End Function

    Private Sub addvendorgroup()
        Dim dttemp As DataTable
        strSql = "select Groupname from OCRG where GroupType='S'"
        dttemp = New DataTable
        da = New SqlDataAdapter(strSql, con)
        da.Fill(dttemp)
        If dttemp.Rows.Count > 0 Then
            Dim groupflag As Boolean
            For i As Integer = 0 To dttemp.Rows.Count - 1
                If dttemp.Rows(i).Item("Groupname") = "HallMark" Then
                    groupflag = True
                    Exit For
                Else
                    groupflag = False
                End If
            Next
            If groupflag = False Then
                Dim objbp As SAPbobsCOM.BusinessPartnerGroups
                objbp = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)
                objbp.Name = "HallMark"
                objbp.Type = SAPbobsCOM.BoBusinessPartnerGroupTypes.bbpgt_VendorGroup
                objbp.Add()
            End If
        End If
    End Sub

    Public Sub New()

    End Sub
End Class




