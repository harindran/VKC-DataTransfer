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


            ''**********************************************************************************
            'objCompany = objSBOConnector.GetCompany(user)
            'If user.ToString = "N" Then
            '    objCompanyMain = objCompany
            'Else
            '    objCompanyMain = CompanyMainCon 'objSBOConnector.OBJMainCompanys()
            'End If


            'If objCompany.CompanyDB.ToString <> "" Then
            '    ' objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
            '    conns = True
            'Else
            '    conns = False
            '    MsgBox("You are not connected to a company")
            '    Application.Exit()
            'End If

            'LoadInitialize()
            'objstart1.Close()

        Catch ex As Exception
            'MsgBox(ex.ToString)
            MsgBox("You are not connected to a company")
            Exit Sub
        End Try
    End Sub

    Public Function OBJMainCompanys(ByVal MDBMainName As String) As SAPbobsCOM.Company

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
            objrs = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                    'objLct.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            'objApplication.MessageBox(ex.ToString)
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
            MsgBox(objCompany.GetLastErrorDescription)
        End Try
    End Function
    Public Function createsqlTable()



        'Try
        '    Tran = con.BeginTransaction
        '    str = " if not exists(select * from sys.all_columns  where name='U_APPVALE') "
        '    str += vbCrLf + "    begin"
        '    str += vbCrLf + "ALTER TABLE dbo.MIPLDTAGH ADD"
        '    str += vbCrLf + "U_APPVALE NUMERIC(38,6) NULL"
        '    str += vbCrLf + "end "
        '    cmd = New SqlCommand(str, con, Tran)
        '    cmd.ExecuteNonQuery()
        '    Tran.Commit()
        '    Tran.Dispose()
        'Catch ex As Exception
        '    objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '    Tran.Rollback()
        '    Tran.Dispose()
        'End Try




        Try
            Tran = con.BeginTransaction
            str = " if not exists (select * from sysobjects where name='MIPLLOGS' and xtype='U')"
            str += vbCrLf + "   CREATE TABLE [dbo].[MIPLLOGS]("
            str += vbCrLf + "   [Code] [bigint] IDENTITY(1,1) NOT NULL,"
            str += vbCrLf + " [CreateDate] [datetime] NULL,"
            str += vbCrLf + " [CreateTime] [nvarchar](100) NULL,"
            str += vbCrLf + " [DocType] [nvarchar](100) NULL,"
            str += vbCrLf + " 	[ToDocType] [nvarchar](100) NULL,"
            str += vbCrLf + " [BaseEntry] [nvarchar](100) NULL,"
            str += vbCrLf + " [PostDB] [nvarchar](100) NULL,"
            str += vbCrLf + " [Flag] [nvarchar](2) NULL,"
            str += vbCrLf + " [ErrorLog] [ntext] NULL"
            str += vbCrLf + " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
            cmd = New SqlCommand(str, con, Tran)
            cmd.ExecuteNonQuery()
            Tran.Commit()
            Tran.Dispose()
        Catch ex As Exception
            objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Tran.Rollback()
            Tran.Dispose()
        End Try
    End Function
    Public Function createTables()
        Try
            Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objFromCompany)
            'objApplication.SetStatusBarMessage("Checking Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            'BaseEntry
            'PostDB
            'PostEntry
            'Flag
            'BaseDB

            'ErrorLog

            objUDFEngine.AddAlphaField("OINV", "BaseEntry", "Work Order Type", 30)
            objUDFEngine.AddAlphaField("OINV", "PostDB", "Post DB", 100)
            objUDFEngine.AddAlphaField("OINV", "PostEntry", "PostEntry", 30)
            objUDFEngine.AddAlphaField("OINV", "Flag", "Flag", 15)
            objUDFEngine.AddAlphaField("OINV", "BaseDB", "Base DB", 100)
            objUDFEngine.AddAlphaMemoField("OINV", "ErrorLog", "Error Log", 999999999)
            objUDFEngine.addField("OINV", "Trantype", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "S,M,N", "Single,Multi,No", "N")
            ''Transaction Tables End*****************************************************
            'objApplication.SetStatusBarMessage("Checking Process Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Function


    Public Function createTables_to()
        Try
            Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objToCompany)
            'objApplication.SetStatusBarMessage("Checking Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)

            'BaseEntry
            'PostDB
            'PostEntry
            'Flag
            'BaseDB

            'ErrorLog

            objUDFEngine.AddAlphaField("opdf", "BaseEntry", "Work Order Type", 30)
            objUDFEngine.AddAlphaField("opdf", "PostDB", "Post DB", 100)
            objUDFEngine.AddAlphaField("opdf", "PostEntry", "PostEntry", 30)
            objUDFEngine.AddAlphaField("opdf", "Flag", "Flag", 15)
            objUDFEngine.AddAlphaField("opdf", "BaseDB", "Base DB", 100)
            objUDFEngine.AddAlphaMemoField("opdf", "ErrorLog", "Error Log", 999999999)
            objUDFEngine.addField("opdf", "Trantype", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "S,M,N", "Single,Multi,No", "N")
            ''Transaction Tables End*****************************************************
            'objApplication.SetStatusBarMessage("Checking Process Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function

    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        'If eventInfo.FormUID.Contains(Clscent.FormType) Then
        '    objcent.RightClickEvent(eventInfo, BubbleEvent)
        'End If
        'If eventInfo.FormUID.Contains(ClsPurchase_Wastage.Formtype) Then
        '    objPur.RightClickEvent(eventInfo, BubbleEvent)
        'End If
    End Sub

    Private Sub createUDOG(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False, Optional ByVal LogForm As Boolean = True)
        objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        'Dim i As Integer
        Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
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
                MsgBox(objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objApplication.Forms.AddEx(creationPackage)
            End If
        End If
        objApplication.SetStatusBarMessage("UDO Created Successfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Sub createUDOC(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
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
                MsgBox("Error Desc : " + objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objApplication.Forms.AddEx(creationPackage)
            End If
        End If
        objApplication.SetStatusBarMessage("UDO Created Suceessfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
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

    

    Public Sub New()

    End Sub
End Class




