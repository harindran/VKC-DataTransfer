Imports System.IO
Imports System.Data.SqlClient
Namespace Mukesh.SBOLib
    Public Class SBOConnector
        Dim objGVar As New clsGlobalVariables

        Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
        Dim lretcode As Long
        Dim retStr As String
        Public Function GetApplication(ByVal ConnectionStr As String) As SAPbouiCOM.Application
            Dim objGUIAPI As SAPbouiCOM.SboGuiApi
            Dim objApp As SAPbouiCOM.Application

            'Try
            '    objGUIAPI = New SAPbouiCOM.SboGuiApi
            '    objGUIAPI.Connect(ConnectionStr)
            '    objApp = objGUIAPI.GetApplication(-1)
            '    If Not objApp Is Nothing Then Return objApp
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            '    End
            'End Try
            Return Nothing
        End Function

        'Public Function GetCompany(ByVal user As String, ByVal DBNAMES As String) As SAPbobsCOM.Company
        '    Try

        '        '  Dim objmaincom As New clsMainCompanyvar()



        '        objGVar.MDDBMain = objMUtil.xDBName.ToString.Trim
        '        objGVar.MServerName = objMUtil.xServerName.ToString.Trim

        '        If user.ToString.ToString = "N" Then


        '            objGVar.MDBName = objMUtil.xDBName.ToString.Trim
        '            objGVar.MUID = objMUtil.xUID.ToString.Trim
        '            objGVar.MPWD = objMUtil.xPWD.ToString.Trim
        '            objGVar.MSAPUID = objMUtil.xSAPUser.ToString.Trim
        '            objGVar.MSAPPWd = objMUtil.xSAPPwd.ToString.Trim
        '            Dim objCompany As New SAPbobsCOM.Company

        '            objCompany.Server = objMUtil.xServerName.ToString.Trim
        '            objCompany.LicenseServer = objMUtil.xServerName.ToString.Trim + ":30000"
        '            objCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        '            objCompany.UseTrusted = False
        '            objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
        '            objCompany.DbUserName = objMUtil.xUID.ToString.Trim
        '            objCompany.DbPassword = objMUtil.xPWD.ToString.Trim
        '            objCompany.CompanyDB = objGVar.MDBName.ToString.Trim
        '            objCompany.UserName = objMUtil.xSAPUser.ToString.Trim
        '            objCompany.Password = objMUtil.xSAPPwd.ToString.Trim


        '            lretcode = objCompany.Connect()


        '            If lretcode <> 0 Then
        '                objCompany.GetLastError(lretcode, retStr)
        '                MsgBox("Error " & lretcode & " " & retStr)
        '                Exit Function
        '            End If


        '            CompanyMainCon = objCompany
        '            Return objCompany


        '        End If




        '    Catch ex As Exception
        '        MsgBox(ex.Message & vbLf & ex.StackTrace)
        '    End Try
        '    Return Nothing
        'End Function

        'Public Function OBJMainCompanys() As SAPbobsCOM.Company
        '    Dim lretcodes As Long
        '    Dim OBJMains As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
        '    objGVar.MAINobjCompany.Server = OBJMains.xServerName.ToString.Trim
        '    objGVar.MAINobjCompany.LicenseServer = OBJMains.xServerName.ToString.Trim + ":30000"
        '    objGVar.MAINobjCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        '    objGVar.MAINobjCompany.UseTrusted = False
        '    objGVar.MAINobjCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
        '    objGVar.MAINobjCompany.DbUserName = OBJMains.xUID.ToString.Trim
        '    objGVar.MAINobjCompany.DbPassword = OBJMains.xPWD.ToString.Trim
        '    'objGVar.MAINobjCompany.CompanyDB = MDBMainName.ToString.Trim
        '    objGVar.MAINobjCompany.UserName = OBJMains.xSAPUser.ToString.Trim
        '    objGVar.MAINobjCompany.Password = OBJMains.xSAPPwd.ToString.Trim
        '    lretcodes = objGVar.MAINobjCompany.Connect()
        '    If lretcodes <> 0 Then
        '        Exit Function
        '    Else
        '        Return objGVar.MAINobjCompany
        '    End If
        'End Function

    End Class
End Namespace