Imports System.Data.SqlClient

Module modEnum
    'Company connections
    Public objcompany2 As SAPbobsCOM.Company
    Public CompanyMainCon As SAPbobsCOM.Company
    Public objcompany1 As SAPbobsCOM.Company
    Public objFromCompany As SAPbobsCOM.Company
    Public objToCompany As SAPbobsCOM.Company
    Public objBRcompany3 As SAPbobsCOM.Company
    Public lretcode As Long
    'Sql Connection
    Public da As SqlDataAdapter
    Public cmd As SqlCommand
    Public dt As DataTable
    Public con As New SqlConnection
    'Details Getting From Notepad
    Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
    Public MServerName As String = objMUtil.xServerName
    Public MDBName1 As String = objMUtil.xDBName
    Public MDBName2 As String = ""
    Public DBSERVERNAME As String = ""
    Public series As Integer

    Public MUID As String = objMUtil.xUID
    Public MPWD As String = objMUtil.xPWD
    Public MLoginType As String = objMUtil.xLoginType
    Public MSAPUID As String = objMUtil.xSAPUser
    Public MSAPPWd As String = objMUtil.xSAPPwd
    'Recordset Connection
    Public Objrs1 As SAPbobsCOM.Recordset
    Public Objrs11 As SAPbobsCOM.Recordset
    Public Objrs2 As SAPbobsCOM.Recordset
    Public Objrs3 As SAPbobsCOM.Recordset
    Public Objrs4 As SAPbobsCOM.Recordset
    Public Objrs5 As SAPbobsCOM.Recordset
    Public strsql1 As String
    Public strsql2 As String
    Public strsql23 As String
    Public oGeneralService1 As SAPbobsCOM.GeneralService
    Public oGeneralData1 As SAPbobsCOM.GeneralData
    Public oGeneralParams1 As SAPbobsCOM.GeneralDataParams
    Public oGeneralService2 As SAPbobsCOM.GeneralService
    Public oGeneralData2 As SAPbobsCOM.GeneralData
    Public oGeneralParams2 As SAPbobsCOM.GeneralDataParams

    Public xmlstring As String
    Public form_Name As String = ""
End Module
