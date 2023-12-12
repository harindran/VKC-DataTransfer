Imports System.Data.SqlClient
Public Class clsGlobalVariables

    Public MServerName As String
    Public MDBName As String
    Public MDDBMain As String
    Public MServerMainServer As String
    Public MServerNamedb As String
    Public MUID As String
    Public MPWD As String
    Public MLoginType As String
    Public MSAPUID As String
    Public MSAPPWd As String
    Public ConString As String
    Public MIPL_FocusColor As Color = Color.FromArgb(254, 240, 158)
    Public MIPL_LostFocusColor As Color = Color.White
    Public Finger As Integer
    Public MAINobjCompany As New SAPbobsCOM.Company
    Public dttempsinglevale As New DataTable
    Public dtrate As New DataTable


End Class
