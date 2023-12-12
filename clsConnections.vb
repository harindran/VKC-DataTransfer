Imports System.IO
Imports System.Data.SqlClient

Public Class clsConnections
    Dim Ss As New clsAddOn
    Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
    Dim MServerName As String = objMUtil.xServerName
    Public MDBName1 As String = objMUtil.xDBName
    Public MDBName2 As String = MDBName1
    Public DBSERVERNAME As String = ""
    Public series As Integer

    Public MUID As String = objMUtil.xUID
    Public MPWD As String = objMUtil.xPWD
    Public MLoginType As String = objMUtil.xLoginType
    Public MSAPUID As String = objMUtil.xSAPUser
    Public MSAPPWd As String = objMUtil.xSAPPwd
    Public Sub connection()
        ' dbconnection() 'Database Connection1(sqlconnection)
        CompanyConnection_notepad() 'Company Connection NOtepad(SAP)
        '  CompanyConnection() 'Company Connection2(SAP)
    End Sub

    Private Sub CompanyConnection_notepad()
        Try
            objcompany1 = New SAPbobsCOM.Company
            objcompany1.Server = MServerName.Trim.ToString
            objcompany1.LicenseServer = "SVKCTWR005:30000" '"SVKCTWR007:30000" ' "SVKCTWR005:30000"
            objcompany1.SLDServer = "SVKCTWR005:40000" '"SVKCTWR007:40000"  '"SVKCTWR005:30010"
            objcompany1.UseTrusted = False
            objcompany1.CompanyDB = MDBName1.Trim.ToString
            objcompany1.UserName = MSAPUID.Trim.ToString
            objcompany1.Password = MSAPPWd.Trim.ToString
            objcompany1.DbUserName = MUID.Trim.ToString
            objcompany1.DbPassword = MPWD.Trim.ToString
            objcompany1.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
            lretcode = objcompany1.Connect()
            If lretcode <> 0 Then
                MsgBox(objcompany1.GetLastErrorDescription)
            Else
                ' dbname()

            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub CompanyConnection(ByVal DBName As String)
        Try
            Dim objtransfer As New clsTransfer
            'DBSERVERNAME
            objFromCompany = New SAPbobsCOM.Company
            'objcompany2.Server = MServerName.Trim.ToString
            'objcompany2.LicenseServer = MServerName.Trim.ToString + ":30000" ' dtheader.Rows(0)("LSRV")
            objFromCompany.Server = MServerName.Trim.ToString
            objFromCompany.LicenseServer = "SVKCTWR005:30000" '"SVKCTWR007:30000"
            objFromCompany.SLDServer = "SVKCTWR005:40000" '"SVKCTWR007:40000"
            objFromCompany.UseTrusted = False
            objFromCompany.CompanyDB = DBName.Trim.ToString
            objFromCompany.UserName = MSAPUID.Trim.ToString
            objFromCompany.Password = MSAPPWd.Trim.ToString
            objFromCompany.DbUserName = MUID.Trim.ToString
            objFromCompany.DbPassword = MPWD.Trim.ToString
            objFromCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
            lretcode = objFromCompany.Connect()
            If lretcode <> 0 Then
                objtransfer.write_log("From DB Connection Failed")
                ' MsgBox(objcompany2.GetLastErrorDescription)
            Else
                objtransfer.write_log("From DB Connected!!!")
            End If
        Catch ex As Exception
            'MsgBox(ex.Message & vbLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub CompanyToConnection(ByVal DBName As String)
        Try
            'DBSERVERNAME
            objToCompany = New SAPbobsCOM.Company
            'objcompany2.Server = MServerName.Trim.ToString
            'objcompany2.LicenseServer = MServerName.Trim.ToString + ":30000" ' dtheader.Rows(0)("LSRV")
            objToCompany.Server = MServerName.Trim.ToString
            objToCompany.LicenseServer = "SVKCTWR005:30000" ' "SVKCTWR007:30000"
            objToCompany.SLDServer = "SVKCTWR005:40000" ' "SVKCTWR007:40000"
            objToCompany.UseTrusted = False
            objToCompany.CompanyDB = DBName.Trim.ToString
            objToCompany.UserName = MSAPUID.Trim.ToString
            objToCompany.Password = MSAPPWd.Trim.ToString
            objToCompany.DbUserName = MUID.Trim.ToString
            objToCompany.DbPassword = MPWD.Trim.ToString
            objToCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
            lretcode = objToCompany.Connect()
            If lretcode <> 0 Then
                ' MsgBox(objcompany2.GetLastErrorDescription)
            Else
            End If
        Catch ex As Exception
            'MsgBox(ex.Message & vbLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub dbconnection()
        Try
            If MLoginType.ToUpper.Trim <> "S" Then
                con = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName1 + "; INTEGRATED SECURITY = TRUE;")
            Else
                con = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName1 + "; USER ID=" + MUID + "; PASSWORD=" + MPWD + ";")
            End If
            con.Open()
        Catch ex As Exception
            MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
        End Try
    End Sub


    Public Sub dbname()
        strsql1 = "select U_dbname,U_DBSERVER from oadm"
        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs1.DoQuery(strsql1)
        If Objrs1.RecordCount > 0 Then
            MDBName2 = Objrs1.Fields.Item("U_dbname").Value
            DBSERVERNAME = Objrs1.Fields.Item("U_DBSERVER").Value
        End If
    End Sub

    Public Function GetSingleValue_SQL_dt(ByVal Str As String)

        If con.State = ConnectionState.Open Then
            'con.Dispose()
            con.Close()
        End If
        con.Open()

        Dim dtemd As New DataTable
        Dim dates As SqlDataAdapter
        Dim dr As SqlDataReader
        Try










            dates = New SqlDataAdapter(Str, con)
            'Dim commandBuilder As New SqlCommandBuilder(dates)

            dtemd = New DataTable
            'dtemd.Locale = System.Globalization.CultureInfo.InvariantCulture
            dates.Fill(dtemd)

            If con.State = ConnectionState.Open Then
                'con.Dispose()
                con.Close()
            End If
            con.Open()
            con.ResetStatistics()
            'con.
            If dtemd.Rows.Count = 0 Then Return dtemd

            Return dtemd  'objRS.Fields.Item(0).Value.ToString

        Catch ex As Exception
            Return dtemd
        End Try
    End Function
    Public Sub Fun_ErrorLog(ByVal Errorlog As String)

        strsql1 = Errorlog
        cmd = New SqlCommand(strsql1, con)
        cmd.ExecuteNonQuery()
        ''A\R Invoice_" & Replace(objcompany.GetLastErrorDescription, "'", "") & "'
    End Sub
    Public Function GetSingleValue_SQL_SQL(ByVal Str As String) As String

        If con.State = ConnectionState.Open Then
            'con.Dispose()
            con.Close()
        End If
        con.Open()
        Dim dtemd As New DataTable
        Dim dates As SqlDataAdapter
        Dim dr As SqlDataReader
        Try


            dates = New SqlDataAdapter(Str, con)
            Dim commandBuilder As New SqlCommandBuilder(dates)

            dtemd = New DataTable
            'dtemd.Locale = System.Globalization.CultureInfo.InvariantCulture
            dates.Fill(dtemd)

            If con.State = ConnectionState.Open Then
                'con.Dispose()
                con.Close()
            End If
            con.Open()
            con.ResetStatistics()
            'con.
            If dtemd.Rows.Count = 0 Then Return ""

            Return dtemd.Rows(0)(0).ToString  'objRS.Fields.Item(0).Value.ToString
        Catch ex As Exception
            Return ""
        End Try
    End Function
End Class
