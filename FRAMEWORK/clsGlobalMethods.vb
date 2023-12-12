Imports System.Data.SqlClient

Public Class clsGlobalMethods
    Dim objRS As SAPbobsCOM.Recordset
    Dim drow As DataRow
    Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
    Dim MServerName As String
    Dim MDBName As String
    Dim MUID As String
    Dim MPWD As String
    Dim MLoginType As String
    Dim ConString As String


    Public Sub New()
        MServerName = objMUtil.xServerName
        MDBName = objMUtil.xDBName
        MUID = objMUtil.xUID
        MPWD = objMUtil.xPWD
        MLoginType = objMUtil.xLoginType

        ''MServerName = objAddOn.objCompany.Server.ToString
        ''MDBName = objAddOn.objCompany.CompanyDB.ToString
        ''MUID = objAddOn.objCompany.DbUserName.ToString
        ''MPWD = objAddOn.objCompany.DbPassword
        ''MLoginType = IIf(objAddOn.objCompany.DbServerType = 6, "S", "W")
    End Sub

    Public Function ConvertToDatatable(ByVal Str As String, Optional ByVal SnoRequired As Boolean = False) As DataTable
        Dim dtConvert As New DataTable

        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(Str)

        dtConvert = New DataTable

        For i As Integer = 0 To objRS.Fields.Count - 1
            dtConvert.Columns.Add(objRS.Fields.Item(i).Name.ToString)
        Next
        If objRS.EoF Then Return dtConvert
        While Not objRS.EoF
            drow = dtConvert.NewRow
            For i As Integer = 0 To objRS.Fields.Count - 1
                drow.Item(i) = objRS.Fields.Item(i).Value.ToString
            Next
            dtConvert.Rows.Add(drow)
            objRS.MoveNext()
        End While
        Return dtConvert
    End Function

    Public Function GetSingleValue(ByVal Str As String) As String
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(Str)

        If objRS.EoF Then Return ""
        Return objRS.Fields.Item(0).Value.ToString
    End Function

    Public Function Connection() As Boolean
        Try
            If objMUtil.xLoginType.ToUpper.Trim <> "S" Then
                con = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; INTEGRATED SECURITY = TRUE;")
            Else
                con = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; USER ID=" + MUID + "; PASSWORD=" + MPWD + ";")
            End If
            con.Open()
            Return True
        Catch ex As Exception
            MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
            Return False
        End Try
    End Function

    Function funcGlobal_CheckValue_Available(ByVal strTableName As String, ByVal strFieldName As String, ByVal strValue As String, Optional ByVal strExceptField As String = "", Optional ByVal strExceptValue As String = "") As Boolean
        Dim str As String = "SELECT 1 FROM " + strTableName + " WHERE " + strFieldName + " = '" + strValue + "'"
        If strExceptField <> "" And strExceptValue <> "" Then
            str += " AND " + strExceptField + " NOT IN ('" + strExceptValue + "')"
        End If
        da = New SqlDataAdapter(str, con)
        dt = New DataTable
        da.Fill(dt)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
