Imports System.IO
Imports System.Data.SqlClient

Public Class clsTransfer
    Dim errorid As String
    Dim objconnection1 As New clsConnections
    Dim oss As New clsAddOn
    Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\DBInfo.ini", 20)
    Dim MainDB2 As String = objMUtil.xDBName
    Dim Errorlogs As String = ""
    Dim dt_time As String = ""
    Dim db, chatlog As String
    Dim fs As FileStream
    Dim objWriter As System.IO.StreamWriter
    Public Sub transfer_document()
        Dim code As String = ""
        Try
            Dim DBS As String = ""

            strsql23 = "select Code,ToDocType ,BaseEntry ,PostDB,DocType  from MIPLLOGS where  Flag ='N' order by code asc"
            da = New SqlDataAdapter(strsql23, con)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then

                For i As Integer = 0 To dt.Rows.Count - 1
                    If DBS = "FORTUNE_KINALOOR" Then Continue For
                    If i = 0 Or i = 1 Or i = 2 Or i = 3 Then Continue For
                    DBS = dt.Rows(i).Item("PostDB")
                    MDBName1 = dt.Rows(i).Item("PostDB")
                    code = dt.Rows(i).Item("Code")
                    write_log("From DB" & DBS)
                    objconnection1.CompanyConnection(MDBName1)
                    write_log("From DB Company Connected:" & DBS)
                    ' cmdDraftTopAYMENT()
                    oss.createTables()
                    If dt.Rows(i).Item("DocType") = "AR" Then
                        Fun_AR_To_GRPO(DBS, dt.Rows(i).Item("BaseEntry"), code)
                    ElseIf dt.Rows(i).Item("DocType") = "GI" Then
                        Fun_GI_To_GR(DBS, dt.Rows(i).Item("BaseEntry"), code)
                    End If
                    'objFromCompany.Disconnect()

                Next
            End If


        Catch ex As Exception
            strsql1 = "update mipllogs set Flag='N',Errorlog='EX" & Replace(ex.ToString, "'", "") & "' where code='" & code & "'"
            objconnection1.Fun_ErrorLog(strsql1)
            write_log(Replace(ex.ToString, "'", ""))
            'MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub cmdDraftTopAYMENT()


        Dim pDraft As SAPbobsCOM.Payments
        pDraft = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
        'pDraft.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
        Dim pOrder As SAPbobsCOM.Payments

        objFromCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        objFromCompany.XMLAsString = False
        pDraft.GetByKey(270)
        pDraft.SaveXML("D:\Prabhu\Pdrafts.xml")

        'Here you should add a code that will change the Object's
        'value from 112 (Drafts) to 17 (Orders) and also you should
        'remove the DocObjectCode node from the xml. You can use any
        'xml parser.
        '
        'Create a new order
        pOrder = objFromCompany.GetBusinessObjectFromXML("D:\Prabhu\Pdrafts.xml", 0)
        pOrder = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

        pOrder.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
        pOrder.TransferAccount = "12420008"
        pOrder.TransferSum = 1005
        pOrder.TransferReference = "PPS"

        Dim retvale As Long
        Dim retvale1 As Long
        retvale = pOrder.Add()
        If retvale <> 0 Then
            strsql1 = Replace(objToCompany.GetLastErrorDescription, "'", "")
            objconnection1.Fun_ErrorLog(strsql1)
        Else
            objFromCompany.GetNewObjectCode(retvale1)
        End If
    End Sub
    Private Sub PurchaseInvoice()
        Try
            strsql1 = "select DocEntry  from [@MIPLopch] where isnull(U_TRANSFER2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLOPCH")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLOPCH")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    oGeneralService2.Add(oGeneralData2)

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLOPCH] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Transfer()
        strsql1 = "select a.docentry,a.U_TRANDATE'Transdate' ,a.U_MODE'Transtype',b.U_BARCODE'transbarcode' ,'C' AS transstaus,a.U_TRANTOBRANCH'translocation',b.U_TOCOUNTER'transcounter',"
        strsql1 += " b.U_GROSSWEIGHT 'transwt',b.U_PIECES'TRANSPCS',a.U_SAPTRNNO'Docnum',a.U_EMPID'Userid',a.U_NODECOUNTER'NODEID'    From [@MIPLDGT] a join [@MIPLDGT1] b on a.DocEntry =b.DocEntry where  a.U_STATUS !='Cancelled' and   isnull(a.U_TRANSFER1,'')!='0' "
        Objrs3 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs3.DoQuery(strsql1)
        For i As Integer = 0 To Objrs3.RecordCount - 1
            TagTrans(Objrs3.Fields.Item("Transdate").Value.ToString, Objrs3.Fields.Item("Transtype").Value.ToString, Objrs3.Fields.Item("transbarcode").Value.ToString, Objrs3.Fields.Item("transstaus").Value.ToString, Objrs3.Fields.Item("translocation").Value.ToString, Objrs3.Fields.Item("transcounter").Value.ToString, Objrs3.Fields.Item("transwt").Value.ToString, Objrs3.Fields.Item("TRANSPCS").Value.ToString, Objrs3.Fields.Item("Docnum").Value.ToString, Objrs3.Fields.Item("Userid").Value.ToString, Objrs3.Fields.Item("NODEID").Value.ToString)
            strsql2 = " UPDATE [@MIPLDGT] SET U_TRANSFER1='" & errorid.ToString & "' where docentry='" & Objrs3.Fields.Item("docentry").Value.ToString & "'"
            Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            Objrs3.MoveNext()
        Next i
    End Sub

    Public Sub TagTrans(ByVal Transdate As Date, ByVal Transtype As String, ByVal transbarcode As String, ByVal transtaus As String, ByVal translocation As String, ByVal transcounter As String, ByVal transwt As Double, ByVal TRANSPCS As Double, ByVal Docnum As String, ByVal Userid As String, ByVal NODEID As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim estdocnum As String = ""
        Dim transfertype As String = ""
        errorid = ""
        Try
TRY_AGAIN:
            If Not objcompany1.InTransaction Then objcompany1.StartTransaction()
            oGeneralService = objcompany1.GetCompanyService.GetGeneralService("MIPLDTAGTRAN")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            strsql1 = "SELECT isnull(max(CONVERT(INT,ISNULL(DOCENTRY,0)))+1,0) MAXCODE FROM [@MIPLDTAGTRAN] with (NOLOCK)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)

            If Objrs1.Fields.Item("MAXCODE").Value = 0 Then
                oGeneralData.SetProperty("DocNum", 1)
            Else
                oGeneralData.SetProperty("DocNum", Objrs1.Fields.Item("MAXCODE").Value)
            End If
            If Transtype.ToString = "Transfer" Then
                transfertype = "GT"
            Else
                transfertype = "GR"
            End If

            oGeneralData.SetProperty("U_BARCODE", transbarcode.ToString)
            oGeneralData.SetProperty("U_TRANSDATE", CDate(Transdate.Date).ToString("yyyy-MM-dd"))
            oGeneralData.SetProperty("U_TRNSTYPE", transfertype.ToString)
            oGeneralData.SetProperty("U_TRANSLOCATION", translocation.ToString)
            oGeneralData.SetProperty("U_DOCNUM", Docnum.ToString)

            If transbarcode.Substring(0, 2) = "BT" Then
                strsql1 = "select U_PRODCODE From [@MIPLBAT] where U_BARCODE ='" & transbarcode.ToString & "' and U_COUNTERCODE='" & transcounter.ToString & "'" 'and U_BRANCHNAME='" & translocation.ToString & "'"
            Else
                strsql1 = "select U_PRODCODE From [@MIPLDTAG] where U_BARCODE ='" & transbarcode.ToString & "'"
            End If

            oGeneralData.SetProperty("U_TRANSPRODCODE", GetSingleValue(strsql1).ToString)
            strsql1 = "select U_METALTYPE  From [@MIPLIM] where U_PRODCODE ='" & GetSingleValue(strsql1).ToString & "'"
            oGeneralData.SetProperty("U_METALTYPE", GetSingleValue(strsql1).ToString)
            oGeneralData.SetProperty("U_TRANSTATUS", transtaus.ToString)
            oGeneralData.SetProperty("U_TRANSCOUNTER", transcounter.ToString)
            If transbarcode.Substring(0, 2) = "BT" Then
                strsql1 = "select U_SUBPRODCODE From [@MIPLBAT] where U_BARCODE ='" & transbarcode.ToString & "'and U_COUNTERCODE='" & transcounter.ToString & "'" ' and U_BRANCHNAME='" & translocation.ToString & "'"
                estdocnum = GetSingleValue(strsql1).ToString
                oGeneralData.SetProperty("U_TRANSWTPCS", Val(transwt))
                oGeneralData.SetProperty("U_TRANSPCS", Val(TRANSPCS))
            Else
                strsql1 = "select U_SUBPRODCODE From [@MIPLDTAG] where U_BARCODE ='" & transbarcode.ToString & "'"
                estdocnum = GetSingleValue(strsql1).ToString
                oGeneralData.SetProperty("U_TRANSWTPCS", Val(transwt))
                oGeneralData.SetProperty("U_TRANSPCS", Val(TRANSPCS))
            End If
            If Transtype.ToString = "TAG" Or Transtype.ToString = "GR" Or Transtype.ToString = "CR" Then
                oGeneralData.SetProperty("U_TRANSINWT", Val(transwt))
                oGeneralData.SetProperty("U_TRANSINPCS", Val(TRANSPCS))
            ElseIf Transtype.ToString = "SI" Then
                oGeneralData.SetProperty("U_TRANSOUTWT", Val(transwt))
                oGeneralData.SetProperty("U_TRANSOUTPCS", Val(TRANSPCS))
            ElseIf Transtype.ToString = "GT" Then
                oGeneralData.SetProperty("U_GROSSWT", Val(transwt))
                oGeneralData.SetProperty("U_PIECES", Val(TRANSPCS))
            End If
            oGeneralData.SetProperty("U_TRANSSUBPCODE", estdocnum)
            strsql1 = "(SELECT UnitName  FROM OWGT WHERE UnitCode=(SELECT InvntryUom  FROM OITM WHERE ITEMCODE='" & estdocnum.ToString & "'))"
            oGeneralData.SetProperty("U_TRANSUOM", GetSingleValue(strsql1).ToString)
            oGeneralData.SetProperty("U_TRANEMPID", Userid.ToString)
            oGeneralData.SetProperty("U_NODECOUNTER", NODEID.ToString)
            Dim PCS As Double = 0
            Dim WT As Double = 0
            strsql1 = "select SUM(U_PIECES)'PCS',SUM(U_WEIGHT)'WT'  From [@MIPLDTAG1] where U_BARCODE1 ='" & transbarcode.ToString & "' and U_BOM='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            If Val(Objrs1.RecordCount) > 0 Then
                PCS = Val(Objrs1.Fields.Item("PCS").Value)
                WT = Val(Objrs1.Fields.Item("WT").Value)
            End If
            If transfertype.ToString = "TAG" Or transfertype.ToString = "GR" Or transfertype.ToString = "CR" Then
                oGeneralData.SetProperty("U_TRANSSTINWT", Val(WT))
                oGeneralData.SetProperty("U_TRANSSTINPCS", Val(PCS))
            ElseIf Transtype.ToString = "SI" Then
                oGeneralData.SetProperty("U_TRANSSTOUTWT", Val(WT))
                oGeneralData.SetProperty("U_TRANSSTOUTPCS", Val(PCS))
            ElseIf transfertype.ToString = "GT" Then
            End If
            oGeneralService.Add(oGeneralData)
        Catch When Err.Number Like "*2038*"
            Threading.Thread.Sleep(10)
            If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            GoTo TRY_AGAIN
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            errorid = 1
            If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Exit Sub
        Finally
        End Try
        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        errorid = 0
    End Sub
    Public Function GetSingleValue(ByVal Str As String) As String
        Objrs4 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs4.DoQuery(Str)

        If Objrs4.EoF Then Return ""
        Return Objrs4.Fields.Item(0).Value.ToString
    End Function
  
    Private Sub MaterialReceipt12()
        Try
            Dim objrswo As SAPbobsCOM.Recordset
            strsql1 = "select DocEntry,U_receipttype  from [@MIPLOIGN] where isnull(U_TRANSFER2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLOIGN")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    If addgoodsreceipt(Objrs1.Fields.Item("Docentry").Value.ToString) Then
                    Else
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Exit Sub
                    End If
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLOIGN")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    oGeneralData2.SetProperty("U_RUNNO", "0")
                    oGeneralService2.Add(oGeneralData2)
                    If Objrs1.Fields.Item("U_receipttype").Value.ToString.ToUpper = "W" Then
                        strsql1 = "select U_workorderno,U_workorderlineno from " & MDBName1.ToString & "..[@MIPLign1] where docentry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        objrswo = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrswo.DoQuery(strsql1)
                        If objrswo.RecordCount > 0 Then
                            For wo As Integer = 0 To objrswo.RecordCount - 1
                                strsql1 = "update [@MIPLDWO1] set U_linestatus=(select U_linestatus from " & MDBName1.Trim.ToString & "..[@mipldwo1] where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "' and lineid='" & objrswo.Fields.Item("U_workorderlineno").Value.ToString & "'),"
                                strsql1 += vbCrLf + " U_openqty=(select U_openqty from " & MDBName1.Trim.ToString & "..[@mipldwo1] where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "' and lineid='" & objrswo.Fields.Item("U_workorderlineno").Value.ToString & "'),"
                                strsql1 += vbCrLf + " U_openweight=(select U_openweight from " & MDBName1.Trim.ToString & "..[@mipldwo1] where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "' and lineid='" & objrswo.Fields.Item("U_workorderlineno").Value.ToString & "')"
                                strsql1 += vbCrLf + " where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "' and lineid='" & objrswo.Fields.Item("U_workorderlineno").Value.ToString & "'"
                                strsql1 += vbCrLf + " update [@MIPLDWO] set U_status=(select U_status from " & MDBName1.Trim.ToString & "..[@mipldwo] where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "')"
                                strsql1 += vbCrLf + " where docentry='" & objrswo.Fields.Item("U_workorderno").Value.ToString & "'"
                                Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Objrs2.DoQuery(strsql1)
                                objrswo.MoveNext()
                            Next
                        End If
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLOIGN] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)

                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub MaterialReceipt21()
        Try
            strsql2 = "select DocEntry,u_DocEntry  from [@MIPLOIGN] where isnull(U_TRANSFER1,'')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLOIGN")
                oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                xmlstring = oGeneralData2.ToXMLString()
                Try
                    If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                    oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLOIGN")
                    oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData1.FromXMLString(xmlstring)
                    oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                    'oGeneralData1.SetProperty("U_RUNNO", "0")
                    oGeneralService1.Add(oGeneralData1)

                    strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLOIGN] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)
                    

                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    strsql1 = "Select docentry from  " & MDBName1.Trim.ToString & "..[@MIPLOIGN] where u_DocEntry='" & Objrs2.Fields.Item("u_DocEntry").Value.ToString & "'"
                    Objrs3 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs3.DoQuery(strsql1)

                    strsql23 = "update " & MDBName1.Trim.ToString & "..[@MIPLOIGN] set U_RUNNO='" & Objrs3.Fields.Item("DocEntry").Value.ToString & "' where u_DocEntry='" & Objrs2.Fields.Item("u_DocEntry").Value.ToString & "'"
                    Objrs4 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs4.DoQuery(strsql23)
                    Objrs2.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub LotStatus21()
        Try
            strsql2 = "SELECT DocEntry  fROM [@MIPLOIGN] WHERE U_DOCENTRY  IN(select U_DOCENTRY   from  " & MDBName1.Trim.ToString & "..[@MIPLOIGN] where U_LOTSTATUS ='C')"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                
                Try

                    strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLOIGN] set U_LOTSTATUS='C' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)



                    Objrs2.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub



    Private Sub LotStatus12()
        Try
            strsql1 = "select U_DOCENTRY  from " & MDBName1.Trim.ToString & "..[@MIPLOIGN] where "
            strsql1 += "DocEntry in(select U_GRNO  from " & MDBName1.Trim.ToString & "..[@MIPLDLOT] where U_STATUS='cancelled')"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1

                Try

                    strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLOIGN] set U_LOTSTATUS='O' where U_DOCENTRY='" & Objrs1.Fields.Item("U_DOCENTRY").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)



                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub MR_ReturnLotStatus21()
        Try
            strsql2 = "SELECT  DocEntry  fROM [@MIPLOIGN] WHERE U_DOCENTRY  IN(select U_DOCENTRY   from " & MDBName2.Trim.ToString & "..[@MIPLOIGN] where U_LOTSTATUS ='C')"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql2)
            For i As Integer = 0 To Objrs1.RecordCount - 1

                Try

                    strsql1 = "update " & MDBName1.Trim.ToString & "..[@MIPLOIGN] set U_LOTSTATUS='C' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql1)



                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub



    Private Function addgoodsreceipt(ByVal docentry As String)
        Try
            Dim objrs1 As SAPbobsCOM.Recordset
            Dim objrs11 As SAPbobsCOM.Recordset
            strsql1 = "select U_RECEIPTTYPE,U_ORDERNO,U_CARDCODE,U_CARDNAME,U_VENDORREFNO,U_VENDORSHTNAME,U_DOCNO,U_DOCENTRY,U_DOCDATE,U_series,U_POSTINGDATE,U_TRANSTYPE,"
            strsql1 += vbCrLf + " U_LOTSTATUS, U_OPENINGBALANCE, U_OPENINGWEIGHT, U_OPENINGWEISIL, U_NETWT, U_PUREWT, U_MAKINGCHRG, U_OTHERCHRG, U_TOTAL, U_EMPNAME, U_REMARKS"
            strsql1 += vbCrLf + " from [@MIPLOIGN] where Docentry='" & docentry.ToString & "'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim omaterialreceipt As SAPbobsCOM.Documents
                omaterialreceipt = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

                omaterialreceipt.UserFields.Fields.Item("U_WORKORDERTYPE").Value = Objrs1.Fields.Item("U_RECEIPTTYPE").Value.ToString
                If Objrs1.Fields.Item("U_RECEIPTTYPE").Value.ToString = "P" Then
                    omaterialreceipt.UserFields.Fields.Item("U_WORKORDERNO").Value = Objrs1.Fields.Item("U_ORDERNO").Value.ToString
                    omaterialreceipt.UserFields.Fields.Item("U_WORKORDERDATE").Value = Objrs1.Fields.Item("U_DOCDATE").Value
                ElseIf Objrs1.Fields.Item("U_RECEIPTTYPE").Value.ToString.ToUpper = "I" Then
                    omaterialreceipt.UserFields.Fields.Item("U_INVOICENO").Value = Objrs1.Fields.Item("U_ORDERNO").Value.ToString
                    omaterialreceipt.UserFields.Fields.Item("U_INVOICEDATE").Value = Objrs1.Fields.Item("U_DOCDATE").Value
                Else
                    omaterialreceipt.UserFields.Fields.Item("U_INVOICEDATE").Value = Objrs1.Fields.Item("U_DOCDATE").Value
                End If

                omaterialreceipt.UserFields.Fields.Item("U_VENDORCODE").Value = Objrs1.Fields.Item("U_CARDCODE").Value
                omaterialreceipt.UserFields.Fields.Item("U_VENDORNAME").Value = Objrs1.Fields.Item("U_CARDNAME").Value
                omaterialreceipt.UserFields.Fields.Item("U_VENDORSHORTNAME").Value = Objrs1.Fields.Item("U_VENDORSHTNAME").Value
                omaterialreceipt.UserFields.Fields.Item("U_VENDORREF").Value = Objrs1.Fields.Item("U_VENDORREFNO").Value
                omaterialreceipt.UserFields.Fields.Item("U_FORMTYPE").Value = "MRT1"
                omaterialreceipt.UserFields.Fields.Item("U_OPENINGBALANCE").Value = objrs1.Fields.Item("U_OPENINGBALANCE").Value.ToString
                omaterialreceipt.UserFields.Fields.Item("U_OPENINGWEIGHT").Value = objrs1.Fields.Item("U_OPENINGWEIGHT").Value.ToString
                omaterialreceipt.UserFields.Fields.Item("U_OPENINGWEISIL").Value = objrs1.Fields.Item("U_OPENINGWEISIL").Value.ToString
                omaterialreceipt.UserFields.Fields.Item("U_GSNO").Value = docentry.ToString
                omaterialreceipt.UserFields.Fields.Item("U_LOTSTATUS").Value = "C"
                'omaterialreceipt.Series = objGM.GetSingleValue(strSql)
                omaterialreceipt.DocDate = Objrs1.Fields.Item("U_POSTINGDATE").Value
                omaterialreceipt.UserFields.Fields.Item("U_TRANSACTIONTYPE").Value = Objrs1.Fields.Item("U_TRANSTYPE").Value
                omaterialreceipt.UserFields.Fields.Item("U_EMPLOYEENAME").Value = Objrs1.Fields.Item("U_EMPNAME").Value
                omaterialreceipt.Comments = Objrs1.Fields.Item("U_REMARKS").Value

                strsql1 = "select isnull(U_WORKORDERNO,'')'U_WORKORDERNO',isnull(U_WORKORDERLINENO,'')'U_WORKORDERLINENO',isnull(U_orderno,'')'U_orderno',"
                strsql1 += vbCrLf + " isnull(U_ORDERLINENO,'')'U_ORDERLINENO',ISNULL(CONVERT(VARCHAR,U_orderdate,103),'')'U_orderdate',U_PRODCODE,U_PRODNAME,U_SUBPRODCODE,U_SUBPRODNAME,ISNULL(U_UOM,'')'U_UOM',"
                strsql1 += vbCrLf + " convert(decimal(38,3),U_PIECES)U_PIECES,U_GROSSWEIGHT,U_NETWEIGHT,U_LESSWEIGHT,U_STONEAMOUNT,U_WASTAGETYPE,U_MCTYPE,U_WASTAGEPERPCS,U_WASTAGEPERCEN,U_WASTAGEGRAM,U_WASTAGEAMOUNT,U_PURITYPERCEN,"
                strsql1 += vbCrLf + " U_PUREWEIGHT,U_UNITPRICE,U_MAKINGCHARGE,U_MAKINGCHRGAMT,U_OTHERCHARGES,U_HALLMARKVENDOR,U_LINETOTAL,U_LOTSTATUS,isnull(U_REMARKS,'')'U_REMARKS',"
                strsql1 += vbCrLf + " U_WHSCODE, U_COUNTERNAME, U_BRANCH, U_BOMPRODUCT"
                strsql1 += vbCrLf + " from [@MIPLIGN1] where DocEntry='" & docentry.ToString & "'"
                Objrs11 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs11.DoQuery(strsql1)
                For j As Integer = 0 To Objrs11.RecordCount - 1
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WORKORDERNO").Value = Objrs11.Fields.Item("U_WORKORDERNO").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WORKORDERLINENO").Value = Objrs11.Fields.Item("U_WORKORDERLINENO").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_ORDERNO").Value = Objrs11.Fields.Item("U_orderno").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_ORDERLINENO").Value = Objrs11.Fields.Item("U_ORDERLINENO").Value
                    If Objrs11.Fields.Item("U_orderdate").Value.ToString <> "" Then omaterialreceipt.Lines.UserFields.Fields.Item("U_ORDERDATE").Value = CDate(Objrs11.Fields.Item("U_orderdate").Value).ToString("yyyy/MM/dd")
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_PRODCODE").Value = Objrs11.Fields.Item("U_PRODCODE").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_PRODNAME").Value = Objrs11.Fields.Item("U_PRODNAME").Value
                    omaterialreceipt.Lines.ItemCode = Objrs11.Fields.Item("U_SUBPRODCODE").Value

                    omaterialreceipt.Lines.UserFields.Fields.Item("U_BOMPRODUCT").Value = Objrs11.Fields.Item("U_BOMPRODUCT").Value
                    If Objrs11.Fields.Item("U_UOM").Value.ToString.ToUpper = "PIECES" Then
                        If Val(Objrs11.Fields.Item("U_PIECES").Value) = 0 Then
                            omaterialreceipt.Lines.Quantity = 1
                        Else
                            omaterialreceipt.Lines.Quantity = Objrs11.Fields.Item("U_PIECES").Value
                        End If
                    Else
                        If Val(Objrs11.Fields.Item("U_GROSSWEIGHT").Value) = 0 Then
                            omaterialreceipt.Lines.Quantity = 1
                        Else
                            omaterialreceipt.Lines.Quantity = Objrs11.Fields.Item("U_GROSSWEIGHT").Value
                        End If
                    End If

                    omaterialreceipt.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value = Val(Objrs11.Fields.Item("U_GROSSWEIGHT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value = Val(Objrs11.Fields.Item("U_LESSWEIGHT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value = Val(Objrs11.Fields.Item("U_NETWEIGHT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WASTAGEPERPCS").Value = Val(Objrs11.Fields.Item("U_WASTAGEPERPCS").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value = Val(Objrs11.Fields.Item("U_WASTAGEPERCEN").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value = Val(Objrs11.Fields.Item("U_WASTAGEGRAM").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WASTAGEAMOUNT").Value = Val(Objrs11.Fields.Item("U_WASTAGEAMOUNT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_WASTAGETYPE").Value = Objrs11.Fields.Item("U_WASTAGETYPE").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_MAKINGCHRGTYPE").Value = Objrs11.Fields.Item("U_MCTYPE").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_STONECHRGAMT").Value = Val(Objrs11.Fields.Item("U_STONEAMOUNT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_PURITYPERCEN").Value = Val(Objrs11.Fields.Item("U_PURITYPERCEN").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_PUREWEIGHT").Value = Val(Objrs11.Fields.Item("U_PUREWEIGHT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_PIECES").Value = objrs11.Fields.Item("U_PIECES").Value.ToString

                    omaterialreceipt.Lines.UserFields.Fields.Item("U_UNITPRICE").Value = Val(Objrs11.Fields.Item("U_UNITPRICE").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_MAKINGCHARGES").Value = Val(Objrs11.Fields.Item("U_MAKINGCHARGE").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_MAKINGCHRGAMT").Value = Val(Objrs11.Fields.Item("U_MAKINGCHRGAMT").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_OTHERCHARGES").Value = Val(objrs11.Fields.Item("U_OTHERCHARGES").Value)
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_HALLMARKVENDOR").Value = objrs11.Fields.Item("U_HALLMARKVENDOR").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_LOTSTATUS").Value = "C"

                    omaterialreceipt.Lines.WarehouseCode = Objrs11.Fields.Item("U_WHSCODE").Value
                    omaterialreceipt.Lines.UnitPrice = 0
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_OTHERGS").Value = Objrs11.Fields.Item("U_OTHERCHARGES").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value = Objrs11.Fields.Item("U_COUNTERNAME").Value
                    omaterialreceipt.Lines.LocationCode = Objrs11.Fields.Item("U_BRANCH").Value
                    omaterialreceipt.Lines.UserFields.Fields.Item("U_REMARKS").Value = Objrs11.Fields.Item("U_REMARKS").Value
                    omaterialreceipt.Lines.Add()
                    Objrs11.MoveNext()
                Next
                lretcode = omaterialreceipt.Add()
                If lretcode <> 0 Then
                    MsgBox(objcompany2.GetLastErrorDescription)
                    Return False
                Else
                    'objcompany2.GetNewObjectCode(goodsreceiptno)
                End If
                Objrs1.MoveNext()
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
        Return True
    End Function

    Private Sub Lot12()
        Try
            strsql1 = "select DocEntry,U_Lotno from [@mipldlot] where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDLOT")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDLOT")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                    'oGeneralData2.SetProperty("U_STATUS", "Closed")
                    'oGeneralData2.SetProperty("U_TAGSTATUS", "C")
                    strsql2 = "select 1 from [@MIPLDLOT] where U_LOTNO='" & Objrs1.Fields.Item("U_LOTNO").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs1.Fields.Item("Docentry").Value)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")

                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDLOT] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub LOT21()
        Try
            strsql2 = "select DocEntry ,U_LOTNO from [@MIPLDLOT] where isnull(U_TRANSFER1,'')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDLOT")
                oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                xmlstring = oGeneralData2.ToXMLString()
                Try
                    If Not objcompany1.InTransaction Then objcompany1.StartTransaction()
                    oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDLOT")
                    oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    'oGeneralData1.SetProperty("U_STATUS", "Closed")
                    'oGeneralData1.SetProperty("U_TAGSTATUS", "C")

                    strsql1 = "select 1 from [@MIPLDLOT] where U_LOTNO='" & Objrs2.Fields.Item("U_LOTNO").Value.ToString & "'"
                    Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)
                    If Objrs1.RecordCount = 0 Then
                        oGeneralData1.FromXMLString(xmlstring)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Add(oGeneralData1)
                    Else
                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs2.Fields.Item("Docentry").Value)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams2)

                        oGeneralData1.FromXMLString(xmlstring)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")

                        oGeneralService1.Update(oGeneralData1)
                    End If

                    strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLDLOT] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)

                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs2.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Tag12()
        Try
            strsql1 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_TAGDATE,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS from [@mipldtag] where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Try
                    strsql2 = "select docentry,U_barcode,U_FORMBASISNO from [@mipldtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_barcode").Value.ToString & "' and U_FORMBASISNO='" & Objrs1.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()
                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Add(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs1.MoveNext()
                    Else
                        'MsgBox(Objrs1.Fields.Item("date").Value)

                        ' strsql2 = ""

                        strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs1.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs1.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs1.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs1.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='" & Objrs1.Fields.Item("U_ESTSTATUS").Value.ToString & "',U_ESTOPENPCS='" & Objrs1.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs1.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs2.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        Objrs1.MoveNext()
                    End If
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
             Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub EstTag12()
        Try
            strsql1 = "select b.U_TAGNO,a.U_DOCSTATUS    From [@MIPLDSE] a join [@MIPLDSE1] b on a.DocEntry =b.DocEntry where b.u_tagno not like ('%BT%') and U_DOCSTATUS='OPEN'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Try
                    If Objrs1.Fields.Item("U_DOCSTATUS").Value.ToString = "CANCELLED" Then
                        strsql2 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_TAGNO").Value.ToString & "' "
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)

                        strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs2.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs2.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs2.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs2.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='OPEN',U_ESTSTATUS ='OPEN',U_ESTOPENPCS='" & Objrs2.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs2.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs2.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)

                        Objrs1.MoveNext()
                    Else
                        strsql2 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_TAGNO").Value.ToString & "' "
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)

                        strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs2.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs2.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs2.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs2.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs2.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='CLOSED',U_ESTOPENPCS='" & Objrs2.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs2.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs2.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        Objrs1.MoveNext()
                    End If
                    

                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub EstTag21()
        Try
            strsql2 = "select b.U_TAGNO,a.U_DOCSTATUS    From [@MIPLDSE] a join [@MIPLDSE1] b on a.DocEntry =b.DocEntry "
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                Try
                    If Objrs2.Fields.Item("U_DOCSTATUS").Value.ToString = "CANCELLED" Then
                        strsql1 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_TAGNO").Value.ToString & "' "
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        strsql1 = "update " & MDBName1.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs1.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs1.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs1.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='OPEN',U_ESTSTATUS ='OPEN',U_ESTOPENPCS='" & Objrs1.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs1.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs1.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        Objrs2.MoveNext()
                    Else
                        strsql1 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_TAGNO").Value.ToString & "' "
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        strsql1 = "update " & MDBName1.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs1.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs1.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs1.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs1.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='Closed',U_ESTOPENPCS='" & Objrs1.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs1.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs1.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        Objrs2.MoveNext()
                    End If


                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Tag21()
        Try
            strsql2 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where isnull(U_TRANSFER1,'')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                Try
                    strsql1 = "select docentry,U_barcode,U_FORMBASISNO from [@mipldtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs2.Fields.Item("U_barcode").Value.ToString & "' and U_FORMBASISNO='" & Objrs2.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                    Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)
                    If Objrs1.RecordCount = 0 Then
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData1.FromXMLString(xmlstring)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        'oGeneralData1.SetProperty("U_TAGSTATUS", "Closed")
                        oGeneralService1.Add(oGeneralData1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@mipldtag] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        strsql1 = " update " & MDBName1.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs2.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_plusdate='" & Microsoft.VisualBasic.Right(Objrs2.Fields.Item("date").Value, 4).ToString + Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(Objrs2.Fields.Item("date").Value, 5).ToString), 2).ToString + Microsoft.VisualBasic.Left(Objrs2.Fields.Item("date").Value, 2).ToString & "',U_ORDERNO='" & Objrs2.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs2.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs2.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='" & Objrs2.Fields.Item("U_ESTSTATUS").Value.ToString & "',U_ESTOPENPCS='" & Objrs2.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs2.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  docentry='" & Objrs1.Fields.Item("docentry").Value.ToString & "'"
                        'strsql1 += vbCrLf + "update " & MDBName2.Trim.ToString & "..[@mipldtag] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs2.MoveNext()
                    Else

                        strsql1 = " update " & MDBName1.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs2.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_plusdate='" & Microsoft.VisualBasic.Right(Objrs2.Fields.Item("date").Value, 4).ToString + Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(Objrs2.Fields.Item("date").Value, 5).ToString), 2).ToString + Microsoft.VisualBasic.Left(Objrs2.Fields.Item("date").Value, 2).ToString & "',U_ORDERNO='" & Objrs2.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs2.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs2.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='" & Objrs2.Fields.Item("U_ESTSTATUS").Value.ToString & "',U_ESTOPENPCS='" & Objrs2.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs2.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  docentry='" & Objrs1.Fields.Item("docentry").Value.ToString & "'"
                        'strsql1 += vbCrLf + "update " & MDBName2.Trim.ToString & "..[@mipldtag] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        strsql1 = " update " & MDBName2.Trim.ToString & "..[@mipldtag] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)
                        Objrs2.MoveNext()
                    End If
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub invoice()
        Try
            Dim Docentry As String
            strsql1 = "select DocEntry from oinv where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim oinvoice1 As SAPbobsCOM.Documents
                oinvoice1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oinvoice1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)
                Docentry = Objrs1.Fields.Item("DocEntry").Value.ToString
                oinvoice1.SaveXML("C:\Users\praveen.e\Desktop\XMl\" & Objrs1.Fields.Item("DocEntry").Value.ToString & ".xml")
                'oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("OINV")
                'oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                'oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                'oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                'oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                'xmlstring = oGeneralData1.ToXMLString()

                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    Dim oinvoice2 As SAPbobsCOM.Documents
                    oinvoice2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    objcompany2.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    oinvoice2 = objcompany2.GetBusinessObjectFromXML("C:\Users\praveen.e\Desktop\XMl\" & Objrs1.Fields.Item("DocEntry").Value.ToString & ".xml", 0)
                    'oinvoice2 = objcompany2.GetBusinessObjectFromXML(String.Format("C:\Users\praveen.e\Desktop\XMl\" & Objrs1.Fields.Item("DocEntry").Value.ToString & ".xml", Docentry), 0)
                    oinvoice2.Expenses.SetCurrentLine(0)
                    oinvoice2.Expenses.ExpenseCode = 1
                    oinvoice2.Expenses.Remarks = "te"
                    oinvoice2.Expenses.TaxCode = "ZEROTAX"
                    oinvoice2.Expenses.LineTotal = 10
                    oinvoice2.Expenses.Add()

                    lretcode = oinvoice2.Add()
                    'oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("OINV")
                    'oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    'oGeneralData2.FromXMLString(xmlstring)
                    'oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    'oGeneralService2.Add(oGeneralData2)
                    If (lretcode) = 0 Then
                        strsql2 = "update " & MDBName1.Trim.ToString & "..OINV set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs1.MoveNext()
                    Else
                        'MsgBox(lretcode)
                        'MsgBox(objcompany2.GetLastErrorDescription)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        'Exit Sub
                        Objrs1.MoveNext()
                    End If

                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(objcompany2.GetLastErrorDescription)
                End Try
            Next
        Catch ex As Exception
            MsgBox(objcompany1.GetLastErrorDescription)
        End Try
    End Sub

    Private Sub WorkOrder12()
        Try
            strsql1 = "select DocEntry from [@MIPLDWO] where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDWO")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDWO")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    oGeneralService2.Add(oGeneralData2)

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDWO] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub invoice12()
        ''Add Sales Invoice
        Try

            strsql1 = "select DocEntry from oinv where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim oinvoice1 As SAPbobsCOM.Documents
                oinvoice1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oinvoice1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim oinvoice2 As SAPbobsCOM.Documents
                oinvoice2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oinvoice2.UserFields.Fields.Item("U_FORMTYPE").Value = oinvoice1.UserFields.Fields.Item("U_FORMTYPE").Value
                oinvoice2.UserFields.Fields.Item("U_ESTTYPE").Value = oinvoice1.UserFields.Fields.Item("U_ESTTYPE").Value
                oinvoice2.UserFields.Fields.Item("U_ESTNO").Value = oinvoice1.UserFields.Fields.Item("U_ESTNO").Value
                oinvoice2.UserFields.Fields.Item("U_ESTDATE").Value = oinvoice1.UserFields.Fields.Item("U_ESTDATE").Value
                oinvoice2.UserFields.Fields.Item("U_CASHCOUNTER").Value = oinvoice1.UserFields.Fields.Item("U_CASHCOUNTER").Value
                oinvoice2.UserFields.Fields.Item("U_PRODUCT").Value = oinvoice1.UserFields.Fields.Item("U_PRODUCT").Value
                oinvoice2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                oinvoice2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                oinvoice2.UserFields.Fields.Item("U_TRANSFER3").Value = "1"

                oinvoice2.UserFields.Fields.Item("U_DELIVERY").Value = oinvoice1.UserFields.Fields.Item("U_DELIVERY").Value
                oinvoice2.UserFields.Fields.Item("U_REMARKS").Value = oinvoice1.UserFields.Fields.Item("U_REMARKS").Value
                oinvoice2.UserFields.Fields.Item("U_PRIVILEGECARDNO").Value = oinvoice1.UserFields.Fields.Item("U_PRIVILEGECARDNO").Value
                oinvoice2.CardCode = oinvoice1.CardCode
                oinvoice2.UserFields.Fields.Item("U_EMPLOYEENAME").Value = oinvoice1.UserFields.Fields.Item("U_EMPLOYEENAME").Value
                oinvoice2.Address = oinvoice1.Address
                oinvoice2.NumAtCard = oinvoice1.NumAtCard
                oinvoice2.ControlAccount = oinvoice1.ControlAccount

                oinvoice2.DocDate = oinvoice1.DocDate
                oinvoice2.DocDueDate = oinvoice1.DocDueDate
                oinvoice2.TaxDate = oinvoice1.TaxDate
                oinvoice2.Series = oinvoice1.Series


                oinvoice2.UserFields.Fields.Item("U_BROKERGCODE").Value = oinvoice1.UserFields.Fields.Item("U_BROKERGCODE").Value
                oinvoice2.UserFields.Fields.Item("U_BROKERGNAME").Value = oinvoice1.UserFields.Fields.Item("U_BROKERGNAME").Value
                oinvoice2.UserFields.Fields.Item("U_REASONOGFRE").Value = oinvoice1.UserFields.Fields.Item("U_REASONOGFRE").Value
                ''Series,Doc No Added Automatically 

                For j As Integer = 0 To oinvoice1.Lines.Count - 1
                    oinvoice2.Lines.UserFields.Fields.Item("U_TAGNO").Value = oinvoice1.Lines.UserFields.Fields.Item("U_TAGNO").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_PRODCODE").Value = oinvoice1.Lines.UserFields.Fields.Item("U_PRODCODE").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_PRODNAME").Value = oinvoice1.Lines.UserFields.Fields.Item("U_PRODNAME").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_PIECES").Value = oinvoice1.Lines.UserFields.Fields.Item("U_PIECES").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_STONEWEIGHT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_STONEWEIGHT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value = oinvoice1.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value = oinvoice1.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value = oinvoice1.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_WASTAGEAMT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_WASTAGEAMT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_STONECHRGAMT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_STONECHRGAMT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_MAKINGCHRG").Value = oinvoice1.Lines.UserFields.Fields.Item("U_MAKINGCHRG").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_HALLMARKCHRG").Value = oinvoice1.Lines.UserFields.Fields.Item("U_HALLMARKCHRG").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_UNITPRICE").Value = oinvoice1.Lines.UserFields.Fields.Item("U_UNITPRICE").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGR1").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGR1").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT1").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT1").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGR2").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGR2").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT2").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT2").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGR3").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGR3").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT3").Value = oinvoice1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT3").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_REMARKS").Value = oinvoice1.Lines.UserFields.Fields.Item("U_REMARKS").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_DISCOUNTAMT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_DISCOUNTAMT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_DISCOUNTAMT1").Value = oinvoice1.Lines.UserFields.Fields.Item("U_DISCOUNTAMT1").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_DISCOUNTAMT2").Value = oinvoice1.Lines.UserFields.Fields.Item("U_DISCOUNTAMT2").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_DISCOUNTAMT3").Value = oinvoice1.Lines.UserFields.Fields.Item("U_DISCOUNTAMT3").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_DELIVERY").Value = oinvoice1.Lines.UserFields.Fields.Item("U_DELIVERY").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_ESTNO").Value = oinvoice1.Lines.UserFields.Fields.Item("U_ESTNO").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_LINENO").Value = oinvoice1.Lines.UserFields.Fields.Item("U_LINENO").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_ORDERNO").Value = oinvoice1.Lines.UserFields.Fields.Item("U_ORDERNO").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_BOMPRODUCT").Value = oinvoice1.Lines.UserFields.Fields.Item("U_BOMPRODUCT").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_ITEMCODE").Value = oinvoice1.Lines.UserFields.Fields.Item("U_ITEMCODE").Value
                    oinvoice2.Lines.UserFields.Fields.Item("U_ITEMNAME").Value = oinvoice1.Lines.UserFields.Fields.Item("U_ITEMNAME").Value

                    oinvoice2.Lines.ItemCode = oinvoice1.Lines.ItemCode
                    oinvoice2.Lines.Quantity = oinvoice1.Lines.Quantity
                    oinvoice2.Lines.UnitPrice = oinvoice1.Lines.UnitPrice

                    oinvoice2.Lines.TaxCode = oinvoice1.Lines.TaxCode
                    oinvoice2.Lines.LineTotal = oinvoice1.Lines.LineTotal
                    oinvoice2.Lines.WarehouseCode = oinvoice1.Lines.WarehouseCode
                    oinvoice2.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value = oinvoice1.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value
                    oinvoice2.Lines.LocationCode = oinvoice1.Lines.LocationCode
                    ''End Standard Fields

                    oinvoice2.Lines.Add()
                Next
                'For i As Integer = 0 To 2
                '    strSql = "select DocEntry  from dpi1 where BaseEntry='" & dgvadjustment.Rows(i).Cells.Item("Docentry").Value & "'"
                '    objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '    objrs.DoQuery(strSql)
                '    If objrs.RecordCount > 0 Then
                '        For J = 0 To objrs.RecordCount - 1
                '            Dim dpi As SAPbobsCOM.DownPaymentsToDraw
                '            dpi = objarinvoice.DownPaymentsToDraw
                '            dpi.DocEntry = objrs.Fields.Item("DocEntry").Value
                '            dpi.Add()
                '            objrs.MoveNext()
                '        Next
                '    End If
                'Next
                oinvoice2.Rounding = oinvoice1.Rounding
                oinvoice2.RoundingDiffAmount = oinvoice1.RoundingDiffAmount
                For fr As Integer = 0 To oinvoice1.Expenses.Count - 1
                    If oinvoice1.Expenses.LineTotal > 0 Then
                        oinvoice2.Expenses.ExpenseCode = oinvoice1.Expenses.ExpenseCode
                        oinvoice2.Expenses.Remarks = oinvoice1.Expenses.Remarks
                        oinvoice2.Expenses.TaxCode = oinvoice1.Expenses.TaxCode
                        oinvoice2.Expenses.LineTotal = oinvoice1.Expenses.LineTotal
                        oinvoice2.Expenses.Add()
                    End If
                Next
                lretcode = oinvoice2.Add
                If lretcode <> 0 Then
                    'MsgBox(objcompany2.GetLastErrorDescription)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..OINV set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Goodsretrun12_metalsettlement12()
        Try
            strsql1 = "select DocEntry from oige where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)

            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim oissue1 As SAPbobsCOM.Documents
                oissue1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                oissue1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim oissue2 As SAPbobsCOM.Documents
                oissue2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

                oissue2.UserFields.Fields.Item("U_ISSUETYPE").Value = oissue1.UserFields.Fields.Item("U_ISSUETYPE").Value
                oissue2.UserFields.Fields.Item("U_ISSUEDATE").Value = oissue1.UserFields.Fields.Item("U_ISSUEDATE").Value
                oissue2.UserFields.Fields.Item("U_RMRDOCNO").Value = oissue1.UserFields.Fields.Item("U_RMRDOCNO").Value
                oissue2.UserFields.Fields.Item("U_FORMTYPE").Value = oissue1.UserFields.Fields.Item("U_FORMTYPE").Value

                oissue2.UserFields.Fields.Item("U_GSNO").Value = oissue1.UserFields.Fields.Item("U_GSNO").Value

                oissue2.UserFields.Fields.Item("U_VENDORCODE").Value = oissue1.UserFields.Fields.Item("U_VENDORCODE").Value
                oissue2.UserFields.Fields.Item("U_VENDORNAME").Value = oissue1.UserFields.Fields.Item("U_VENDORNAME").Value
                oissue2.UserFields.Fields.Item("U_VENDORSHORTNAME").Value = oissue1.UserFields.Fields.Item("U_VENDORSHORTNAME").Value
                oissue2.UserFields.Fields.Item("U_VENDORREF").Value = oissue1.UserFields.Fields.Item("U_VENDORREF").Value
                oissue2.Series = oissue1.Series
                oissue2.DocDate = oissue1.DocDate
                oissue2.DocDueDate = oissue1.DocDueDate
                oissue2.DocumentsOwner = oissue1.DocumentsOwner

                oissue2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                oissue2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                oissue2.UserFields.Fields.Item("U_TRANSFER3").Value = oissue1.UserFields.Fields.Item("U_TRANSFER3").Value
                oissue2.UserFields.Fields.Item("U_OPENINGBALANCE").Value = oissue1.UserFields.Fields.Item("U_OPENINGBALANCE").Value
                oissue2.UserFields.Fields.Item("U_OPENINGWEIGHT").Value = oissue1.UserFields.Fields.Item("U_OPENINGWEIGHT").Value

                For j As Integer = 0 To oissue1.Lines.Count - 1
                    oissue2.Lines.UserFields.Fields.Item("U_LINENO").Value = oissue1.Lines.UserFields.Fields.Item("U_LINENO").Value
                    oissue2.Lines.UserFields.Fields.Item("U_TAGNO").Value = oissue1.Lines.UserFields.Fields.Item("U_TAGNO").Value
                    oissue2.Lines.UserFields.Fields.Item("U_PRODCODE").Value = oissue1.Lines.UserFields.Fields.Item("U_PRODCODE").Value
                    oissue2.Lines.UserFields.Fields.Item("U_PRODNAME").Value = oissue1.Lines.UserFields.Fields.Item("U_PRODNAME").Value
                    oissue2.Lines.ItemCode = oissue1.Lines.ItemCode
                    oissue2.Lines.Quantity = oissue1.Lines.Quantity

                    oissue2.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value = oissue1.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value = oissue1.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value = oissue1.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_WASTAGEPERPCS").Value = oissue1.Lines.UserFields.Fields.Item("U_WASTAGEPERPCS").Value
                    oissue2.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value = oissue1.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value
                    oissue2.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value = oissue1.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value
                    oissue2.Lines.UserFields.Fields.Item("U_WASTAGEAMOUNT").Value = oissue1.Lines.UserFields.Fields.Item("U_WASTAGEAMOUNT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_WASTAGETYPE").Value = oissue1.Lines.UserFields.Fields.Item("U_WASTAGETYPE").Value
                    oissue2.Lines.UserFields.Fields.Item("U_MAKINGCHRGTYPE").Value = oissue1.Lines.UserFields.Fields.Item("U_MAKINGCHRGTYPE").Value
                    oissue2.Lines.UserFields.Fields.Item("U_STONECHRGAMT").Value = oissue1.Lines.UserFields.Fields.Item("U_STONECHRGAMT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_PURITYPERCEN").Value = oissue1.Lines.UserFields.Fields.Item("U_PURITYPERCEN").Value
                    oissue2.Lines.UserFields.Fields.Item("U_PUREWEIGHT").Value = oissue1.Lines.UserFields.Fields.Item("U_PUREWEIGHT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_ALLOYWEIGHT").Value = oissue1.Lines.UserFields.Fields.Item("U_ALLOYWEIGHT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_PIECES").Value = oissue1.Lines.UserFields.Fields.Item("U_PIECES").Value
                    oissue2.Lines.UserFields.Fields.Item("U_UNITPRICE").Value = oissue1.Lines.UserFields.Fields.Item("U_UNITPRICE").Value
                    oissue2.Lines.UserFields.Fields.Item("U_MAKINGCHARGES").Value = oissue1.Lines.UserFields.Fields.Item("U_MAKINGCHARGES").Value
                    oissue2.Lines.UserFields.Fields.Item("U_MAKINGCHRGAMT").Value = oissue1.Lines.UserFields.Fields.Item("U_MAKINGCHRGAMT").Value
                    oissue2.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value = oissue1.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value
                    oissue2.Lines.WarehouseCode = oissue1.Lines.WarehouseCode
                    oissue2.Lines.LineTotal = oissue1.Lines.LineTotal
                    oissue2.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value = oissue1.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value
                    oissue2.Lines.LocationCode = oissue1.Lines.LocationCode

                    oissue2.Lines.Add()
                Next
                lretcode = oissue2.Add
                If lretcode <> 0 Then
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(objcompany2.GetLastErrorDescription)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..oige set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub inventrytransfer12()
        Try
            strsql1 = "select DocEntry from owtr where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)

            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim oinventrytransfer1 As SAPbobsCOM.StockTransfer
                oinventrytransfer1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                oinventrytransfer1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim oinventrytransfer2 As SAPbobsCOM.StockTransfer
                oinventrytransfer2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                oinventrytransfer2.UserFields.Fields.Item("U_FORMTYPE").Value = oinventrytransfer1.UserFields.Fields.Item("U_FORMTYPE").Value
                oinventrytransfer2.UserFields.Fields.Item("U_TRANSACTIONTYPE").Value = oinventrytransfer1.UserFields.Fields.Item("U_TRANSACTIONTYPE").Value
                oinventrytransfer2.UserFields.Fields.Item("U_FROMBRANCH").Value = oinventrytransfer1.UserFields.Fields.Item("U_FROMBRANCH").Value
                oinventrytransfer2.UserFields.Fields.Item("U_TOBRANCH").Value = oinventrytransfer1.UserFields.Fields.Item("U_TOBRANCH").Value
                oinventrytransfer2.UserFields.Fields.Item("U_JEWELLITEMS").Value = oinventrytransfer1.UserFields.Fields.Item("U_JEWELLITEMS").Value

                oinventrytransfer2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                oinventrytransfer2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                oinventrytransfer2.UserFields.Fields.Item("U_TRANSFER3").Value = "1"

                oinventrytransfer2.UserFields.Fields.Item("U_GINO").Value = oinventrytransfer1.UserFields.Fields.Item("U_GINO").Value
                oinventrytransfer2.UserFields.Fields.Item("U_PRODUCT").Value = oinventrytransfer1.UserFields.Fields.Item("U_PRODUCT").Value

                oinventrytransfer2.Comments = oinventrytransfer1.Comments
                oinventrytransfer2.JournalMemo = oinventrytransfer1.JournalMemo

                oinventrytransfer2.FromWarehouse = oinventrytransfer1.FromWarehouse
                For j As Integer = 0 To oinventrytransfer1.Lines.Count - 1
                    oinventrytransfer2.Lines.ItemCode = oinventrytransfer1.Lines.ItemCode
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_PRODCODE").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_PRODCODE").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_PRODNAME").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_PRODNAME").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_REFNO").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_REFNO").Value
                    oinventrytransfer2.Lines.WarehouseCode = oinventrytransfer1.Lines.WarehouseCode
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_REPAIRTYPE").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_REPAIRTYPE").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_TAGORIGINAL").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_TAGORIGINAL").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_TAGORIGINALNAME").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_TAGORIGINALNAME").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_ITEMCODE").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_ITEMCODE").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_ITEMNAME").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_ITEMNAME").Value

                    oinventrytransfer2.Lines.Quantity = oinventrytransfer1.Lines.Quantity

                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_TAGNO").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_TAGNO").Value
                    oinventrytransfer2.Lines.UserFields.Fields.Item("U_VISILBLE").Value = oinventrytransfer1.Lines.UserFields.Fields.Item("U_VISILBLE").Value

                    oinventrytransfer2.Lines.Add()
                Next
                lretcode = oinventrytransfer2.Add
                If lretcode <> 0 Then
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    ' MsgBox(objcompany2.GetLastErrorDescription)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..owtr set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SalesEstimation12()

        Try
            'UpdateProcessTobeDone("5 : " + form_Name)
            strsql1 = "select DocEntry,U_dailyno,U_branch,U_docstatus,convert(varchar(10),U_ESTDATE,120)'date'   from [@MIPLDSE] where isnull(U_TRANSFER2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            'UpdateProcessTobeDone("Count : " + Objrs1.RecordCount.ToString + " - " + form_Name)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                'UpdateProcessTobeDone("For : " + i.ToString + " - " + form_Name)
                Try
                    strsql2 = "select DocEntry,DocNum,Series  from [@MIPLDSE] where U_branch='" & Objrs1.Fields.Item("U_branch").Value.ToString & "' and U_DAILYNO='" & Objrs1.Fields.Item("U_dailyno").Value.ToString & "' and U_ESTDATE='" & (Objrs1.Fields.Item("date").Value).ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    'UpdateProcessTobeDone("Count1 : " + Objrs2.RecordCount.ToString + " - " + form_Name)
                    'UpdateProcessTobeDone("DocEntry : " + Objrs1.Fields.Item("DocEntry").Value.ToString + " - " + form_Name)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)
                        'UpdateProcessTobeDone("6 : " + form_Name)
                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDSE' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_branch").Value.ToString & "'"
                        Objrs5 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        SERIES = Objrs5.Fields.Item("Series").Value
                        oGeneralData2.SetProperty("Series", SERIES)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")

                        oGeneralService2.Add(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDSE] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        'UpdateProcessTobeDone("7 : " + form_Name)
                    Else
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)
                        'UpdateProcessTobeDone("6 : " + form_Name)
                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService1.GetByParams(oGeneralParams2)

                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("DocNum", Objrs2.Fields.Item("DocNum").Value.ToString)
                        oGeneralData2.SetProperty("Series", Objrs2.Fields.Item("Series").Value.ToString)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")

                        oGeneralService2.Update(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDSE] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If

                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                    'UpdateProcessTobeDone("Ex1 : " + ex.ToString + " - " + form_Name)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
            'UpdateProcessTobeDone("Ex2 : " + ex.ToString + " - " + form_Name)
        End Try
    End Sub
    Private Sub allbranchconnection()

        strsql1 = "select Code'Locationcode',Location,U_DBIPADDR'IPaddr',U_DBNAME'DBName'  From OLCT "
        Objrs3 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs3.DoQuery(strsql1)
        'dt = New DataTable
        'da = New SqlDataAdapter(strsql1, con)
        'da.Fill(dt)
        Dim location As String = ""
        'Dim DGV As New DataGridView
        'DGV.DataSource = dt
        For i As Integer = 0 To Objrs3.RecordCount - 1

            If Objrs3.Fields.Item("IPaddr").Value.ToString <> "" Then

                objBRcompany3 = New SAPbobsCOM.Company
                objBRcompany3.Server = Objrs3.Fields.Item("IPaddr").Value.ToString
                objBRcompany3.LicenseServer = Objrs3.Fields.Item("IPaddr").Value.ToString + ":30000" ' dtheader.Rows(0)("LSRV")
                objBRcompany3.UseTrusted = False
                objBRcompany3.CompanyDB = Objrs3.Fields.Item("DBName").Value.ToString
                objBRcompany3.UserName = MSAPUID.Trim.ToString
                objBRcompany3.Password = MSAPPWd.Trim.ToString
                objBRcompany3.DbUserName = MUID.Trim.ToString
                objBRcompany3.DbPassword = MPWD.Trim.ToString
                objBRcompany3.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                lretcode = objBRcompany3.Connect()
                If lretcode <> 0 Then
                    MsgBox(objBRcompany3.GetLastErrorDescription)
                Else
                    location = Objrs3.Fields.Item("Location").Value.ToString
                    'saleswaste(location)
                    BranchTransfer(location)


                End If
            End If

            Objrs3.MoveNext()

        Next
    End Sub
    Private Sub BranchTransfer(ByVal location As String)
        Try
            strsql1 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_TAGDATE,103)'date',convert(varchar,U_PLUSDATE,103)'pdate',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS from [@mipldtag] where U_BRANCHSTOCKNOW='" & location.ToString & "' AND U_BRSTATUS='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Try

                    strsql2 = "select docentry,U_barcode,U_FORMBASISNO from [@mipldtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_barcode").Value.ToString & "' and U_FORMBASISNO='" & Objrs1.Fields.Item("U_FORMBASISNO").Value.ToString & "' "
                    Objrs2 = objBRcompany3.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()
                        If Not objBRcompany3.InTransaction Then objBRcompany3.StartTransaction()
                        oGeneralService2 = objBRcompany3.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Add(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y',U_BRSTATUS='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objBRcompany3.InTransaction Then objBRcompany3.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs1.MoveNext()
                    Else
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()
                        If Not objBRcompany3.InTransaction Then objBRcompany3.StartTransaction()
                        oGeneralService2 = objBRcompany3.GetCompanyService.GetGeneralService("MIPLDTAG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralParams1.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Update(oGeneralData2)


                  
                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y',U_BRSTATUS='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objBRcompany3.InTransaction Then objBRcompany3.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs1.MoveNext()
                    End If
                Catch ex As Exception
                    If objBRcompany3.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub OGEstimation12()
        Try
            'UpdateProcessTobeDone("5 : " + form_Name)
            strsql1 = "Select DocEntry,U_DOCNUM,U_NBRANCH,U_SERIES,convert(varchar(10),U_DOCDATE,103) U_DOCDATE from [@MIPLDOG] where isnull(U_TRANSFER2,'N')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            'UpdateProcessTobeDone("Count : " + Objrs1.RecordCount.ToString + " - " + form_Name)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                'UpdateProcessTobeDone("For : " + i.ToString + " - " + form_Name)
                Try
                    strsql2 = "Select DocEntry,DocNum,Series from [@MIPLDOG]  where U_NBRANCH='" & Objrs1.Fields.Item("U_NBRANCH").Value.ToString & "' and U_DOCNUM='" & Objrs1.Fields.Item("U_DOCNUM").Value.ToString & "' "
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    'UpdateProcessTobeDone("Count1 : " + Objrs2.RecordCount.ToString + " - " + form_Name)
                    'UpdateProcessTobeDone("DocEntry : " + Objrs1.Fields.Item("DocEntry").Value.ToString + " - " + form_Name)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)
                        'UpdateProcessTobeDone("6 : " + form_Name)
                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDOG' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("U_DOCDATE").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_NBRANCH").Value.ToString & "'"
                        Objrs5 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        series = Objrs5.Fields.Item("Series").Value
                        oGeneralData2.SetProperty("Series", series.ToString)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")

                        oGeneralService2.Add(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDOG] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        'UpdateProcessTobeDone("7 : " + form_Name)
                    Else
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)
                        'UpdateProcessTobeDone("6 : " + form_Name)
                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService1.GetByParams(oGeneralParams2)

                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralData2.SetProperty("DocNum", Objrs2.Fields.Item("DocNum").Value.ToString)
                        oGeneralData2.SetProperty("Series", Objrs2.Fields.Item("Series").Value.ToString)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")

                        oGeneralService2.Update(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDOG] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If


                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                    'UpdateProcessTobeDone("Ex1 : " + ex.ToString + " - " + form_Name)
                End Try
                Objrs1.MoveNext()
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
            'UpdateProcessTobeDone("Ex2 : " + ex.ToString + " - " + form_Name)
        End Try
    End Sub

    Private Sub OGEstimation21()
        Try

            strsql2 = "Select DocEntry,U_DOCNUM,U_NBRANCH,U_SERIES,convert(varchar(10),U_DOCDATE,120) U_DOCDATE from [@MIPLDOG] where isnull(U_TRANSFER1,'N')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                strsql1 = "Select DocEntry,DocNum,Series from [@MIPLDOG]  where U_NBRANCH='" & Objrs1.Fields.Item("U_NBRANCH").Value.ToString & "' and U_DOCNUM='" & Objrs1.Fields.Item("U_DOCNUM").Value.ToString & "' and U_SERIES='" & Objrs1.Fields.Item("U_SERIES").Value.ToString & "'"
                Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs1.DoQuery(strsql1)
                Try
                    If Objrs1.RecordCount = 0 Then
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData1.FromXMLString(xmlstring)

                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDOG' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("U_DOCDATE").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_NBRANCH").Value.ToString & "'"
                        Objrs5 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        series = Objrs5.Fields.Item("Series").Value
                        oGeneralData1.SetProperty("Series", series)

                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Add(oGeneralData1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLDOG] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Else
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDOG")
                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams2)

                        oGeneralData1.FromXMLString(xmlstring)
                        oGeneralData1.SetProperty("DocNum", Objrs1.Fields.Item("DocNum").Value.ToString)
                        oGeneralData1.SetProperty("Series", Objrs1.Fields.Item("Series").Value.ToString)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Update(oGeneralData1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLDOG] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    Objrs2.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next

        Catch ex As Exception
            'MsgBox(ex.ToString)
            'UpdateProcessTobeDone("Ex2 : " + ex.ToString + " - " + form_Name)
        End Try

    End Sub
    Private Sub SalesEstimation21()

        Try
            strsql2 = "select DocEntry,U_dailyno,U_branch,U_docstatus,convert(varchar(10),U_ESTDATE,120)'date'  from [@MIPLDSE] where isnull(U_TRANSFER1,'')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                strsql1 = "select DocEntry,DocNum,Series  from [@MIPLDSE] where U_branch='" & Objrs2.Fields.Item("U_branch").Value.ToString & "' and U_DAILYNO='" & Objrs2.Fields.Item("U_dailyno").Value.ToString & "' and U_ESTDATE='" & (Objrs2.Fields.Item("date").Value).ToString & "'"
                Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs1.DoQuery(strsql1)
                Try
                    If Objrs1.RecordCount = 0 Then
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData1.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDSE' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs2.Fields.Item("date").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs2.Fields.Item("U_branch").Value.ToString & "'"
                        Objrs5 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        SERIES = Objrs5.Fields.Item("Series").Value
                        oGeneralData1.SetProperty("Series", SERIES)

                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Add(oGeneralData1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLDSE] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Else
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)

                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()

                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams2)

                        oGeneralData1.FromXMLString(xmlstring)
                        oGeneralData1.SetProperty("DocNum", Objrs1.Fields.Item("DocNum").Value.ToString)
                        oGeneralData1.SetProperty("Series", Objrs1.Fields.Item("Series").Value.ToString)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Update(oGeneralData1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLDSE] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    Objrs2.MoveNext()
                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try

            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SalesEstimation12_check()
        Dim SERIES As Integer
        Try
            strsql1 = "select a.DocEntry,U_dailyno,b.u_tagno,U_branch,U_docstatus,convert(varchar,U_ESTDATE,103)'date',a.U_ESTNO   from [@MIPLDSE] a join [@MIPLDSE1] b on a.docentry=b.docentry where isnull(U_TRANSFER2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Try
                    strsql2 = "select docentry  from [@MIPLDSE] where U_branch='" & Objrs1.Fields.Item("U_branch").Value.ToString & "' and U_DAILYNO='" & Objrs1.Fields.Item("U_dailyno").Value.ToString & "' and U_ESTDATE='" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDSE' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_branch").Value.ToString & "'"
                        Objrs5 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        SERIES = Objrs5.Fields.Item("Series").Value
                        oGeneralData2.SetProperty("Series", SERIES)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Add(oGeneralData2)
                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDSE] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Else
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        'oGeneralParams2.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        strsql2 = "select DocEntry  From [@MIPLDSE] where U_ESTNO='" & Objrs1.Fields.Item("U_ESTNO").Value.ToString & "' "
                        Objrs3 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs3.DoQuery(strsql2)

                        oGeneralParams2.SetProperty("DocEntry", Objrs3.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)


                        oGeneralData2.FromXMLString(xmlstring)
                        oGeneralService2.Update(oGeneralData2)


                        'oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDSE")
                        'oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        'oGeneralData2.FromXMLString(xmlstring)
                        ''oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        'oGeneralService2.Update(oGeneralData2)


                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    If Objrs1.Fields.Item("U_TAGNO").Value.Substring(0, 2) = "BT" Then
                    Else
                        strsql2 = "select DocEntry,U_barcode,U_FORMBASISNO,convert(varchar,U_plusdate,103)'date',U_ORDERNO,U_ORDERLINENO,U_INVSTATUS,U_ESTSTATUS ,U_ESTOPENPCS,U_ESTOPENQTY,U_TAGSTATUS   from [@MIPLDtag] where U_FORMTYPE<>'GIFT' and U_barcode='" & Objrs1.Fields.Item("U_TAGNO").Value.ToString & "' "
                        Objrs3 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs3.DoQuery(strsql2)

                        strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLDtag] set U_TAGSTATUS='" & Objrs3.Fields.Item("U_TAGSTATUS").Value.ToString & "',U_TAGDATE='" & CDate(Objrs3.Fields.Item("date").Value).ToString("yyyyMMdd") & "',U_ORDERNO='" & Objrs3.Fields.Item("U_ORDERNO").Value.ToString & "',U_ORDERLINENO='" & Objrs3.Fields.Item("U_ORDERLINENO").Value.ToString & "',U_INVSTATUS='" & Objrs3.Fields.Item("U_INVSTATUS").Value.ToString & "',U_ESTSTATUS ='" & Objrs3.Fields.Item("U_ESTSTATUS").Value.ToString & "',U_ESTOPENPCS='" & Objrs3.Fields.Item("U_ESTOPENPCS").Value.ToString & "',U_ESTOPENQTY='" & Objrs3.Fields.Item("U_ESTOPENQTY").Value.ToString & "' where  U_FORMBASISNO='" & Objrs3.Fields.Item("U_FORMBASISNO").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@mipldtag] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs4 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs4.DoQuery(strsql2)
                    End If



                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub batch12()
        Try
            strsql1 = "select U_BARCODE,U_COuntercode,U_BRANCHNAME,DocEntry,U_Grossweight,U_netweight,U_pieces,U_POSTDATE'date'   from [@MIPLBAT] where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Try
                    strsql2 = "select DocEntry  from [@miplbat] where U_BARCODE='" & Objrs1.Fields.Item("U_BARCODE").Value.ToString & "' and U_BRANCHNAME='" & Objrs1.Fields.Item("U_BRANCHNAME").Value.ToString & "' and U_COuntercode='" & Objrs1.Fields.Item("U_COuntercode").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLBAT")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                        oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams1.SetProperty("DocEntry", Objrs1.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                        xmlstring = oGeneralData1.ToXMLString()

                        If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLBAT")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData2.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDSE' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_BRANCHNAME").Value.ToString & "'"
                        Objrs5 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        series = Objrs5.Fields.Item("Series").Value
                        oGeneralData2.SetProperty("Series", series)
                        oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                        oGeneralService2.Add(oGeneralData2)

                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        'Objrs1.MoveNext()
                    Else
                        strsql2 = "update " & MDBName2.Trim.ToString & "..[@MIPLBAT] set U_NETWEIGHT='" & Objrs1.Fields.Item("U_netweight").Value.ToString & "',U_GROSSWEIGHT='" & Objrs1.Fields.Item("U_Grossweight").Value.ToString & "',U_PIECES='" & Objrs1.Fields.Item("U_pieces").Value.ToString & "' where  DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        'strsql2 += vbCrLf + "update " & MDBName1.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)


                        strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs2.DoQuery(strsql2)
                        'If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                    End If
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLBAT] set U_TRNREMARKS='" & ex.ToString & "' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    MsgBox(ex.ToString)
                End Try
                Objrs1.MoveNext()
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub batch21()
        Try
            strsql2 = "select U_BARCODE,U_COuntercode,U_BRANCHNAME,DocEntry,U_Grossweight,U_netweight,U_pieces,U_POSTDATE'date'  from [@MIPLBAT] where isnull(U_TRANSFER1,'')='N'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            For i As Integer = 0 To Objrs2.RecordCount - 1
                Try
                    strsql1 = "select DocEntry  from [@miplbat] where U_BARCODE='" & Objrs2.Fields.Item("U_BARCODE").Value.ToString & "' and U_BRANCHNAME='" & Objrs2.Fields.Item("U_BRANCHNAME").Value.ToString & "' and U_COuntercode='" & Objrs2.Fields.Item("U_COuntercode").Value.ToString & "'"
                    Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs1.DoQuery(strsql1)
                    If Objrs1.RecordCount = 0 Then
                        oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLBAT")
                        oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralParams2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        oGeneralParams2.SetProperty("DocEntry", Objrs2.Fields.Item("DocEntry").Value.ToString)
                        oGeneralData2 = oGeneralService2.GetByParams(oGeneralParams2)
                        xmlstring = oGeneralData2.ToXMLString()
                        If Not objcompany1.InTransaction Then objcompany1.StartTransaction()
                        oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLBAT")
                        oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                        oGeneralData1.FromXMLString(xmlstring)
                        strsql1 = "select Series  from NNM1 where ObjectCode ='MIPLDSE' and  indicator=(select Indicator  from OFPR where '" & CDate(Objrs1.Fields.Item("date").Value).ToString("yyyyMMdd") & "' between F_RefDate  and T_RefDate  ) AND remark='" & Objrs1.Fields.Item("U_BRANCHNAME").Value.ToString & "'"
                        Objrs5 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs5.DoQuery(strsql1)
                        series = Objrs5.Fields.Item("Series").Value
                        oGeneralData1.SetProperty("Series", series)
                        oGeneralData1.SetProperty("U_TRANSFER1", "Y")
                        oGeneralService1.Add(oGeneralData1)
                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)
                        If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Objrs2.MoveNext()
                    Else
                        strsql1 = "update " & MDBName1.Trim.ToString & "..[@MIPLBAT] set U_NETWEIGHT='" & Objrs2.Fields.Item("U_netweight").Value.ToString & "',U_GROSSWEIGHT='" & Objrs2.Fields.Item("U_Grossweight").Value.ToString & "',U_PIECES='" & Objrs2.Fields.Item("U_pieces").Value.ToString & "' where  DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                        'strsql1 += vbCrLf + "update " & MDBName2.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        strsql1 = "update " & MDBName2.Trim.ToString & "..[@MIPLBAT] set U_TRANSFER1='Y' where DocEntry='" & Objrs2.Fields.Item("DocEntry").Value.ToString & "'"
                        Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Objrs1.DoQuery(strsql1)

                        Objrs2.MoveNext()
                    End If

                Catch ex As Exception
                    If objcompany1.InTransaction Then objcompany1.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub metal_master12()
        Try
            strsql1 = "select Code  from [@MIPLMT] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLMT")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLMT")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@miplmt] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLMT] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Purity_master12()
        Try
            strsql1 = "select Code  from [@MIPLPM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLPM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLPM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLPM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLPM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Incentive_master12()
        Try
            strsql1 = "select Code  from [@MIPLICM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLICM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLICM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLICM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLICM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub HallMark_master12()
        Try
            strsql1 = "select Code  from [@MIPLHMC] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLHMC")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLHMC")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLHMC] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLHMC] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Discount_master12()
        Try
            strsql1 = "select Code  from [@MIPLDM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLDM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLDM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLDM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLDM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub reason_master12()
        Try
            strsql1 = "select Code  from [@MIPLREASON] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLREASON")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLREASON")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLREASON] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLREASON] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub NOde_master12()
        Try
            strsql1 = "select Code  from [@MIPLNM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLNM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLNM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLNM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLNM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Nodecounter_master12()
        Try
            strsql1 = "select Code  from [@MIPLNCM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLNCM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLNCM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLNCM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLNCM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CustomerPrivilege_master12()
        Try
            strsql1 = "select Code  from [@MIPLPRm] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLPRM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLPRM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLPRm] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLPRm] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub VendorPrivilege_master12()
        Try
            strsql1 = "select Code  from [@MIPLVPM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLVPM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLVPM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLVPM] where code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLVPM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Product_master12()
        Try
            strsql1 = "select Code,U_PRODCODE  from [@MIPLIM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLIM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLIM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLIM] where U_PRODCODE='" & Objrs1.Fields.Item("U_PRODCODE").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLIM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub GiftItem_master12()
        Try
            strsql1 = "select Code,U_GIFTCODE  from [@MIPLGM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLGM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLGM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLGM] where U_GIFTCODE='" & Objrs1.Fields.Item("U_GIFTCODE").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLGM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Size_master12()
        Try
            strsql1 = "select Code,U_SIZECODE  from [@MIPLSM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLSM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLSM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLSM] where U_SIZECODE='" & Objrs1.Fields.Item("U_SIZECODE").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLSM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Subproduct_master12()
        ''Add Sales Invoice
        Try

            strsql1 = "select ItemCode from oitm where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim Oitem1 As SAPbobsCOM.Items
                Oitem1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                Oitem1.GetByKey(Objrs1.Fields.Item("ItemCode").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim oitem2 As SAPbobsCOM.Items
                oitem2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                strsql2 = "select 1 from oitm where itemcode='" & Objrs1.Fields.Item("ItemCode").Value.ToString & "'"
                Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs2.DoQuery(strsql2)
                If Objrs2.RecordCount > 0 Then
                    oitem2.GetByKey(Objrs1.Fields.Item("ItemCode").Value.ToString)
                Else
                    oitem2.ItemCode = Oitem1.ItemCode
                End If
                'Header
                oitem2.ItemName = Oitem1.ItemName
                oitem2.UserFields.Fields.Item("U_PRODCODE").Value = Oitem1.UserFields.Fields.Item("U_PRODCODE").Value
                oitem2.UserFields.Fields.Item("U_PRODNAME").Value = Oitem1.UserFields.Fields.Item("U_PRODNAME").Value
                oitem2.InventoryUOM = Oitem1.InventoryUOM
                oitem2.PurchaseUnit = Oitem1.PurchaseUnit
                oitem2.SalesUnit = Oitem1.SalesUnit
                oitem2.PurchaseItem = Oitem1.PurchaseItem
                oitem2.SalesItem = Oitem1.SalesItem
                oitem2.InventoryItem = Oitem1.InventoryItem
                oitem2.AssetItem = Oitem1.AssetItem
                oitem2.ForeignName = Oitem1.ForeignName
                'oitem2.ItemsGroupCode = Oitem1.ItemsGroupCode
                oitem2.UserFields.Fields.Item("U_METALTYPE").Value = Oitem1.UserFields.Fields.Item("U_METALTYPE").Value
                oitem2.UserFields.Fields.Item("U_CATEGORY").Value = Oitem1.UserFields.Fields.Item("U_CATEGORY").Value
                oitem2.UserFields.Fields.Item("U_STOCKTYPE").Value = Oitem1.UserFields.Fields.Item("U_STOCKTYPE").Value
                oitem2.UserFields.Fields.Item("U_SALESMODE").Value = Oitem1.UserFields.Fields.Item("U_SALESMODE").Value
                oitem2.UserFields.Fields.Item("U_PURCHASEMODE").Value = Oitem1.UserFields.Fields.Item("U_PURCHASEMODE").Value
                oitem2.UserFields.Fields.Item("U_MCTYPE").Value = Oitem1.UserFields.Fields.Item("U_MCTYPE").Value
                oitem2.UserFields.Fields.Item("U_OGRATE").Value = Oitem1.UserFields.Fields.Item("U_OGRATE").Value
                oitem2.UserFields.Fields.Item("U_BRAND").Value = Oitem1.UserFields.Fields.Item("U_BRAND").Value
                oitem2.UserFields.Fields.Item("U_DEFPCS").Value = Oitem1.UserFields.Fields.Item("U_DEFPCS").Value
                oitem2.UserFields.Fields.Item("U_TAXCODE").Value = Oitem1.UserFields.Fields.Item("U_TAXCODE").Value
                oitem2.UserFields.Fields.Item("U_PURITYID").Value = Oitem1.UserFields.Fields.Item("U_PURITYID").Value
                oitem2.UserFields.Fields.Item("U_DESCRIPTION").Value = Oitem1.UserFields.Fields.Item("U_DESCRIPTION").Value
                oitem2.UserFields.Fields.Item("U_PURITY").Value = Oitem1.UserFields.Fields.Item("U_PURITY").Value
                oitem2.UserFields.Fields.Item("U_HALLMARKID").Value = Oitem1.UserFields.Fields.Item("U_HALLMARKID").Value
                oitem2.UserFields.Fields.Item("U_SIZE").Value = Oitem1.UserFields.Fields.Item("U_SIZE").Value
                oitem2.UserFields.Fields.Item("U_STONES").Value = Oitem1.UserFields.Fields.Item("U_STONES").Value
                oitem2.UserFields.Fields.Item("U_HALLMKCHRGE").Value = Oitem1.UserFields.Fields.Item("U_HALLMKCHRGE").Value
                oitem2.UserFields.Fields.Item("U_WASTAGE").Value = Oitem1.UserFields.Fields.Item("U_WASTAGE").Value
                oitem2.UserFields.Fields.Item("U_OTHERCHRGE").Value = Oitem1.UserFields.Fields.Item("U_OTHERCHRGE").Value
                oitem2.UserFields.Fields.Item("U_BRATEDIFF").Value = Oitem1.UserFields.Fields.Item("U_BRATEDIFF").Value
                oitem2.UserFields.Fields.Item("U_BWASTAGE").Value = Oitem1.UserFields.Fields.Item("U_BWASTAGE").Value
                oitem2.UserFields.Fields.Item("U_LESSTAXONRATE").Value = Oitem1.UserFields.Fields.Item("U_LESSTAXONRATE").Value
                oitem2.UserFields.Fields.Item("U_MULTIMETAL").Value = Oitem1.UserFields.Fields.Item("U_MULTIMETAL").Value
                oitem2.UserFields.Fields.Item("U_REASON1").Value = Oitem1.UserFields.Fields.Item("U_REASON1").Value
                oitem2.UserFields.Fields.Item("U_REASON2").Value = Oitem1.UserFields.Fields.Item("U_REASON2").Value
                oitem2.UserFields.Fields.Item("U_REASON3").Value = Oitem1.UserFields.Fields.Item("U_REASON3").Value
                oitem2.UserFields.Fields.Item("U_MANUALWASTAGE").Value = Oitem1.UserFields.Fields.Item("U_MANUALWASTAGE").Value
                oitem2.UserFields.Fields.Item("U_BOUGHTNOTE").Value = Oitem1.UserFields.Fields.Item("U_BOUGHTNOTE").Value
                oitem2.UserFields.Fields.Item("U_BOM").Value = Oitem1.UserFields.Fields.Item("U_BOM").Value
                oitem2.UserFields.Fields.Item("U_SALESRETURN").Value = Oitem1.UserFields.Fields.Item("U_SALESRETURN").Value
                oitem2.UserFields.Fields.Item("U_SALESRETDAYS").Value = Oitem1.UserFields.Fields.Item("U_SALESRETDAYS").Value
                oitem2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                oitem2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                oitem2.UserFields.Fields.Item("U_TRANSFER3").Value = Oitem1.UserFields.Fields.Item("U_TRANSFER3").Value
                oitem2.User_Text = Oitem1.User_Text


                oitem2.Valid = Oitem1.Valid
                oitem2.ValidFrom = Oitem1.ValidFrom
                oitem2.ValidTo = Oitem1.ValidTo
                oitem2.ValidRemarks = Oitem1.ValidRemarks
                oitem2.Frozen = Oitem1.Frozen
                oitem2.FrozenFrom = Oitem1.FrozenFrom
                oitem2.FrozenTo = Oitem1.FrozenTo
                oitem2.FrozenRemarks = Oitem1.FrozenRemarks
                oitem2.UserFields.Fields.Item("U_PRICELIST").Value = Oitem1.UserFields.Fields.Item("U_PRICELIST").Value
                oitem2.UserFields.Fields.Item("U_PRICELISTNAME").Value = Oitem1.UserFields.Fields.Item("U_PRICELISTNAME").Value
                oitem2.UserFields.Fields.Item("U_PRICEVALUE").Value = Oitem1.UserFields.Fields.Item("U_PRICEVALUE").Value
                'For j As Integer = 0 To Oitem1.PriceList.Count - 1
                '    oitem2.PriceList.SetCurrentLine(j)
                '    oitem2.PriceList.Price = Oitem1.PriceList.Price
                'Next
                'If Objrs2.RecordCount = 0 Then
                '    For j As Integer = 0 To Oitem1.WhsInfo.Count - 1
                '        Oitem1.WhsInfo.SetCurrentLine(j)
                '        oitem2.WhsInfo.WarehouseCode = Oitem1.WhsInfo.WarehouseCode
                '        oitem2.WhsInfo.Locked = Oitem1.WhsInfo.Locked
                '        oitem2.WhsInfo.Add()
                '    Next
                'Else
                '    For j As Integer = 0 To Oitem1.WhsInfo.Count - 1
                '        Oitem1.WhsInfo.SetCurrentLine(j)
                '        oitem2.WhsInfo.SetCurrentLine(j)
                '        oitem2.WhsInfo.WarehouseCode = Oitem1.WhsInfo.WarehouseCode
                '        oitem2.WhsInfo.Locked = Oitem1.WhsInfo.Locked
                '    Next
                'End If
                If Objrs2.RecordCount > 0 Then
                    lretcode = oitem2.Update
                Else
                    lretcode = oitem2.Add
                End If
                If lretcode <> 0 Then
                    MsgBox(objcompany2.GetLastErrorDescription)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..OITM set U_TRANSFER2='Y' where Itemcode='" & Objrs1.Fields.Item("ItemCode").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Rate_master12()
        Try
            strsql1 = "select Code  from [@MIPLRM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLRM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLRM")
                    'oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    oGeneralService2.Add(oGeneralData2)

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLRM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub RateBranchwisesettings_master12()
        Try
            strsql1 = "select Code  from [@MIPLBRM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLBRM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLBRM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLBRM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLBRM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Purchasecent_master12()
        Try
            strsql1 = "select Code  from [@MIPLPCM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLPCM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLPCM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLPCM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLPCM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Salescent_master12()
        Try
            strsql1 = "select Code  from [@MIPLSCM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLSCM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLSCM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLSCM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLSCM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Purchasewastage_master12()
        Try
            strsql1 = "select Code  from [@MIPLPWM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLPWM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLPWM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLPWM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLPWM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Saleswastage_master12()
        Try
            strsql1 = "select Code  from [@MIPLSWM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLSWM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLSWM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLSWM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLSWM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SpecialSaleswastage_master12()
        Try
            strsql1 = "select Code  from [@MIPLSSWM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLSSWM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLSSWM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLSSWM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLSSWM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Repairwastage_master12()
        Try
            strsql1 = "select Code  from [@MIPLRWM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLRWM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLRWM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLRWM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLRWM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SalesEstimationparameter_master12()
        Try
            strsql1 = "select Code  from [@MIPLPSEST] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLPSEST")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLPSEST")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLPSEST] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLPSEST] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BP_master12()
        ''Add Bp Master from 1 to 2
        Try

            strsql1 = "select cardcode from ocrd where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim bp1 As SAPbobsCOM.BusinessPartners
                bp1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                bp1.GetByKey(Objrs1.Fields.Item("cardcode").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim bp2 As SAPbobsCOM.BusinessPartners
                bp2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                strsql2 = "select 1 from ocrd where cardcode='" & Objrs1.Fields.Item("cardcode").Value.ToString & "'"
                Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs2.DoQuery(strsql2)
                If Objrs2.RecordCount > 0 Then
                    bp2.GetByKey(Objrs1.Fields.Item("cardcode").Value.ToString)
                Else
                    bp2.CardCode = bp1.CardCode
                End If
                'Header

                bp2.SubjectToWithholdingTax = bp1.SubjectToWithholdingTax
                bp2.CardName = bp1.CardName
                bp2.CardType = bp1.CardType
                bp2.GroupCode = bp1.GroupCode
                If bp1.CardType = SAPbobsCOM.BoCardTypes.cSupplier Then
                    bp2.DebitorAccount = bp1.DebitorAccount
                    bp2.AccountRecivablePayables.Add()
                    bp2.DownPtpanaymentClearAct = bp1.DownPtpanaymentClearAct
                ElseIf bp1.CardType = SAPbobsCOM.BoCardTypes.cCustomer Then
                    bp2.DebitorAccount = bp2.DebitorAccount
                    bp2.DownPaymentClearAct = bp2.DownPaymentClearAct
                End If
                'BP.CardForeignName = txtDesignerCode.Text
                bp2.Phone1 = bp1.Phone1
                bp2.Cellular = bp1.Cellular
                bp2.EmailAddress = bp1.EmailAddress
                bp2.Addresses.AddressName = bp1.Addresses.AddressName
                bp2.Addresses.AddressName2 = bp1.Addresses.AddressName2
                bp2.Addresses.Street = bp1.Addresses.Street ''AREA MAPPING IN STREET FIELD IN BP MASTER
                bp2.Addresses.City = bp1.Addresses.City
                bp2.Addresses.ZipCode = bp1.Addresses.ZipCode

                bp2.FiscalTaxID.TaxId0 = bp1.FiscalTaxID.TaxId0
                bp2.FiscalTaxID.Add()
                bp2.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo

                bp2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                bp2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                bp2.UserFields.Fields.Item("U_TRANSFER3").Value = "1"

                If Objrs2.RecordCount > 0 Then
                    lretcode = bp2.Update
                Else
                    lretcode = bp2.Add
                End If
                If lretcode <> 0 Then
                    'MsgBox(objcompany2.GetLastErrorDescription)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..OCRD set U_TRANSFER2='Y' where cardcode='" & Objrs1.Fields.Item("cardcode").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Customermaster_MIPLCM_12()
        Try
            strsql1 = "select Code  from [@MIPLCM] where isnull(U_TRANSFER2,'')='N' order by convert(int,Code)"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                oGeneralService1 = objcompany1.GetCompanyService.GetGeneralService("MIPLCM")
                oGeneralData1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralParams1 = oGeneralService1.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams1.SetProperty("Code", Objrs1.Fields.Item("Code").Value.ToString)
                oGeneralData1 = oGeneralService1.GetByParams(oGeneralParams1)

                xmlstring = oGeneralData1.ToXMLString()
                Try
                    If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                    oGeneralService2 = objcompany2.GetCompanyService.GetGeneralService("MIPLCM")
                    oGeneralData2 = oGeneralService2.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    oGeneralData2.FromXMLString(xmlstring)
                    oGeneralData2.SetProperty("U_TRANSFER2", "Y")
                    strsql2 = "select 1 from [@MIPLCM] where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If Objrs2.RecordCount = 0 Then
                        oGeneralService2.Add(oGeneralData2)
                    Else
                        oGeneralService2.Update(oGeneralData2)
                    End If

                    strsql2 = "update " & MDBName1.Trim.ToString & "..[@MIPLCM] set U_TRANSFER2='Y' where Code='" & Objrs1.Fields.Item("Code").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                Catch ex As Exception
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    'MsgBox(ex.ToString)
                End Try
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Salesorder12()
        ''Add Sales Invoice
        Try

            strsql1 = "select DocEntry from ordr where isnull(U_transfer2,'')='N'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                strsql2 = "select 1 from ordr where docentry='" & Objrs1.Fields.Item("docentry").Value & "'"
                Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs2.DoQuery(strsql2)
                If Objrs2.RecordCount > 0 Then Continue For
                Dim oorder1 As SAPbobsCOM.Documents
                oorder1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                oorder1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)

                If Not objcompany2.InTransaction Then objcompany2.StartTransaction()
                Dim oorder2 As SAPbobsCOM.Documents
                oorder2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                oorder2.CardCode = oorder1.CardCode
                oorder2.Address = oorder1.Address
                oorder2.Address2 = oorder1.Address2

                oorder2.UserFields.Fields.Item("U_CONTACTNO").Value = oorder1.UserFields.Fields.Item("U_CONTACTNO").Value
                oorder2.UserFields.Fields.Item("U_TELLNO").Value = oorder1.UserFields.Fields.Item("U_TELLNO").Value
                oorder2.UserFields.Fields.Item("U_EMAIL").Value = oorder1.UserFields.Fields.Item("U_EMAIL").Value
                oorder2.UserFields.Fields.Item("U_ESTTYPE").Value = oorder1.UserFields.Fields.Item("U_ESTTYPE").Value
                oorder2.UserFields.Fields.Item("U_STATUS").Value = oorder1.UserFields.Fields.Item("U_STATUS").Value

                oorder2.UserFields.Fields.Item("U_TRANSFER1").Value = "Y"
                oorder2.UserFields.Fields.Item("U_TRANSFER2").Value = "Y"
                oorder2.UserFields.Fields.Item("U_TRANSFER3").Value = "1"

                oorder2.NumAtCard = oorder1.NumAtCard

                oorder2.DocDate = oorder1.DocDate
                oorder2.DocDueDate = oorder1.DocDueDate
                oorder2.TaxDate = oorder1.TaxDate
                oorder2.Series = oorder1.Series
                oorder2.DocNum = oorder1.DocNum
                oorder2.UserFields.Fields.Item("U_WORKORDERTYPE").Value = oorder1.UserFields.Fields.Item("U_WORKORDERTYPE").Value
                oorder2.DocumentsOwner = oorder1.DocumentsOwner

                oorder2.UserFields.Fields.Item("U_ADVANCE").Value = oorder1.UserFields.Fields.Item("U_ADVANCE").Value
                oorder2.UserFields.Fields.Item("U_BOOKED").Value = oorder1.UserFields.Fields.Item("U_BOOKED").Value
                oorder2.UserFields.Fields.Item("U_PAYMENTSTATUS").Value = oorder1.UserFields.Fields.Item("U_PAYMENTSTATUS").Value
                oorder2.Comments = oorder1.Comments

                For j As Integer = 0 To oorder1.Lines.Count - 1
                    oorder2.Lines.UserFields.Fields.Item("U_TAGNO").Value = oorder1.Lines.UserFields.Fields.Item("U_TAGNO").Value
                    oorder2.Lines.UserFields.Fields.Item("U_DOCNO").Value = oorder1.Lines.UserFields.Fields.Item("U_DOCNO").Value
                    oorder2.Lines.UserFields.Fields.Item("U_ORDERNO").Value = oorder1.Lines.UserFields.Fields.Item("U_ORDERNO").Value
                    oorder2.Lines.UserFields.Fields.Item("U_PRODCODE").Value = oorder1.Lines.UserFields.Fields.Item("U_PRODCODE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_PRODNAME").Value = oorder1.Lines.UserFields.Fields.Item("U_PRODNAME").Value
                    oorder2.Lines.ItemCode = oorder1.Lines.ItemCode
                    oorder2.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value = oorder1.Lines.UserFields.Fields.Item("U_NOOFPIECES").Value
                    oorder2.Lines.UserFields.Fields.Item("U_SIZENAME").Value = oorder1.Lines.UserFields.Fields.Item("U_SIZENAME").Value
                    oorder2.Lines.WarehouseCode = oorder1.Lines.WarehouseCode
                    oorder2.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value = oorder1.Lines.UserFields.Fields.Item("U_COUNTERNAME").Value
                    oorder2.Lines.LocationCode = oorder1.Lines.LocationCode
                    oorder2.Lines.Quantity = oorder1.Lines.Quantity
                    oorder2.Lines.TaxCode = oorder1.Lines.TaxCode
                    oorder2.Lines.LineTotal = oorder1.Lines.LineTotal
                    oorder2.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value = oorder1.Lines.UserFields.Fields.Item("U_GROSSWEIGHT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value = oorder1.Lines.UserFields.Fields.Item("U_LESSWEIGHT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value = oorder1.Lines.UserFields.Fields.Item("U_NETWEIGHT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value = oorder1.Lines.UserFields.Fields.Item("U_WASTAGEPERCEN").Value
                    oorder2.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value = oorder1.Lines.UserFields.Fields.Item("U_WASTAGEGRAM").Value
                    oorder2.Lines.UserFields.Fields.Item("U_WASTAGEAMT").Value = oorder1.Lines.UserFields.Fields.Item("U_WASTAGEAMT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_MAKINGCHRGAMT").Value = oorder1.Lines.UserFields.Fields.Item("U_MAKINGCHRGAMT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_STONECHARGES").Value = oorder1.Lines.UserFields.Fields.Item("U_STONECHARGES").Value
                    oorder2.Lines.UserFields.Fields.Item("U_HALLMARKCHRG").Value = oorder1.Lines.UserFields.Fields.Item("U_HALLMARKCHRG").Value
                    oorder2.Lines.UserFields.Fields.Item("U_PURITYPERCEN").Value = oorder1.Lines.UserFields.Fields.Item("U_PURITYPERCEN").Value
                    oorder2.Lines.UserFields.Fields.Item("U_PUREWEIGHT").Value = oorder1.Lines.UserFields.Fields.Item("U_PUREWEIGHT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_UNITPRICE").Value = oorder1.Lines.UserFields.Fields.Item("U_UNITPRICE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_DISCOUNTAMT").Value = oorder1.Lines.UserFields.Fields.Item("U_DISCOUNTAMT").Value
                    oorder2.Lines.UserFields.Fields.Item("U_BEFOREDISC").Value = oorder1.Lines.UserFields.Fields.Item("U_BEFOREDISC").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGR1").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGR1").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT1").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT1").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGR2").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGR2").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT2").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT2").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGR3").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGR3").Value
                    oorder2.Lines.UserFields.Fields.Item("U_OTHCHRGAMT3").Value = oorder1.Lines.UserFields.Fields.Item("U_OTHCHRGAMT3").Value
                    oorder2.Lines.UserFields.Fields.Item("U_DELIVERYDATE").Value = oorder1.Lines.UserFields.Fields.Item("U_DELIVERYDATE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_REMARKS").Value = oorder1.Lines.UserFields.Fields.Item("U_REMARKS").Value
                    oorder2.Lines.UserFields.Fields.Item("U_WASTAGEPCS").Value = oorder1.Lines.UserFields.Fields.Item("U_WASTAGEPCS").Value
                    oorder2.Lines.UserFields.Fields.Item("U_WASTAGETYPE").Value = oorder1.Lines.UserFields.Fields.Item("U_WASTAGETYPE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_MAKINGCHRGTYPE").Value = oorder1.Lines.UserFields.Fields.Item("U_MAKINGCHRGTYPE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_MC").Value = oorder1.Lines.UserFields.Fields.Item("U_MC").Value
                    oorder2.Lines.UserFields.Fields.Item("U_ORIGINALWASTAGE").Value = oorder1.Lines.UserFields.Fields.Item("U_ORIGINALWASTAGE").Value
                    oorder2.Lines.UserFields.Fields.Item("U_ORIGINALMC").Value = oorder1.Lines.UserFields.Fields.Item("U_ORIGINALMC").Value
                    oorder2.Lines.UserFields.Fields.Item("U_EMPNAME").Value = oorder1.Lines.UserFields.Fields.Item("U_EMPNAME").Value
                    oorder2.Lines.UserFields.Fields.Item("U_LINENO").Value = oorder1.Lines.UserFields.Fields.Item("U_LINENO").Value
                    oorder2.Lines.UserFields.Fields.Item("U_LINESTATUS").Value = oorder1.Lines.UserFields.Fields.Item("U_LINESTATUS").Value

                    oorder2.Lines.Add()
                Next

                oorder2.DiscountPercent = oorder1.DiscountPercent
                If oorder1.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    oorder2.RoundingDiffAmount = oorder1.RoundingDiffAmount
                Else
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                lretcode = oorder2.Add
                If lretcode <> 0 Then
                    'MsgBox(objcompany2.GetLastErrorDescription)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    strsql2 = "update " & MDBName1.Trim.ToString & "..ORDR set U_TRANSFER2='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)
                    If objcompany2.InTransaction Then objcompany2.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    Objrs1.MoveNext()
                End If
            Next
            updatesalesorder12()
            updatesalesorder21()
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub


    Public Sub write_log(ByVal status As String)
        Try
            If dt_time = "" Then dt_time = Today.ToString("yyyyMMdd") & "\Log_" & Now.ToString("HH_mm_ss")
            Dim di As DirectoryInfo = New DirectoryInfo(Application.StartupPath.ToString + "\logfiles\SAP_DB_ToSAP DB_" & Today.ToString("yyyyMMdd") & "")
            If di.Exists Then
            Else
                di.Create()
            End If
            chatlog = Application.StartupPath.ToString + "\logfiles\SAP_DB_ToSAP DB_" & dt_time & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
                'objWriter = New System.IO.StreamWriter(chatlog, True)
                'objWriter.WriteLine("Current date & time       Transaction-Docentry       LogDetails")
                'objWriter.WriteLine("-------------------       --------------------       ----------")
                'objWriter.Close()
            End If
            objWriter = New System.IO.StreamWriter(chatlog, True)
            If status <> "" Then objWriter.WriteLine(Now & " : " & status)
            objWriter.WriteLine(" ")
            objWriter.Close()
        Catch ex As Exception
            MsgBox("createlog :" + ex.ToString)
        End Try
    End Sub
    Private Sub Fun_GI_To_GR(ByVal DBC As String, ByVal MBaseEntry As String, ByVal MainCode As String)
        ''Add Sales Invoice
        Dim updatestatus As Boolean = False
        Try
            ' objconnection1.CompanyToConnection(DBC)

            strsql1 = "select a.DocEntry,b.ProfitCode'COST',a.U_TrnsctnCtgry from OIGE a join jdt1 b on a.transid=b.transid and b.line_ID=0 where a.DocEntry='" & MBaseEntry & "'"
            Objrs1 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim vendorcode As String = ""
                Dim cvendorcode As String = ""
                Dim whscode As String = ""
                Dim costcenter As String = ""
                Dim ToDB As String = ""
                Dim series As String = ""
                Dim series1 As String = ""
                Dim RMWHSCode As String = ""
                Dim mainquers As String = ""

                Dim LinsSql As String = ""
                LinsSql = "select Itemcode,Quantity,LineTotal,isnull(TaxCode,'0') TaxCode,U_PairQty,U_trnsctnCtgry From [" & DBC & "]..IGE1 where docentry='" & Objrs1.Fields.Item("DocEntry").Value & "' order by Linenum asc"

                mainquers = " select distinct a.DocDate,b.U_DB'ToDB',b.U_CostCenter'Cost Center',d.U_Acctcode'Acc',a.ItemCode,a.Quantity,b.U_DfltWhs'Default Whs',b.U_RmWhs 'RMWHS',isnull(convert(varchar,b.U_BPLid),'N') 'BPLid' from /*From DB*/[" & DBC & "]..ige1 a"
                mainquers += " inner join [" & DBC & "]..oige c on a.docentry=c.docentry"
                mainquers += "  inner join [VKCGROUPSOPKERALA(D)]..[@MIPL_INTERCOMPANY] b on a.acctcode=b.u_acctcode"
                mainquers += " left join [VKCGROUPSOPKERALA(D)]..[@MIPL_INTERCOMPANY] d on c.U_COSTCENTER=d.U_CostCenter"
                mainquers += " where a.docentry='" & MBaseEntry & "'"


                ' strsql1 = "select U_DfltWhs 'ToWhscode',U_CostCenter 'ToCostcentre',U_DB 'ToDB' from [@MIPL_INTERCOMPANY] where U_CardCode='" & Objrs1.Fields.Item("CardCode").Value & "'"
                Dim dts As New DataTable
                dts = objconnection1.GetSingleValue_SQL_dt(mainquers)
                Dim dtLines As New DataTable
                'dts = objconnection1.GetSingleValue_SQL_dt(strsql1)
                dtLines = objconnection1.GetSingleValue_SQL_dt(LinsSql)


                cvendorcode = Objrs1.Fields.Item("COST").Value
                Dim TransType As String = ""
                TransType = Objrs1.Fields.Item("U_TrnsctnCtgry").Value
                costcenter = dts.Rows(0).Item("Cost Center")
                write_log("Cost Center:" & costcenter)
                whscode = dts.Rows(0).Item("Default Whs")
                RMWHSCode = dts.Rows(0).Item("RMWHS")
                write_log("Whs Code:" & whscode)

                write_log("Vendor Code:" & vendorcode)
                write_log("TO DB:" & dts.Rows(0).Item("ToDB"))
                Dim oorder1 As SAPbobsCOM.Documents
                If dts.Rows.Count > 0 Then
                    oorder1 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    oorder1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)
                    ' vendorcode = dts.Rows(0).Item("U_CardCode")
                    ToDB = dts.Rows(0).Item("ToDB")
                    ' costcenter = dts.Rows(0).Item("U_CostCenter")
                    'whscode = dts.Rows(0).Item("U_DfltWhs")
                    objconnection1.CompanyToConnection(ToDB)
                    oss.createTables_to()

                    strsql2 = "SELECT distinct  T0.[U_Series] FROM [" & ToDB & "]..[@COSTCENTRE]  T0 where 	T0.U_BranchCode='" & costcenter & "' and T0.U_Doctype='59'"
                    series = objconnection1.GetSingleValue_SQL_SQL(strsql2)
                    strsql2 = "SELECT Series FROM [" & ToDB & "]..NNM1 where 	SeriesName='" & series & "'"

                    series1 = objconnection1.GetSingleValue_SQL_SQL(strsql2)
                End If
                'strsql2 = "select 1 from ODRF where U_BaseEntry='" & Objrs1.Fields.Item("docentry").Value & "'"
                'Objrs2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Objrs2.DoQuery(strsql2)
                'If Objrs2.RecordCount > 0 Then Continue For

                If Not objToCompany.InTransaction Then objToCompany.StartTransaction()
                Dim oorder2 As SAPbobsCOM.Documents
                'oorder2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                oorder2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                oorder2.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry
                'oorder2.CardCode = vendorcode
                updatestatus = True
                oorder2.Series = series1.ToString

                'oorder2.UserFields.Fields.Item("U_BaseEntry").Value = Objrs1.Fields.Item("docentry").Value.ToString
                'oorder2.UserFields.Fields.Item("U_BaseDB").Value = DBC
                If dts.Rows(0).Item("BPLid") = "N" Then
                Else
                    oorder2.BPL_IDAssignedToInvoice = dts.Rows(0).Item("BPLid").ToString
                End If
                oorder2.UserFields.Fields.Item("U_COSTCENTER").Value = costcenter
                oorder2.UserFields.Fields.Item("U_TrnsctnCtgry").Value = TransType
                'U_TrnsctnCtgry
                'U_COSTCENTER


                oorder2.NumAtCard = oorder1.NumAtCard

                oorder2.DocDate = oorder1.DocDate
                oorder2.DocDueDate = oorder1.DocDueDate
                oorder2.TaxDate = oorder1.TaxDate
                'oorder2.Comments = oorder1.Comments

                'For j As Integer = 0 To dtline

                '    oorder2.Lines.ItemCode = oorder1.Lines.ItemCode

                '    oorder2.Lines.WarehouseCode = whscode 'oorder1.Lines.WarehouseCode
                '    Dim loc As String = ""
                '    loc = objconnection1.GetSingleValue_SQL_SQL("select Location From [" & ToDB & "]..owhs where whscode='" & whscode & "'")
                '    oorder2.Lines.LocationCode = loc.ToString
                '    oorder2.Lines.Quantity = oorder1.Lines.Quantity
                '    'oorder2.Lines.TaxCode = oorder1.Lines.TaxCode
                '    oorder2.Lines.LineTotal = oorder1.Lines.LineTotal
                '    oorder2.Lines.AccountCode = dts.Rows(0).Item("Acc")
                '    oorder2.Lines.Add()
                'Next


                For j As Integer = 0 To dtLines.Rows.Count - 1

                    oorder2.Lines.ItemCode = dtLines.Rows(j).Item("Itemcode")

                    'oorder2.Lines.WarehouseCode = whscode 'oorder1.Lines.WarehouseCode

                    Dim Mainwhscode As String = ""
                    Mainwhscode = objconnection1.GetSingleValue_SQL_SQL("select U_Priority From [" & ToDB & "]..oitm where ItemCode='" & dtLines.Rows(j).Item("Itemcode") & "'")
                    If Mainwhscode = "2" Then
                        oorder2.Lines.WarehouseCode = RMWHSCode 'oorder1.Lines.WarehouseCode
                    Else
                        oorder2.Lines.WarehouseCode = whscode 'oorder1.Lines.WarehouseCode
                    End If


                    Dim loc As String = ""
                    loc = objconnection1.GetSingleValue_SQL_SQL("select Location From [" & ToDB & "]..owhs where whscode='" & whscode & "'")
                    oorder2.Lines.LocationCode = loc.ToString
                    oorder2.Lines.Quantity = dtLines.Rows(j).Item("Quantity")
                    If dtLines.Rows(j).Item("TaxCode").ToString = "0" Then
                    Else
                        oorder2.Lines.TaxCode = dtLines.Rows(j).Item("TaxCode")
                    End If

                    oorder2.Lines.LineTotal = dtLines.Rows(j).Item("LineTotal")
                    oorder2.Lines.AccountCode = dts.Rows(0).Item("Acc").ToString
                    oorder2.Lines.UserFields.Fields.Item("U_PairQty").Value = dtLines.Rows(j).Item("U_PairQty").ToString
                    oorder2.Lines.UserFields.Fields.Item("U_TrnsctnCtgry").Value = dtLines.Rows(j).Item("U_trnsctnCtgry").ToString

                    'U_PairQty,U_trnsctnCtgry
                    oorder2.Lines.Add()
                Next

                oorder2.DiscountPercent = oorder1.DiscountPercent
                If oorder1.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    oorder2.RoundingDiffAmount = oorder1.RoundingDiffAmount
                Else
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                lretcode = oorder2.Add
                Dim goodsreceiptno As Long
                If lretcode <> 0 Then

                    strsql1 = "update mipllogs set Flag='N',Errorlog='GR_ADD" & Replace(objToCompany.GetLastErrorDescription, "'", "") & "' where code='" & MainCode & "'"
                    objconnection1.Fun_ErrorLog(strsql1)
                    ' MsgBox(objToCompany.GetLastErrorDescription)
                    write_log("Transaction Failed:" & objToCompany.GetLastErrorDescription)
                    If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    objToCompany.GetNewObjectCode(goodsreceiptno)
                    Dim strsqlT As String = ""
                    strsqlT = "update DRF1 set OcrCode='" & costcenter & "' where DocEntry='" & goodsreceiptno & "'"
                    Objrs2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsqlT)

                    strsql2 = "update OIGE set U_Flag='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    Objrs2 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsql2)

                    strsql1 = "update mipllogs set Flag='Y',Errorlog='Success',ToDB='" & ToDB & "',TargetEntry='" & goodsreceiptno & "' where code='" & MainCode & "'"
                    cmd = New SqlCommand(strsql1, con)
                    cmd.ExecuteNonQuery()
                    If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    write_log("Transaction Success!!!")
                    Objrs1.MoveNext()
                End If
            Next
            objToCompany.Disconnect()
            'updatesalesorder12()
            'updatesalesorder21()
        Catch ex As Exception
            If updatestatus = True Then
                If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objToCompany.Disconnect()
            End If
            strsql1 = "update mipllogs set Flag='N',Errorlog='EX" & Replace(ex.ToString, "'", "") & "' where code='" & MainCode & "'"
            objconnection1.Fun_ErrorLog(strsql1)
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Fun_AR_To_GRPO(ByVal DBC As String, ByVal MBaseEntry As String, ByVal MainCode As String)
        ''Add Sales Invoice
        Dim updatestatus As Boolean = False
        Dim Errorlog As String = ""
        Try
            ' objconnection1.CompanyToConnection(DBC)

            strsql1 = "select a.DocEntry,a.CardCode,b.ProfitCode'COST',a.DiscPrcnt from oinv a join jdt1 b on a.transid=b.transid and b.line_ID=0  where a.DocEntry='" & MBaseEntry & "'"
            Objrs1 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            For i As Integer = 0 To Objrs1.RecordCount - 1
                Dim vendorcode As String = ""
                Dim cvendorcode As String = ""
                Dim whscode As String = ""
                Dim costcenter As String = ""
                Dim ToDB As String = ""
                Dim series As String = ""
                Dim series1 As String = ""
                Dim LinsSql As String = ""
                Dim FrightSql As String = ""
                Dim FrightSql_code As String = ""
                Dim FrightSqlC As String = ""
                Dim FrightTaxD As String = ""
                Dim RMWHSCode As String = ""
                LinsSql = "select Itemcode,Quantity,LineTotal,TaxCode,DiscPrcnt,priceBefdi,U_HSNCode,U_PairPrice,U_PairQty,U_trnsctnCtgry,U_MRP From [" & DBC & "]..inv1 where docentry='" & Objrs1.Fields.Item("DocEntry").Value & "' and isnull(Itemcode,'')!='' order by Linenum asc"
                strsql1 = "select U_DfltWhs 'ToWhscode',U_RmWhs'RMWHS',U_CostCenter 'ToCostcentre',U_DB 'ToDB',isnull(convert(varchar,U_BPLid),'N') 'BPLid' from [@MIPL_INTERCOMPANY] where U_CardCode='" & Objrs1.Fields.Item("CardCode").Value & "'"
                FrightSql = "select LineTotal,TaxCode,(select top 1 ExpnsName from oexd where ExpnsName like 'Delivery%') ExpName,U_STapplcble,Comments,OcrCode,Distrbmthd,TaxDistMtd From [" & DBC & "]..inv3 where docentry='" & Objrs1.Fields.Item("DocEntry").Value & "' order by Linenum asc"
                Dim dts As New DataTable
                Dim dtLines As New DataTable
                Dim dtFright As New DataTable
                dts = objconnection1.GetSingleValue_SQL_dt(strsql1)
                dtLines = objconnection1.GetSingleValue_SQL_dt(LinsSql)
                dtFright = objconnection1.GetSingleValue_SQL_dt(FrightSql)
                If dts.Rows.Count = 0 Then
                    Errorlogs = "To DB Whs Code Missing : " & cvendorcode
                    strsql1 = "update mipllogs set Flag='N',Errorlog='EX" & Replace(Errorlogs.ToString, "'", "") & "' where code='" & MainCode & "'"
                    objconnection1.Fun_ErrorLog(strsql1)
                    write_log(Replace(Errorlogs.ToString, "'", ""))
                    Exit Sub
                End If
                cvendorcode = Objrs1.Fields.Item("COST").Value
                costcenter = dts.Rows(0).Item("ToCostcentre")

                write_log("Cost Center:" & costcenter)
                whscode = dts.Rows(0).Item("ToWhscode")
                RMWHSCode = dts.Rows(0).Item("RMWHS")
                write_log("Whs Code:" & whscode)

                Errorlog = Errorlog + "To Cost:" + costcenter + " : To Whs Code:" + whscode + ": To Whs Code 1: " + RMWHSCode
                strsql1 = "select U_CardCode  From [@MIPL_INTERCOMPANY] where U_CardType='S'  AND U_CostCenter='" & cvendorcode & "'"
                vendorcode = objconnection1.GetSingleValue_SQL_SQL(strsql1)
                Errorlog = Errorlog + " : To Vendor code: " + vendorcode
                If objconnection1.GetSingleValue_SQL_SQL(strsql1) = "" Then
                    Errorlogs = "To DB Vendor Code Missing : " & cvendorcode
                    strsql1 = "update mipllogs set Flag='N',Errorlog='EX" & Replace(Errorlogs.ToString, "'", "") & "' where code='" & MainCode & "'"
                    objconnection1.Fun_ErrorLog(strsql1)
                    write_log(Replace(Errorlogs.ToString, "'", ""))
                    Exit Sub
                End If
                write_log("Vendor Code:" & vendorcode)
                write_log("TO DB:" & dts.Rows(0).Item("ToDB"))
                Dim oorder1 As SAPbobsCOM.Documents
                If dts.Rows.Count > 0 Then
                    oorder1 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    oorder1.GetByKey(Objrs1.Fields.Item("DocEntry").Value.ToString)
                    ' vendorcode = dts.Rows(0).Item("U_CardCode")
                    ToDB = dts.Rows(0).Item("ToDB")
                    ' costcenter = dts.Rows(0).Item("U_CostCenter")
                    'whscode = dts.Rows(0).Item("U_DfltWhs")
                    objconnection1.CompanyToConnection(ToDB)
                    oss.createTables_to()

                    strsql2 = "SELECT distinct  T0.[U_Series] FROM [" & ToDB & "]..[@COSTCENTRE]  T0 where 	T0.U_BranchCode='" & costcenter & "' and T0.U_Doctype='20'"
                    series = objconnection1.GetSingleValue_SQL_SQL(strsql2)
                    strsql2 = "SELECT Series FROM [" & ToDB & "]..NNM1 where 	SeriesName='" & series & "'"

                    series1 = objconnection1.GetSingleValue_SQL_SQL(strsql2)
                End If
                'strsql2 = "select 1 from ODRF where U_BaseEntry='" & Objrs1.Fields.Item("docentry").Value & "'"
                'Objrs2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Objrs2.DoQuery(strsql2)
                'If Objrs2.RecordCount > 0 Then Continue For

                If Not objToCompany.InTransaction Then objToCompany.StartTransaction()
                Dim oorder2 As SAPbobsCOM.Documents
                'oorder2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                oorder2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                oorder2.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
                oorder2.CardCode = vendorcode
                updatestatus = True
                oorder2.Series = series1.ToString
                If dts.Rows(0).Item("BPLid") = "N" Then
                Else
                    oorder2.BPL_IDAssignedToInvoice = dts.Rows(0).Item("BPLid").ToString
                End If
                'oorder2.UserFields.Fields.Item("U_BaseEntry").Value = Objrs1.Fields.Item("docentry").Value.ToString
                'oorder2.UserFields.Fields.Item("U_BaseDB").Value = DBC
                oorder2.UserFields.Fields.Item("U_COSTCENTER").Value = costcenter
                'U_COSTCENTER


                oorder2.NumAtCard = oorder1.NumAtCard

                oorder2.DocDate = oorder1.DocDate
                oorder2.DocDueDate = oorder1.DocDueDate
                oorder2.TaxDate = oorder1.TaxDate
                'oorder2.Comments = oorder1.Comments

                For j As Integer = 0 To dtLines.Rows.Count - 1

                    oorder2.Lines.ItemCode = dtLines.Rows(j).Item("Itemcode")
                    Dim todbitemcode As String = ""
                    todbitemcode = objconnection1.GetSingleValue_SQL_SQL("select ItemCode From [" & ToDB & "]..oitm where ItemCode='" & dtLines.Rows(j).Item("Itemcode").ToString.Trim & "'")
                    If todbitemcode = "" Then
                        Errorlog = ": ToDB ItemCode Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + "TODB:" + ToDB : write_log(": ToDB ItemCode Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + "TODB:" + ToDB)
                    End If

                    Dim todbitemcode1 As String = ""
                    todbitemcode1 = objconnection1.GetSingleValue_SQL_SQL("select U_Priority From [" & ToDB & "]..oitm where ItemCode='" & dtLines.Rows(j).Item("Itemcode").ToString.Trim & "'")
                    If todbitemcode = "" Then
                    Else
                        If todbitemcode1 = "" Then
                            Errorlog = ": ToDB Priority Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + "TODB:" + ToDB : write_log(": ToDB Priority Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + "TODB:" + ToDB)
                        End If
                    End If

                    Dim Mainwhscode As String = ""
                    Mainwhscode = objconnection1.GetSingleValue_SQL_SQL("select U_Priority From [" & ToDB & "]..oitm where ItemCode='" & dtLines.Rows(j).Item("Itemcode") & "'")
                    If Mainwhscode = "" Then Errorlog = ": ToDB Warehouse Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + " TODB:" + ToDB : write_log(": ToDB Warehouse Missing for this Itemcode: " + dtLines.Rows(j).Item("Itemcode") + Space(2) + " TODB:" + ToDB)

                    If Mainwhscode = "2" Then
                        oorder2.Lines.WarehouseCode = RMWHSCode 'oorder1.Lines.WarehouseCode
                    Else
                        oorder2.Lines.WarehouseCode = whscode 'oorder1.Lines.WarehouseCode
                    End If
                    ' oorder2.Lines.WarehouseCode = whscode 'oorder1.Lines.WarehouseCode
                    Dim loc As String = ""
                    loc = objconnection1.GetSingleValue_SQL_SQL("select Location From [" & ToDB & "]..owhs where whscode='" & whscode & "'")
                    oorder2.Lines.LocationCode = loc.ToString
                    oorder2.Lines.Quantity = dtLines.Rows(j).Item("Quantity")
                    oorder2.Lines.UnitPrice = dtLines.Rows(j).Item("priceBefdi")
                    ' oorder2.Lines.DiscountPercent = dtLines.Rows(j).Item("DiscPrcnt")
                    oorder2.Lines.TaxCode = dtLines.Rows(j).Item("TaxCode")
                    oorder2.Lines.LineTotal = dtLines.Rows(j).Item("LineTotal")
                    oorder2.Lines.UserFields.Fields.Item("U_HSNCode").Value = dtLines.Rows(j).Item("U_HSNCode")
                    oorder2.Lines.UserFields.Fields.Item("U_PairPrice").Value = dtLines.Rows(j).Item("U_PairPrice").ToString
                    oorder2.Lines.UserFields.Fields.Item("U_PairQty").Value = dtLines.Rows(j).Item("U_PairQty").ToString
                    oorder2.Lines.UserFields.Fields.Item("U_MRP").Value = dtLines.Rows(j).Item("U_MRP").ToString
                    'U_MRP
                    'U_HSNCode,U_PairPrice,U_PairQty,U_trnsctnCtgry
                    oorder2.Lines.Add()
                Next

                oorder2.DiscountPercent = oorder1.DiscountPercent
                If oorder1.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    oorder2.RoundingDiffAmount = oorder1.RoundingDiffAmount
                Else
                    oorder2.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If


                If Val(dtFright.Rows.Count) > 0 Then
                    For fr As Integer = 0 To dtFright.Rows.Count - 1
                        If dtFright.Rows(fr).Item("ExpName").ToString = "DELIVERY CHARGES" Then
                            Dim FC As String = ""
                            FC = "select ExpnsCode from [" & ToDB & "]..OEXD where ExpnsName like 'Carriage%'"
                            FrightSqlC = objconnection1.GetSingleValue_SQL_SQL(FC)
                            oorder2.Expenses.ExpenseCode = FrightSqlC
                            oorder2.Expenses.TaxCode = dtFright.Rows(fr).Item("TaxCode").ToString
                            oorder2.Expenses.LineTotal = dtFright.Rows(fr).Item("LineTotal").ToString
                            oorder2.Expenses.Remarks = dtFright.Rows(fr).Item("Comments").ToString
                            oorder2.Expenses.UserFields.Fields.Item("U_STapplcble").Value = dtFright.Rows(fr).Item("U_STapplcble").ToString
                            oorder2.Expenses.DistributionRule = costcenter 'dtFright.Rows(fr).Item("OcrCode").ToString
                            'oorder2.Expenses.VatGroup = dtFright.Rows(fr).Item("TaxDistMtd").ToString
                            FrightTaxD = dtFright.Rows(fr).Item("TaxDistMtd").ToString
                            'oorder2.Expenses.DistributionRule5 = dtFright.Rows(fr).Item("TaxDistMtd").ToString
                            oorder2.Expenses.Remarks = dtFright.Rows(fr).Item("Comments").ToString
                            If dtFright.Rows(fr).Item("Distrbmthd").ToString = "Q" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_Quantity
                            ElseIf dtFright.Rows(fr).Item("Distrbmthd").ToString = "W" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_Weight
                            ElseIf dtFright.Rows(fr).Item("Distrbmthd").ToString = "E" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_Equally
                            ElseIf dtFright.Rows(fr).Item("Distrbmthd").ToString = "N" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_None
                            ElseIf dtFright.Rows(fr).Item("Distrbmthd").ToString = "R" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_RowTotal

                            ElseIf dtFright.Rows(fr).Item("Distrbmthd").ToString = "V" Then
                                oorder2.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_Volume
                            End If

                            'OcrCode,Distrbmthd,TaxDistMtd
                            'U_STapplcble,Comments
                            oorder2.Expenses.Add()
                        End If

                    Next
                End If

                lretcode = oorder2.Add
                Dim goodsreceiptno As Long
                If lretcode <> 0 Then

                    strsql1 = "update mipllogs set Flag='N',Errorlog='GRPO_ADD" & Errorlog & " ErrorDescription-" & Replace(objToCompany.GetLastErrorDescription, "'", "") & "' where code='" & MainCode & "'"
                    objconnection1.Fun_ErrorLog(strsql1)
                    ' MsgBox(objToCompany.GetLastErrorDescription)
                    write_log("Transaction Failed:" & objToCompany.GetLastErrorDescription)
                    If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    Objrs1.MoveNext()
                    'Exit Sub
                Else
                    objToCompany.GetNewObjectCode(goodsreceiptno)
                    Dim strsqlT As String = ""
                    strsqlT = "UPDATE DRF3 SET TaxDistMtd='" & FrightTaxD & "' where docentry='" & goodsreceiptno & "'"
                    Objrs2 = objToCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Objrs2.DoQuery(strsqlT)
                    'strsql2 = "update OINV set U_Flag='Y' where DocEntry='" & Objrs1.Fields.Item("DocEntry").Value.ToString & "'"
                    'Objrs2 = objFromCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'Objrs2.DoQuery(strsql2)
                    strsql1 = "update mipllogs set Flag='Y',Errorlog='Success',ToDB='" & ToDB & "',TargetEntry='" & goodsreceiptno & "' where code='" & MainCode & "'"
                    cmd = New SqlCommand(strsql1, con)
                    cmd.ExecuteNonQuery()
                    If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    write_log("Transaction Success!!!")
                    Objrs1.MoveNext()
                End If
            Next

            objToCompany.Disconnect()
            'updatesalesorder12()
            'updatesalesorder21()
        Catch ex As Exception
            If updatestatus = True Then
                If objToCompany.InTransaction Then objToCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objToCompany.Disconnect()
            End If
            strsql1 = "update mipllogs set Flag='N',Errorlog='EX" & Errorlog + Space(2) & "' where code='" & MainCode & "'"
            objconnection1.Fun_ErrorLog(strsql1)
            write_log("Transaction Exception:" & ex.Message & "Error Log-" & Errorlog)
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub updatesalesorder12()
        strsql1 = "select docentry,U_STATUS from ordr  where U_transfer2='N'"
        Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs1.DoQuery(strsql1)
        For i As Integer = 0 To Objrs1.RecordCount - 1
            strsql2 = "select 1 from ordr where docentry='" & Objrs1.Fields.Item("docentry").Value & "'"
            Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs2.DoQuery(strsql2)
            If Objrs2.RecordCount > 0 Then
                strsql1 = "select linenum,U_LINESTATUS  from rdr1 where docentry='" & Objrs1.Fields.Item("docentry").Value & "'"
                Objrs11 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs11.DoQuery(strsql1)
                strsql2 = ""
                For j As Integer = 0 To Objrs11.RecordCount - 1
                    strsql2 += vbCrLf + "update RDR1 set U_LINESTATUS ='" & Objrs11.Fields.Item("U_LINESTATUS").Value & "' where DocEntry = '" & Objrs1.Fields.Item("docentry").Value & "' and LineNum='" & Objrs11.Fields.Item("linenum").Value & "'"
                Next
                strsql2 += vbCrLf + " update ordr set U_STATUS='" & Objrs1.Fields.Item("U_STATUS").Value & "' where docentry= '" & Objrs1.Fields.Item("docentry").Value & "'"
                strsql2 += vbCrLf + " update " & MDBName1.ToString & "..ordr set U_transfer2='Y' where docentry= '" & Objrs1.Fields.Item("docentry").Value & "'"
                Objrs2 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs2.DoQuery(strsql2)
            End If
            Objrs1.MoveNext()
        Next
    End Sub

    Private Sub updatesalesorder21()
        strsql2 = "select docentry,U_STATUS from ordr  where U_transfer1='N'"
        Objrs2 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Objrs2.DoQuery(strsql2)
        For i As Integer = 0 To Objrs2.RecordCount - 1
            strsql1 = "select 1 from ordr where docentry='" & Objrs2.Fields.Item("docentry").Value & "'"
            Objrs1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Objrs1.DoQuery(strsql1)
            If Objrs1.RecordCount > 0 Then
                strsql2 = "select linenum,U_LINESTATUS  from rdr1 where docentry='" & Objrs2.Fields.Item("docentry").Value & "'"
                Objrs11 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs11.DoQuery(strsql2)
                strsql1 = ""
                For j As Integer = 0 To Objrs11.RecordCount - 1
                    strsql1 += vbCrLf + "update RDR1 set U_LINESTATUS ='" & Objrs11.Fields.Item("U_LINESTATUS").Value & "' where DocEntry = '" & Objrs2.Fields.Item("docentry").Value & "' and LineNum='" & Objrs11.Fields.Item("linenum").Value & "'"
                    Objrs11.MoveNext()
                Next
                strsql1 += vbCrLf + " update ordr set U_STATUS='" & Objrs2.Fields.Item("U_STATUS").Value & "' where docentry= '" & Objrs2.Fields.Item("docentry").Value & "'"
                strsql1 += vbCrLf + " update " & MDBName2.ToString & "..ordr set U_transfer1='Y' where docentry= '" & Objrs2.Fields.Item("docentry").Value & "'"
                Objrs1 = objcompany2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Objrs1.DoQuery(strsql1)
            End If
            Objrs2.MoveNext()
        Next
    End Sub

End Class

