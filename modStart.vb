Imports System.Data.SqlClient
Module modStart
    Public Sub Main()
        Try
            Dim remoteAll As Process() = Process.GetProcesses()

            'Connection Enabling functions
            Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
            If CommandLineArgs.Count > 0 Then
                form_Name = CommandLineArgs(0)
                ' form_Name = "MATERIALRECEIPT21"
            End If

            Dim CurrentProcess As Process()
            Dim CurrentProccessName As String = Process.GetCurrentProcess.ProcessName
            CurrentProcess = Process.GetProcessesByName(CurrentProccessName)

            Dim objconnection As clsConnections
            Dim objs As New clsAddOn
            objconnection = New clsConnections
            objconnection.dbconnection()

            If con.State <> ConnectionState.Open Then
                con.Close()
                Application.Exit()
            End If

            'If CurrentProcess.Length > 0 Then
            '    UpdateProcessTobeDone(form_Name)
            '    con.Close()
            '    Application.Exit()
            'End If

            'objconnection.connection()

            'If objcompany1.Connected = False Then
            '    objcompany1.Disconnect()
            '    Application.Exit()
            'End If

            'If objcompany2.Connected = False Then
            '    objcompany2.Disconnect()
            '    Application.Exit()
            'End If

            'Transfer Functions Calling
            'If remoteAll.ToString <> "DataTransfer" Then
            If remoteAll.ToString = "DataTransfer" Then
            Else
                Dim objtransfer As New clsTransfer
                objtransfer.transfer_document() ''Transfer Functions Calling
            End If

            'Else
            'End
            'End If

            'objFromCompany.Disconnect()
            ' objToCompany.Disconnect()
            Application.Exit()

        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub
    Public Function GetPendingProcess(Optional ByVal Process As String = "") As String
        Try
            Using comm As New SqlCommand
                comm.Connection = con
                comm.CommandText = "SP_DT_Process_Pending"
                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.Add("@Process", SqlDbType.VarChar).Value = Process
                Return comm.ExecuteScalar().ToString()
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function UpdateProcessTobeDone(ByVal Process As String) As String
        Try
            Using comm As New SqlCommand
                comm.Connection = con
                comm.CommandText = "SP_DT_Process_Insert"
                comm.CommandType = CommandType.StoredProcedure
                comm.Parameters.Add("@Process", SqlDbType.VarChar).Value = Process
                Return comm.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

End Module
