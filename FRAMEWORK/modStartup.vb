Module modStartup
    Public objAddOn As clsAddOn
    Public objMdi_Form As MDI_JEWELADDON
    Public Sub Main()
        Try
            objAddOn = New clsAddOn
            objAddOn.Intialize("N")
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Module
