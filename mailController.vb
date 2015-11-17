Imports System.IO
''' <summary>
''' Controle de envio de email pelo exchange
''' </summary>
''' <remarks>enjoy</remarks>
Public Class mailController
    Public Sub sent(ByVal strDest As String, ByVal strAssnt As String, ByVal strBody As String)

        Dim olapp As Object
        Dim oitem As Object
        Dim errmy As Integer = 0
        Dim DADOS As New DataTable("DADOS")
        Dim VARDEST As String = ""
        olapp = CreateObject("Outlook.Application")
        oitem = olapp.CreateItem(0)

        VARDEST = "fagner.silva@net.com.br;"

        With oitem
            .Subject = "ERR " & strAssnt & " - " & Environment.MachineName
            .To = VARDEST
            .HTMLBody = strBody
        End With

        oitem.Send()
        oitem.dispose()

        Threading.Thread.Sleep(6000) 'AGUARDA de 1

    End Sub
End Class
