Imports MySql.Data.MySqlClient
Public Class BUSCA_CIDADES
    Inherits objConnection

    Public Function GetCidades() As List(Of cfgCidades)
        Dim cidList As New List(Of cfgCidades)

        Dim QryString As New Dictionary(Of String, String)
        Dim JsonString As New Dictionary(Of String, String)

        QryString("Select") = "CID_CONTRATO, COD_OPERADORA, NOME_CIDADE, LOWER(BASE) AS BASE"
        QryString("From") = "ora_cfg.cfg_cidades"
        QryString("Where") = "ESPELHO = 1 order by CID_CONTRATO "
        Dim QryStr As String = QryConstructor(QryString)

        Using MyConn As New MySqlConnection(MySqlConnection("ora_cfg"))
            MyConn.Open()

            Dim myComm As New MySqlCommand
            myComm = New MySqlCommand(QryStr, MyConn)

            Using sqlResult As MySqlDataReader = myComm.ExecuteReader()
                While sqlResult.Read()
                    cidList.Add(New cfgCidades(sqlResult.Item("CID_CONTRATO"), sqlResult.Item("COD_OPERADORA"), sqlResult.Item("BASE")))
                End While
            End Using
        End Using

        MySql.Data.MySqlClient.MySqlConnection.ClearAllPools()

        Return cidList
    End Function

    '// PROVISORIO - ATÉ TER TODAS AS CIDADES NO BANCO DE DADOS
    Function cidades_cancelado() As Dictionary(Of String, String)

        Dim cid_cancelado = New Dictionary(Of String, String)

        cid_cancelado("06294") = "netisp"
        cid_cancelado("06413") = "netsoc" 'CASCAVEL
        cid_cancelado("07193") = "netisp" 'ITAJAI
        cid_cancelado("06430") = "netsoc"
        cid_cancelado("06531") = "netsoc"
        cid_cancelado("06794") = "netsoc"
        cid_cancelado("07052") = "netisp"
        cid_cancelado("07278") = "netisp"
        cid_cancelado("07363") = "netisp"
        cid_cancelado("07584") = "netisp"
        cid_cancelado("07603") = "netisp"
        cid_cancelado("07757") = "netisp"
        cid_cancelado("07758") = "netisp"
        cid_cancelado("08330") = "netisp"
        cid_cancelado("08332") = "netisp"
        cid_cancelado("53902") = "netsul"
        cid_cancelado("55298") = "netsul"
        cid_cancelado("56995") = "netsul"
        cid_cancelado("57304") = "netsul"
        cid_cancelado("66770") = "netsul"
        cid_cancelado("67040") = "netsul"
        cid_cancelado("67784") = "netsul"
        cid_cancelado("68020") = "netsul"
        cid_cancelado("68659") = "netsul"
        cid_cancelado("69019") = "netsul"
        cid_cancelado("69337") = "netsul"
        cid_cancelado("70408") = "netsul"
        cid_cancelado("71242") = "netsul"
        cid_cancelado("71587") = "netsul"
        cid_cancelado("71706") = "netsul"
        cid_cancelado("71986") = "netsul"
        cid_cancelado("72451") = "netsul"
        cid_cancelado("72842") = "netsul"
        cid_cancelado("72907") = "netsul"
        cid_cancelado("74748") = "netsul"
        cid_cancelado("75680") = "netsul"
        cid_cancelado("75681") = "netsul"
        cid_cancelado("76066") = "netsul"
        cid_cancelado("76180") = "netsul"
        cid_cancelado("76414") = "netsul"
        cid_cancelado("77127") = "netsul"
        cid_cancelado("89710") = "netisp"

        Return cid_cancelado
    End Function

End Class
