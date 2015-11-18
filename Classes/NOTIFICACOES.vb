Imports System
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types


Public Class NOTIFICACOES
    Inherits objConnection

    Private NotificacoesQuery As New Notificacoes_Query
    Private MyCommand As New MySqlCommand
    Private strComando As String

    Sub ATUALIZA_NOTIFICACOES(ByVal dtHora As String, ByVal codCidade As String, ByVal BASE As String)

        Dim QryStr As String = NotificacoesQuery.NotificacoesQuery(codCidade, dtHora)

        Try
            '--------------------------
            '#  DECLARA VARIAVEIS ORA
            Using Conn As New OracleConnection(oraConnection(BASE)) : Conn.Open()
                Dim com As New OracleCommand
                com = New OracleCommand(QryStr, Conn)

                Using GerRow As OracleDataReader = com.ExecuteReader()
                    Using MyConn As New MySqlConnection(MySqlConnection("ora_ord"))
                        MyConn.Open()
                        While GerRow.Read()
                            ATUALIZA_NOTIFICACOES( _
                                GerRow.Item("ID_NOTIFICACAO").ToString, _
                                GerRow.Item("CI_CODIGO").ToString, _
                                GerRow.Item("DESCRICAO").ToString, _
                                GerRow.Item("STN_DESCRICAO").ToString, _
                                GerRow.Item("US_CODIGO").ToString, _
                                GerRow.Item("fechamento").ToString, _
                                GerRow.Item("OBS").ToString, _
                                GerRow.Item("DT_AGENDAMENTO").ToString, _
                                GerRow.Item("abertura").ToString, _
                                GerRow.Item("ID_SINTOMA_NOTIFICACAO").ToString, _
                                BASE, MyConn)
                        End While
                        MyConn.Dispose()
                    End Using
                    Conn.Dispose()
                End Using
            End Using
        Catch ex As Exception
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, QryStr)
        End Try

    End Sub

    Sub ATUALIZA_NOTIFICACOES(ByVal ID_NOTIFICACAO As String, ByVal CI_CODIGO As String, ByVal DESCRICAO As String, ByVal STN_DESCRICAO As String, ByVal US_CODIGO As String, ByVal fechamento As String, ByVal OBS As String, ByVal DT_AGENDAMENTO As String, ByVal abertura As String, ByVal ID_SINTOMA_NOTIFICACAO As String, ByVal BASE As String, ByVal strConn As MySqlConnection)

        Dim M As String
        If ID_SINTOMA_NOTIFICACAO <> "" Then M = ", " & ID_SINTOMA_NOTIFICACAO Else M = ", null"

        strComando = "INSERT INTO ORA_ORD.ORD_NOTIFICACAO( " & _
            "ID_NOTIFICACAO,CID_CONTRATO,DESCRICAO,STN_DESCRICAO, " & _
            "US_CODIGO,fechamento,OBS,DT_AGENDAMENTO, " & _
            "abertura,ID_SINTOMA_NOTIFICACAO, dt_atualizacao, base " & _
            ")VALUES( " & _
            ID_NOTIFICACAO & ", " & CI_CODIGO & ", '" & DESCRICAO & "', '" & STN_DESCRICAO & "', '" & _
            US_CODIGO & "', '" & fechamento & "', '" & OBS.Replace("'", ".") & "', '" & DT_AGENDAMENTO & "', '" & _
            abertura & "'" & M & ", now(), '" & BASE & "' )" & _
            "on duplicate key update " & _
            "DESCRICAO  =     '" & DESCRICAO & "', " & _
            "STN_DESCRICAO =  '" & STN_DESCRICAO & "', " & _
            "US_CODIGO =      '" & US_CODIGO & "', " & _
            "fechamento =     '" & fechamento & "', " & _
            "OBS =            '" & OBS.Replace("'", ".") & "', " & _
            "DT_AGENDAMENTO = '" & DT_AGENDAMENTO & "', " & _
            "ID_SINTOMA_NOTIFICACAO = " & M.Replace(",", "") & ", " & _
            "dt_atualizacao = now();"

        strComando = strComando.Replace("\", "")

        MyCommand = New MySqlCommand(strComando, strConn)
        MyCommand.ExecuteNonQuery()
    End Sub

End Class
