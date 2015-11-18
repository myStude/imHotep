Imports System
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
''' <summary>
''' CLASSE IEs
''' </summary>
''' <remarks></remarks>
Public Class IES_
    Inherits objConnection

    Private IEsquery As New IEs_query
    Private MyCommand As New MySqlCommand

    Private strComando As String

    ''' <summary>
    ''' ATUALIZA IEs
    ''' </summary>
    ''' <param name="dtHora">DATA DD/MM/YYYY</param>
    ''' <param name="codCidade">INT 6</param>
    ''' <param name="BASE">STR 6</param>
    ''' <remarks>SE NAO TIVER CONEXAO, FALHA</remarks>
    Sub IES(ByVal dtHora As String, ByVal codCidade As String, ByVal BASE As String)
        Try
            '--------------------------
            '#  DECLARA VARIAVEIS ORA
            Using Conn As New OracleConnection(oraConnection(BASE)) : Conn.Open()
                Dim com As New OracleCommand
                Dim QryStr As String = IEsquery.IEsQuery(dtHora, codCidade, BASE)
                com = New OracleCommand(QryStr, Conn)
                Using GerRow As OracleDataReader = com.ExecuteReader()
                    Using MyConn As New MySqlConnection(MySqlConnection("ora_ord"))
                        MyConn.Open()
                        While GerRow.Read
                            ATUALIZA_IES( _
                                GerRow.Item("ID_OCORRENCIA").ToString, _
                                GerRow.Item("ID_TIPO_OCORRENCIA").ToString, _
                                GerRow.Item("NUM_CONTRATO").ToString, _
                                GerRow.Item("ID_ASSINANTE").ToString, _
                                GerRow.Item("DT_OCORRENCIA").ToString, _
                                GerRow.Item("RESOLUCAO").ToString, _
                                GerRow.Item("NOTIFICACAO").ToString, _
                                GerRow.Item("COD_NODE").ToString, _
                                GerRow.Item("CID_CONTRATO").ToString, BASE, MyConn)
                        End While
                        MyConn.Dispose()
                    End Using
                    Conn.Dispose()
                End Using
            End Using
        Catch ex As Exception
            ' ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, "")
        End Try

    End Sub

    Private Sub ATUALIZA_IES(ByVal ID_OCORRENCIA As String, ByVal ID_TIPO_OCORRENCIA As String, ByVal NUM_CONTRATO As String, ByVal ID_ASSINANTE As String, ByVal DT_OCORRENCIA As String, ByVal RESOLUCAO As String, ByVal NOTIFICACAO As String, ByVal NODE As String, ByVal CID_CONTRATO As String, ByVal BASE As String, ByVal SqlConn As MySqlConnection)
        If RESOLUCAO = "" Then RESOLUCAO = "null"
        If NOTIFICACAO = "" Then NOTIFICACAO = "null"

        strComando = "INSERT INTO ora_ord.ord_ies( " & vbNewLine & _
                        "ID_OCORRENCIA,ID_TIPO_OCORRENCIA,NUM_CONTRATO, " & vbNewLine & _
                        "ID_ASSINANTE,DT_OCORRENCIA,RESOLUCAO, " & vbNewLine & _
                        "NOTIFICACAO,NODE,LAST_UPDATE, CID_CONTRATO, BASE " & vbNewLine & _
                        ") " & vbNewLine & _
                        " VALUES( " & vbNewLine & _
                        ID_OCORRENCIA & ", " & ID_TIPO_OCORRENCIA & ", " & NUM_CONTRATO & ", " & vbNewLine & _
                        ID_ASSINANTE & ", '" & DT_OCORRENCIA & "', " & RESOLUCAO & ", " & vbNewLine & _
                        NOTIFICACAO & ", '" & NODE & "', NOW(), " & CID_CONTRATO & ", '" & BASE & "') " & vbNewLine & _
                        "ON DUPLICATE KEY UPDATE " & vbNewLine & _
                        "RESOLUCAO = " & RESOLUCAO & ", NOTIFICACAO = " & NOTIFICACAO & ", LAST_UPDATE = NOW() ; "

        MyCommand = New MySqlCommand(strComando, SqlConn)
        MyCommand.ExecuteNonQuery()
    End Sub

End Class
