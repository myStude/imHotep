Imports System
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
''' <summary>
''' CLASS IMOVEIS
''' </summary>
''' <remarks>SE NAO TIVER CONEXAO, FALHA</remarks>
Public Class IMOVEIS
    Inherits objConnection

    Private strComando As String

    ''' <summary>
    ''' ATUALIZA IMOVEIS
    ''' </summary>
    ''' <param name="CodOperadora">INT 3</param>
    ''' <param name="CodCidade">INT 6</param>
    ''' <remarks>SE NAO TIVER CONEXAO, FALHA</remarks>
    Sub importaImoveis(ByVal CodCidade As String, ByVal CodOperadora As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)

        Dim FIELDS, TABLE, JOINS, WHERES As String

        FIELDS = " COD_OPERADORA, COD_IMOVEL, COD_NODE, COD_CIDADE, COD_CELULA, COD_LOGRADOURO, COD_ENDERECO, IND_TIPO_EDIFICACAO, to_char(DT_ATUALIZACAO,'YYYY-MM-DD HH24:MI:SS') DT_ATUALIZACAO_ORA  "
        TABLE = " GED.IMOVEL "
        JOINS = ""
        WHERES = " cod_operadora in(" & CodOperadora & ") and cod_cidade in('" & CodCidade & "') " _
                & " AND DT_ATUALIZACAO BETWEEN TO_DATE('" & CDate(Format(Now, "dd/MM/yyyy").ToString).AddDays(-3) & " 00','DD/MM/YYYY HH24') " _
                & " AND TO_DATE('" & Format(Now, "dd/MM/yyyy").ToString & " " & Format(Now, "HH").ToString & "','DD/MM/YYYY HH24')"


        Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES

        Dim tbInsert As String = "ora_cfg.cfg_imoveis (COD_OPERADORA,  COD_IMOVEL,  COD_NODE,  COD_CIDADE,  COD_CELULA,  COD_LOGRADOURO,  COD_ENDERECO,  IND_TIPO_EDIFICACAO, DT_UPDATE, DT_ATUALIZACAO_ORA)"
        Dim strSql, strDetail As String
        Try
            '--------------------------
            '#  DECLARA VARIAVEIS ORA
            oraComm = New OracleCommand(QryStr, oraConn)

            Using GerRow As OracleDataReader = oraComm.ExecuteReader()
                While GerRow.Read()

                    strSql = "('" & GerRow.Item("COD_OPERADORA").ToString & "','" & GerRow.Item("COD_IMOVEL") & "'," & _
                                "'" & GerRow.Item("COD_NODE") & "','" & GerRow.Item("COD_CIDADE") & "','" & GerRow.Item("COD_CELULA") & "'," & _
                                "'" & GerRow.Item("COD_LOGRADOURO") & "','" & GerRow.Item("COD_ENDERECO") & "'," & _
                                "'" & GerRow.Item("IND_TIPO_EDIFICACAO") & "','" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'," & _
                                "'" & GerRow.Item("DT_ATUALIZACAO_ORA") & "') "

                    strDetail = " ON DUPLICATE KEY UPDATE " _
                                    & "COD_NODE = '" & GerRow.Item("COD_NODE") & "', COD_CELULA = '" & GerRow.Item("COD_CELULA") & "', " _
                                    & "COD_LOGRADOURO = '" & GerRow.Item("COD_LOGRADOURO") & "', " _
                                    & "COD_ENDERECO = '" & GerRow.Item("COD_ENDERECO") & "', " _
                                    & "IND_TIPO_EDIFICACAO = '" & GerRow.Item("IND_TIPO_EDIFICACAO") & "', " _
                                    & "DT_UPDATE = '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "' , " _
                                    & "DT_ATUALIZACAO_ORA = '" & GerRow.Item("DT_ATUALIZACAO_ORA") & "'"

                    MysqlAdd("", tbInsert, strSql, strDetail, myComm, MyConn)

                End While
            End Using
        Catch ex As Exception
            '# ErrCatcher
            'MsgBox(ex.Message)

            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, strComando)
        End Try


    End Sub

End Class