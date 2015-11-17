Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
''' <summary>
''' CLASSE BAIRRO
''' </summary>
''' <remarks>IF NO CONECTION</remarks>
Public Class ENDERECO_BAIRRO
    Inherits objConnection

    ''' <summary>
    ''' ATAUALIZA TABELA BAIRRO
    ''' </summary>
    ''' <param name="cidContrato">INT 6</param>
    ''' <remarks>ERRO QUANDO PERDE A CONEXAO</remarks>
    Sub importaBairro(ByVal cidContrato As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)


        Dim FIELDS, TABLE, JOINS, WHERES As String

        FIELDS = " BRR.COD_BAIRRO, BRR.COD_CIDADE, BRR.NOM_BAIRRO, BRR.NOM_BAIRRO_ABREV "
        TABLE = " GED.BAIRRO BRR "
        JOINS = ""
        WHERES = " COD_CIDADE = '" & cidContrato & "'"


        Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & JOINS & " WHERE " & WHERES


        Try
            '--------------------------
            '#  DECLARA VARIAVEIS ORA
            oraComm = New OracleCommand(QryStr, oraConn)

            Using GerRow As OracleDataReader = oraComm.ExecuteReader() ' leitor ora
                While GerRow.Read()
                    atualizaBairro(GerRow.Item("COD_BAIRRO").ToString, GerRow.Item("COD_CIDADE").ToString, _
                                        GerRow.Item("NOM_BAIRRO").ToString, GerRow.Item("NOM_BAIRRO_ABREV").ToString, MyConn, myComm)
                End While
            End Using
        Catch ex As Exception
            '# - -  ERR - -
            'MsgBox(ex.Message)
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, "")
        End Try

    End Sub

    Private Sub atualizaBairro(ByVal COD_BAIRRO As String, ByVal COD_CIDADE As String, ByVal NOM_BAIRRO As String, ByVal NOM_BAIRRO_ABREV As String, ByVal myConn As MySqlConnection, ByVal myComm As MySqlCommand)
        Dim strComando As String
        Dim ERRO As Integer = 0

        NOM_BAIRRO = Replace(NOM_BAIRRO, "'", "\'")
        NOM_BAIRRO_ABREV = Replace(NOM_BAIRRO_ABREV, "'", "\'")

        strComando = "INSERT INTO " & _
            "  ora_cfg.cfg_bairro(COD_BAIRRO, COD_CIDADE, NOM_BAIRRO, NOM_BAIRRO_ABREV, DATA_UPDATE) " & _
            "VALUES " & _
            "  (" & COD_BAIRRO & ", " & COD_CIDADE & ", '" & NOM_BAIRRO & "', '" & NOM_BAIRRO_ABREV & "', CURDATE() ) " & _
            " ON DUPLICATE KEY UPDATE " & _
            "NOM_BAIRRO = '" & NOM_BAIRRO & "'," & _
            "NOM_BAIRRO_ABREV = '" & NOM_BAIRRO_ABREV & "'," & _
            "DATA_UPDATE = CURDATE() "

        myComm = New MySqlCommand(strComando, myConn)
        myComm.ExecuteNonQuery()
    End Sub
End Class
