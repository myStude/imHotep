Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
''' <summary>
''' CALSSE LOGRADOURO
''' </summary>
''' <remarks>IF NO CONNECTION IT FAIL</remarks>
Public Class ENDERECO_LOGRADOURO
    Inherits objConnection

    ''' <summary>
    ''' ATUALIZA LOGRADOURO
    ''' </summary>
    ''' <param name="cidContrato">INT 6</param>
    ''' <remarks>ERRO QUANDO NAO HA CONEXAO</remarks>
    Sub impotaEnderLogradouro(ByVal cidContrato As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)

        Dim FIELDS, TABLE, JOINS, WHERES As String


        FIELDS = "  L.COD_LOGRADOURO, L.COD_CIDADE,L.COD_TIPO_LOGR,  L.COD_OPERADORA,  L.COD_TITULO,   " & _
         "  L.NOM_LOGRADOURO, L.NOM_LOGR_ABREV, L.NOM_LOGR_COMPLETO, E.COD_ENDERECO, E.NUM_ENDERECO,  " & _
         "  E.COD_TIPO_COMPL1, E.TXT_TIPO_COMPL1, E.COD_TIPO_COMPL2, E.TXT_TIPO_COMPL2,  " & _
         "  E.COD_TIPO_COMPL3, E.TXT_TIPO_COMPL3, E.COD_TIPO_COMPL4, E.TXT_TIPO_COMPL4,  " & _
         "  E.TXT_COMPL_NUM, E.COD_IMOVEL, E.NOM_COMPLETO, COALESCE(E.COD_BAIRRO,0) AS COD_BAIRRO, " & _
         "  to_char(E.DT_ATUALIZACAO,'YYYY-MM-DD HH24:MI:SS') DT_ATUALIZACAO_ORA  "

        TABLE = " GED.ENDERECO E  "

        JOINS = " INNER JOIN GED.LOGRADOURO L on L.cod_logradouro = E.cod_logradouro AND E.COD_CIDADE = L.COD_CIDADE "
        WHERES = " E.COD_CIDADE = '" & cidContrato & "' " & _
                " AND E.DT_ATUALIZACAO BETWEEN TO_DATE('" & CDate(Format(Now, "dd/MM/yyyy").ToString).AddDays(-3) & " 00','DD/MM/YYYY HH24') " & _
                " AND TO_DATE('" & Format(Now, "dd/MM/yyyy").ToString & " " & Format(Now, "HH").ToString & "','DD/MM/YYYY HH24')"


        Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & JOINS & " WHERE " & WHERES




        Try
            '--------------------------
            '#  DECLARA VARIAVEIS ORA
            oraComm = New OracleCommand(QryStr, oraConn)

            Using GerRow As OracleDataReader = oraComm.ExecuteReader()

                'COM A CONEXAO MySQL, ENQUANTO HÁ LEITURA NO GerRow, INSERE MySQL
                While GerRow.Read()
                    atualizaEnderLogradouro(GerRow.Item("COD_LOGRADOURO").ToString, GerRow.Item("COD_CIDADE").ToString, _
                             GerRow.Item("COD_TIPO_LOGR").ToString, GerRow.Item("COD_OPERADORA").ToString, _
                             GerRow.Item("COD_TITULO").ToString, GerRow.Item("NOM_LOGRADOURO").ToString, _
                             GerRow.Item("NOM_LOGR_ABREV").ToString, GerRow.Item("NOM_LOGR_COMPLETO").ToString, _
                             GerRow.Item("COD_ENDERECO").ToString, GerRow.Item("NUM_ENDERECO"), GerRow.Item("COD_TIPO_COMPL1").ToString, _
                             GerRow.Item("TXT_TIPO_COMPL1").ToString, GerRow.Item("COD_TIPO_COMPL2").ToString, _
                             GerRow.Item("TXT_TIPO_COMPL2").ToString, GerRow.Item("COD_TIPO_COMPL3").ToString, _
                             GerRow.Item("TXT_TIPO_COMPL3").ToString, GerRow.Item("COD_TIPO_COMPL4").ToString, _
                             GerRow.Item("TXT_TIPO_COMPL4").ToString, GerRow.Item("TXT_COMPL_NUM").ToString, _
                             GerRow.Item("COD_IMOVEL").ToString, GerRow.Item("NOM_COMPLETO").ToString, _
                             GerRow.Item("COD_BAIRRO").ToString, GerRow.Item("DT_ATUALIZACAO_ORA").ToString, MyConn, myComm)
                End While
            End Using
        Catch ex As Exception
            '# - - err - -
            'MsgBox(ex.Message)
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, strComando)
        End Try

    End Sub

    Private Sub atualizaEnderLogradouro(ByVal COD_LOGRADOURO As String, ByVal COD_CIDADE As String, ByVal COD_TIPO_LOGR As String, ByVal COD_OPERADORA As String, ByVal COD_TITULO As String, ByVal NOM_LOGRADOURO As String, ByVal NOM_LOGR_ABREV As String, ByVal NOM_LOGR_COMPLETO As String, ByVal COD_ENDERECO As String, ByVal NUM_ENDERECO As String, ByVal COD_TIPO_COMPL1 As String, ByVal TXT_TIPO_COMPL1 As String, ByVal COD_TIPO_COMPL2 As String, ByVal TXT_TIPO_COMPL2 As String, ByVal COD_TIPO_COMPL3 As String, ByVal TXT_TIPO_COMPL3 As String, ByVal COD_TIPO_COMPL4 As String, ByVal TXT_TIPO_COMPL4 As String, ByVal TXT_COMPL_NUM As String, ByVal COD_IMOVEL As String, ByVal NOM_COMPLETO As String, ByVal COD_BAIRRO As String, ByVal DT_ATUALIZACAO_ORA As String, ByVal MyConn As MySqlConnection, ByVal MyComm As MySqlCommand)

        Dim ERRO As Integer = 0
        NOM_LOGRADOURO = NOM_LOGRADOURO.Replace("\", "#").ToString
        NOM_LOGRADOURO = NOM_LOGRADOURO.Replace("'", "#").ToString
        NOM_LOGR_ABREV = NOM_LOGR_ABREV.Replace("'", "#").ToString()
        NOM_LOGR_ABREV = NOM_LOGR_ABREV.Replace("\", "#").ToString()
        NOM_LOGR_COMPLETO = NOM_LOGR_COMPLETO.Replace("'", "#").ToString()
        NOM_LOGR_COMPLETO = NOM_LOGR_COMPLETO.Replace("\", "#").ToString()
        COD_ENDERECO = COD_ENDERECO.Replace("'", "#").ToString()
        COD_ENDERECO = COD_ENDERECO.Replace("\", "#").ToString()
        NUM_ENDERECO = NUM_ENDERECO.Replace("'", "#").ToString()
        NUM_ENDERECO = NUM_ENDERECO.Replace("\", "#").ToString()
        NUM_ENDERECO = NUM_ENDERECO.Replace("'", "#").ToString()
        If NUM_ENDERECO = "S/N" Then NUM_ENDERECO = 0
        COD_TIPO_COMPL1 = COD_TIPO_COMPL1.Replace("\", "#").ToString()
        TXT_TIPO_COMPL1 = TXT_TIPO_COMPL1.Replace("'", "#").ToString()
        TXT_TIPO_COMPL1 = TXT_TIPO_COMPL1.Replace("\", "#").ToString()
        COD_TIPO_COMPL2 = COD_TIPO_COMPL2.Replace("'", "#").ToString()
        COD_TIPO_COMPL2 = COD_TIPO_COMPL2.Replace("\", "#").ToString()
        TXT_TIPO_COMPL2 = TXT_TIPO_COMPL2.Replace("'", "#").ToString()
        TXT_TIPO_COMPL2 = TXT_TIPO_COMPL2.Replace("\", "#").ToString()
        COD_TIPO_COMPL3 = COD_TIPO_COMPL3.Replace("'", "#").ToString()
        COD_TIPO_COMPL3 = COD_TIPO_COMPL3.Replace("\", "#").ToString()
        TXT_TIPO_COMPL3 = TXT_TIPO_COMPL3.Replace("'", "#").ToString()
        TXT_TIPO_COMPL3 = TXT_TIPO_COMPL3.Replace("\", "#").ToString()
        COD_TIPO_COMPL4 = COD_TIPO_COMPL4.Replace("'", "#").ToString()
        COD_TIPO_COMPL4 = COD_TIPO_COMPL4.Replace("\", "#").ToString()
        TXT_TIPO_COMPL4 = TXT_TIPO_COMPL4.Replace("'", "#").ToString()
        TXT_TIPO_COMPL4 = TXT_TIPO_COMPL4.Replace("\", "#").ToString()
        TXT_COMPL_NUM = TXT_COMPL_NUM.Replace("'", "#").ToString()
        TXT_COMPL_NUM = TXT_COMPL_NUM.Replace("\", "#").ToString()
        NOM_COMPLETO = NOM_COMPLETO.Replace("'", "#").ToString()
        NOM_COMPLETO = NOM_COMPLETO.Replace("\", "#").ToString()

        Dim strComando As String = "INSERT INTO ora_cfg.cfg_endereco (" & _
        "COD_LOGRADOURO,    COD_CIDADE,         COD_TIPO_LOGR,      COD_OPERADORA,      COD_TITULO,     NOM_LOGRADOURO, " & _
        "NOM_LOGR_ABREV,    NOM_LOGR_COMPLETO,  COD_ENDERECO,       NUM_ENDERECO,       COD_TIPO_COMPL1, " & _
        "TXT_TIPO_COMPL1,   COD_TIPO_COMPL2,    TXT_TIPO_COMPL2,    COD_TIPO_COMPL3,    TXT_TIPO_COMPL3, " & _
        "COD_TIPO_COMPL4,   TXT_TIPO_COMPL4,    TXT_COMPL_NUM,      COD_IMOVEL,         NOM_COMPLETO,   COD_BAIRRO, DATA_UPDATE, DT_ATUALIZACAO_ORA ) " & _
        "VALUES ('" & COD_LOGRADOURO & "','" & COD_CIDADE & "','" & COD_TIPO_LOGR & _
        "','" & COD_OPERADORA & "','" & COD_TITULO & "','" & NOM_LOGRADOURO.Replace("'", "#").ToString & _
        "','" & NOM_LOGR_ABREV.Replace("'", "#").ToString & "','" & NOM_LOGR_COMPLETO.Replace("'", "#").ToString & _
        "','" & COD_ENDERECO.Replace("'", "#").ToString & "','" & NUM_ENDERECO.Replace("'", "#").ToString & _
        "','" & COD_TIPO_COMPL1.Replace("'", "#").ToString & "','" & TXT_TIPO_COMPL1.Replace("'", "#").ToString & _
        "','" & COD_TIPO_COMPL2.Replace("'", "#").ToString & "','" & TXT_TIPO_COMPL2.Replace("'", "#").ToString & _
        "','" & COD_TIPO_COMPL3.Replace("'", "#").ToString & "','" & TXT_TIPO_COMPL3.Replace("'", "#").ToString & _
        "','" & COD_TIPO_COMPL4.Replace("'", "#").ToString & "','" & TXT_TIPO_COMPL4.Replace("'", "#").ToString & _
        "','" & TXT_COMPL_NUM.Replace("'", "#").ToString & "','" & COD_IMOVEL & "','" & _
        NOM_COMPLETO.Replace("'", "#").ToString & "', '" & COD_BAIRRO & "', CURDATE(), '" & DT_ATUALIZACAO_ORA & "') " & _
        " ON DUPLICATE KEY UPDATE " & _
        "  cod_logradouro	 = '" & COD_LOGRADOURO & "', cod_operadora	= '" & COD_OPERADORA & "'" & _
        ", cod_tipo_logr	 = '" & COD_TIPO_LOGR & "', cod_titulo	= '" & COD_TITULO & "'" & _
        ", nom_logradouro	 = '" & NOM_LOGRADOURO.Replace("'", "#").ToString & "', nom_logr_abrev	    = '" & NOM_LOGR_ABREV.Replace("'", "#").ToString & "'" & _
        ", nom_logr_completo = '" & NOM_LOGR_COMPLETO.Replace("'", "#").ToString & "', num_endereco	= '" & NUM_ENDERECO.Replace("'", "#").ToString & "'" & _
        ", cod_tipo_compl1	 = '" & COD_TIPO_COMPL1.Replace("'", "#").ToString & "', txt_tipo_compl1	= '" & TXT_TIPO_COMPL1.Replace("'", "#").ToString & "'" & _
        ", cod_tipo_compl2	 = '" & COD_TIPO_COMPL2.Replace("'", "#").ToString & "', txt_tipo_compl2	= '" & TXT_TIPO_COMPL2.Replace("'", "#").ToString & "'" & _
        ", cod_tipo_compl3	 = '" & COD_TIPO_COMPL3.Replace("'", "#").ToString & "', txt_tipo_compl3	= '" & TXT_TIPO_COMPL3.Replace("'", "#").ToString & "'" & _
        ", cod_tipo_compl4	 = '" & COD_TIPO_COMPL4.Replace("'", "#").ToString & "', txt_tipo_compl4	= '" & TXT_TIPO_COMPL4.Replace("'", "#").ToString & "'" & _
        ", txt_compl_num	 = '" & TXT_COMPL_NUM.Replace("'", "#").ToString & "', nom_completo	= '" & NOM_COMPLETO.Replace("'", "#").ToString & "'" & _
        ", cod_bairro	= '" & COD_BAIRRO & "', data_update	= curdate(), DT_ATUALIZACAO_ORA = '" & DT_ATUALIZACAO_ORA & "'"

        MyComm = New MySqlCommand(strComando, MyConn)
        MyComm.ExecuteNonQuery()

    End Sub



End Class
