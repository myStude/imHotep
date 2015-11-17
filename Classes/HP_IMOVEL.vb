Imports Oracle.DataAccess.Client
Imports MySql.Data.MySqlClient

Public Class HP_IMOVEL

    Private wCtrl As New writerController

    ''' <summary>
    ''' FUNÇÃO QUE ATUALIZA TABELA 
    ''' </summary>
    ''' <remarks></remarks>
    Sub Atualiza(ByVal CodCidade As String, ByVal strBase As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal MyComm As MySqlCommand)

        Dim QryStr As String = StrQry(CodCidade)
        Try
            oraComm = New OracleCommand(QryStr, oraConn)
            Using GerRow As OracleDataReader = oraComm.ExecuteReader()
                While GerRow.Read()
                    insereHPImovel( _
                    GerRow.Item("COD_OPERADORA").ToString, GerRow.Item("COD_HP").ToString, _
                    GerRow.Item("COD_TIPO_IMOVEL").ToString, GerRow.Item("COD_IMOVEL").ToString, _
                    GerRow.Item("COD_TIPO_COMPL1").ToString, GerRow.Item("TXT_COMPL1").ToString, _
                    GerRow.Item("COD_TIPO_COMPL2").ToString, GerRow.Item("TXT_COMPL2").ToString, _
                    GerRow.Item("COD_HP_SMS").ToString, GerRow.Item("DT_ATUALIZACAO_ORA").ToString, MyConn, MyComm)
                End While
            End Using

        Catch ex As Exception
            '// ESCREVE RELATORIO CASO DE ERRO
            wCtrl.write(Format(Now, "dd/MM/yyyy HH:mm"), System.Reflection.MethodBase.GetCurrentMethod.Name & ": " & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' query de extração
    ''' </summary>
    ''' <returns>query</returns>
    ''' <remarks>enjoy</remarks>
    Private Function StrQry(ByVal CodCidade As String) As String

        Dim FIELDS, TABLE, JOINS, WHERES As String

        FIELDS = " COD_OPERADORA, COD_HP, COD_TIPO_IMOVEL, COD_IMOVEL, COD_TIPO_COMPL1, " _
                & " TXT_COMPL1, COD_TIPO_COMPL2, TXT_COMPL2, COD_HP_SMS , " _
                & " to_char(GED.HP_IMOVEL.DT_ATUALIZACAO,'YYYY-MM-DD HH24:MI:SS') DT_ATUALIZACAO_ORA "
        TABLE = " GED.HP_IMOVEL "
        JOINS = ""
        WHERES = " GED.HP_IMOVEL.COD_OPERADORA IN (" & CodCidade & ") AND GED.HP_IMOVEL.DT_ATUALIZACAO >= TO_DATE('" & CDate(Format(Now, "dd-MM-yyyy").ToString).AddDays(-7) & " 00:00:00','DD/MM/YYYY HH24:MI:SS') "

        Return "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES

    End Function

    ''' <summary>
    ''' FUNÇÃO INSERE NO MYSQL
    ''' </summary>
    ''' <param name="COD_OPERADORA">INT 3</param>
    ''' <param name="COD_HP">BIGINT 10</param>
    ''' <param name="COD_TIPO_IMOVEL">INT</param>
    ''' <param name="COD_IMOVEL">INT</param>
    ''' <param name="COD_TIPO_COMPL1">STRING</param>
    ''' <param name="TXT_COMPL1">STRING</param>
    ''' <param name="COD_TIPO_COMPL2">STRING</param>
    ''' <param name="TXT_COMPL2">STRING</param>
    ''' <param name="COD_HP_SMS">INT</param>
    ''' <param name="DT_ATUALIZACAO_ORA">DATE</param>
    ''' <param name="strConn">MySQL CONNECTION</param>
    ''' <param name="MyCom">MySQL COMMAND</param>
    ''' <returns>row affected or error</returns>
    ''' <remarks>enjoy</remarks>
    Private Function insereHPImovel(ByVal COD_OPERADORA As String, ByVal COD_HP As String, ByVal COD_TIPO_IMOVEL As String, ByVal COD_IMOVEL As String, ByVal COD_TIPO_COMPL1 As String, ByVal TXT_COMPL1 As String, ByVal COD_TIPO_COMPL2 As String, ByVal TXT_COMPL2 As String, ByVal COD_HP_SMS As String, ByVal DT_ATUALIZACAO_ORA As String, ByVal strConn As MySqlConnection, ByVal MyCom As MySqlCommand)
        Dim Affected As String
        Dim strComando = _
        "INSERT INTO ora_cfg.cfg_hp_imovel " & _
        "( COD_OPERADORA, COD_HP,COD_TIPO_IMOVEL,COD_IMOVEL,COD_TIPO_COMPL1, TXT_COMPL1,COD_TIPO_COMPL2,TXT_COMPL2,COD_HP_SMS, DT_UPDATE ) " _
        & "VALUES ('" & COD_OPERADORA & "','" & COD_HP & "','" & COD_TIPO_IMOVEL & "','" & _
        COD_IMOVEL & "','" & COD_TIPO_COMPL1.Replace("'", "#") & "','" & TXT_COMPL1.Replace("'", "#") & "','" & _
        COD_TIPO_COMPL2.Replace("'", "#") & "','" & TXT_COMPL2.Replace("'", "#") & "','" & COD_HP_SMS & "', '" & _
        Format(Now, "yyyy-MM-dd HH:mm:ss") & "') " & _
        " ON DUPLICATE KEY UPDATE " & _
        " COD_IMOVEL = '" & COD_IMOVEL & "', " & _
        " COD_TIPO_IMOVEL = '" & COD_TIPO_IMOVEL & "', " & _
        " COD_TIPO_COMPL1 = '" & COD_TIPO_COMPL1.Replace("'", "#") & "' , " & _
        " TXT_COMPL1 = '" & TXT_COMPL1.Replace("'", "#") & "', " & _
        " COD_TIPO_COMPL2 = '" & COD_TIPO_COMPL2.Replace("'", "#") & "', " & _
        " TXT_COMPL2 = '" & TXT_COMPL2.Replace("'", "#") & "', " & _
        " COD_HP_SMS = '" & COD_HP_SMS & "', " & _
        " DT_ATUALIZACAO_ORA = '" & DT_ATUALIZACAO_ORA & "', " & _
        " DT_UPDATE = '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"

        strComando = Replace(strComando, "\", " # ")

        MyCom = New MySqlCommand(strComando, strConn)
        Try
            Affected = MyCom.ExecuteNonQuery()
        Catch ex As Exception
            Return ex.Message
            Exit Function
        End Try

        Return Affected

    End Function

End Class
