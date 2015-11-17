Imports System
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
''' <summary>
''' CLASSE ENDER
''' </summary>
''' <remarks>SE NAO TIVER CONEXAO, FALHA</remarks>
Public Class ENDERECOS
    Inherits objConnection

    ''' <summary>
    ''' ATUALIZA CODIFICAÇÃO DE ENDEREÇOS FROM ORA
    ''' </summary>
    ''' <param name="idEnder"></param>
    ''' <param name="codOperadora"></param>
    ''' <remarks></remarks>
    Sub importaEnderecos(ByVal idEnder As String, ByVal codOperadora As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)

        Dim RETORNO As String = ""
        Dim ERRO As Integer = 0
        Dim ID_ENDER, COD_HP, COD_OPERADORA, COD_IMOVEL, DT_ATUALIZACAO_ORA As String
        Dim FIELDS, TABLE, JOINS, WHERES As String
        Dim tbStr As String = "ora_cfg.cfg_ender(ID_ENDER, ID_EDIFICACAO_HP, COD_OPERADORA, COD_IMOVEL, DATA_UPDATE, DT_ATUALIZACAO_ORA)"
        Dim strSql, strDtl As String

        '//SQLSTR TO ORA
        FIELDS = " EN.ID_ENDER, EN.ID_EDIFICACAO AS COD_HP, IM.COD_OPERADORA, IM.COD_IMOVEL, to_char(IM.DT_ATUALIZACAO,'YYYY-MM-DD HH24:MI:SS') DT_ATUALIZACAO_ORA "
        TABLE = " PROD_JD.SN_ENDER EN "
        JOINS = " INNER JOIN GED.HP_IMOVEL IM on IM.COD_HP = EN.ID_EDIFICACAO "
        '// SE NAO TIVER ID_ENDER, BUSCA POR DATA
        If idEnder = "" Then _
            WHERES = " IM.COD_OPERADORA in (" & codOperadora & ") " & _
                     " AND IM.DT_ATUALIZACAO BETWEEN TO_DATE('" & CDate(Format(Now, "dd-MM-yyyy").ToString).AddDays(-30) & " 00','DD/MM/YYYY HH24') " & _
                     " AND TO_DATE('" & Format(Now, "dd/MM/yyyy").ToString & " " & Format(Now, "HH").ToString & "','DD/MM/YYYY HH24') " _
                            Else WHERES = " IM.COD_OPERADORA = " & codOperadora & "  AND EN.ID_ENDER = " & idEnder
        '// SQLSTR MOUNT
        Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & JOINS & " WHERE " & WHERES

        Try

            '--------------------------
            oraComm = New OracleCommand(QryStr, oraConn)

            '// READ ENGINE
            Using GerRow As OracleDataReader = oraComm.ExecuteReader()
                'ENQUANTO HÁ LEITURA NO GerRow INSERE NO MySQL
                While GerRow.Read()
                    If GerRow.Item("ID_ENDER") Is DBNull.Value Then ID_ENDER = "" Else ID_ENDER = GerRow.Item("ID_ENDER")
                    If GerRow.Item("COD_HP") Is DBNull.Value Then COD_HP = "" Else COD_HP = GerRow.Item("COD_HP")
                    If GerRow.Item("COD_OPERADORA") Is DBNull.Value Then COD_OPERADORA = "" Else COD_OPERADORA = GerRow.Item("COD_OPERADORA")
                    If GerRow.Item("COD_IMOVEL") Is DBNull.Value Then COD_IMOVEL = "" Else COD_IMOVEL = GerRow.Item("COD_IMOVEL")
                    If GerRow.Item("DT_ATUALIZACAO_ORA") Is DBNull.Value Then DT_ATUALIZACAO_ORA = "0000-00-00 00:00:00" Else DT_ATUALIZACAO_ORA = GerRow.Item("DT_ATUALIZACAO_ORA")

                    strSql = "('" & ID_ENDER & "','" & COD_HP & "','" & COD_OPERADORA & "','" & COD_IMOVEL & "', CURDATE(), '" & DT_ATUALIZACAO_ORA & "') "

                    strDtl = " ON DUPLICATE KEY UPDATE " _
                            & " ID_EDIFICACAO_HP = '" & COD_HP & "', COD_IMOVEL = '" & COD_IMOVEL & "', " _
                            & " DATA_UPDATE = CURDATE(), DT_ATUALIZACAO_ORA = '" & DT_ATUALIZACAO_ORA & "'"

                    MysqlAdd("", tbStr, strSql, strDtl, myComm, MyConn)

                    'ATUALIZA_CFG_ENDER(ID_ENDER, COD_HP, COD_OPERADORA, COD_IMOVEL, DT_ATUALIZACAO_ORA, MyConn)

                End While
            End Using
        Catch ex As Exception
            '# -- ERR --
            'MsgBox(ex.Message)
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, "")
        End Try

    End Sub


End Class
