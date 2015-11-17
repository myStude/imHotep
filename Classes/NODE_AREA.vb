Imports MySql.Data.MySqlClient
Imports Oracle.DataAccess.Client

Public Class NODE_AREA
    Inherits objConnection

    Sub importaNodeArea(ByVal cidContrato As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)

        Dim tbStr As String = "ora_cfg.cfg_node_area (CID_CONTRATO,  AREA,  NODE, DT_UPDATE)"

        Dim FIELDS, TABLE, JOINS, WHERES, strSql, hvDtl As String

        FIELDS = " ID_AREA_DESPACHO, CID_CONTRATO, ID_CELULA "
        TABLE = " PROD_JD.SN_REL_AREA_CELULA_DESPACHO "
        JOINS = ""
        WHERES = " CID_CONTRATO  = " & cidContrato & " AND ID_TIPO_CELULA = 'CDP'"

        Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & JOINS & " WHERE " & WHERES

        Try
            oraComm = New OracleCommand(QryStr, oraConn)

            Using GerRow As OracleDataReader = oraComm.ExecuteReader()
                Dim udt As Integer = 0
                While GerRow.Read()
                    If udt = 0 Then udt = 1 : MysqlUpdate("ora_cfg.cfg_node_area", "DT_UPDATE = '1999-01-01 00:00:00'", "CID_CONTRATO IN (" & cidContrato & ")", MyConn, myComm)
                    strSql = "('" & GerRow.Item("CID_CONTRATO") & "','" & GerRow.Item("ID_AREA_DESPACHO") & "','" & GerRow.Item("ID_CELULA").ToString & "','" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "')"
                    hvDtl = "ON DUPLICATE KEY UPDATE DT_UPDATE = '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
                    MysqlAdd("", tbStr, strSql, hvDtl, myComm, MyConn)
                End While
                If udt = 1 Then MysqlDelete("ora_cfg.cfg_node_area", "DT_UPDATE = '1999-01-01 00:00:00'", MyConn, myComm)
            End Using
        Catch ex As Exception
            '# - - ERR - -
            'MsgBox(ex.Message)
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, "")
        End Try

    End Sub

End Class
