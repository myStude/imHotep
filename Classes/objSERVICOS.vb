Imports System.IO
Imports System.Data
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Imports MySql.Data.MySqlClient


Public Class objSERVICOS
    Inherits objConnection

    Private mCtrl As New mailController


    ''' <summary>
    ''' IMPORTA ORA GERADAS
    ''' </summary>
    ''' <param name="CIDADE"></param>
    ''' <param name="MAXSOLIC"></param>
    ''' <param name="DATAORA"></param>
    ''' <param name="DATAMY"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function importaGeradas(ByVal CIDADE As String, ByVal MAXSOLIC As String, ByVal DATAORA As String, ByVal DATAMY As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal MyComm As MySqlCommand)

        '/*
        ' *  FALTA RELATORIO DE ERROS
        '*/

        Dim tentativas As Integer = 0
        Dim Data As String = Format(CDate(DATAORA), "dd/MM/yyyy").ToString
        Dim FIELDS, TABLE, JOINS, WHERES, OBS As String

        Dim RETORNO As String = ""
        Dim EStr As String = "" : Dim BsodStr As String = ""
        Dim TORA As String = ""
        Dim EQPN, BSODN As Integer : EQPN = 0 : BSODN = 0
        Dim QTDEINSERT, QTDEUPDATE, VARREGISTRO As Double

        Dim errora, errmy As Integer : errora = 0 : errmy = 0
        Dim corpo As String
        Dim HORAATEND, PACOTEPONTO, TIPOPONTO, DTATEND, OS, DT, CID, TIPOOS, CONTRATO, STATUS, PERIODO, IMEDIATA, IDASSINANTE, IDSOLIC, IDENDER, TIPOCLIENTE, IDPERIODO, IDTIPOFECH, IDPONTO, IDEQUIPE, IDOCORRENCIA, CONV As String
        Dim ERRO As Integer = 0
        Dim ARRAYPONTO(2) As String
        Dim HORAINICIO As String
        Dim DIFHORA As TimeSpan

        HORAINICIO = Now
        errmy = 0
        QTDEINSERT = 0
        QTDEUPDATE = 0
        VARREGISTRO = 0

        Dim InsertStr As String : Dim MltQry As String = "" : Dim X As Integer = 0 : Dim Rol As Double = 1 ': Dim Blocos As Integer = Int(TORA / 100) + 1


        FIELDS = "OS.OBS," _
             & "(SELECT CA.id_caracteristica || '#' || P.DESCRICAO FROM PROD_JD.SN_REL_PONTO_PRODUTO PP " _
             & "INNER JOIN PROD_JD.SN_PRODUTO P ON P.ID_PRODUTO = PP.ID_PRODUTO " _
             & "INNER JOIN PROD_JD.SN_CARACTERISTICA CA ON CA.ID_CARACTERISTICA = P.ID_CARACTERISTICA " _
             & "WHERE ID_PONTO = OS.ID_PONTO AND INSTALADO = 1 " _
             & "AND DT_FIM > sysdate AND ID_TIPO_PRODUTO = 1 AND ROWNUM = 1) as TIPO_PONTO, " _
             & "OS.COD_OS,to_char(OS.DT_ATEND,'YYYY-MM-DD HH24:MI:SS') DT_ATEND,PE.DT,CT.CID_CONTRATO,OS.ID_TIPO_OS, " _
             & "CT.NUM_CONTRATO,OS.STATUS,PE.ID_TIPO_PERIODO, OS.IMEDIATA,OS.ID_ASSINANTE,OS.ID_SOLICITACAO_ASS,OS.ID_ENDER, " _
             & "OS.ID_TIPO_CLIENTE, OS.ID_PERIODO,OS.ID_TIPO_FECHAMENTO,OS.ID_PONTO, " _
             & "OS.ID_EQUIPE,OS.ID_OCORRENCIA, HIS.fn_conveniencia AS CONV "

        TABLE = " PROD_JD.SN_OS OS "

        JOINS = "INNER JOIN PROD_JD.SN_CONTRATO CT ON OS.ID_ASSINANTE = CT.ID_ASSINANTE " _
              & "INNER JOIN PROD_JD.SN_PERIODO  PE ON OS.ID_PERIODO   = PE.ID_PERIODO " _
              & "INNER JOIN PROD_JD.TBSN_HISTORICO_AGENDAMENTO_OS HIS ON OS.COD_OS = HIS.COD_OS "

        If MAXSOLIC = "" Then
            WHERES = "OS.DT_ATEND BETWEEN TO_DATE('" & DATAORA & " 00:00:00','DD/MM/YYYY HH24:MI:SS') " _
                   & "               AND TO_DATE('" & DATAORA & " 23:59:59','DD/MM/YYYY HH24:MI:SS') " _
                   & "AND CT.CID_CONTRATO = '" & CIDADE & "'"
        Else
            WHERES = "OS.ID_SOLICITACAO_ASS > " & MAXSOLIC _
                   & " AND CT.CID_CONTRATO = '" & CIDADE & "'"
        End If


        Try

            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand
            Dim CmdStr As New OracleCommand
            Dim CmdStr2 As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text
            CmdStr.CommandType = CommandType.Text
            CmdStr2.CommandType = CommandType.Text

            '#  LE QUERY
            Using GerRow As OracleDataReader = com.ExecuteReader()

                '#  ENQUANTO LE...
                While GerRow.Read()
                    VARREGISTRO = VARREGISTRO + 1
                    QTDEINSERT = QTDEINSERT + 1

                    OS = GerRow.Item("COD_OS")
                    If GerRow.Item("OBS") Is DBNull.Value Then OBS = "-" Else OBS = GerRow.Item("OBS")
                    If GerRow.Item("DT") Is DBNull.Value Then DT = "0000-00-00" Else DT = Format(CDate(GerRow.Item("DT")), "yyyy-MM-dd")
                    DTATEND = Format(CDate(GerRow.Item("DT_ATEND")), "yyyy-MM-dd")
                    HORAATEND = Format(CDate(GerRow.Item("DT_ATEND")), "HH:mm:ss")
                    CID = GerRow.Item("CID_CONTRATO")
                    TIPOOS = GerRow.Item("ID_TIPO_OS")
                    CONTRATO = GerRow.Item("NUM_CONTRATO")
                    STATUS = GerRow.Item("STATUS")
                    If GerRow.Item("ID_TIPO_PERIODO") Is DBNull.Value Then PERIODO = 0 Else PERIODO = GerRow.Item("ID_TIPO_PERIODO").ToString
                    If GerRow.Item("IMEDIATA") Is DBNull.Value Then IMEDIATA = 0 Else IMEDIATA = 1
                    IDASSINANTE = GerRow.Item("ID_ASSINANTE")
                    IDSOLIC = GerRow.Item("ID_SOLICITACAO_ASS")
                    IDENDER = GerRow.Item("ID_ENDER")
                    If GerRow.Item("ID_TIPO_CLIENTE") Is DBNull.Value Then TIPOCLIENTE = 0 Else TIPOCLIENTE = GerRow.Item("ID_TIPO_CLIENTE").ToString
                    If GerRow.Item("ID_PERIODO") Is DBNull.Value Then IDPERIODO = 0 Else IDPERIODO = GerRow.Item("ID_PERIODO").ToString
                    If GerRow.Item("ID_TIPO_FECHAMENTO") Is DBNull.Value Then IDTIPOFECH = "0" Else IDTIPOFECH = GerRow.Item("ID_TIPO_FECHAMENTO")
                    If GerRow.Item("ID_PONTO") Is DBNull.Value Then IDPONTO = "0" Else IDPONTO = GerRow.Item("ID_PONTO")
                    If GerRow.Item("ID_EQUIPE") Is DBNull.Value Then IDEQUIPE = "0" Else IDEQUIPE = GerRow.Item("ID_EQUIPE")
                    If GerRow.Item("ID_OCORRENCIA") Is DBNull.Value Then IDOCORRENCIA = "0" Else IDOCORRENCIA = GerRow.Item("ID_OCORRENCIA")
                    If GerRow.Item("CONV") Is DBNull.Value Then CONV = "0" Else CONV = GerRow.Item("CONV")

                    If GerRow.Item("TIPO_PONTO") Is DBNull.Value Then
                        TIPOPONTO = "-"
                    Else
                        ARRAYPONTO = Split(GerRow.Item("TIPO_PONTO"), "#")
                        PACOTEPONTO = ARRAYPONTO(1)
                        TIPOPONTO = ARRAYPONTO(0)
                        If TIPOPONTO = "2" Then
                            If InStr(ARRAYPONTO(1), "HD") > 0 Then
                                TIPOPONTO = "5"
                            Else
                                TIPOPONTO = "2"
                            End If
                        End If
                    End If

                    'INSERE MYSQL -------------------------------------------------------------------------
                    InsertStr = "ora_ord.ord_geradas " _
                        & "(COD_OS,DT_ATEND,HORA_ATEND,DT_AGENDA,CID_CONTRATO,ID_TIPO_OS,NUM_CONTRATO, " _
                        & "STATUS,ID_TIPO_PERIODO,IMEDIATA,ID_ASSINANTE,ID_SOLICITACAO_ASS,ID_ENDER,TIPO_CLIENTE, " _
                        & "ID_PERIODO,ID_TIPO_FECHAMENTO,ID_PONTO,ID_EQUIPE,ID_OCORRENCIA,DT_UPDATE, TIPO_PONTO, CONVENIENCIA) "

                    MltQry = "(" & OS & ", '" & DTATEND & "', '" & HORAATEND & "', '" & DT & "', " & CID & ", " & TIPOOS & ", " & CONTRATO & ", " & _
                            "'" & STATUS & "', " & PERIODO & ", " & IMEDIATA & ", " & IDASSINANTE & ", " & IDSOLIC & ", " & IDENDER & ",  " & _
                                TIPOCLIENTE & ", " & IDPERIODO & ", " & IDTIPOFECH & ", " & IDPONTO & ", " & IDEQUIPE & ", " & IDOCORRENCIA & ", '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', '" & _
                                    TIPOPONTO & "', " & CONV & ")"
                    Dim OnDuplcate As String = ""
                    MysqlAdd("IGNORE", InsertStr, MltQry, "", MyComm, MyConn)
                    'strMY.INSERIR_NO_MYSQL("IGNORE", InsertStr, MltQry, OnDuplcate, MyConn) '<- DEPRECATED
                    '--------------------------------------------------------------------------------------

                    '#################################################################################################

                    ' EQUIPAMENTOS -----------------------------------------------------------------------------
                    If IDPONTO <> "0" Then

                        EStr = "select * from (SELECT EN.CD_ENDERECAVEL,T.NR_SERIE,T.ID_MODELO_EQUIPAMENTO, " _
                            & "PC.ID_PONTO,PH.TM_ID_ASSOC,PH.TM_ID,PC.NUM_CONTRATO,PC.CID_CONTRATO  " _
                            & "FROM PROD_JD.SN_PONTO_CONTR PC   " _
                            & "INNER JOIN PROD_JD.SN_PONTO_HISTORICO PH ON PH.ID_PONTO = PC.ID_PONTO  " _
                            & "INNER JOIN INTEGRACAO_ATLAS.MW_EQUIPAMENTO T ON T.ID_EQUIPAMENTO = PH.TM_ID  " _
                            & "LEFT JOIN INTEGRACAO_ATLAS.MW_ENDERECAVEL EN ON EN.id_equipamento = PH.TM_ID " _
                            & "WHERE PC.CID_CONTRATO = " & CID & " AND PH.DT_FIM > SYSDATE  and PH.INSTALADO = 1" _
                            & "UNION ALL " _
                            & "SELECT EN.CD_ENDERECAVEL,T.NR_SERIE,T.ID_MODELO_EQUIPAMENTO, " _
                            & "PC.ID_PONTO,PH.TM_ID_ASSOC,PH.TM_ID,PC.NUM_CONTRATO,PC.CID_CONTRATO " _
                            & "FROM PROD_JD.SN_PONTO_CONTR PC  " _
                            & "INNER JOIN PROD_JD.SN_PONTO_HISTORICO PH ON PH.ID_PONTO = PC.ID_PONTO  " _
                            & "INNER JOIN INTEGRACAO_ATLAS.MW_EQUIPAMENTO T ON T.ID_EQUIPAMENTO = PH.TM_ID_ASSOC " _
                            & "LEFT JOIN INTEGRACAO_ATLAS.MW_ENDERECAVEL EN ON EN.id_equipamento = PH.TM_ID_ASSOC " _
                            & "WHERE PC.CID_CONTRATO = " & CID & " AND PH.DT_FIM > SYSDATE and PH.INSTALADO = 1) WHERE ID_PONTO = " & IDPONTO
                        CmdStr = New OracleCommand(EStr, oraConn)
                        CmdStr.CommandType = CommandType.Text

                        Dim strComando As String
                        Dim eqTable As String = "ora_ord.ord_equipamentos (CD_ENDERECAVEL,NR_SERIE,ID_MODELO_EQUIPAMENTO,ID_PONTO,TM_ID_ASSOC,TM_ID,NUM_CONTRATO,CID_CONTRATO,DT_UPDATE,ID_SOLIC)"

                        '#  RODA CONSULTA EQUIPAMENTO
                        Using e As OracleDataReader = CmdStr.ExecuteReader()
                            EQPN = 0
                            While e.Read()
                                EQPN = EQPN + 1
                                strComando = "('" & e.Item("CD_ENDERECAVEL") & "','" & _
                                            e.Item("NR_SERIE") & "','" & _
                                            e.Item("ID_MODELO_EQUIPAMENTO") & "','" & _
                                            e.Item("ID_PONTO") & "','" & _
                                            e.Item("TM_ID_ASSOC") & "','" & _
                                            e.Item("TM_ID") & "','" & _
                                            e.Item("NUM_CONTRATO") & "','" & _
                                            e.Item("CID_CONTRATO") & "','" & _
                                            Format(Now, "yyyy-MM-dd") & "','" & _
                                            IDSOLIC & "')"

                                MysqlAdd("IGNORE", eqTable, strComando, "", MyComm, MyConn)

                            End While
                            e.Dispose()
                            e.Close()
                        End Using

                    End If

                    '#################################################################################################

                    ' VERIFICA BSOD ---------------------------------------------------------------------
                    If TIPOOS = "22" Or TIPOOS = "26" Or TIPOOS = "38" Or TIPOOS = "27" Or TIPOOS = "10" Or TIPOOS = "48" Or TIPOOS = "50" Or TIPOOS = "62" Then

                        FIELDS = " RP.NUM_CONTRATO, PR.DESCRICAO "
                        TABLE = " PROD_JD.SN_REL_PONTO_PRODUTO RP "
                        JOINS = " LEFT JOIN PROD_JD.SN_PRODUTO                   PR ON  RP.ID_PRODUTO = PR.ID_PRODUTO " _
                        & " LEFT JOIN PROD_JD.VPP_ASS                      VA ON  RP.NUM_CONTRATO = VA.NUM_CONTRATO AND RP.CID_CONTRATO = VA.CID_CONTRATO " _
                        & " LEFT JOIN PROD_JD.SN_REL_ASSINANTE_SEGMENTACAO RA ON  RP.NUM_CONTRATO = RA.NUM_CONTRATO AND RP.CID_CONTRATO = RA.CID_CONTRATO "
                        WHERES = " RA.ID_TIPO_SEGMENTO = 11 " _
                            & " AND RP.CID_CONTRATO = '" & CID & "' " _
                            & " AND RP.NUM_CONTRATO = '" & CONTRATO & "' and dt_fim >= SYSDATE "

                        Dim Str2 As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
                        CmdStr2 = New OracleCommand(Str2, oraConn)
                        CmdStr2.CommandType = CommandType.Text

                        corpo = "Cidade: " & CID & "<TABLE BORDER=1 BGCOLOR='#B0C4DE'><TR><FONT SIZE='2'><TH>CONTRATO</TH><TH>OS</TH><TH>DT AGENDA</TH></FONT></TR>"
                        Using b As OracleDataReader = CmdStr2.ExecuteReader()
                            BSODN = 0
                            While b.Read()
                                BSODN = BSODN + 1
                                corpo += "<tr><FONT SIZE='2'><td>" & b.Item("NUM_CONTRATO") & "</td><td>" & OS & "</td><td>" & b.Item("DESCRICAO") & "</td><td>" & DT & "</td></FONT></tr>"
                            End While
                            If BSODN > 0 Then
                                corpo += "</TABLE>"

                                '# SEND EMAIL BSOD
                                mCtrl.sent("fagner.silva@net.com.br", "BSOD", corpo)
                            End If
                            b.Dispose()
                            b.Close()
                        End Using
                    End If
                    '-------------------------------------------------------------------

                    '#################################################################################################

                    ' OBSERVAÇÃO -----------------------------------------------------

                    Dim obsTable As String = "ora_ord.ord_observacao (COD_OS,NUM_CONTRATO,OBS)"

                    Dim values As String = "( '" & OS & "','" & CONTRATO & "','" & OBS & "')"
                    Dim hvDtl As String = " ON DUPLICATE KEY UPDATE OBS = '" & OBS & "'"

                    MysqlAdd("", obsTable, values, hvDtl, MyComm, MyConn)

                End While
            End Using


            DIFHORA = Now - CDate(HORAINICIO)

DISPOSE:

        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & CIDADE & ") - " & ex.Message
            Exit Function
        End Try

        Return Now & " - " & System.Reflection.MethodBase.GetCurrentMethod.Name & RETORNO & " / CIDADE " & CIDADE

    End Function

    ''' <summary>
    ''' IMPORTA ORA PENDENTES
    ''' </summary>
    ''' <param name="cfgcidade"></param>
    ''' <param name="DATAORA"></param>
    ''' <param name="DATAMY"></param>
    ''' <param name="BSOD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function importaPendentes(ByVal cfgcidade As String, ByVal DATAORA As String, ByVal DATAMY As String, ByVal BSOD As Integer, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand) As String

        Dim RETORNO As String = ""
        Dim TORA As Integer = 0
        Dim QTDEINSERT, QTDEUPDATE, VARREGISTRO As Double
        Dim CONV As String
        Dim FIELDS, TABLE, JOINS, WHERES As String

        Dim CntBSOD As Integer
        Dim DIMPORTA As New DataTable("DIMPORTA")

        Dim TABLETEMP As New DataTable("TABLETEMP")
        Dim ROWTEMP As DataRow
        Dim HORAATEND, DTATEND, OS, DT, CID, TIPOOS, CONTRATO, STATUS, PERIODO, IMEDIATA, IDASSINANTE, IDSOLIC, IDENDER, TIPOCLIENTE, IDPERIODO, IDTIPOFECH, IDPONTO, IDEQUIPE, IDOCORRENCIA As String
        Dim ERRO As Integer = 0
        Dim HORAINICIO, StrBSOD, ERROBSOD, MltQry As String : StrBSOD = "" : ERROBSOD = ""
        Dim CORPO As String = ""
        Dim DIFHORA As TimeSpan = Nothing
        HORAINICIO = Now

        QTDEINSERT = 0
        QTDEUPDATE = 0

        Dim DATA, DATA2 As String
        DATA = Format(CDate(DATAORA), "dd/MM/yyyy").ToString
        DATA2 = Format(CDate(DATAORA).AddYears(1), "dd/MM/yyyy").ToString

        '   NOVA QUERY
        'Dim strComando As String

        FIELDS = " OS.COD_OS,PE.DT,CT.CID_CONTRATO,OS.ID_TIPO_OS,CT.NUM_CONTRATO,OS.STATUS, " & _
              " PE.ID_TIPO_PERIODO,OS.IMEDIATA,OS.ID_ASSINANTE,OS.ID_SOLICITACAO_ASS,OS.ID_ENDER,OS.ID_TIPO_CLIENTE,  " & _
              " OS.ID_PERIODO,OS.ID_TIPO_FECHAMENTO,OS.ID_PONTO,OS.ID_EQUIPE,OS.ID_OCORRENCIA, HIS.fn_conveniencia AS CONV, " & _
              " to_char(OS.DT_ATEND,'YYYY-MM-DD HH24:MI:SS') DT_ATEND  "

        TABLE = "PROD_JD.SN_OS OS "

        JOINS = "INNER JOIN PROD_JD.SN_CONTRATO CT ON CT.ID_ASSINANTE = OS.ID_ASSINANTE " _
              & "INNER JOIN PROD_JD.SN_PERIODO  PE ON PE.ID_PERIODO   = OS.ID_PERIODO " _
              & "LEFT JOIN PROD_JD.TBSN_HISTORICO_AGENDAMENTO_OS HIS ON OS.COD_OS = HIS.COD_OS "

        WHERES = " PE.DT BETWEEN TO_DATE('" & DATA & "','DD/MM/YYYY') AND TO_DATE('" & DATA2 & "','DD/MM/YYYY') " _
               & "AND OS.STATUS IN ('A','D','C') AND CT.CID_CONTRATO = '" & cfgcidade & "'"


        '#########################################################################################
        '#  USING CLIENT ORA <------------<<
        Try
            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand
            Dim CmdStr2 As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text



            'ATUALIZA MYSQL
            'MUDA STATUS PARA '-' PARA EXCLUIR DEPOIS DA ATUALIZALÇAO --------------------------
            MysqlUpdate("ora_ord.ord_pendentes", "STATUS = '-'", "CID_CONTRATO IN ('" & cfgcidade & "')", MyConn, myComm)

            'For Each PenRows.Item In IMPORTA.Rows
            Using PenRows As OracleDataReader = com.ExecuteReader()
                While PenRows.Read()
                    TIPOOS = PenRows.Item("ID_TIPO_OS")
                    CONTRATO = PenRows.Item("NUM_CONTRATO")
                    VARREGISTRO = VARREGISTRO + 1
                    QTDEINSERT = QTDEINSERT + 1
                    OS = PenRows.Item("COD_OS")

                    DT = PenRows.Item("DT")
                    CID = PenRows.Item("CID_CONTRATO")
                    STATUS = PenRows.Item("STATUS")
                    PERIODO = PenRows.Item("ID_TIPO_PERIODO")

                    DTATEND = Format(CDate(PenRows.Item("DT_ATEND")), "yyyy-MM-dd")
                    HORAATEND = Format(CDate(PenRows.Item("DT_ATEND")), "HH:mm:ss")

                    If PenRows.Item("IMEDIATA") Is DBNull.Value Then IMEDIATA = 0 Else IMEDIATA = 1
                    IDASSINANTE = PenRows.Item("ID_ASSINANTE")
                    IDSOLIC = PenRows.Item("ID_SOLICITACAO_ASS")
                    IDENDER = PenRows.Item("ID_ENDER")
                    TIPOCLIENTE = PenRows.Item("ID_TIPO_CLIENTE")
                    IDPERIODO = PenRows.Item("ID_PERIODO")
                    If PenRows.Item("ID_TIPO_FECHAMENTO") Is DBNull.Value Then IDTIPOFECH = "0" Else IDTIPOFECH = PenRows.Item("ID_TIPO_FECHAMENTO")
                    If PenRows.Item("ID_PONTO") Is DBNull.Value Then IDPONTO = "0" Else IDPONTO = PenRows.Item("ID_PONTO")
                    If PenRows.Item("ID_EQUIPE") Is DBNull.Value Then IDEQUIPE = "0" Else IDEQUIPE = PenRows.Item("ID_EQUIPE")
                    If PenRows.Item("ID_OCORRENCIA") Is DBNull.Value Then IDOCORRENCIA = "0" Else IDOCORRENCIA = PenRows.Item("ID_OCORRENCIA")
                    If PenRows.Item("CONV") Is DBNull.Value Then CONV = "0" Else CONV = PenRows.Item("CONV")

                    ' B S O D ------------------------------------
                    If TIPOOS = "43" Or TIPOOS = "44" Or TIPOOS = "71" Then
                        If BSOD = 1 Then

                            FIELDS = " PROD_JD.SN_REL_PONTO_PRODUTO.NUM_CONTRATO, PROD_JD.SN_PRODUTO.DESCRICAO "
                            TABLE = " PROD_JD.SN_REL_PONTO_PRODUTO "
                            JOINS = " INNER JOIN PROD_JD.SN_PRODUTO ON PROD_JD.SN_REL_PONTO_PRODUTO.ID_PRODUTO = PROD_JD.SN_PRODUTO.ID_PRODUTO "
                            WHERES = " PROD_JD.SN_PRODUTO.DESCRICAO LIKE '%BSOD%' " _
                            & " AND PROD_JD.SN_REL_PONTO_PRODUTO.CID_CONTRATO = '" & cfgcidade & "' " _
                            & " AND NUM_CONTRATO = '" & CONTRATO & "' and dt_fim >= SYSDATE"

                            Dim Str2 As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
                            CmdStr2 = New OracleCommand(Str2, oraConn)
                            'CmdStr2.CommandText = Str2
                            CmdStr2.CommandType = CommandType.Text
                            CORPO = "Cidade: " & CID & "<TABLE BORDER=1 BGCOLOR='#B0C4DE'><TR><FONT SIZE='2'><TH>CONTRATO</TH><TH>OS</TH><TH>DT AGENDA</TH></FONT></TR>"
                            Using b As OracleDataReader = CmdStr2.ExecuteReader()
                                CntBSOD = 0
                                While b.Read()
                                    For Each ROWTEMP In TABLETEMP.Rows
                                        CntBSOD += 1
                                        CORPO += "<tr><FONT SIZE='2'><td>" & b.Item("NUM_CONTRATO") & "</td><td>" & OS & "</td><td>" & DT & "</td></FONT></tr>"
                                    Next
                                    CORPO += "</TABLE>"
                                End While
                                If CntBSOD > 0 Then
                                    mCtrl.sent("fagner.silva@net.com.br", "BSOD", CORPO)
                                End If
                            End Using
                        End If
                    End If
                    '------------------------------------------------------------------------------------------

                    '#  MYSQL QUERY RUN -----------------------------------------------------------------------
                    Dim OnDuplcate As String = " ON DUPLICATE KEY UPDATE DT_AGENDA = values(DT_AGENDA), " & _
                            "STATUS             = values(STATUS), ID_TIPO_PERIODO  = values(ID_TIPO_PERIODO), " & _
                            "DT_ATEND           = values(DT_ATEND), HORA_ATEND     = values(HORA_ATEND), " & _
                            "IMEDIATA           = values(IMEDIATA), ID_SOLICITACAO_ASS = values(ID_SOLICITACAO_ASS), " & _
                            "TIPO_CLIENTE       = values(TIPO_CLIENTE), ID_PERIODO     = values(ID_PERIODO), " & _
                            "ID_TIPO_FECHAMENTO = values(ID_TIPO_FECHAMENTO), ID_PONTO = values(ID_PONTO), " & _
                            "ID_EQUIPE          = values(ID_EQUIPE), ID_OCORRENCIA     = values(ID_OCORRENCIA) "

                    MltQry = "(" & OS & ", '" & Format(CDate(DT), "yyyy-MM-dd") & "', " & CID & ", " & _
                        TIPOOS & ", " & CONTRATO & ", '" & STATUS & "', " & PERIODO & ", " & _
                        IMEDIATA & ", " & IDASSINANTE & ", " & IDSOLIC & ", " & IDENDER & ", " & _
                        TIPOCLIENTE & ", " & IDPERIODO & ", " & IDTIPOFECH & ", " & IDPONTO & ", " & _
                        IDEQUIPE & ", " & IDOCORRENCIA & ", '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', " & _
                        CONV & ", '" & DTATEND & "', '" & HORAATEND & "')"

                    MysqlAdd("IGNORE", "ora_ord.ord_pendentes", MltQry, OnDuplcate, myComm, MyConn)
                    'strMY.INSERIR_NO_MYSQL("IGNORE", "ora_ord.ord_pendentes ", MltQry, OnDuplcate, myConn)

                End While
            End Using

            '-------------- DELETA AS NAO ATUALIZADAS
            'strMY.DELETA_TABELA("ora_ord.ord_pendentes", cfgcidade, "", myConn)
            MysqlDelete("ora_ord.ord_pendentes", "STATUS = '-'", MyConn, myComm)

            DIFHORA = Now - CDate(HORAINICIO)

            Return Now & " - " & System.Reflection.MethodBase.GetCurrentMethod.Name & RETORNO & " / CIDADE " & cfgcidade
            Exit Function
        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & cfgcidade & ") - " & ex.Message
            Exit Function
        End Try

    End Function

    ''' <summary>
    ''' IMPORTA ORA BAIXADAS
    ''' </summary>
    ''' <param name="cfgcidade"></param>
    ''' <param name="DATAORA"></param>
    ''' <param name="DATAMY"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function importaBaixadas(ByVal cfgcidade As String, ByVal DATAORA As String, ByVal DATAMY As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand) As String

        Dim RETORNO As String = ""
        Dim TORA As String = ""
        Dim QTDEINSERT, QTDEUPDATE As Double

        Dim DT_BAIXA, OS, CID, COD_BAIXA, CONTRATO, ID_EMP_EXECUCAO, ID_EQUIPE, USR_BAIXA, HORA_INI_EXEC, HORA_FIM_EXEC As String
        Dim ERRO As Integer = 0

        Dim FIELDS, TABLE, JOINS, WHERES As String

        QTDEINSERT = 0
        QTDEUPDATE = 0

        FIELDS = "CT.CID_CONTRATO, OS.COD_OS, CT.NUM_CONTRATO,to_char(OS.DT_BAIXA,'YYYY-MM-DD HH24:MI:SS') AS DT_BAIXA, " _
                 & " (SELECT BA.COD_BAIXA FROM PROD_JD.SN_REL_OS_BAIXA BA  " _
                 & " WHERE BA.COD_OS = OS.COD_OS AND ROWNUM = 1 ) AS COD_BAIXA, OS.USR_BAIXA, " _
                 & " to_char(OS.HR_INICIO_EXECUCAO,'YYYY-MM-DD HH24:MI:SS') HORA_INI_EXEC,  " _
                 & " to_char(OS.HR_TERMINO_EXECUCAO,'YYYY-MM-DD HH24:MI:SS') HORA_FIM_EXEC, " _
                 & " (SELECT ID_EMP_EXECUCAO FROM PROD_JD.SN_REL_OS_TAREFA T WHERE T.COD_OS = OS.COD_OS AND ROWNUM = 1) ID_EMP_EXECUCAO, " _
                 & " (SELECT ID_EQUIPE       FROM PROD_JD.SN_REL_OS_TAREFA T WHERE T.COD_OS = OS.COD_OS AND ROWNUM = 1) ID_EQUIPE, " _
                 & " (SELECT NM_LOCAL        FROM INTEGRACAO_ATLAS.MW_LOCAL where id_local = " _
                 & " (SELECT ID_EQUIPE       FROM PROD_JD.SN_REL_OS_TAREFA T WHERE T.COD_OS = OS.COD_OS AND ROWNUM = 1)) as NOME_TECNICO "

        TABLE = " PROD_JD.SN_OS OS "

        JOINS = " INNER JOIN PROD_JD.SN_ASSINANTE ASS ON ASS.ID_ASSINANTE = OS.ID_ASSINANTE " _
                 & " INNER JOIN PROD_JD.SN_CONTRATO CT   ON CT.ID_ASSINANTE  = OS.ID_ASSINANTE "

        WHERES = " DT_BAIXA BETWEEN TO_DATE('" & DATAORA & " 00:00:00','DD/MM/YYYY HH24:MI:SS') " _
                 & "            AND TO_DATE('" & DATAORA & " 23:59:59','DD/MM/YYYY HH24:MI:SS') " _
                 & "            AND CT.CID_CONTRATO  = '" & cfgcidade & "'"


        '#########################################################################################
        '#  USING CLIENT ORA <------------<<
        Try

            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text

            Using BxdRows As OracleDataReader = com.ExecuteReader()
                While BxdRows.Read()
                    OS = BxdRows.Item("COD_OS")
                    DT_BAIXA = Format(CDate(BxdRows.Item("DT_BAIXA")), "yyyy-MM-dd HH:mm:ss")
                    CID = BxdRows.Item("CID_CONTRATO")
                    CONTRATO = BxdRows.Item("NUM_CONTRATO")
                    If BxdRows.Item("COD_BAIXA") Is DBNull.Value Then COD_BAIXA = "0" Else COD_BAIXA = BxdRows.Item("COD_BAIXA")
                    If BxdRows.Item("ID_EMP_EXECUCAO") Is DBNull.Value Then ID_EMP_EXECUCAO = "0" Else ID_EMP_EXECUCAO = BxdRows.Item("ID_EMP_EXECUCAO").ToString
                    If BxdRows.Item("ID_EQUIPE") Is DBNull.Value Then ID_EQUIPE = "0" Else ID_EQUIPE = BxdRows.Item("ID_EQUIPE").ToString
                    USR_BAIXA = BxdRows.Item("USR_BAIXA").ToString
                    If BxdRows.Item("HORA_INI_EXEC") Is DBNull.Value Then HORA_INI_EXEC = "0000-00-00 00:00:00" Else HORA_INI_EXEC = BxdRows.Item("HORA_INI_EXEC").ToString
                    If BxdRows.Item("HORA_FIM_EXEC") Is DBNull.Value Then HORA_FIM_EXEC = "0000-00-00 00:00:00" Else HORA_FIM_EXEC = BxdRows.Item("HORA_FIM_EXEC").ToString

                    If ID_EQUIPE > 0 Then
                        MysqlAdd("", "ora_cfg.cfg_equipes (ID_EQUIPE,NOME_EQUIPE,SKILL)", "(" & ID_EQUIPE & ",'" & BxdRows.Item("NOME_TECNICO").ToString & "','INDOR')", _
                                        "ON DUPLICATE KEY UPDATE NOME_EQUIPE = '" & BxdRows.Item("NOME_TECNICO").ToString & "'", myComm, MyConn)
                    End If

                    MysqlAdd("IGNORE", "ora_ord.ord_baixadas (COD_OS, DT_BAIXA, CID_CONTRATO, NUM_CONTRATO, COD_BAIXA, ID_EMP_EXECUCAO, ID_EQUIPE, USR_BAIXA, HORA_INI_EXEC, HORA_FIM_EXEC,DT_UPDATE)", _
                            "(" & OS & ",'" & DT_BAIXA & "'," & CID & "," & CONTRATO & "," & COD_BAIXA & _
                            "," & ID_EMP_EXECUCAO & "," & ID_EQUIPE & ",'" & USR_BAIXA & "','" & HORA_INI_EXEC & "','" & HORA_FIM_EXEC & "','" & Format(CDate(Now), "yyyy-MM-dd HH:mm:ss") & "')", "", myComm, MyConn)
                End While
            End Using
            Return Now & " - " & System.Reflection.MethodBase.GetCurrentMethod.Name & " CIDADE " & cfgcidade
        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & cfgcidade & ") - " & ex.Message
        End Try

    End Function

    ''' <summary>
    ''' IMPORTA ORA AGENDADAS
    ''' </summary>
    ''' <param name="cfgcidade">CID_CONTRATO</param>
    ''' <param name="DATAORA">DATA DD/MM/YYYY</param>
    ''' <param name="DATAMY">DATA YYYY-MM-DD</param>
    ''' <returns>SUCESS/FAILS</returns>
    ''' <remarks>NAO DEVE DEMORAR</remarks>
    Public Function importaAgendadas(ByVal cfgcidade As String, ByVal DATAORA As String, ByVal DATAMY As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand) As String

        Dim TORA As Integer = 0
        Dim REPORT As String = ""

        Dim tabToinsert As String = "ora_ord.ord_agendadas (COD_OS,DT_AGENDA,CID_CONTRATO,ID_TIPO_OS,NUM_CONTRATO,STATUS,ID_TIPO_PERIODO,IMEDIATA,ID_ASSINANTE,ID_SOLICITACAO_ASS,ID_ENDER,TIPO_CLIENTE,ID_PERIODO,ID_TIPO_FECHAMENTO,ID_PONTO,ID_EQUIPE,ID_OCORRENCIA,DT_UPDATE, CONVENIENCIA, DT_ATEND ,HORA_ATEND)"
        Dim HORAATEND, DTATEND, OS, DT, CID, TIPOOS, CONTRATO, STATUS, PERIODO, IMEDIATA, IDASSINANTE, IDSOLIC, IDENDER, TIPOCLIENTE, IDPERIODO, IDTIPOFECH, IDPONTO, IDEQUIPE, IDOCORRENCIA, CONV, strduplicada As String : strduplicada = Nothing
        Dim HORAINICIO, strData, dpKey As String
        HORAINICIO = Now

        Dim FIELDS, TABLE, JOINS, WHERES As String

        FIELDS = " OS.COD_OS,PE.DT,CT.CID_CONTRATO,OS.ID_TIPO_OS,CT.NUM_CONTRATO,OS.STATUS, " _
               & "PE.ID_TIPO_PERIODO,OS.IMEDIATA,OS.ID_ASSINANTE,OS.ID_SOLICITACAO_ASS,OS.ID_ENDER,OS.ID_TIPO_CLIENTE, " _
               & "OS.ID_PERIODO,OS.ID_TIPO_FECHAMENTO,OS.ID_PONTO,OS.ID_EQUIPE,OS.ID_OCORRENCIA, HIS.fn_conveniencia AS CONV, to_char(OS.DT_ATEND,'YYYY-MM-DD HH24:MI:SS') DT_ATEND "

        TABLE = " PROD_JD.SN_OS OS "

        JOINS = " INNER JOIN PROD_JD.SN_CONTRATO CT ON CT.ID_ASSINANTE = OS.ID_ASSINANTE " _
              & " INNER JOIN PROD_JD.SN_PERIODO  PE ON PE.ID_PERIODO = OS.ID_PERIODO " _
              & " LEFT JOIN PROD_JD.TBSN_HISTORICO_AGENDAMENTO_OS HIS ON OS.COD_OS = HIS.COD_OS "

        WHERES = " PE.DT = TO_DATE('" & DATAORA & "','DD/MM/YYYY') AND CT.CID_CONTRATO = '" & cfgcidade & "' "


        '#########################################################################################
        '#  USING CLIENT ORA <------------<<
        Try
            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text

            '#  READ AGENDADAS
            Using AgnRows As OracleDataReader = com.ExecuteReader()
                While AgnRows.Read()
                    OS = AgnRows.Item("COD_OS")
                    DT = AgnRows.Item("DT")
                    CID = AgnRows.Item("CID_CONTRATO")
                    TIPOOS = AgnRows.Item("ID_TIPO_OS")
                    CONTRATO = AgnRows.Item("NUM_CONTRATO")
                    STATUS = AgnRows.Item("STATUS")
                    PERIODO = AgnRows.Item("ID_TIPO_PERIODO")

                    DTATEND = Format(CDate(AgnRows.Item("DT_ATEND")), "yyyy-MM-dd")
                    HORAATEND = Format(CDate(AgnRows.Item("DT_ATEND")), "HH:mm:ss")

                    If AgnRows.Item("IMEDIATA") Is DBNull.Value Then IMEDIATA = 0 Else IMEDIATA = 1 'verifica se imediata
                    IDASSINANTE = AgnRows.Item("ID_ASSINANTE")
                    IDSOLIC = AgnRows.Item("ID_SOLICITACAO_ASS")
                    IDENDER = AgnRows.Item("ID_ENDER")
                    TIPOCLIENTE = AgnRows.Item("ID_TIPO_CLIENTE")
                    IDPERIODO = AgnRows.Item("ID_PERIODO")
                    If AgnRows.Item("ID_TIPO_FECHAMENTO") Is DBNull.Value Then IDTIPOFECH = "0" Else IDTIPOFECH = AgnRows.Item("ID_TIPO_FECHAMENTO")
                    If AgnRows.Item("ID_PONTO") Is DBNull.Value Then IDPONTO = "0" Else IDPONTO = AgnRows.Item("ID_PONTO")
                    If AgnRows.Item("ID_EQUIPE") Is DBNull.Value Then IDEQUIPE = "0" Else IDEQUIPE = AgnRows.Item("ID_EQUIPE")
                    If AgnRows.Item("ID_OCORRENCIA") Is DBNull.Value Then IDOCORRENCIA = "0" Else IDOCORRENCIA = AgnRows.Item("ID_OCORRENCIA")
                    If AgnRows.Item("CONV") Is DBNull.Value Then CONV = "0" Else CONV = AgnRows.Item("CONV")

                    '// INSERE

                    strData = "( '" & OS & "','" & Format(CDate(DT), "yyyy-MM-dd") & "','" & CID & "','" & TIPOOS & "','" & CONTRATO & "','" & STATUS & "'," & _
                        "'" & PERIODO & "','" & IMEDIATA & "','" & IDASSINANTE & "','" & IDSOLIC & "','" & IDENDER & "'," & _
                        "'" & TIPOCLIENTE & "','" & IDPERIODO & "','" & IDTIPOFECH & "','" & IDPONTO & "','" & IDEQUIPE & "'," & _
                        "'" & IDOCORRENCIA & "','" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', " & CONV & ", '" & DTATEND & "', '" & HORAATEND & "') "

                    dpKey = "on duplicate key update " _
                                & "DT_AGENDA  = '" & Format(CDate(DT), "yyyy-MM-dd") & "', " _
                                & "STATUS     = '" & STATUS & "',   ID_TIPO_PERIODO    = '" & PERIODO & "',     DT_ATEND     = '" & DTATEND & "', HORA_ATEND = '" & HORAATEND & "', " _
                                & "IMEDIATA   = '" & IMEDIATA & "', ID_SOLICITACAO_ASS = '" & IDSOLIC & "',     TIPO_CLIENTE = '" & TIPOCLIENTE & "', " _
                                & "ID_PERIODO = '" & IDPERIODO & "',ID_TIPO_FECHAMENTO = '" & IDTIPOFECH & "',  ID_PONTO     = '" & IDPONTO & "', " _
                                & "ID_EQUIPE  = '" & IDEQUIPE & "', ID_OCORRENCIA      = '" & IDOCORRENCIA & "',DT_UPDATE    = '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "', CONVENIENCIA =  " & CONV


                    MysqlAdd("", tabToinsert, strData, dpKey, myComm, MyConn)
                End While
                AgnRows.Dispose()
                AgnRows.Close()
            End Using

            Return Now & " ATUALIZADO: " & System.Reflection.MethodBase.GetCurrentMethod.Name & " CIDADE " & cfgcidade

        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & cfgcidade & ") - " & ex.Message
        End Try

    End Function

    ''' <summary>
    ''' IMPORTA ORA REAGENDADAS
    ''' </summary>
    ''' <param name="CIDADE">INT [5]</param>
    ''' <param name="DATAORA">String [dd/MM/yyyy]</param>
    ''' <param name="DATAMY">String [yyyy-MM-dd]</param>
    ''' <param name="MyConn">MySQL CLIENT CONNECTION</param>
    ''' <param name="oraConn">ORA CLIENT CONNECTION</param>
    ''' <param name="oraComm">ORA CLIENT COMMAND</param>
    ''' <param name="myComm">MySQL CLIENT COMMAND</param>
    ''' <returns>NOTHING</returns>
    ''' <remarks>ENJOY</remarks>
    Public Function importaReagendadas(ByVal CIDADE As String, ByVal DATAORA As String, ByVal DATAMY As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand) As String
        Dim RETORNO As String = ""
        Dim TORA As String = ""

        Dim FIELDS, TABLE, JOINS, WHERES As String

        Dim DT_BAIXA, OS, ID_MOTIVO, ID_USR, ID_EMP_DESPACHO, CID, CONTRATO, ID_EQUIPE As String
        Dim strTable As String = "ora_ord.ord_reagendadas(COD_OS, DT_REAGEND, CID_CONTRATO, NUM_CONTRATO, ID_MOTIVO, ID_USR, ID_EMP_DESPACHO, ID_EQUIPE,DT_UPDATE)"
        Dim strData As String = ""

        Dim ERRO As Integer = 0

        '#  STR QUERY
        FIELDS = " R.NUM_CONTRATO, R.CID_CONTRATO, R.COD_OS, " & _
              " TO_CHAR(R.DATA,'YYYY-MM-DD HH24:MI:SS') DATA, " & _
              " R.ID_MOTIVO, R.ID_USR, R.ID_EMP_DESPACHO, R.ID_EQUIPE "

        TABLE = " PROD_JD.SN_LOG_REAGENDAMENTO_OS R "

        JOINS = ""

        WHERES = " R.DATA BETWEEN TO_DATE('" & DATAORA & " 00:00:00','DD/MM/YY HH24:MI:SS')  " _
               & "           AND TO_DATE('" & DATAORA & " 23:59:59','DD/MM/YY HH24:MI:SS')  " _
               & " AND R.CID_CONTRATO = '" & CIDADE & "'"

        Try

            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text

            Using ReadRows As OracleDataReader = com.ExecuteReader()
                While ReadRows.Read()
                    OS = ReadRows.Item("COD_OS")
                    DT_BAIXA = Format(CDate(ReadRows.Item("DATA")), "yyyy-MM-dd HH:mm:ss")
                    CID = ReadRows.Item("CID_CONTRATO")
                    CONTRATO = ReadRows.Item("NUM_CONTRATO")
                    ID_MOTIVO = ReadRows.Item("ID_MOTIVO")
                    ID_USR = ReadRows.Item("ID_USR").ToString
                    ID_EMP_DESPACHO = ReadRows.Item("ID_EMP_DESPACHO").ToString
                    ID_EQUIPE = ReadRows.Item("ID_EQUIPE").ToString

                    strData = "( '" & OS & "','" & Format(CDate(DT_BAIXA), "yyyy-MM-dd HH:mm:ss") & "','" & CID & "','" & CONTRATO & "'," & _
                        "'" & ID_MOTIVO & "','" & ID_USR & "','" & ID_EMP_DESPACHO & "','" & ID_EQUIPE & "','" & Format(CDate(Now), "yyyy-MM-dd HH:mm:ss") & "')"

                    MysqlAdd("IGNORE", strTable, strData, "", myComm, MyConn)

                End While
                ReadRows.Dispose()
                ReadRows.Close()
            End Using
            Return Nothing
        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & CIDADE & ") - " & ex.Message
        End Try

    End Function

    ''' <summary>
    ''' IMPORTA ORA SOLICITACOES
    ''' </summary>
    ''' <param name="CIDADE"></param>
    ''' <param name="DATAORA"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function importaSolicitacao(ByVal CIDADE As String, ByVal DATAORA As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand) As String

        Dim dtInsert As String = ""
        Dim Qry As String = ""
        Dim TORA As Integer = 0

        Dim FIELDS, TABLE, JOINS, WHERES As String


        '# = = =  RETURN ONE ROW OF QUERY = = = = = =
        Dim myFast As String = "SELECT max(id_solicitacao_ass) solic FROM ord_solic WHERE cid_contrato = " & CIDADE
        myComm = New MySqlCommand(myFast, MyConn)
        Dim mxSolic As String = Convert.ToString(myComm.ExecuteScalar())
        '# = = = = = = = = = = = = = = = = = = = = = =


        FIELDS = " ASS.CID_CONTRATO, ASS.ID_SOLICITACAO_ASS, ASS.DT_CADASTRO, ASS.NUM_CONTRATO, ASS.ID_TIPO_SOLIC "

        TABLE = " PROD_JD.SN_SOLICITACAO_ASS ASS "

        JOINS = ""

        If mxSolic = "" Then
            WHERES = " ASS.CID_CONTRATO = " & CIDADE & " AND ASS.DT_CADASTRO = '" & Format(CDate(DATAORA), "dd/MM/yyyy") & "' ORDER BY ASS.ID_SOLICITACAO_ASS "
        Else
            WHERES = " ASS.CID_CONTRATO = " & CIDADE & " AND   ASS.ID_SOLICITACAO_ASS >= " & mxSolic & " ORDER BY ASS.ID_SOLICITACAO_ASS "
        End If

        '#########################################################################################
        '#  USING CLIENT ORA <------------<<
        Try

            '#  DECLARA VARIAVEIS ORA
            Dim com As New OracleCommand

            '#  RODA QUERY GERAL
            Dim QryStr As String = "SELECT " & FIELDS & " FROM " & TABLE & " " & JOINS & " WHERE " & WHERES
            com = New OracleCommand(QryStr, oraConn)
            com.CommandType = CommandType.Text

            Using SolRows As OracleDataReader = com.ExecuteReader()
                While SolRows.Read()

                    dtInsert = " (" & SolRows.Item("ID_SOLICITACAO_ASS") & "," & _
                            SolRows.Item("CID_CONTRATO") & "," & _
                            "'" & Format(CDate(SolRows.Item("DT_CADASTRO")), "yyyy-MM-dd") & "'," & _
                            SolRows.Item("NUM_CONTRATO") & "," & _
                            SolRows.Item("ID_TIPO_SOLIC") & ")"

                    MysqlAdd("IGNORE", "ora_ord.ord_solic", dtInsert, "", myComm, MyConn)

                End While
            End Using

            Return Now & " - ATUALIZADO: " & System.Reflection.MethodBase.GetCurrentMethod.Name & " CIDADE " & CIDADE

        Catch ex As Exception
            Return " ERRO " & System.Reflection.MethodBase.GetCurrentMethod.Name & " (cidade " & CIDADE & ") - " & ex.Message
        End Try

    End Function

    ''' <summary>
    ''' IMPORTA ORA CANCELADAS
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="cidContrato"></param>
    ''' <param name="MyConn"></param>
    ''' <param name="oraConn"></param>
    ''' <param name="oraComm"></param>
    ''' <param name="myComm"></param>
    ''' <remarks></remarks>
    Sub importaCanceladas(ByVal dt As String, ByVal cidContrato As String, ByVal MyConn As MySqlConnection, ByVal oraConn As OracleConnection, ByVal oraComm As OracleCommand, ByVal myComm As MySqlCommand)

        Dim count As Integer = 0
        Dim tbInsert As String = "ora_ord.ord_canceladas(CID_CONTRATO, COD_OS, NUM_CONTRATO, ID_ASSINANTE, ID_TIPO_OS, ID_ENDER, ID_OCORRENCIA, DT_BAIXA, ID_SOLICITACAO_ASS, OBS, STATUS, USR_BAIXA, DESCRICAO, ID_EMP_EXECUCAO)"


        Dim FIELDS, TABLE, JOINS, WHERES, OBS, strSql As String
        Dim ID_TIPO_OS As String = "10, 21, 22, 26, 27, 38, 42, 48, 49, 50, 61, 62, 67, 69, 204" '// VISITAS


        FIELDS = " CT.CID_CONTRATO, CT.NUM_CONTRATO, OS.COD_OS, OS.ID_ASSINANTE, OS.ID_TIPO_OS, " & _
                    " OS.ID_ENDER, OS.ID_OCORRENCIA, to_char(OS.DT_BAIXA,'YYYY-MM-DD HH24:MI:SS') AS DT_BAIXA,  " & _
                    " OS.USR_BAIXA, OS.ID_SOLICITACAO_ASS, OS.OBS, F.DESCRICAO AS STATUS, RZC.DESCRICAO, " & _
                    " (SELECT ID_EMP_EXECUCAO FROM PROD_JD.SN_REL_OS_TAREFA T WHERE T.COD_OS = OS.COD_OS AND ROWNUM = 1)  AS ID_EMP_EXECUCAO "

        TABLE = " PROD_JD.SN_OS OS "

        JOINS = " INNER JOIN PROD_JD.SN_ASSINANTE                ASS  ON ASS.ID_ASSINANTE = OS.ID_ASSINANTE " & _
                    " INNER JOIN PROD_JD.SN_CONTRATO             CT   ON CT.ID_ASSINANTE  = OS.ID_ASSINANTE " & _
                    " LEFT JOIN PROD_JD.SN_TIPO_FECHAMENTO       F    ON OS.ID_TIPO_FECHAMENTO = F.ID_TIPO_FECHAMENTO " & _
                    " LEFT JOIN PROD_JD.SN_SOLICITACAO_ASS       SOL  ON OS.ID_SOLICITACAO_ASS = SOL.ID_SOLICITACAO_ASS AND CT.CID_CONTRATO = SOL.CID_CONTRATO " & _
                    " LEFT JOIN PROD_JD.SN_RAZAO_CANCELAMENTO_OS RZC  ON SOL.ID_RAZAO_CANCELAMENTO = RZC.ID_RAZAO_CANCELAMENTO_OS  "

        WHERES = "  OS.DT_BAIXA BETWEEN TO_DATE('" & dt & " 00:00:00','DD/MM/YYYY HH24:MI:SS') " & _
                    " AND TO_DATE('" & dt & " 23:59:59','DD/MM/YYYY HH24:MI:SS') " & _
                    " AND CT.CID_CONTRATO  = '" & cidContrato & "' AND OS.ID_TIPO_FECHAMENTO = 3 " & _
                    " AND OS.ID_TIPO_OS IN (" & ID_TIPO_OS & ") "

        Dim QryStr As String = "Select " & FIELDS & " From " & TABLE & JOINS & " Where " & WHERES

        Try

            oraComm = New OracleCommand(QryStr, oraConn)

            Using GerRow As OracleDataReader = oraComm.ExecuteReader()
                While GerRow.Read()
                    count += 1

                    OBS = Replace(GerRow.Item("OBS").ToString, "'", ".")
                    OBS = Replace(OBS, "\", "-")

                    If Len(OBS) < 1 Then OBS = "null"

                    strSql = "( " & GerRow.Item("CID_CONTRATO") & ", " & GerRow.Item("COD_OS") & ", " & GerRow.Item("NUM_CONTRATO") & ", " & _
                                GerRow.Item("ID_ASSINANTE") & ", " & GerRow.Item("ID_TIPO_OS") & ", " & GerRow.Item("ID_ENDER") & ", " & _
                                IIf(IsDBNull(GerRow.Item("ID_OCORRENCIA")) = True, "0", GerRow.Item("ID_OCORRENCIA")) & ", '" & GerRow.Item("DT_BAIXA") & "', " & _
                                GerRow.Item("ID_SOLICITACAO_ASS") & ", '" & OBS & "', '" & GerRow.Item("STATUS") & "', '" & GerRow.Item("USR_BAIXA") & "', '" & GerRow.Item("DESCRICAO") & "', " & _
                                IIf(IsDBNull(GerRow.Item("ID_EMP_EXECUCAO")) = True, "0", GerRow.Item("ID_EMP_EXECUCAO")) & ")"

                    MysqlAdd("IGNORE", tbInsert, strSql, "", myComm, MyConn)

                End While
            End Using
        Catch ex As Exception
            '# ERRO AQUI
            'MsgBox(ex.Message)
            'ErrCatcher.LogErro(System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message, QryStr)
        End Try

    End Sub



End Class
