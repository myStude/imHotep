Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Imports MySql.Data.MySqlClient
Imports System.Threading

Public Class updateController
    Inherits objConnection

    Private eCtrl As New errController


    'ORA_ORD
    Private objSrv As New objSERVICOS

    'ORA_CFG
    Private Imoveis As New IMOVEIS
    Private hpImovel As New HP_IMOVEL
    Private Endereco As New ENDERECOS
    Private nodeArea As New NODE_AREA
    Private enderLogradouro As New ENDERECO_LOGRADOURO
    Private enderBairro As New ENDERECO_BAIRRO

    Private threadSleep As Integer() = {600000, 120000, 1200000} '600.000 Milisegundos = 10 Minutos

    ''' <summary>
    ''' Controlador dos metodos de Espelho
    ''' </summary>
    ''' <creator>Fagenrdin [n5774826]</creator>
    ''' <param name="cidContrato">int[5]</param>
    ''' <param name="codCidade">int[3]</param>
    ''' <param name="netBase">String[6]</param>
    ''' <param name="dtUpdate">Curdate[yyyy-MM-yy]</param>
    ''' <remarks>Enjoy</remarks>
    Sub espelhoORA(ByVal cidContrato As String, ByVal codCidade As String, ByVal netBase As String, ByVal dtUpdate As String)

        Dim DATAORA As String = Format(CDate(dtUpdate), "dd/MM/yyyy")
        Dim DATAMY As String = Format(CDate(dtUpdate), "yyyy-MM-dd")

        Dim dtGrid As Object = My.Forms.FrmHome.dtaGrid

        Try
            Using oraConn As New OracleConnection(oraConnection(netBase))
                oraConn.Open()
                Dim oraComm As New OracleCommand
                Using MyConn As New MySqlConnection(MySqlConnection("ora_ord"))
                    MyConn.Open()
                    Dim myComm As New MySqlCommand
                    With objSrv
                        .importaReagendadas(cidContrato, DATAORA, DATAMY, MyConn, oraConn, oraComm, myComm)
                        .importaAgendadas(cidContrato, DATAORA, DATAMY, MyConn, oraConn, oraComm, myComm)
                        .importaBaixadas(cidContrato, DATAORA, DATAMY, MyConn, oraConn, oraComm, myComm)
                        .importaPendentes(cidContrato, DATAORA, DATAMY, 0, MyConn, oraConn, oraComm, myComm)
                        .importaGeradas(cidContrato, "", DATAORA, DATAMY, MyConn, oraConn, oraComm, myComm)
                        .importaSolicitacao(cidContrato, DATAORA, MyConn, oraConn, oraComm, myComm)
                        .importaCanceladas(DATAORA, cidContrato, MyConn, oraConn, oraComm, myComm)
                    End With
                    MyConn.Dispose()
                End Using
                oraConn.Dispose()
            End Using
            Thread.Sleep(threadSleep(0))
        Catch ex As Exception
            dtGrid.rows.add(ex.Message)
            'eCtrl.logController(System.Reflection.MethodBase.GetCurrentMethod.Name, ex.Message)
            Thread.Sleep(threadSleep(1))
        End Try


    End Sub

    ''' <summary>
    ''' CONTROLADOR ESPELHO CFGs
    ''' </summary>
    ''' <param name="cidContrato">INT[5]</param>
    ''' <param name="codCidade">INT[3]</param>
    ''' <param name="netBase">STR[6]</param>
    ''' <param name="dtUpdate">DATE[DD-MM-YYYY]</param>
    ''' <remarks>ENJOY</remarks>
    Sub espelhoCFG(ByVal cidContrato As String, ByVal codCidade As String, ByVal netBase As String, ByVal dtUpdate As String)
        Try
            Using oraConn As New OracleConnection(oraConnection(netBase))
                oraConn.Open()
                Dim oraComm As New OracleCommand
                Using MyConn As New MySqlConnection(MySqlConnection("ora_cfg"))
                    MyConn.Open()
                    Dim myComm As New MySqlCommand

                    hpImovel.Atualiza(codCidade, netBase, MyConn, oraConn, oraComm, myComm)
                    Imoveis.importaImoveis(cidContrato, codCidade, MyConn, oraConn, oraComm, myComm)
                    Endereco.importaEnderecos("", codCidade, MyConn, oraConn, oraComm, myComm)
                    nodeArea.importaNodeArea(cidContrato, MyConn, oraConn, oraComm, myComm)
                    enderLogradouro.impotaEnderLogradouro(cidContrato, MyConn, oraConn, oraComm, myComm)
                    enderBairro.importaBairro(cidContrato, MyConn, oraConn, oraComm, myComm)

                    MyConn.Dispose()
                End Using
                oraConn.Dispose()
            End Using
            Thread.Sleep(threadSleep(2))
        Catch ex As Exception
            Dim a As String = ex.Message
            Thread.Sleep(threadSleep(1))
        End Try
    End Sub


End Class
