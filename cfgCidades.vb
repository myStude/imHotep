''' <summary>
''' Array list com a configuração das cidades
''' </summary>
''' <remarks>enjoy</remarks>
Public Class cfgCidades

    Public Property cidContrato() As Integer
    Public Property codCidade() As Integer
    Public Property strBase() As String

    Public Sub New(vcidContrato As Integer, vcodCidade As Integer, vstrBase As String)
        cidContrato = vcidContrato
        codCidade = vcodCidade
        strBase = vstrBase
    End Sub
End Class