Public Class errController

    Private mCtrl As New mailController
    Private wCtrl As New writerController

    Public Sub logController(ByVal strFuncao As String, ByVal strDetalhe As String)
        'mCtrl.sent("fagner.sivla@net.com.br", "email test", "corpo da mensagem")
        wCtrl.write(Format(Now, "yyyy-MM-dd"), strFuncao & " - " & strDetalhe)
    End Sub

    

End Class
